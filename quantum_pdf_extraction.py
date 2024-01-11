#!/Library/Frameworks/Python.framework/Versions/3.10/bin/python3.10


"""

Quantum Desktop Playback - data reporter

This code will parse a PDF file and create an Excel worksheet.
The PDF file is created by running the QDP software, selecting the timescale
required using tags (or select the entire file) then printing it to a PDF file

When setting up the printed page, select ALL variables logged - this code will weed out those
that are not required.

The header page of the PDF report is used ot make adjustments to the wheel speed from that
recorded in the data (which can be modified if the data is extracted using QDP but cannot if
it's downloaded using the QRST code).

All logger data is in imperial units, so we need to convert to metric.

The cfg file contains all the data required to drive the application operation.


							Maintenance History
							
March 2023	GJN	Initial Creation
2023/03/23  GJN Added Throttle position translation for text fields
                Use RE to solve issue of 2 character TP values stuffing up line splitting.
2023/03/26  GJN Streamline worksheet writing code.
                Hide unwanted columns
                Protect code is written but not implemented as I'm still working out
                how to allow users to filter protected data...
                Add progress bar.

2023/03/27  GJN Stop using PDFReader as it was munging the data in cases where the TMC exceeded 3 digits
                or the Throttle Position was 2 characters.
                Now we use pdfplumber instead to extract the PDF data

2023/06/26  GJN For each annotation detected, print the time of the previous event and
                the timestamp delta. This will help when diagnosing the power up events.

2023/07/26  GJN When printing annotation events on event page, only print inter-event
                details for Power related events.

2023/08/31  GJN When writing to the Excel sheet, for the digital inputs, use a range function
                rather than individual writes.

2024/01/11  GJN Add ability to apply timeof day offset to logger records to compensate for
                the loss of accuracy of the logger's RTC.
                Offset is recorded in the configuration file.
                Compensation and other run time modifiers are recorded in 3rd worksheet
                as a matter of record

NB: Low Idle position allows the engine to idle lower than normal to save fuel. 
	Not used on our 830 or 930 class locomotives.

	DYN refers to dynamic braking, the 830 class does not have this. The instances
	seen in the data were caused by spurious voltages appearing on the dynamic brake
	signal leads connected to the MU cable in 844's electrical cabinet. The presence
	of +74VDC on these wires masked the correct throttle settings. (Refer to the TP
	truth table in the Quantum Speedometer manuai). We disconnected these 2 signal feeds
	into the data logger in June/July 2023 to fix this problem




"""

from progress.bar import Bar
import pprint
import os
import xlsxwriter
from datetime import datetime
import pdfplumber
import quantum_pdf_extraction_cfg as cfg






def main():

   # pp = pprint.PrettyPrinter(indent=4)

    loco_name=""
    start_epoch=get_epoch(cfg.start_timestamp)
    end_epoch=get_epoch(cfg.end_timestamp)

    if cfg.filter_dates:
        print("Start from "+cfg.start_timestamp)
        print("End at "+cfg.end_timestamp)
    else:
        print("No record filtering required")
    print("Input = "+cfg.source_file)
    pdf=pdfplumber.open(cfg.source_file)
    pages=len(pdf.pages)
    wb_name=""
    # Iterate through each page
    old_page=0
    with Bar('Processing...', max=pages, width=80) as bar:
        for page in range(pages):
            #print("Processing page "+str(page+1)+" of "+str(pages))
            if page == 1 and old_page ==0:
                wb_name=cfg.workbook_name+" "+loco_name+" "+datetime.now().strftime("%Y%m%d%H%M")+".xlsx"
                wb=xlsxwriter.Workbook(wb_name,{'strings_to_numbers':True})
                ws=wb.add_worksheet(cfg.worksheet_name)
                lalign=wb.add_format({'align':'left'})
                ws_ann=wb.add_worksheet("Logger Events")
                parts=os.path.split(cfg.source_file)
                ws_modifiers=wb.add_worksheet("Runtime modifiers")
                ws_row=write_header(wb,ws,"Data extract from Quantum Data Recorder","Locomotive "+loco_name+". Source file "+parts[1])
                ws_row_ann=write_header_ann(wb,ws_ann,"Data extract from Quantum Data Recorder",loco_name)
                ws_row_modifiers=write_header_modifiers(wb,ws_modifiers,"Runtime modifiers and events")
                if cfg.filter_dates:
                    ws_modifiers.write(ws_row_modifiers,0,"Records selected from "+cfg.start_timestamp+" to "+cfg.end_timestamp)
                    ws_row_modifiers+=1
                else:
                    ws_modifiers.write(ws_row_modifiers,0,"No record filtering in place")
                    ws_row_modifiers+=1
                ws_modifiers.write(ws_row_modifiers,0,"Record timestamp offset applied is "+str(cfg.ts_adjustment)+" seconds")
                ws_row_modifiers += 1
                ws_modifiers.write(ws_row_modifiers, 0,
                               "Speed adjustment factor applied. QDP defined wheel diameter = " + str(
                                   wheel_diameter_qdp_inches * 25.4) + ". Measured wheel diameter = " + str(
                                   cfg.wheel_dia_actual_mm) + ". Adjustment factor = " + str(
                                   cfg.speed_adjustment_factor) + ".")
                ws_row_modifiers += 1
            #print("Input = " + cfg.source_file)
            old_page=page

            page_contents = pdf.pages[page]
            text = page_contents.extract_text()
            lines = text.split('\n')

            # Iterate through each line for this page
            for line in lines:
                if len(line) == 0:
                    continue

                # Search for wheel diameter figure in 1st page, but only if adjustment factor has not already been established
                if page == 0:
                    if "Locomotive Number" in line:
                        words=line.split()
                        loco_name=words[-1]
                    if cfg.speed_adjustment_factor == 0:
                        if "Circumference" in line and "Diameter" in line:
                            words=line.split()
                            #pp.pprint(words)
                            wheel_diameter_qdp_inches=float(words[-1])    # wheel diameter according to the QDP software
                            cfg.speed_adjustment_factor = cfg.wheel_dia_actual_mm / (wheel_diameter_qdp_inches*25.4)
                            #pp.pprint(cfg.speed_adjustment_factor)
                            continue
                    continue        # We don't want anything else from page 0

                # Skip lines with strings we are not interested in
                if skip_line_found(line):
                    continue

                # If 1st character in the line is non-numeric we treat the line as an annotation
                # for example - recorder power up, laptop connection etc.
                if not line[0].isnumeric():
                    # Handle annotations
                    # Extract date and time - last word in string in format HH:MM:SS-mm/dd/yyyy
                    words=line.split()
                    record_date=convert_date(words[-1])
                    record_time=words[-2].replace("-","")
                    record_date,record_time=apply_time_adjustment(record_date,record_time)
                    record_epoch=get_epoch(record_date+" "+record_time)
                    if start_epoch > 0 and ((record_epoch < start_epoch) or (record_epoch > end_epoch)):
                        continue

                    ws.write(ws_row,0,record_date)
                    ws.write(ws_row,1,record_time)
                    ws.write(ws_row,2,' '.join(words[:-2]),lalign)
                    ws_row+=1

                    # Calculate offset between this annotation and the previous record.
                    s = datetime.strptime(old_record_date + " " + old_record_time, "%Y/%m/%d %H:%M:%S")
                    e = datetime.strptime(record_date + " " + record_time, "%Y/%m/%d %H:%M:%S")
                    offset=str(e-s)

                    ws_ann.write(ws_row_ann,0,record_date)
                    ws_ann.write(ws_row_ann,1,record_time)
                    ws_ann.write(ws_row_ann,2," ".join(words[:-2]))
                    # Only write inter-event interval for power related events.
                    if words[0][:5] == 'Power':
                        ws_ann.write(ws_row_ann,3,old_record_date)
                        ws_ann.write(ws_row_ann,4,old_record_time)
                        ws_ann.write(ws_row_ann,5,offset)
                    ws_row_ann+=1

                    continue

                # Split the line into words on whitespace
                words=line.split()
                #pp.pprint(words)
                record_date=convert_date(words[1])
                record_time=words[0].replace("-","")
                record_date, record_time = apply_time_adjustment(record_date, record_time)
                old_record_date=record_date
                old_record_time=record_time
                record_epoch=get_epoch(record_date+" "+record_time)
                if cfg.filter_dates and ((record_epoch < start_epoch) or (record_epoch > end_epoch)):
                    continue
                # Write record to spreadsheet
                ws_row=write_record(ws,ws_row,words,record_date,record_time)
            bar.next()
        bar.finish()
    hide_columns(ws,cfg.headers)
 #   ws.protect(cfg.protect_string,cfg.protection_mode)
  #  ws_ann.protect(cfg.protect_string,cfg.protection_mode)
    wb.close()
    print("Written file : "+wb_name)

def hide_columns(ws,headers):
    """ Hide any column with False in the header tuple """
    for column, record in enumerate(headers):
        if not record[1]:
            ws.set_column(column,column,None,None,{'hidden':True})


def    write_record(ws,ws_row,words,record_date,record_time):
    """ Write spreadsheet row, return updated row number """


    # Time - has a dash appended
    # Date - is in mm/dd/yyyy format
    # Miles - 2 decimal places
    # Speed - MPH?
    # Traction motor current
    # Brake pipe pressure
    # Independent brake pressure
    # Throttle notch

    # Flags are either 1 or 0 for ON or OFF
    # Reverser in reverse
    # Engineer induced emergency
    # Pressure control switch (set when the BP air drops below 45 psi)
    # Headlight on - short end
    # Reverser in forward
    # Headlight on - long end
    # Horn on
    # Digital spare 1
    # Digital spare 2
    # Vigilance Control Alert acknowledge
    # Axle drive type

    ws_col = 0
    ws.write_string(ws_row, ws_col, record_date)  # AUS Date stamp
    ws_col += 1
    ws.write_string(ws_row, ws_col, record_time)  # Timestamp
    ws_col += 1
    #ws.write(ws_row, ws_col, "{:.2f}".format(float(words[2]) * 1.6))  # Mileage converted to km units
    ws.write_number(ws_row, ws_col, float(words[2]) * 1.6)  # Mileage converted to km units
    ws_col += 1
    # Speed, converted to kph and adjusted according to the difference between the real wheel diameter
    # and the diameter reported by the QDP software. NB: The reported wheel diameter can be set when
    # downloading the data via QDP but not when downloading via the QRST software.
    ws.write_number(ws_row, ws_col, round(int(words[3]) * 1.6 * cfg.speed_adjustment_factor))
    ws_col += 1
    ws.write_number(ws_row, ws_col, int(words[4]))  # TMC
    ws_col += 1
    ws.write_number(ws_row, ws_col, int(words[5]))  # Brake pipe pressure
    ws_col += 1
    ws.write_number(ws_row, ws_col, int(words[6]))  # Independent brake pressure
    ws_col += 1
    ws.write(ws_row,ws_col,translate_tp(words[7]))  # Throttle position
    ws_col += 1
    # Digital inputs follow
    for i in range(8,18,1):
        ws.write(ws_row, ws_col, "Y" if words[i] == "1" else "N")
        ws_col += 1
    ws_row += 1
    return ws_row

def translate_tp(tp):
    """ take a throtle position. If it's a number, then return that number.
        If it's a letter then returnn the corrsponding text """
    if tp.isnumeric():
        return tp
    if tp.upper() in cfg.tp_translations.keys():
        return cfg.tp_translations[tp.upper()]
    return tp+" (Unknown)"

def convert_date(us_date):
    """ Convert US formatted date to AUS formatted date """
    parts=us_date.split("/")
    if len(parts)!=3:
        return "invalid"
    aus_date="{:0>4d}/{:0>2d}/{:0>2d}".format(int(parts[2]),int(parts[0]),int(parts[1]))
    return aus_date

def write_header(wb,ws,text,loco_name):

    wb.set_size(1920,1080)
    lalign=wb.add_format({'align':'left'})
    ws.set_column('A:B',15,lalign)
    numfmt=wb.add_format({'num_format':'0.00'})
    ws.set_column('C:C',10,numfmt)
    ralign=wb.add_format({'align':'right'})
    ws.set_column('D:H',10,ralign)
    calign=wb.add_format({'align':'center'})
    ws.set_column('I:S',10,calign)
    ws.set_column('T:T',20,lalign)

    calign_b=wb.add_format({'align':'center','bold':True})
    ws.set_row(1,None,calign_b)

    header_format=wb.add_format({'font_size':14,'bold':True})

    ws.freeze_panes(3,0)

    """ Write header line to the worksheet. Return the next row number (0 based) """
    ws.write(0,0,text+" : "+loco_name,header_format)

    for column, record in enumerate(cfg.headers):
        ws.write(1,column,record[0])
    return 3


def write_header_modifiers(wb,ws,text):

    lalign=wb.add_format({'align':'left'})
    ws.set_column('A:A',150,lalign)
    header_format_modifiers=wb.add_format({'font_size':14,'bold':True})
    ws.freeze_panes(3,0)
    """ Write header line to the worksheet. Return the next row number (0 based) """
    ws.write(0,0,text,header_format_modifiers)
    return 3

def write_header_ann(wb,ws,text,loco_name):

    lalign=wb.add_format({'align':'left'})
    ws.set_column('A:B',15,lalign)
    ws.set_column('C:C',50,lalign)
    ws.set_column('D:F',15,lalign)

    header_format_ann=wb.add_format({'font_size':14,'bold':True})
    ws.freeze_panes(3,0)
    """ Write header line to the worksheet. Return the next row number (0 based) """
    ws.write(0,0,text+" : "+loco_name,header_format_ann)
    ws.write(1,0,"Event Date",header_format_ann)
    ws.write(1,1,"Event Time",header_format_ann)
    ws.write(1,2,"Event Type",header_format_ann)
    ws.write(1,3,"Prev Evt Date",header_format_ann)
    ws.write(1,4,"Prev Evt Time",header_format_ann)
    ws.write(1,5,"Offset", header_format_ann)
    return 3

def skip_line_found(line):
    """
        Search for existence of skip_list word(s) in the line variable passed into the function.
        If found then return True, otherwise return False
    """
    for word in cfg.skip_list_words:
        if word in line:
            return True
    return False


def get_epoch(timestamp):
    """ Take string in format yyyy/mm/dd hh:mm:ss and return epoch seconds (or 0 if flag is false) """
    if not cfg.filter_dates:
        return 0
    d = datetime.strptime(timestamp,"%Y/%m/%d %H:%M:%S")
    epoch=datetime(d.year,d.month,d.day,d.hour,d.minute,d.second).timestamp()
    return int(epoch)


def isfloat(num):
    try:
        float(num)
        return True
    except ValueError:
        return False

# The logger realtime clock can vary from the actual time so the timestamps
# are not accurate. This function will normalise the timestamps based on the
# TOD adjustment factor. If the logger clock is behind the real clock then a
# positive adjustment is made, if ahead then a negative adjustment is made.
# The adjustment factor is in seconds.
def apply_time_adjustment(date,time):
    d = datetime.strptime(date+" "+time,"%Y/%m/%d %H:%M:%S")
    epoch=datetime(d.year,d.month,d.day,d.hour,d.minute,d.second).timestamp()
    epoch+=cfg.ts_adjustment
    datetime_obj = datetime.fromtimestamp(epoch)
    date = datetime_obj.strftime("%Y/%m/%d")
    time = datetime_obj.strftime("%H:%M:%S")
    return date,time

if __name__ == '__main__':
    main()
