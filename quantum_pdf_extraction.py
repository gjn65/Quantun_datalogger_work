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
"""

from progress.bar import Bar
#import pprint
import xlsxwriter
from datetime import datetime
import pdfplumber
import quantum_pdf_extraction_cfg as cfg






def main():

    loco_name=""
    start_epoch=get_epoch(cfg.start_timestamp)
    end_epoch=get_epoch(cfg.end_timestamp)

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
                ws_row=write_header(wb,ws,"Data extract from Quantum Data Recorder",loco_name)
                ws_row_ann=write_header_ann(wb,ws_ann,"Data extract from Quantum Data Recorder",loco_name)
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
                    record_epoch=get_epoch(record_date+" "+record_time)
                    if start_epoch > 0 and ((record_epoch < start_epoch) or (record_epoch > end_epoch)):
                        continue

                    ws.write(ws_row,0,record_date)
                    ws.write(ws_row,1,record_time)
                    ws.write(ws_row,2,' '.join(words[:-2]),lalign)
                    ws_row+=1
                    ws_ann.write(ws_row_ann,0,record_date)
                    ws_ann.write(ws_row_ann,1,record_time)
                    ws_ann.write(ws_row_ann,2," ".join(words[:-2]))
                    ws_row_ann+=1
                    continue

                # Split the line into words on whitespace
                words=line.split()
                #pp.pprint(words)
                record_date=convert_date(words[1])
                record_time=words[0].replace("-","")
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
    ws.write(ws_row, ws_col, "Y" if words[8] == "1" else "N")  # Reverse
    ws_col += 1
    ws.write(ws_row, ws_col, "Y" if words[9] == "1" else "N")  # EIE
    ws_col += 1
    ws.write(ws_row, ws_col, "Y" if words[10] == "1" else "N")  # Pressure control switch
    ws_col += 1
    ws.write(ws_row, ws_col, "Y" if words[11] == "1" else "N")  # Headlight on (short end)
    ws_col += 1
    ws.write(ws_row, ws_col, "Y" if words[12] == "1" else "N")  # Forward
    ws_col += 1
    ws.write(ws_row, ws_col, "Y" if words[13] == "1" else "N")  # Headlight on (long end)
    ws_col += 1
    ws.write(ws_row, ws_col, "Y" if words[14] == "1" else "N")  # Horn on
    ws_col += 1
    ws.write(ws_row, ws_col, "Y" if words[15] == "1" else "N")  # Digital Spare 1
    ws_col += 1
    ws.write(ws_row, ws_col, "Y" if words[16] == "1" else "N")  # Digital Spare 2
    ws_col += 1
    ws.write(ws_row, ws_col, "Y" if words[17] == "1" else "N")  # VC alert acknowledge
    ws_col += 1
    ws.write(ws_row, ws_col, "Y" if words[18] == "1" else "N")  # Axle Drive type
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


def write_header_ann(wb,ws,text,loco_name):

    lalign=wb.add_format({'align':'left'})
    ws.set_column('A:B',15,lalign)
    ws.set_column('C:C',100,lalign)
    header_format_ann=wb.add_format({'font_size':14,'bold':True})
    ws.freeze_panes(3,0)
    """ Write header line to the worksheet. Return the next row number (0 based) """
    ws.write(0,0,text+" : "+loco_name,header_format_ann)
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



if __name__ == '__main__':
    main()
