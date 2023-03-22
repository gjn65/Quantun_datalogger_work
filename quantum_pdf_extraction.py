'''

Quantum Desktop Playback - data reporter

This code will parse a PDF file and create an Excel worksheet.
The PDF file is created by running the QDP software, selecting the timescale
required using tags (or select the entire file) then printing it to a PDF file

When setting up the printed page, select ALL variables logged - this code will weed out those
that are not required.

Th header page of the PDF report is used ot make adjustments to the wheel speed from that
recorded in the data (which can be modified if the data is extracted using QDP but cannot if
it's downloaded using the QRST code).

All logger data is in imperial units so we need to convert to metric.

The cfg file contains all the data required to drive the application operation.


							Maintenance History
							
March 2023	GJN	Initial Creation
2023/03/23  GJN Added Throttle position translation for text fields
                Use RE to solve issue of 2 character TP values stuffing up line splitting.


'''

import re
import pprint
from pypdf import PdfReader
import xlsxwriter
from datetime import datetime
import quantum_pdf_extraction_cfg as cfg

pp=pprint.PrettyPrinter(indent=4)

def main():

    loco_name=""
    start_epoch=get_epoch(cfg.start_timestamp)
    end_epoch=get_epoch(cfg.end_timestamp)

    reader=PdfReader(cfg.source_file)
    pages=len(reader.pages)
    # Iterate through each page
    old_page=0
    for page in range(pages):
        print("Processing page "+str(page+1)+" of "+str(pages))
        if page == 1 and old_page ==0:
            wb=xlsxwriter.Workbook(cfg.workbook_name+" "+loco_name+" "+datetime.now().strftime("%Y%m%d%H%M")+".xlsx")
            ws=wb.add_worksheet(cfg.worksheet_name)
            ws_ann=wb.add_worksheet(("Events"))
            ws_row=write_header(wb,ws,loco_name)
            ws_col=0
            ws_row_ann=0
        old_page=page
        lines=reader.pages[page].extract_text().split('\n')
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
            if skip_line_found(line) == True:
                continue

            # If 1st character in the line is non numeric we treat the line as an annotation
            # for example - recorder power up, laptop connection etc.
            if line[0].isnumeric() == False:
                # Handle annotations
                # Extract date and time - last word in string in format HH:MM:SS-mm/dd/yyyy
                words=line.split()
                record_date=convert_date(words[-1])
                record_time=words[-2].replace("-","")
                record_epoch=get_epoch(record_date+" "+record_time)
                if start_epoch > 0 and ((record_epoch < start_epoch) or (record_epoch > end_epoch)):
                    continue

                ws.write(ws_row,0,line)
                ws_row+=1
                ws_ann.write(ws_row_ann,0,line)
                ws_row_ann+=1
                continue

            # Split the line into words on whitespace
            # We want the first 8 fields separately:
            # Time - has a dash appended
            # Date - is in mm/dd/yyyy format
            # Miles - 2 decimal places
            # Speed - MPH?
            # Traction motor current
            # Brake pipe pressure
            # Independent brake pressure
            # Throttle notch (may be 'D')
            # The rest of the fields are either 1 or 0 for ON or OFF but the PDF extractor
            # messes up the separators so we will squeeze the spaces out then use positional
            # references
            #
            # NOTE: - If the Throttle Position value is 2 letters (currently ID has been found as a case)
            #         the PDF parser will concatenate that with the preceding field (the IBRK value)
            #         This will stuff up the splitting by words so we will search for the string ID preceded
            #         by a digit and if this is found then we will replace "ID" with " I"
            if re.search("[0-9]ID",line) != None:
                line=line.replace("ID"," I")
            words=line.split(maxsplit=8)
            #pp.pprint(words)
            record_date=convert_date(words[1])
            record_time=words[0].replace("-","")
            record_epoch=get_epoch(record_date+" "+record_time)
            if start_epoch > 0 and ((record_epoch < start_epoch) or (record_epoch > end_epoch)):
                continue
            ws_col=0
            ws.write_string(ws_row,ws_col,record_date)       # AUS Date stamp
            ws_col+=1
            ws.write_string(ws_row,ws_col,record_time)     # Timestamp
            ws_col+=1
            ws.write(ws_row,ws_col,"{:.2f}".format(float(words[2])*1.6))     # Mileage converted to km units
            ws_col+=1
            # Speed, converted to kph and adjusted according to the difference between the real wheel diameter
            # and the diameter reported by the QDP software. NB: The reported wheel diameter can be set when
            # downloading the data via QDP but not when downloading via the QRST software.
            ws.write_number(ws_row,ws_col,round(int(words[3])*1.6*cfg.speed_adjustment_factor))
            ws_col+=1
            ws.write_number(ws_row,ws_col,int(words[4]))        # TMC
            ws_col+=1
            ws.write_number(ws_row,ws_col,int(words[5]))        # Brake pipe pressure
            ws_col+=1
            ws.write_number(ws_row,ws_col,int(words[6]))        # Independent brake pressure
            ws_col+=1
            ws.write(ws_row,ws_col,translate_tp(words[7]))             # Throttle position
            ws_col+=1
            flags=words[-1].replace(" ","")

            #pp.pprint(flags)
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
            for flag in flags:
                ws.write(ws_row,ws_col,"Y" if flag =="1" else "N")             # Throttle position
                ws_col+=1

            ws_row+=1

    wb.close()

def translate_tp(tp):
    ''' take a throtle position. If it's a number then return that number.
        If it's a letter then returnn the corrsponding text '''
    if tp.isnumeric():
        return tp
    if tp.upper() in cfg.tp_translations.keys():
        return cfg.tp_translations[tp.upper()]
    return tp+" (Unknown)"

def convert_date(us_date):
    ''' Convert US formatted date to AUS formatted date '''
    parts=us_date.split("/")
    if len(parts)!=3:
        return "invalid"
    aus_date="{:0>4d}/{:0>2d}/{:0>2d}".format(int(parts[2]),int(parts[0]),int(parts[1]))
    return aus_date

def write_header(wb,ws,loco_name):

    wb.set_size(1920,1080)
    lalign=wb.add_format({'align':'left'})
    ws.set_column('A:B',15,lalign)
    ralign=wb.add_format({'align':'right'})
    ws.set_column('C:H',10,ralign)
    calign=wb.add_format({'align':'center'})
    ws.set_column('I:S',10,calign)

    calign_b=wb.add_format({'align':'center','bold':True})
    ws.set_row(1,None,calign_b)

    header_format=wb.add_format({'font_size':14,'bold':True})

    ws.freeze_panes(3,0)

    ''' Write header line to the worksheet. Return the next row number (0 based) '''
    ws.write(0,0,"Data extract from Quantum Data Recorder : "+loco_name,header_format)
    ws.write(1,0,"Date")
    ws.write(1,1,"Time")
    ws.write(1,2,"Kilometres")
    ws.write(1,3,"Speed (kph)")
    ws.write(1,4,"TMC (A)")
    ws.write(1,5,"ABrk (psi)")
    ws.write(1,6,"IBrk (psi)")
    ws.write(1,7,"Throttle")
    ws.write(1,8,"Reverse")
    ws.write(1,9,"EIE")
    ws.write(1,10,"PCS")
    ws.write(1,11,"Headlight (S)")
    ws.write(1,12,"Forward")
    ws.write(1,13,"Headlight (L)")
    ws.write(1,14,"Horn")
    ws.write(1,15,"DS1")
    ws.write(1,16,"DS2")
    ws.write(1,17,"VC Ack")
    ws.write(1,18,"Axle Drive")
    return 3

def skip_line_found(line):
    '''
        Search for existence of skip_list word(s) in the passed line.
        If found then return True, other wise return False
    '''
    for word in cfg.skip_list_words:
        if word in line:
            return True
    return False


def get_epoch(timestamp):
    ''' Take string in format yyyy/mm/dd hh:mm:ss and return epoch seconds (or 0 if flag is false) '''
    if cfg.between_dates==False:
        return 0
    d = datetime.strptime(timestamp,"%Y/%m/%d %H:%M:%S")
    epoch=datetime(d.year,d.month,d.day,d.hour,d.minute,d.second).timestamp()
    return int(epoch)


if __name__ == '__main__':
    main()
