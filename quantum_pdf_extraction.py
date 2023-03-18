'''

Quantum Desktop Playback - data reporter

This code will parse a PDF file and create an Excel worksheet.
The PDF file is created by running the QDP software, selecting the timescale
required using tags (or select the entire file) then printing it to a PDF file

When setting up the printed page, select ALL variables logged - this code will weed out those
that are not required.

Th eheader page of the PDF report is used ot make adjustments to the wheel speed from that
recorded in the data (which can be modified if the data is extracted using QDP but cannot if
it's downloaded using the QRST code).

All logger data is in imperial units so we need to convert to metric.

The cfg file contains all the data required to drive the application operation.




'''


import pprint
from pypdf import PdfReader
import xlsxwriter
import quantum_pdf_extraction_cfg as cfg

pp=pprint.PrettyPrinter(indent=4)

def main():

    loco_name=""
    wb=xlsxwriter.Workbook(cfg.workbook_name)
    ws=wb.add_worksheet(cfg.worksheet_name)
    ws_ann=wb.add_worksheet(("Events"))
    ws_row=write_header(wb,ws)
    ws_col=0
    ws_row_ann=0

    reader=PdfReader(cfg.source_file)
    pages=len(reader.pages)
    # Iterate through each page
    for page in range(pages):
        print("Processing page "+str(page))
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
                    loco_fmt=wb.add_format({'font_size':12,'bold':True})
                    ws.write(ws_row,0,"Locomotive - "+loco_name,loco_fmt)
                    ws_row+=1
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
            words=line.split(maxsplit=8)
            #pp.pprint(words)
            ws_col=0
            ws.write_string(ws_row,ws_col,convert_date(words[1]))       # AUS Date stamp
            ws_col+=1
            ws.write_string(ws_row,ws_col,words[0].replace("-",""))     # Timestamp
            ws_col+=1
            ws.write_string(ws_row,ws_col,"{:.2f}".format(float(words[2])*1.6))     # Mileage converteed to km units
            ws_col+=1
            if int(words[3])==0:    # Speed, convert to kph and apply wheel diameter adjustment factor
                x=0
            else:
                x=round(int(words[3])*1.6*cfg.speed_adjustment_factor)
            ws.write_number(ws_row,ws_col,x)
            ws_col+=1
            ws.write_number(ws_row,ws_col,int(words[4]))        # TMC
            ws_col+=1
            ws.write_number(ws_row,ws_col,int(words[5]))        # Brake pipe pressure
            ws_col+=1

            # This gets messy, some records show the Throttle Position as 'ID'
            # When that happens, the PDF reder concatenates that value to the inpendent brake pressure
            # (eg. "53ID") and this then stuffs up the parsing of the following flags as they end up
            # in 2 variables, not 1.
            # So if this is the case we need to split word[6] into the IBRK and TP components AND
            # we need to amalgamate the 2 flag vars onto 1
            if len(words[6])>2 and words[6][-2:]=="ID":
                # This is one of the wierd cases
                ibrk=words[6][:-2]
                tp=words[6][-2:]
                ws.write_number(ws_row,ws_col,int(ibrk))        # Independent brake pressure
                ws_col+=1
                ws.write(ws_row,ws_col,tp)             # Throttle position
                ws_col+=1
                flags=words[7].replace(" ","") + words[-1].replace(" ","")

            else:
                ws.write_number(ws_row,ws_col,int(words[6]))        # Independent brake pressure
                ws_col+=1
                ws.write(ws_row,ws_col,words[7])             # Throttle position
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



def convert_date(us_date):
    ''' Convert US formatted date to AUS formatted date '''
    parts=us_date.split("/")
    if len(parts)!=3:
        return "invalid"
    aus_date="{:0>4d}/{:0>2d}/{:0>2d}".format(int(parts[2]),int(parts[0]),int(parts[1]))
    return aus_date

def write_header(wb,ws):

    wb.set_size(1920,1080)
    lalign=wb.add_format({'align':'left'})
    ws.set_column('A:C',15,lalign)
    ralign=wb.add_format({'align':'right'})
    ws.set_column('D:H',10,ralign)
    calign=wb.add_format({'align':'center'})
    ws.set_column('I:S',10,calign)

    calign_b=wb.add_format({'align':'center','bold':True})
    ws.set_row(1,None,calign_b)

    header_format=wb.add_format({'font_size':14,'bold':True})

    ws.freeze_panes(4,0)

    ''' Write header line to the worksheet. Return the next row number (0 based) '''
    ws.write(0,0,"Data extract from Quantum Data Recorder",header_format)
    ws.write(1,0,"Date")
    ws.write(1,1,"Time")
    ws.write(1,2,"Kilometres")
    ws.write(1,3,"Speed (kph)")
    ws.write(1,4,"TMC")
    ws.write(1,5,"ABrk")
    ws.write(1,6,"IBrk")
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

if __name__ == '__main__':
    main()
