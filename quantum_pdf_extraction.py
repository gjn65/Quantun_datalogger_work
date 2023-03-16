import pprint
from pypdf import PdfReader
import quantum_pdf_extraction_cfg as cfg

pp=pprint.PrettyPrinter(indent=4)

def main():

    reader=PdfReader(cfg.source_file)
    print(len(reader.pages))
    pages=len(reader.pages)
    # Iterate through each page
    for page in range(pages):
        print("Processing page "+str(page))
        lines=reader.pages[page].extract_text().split('\n')
        # Iterate through each line for this page
        for line in lines:
            print(line)
            print(str(len(line)))
            if len(line) == 0:
                continue

            # Search for wheel diameter figure in 1st page, but only if adjustment factor has not already been established
            if page == 0:
                if cfg.speed_adjustment_factor == 0:
                    if "Circumference" in line and "Diameter" in line:
                        words=line.split()
                        pp.pprint(words)
                        wheel_diameter_qdp_inches=float(words[-1])    # wheel diameter according to the QDP software
                        cfg.speed_adjustment_factor = cfg.wheel_dia_actual_mm / (wheel_diameter_qdp_inches*25.4)
                        pp.pprint(cfg.speed_adjustment_factor)
                        continue
                continue        # We don't want anything else from page 0

            # Skip lines with strings we are not interested in
            if skip_line_found(line) == True:
                continue

            # If 1st character in the line is non numeric we treat the line as an annotation
            # for example - recorder power up, laptop connection etc.
            if line[0].isnumeric() == False:
                # Handle annotations
                print("Annotataion : "+line)
                continue

            record=dict()
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
            pp.pprint(words)
            record["time"]=words[0].replace("-","")
            record["date"]=convert_date(words[1])
            record["miles"]=words[2]
            record["speed"]=words[3]
            record["tmc"]=words[4]
            record["abrk"]=words[5]
            record["ibrk"]=words[6]
            record["throttle"]=words[7]

            flags=words[-1].replace(" ","")
            pp.pprint(flags)
            # Flags are either 1 or 0 for ON or OFF
            # Reverser in reverse
            # Engineer induced emergency - NOT USED
            # Preesure control switch (set when the BP air drops below 45 psi)
            # Headlight on - short end
            # Reverser in forward
            # Headlight on - long end
            # Horn on
            # Digital spare 1 - NOT USED
            # Digital spare 2 - NOT USED
            # Vigilance Control Alert acknowledge - NOT USED
            # Ad Ty - NOT USED
            record["rev"]=flags[0]
            record["pcs"]=flags[2]
            record["head_short"]=flags[3]
            record["fwd"]=flags[4]
            record["head_long"]=flags[5]
            record["horn"]=flags[6]




def convert_date(us_date):
    ''' Convert US formatted date to AUS formatted date '''
    parts=us_date.split("/")
    if len(parts)!=3:
        return "invalid"
    aus_date=parts[1]+"/"+parts[0]+"/"+parts[2]
    return aus_date



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
