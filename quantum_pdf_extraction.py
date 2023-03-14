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
