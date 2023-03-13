import pprint
import PyPDF2
from pypdf import PdfReader

pp=pprint.PrettyPrinter(indent=4)

def main():

    reader=PdfReader('test1.pdf')
    print(len(reader.pages))
    pages=len(reader.pages)
    for page in range(pages):
        print("Processing page "+str(page))
        lines=reader.pages[page].extract_text().split('\n')
        print(lines)



if __name__ == '__main__':
    main()
