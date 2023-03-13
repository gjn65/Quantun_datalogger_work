import pprint
import PyPDF2
from pypdf import PdfReader

pp=pprint.PrettyPrinter(indent=4)

def main():
    ''' pdf_obj = open('test1.pdf','rb')
    pdf_reader = PyPDF2.PdfReader(pdf_obj)
    print(len(pdf_reader.pages))
    page_obj = pdf_reader.pages[0]
    print(page_obj.extract_text())
    pdf_obj.close()
    '''

    reader=PdfReader('test1.pdf')
    print(len(reader.pages))
    pages=len(reader.pages)
    for page in range(pages):
        print("Processing page "+str(page))
        lines=reader.pages[page].extract_text().split('\n')
        print(lines)



if __name__ == '__main__':
    main()
