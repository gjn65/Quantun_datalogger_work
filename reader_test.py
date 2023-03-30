import pdfplumber
import pprint


pp=pprint.PrettyPrinter(indent=4)

#with pdfplumber.open(r'844 high tmc.pdf') as pdf:

pdf=pdfplumber.open('844 high tmc.pdf')
for page in range(len(pdf.pages)):
    print("Page : "+str(page))
    page_contents=pdf.pages[page]
    print("DEBUG START")
    pp.pprint(page_contents)
    x=page_contents.extract_text()
    pp.pprint(x)
    lines=x.split('\n')
    pp.pprint(lines)
    for line in lines:
        print(line)
    print("DEBUG END")
      #  print(page_contents.extract_text())