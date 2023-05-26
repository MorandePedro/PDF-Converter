import aspose.pdf as ap
import os


for filename in os.listdir("..\\PDFs"):

    filename = filename.split('.')[0]
    input = f"..\\PDFs\\{filename}.pdf"
    output = f"..\\EXCEL\\{filename}.xlsx"

    document = ap.Document(input)

    save_option = ap.ExcelSaveOptions()

    document.save(output, save_option)
    print("**********************\n")
    print(f'{filename}.pdf CONVERTIDO!\n')