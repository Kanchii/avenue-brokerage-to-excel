import tkinter
import tkinter.filedialog
import re
import pandas as pd
import PyPDF2

class PdfToXLSX():
    def __init__(self):
        filename = tkinter.filedialog.askopenfilename()
        if(filename == None or filename == "" or filename[-3:] != "pdf"):
            raise Exception("You need to select a PDF file")
        self.pdf_filename = filename.split("/")[-1].split(".")[0]

    def convert(self):
        self.__write_pdf_to_csv(self.__extract_text_from_pdf())
        self.__convert_csv_to_xlsx()
    
    def __extract_text_from_pdf(self):
        pdf_file_obj = open(f'PDF/{self.pdf_filename}.pdf', 'rb')
        pdf_reader = PyPDF2.PdfReader(pdf_file_obj)
        text = ""
        for page_num in range(len(pdf_reader.pages)):
            page_obj = pdf_reader.pages[page_num]
            text += page_obj.extract_text()
        pdf_file_obj.close()
        return text

    def __write_pdf_to_csv(self, pdf_text):
        pattern = r"\d+ ([SB]{1}) (\d{2}\/\d{2}\/\d{2,4}) (\d{2}\/\d{2}\/\d{2,4}) (\d+\.\d+) ([A-Z]+) (\d+(?:,\d+)*\.\d+) (\d+(?:,\d+)*\.\d+) (\d+\.\d+)"
        matches = re.findall(pattern, pdf_text)
        if(matches):
            with open(f"CSV/{self.pdf_filename}.csv", "w", encoding="UTF-8") as f:
                f.write("Ticker,Data (DD/MM/YYYY),Tipo da Transação,Quantidade,Preço Unitário,Taxas,Corretora (Utilizar mesmo nome utilizado na plataforma)\n")
                for match in matches:
                    operation, op_date, settle_date, shares, ticker, cost_per_share, total_amount, fee = match
                    s_splitted = settle_date.split("/")
                    settle_date = '/'.join([s_splitted[1],s_splitted[0],s_splitted[2]])
                    f.write(f'{ticker},{settle_date},{"APORTE" if operation == "B" else "RETIRADA"},{shares},{cost_per_share},{fee},AVENUE SECURITIES\n')
        else:
            raise Exception("This script only works for Avenue's brokerage statement")

    def __convert_csv_to_xlsx(self):
        df = pd.read_csv(f'CSV/{self.pdf_filename}.csv', sep=',')
        df.to_excel(f'XLSX/{self.pdf_filename}.xlsx', index=False)

if __name__ == "__main__":
    PdfToXLSX().convert()
    print("PDF converted to XLSX successfully!")