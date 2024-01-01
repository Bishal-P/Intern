import xlrd
import xlwt
import time
from threading import Thread
from googletrans import Translator
translator = Translator()
translator.raise_exception = True


file_path = 'Order_Export.xls'
workbook = xlrd.open_workbook(file_path)
workbook_write = xlwt.Workbook()
sheet_read = workbook.sheet_by_index(0)
sheet_write = workbook_write.add_sheet('sheet1')
num_rows = sheet_read.nrows
num_cols = sheet_read.ncols

def translate_text(text, target_language='en'):
    text=str(text)
    translation = translator.translate(text, dest=target_language)
    
    return translation.text

def translate_excel(row_number):
    for i in range(num_cols):
        a=sheet_read.cell_value(row_number,i)
        if(a==""):
            sheet_write.write(row_number,i,"")
        elif(str(a).isdigit()):
            sheet_write.write(row_number,i,a)
        else:
            retries = 0
            while(retries<30):
                try:
                    sheet_write.write(row_number,i,translate_text(a))
                    break
                except:
                    retries+=1
                    time.sleep(0.1)


if __name__ == "__main__":
    time_start = time.time()
    threads = []
    for i in range(num_rows):
        t = Thread(target=translate_excel, args=(i,))
        threads.append(t)
        t.start()
    for thread in threads:
        thread.join()
    workbook_write.save('translated_excel_sheet.xls')
    time_end = time.time()
    print(f"Time taken to translate and save: {time_end - time_start:.2f} seconds")
    
