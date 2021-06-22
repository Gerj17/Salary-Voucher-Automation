import win32com.client
from pathlib import Path as p
import copy

"""
This file opens an Excel work book, takes the indexed sheets and convert them to a pdf keeping the formatting.

pdf_name : The name of the pdf file when printed(important)

"""

def ExcelPrintPdf(abs_wb_path, pdf_name, ):
    """

    :param wb_path: Absolute Path to the workbook (prefer Pathlib)
    :param pdf_name: Name of PDF document
    :param pdf_save_loc: Location to dave pdf document
    :output PDF version of excel sheet with according to sheet print formatting
    :return:
    """
    o = win32com.client.gencache.EnsureDispatch("Excel.Application")
    #o = win32com.client.Dispatch("Excel.Application")
    o.Visible = False

        # Used for testing puposes
        #pdf_name = "p.pdf"
        #wb_path = r"C:\Users\Gerard\OneDrive\Desktop\Ideal\Ideal Bakery Wage Payment Voucher copy.xlsx"
        #wb_path = p.cwd().joinpath(r"Ideal Bakery Wage Payment Voucher copy.xlsx")
        #path_to_pdf = r'C:\\Users\\Gerard\\OneDrive\Desktop\\Ideal\\' + pdf_name

   # pdf_name = str(pdf_name +'.pdf')
    pdf_name = str(pdf_name)
    wb_path = str(abs_wb_path)
    path_to_pdf = str(p.cwd().joinpath('PDFS',pdf_name))

    #print('path_to_pdf :',path_to_pdf)
    wb = o.Workbooks.Open(wb_path)

    ws_index_list = [0]  # say you want to print these sheets Starts from 1

    #wb.WorkSheets(ws_index_list).Select()  # selects items from index and concatenates into single print file

    #wb.ActiveSheet.ExportAsFixedFormat(0, "C:\\Users\\Gerard\\OneDrive\\Desktop\\Ideal\\test - Copy.pdf")
    try:
        wb.ActiveSheet.ExportAsFixedFormat(0, str(path_to_pdf))
    except :
        print('Create Folder named __ PDFS__')
        wb.ActiveSheet.ExportAsFixedFormat(0, str(str(p.cwd().joinpath(pdf_name))))
    finally:
        wb.Close()

#ExcelPrintPdf(p.cwd().joinpath('Ideal Bakery Wage Payment Voucher copy.xlsx'),'hello.pdf',)