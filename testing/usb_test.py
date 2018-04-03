from os.path import realpath
import xlsxwriter as xlsx

wb = xlsx.Workbook(realpath("/media/pi/5443-4C2A1/test.xlsx"))
ws = wb.add_worksheet()
wb.close()
