import xlwings as xw
from pathlib import Path

filename = Path("TemizlikRobotuDeneme.xlsx")
wb = xw.Book(filename)
sht = wb.sheets["Sayfa1"]
myrange = sht.range("B2:G16")

fs_count_row = 14
fs_count_col = 5

for row in range(0,fs_count_row+1):
    for column in range(fs_count_col, -1, -1):
        cell_color = myrange[row,column].color

        if(cell_color == (255,255,255) or cell_color == (255,0,0)):
            continue
        elif(cell_color == (255,217,102)):
            print("There is a trash in indeks:[{},{}]".format(row,column))
            myrange[row,column].color = (255,255,255)
