#this is used to find the column NO of the true cells

import xlwings as xw

# the original workbook
wb = xw.Book('cellNO.xlsm')

# the new book to save all the spikes

#print(wb.sheets[0].range(3,3).value)
#
#print(wb.sheets[0].cells.last_cell.address)
#print(wb.sheets[0].cells.rows)

#the threshhold to identify a spike
threshhold = 2.5


# column and row No. in wb
wbcol = 1
wbrow = 1

# column and row No. in spbook
spcol = 1
sprow = 2

#array to save the column NO. of the true cells
cells = []

i = 0

while True:
    wbrow += 1
    tem1 = wb.sheets[1].range(wbrow, 1)
    if tem1.value is None:
        break
    else:
        #print(wbrow)
        cells.append(int(tem1.value))
print(cells)
        


        
        
        
        
