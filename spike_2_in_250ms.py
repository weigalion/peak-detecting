import xlwings as xw
import cell_NO

# the original workbook
wb = xw.Book('xiaoyu_proj.xlsm')

# the new book to save all the spikes
spbook = xw.Book('spike_2_in_250ms.xlsm')

# the array to save the column NO of the true cells
cellcol = cell_NO.cell_NO

#print(cellcol)

#print(wb.sheets[0].range(3,3).value)
#
#print(wb.sheets[0].cells.last_cell.address)
#print(wb.sheets[0].cells.rows)

#the threshhold to identify a spike
threshhold = 2

# index of the cellcol array
index = 39

# column and row No. in wb
wbcol = cellcol[index]
wbrow = 1




# column and row No. in spbook
spcol = 1
sprow = 2



while index < len(cellcol):
    print('+++++++++++++++++++++++++++++++++++++++++++++')
    wbrow = 1
    sprow = 2
    spcol = index*2 + 1
    cellname = 'CELL'
    cellname = cellname + str(cellcol[index])
    wbcol = cellcol[index] + 1
    spbook.sheets[0].range(1,spcol).value = cellname
    spbook.sheets[0].range(1,spcol).color = (100,100,100)
    while True:
        tem1 = wb.sheets[0].range(wbrow, wbcol)
        tem2 = wb.sheets[0].range(wbrow + 1, wbcol)
        if tem2.value:
            if tem2.value - tem1.value >= threshhold:
                spbook.sheets[0].range(sprow, spcol).value = wb.sheets[0].range(wbrow, 1).value
                spbook.sheets[0].range(sprow, spcol).color = (255, 100,100)
                #spbook.sheets[0].range(sprow + 1, spcol).value = wb.sheets[0].range(wbrow + 1, 1).value
                
                spbook.sheets[0].range(sprow, spcol+1).value = wb.sheets[0].range(wbrow, wbcol).value
                spbook.sheets[0].range(sprow, spcol+1).color = (250, 100, 100)
                #spbook.sheets[0].range(sprow + 1, spcol+1).value = wb.sheets[0].range(wbrow + 1, wbcol).value
                sprow += 1
            else:
                #print(wbrow)
                tem3 = wb.sheets[0].range(wbrow + 2, wbcol)
                if tem3.value:
                    if tem3.value - tem1.value >= threshhold:
                        spbook.sheets[0].range(sprow, spcol).value = wb.sheets[0].range(wbrow, 1).value                
                        spbook.sheets[0].range(sprow, spcol).color = (100, 100, 255)                
                        spbook.sheets[0].range(sprow, spcol+1).value = wb.sheets[0].range(wbrow, wbcol).value
                        spbook.sheets[0].range(sprow, spcol+1).color = (100, 100, 255)
                        sprow += 1
                else:
                    break
        else:
            break
        wbrow = wbrow + 1
        print(wbrow)
    index += 1

        
        
        
        
