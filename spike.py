import xlwings as xw

# the original workbook
wb = xw.Book('xiaoyu_proj.xlsm')

# the new book to save all the spikes
spbook = xw.Book('spike.xlsm')

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



while wb.sheets[0].range(1, wbcol) is not None:
    #print('+++++++++++++++++++++++++++++++++++++++++++++')
    wbrow = 1
    sprow = 2
    spcol = wbcol*2 - 1
    cellname = 'CELL'
    cellname = cellname + str(wbcol)
    wbcol = wbcol + 1
    spbook.sheets[0].range(1,spcol).value = cellname
    spbook.sheets[0].range(1,spcol).color = (0,0,255)
    while True:
        tem1 = wb.sheets[0].range(wbrow, wbcol)
        tem2 = wb.sheets[0].range(wbrow + 1, wbcol)
        if tem2.value is None:
            break
        else:
            #print(wbrow)
            if tem2.value - tem1.value >= threshhold:
                spbook.sheets[0].range(sprow, spcol).value = wb.sheets[0].range(wbrow, 1).value                
                spbook.sheets[0].range(sprow + 1, spcol).value = wb.sheets[0].range(wbrow + 1, 1).value
                
                spbook.sheets[0].range(sprow, spcol+1).value = wb.sheets[0].range(wbrow, wbcol).value
                spbook.sheets[0].range(sprow + 1, spcol+1).value = wb.sheets[0].range(wbrow + 1, wbcol).value
                sprow = sprow + 2
            wbrow = wbrow + 1


        
        
        
        
