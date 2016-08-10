import xlwings as xw

# the original workbook
wb = xw.Book('spike.xlsm')

# the new book to save all the spikes
spbook = xw.Book('spike3.xlsm')

#print(wb.sheets[0].range(3,3).value)
#
#print(wb.sheets[0].cells.last_cell.address)
#print(wb.sheets[0].cells.rows)

#the threshhold to identify a spike
threshhold = 3.0


# column and row No. in wb
wbcol = 1
wbrow = 1

# column and row No. in spbook
spcol = 1
sprow = 2



while wb.sheets[0].range(1, wbcol) is not None:
    #print('+++++++++++++++++++++++++++++++++++++++++++++')
    wbrow = 2
    sprow = 2
    spcol = wbcol
    spbook.sheets[0].range(1,spcol).value = wb.sheets[0].range(1,wbcol)
    spbook.sheets[0].range(1,spcol).color = (0,0,255)
    while True:
        tem1 = wb.sheets[0].range(wbrow, wbcol+1)
        tem2 = wb.sheets[0].range(wbrow + 1, wbcol+1)
        print(tem1.value)
        if tem1.value is None:
            break
        else:
            #print(wbrow)
            if tem2.value - tem1.value >= threshhold:
                spbook.sheets[0].range(sprow, spcol).value = wb.sheets[0].range(wbrow, wbcol).value                
                spbook.sheets[0].range(sprow + 1, spcol).value = wb.sheets[0].range(wbrow + 1, wbcol).value
                
                spbook.sheets[0].range(sprow, spcol+1).value = wb.sheets[0].range(wbrow, wbcol+1).value
                spbook.sheets[0].range(sprow + 1, spcol+1).value = wb.sheets[0].range(wbrow + 1, wbcol+1).value
                sprow = sprow + 2
            wbrow = wbrow + 2
    wbcol += 2