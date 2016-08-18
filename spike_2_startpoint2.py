import xlwings as xw

# the original workbook
wb = xw.Book('spike_2_in_250ms.xlsm')

# the new book to save all the spikes
spbook = xw.Book('spike_2_startpoint2.xlsm')




#print(wb.sheets[0].range(3,3).value)
#
#print(wb.sheets[0].cells.last_cell.address)
#print(wb.sheets[0].cells.rows)

#the threshhold to identify a startpoint of a spike
threshhold = 0.4

#the time unit
tmunit = 0.125


# column and row No. in wb
wbcol = 1
wbrow = 1

# sheets[0] in wb
wbsht = wb.sheets[0]



# column and row No. in spbook
spcol = 1
sprow = 1

#sheets[0] in spbook
spsht = spbook.sheets[0]


# a function to get how many  time units a spike extends to
def spikeNO(row):
    count = 1
    while wbsht.range(row + 1, wbcol).value:
        if wbsht.range(row + 1, wbcol).value - wbsht.range(row, wbcol).value <= tmunit:
            count += 1
            row += 1
        else:
            break
    return count




while spbook.sheets[0].range(1, spcol).value:
    print('+++++++++++++++++++++++++++++++++++++++++++++')
    wbrow = 2
    sprow = 2
    spcol = wbcol

    while True:

        if wbsht.range(wbrow + 1, wbcol).value:
            
            #get the extent of the spike here
            extent = spikeNO(wbrow)
            
            if extent == 1:# to identiy if the spike here is continuous
                
                #if not, the present point is a startpoint
                spsht.range(sprow, spcol).value = wbsht.range(wbrow, wbcol).value
                spsht.range(sprow, spcol + 1).value = wbsht.range(wbrow, wbcol + 1).value
                spsht.range(sprow, spcol).color = wbsht.range(wbrow, wbcol).color
                spsht.range(sprow, spcol + 1).color = wbsht.range(wbrow, wbcol + 1).color                
            elif wbsht.range(wbrow + 1, wbcol + 1).value - wbsht.range(wbrow, wbcol + 1).value >=0.4 :#if continuous, testify if the point here is a startpoint under the threshhold
                # if yes, it's a startpoint
                spsht.range(sprow, spcol).value = wbsht.range(wbrow, wbcol).value
                spsht.range(sprow, spcol + 1).value = wbsht.range(wbrow, wbcol + 1).value
                spsht.range(sprow, spcol).color = wbsht.range(wbrow, wbcol).color
                spsht.range(sprow, spcol + 1).color = wbsht.range(wbrow, wbcol + 1).color                
                
                
            else:
                spsht.range(sprow, spcol).value = wbsht.range(wbrow + 1, wbcol).value
                spsht.range(sprow, spcol + 1).value = wbsht.range(wbrow + 1, wbcol + 1).value
                spsht.range(sprow, spcol).color = wbsht.range(wbrow + 1, wbcol).color
                spsht.range(sprow, spcol + 1).color = wbsht.range(wbrow + 1, wbcol + 1).color
            sprow += 1
        else:
            break
        wbrow += extent

        print(wbrow)
    wbcol += 2

        
        
        
        
