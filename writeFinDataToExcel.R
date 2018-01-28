writeFinDataToExcel <- function(sub, closePriceI, sdLogReturn, i, wb, ws){
    rowTableStart = 3
  # export to Excel
    lengthData = length(coredata(sub$close))
    #  createStyle(numFmt="NUMBER")
    writeData(wb, ws, closePriceI, startCol = 5, startRow = i+rowTableStart, rowNames = FALSE)
    writeData(wb, ws, sdLogReturn, startCol = 6, startRow = i+rowTableStart, rowNames = FALSE)
    #write retuns 1D, 5D, 23D, 125D, 250D
    createStyle(numFmt="PERCENTAGE")
    writeData(wb, ws, sum(coredata(last(sub$logReturn))), startCol = 12, startRow = i+rowTableStart, rowNames = FALSE)
    writeData(wb, ws, sum(sub$logReturn[(lengthData-4):lengthData]), startCol = 13, startRow = i+rowTableStart, rowNames = FALSE)
    if(lengthData > 125){
      writeData(wb, ws, sum(sub$logReturn[(lengthData-23):lengthData]), startCol = 14, startRow = i+rowTableStart, rowNames = FALSE)
    }
    if(lengthData > 125){
      writeData(wb, ws, sum(sub$logReturn[(lengthData-125):lengthData]), startCol = 15, startRow = i+rowTableStart, rowNames = FALSE)
    }
    if(lengthData > 250){
      writeData(wb, ws, sum(sub$logReturn[(lengthData-250):lengthData]), startCol = 16, startRow = i+rowTableStart, rowNames = FALSE)
    }
  return(0)
}
