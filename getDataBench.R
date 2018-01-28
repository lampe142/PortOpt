# myGoogleFinData()

################################################################################
# Get Benchmark Data
################################################################################

myGoogleFinDataBench <- function( eDate = Sys.Date(),
                                  sDate = Sys.Date() - 364*3,
                                  loadDataNew = TRUE,
                                  verbose=TRUE,
                                  dataMergInd =c(1,2),
                                  writeToExcel=TRUE
){
  
  library(quantmod)
  library(openxlsx)
  library(ecb)
  wb <- loadWorkbook(file = "2018-01-04 Port.xlsm")
  
  # Currencies
  filter2 <- list(startPeriod = "2015-01-10", endPeriod = "2018-01-15", detail = 'full')
  USEURO <- get_data("EXR.D.USD.EUR.SP00.A", filter2);
  USEURO <- xts(x=USEURO$obsvalue, order.by = as.Date(USEURO$obstime))
  
  # Equity/ETF
  gooFinTicker <- t(read.xlsx(wb, sheet = "CurrentPositions", startRow = 2, colNames = FALSE,
                              rowNames = FALSE, detectDates = FALSE, skipEmptyRows = TRUE,
                              skipEmptyCols = TRUE, rows = 23:27, cols = 7))
  
  assets <- vector("list")
  closePrice <- vector("list")
  dataLoaded <- logical(length = length(gooFinTicker))
  dataLoaded[] <- TRUE
  
  dataLoad <- logical(length = length(gooFinTicker))
  dataLoad[] <- TRUE
  
  dollarToEuro<- logical(length = length(gooFinTicker))
  dollarToEuro[] <- FALSE
  dollarToEuro[c(1,2)] <- TRUE
  
  if(verbose){
    show("LOADING assets")
    show(gooFinTicker[dataLoad])
  }
  
  if(loadDataNew){
    for (i in 1:length(gooFinTicker)){
      #  try(
      if(dataLoad[i]){
        show(gooFinTicker[i])
        sub <- getSymbols(gooFinTicker[i], from=sDate, to=eDate,
                          src='google', getSymbols.warning4.0=FALSE,
                          return.class='xts', adjust=TRUE,
                          auto.assign= FALSE)
        colnames(sub) <- c("open", "high", "low","close", "volume")
        # US Dollar to Euro transformation
        if (dollarToEuro[i]){
          trans <- merge(sub, USEURO, join='left')
          ind = which(is.na(trans$USEURO))
          trans$USEURO[ind] = (coredata(trans$USEURO[ind+1])+coredata(trans$USEURO[ind-1]))/2
          sub$open <- sub$open / trans$USEURO
          sub$high <- sub$high / trans$USEURO
          sub$low <- sub$low / trans$USEURO
          sub$close <- sub$close / trans$USEURO
        }
        
        
        sub$logReturn <- diff(log(sub[,"close"]), lag=1)
        closePrice[i] <- coredata(last(sub$close))
        assets[[i]] <- sub
        
        # export to Excel
        if(writeToExcel){
          lengthData = length(coredata(sub$close))
          #  createStyle(numFmt="NUMBER")
          writeData(wb, "CurrentPositions", closePrice[i], startCol = 13, startRow = i+23, rowNames = FALSE)
          
          #    createStyle(numFmt="DATE")
          writeData(wb, "CurrentPositions", index(last(sub$close)), startCol = 12, startRow = i+23, rowNames = FALSE)
          writeData(wb, "CurrentPositions", index(first(sub$close)), startCol = 11, startRow = i+23, rowNames = FALSE)
          #  createStyle(numFmt="NUMBER")
          writeData(wb, "CurrentPositions", lengthData, startCol = 10, startRow = i+23, rowNames = FALSE)
          # write retunds
          createStyle(numFmt="PERCENTAGE")
          writeData(wb, "CurrentPositions", sum(coredata(last(sub$logReturn))), startCol = 15, startRow = i+23, rowNames = FALSE)
          writeData(wb, "CurrentPositions", sum(sub$logReturn[(lengthData-4):lengthData]), startCol = 16, startRow = i+23, rowNames = FALSE)
          if(lengthData > 125){
            writeData(wb, "CurrentPositions", sum(sub$logReturn[(lengthData-23):lengthData]), startCol = 17, startRow = i+23, rowNames = FALSE)
          }
          if(lengthData > 125){
            writeData(wb, "CurrentPositions", sum(sub$logReturn[(lengthData-125):lengthData]), startCol = 18, startRow = i+23, rowNames = FALSE)
          }
          if(lengthData > 250){
            writeData(wb, "CurrentPositions", sum(sub$logReturn[(lengthData-250):lengthData]), startCol = 19, startRow = i+23, rowNames = FALSE)
          }
        }
      }
      if(writeToExcel){
        saveWorkbook(wb, "2018-01-04 Port.xlsm", overwrite = TRUE)
      }
      save(assets, closePrice, file = "gooFinDataBench.RData")
    }
  } else{
    # loading data from file
    load(file = "gooFinDataBench.RData", verbose=TRUE)
  }
  
  # selecting assets
  if(verbose){
    show("ASSETS returns merged:")
    show(gooFinTicker[dataMergInd])
  }
  
  # merging assets for common timeSeries
  logReturns <- assets[[dataMergInd[1]]]$logReturn
  for (i in dataMergInd[2:length(dataMergInd)]){
    logReturns <- merge(logReturns, assets[[i]]$logReturn, join='inner')
  }
  colnames(logReturns) <-  gooFinTicker[dataMergInd]
  logReturns <- na.trim(logReturns, sides = "both")
  
  dataLoadInd <- 1:length(gooFinTicker)
  dataLoadInd <- dataLoadInd[dataLoad]
  
  return(list(closePrice= closePrice,
              assetsHist= assets,
              lR = logReturns,
              ticker = gooFinTicker,
              tickerMerged = gooFinTicker[dataMergInd],
              tickerLoaded = gooFinTicker[dataLoad],
              dataMergInd = dataMergInd,
              dataLoadInd = dataLoadInd,
              dates =time(logReturns)))
}
