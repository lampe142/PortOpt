# Author: Max Lampe
# Date: 2017-12-23
# Do: download and merge assets from google Finance

#setwd("Documents/Julia/PortOpt")

#getECBCurUp(eDate = Sys.Date(), sDate = Sys.Date() - 364*3)
paVaR <- function(LR, p=0.99, i, wb, ws){
  rowTableStart <- 2
  VaRGaus <- -PerformanceAnalytics:::VaR.Gaussian(R=LR[!is.na(LR)], p=0.99)
  VaRHS <- -PerformanceAnalytics:::VaR.historical(R=LR[!is.na(LR)], p=0.99)
  writeData(wb, ws, VaRGaus, startCol = 16, startRow = i+rowTableStart, rowNames = F, colNames = F)
  writeData(wb, ws, VaRHS, startCol = 17, startRow = i+rowTableStart, rowNames = F, colNames = F)
}



# library("openxlsx")
# b <- getECBCurUp(eDate = Sys.Date(),sDate = Sys.Date() - 364*3)
getECBCurUp <- function(sDate,
                        eDate,
                        loadDataNew=T){



  if(loadDataNew){    ## Currencies
    # EUROCAD
    #  filter1 <- list(startPeriod = "2015-01-10", endPeriod = "2018-01-16", detail = 'full')
    #  EUROCAD <- get_data("EXR.M.CAD.EUR.SP00.E", filter1);
    #  EUROCAD <- xts(x=EUROCAD$obsvalue, order.by = as.Date(EUROCAD$obstime))
    # USEURO
    filter2 <- list(startPeriod = as.character(sDate),
                    endPeriod = as.character(eDate), detail = 'full')
    # US Dollar for many assets
    USEURO <- get_data("EXR.D.USD.EUR.SP00.A", filter2);
    USEURO <- xts(x=USEURO$obsvalue, order.by = as.Date(USEURO$obstime))
    # Hong Kong Dollar for Indonesia ETF position
    HKDEURO <- get_data("EXR.D.HKD.EUR.SP00.A", filter2);
    HKDEURO <- xts(x=HKDEURO$obsvalue, order.by = as.Date(HKDEURO$obstime))


    save(USEURO, file = "ECBcurrData.RData")
  } else{
    load(file = "ECBcurrData.RData", verbose=TRUE)
  }
  return(list(USEURO=USEURO, HKDEURO=HKDEURO))
}

# corrects for in differently quoted currencies
curAdj <- function(curr, sub){
    trans <- merge(sub, curr, join='left')
    ind = which(is.na(trans$curr))
    while(ind[length(ind)] == length(trans$curr)){
      ind = ind[1:(length(ind)-1)]
      trans$curr[1:(length(trans$curr)-1)]
    }
    trans$curr.USEURO[ind] = (coredata(trans$curr[ind+1])+coredata(trans$curr[ind-1]))/2
    sub$open <- sub$open / trans$curr
    sub$high <- sub$high / trans$curr
    sub$low <- sub$low / trans$curr
    sub$close <- sub$close / trans$curr
  return(sub)
}


# myGoogleFinData()
myGoogleFinData <- function( eDate = Sys.Date(),
                             sDate = Sys.Date() - 364*3,
                             loadDataNew = T,
                             verbose=T,
                             writeToExcel=T,
                             getECBCur=T,
                             perfAnaly=T,
                             openXLSXFile=T,
                             nAssets=30,
                             file="Port.xlsx",
                             indexDollarToEuro){
  library(zoo)
  library(xts)
  library(quantmod)
  library(openxlsx)
  library(ecb)

  wb <- loadWorkbook(file = file)
  curr <- c()
# If ECB - Euro Dollar information is active
  if(getECBCur){
    curr<- getECBCurUp(sDate, eDate, loadDataNew=loadDataNew)
  }

  # get Assets ticker for google finance
  gooFinTicker <- t(read.xlsx(wb, sheet = "Risk", startRow = 3, colNames = F,
                              rowNames = F, detectDates = F, skipEmptyRows = T,
                              skipEmptyCols = T, rows = 3:(nAssets+2), cols = 3))

  assets <- vector("list")
  closePrice <- vector("list")
  dataLoaded <- logical(length = length(gooFinTicker))
  dataLoaded[] <- T

  dataLoad <- logical(length = length(gooFinTicker))
  dataLoad[] <- T

  dollarToEuro <- logical(length = length(gooFinTicker))
  dollarToEuro[] <- F
  HKDToEuro <- logical(length = length(gooFinTicker))
  HKDToEuro[] <- F
  HKDToEuro[3] <- F
  indexDollarToEuro
  dollarToEuro[indexDollarToEuro] <- T

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
                          src='google', getSymbols.warning4.0=F,
                          return.class='xts', adjust=T,
                          auto.assign= F)
        colnames(sub) <- c("open", "high", "low","close", "volume")
        # US Dollar to Euro transformation
        if (getECBCur && dollarToEuro[i]){
          trans <- merge(sub, curr$USEURO, join='left')
          ind = which(is.na(trans$curr.USEURO))
          while(ind[length(ind)] == length(trans$curr.USEURO)){
            ind = ind[1:(length(ind)-1)]
            trans$curr.USEURO[1:(length(trans$curr.USEURO)-1)]
          }
          trans$curr.USEURO[ind] = (coredata(trans$curr.USEURO[ind+1])+coredata(trans$curr.USEURO[ind-1]))/2
          sub$open <- sub$open / trans$curr.USEURO
          sub$high <- sub$high / trans$curr.USEURO
          sub$low <- sub$low / trans$curr.USEURO
          sub$close <- sub$close / trans$curr.USEURO
        }
        if (getECBCur && HKDToEuro[i]){
          sub <- curAdj(curr ,sub)
        }
        sub$logReturn <- diff(log(sub[,"close"]), lag=1)
        closePrice[i] <- coredata(last(sub$close))
        assets[[i]] <- sub
        paVaR(LR= sub$logReturn, p=0.99, i=i, wb=wb, ws='Risk')
        # export to excel file
        source("writeFinDataToExcel.R")
        if(writeToExcel){
          writeFinDataToExcel(sub, closePrice[i], sd(sub$logReturn, na.rm = T), i, wb, "Risk");
        }
      }
    }

    if(writeToExcel){
      saveWorkbook(wb, file, overwrite = T)
    }
    save(assets, closePrice, file = "Data/gooFinData.RData")
  }else{
    # loading data from file
    load(file = "Data/gooFinData.RData", verbose=T)
  }

  reMergInfo <- getDataMergInd(assets, wb=wb, ws= 'Risk',file=file)
  dataMergInd <- reMergInfo$dataMergInd
  dataNotMergInd <- reMergInfo$dataNotMergInd

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

if(openXLSXFile){
  system(paste("open", file))
}

  return(list(closePrice= closePrice,
              assetsHist= assets,
              lR = logReturns,
              ticker = gooFinTicker,
              tickerMerged = gooFinTicker[dataMergInd],
              tickerLoaded = gooFinTicker[dataLoad],
              dataMergInd = dataMergInd,
              dataNotMergInd = dataNotMergInd,
              dataLoadInd = dataLoadInd,
              dates =time(logReturns)))
}


# Do: returns the assets indexes to merged based on minimum length and non empty
# and traded Volume not zero
getDataMergInd  <- function(assets, wb, ws, file,
                            minLength=350,
                            minNonZeroVolum=300,
                            verbose=T,
                            writeToExcel=T){
  dataMergInd <- c()
  dataNotMergInd <- c()
  for (i in 1:length(assets)){
  if(length(assets[[i]]$logReturn) >= minLength
  & sum(!is.na(assets[[i]]$logReturn)) >= minLength
  & (sum(!(assets[[i]]$volume==0)) >= minNonZeroVolum)){
      dataMergInd <- cbind(dataMergInd, i)
    } else {
      dataNotMergInd <- cbind(dataNotMergInd, i)
    }
  }
  if(verbose){
  show(dataMergInd)
  show(dataNotMergInd)
  }
  # exporting the merged indicator into Row 3 and Col 22
  if(writeToExcel){
    dataMergIndPrint <- rep(0,length(assets))
    dataMergIndPrint[dataMergInd] = 1
    writeData(wb, ws, dataMergIndPrint, startCol = 22, startRow = 3, rowNames = FALSE)
    saveWorkbook(wb, file, overwrite = TRUE)
  }

  return(list(dataMergInd=dataMergInd, dataNotMergInd= dataNotMergInd))
}
