# Author: Max Lampe
# Date: 2017-12-23
# Do: download and merge assets from google Finance

#setwd("Documents/Julia")

#getECBCurUp(eDate = Sys.Date(), sDate = Sys.Date() - 364*3)

getECBCurUp <- function(sDate,
                        eDate,
                        loadDataNew=TRUE){
  if(loadDataNew){    ## Currencies
    # EUROCAD
    #  filter1 <- list(startPeriod = "2015-01-10", endPeriod = "2018-01-16", detail = 'full')
    #  EUROCAD <- get_data("EXR.M.CAD.EUR.SP00.E", filter1);
    #  EUROCAD <- xts(x=EUROCAD$obsvalue, order.by = as.Date(EUROCAD$obstime))
    # USEURO
    filter2 <- list(startPeriod = as.character(sDate),
                    endPeriod = as.character(eDate), detail = 'full')
    USEURO <- get_data("EXR.D.USD.EUR.SP00.A", filter2);
    USEURO <- xts(x=USEURO$obsvalue, order.by = as.Date(USEURO$obstime))
    save(USEURO, file = "ECBcurrData.RData")
  } else{
    load(file = "ECBcurrData.RData", verbose=TRUE)
  }
  return(list(USEURO=USEURO))
}

# myGoogleFinData()
myGoogleFinData <- function( eDate = Sys.Date(),
                             sDate = Sys.Date() - 364*3,
                             loadDataNew = TRUE,
                             verbose=TRUE,
                             writeToExcel=TRUE,
                             getECBCur=TRUE,
                             file="Port.xlsx"){
  library(zoo)
  library(xts)
  library(quantmod)
  library(openxlsx)
  library(ecb)

  wb <- loadWorkbook(file = file)
  curr <- c()
  if(getECBCur){
    curr<- getECBCurUp(sDate, eDate)
  }

  # Equity/ETF
  gooFinTicker <- t(read.xlsx(wb, sheet = "Risk", startRow = 2, colNames = FALSE,
                              rowNames = FALSE, detectDates = FALSE, skipEmptyRows = TRUE,
                              skipEmptyCols = TRUE, rows = 4:30, cols = 4))

  assets <- vector("list")
  closePrice <- vector("list")
  dataLoaded <- logical(length = length(gooFinTicker))
  dataLoaded[] <- TRUE

  dataLoad <- logical(length = length(gooFinTicker))
  dataLoad[] <- TRUE
  # dataLoad[c(11, 15)] <- FALSE

  dollarToEuro<- logical(length = length(gooFinTicker))
  dollarToEuro[] <- FALSE
  dollarToEuro[c(1,2,3,5,7,9)] <- TRUE

  if(verbose){
    show("LOADING assets")
    show(gooFinTicker[dataLoad])
  }

  if(loadDataNew){
    if(getECBCur){
      curr <- getECBCurUp(sDate, eDate)
    }

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
        sub$logReturn <- diff(log(sub[,"close"]), lag=1)
        closePrice[i] <- coredata(last(sub$close))
        assets[[i]] <- sub

        # export to excel file
        source("writeFinDataToExcel.R")
        if(writeToExcel){
          writeFinDataToExcel(sub, closePrice[i], sd(sub$logReturn, na.rm = TRUE), i, wb, "Risk");
        }
      }
    }
    if(writeToExcel){
      saveWorkbook(wb, file, overwrite = TRUE)
    }
    save(assets, closePrice, file = "Data/gooFinData.RData")
  }else{
    # loading data from file
    load(file = "Data/gooFinData.RData", verbose=TRUE)
  }

  reMergInfo <- getDataMergInd(assets)
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


# Do: returns the assets indexes to merged based on minimum length and non empty and traded Volume not zero
getDataMergInd  <- function(assets,
                            minLength=350,
                            minNonZeroVolum=300,
                            verbose=TRUE){
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
  return(list(dataMergInd=dataMergInd, dataNotMergInd= dataNotMergInd))
}
