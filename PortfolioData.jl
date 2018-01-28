
using TimeSeries, Plots, Quandl, CSV, Stats, XLSXReader, DataFrames, RCall, ExcelReaders

 df = Dates.DateFormat("y-m-d");
 dataFileName = "2017-12-20 Port.xlsm"
cd("$(homedir())/Documents/Julia/PortOpt")

include("importPortData.jl")
(position, nAssets, posNames) =
portData()

######################################################
 # get data for the last 3 years from googleFinance
######################################################

R"source('writeFinDataToExcel.R')";
R"source('getData.R')";
R"data <- myGoogleFinData(loadDataNew = FALSE)";
include("getDataRToJulia.jl")

include("riskMeasure.jl")
(singleVaR, singleES, reSingleVaRConfInt, reSingleESConfInt, cVar, pVar ) =
getVaRESSingle(logReturns, 0.99, closePrice, position, ticker,
dataMergInd, dataNotMergInd)

pVar * position[dataMergInd]' * closePrice[dataMergInd]
sum(cVar)

portReturns = TimeArray(dates, logReturns, ticker[dataMergInd]);
position = sheetsInfos[:,6]
assSimRet = bootStrap(logReturns, 1000, 394)



(portVaR, portES) = getVaRESPort(assSimRet, level, closePrice[dataMergInd],
                    position[dataMergInd])
(singleVaR, singleES) = getVaRES(assSimRet, 0.99)
expVaRES(portVaR, portES, singleVaR, singleES, dataMergInd)


######################################################
#
######################################################
R"benchHist <- bench$assetsHist"
@rget benchHist
R"benchlogReturns <- bench$lR";
@rget benchlogReturns
R"benchDates <- data$dates";
@rget benchDates
R"benchTickerMerged <- bench$tickerMerged"
@rget benchTickerMerged;


## saving closing prices separatly
closePrice = Array{Float64}(size(ticker,1))
closePrice[:] = 0.0;
b = Array{Float64}(14)

for i = 1:14
  b[i] = assetsHist[i][size(assetsHist[i],1),4]
end
closePrice[dataLoadInd] = b
##
logReturnsNotNa = !isna(logReturns)
logReturnsInd = logReturnsNotNa[:,1] & logReturnsNotNa[:,2]

for i =3:size(logReturnsNotNa,2)
 logReturnsInd = logReturnsInd & logReturnsNotNa[:,i]
end
