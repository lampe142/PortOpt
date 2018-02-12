
workspace()
using TimeSeries, Plots, Quandl, CSV, Stats, XLSXReader, DataFrames, RCall
using RCall, Bootstrap
using ExcelReaders

 df = Dates.DateFormat("y-m-d");
 dataFileName = "2017-12-20 Port.xlsm"
cd("$(homedir())/Documents/Julia/PortOpt")

global level, position, closePrice, dataMergInd
include("importPortData.jl")
(position, nAssets, posNames) =
portData(30)

######################################################
 # get data for the last 3 years from googleFinance
######################################################

R"source('writeFinDataToExcel.R')";
R"source('getData.R')";
R"data <- myGoogleFinData(loadDataNew = F, nAssets=30)";
include("getDataRToJulia.jl")

include("riskMeasure.jl")
(singleVaR, singleES, reSingleVaRConfInt, reSingleESConfInt, cVar, pVar ) =
getVaRESSingle(logReturns, 0.99, closePrice, position, ticker,
dataMergInd, dataNotMergInd)


######################################################
# test case
######################################################
getVaRES(assetsHist, logReturns, dataMergInd,
         dataNotMergInd,expToExcel = false)

include("riskMeasure.jl")
test = zeros(Float64, (size(logReturns,1),2))
test[:,1] = logReturns[:,1]
test[:,2] = logReturns[:,2]
#test[:,2] = logReturns[:,2]
closePrice = [100.0; 100.0]
position = [1; 1]
ticker = ticker[[1,2]]
dataMergInd = [1,2]
dataNotMergInd = [0]

(a, b, c, d, cVar2, portfolioVarTest) =
getVaRESSingle(test, 0.99, closePrice, position, ticker,
dataMergInd, dataNotMergInd, expToExcel = false, nBootstrap=2000,
nSim =1000)

sum(a .* closePrice)
sum(cVar2)
portfolioVarTest
######################################################
# test case
######################################################




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
