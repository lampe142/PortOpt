
# get data drom R to Julia
# Financial Data
 R"logReturns <- data$lR";
 @rget logReturns
 R"dates <- data$dates";
 @rget dates
 R"assetsHist <- data$assetsHist"
 @rget assetsHist
 R"closePrice <- data$closePrice"
 @rget closePrice
 closePrice = convert(Array{Float64,1}, closePrice)


 # Ticker
 # R"tickerLoaded <- data$tickerLoaded"
 # @rget tickersLoaded;
 R"tickerMerged <- data$tickerMerged"
@rget tickerMerged;
R"ticker <- data$ticker"
@rget ticker
R"dataMergInd <- data$dataMergInd"
@rget dataMergInd
dataMergInd = vec(dataMergInd')
R"dataNotMergInd <- data$dataNotMergInd"
@rget dataNotMergInd
dataNotMergInd = vec(dataNotMergInd)
R"dataLoadInd <- data$dataLoadInd"
@rget dataLoadInd
