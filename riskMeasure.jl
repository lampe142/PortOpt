# Author: Max Lampe
# Date: 2017-12-23
# Do: compute risk measures for a portfolio

################################################################################
## computre a single bootstrap
################################################################################
function bootStrapSample(logReturns,
                nSim = 1000,
                current= size(logReturns,1),
                nLookback=250,
                ndays=1)
 samTimLength = size(logReturns,1)
 assSimRet = zeros(Float64, (nSim, size(logReturns,2)))
 current = size(logReturns,1)
 nLookback = size(logReturns,1) -1
 # looping of the time length of the simulation (i.e. 1, 5, 10 day)
 for i = 1:ndays
  selBoot = rand((current-nLookback):(current-1), nSim)
  # looping of the difference
  for j = 1:nSim
   assSimRet[j,:] = assSimRet[j,:] + logReturns[selBoot[j],:]
  end
 end
 return assSimRet
end

################################################################################
##
################################################################################
function  getVaRESSingle(logReturns, level,
                         closePrice,
                         position,
                         ticker,
                         dataMergInd,
                         dataNotMergInd,
                         confInt = true,
                         expToExcel = true)
 nAssets = size(dataMergInd,1)
 nTotalAssets = size(closePrice)
 nBootstrap = 500
 singleVaR = zeros(Float64, (nAssets, nBootstrap))
 CVaR = zeros(Float64, (nBootstrap, nAssets))
 singleES = zeros(Float64, (nAssets, nBootstrap))
 portfolioVaR = 0.0;
 portfolioES = 0.0;
 nReturns = 1000;
 confIntValues = [0.05, 0.95]
 portVaR = zeros(Float64, (1, nBootstrap))
 portES = zeros(Float64, (1, nBootstrap))
 reSingleVaR = zeros(Float64, size(closePrice,1))
 reSingleES = zeros(Float64, size(closePrice,1))

## VaR computation for the assets with sufficient data
for j =1:nBootstrap
 assSimRet = bootStrapSample(logReturns, nReturns, 394)
 # single assets
 for i=1:nAssets
  sortReturn = sort(assSimRet[:,i])
  singleVaR[i, j] = quantile(sortReturn, (1-level), sorted=true)
  singleES[i, j] = mean(sortReturn[sortReturn .< singleVaR[i, j]])
 end
 # portfolio VaR and ES and CVaR
 (portVaRj, b, CVaRj) =
 getVaRESPortCVaR(assSimRet, level, closePrice, position, dataMergInd)
 portVaR[j] = portVaRj
 portES[j] = b
 CVaR[j,:] = CVaRj
end
 resCVaR = zeros(Float64, nTotalAssets)
 reSingleVaR[dataMergInd] = map(x -> mean(singleVaR[x,:]), 1:size(singleVaR,1))
 reSingleES[dataMergInd] = map(x -> mean(singleES[x,:]), 1:size(singleES,1))
 resCVaR[dataMergInd] = map(x -> mean(CVaR[:,x]), 1:size(CVaRj,1))

## VaR for the assets with short time series
VaRNotMerg = zeros(Float64, (nBootstrap , size(dataNotMergInd,2)))
ESNotMerg = zeros(Float64, (nBootstrap , size(dataNotMergInd,2)))
for i in 1:size(dataNotMergInd,2)
 for j=1:nBootstrap
  bootSamNonMerg = vec(bootStrapSample(assetsHist[dataNotMergInd[i]][2:end,6]))
  VaRNotMerg[j,i] = quantile(bootSamNonMerg, (1-level))
  ESNotMerg[j,i] = mean(bootSamNonMerg[bootSamNonMerg .< VaRNotMerg[j,i]])
 end
end
cc = zeros(Float64, size(dataNotMergInd,2))
cc = map(x -> mean(VaRNotMerg[:,x]), 1:size(VaRNotMerg,2))
reSingleES = map(x -> mean(singleES[x,:]), 1:size(singleES,1))
resCVaR[dataNotMergInd] = cc .* closePrice[dataNotMergInd]' .* position[dataNotMergInd]'


# compute best estimate over bootStraps

# compute confidence Intervals over bootStraps
if(confInt)
 reSingleVaRConfInt = zeros(Float64, (size(dataMergInd,1), 2))
 reSingleESConfInt = zeros(Float64, (size(dataMergInd,1), 2))
 for i = 1:size(singleVaR,1)
  reSingleVaRConfInt[i,:] = quantile(singleVaR[i,:], confIntValues)
  reSingleESConfInt[i,:] = quantile(singleVaR[i,:], confIntValues)
 end

 if(expToExcel)
  expVaRES(mean(portVaR),
           mean(portES),
           reSingleVaR,
           reSingleES,
           dataMergInd,
           ticker,
           resCVaR)
          # reSingleVaRConfInt,
          # reSingleESConfInt,

 end
 return(reSingleVaR, reSingleES, reSingleVaRConfInt, reSingleESConfInt, resCVaR, mean(portVaR))
end
return(reSingleVaR, reSingleES)
end

################################################################################
## Value at Risk and Expected Shortflall for the Portfolio
################################################################################

#(portVaR, portES, CVaRj) = getVaRESPortCVaR(assSimRet, 0.9,
#closePrice[dataMergInd], assVolums[dataMergInd])
#position = assVolums[dataMergInd]
#closePriceMer = closePrice[dataMergInd]

function  getVaRESPortCVaR(assSimRet, level, closePrice, position, dataMergInd)
 portValue = closePrice[dataMergInd]' * position[dataMergInd]
 posValue = closePrice[dataMergInd] .* position[dataMergInd]
 assSimRetPosValue = zeros(Float64, size(assSimRet,1))
 CVaRj = zeros(Float64, size(assSimRet,2))

 portSimRetSort = sort(vec(assSimRet * posValue / portValue))

 portVaR = quantile(portSimRetSort, (1-level))
 portES = mean(portSimRetSort[portSimRetSort .< portVaR])
 hPosition = zeros(Float64, size(position[dataMergInd],1))

 for i=1:size(assSimRet,2)
  hPosition[:] = position[dataMergInd]
  hPosition[i] = 0.0
  hPosValue = closePrice[dataMergInd] .* hPosition
  CVaRj[i] = quantile(vec(assSimRet * hPosValue), (1-level))
 end
 CVaRj = -(CVaRj .- portVaR * portValue)

 return(portVaR, portES, CVaRj)
end

################################################################################
## Export Values singel and portfolio ES and VaR and conf Intervall
################################################################################
function  expVaRES(portVaR,
                   portES,
                   reSingleVaR,
                   reSingleES,
                   dataMergInd,
                   ticker,
#                   reSingleVaRConfInt,
#                  reSingleESConfInt,
                   resCVaR)

expCVaR = zeros(Float64,size(ticker,2))
inMerg = zeros(Float64,size(ticker,2))

inMerg[dataMergInd]=1
singleVaRConfInt = zeros(Float64, size(ticker,2),2)
singleESConfInt = zeros(Float64, size(ticker,2),2)
# singleVaRConfInt[dataMergInd,:] = reSingleVaRConfInt;
# singleESConfInt[dataMergInd,:] = reSingleESConfInt;
# singleVaRConfInt095 = singleVaRConfInt[:,1]
# singleVaRConfInt005 = singleVaRConfInt[:,2]
# singleESConfInt095 = singleESConfInt[:,1]
#singleESConfInt005 = singleESConfInt[:,2]

# @rput singleVaRConfInt095
# @rput singleVaRConfInt005
# @rput singleESConfInt095
#@rput singleESConfInt005
@rput reSingleVaR
@rput portVaR
@rput portES
@rput singleVaRConfInt
@rput singleESConfInt
@rput resCVaR
@rput inMerg
R"
 expFile <- 'Port.xlsx'
 wb <- loadWorkbook(file = expFile)
 writeData(wb, 'Risk', portVaR, startCol = 4, startRow = 2, rowNames = FALSE)
 writeData(wb, 'Risk', portES, startCol = 5, startRow = 2, rowNames = FALSE)

 writeData(wb, 'Risk', reSingleVaR, startCol = 8, startRow = 4, rowNames = FALSE)
 writeData(wb, 'Risk', resCVaR, startCol = 10, startRow = 4, rowNames = FALSE)
 writeData(wb, 'Risk', inMerg, startCol = 18, startRow = 4, rowNames = FALSE)
 saveWorkbook(wb, expFile, overwrite = TRUE)
"
#  writeData(wb, 'CurrentPositions', expSingleVar, startCol = 21, startRow = 3, rowNames = FALSE)
# writeData(wb, 'CurrentPositions', expSingleES, startCol = 26, startRow = 3, rowNames = FALSE)
# writeData(wb, 'CurrentPositions', singleVaRConfInt005, startCol = 23, startRow = 3, rowNames = FALSE)
# writeData(wb, 'CurrentPositions', singleVaRConfInt095, startCol = 22, startRow = 3, rowNames = FALSE)
# writeData(wb, 'CurrentPositions', singleESConfInt095, startCol = 27, startRow = 3, rowNames = FALSE)
# writeData(wb, 'CurrentPositions', singleESConfInt005, startCol = 28, startRow = 3, rowNames = FALSE)

end
