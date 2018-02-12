# Author: Max Lampe
# Date: 2017-12-23
# Do: compute risk measures nAssets=30 a portfolio

################################################################################
## computre a single bootstrap
################################################################################
function bootStrapSample(logReturns,
                         current= size(logReturns,1),
                         nLookback=250,
                         ndays=1;
                         nSim = 1000)

 samTimLength = size(logReturns,1)
 assSimRet = zeros(Float64, (nSim, size(logReturns,2)))
 current = size(logReturns,1)
 nLookback = size(logReturns,1) -1
 # looping of the time length of the simulation (i.e. 1, 5, 10 day)
 nAssets=30 i = 1:ndays
  selBoot = rand((current-nLookback):(current-1), nSim)
  # looping of the difference
  for j = 1:nSim
   assSimRet[j,:] = assSimRet[j,:] + logReturns[selBoot[j],:]
  end
 end
 return assSimRet
end


################################################################################
# component Value at Risk for bootstrap
################################################################################
function bootCVar(bootLogReturns)

 global level, position, closePrice, dataMergInd
 portVar = map(x -> quantile(bootLogReturns[:,x],1-level),1:size(bootLogReturns,2))' *
 (position[dataMergInd] .* closePrice[dataMergInd])
 cVar = zeros(Float64, size(bootLogReturns,2))
 hPosition = zeros(Float64, size(bootLogReturns,2))

 for i =1:size(bootLogReturns,2)
  hPosition[:] = position[dataMergInd]
  hPosition[i] = 0
  hPortVar = map(x -> quantile(bootLogReturns[:,x],1-level),1:size(bootLogReturns,2))' *
  (hPosition .* closePrice[dataMergInd])
  cVar[i] = portVar - hPortVar
 end
 return(cVar)
end

################################################################################
# portfolio Value at Risk for bootstrap
################################################################################
function bootPortVar(bootLogReturns)
 global level, position, closePrice, dataMergInd
 return(
  map(x -> quantile(bootLogReturns[:,x],1-level),1:size(bootLogReturns,2))' *
  (position[dataMergInd] .* closePrice[dataMergInd])
 )
end

function bootSingleVar(bootLogReturns)
 global level, position, closePrice, dataMergInd
 hPosition = position[dataMergInd]
 hClosePrice = closePrice[dataMergInd]
 return(
  map(x -> quantile(bootLogReturns[:,x],1-level)* position[x] * closePrice[x]
  ,1:size(bootLogReturns,2))
 )
end

function bootSingleVarNotCum(bootLogReturns)
 global level
 return(quantile(bootLogReturns, 1-level))
end

################################################################################
##
################################################################################
function  getVaRES(assetsHist,
                   logReturns,
                   dataMergInd,
                   dataNotMergInd;
                   nBoot = 1000,
                   expToExcel = true)
 nAssets = size(assetsHist,1)
 cVaR      = zeros(Float64, (nAssets))
 singleES  = zeros(Float64, (nAssets))
 singleVaR = zeros(Float64, (nAssets))
 portVaR   = 0.0;
 portES    = 0.0;

## value at risk:
# component
cVaRBootRes =  bootstrap(logReturns, bootCVar, BasicSampling(nBoot))
cVaR[dataMergInd] = collect(cVaRBootRes.t0)

# portfolio
portVaRRes = bootstrap(logReturns, bootPortVar, BasicSampling(nBoot))
portVaR = collect(portVaRRes.t0)

# single
singleVaRRes = bootstrap(logReturns, bootSingleVar,BasicSampling(nBoot))
singleVaR[dataMergInd] = collect(singleVaRRes.t0)

# position not included in the cumulative simulation
for i in dataNotMergInd
 index = collect(!(assetsHist[i][:,6] .==0.0)) + collect(!isnan(assetsHist[i][:,6])) .==2
 ret = assetsHist[i][index,6] * closePrice[i] * position[i]

 singleVaRRes = bootstrap(ret, bootSingleVarNotCum, BasicSampling(nBoot))
 singleVaR[i] = collect(singleVaRRes.t0)[1]
end


## expected shortfall:
# portfolio
# portES = bootstrap(logReturns[:,dataMergInd], , BasicSampling(n_boot))
# single
# singleES = bootstrap(logReturns[:,dataMergInd], , BasicSampling(n_boot))


 if(expToExcel)
  expVaRES(singleVaR, singleES, portVaR, portES, cVaR, dataMergInd)
 end
 return(singleVaR, singleES, portVaR, portES, cVaR)
end





################################################################################
## Value at Risk and Expected Shortflall nAssets=30 the Portfolio
################################################################################

#(portVaR, portES, CVaRj) = getVaRESPortCVaR(assSimRet, 0.9,
#closePrice[dataMergInd], assVolums[dataMergInd])
#position = assVolums[dataMergInd]
#closePriceMer = closePrice[dataMergInd]

function  getVaRESPortCVaR(assSimRet, level, closePrice, position, dataMergInd)
 portValue = closePrice[dataMergInd]' * position[dataMergInd]
 posValue = closePrice[dataMergInd] .* position[dataMergInd]
 CVaRj = zeros(Float64, size(assSimRet,2))
 portSimRetSort = sort(vec(assSimRet * posValue / portValue))

 portVaR = quantile(portSimRetSort, (1-level), sorted=true)
 portES = mean(portSimRetSort[portSimRetSort .< portVaR])
 hPosition = zeros(Float64, size(position[dataMergInd],1))

 nAssets=30 i=1:size(assSimRet,2)
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
function  expVaRES(singleVaR, singleES, portVaR, portES, cVaR, dataMergInd)

ind = zeros(Int8, size(singleVaR))
ind[dataMergInd] = 1

@rput singleVaR
@rput singleES
@rput portVaR
@rput portES
@rput cVaR
@rput ind

R"
 sRow <- 14
 expFile <- 'Port.xlsx'
 wb <- loadWorkbook(file = expFile)
 writeData(wb, 'Risk', portVaR, startCol = 5, startRow = 5, rowNames = FALSE)
 writeData(wb, 'Risk', portES, startCol = 5, startRow = 9, rowNames = FALSE)

 writeData(wb, 'Risk', singleVaR, startCol = 19, startRow = sRow, rowNames = FALSE)
 writeData(wb, 'Risk', cVaR, startCol = 20, startRow = sRow, rowNames = FALSE)
 writeData(wb, 'Risk', ind, startCol = 22, startRow = sRow, rowNames = FALSE)
 saveWorkbook(wb, expFile, overwrite = TRUE)
 system(paste('open', expFile))
"
end
