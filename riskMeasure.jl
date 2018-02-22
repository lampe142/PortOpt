# Author: Max Lampe
# Date: 2017-12-23
# Do: compute risk measures nAssets=30 a portfolio


################################################################################
# component Value at Risk for bootstrap
################################################################################
function bootCVar(bootLogReturns)

 global level, position, closePrice, dataMergInd
 portVar = quantile(bootLogReturns * (position[dataMergInd] .* closePrice[dataMergInd]),1-level)
 cVar = zeros(Float64, size(bootLogReturns,2))
 hPosition = zeros(Float64, size(bootLogReturns,2))

 for i =1:size(bootLogReturns,2)
  hPosition[:] = position[dataMergInd]
  hPosition[i] = 0
  hPortVar = quantile(bootLogReturns * (hPosition .* closePrice[dataMergInd]),1-level)
  cVar[i] = portVar - hPortVar
 end
 return(cVar)
end

################################################################################
# marginal Value at Risk for bootstrap
################################################################################
function bootMVar(bootLogReturns)

 global level, position, closePrice, dataMergInd
 portVar = quantile(bootLogReturns * (position[dataMergInd] .* closePrice[dataMergInd]),1-level)
 mVar = zeros(Float64, size(bootLogReturns,2))
 hPosition = zeros(Float64, size(bootLogReturns,2))

 for i =1:size(bootLogReturns,2)
  hPosition[:] = position[dataMergInd]
  hPosition[i] = hPosition[i] + 1/closePrice[i]
  hPortVar = quantile(bootLogReturns * (hPosition .* closePrice[dataMergInd]),1-level)
  mVar[i] = hPortVar - portVar
 end
 return(mVar)
end


################################################################################
# portfolio Value at Risk for bootstrap
################################################################################
function bootPortVar(bootLogReturns)
 global level, position, closePrice, dataMergInd
 return quantile(bootLogReturns * (position[dataMergInd] .* closePrice[dataMergInd]),1-level)

end

function bootSingleVar(bootLogReturns)
 global level, position, closePrice, dataMergInd
 hPosition = position[dataMergInd]
 hClosePrice = closePrice[dataMergInd]
 return(
  map(x -> quantile(bootLogReturns[:,x],1-level)* hPosition[x] * hClosePrice[x]
  ,1:size(bootLogReturns,2))
 )
end

function bootSingleVar2(bootLogReturns)
 return(quantile(bootLogReturns,1-level))
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
 mVaR      = zeros(Float64, (nAssets))
 singleES  = zeros(Float64, (nAssets))
 singleVaR = zeros(Float64, (nAssets))
 portVaR   = 0.0;
 portES    = 0.0;

## value at risk:
# component
cVaRBootRes =  bootstrap(logReturns, bootCVar, BasicSampling(nBoot))
cVaR[dataMergInd] = collect(cVaRBootRes.t0)

# marginal
mVaRBootRes =  bootstrap(logReturns, bootMVar, BasicSampling(nBoot))
mVaR[dataMergInd] = collect(mVaRBootRes.t0)

# portfolio
portVaRRes = bootstrap(logReturns, bootPortVar, BasicSampling(nBoot))
portVaR = collect(portVaRRes.t0)

# single
#singleVaRRes = bootstrap(logReturns, bootSingleVar,BasicSampling(nBoot))
#singleVaR[dataMergInd] = collect(singleVaRRes.t0)

lRMerg = map( x -> logReturns[:,x] * (closePrice[dataMergInd] .* position[dataMergInd])[x]
, 1:size(logReturns,2))
for i in 1:size(dataMergInd,1)
 singleVaRRes = bootstrap(lRMerg[i], bootSingleVar2, BasicSampling(nBoot))
 singleVaR[dataMergInd[i]] = collect(singleVaRRes.t0)[1]
end

# position not included in the cumulative simulation
if(dataNotMergInd != [0])
 for i in dataNotMergInd
  index = collect(!(assetsHist[i][:,6] .==0.0)) + collect(!isnan(assetsHist[i][:,6])) .==2
  ret = assetsHist[i][index,6] * closePrice[i] * position[i]

  singleVaRRes = bootstrap(ret, bootSingleVarNotCum, BasicSampling(nBoot))
  singleVaR[i] = collect(singleVaRRes.t0)[1]
 end
end
## expected shortfall:
# portfolio
# portES = bootstrap(logReturns[:,dataMergInd], , BasicSampling(n_boot))
# single
# singleES = bootstrap(logReturns[:,dataMergInd], , BasicSampling(n_boot))

# Export Data into the EXcel File 
 if(expToExcel)
  expVaRES(singleVaR, singleES, portVaR, portES, cVaR, mVaR, dataMergInd)
 end

 return(singleVaR, singleES, portVaR, portES, cVaR, mVaR)
end

################################################################################
## Export Values singel and portfolio ES and VaR and conf Intervall
################################################################################
function  expVaRES(singleVaR, singleES, portVaR, portES, cVaR, mVaR, dataMergInd)

ind = zeros(Int8, size(singleVaR))
ind[dataMergInd] = 1

@rput singleVaR
@rput singleES
@rput portVaR
@rput portES
@rput cVaR
@rput mVaR
@rput ind

R"
 sRow <- 3
 expFile <- 'Port.xlsx'
 wb <- loadWorkbook(file = expFile)

 writeData(wb, 'Risk', singleVaR, startCol = 19, startRow = sRow, rowNames = FALSE)
 writeData(wb, 'Risk', cVaR, startCol = 20, startRow = sRow, rowNames = FALSE)
 writeData(wb, 'Risk', ind, startCol = 22, startRow = sRow, rowNames = FALSE)
 writeData(wb, 'Risk', mVaR, startCol = 23, startRow = sRow, rowNames = FALSE)
 saveWorkbook(wb, expFile, overwrite = TRUE)
 system(paste('open', expFile))
"
# writeData(wb, 'Risk', portVaR, startCol = 5, startRow = 5, rowNames = FALSE)
# writeData(wb, 'Risk', portES, startCol = 5, startRow = 9, rowNames = FALSE)

end
