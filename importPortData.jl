######################################################
 # get current portfio positions
######################################################

function portData(nAssets)
 dataFileName = "Port.xlsx"
 f = openxl(dataFileName)
 sheetsInfosCNames = readxl(f, "Risk!B14:B" * string(nAssets+13))
 sheetsInfos = readxl(f, "Risk!B14:D" * string(nAssets+13));
 posNames = sheetsInfos[:,1];
 assVolums = convert(Array{Float64,1}, sheetsInfos[:,3])
 return(assVolums, nAssets, posNames)
end;
