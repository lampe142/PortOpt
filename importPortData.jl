######################################################
 # get current portfio positions
######################################################

function portData(nAssets)
 dataFileName = "Port.xlsx"
 f = openxl(dataFileName)
 sheetsInfosCNames = readxl(f, "Risk!B3:B" * string(nAssets+3))
 sheetsInfos = readxl(f, "Risk!B3:D" * string(nAssets+3));
 posNames = sheetsInfos[:,1];
 assVolums = convert(Array{Float64,1}, sheetsInfos[:,3])
 return(assVolums, posNames)
end;
