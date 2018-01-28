######################################################
 # get current portfio positions
######################################################

function portData()
 dataFileName = "Port.xlsx"
 f = openxl(dataFileName)
 sheetsInfosCNames = readxl(f, "Risk!B2:B4")
 sheetsInfos = readxl(f, "Risk!B4:C30");
 posNames = sheetsInfos[:,1];
 nAssets = size(sheetsInfos,1);
 assVolums = convert(Array{Float64,1}, sheetsInfos[:,2])
 return(assVolums, nAssets, posNames)
end;
