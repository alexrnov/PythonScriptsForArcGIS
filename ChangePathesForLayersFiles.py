import arcpy, os
def changeLayerFiles(fullPath):
  nameWithPath, extension = os.path.splitext(fullPath)
  if extension.lower() == ".lyr":
    layer = arcpy.mapping.Layer(fullPath)
    for lyr in arcpy.mapping.ListLayers(layer):
      lyr.findAndReplaceWorkspacePath("", sdeFile)
      #lyr.findAndReplaceWorkspacePaths(u"C:", u"D:")#do when need change name of disk
      outputFile = os.path.join(outputFolder, os.path.basename(nameWithPath) + ".lyr")
      lyr.saveACopy(outputFile, "10.3")
      arcpy.AddMessage("layer " + os.path.basename(nameWithPath) + ".lyr is change")
      del layer
                
if __name__ == '__main__':
  #input parameters:
	inputFolder = arcpy.GetParameterAsText(0) #folder with input files mxd 
  outputFolder = arcpy.GetParameterAsText(1) #folder with output files mxd
  #sdeFile = r"D:/gis2.sde"#do when need used locale sde-file
  sdeFile = r"\\10.64.68.36\j\ElectronicArchive\Connect\gis1.sde"
  for fileName in os.listdir(inputFolder):
	fullPath = os.path.join(inputFolder, fileName)
    if os.path.isfile(fullPath): changeLayerFiles(fullPath)
        
