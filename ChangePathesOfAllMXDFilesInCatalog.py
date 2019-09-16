import arcpy, os

def changeMXDFile(fullPath):
  nameWithPath, extension = os.path.splitext(fullPath)
  if extension.lower() == ".mxd":
    mxd = arcpy.mapping.MapDocument(fullPath)
    #mxd.findAndReplaceWorkspacePaths(u"\\10.64.68.36\j\ElectronicArchive\FondOldCGM\", u"D:")#do when need change name of disk
    mxd.findAndReplaceWorkspacePaths(u"\\10.64.68.36\j\FondOldCGM\", u"\\10.64.68.36\j\ElectronicArchive\FondOldCGM\")
    outputFile = os.path.join(outputFolder, os.path.basename(nameWithPath) + ".mxd")
    mxd.saveACopy(outputFile, "10.3")
    arcpy.AddMessage("file " + os.path.basename(nameWithPath) + ".mxd is change")
    del mxd
                
if __name__ == '__main__':
  #input parameters:
  inputFolder = arcpy.GetParameterAsText(0) #folder with input files mxd 
  outputFolder = arcpy.GetParameterAsText(1) #folder with output files mxd
  #sdeFile = r"D:/gis2.sde"#do when need used locale sde-file
  sdeFile = r"\\10.64.68.36\j\ElectronicArchive\FondOldCGM\"
	
  for fileName in os.listdir(inputFolder):
    fullPath = os.path.join(inputFolder, fileName)
    if os.path.isfile(fullPath): changeMXDFile(fullPath)
