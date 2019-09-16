import arcpy, os
def getListLayerPathesOfMXD(inputMXDFile):
  mxd = arcpy.mapping.MapDocument(inputMXDFile)
  for layer in arcpy.mapping.ListLayers(mxd):
    if layer.supports("DATASOURCE"): arcpy.AddMessage(layer.dataSource)
        
if __name__ == '__main__':
  inputMXDFile = arcpy.GetParameterAsText(0) 
  if os.path.splitext(inputMXDFile)[1] == ".mxd": getListLayerPathesOfMXD(inputMXDFile)
    else: arcpy.AddError("input file is not map document")
