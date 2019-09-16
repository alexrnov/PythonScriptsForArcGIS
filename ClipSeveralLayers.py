import os
import arcpy as arc
if __name__ == '__main__':
    input_layers = arc.GetParameterAsText(0).encode('cp1251')
    clip_layer = arc.GetParameterAsText(1).encode('cp1251')
    xy_tolerance = arc.GetParameterAsText(2)
    output_db = arc.GetParameterAsText(3).encode('cp1251')
    for lay_for_clipping in input_layers.split(";"):
        arc.AddMessage("Clip for " + lay_for_clipping)
        lay_for_clipping = lay_for_clipping.replace("'",'')
        layer_path, layer_name = os.path.split(lay_for_clipping)
        output_layer = os.path.join(output_db, layer_name)
        arc.Clip_analysis(lay_for_clipping, clip_layer, output_layer, xy_tolerance)
