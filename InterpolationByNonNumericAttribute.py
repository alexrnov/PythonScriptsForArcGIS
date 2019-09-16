import os
import io
import subprocess

import arcpy as arc
import arcgisscripting

def read_map_layer(layer, attribute):
    rows = gp.SearchCursor(layer)
    row = rows.Next()
    layer = gp.describe(layer)
    points = []
    while row:
        feat = row.GetValue(layer.ShapeFieldName)
        point = feat.getpart(0)
        x = point.x
        y = point.y
        z = row.getValue(attribute)
        if (type(z) == unicode):
            z = z.encode('utf-8')
        if z != None and z != '' and z != ' ':
            points.append([x, y, z])
        row = rows.next()
    return points

def write_to_map_layer(points):
    arc.CreateFeatureclass_management(layer_path, layer_name, "POINT", "", "", "ENABLED", coordinat_system)
    gp.AddField(output_map_layer, "x", "FLOAT", 15, 15)
    gp.AddField(output_map_layer, "y", "FLOAT", 15, 15)
    gp.AddField(output_map_layer, "z", "STRING", "", "", 1000)
    rows = gp.InsertCursor(output_map_layer)
    pnt = gp.CreateObject("Point")
    for point in points:
        attributes = point.split(SEPARATOR)
        feat = rows.NewRow()
        x = attributes[0]
        y = attributes[1]
        x = float(x)
        y = float(y)
        feat.SetValue("x", x)
        feat.SetValue("y", y)
        feat.SetValue("z", attributes[2])
        pnt.x = x
        pnt.y = y
        feat.shape = pnt
        rows.InsertRow(feat)
          
if __name__ == '__main__':
    global gp
    gp = arcgisscripting.create(9.3)

    TYPE_OF_TASK = u"Интерполяция по нечисловому атрибуту".encode('cp1251')
    OUTPUT_FILE = "addToolsInterpolation.txt"
    JAR_FILE = "AdditionalTools.jar"
    FILE_WITH_MATRIX = "OutputFileForXYZLayer.txt"
    MAX_AMOUNT_POINTS = 1353627
    global SEPARATOR
    SEPARATOR = ";;";
    
    input_layer = arc.GetParameterAsText(0).encode('cp1251')
    attribute = arc.GetParameterAsText(1).encode('cp1251')
    try:
        cell_size = float(arc.GetParameterAsText(2))
    except ValueError:
        arc.AddMessage(u"Некорректное значение размера растровой ячейки")
        raise SystemExit
    
    global coordinat_system
    coordinat_system = arc.GetParameterAsText(3).encode('cp1251')

    output_raster = arc.GetParameterAsText(4).encode('cp1251')

    global output_map_layer
    path_base, name_raster_file = os.path.split(output_raster)
    output_map_layer = os.path.join(path_base, "FileForConvertToRastr")
    
    global layer_path
    global layer_name
    layer_path, layer_name = os.path.split(output_map_layer)
    
    points = read_map_layer(input_layer, attribute)

    layer_path, layer_name = os.path.split(output_map_layer)
    base_directory, base = os.path.split(layer_path)
    output_file_dir = os.path.join(base_directory, OUTPUT_FILE)
    
    with open(output_file_dir, "w") as text_file:
        for p in points:
            text_file.write(str(p[0]) + SEPARATOR + str(p[1]) + SEPARATOR + str(p[2]) + "\n")

    script_dir = os.path.dirname(os.path.abspath(__file__))
    jar_dir = os.path.join(script_dir, JAR_FILE)
    jar_dir = jar_dir.encode('cp1251')
    code_result = subprocess.call(['java', '-jar', jar_dir,
                                   TYPE_OF_TASK,
                                   output_file_dir,
                                   str(cell_size)])
    if code_result != 0:
        arc.AddMessage(u"Вычисления не выполнены")
        os.remove(output_file_dir)
        raise SystemExit

    os.remove(output_file_dir)
    input_file_dir = os.path.join(base_directory, FILE_WITH_MATRIX)
    with io.open(input_file_dir, "r", encoding='utf-8') as f:
        lines = f.read()
    os.remove(input_file_dir)
    points = lines.split("\n")
    del(points[len(points)-1])

    if (len(points) > MAX_AMOUNT_POINTS):
        arc.AddMessage(u"Вычисления не выполнены - необходимо увеличить размер растровой ячейки")
        raise SystemExit
    
    write_to_map_layer(points)
        
    try:
        arc.PointToRaster_conversion(output_map_layer, "z", output_raster, "MAXIMUM", "", cell_size)
    except RuntimeError:
        arc.Delete_management(output_map_layer)
        arc.AddMessage(u"Не удалось создать растровой слой")
        raise SystemExit
    
    arc.Delete_management(output_map_layer)
