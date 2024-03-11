import xml.etree.ElementTree as ET
import xlwings as xw
import os
import time
from collections import Counter
import shutil
from openpyxl import load_workbook
import pandas as pd
import re
from io import StringIO
from tools import read_one_excel, peakhourfinder, obtener_numero_al_final
from unidecode import unidecode

def writing_campo(vissim_path, turno) -> None:
    tree = ET.parse(vissim_path)
    network_tag = tree.getroot()

    uda_key = network_tag.find(".//userDefinedAttribute[@nameLong='Código']")
    key = uda_key.get('no')
    nodes_info = []

    for node_tag in network_tag.findall("./nodes/node"):
        number_node = node_tag.get("no")
        uda_element = node_tag.find(f"./uda[@key='{key}']")
        codigo_intersection = uda_element.get("value")
        nodes_info.append((number_node,codigo_intersection))

    nombre_subarea = vissim_path
    for _ in range(4):
        nombre_subarea = os.path.dirname(nombre_subarea)
    _,nombre_subarea = os.path.split(nombre_subarea)

    tipicidad = vissim_path
    for _ in range(2):
        tipicidad = os.path.dirname(tipicidad)
    _,tipicidad = os.path.split(tipicidad)

    directorio_proyecto = vissim_path
    for _ in range(6):
        directorio_proyecto = os.path.dirname(directorio_proyecto)

    directorio_flujogramas = os.path.join(directorio_proyecto, "7. Informacion de Campo", nombre_subarea, "Vehicular", tipicidad)

    files = os.listdir(directorio_flujogramas)

    flujogramas_ordered = []

    for node_info in nodes_info:
        for file in files:
            if file[:5] == node_info[1]: #TODO: <------ Se puede mejorar solo captando patrones.
                flujogramas_ordered.append(file)
                break

    excel_paths = []
    for flujograma in flujogramas_ordered:
        flujograma = os.path.join(directorio_flujogramas,flujograma)
        excel_paths.append(flujograma)
    #start_time = time.perf_counter()

    # Finding Peak Hour
    resultados_peakhour = []

    for excel_path in excel_paths:
        resultados_peakhour.append(peakhourfinder(turno, excel_path))

    #end_time = time.perf_counter()

    horas_puntas = []

    horas_puntas = [hora[0] for hora in resultados_peakhour]
    counter_horas = Counter(horas_puntas)
    modas_horas = [hora for hora, frecuencia in counter_horas.items() if frecuencia == max(counter_horas.values())]

    mayor_volumen = 0

    for moda_hora in modas_horas:
        volumenes_moda = [hora[1] for hora in resultados_peakhour if hora[0] == moda_hora]
        max_volumen_moda = max(volumenes_moda)

        if max_volumen_moda > mayor_volumen:
            mayor_volumen = max_volumen_moda
            peakhour = moda_hora

    #hours = int(peakhour//4)
    #minutes = int(((peakhour/4)%1)*100*0.15/0.25)

    #Reading Excels:

    #start_time = time.perf_counter()

    num_veh_classes = 11
    interval = slice(peakhour, peakhour+4)
    data_excel = []
    for excel_path in excel_paths:
        data_excel.append(read_one_excel(excel_path, num_veh_classes, interval))

    #end_time = time.perf_counter()

    #Obtaining vehicle types:
    wb = load_workbook(excel_paths[0],read_only=True,data_only=True)
    ws_ma=wb['V_Ma']
    
    types = [row[0].value for s in [slice("O59","O78")] for row in ws_ma[s]]
    wb.close()

    #Copia del modelo.
    modelo_original = "./images/Modelo.xlsm"
    directorio,_ = os.path.split(vissim_path)
    modelo = os.path.join(directorio,'Reporte_GEH-R2.xlsm')
    shutil.copy2(modelo_original,modelo)

    vehicles_names = 'H8'
    intersection = 'B8'
    od = 'C8'
    vehicle_type = 'D8'
    volumes = 'E8'

    #numero_total = len(data_excel)*len(data_excel[0])*len(data_excel[0][0][2])*len(data_excel[0][0][1])

    #----------------------------------------------------------------------------#
    # Writing field data in GEH #
    #----------------------------------------------------------------------------#

    wb = xw.Book(modelo)
    ws = wb.sheets['GEH']
    all_values = []
    #insercion_data_inicio = time.perf_counter()

    for i in range(len(types)):
        ws.range(vehicles_names).value = types[i]
        vehicles_names = vehicles_names[0] + str(9+i)

    for excel in data_excel:
        for access in excel:
            for i, turns in enumerate(access[1]):
                for j, volume in enumerate(access[2]):
                    all_values.append([access[0], turns, types[j], volume[i]])

    for index, values in enumerate(all_values):
        ws.range(intersection).value = values[0]
        ws.range(od).value = values[1]
        ws.range(vehicle_type).value = values[2]
        ws.range(volumes).value = values[3]

        intersection = intersection[0] + str(9+index)
        od = od[0] + str(9+index)
        vehicle_type = vehicle_type[0] + str(9+index)
        volumes = volumes[0] + str(9+index) 

    wb.save(modelo)
    wb.close()
    xw.App().quit()

    #insercion_data_fin = time.perf_counter()

def writing_model(vissim_path) -> None:
    #---------------------------------------------#
    # Obtaining simulated data #
    #---------------------------------------------#

    #inicio_reading = time.perf_counter()
    directorio = os.path.dirname(vissim_path)

    #Extrayendo nombres de los tipos vehiculares por clases:
    number_by_name = {}
    tree = ET.parse(vissim_path)
    network_tag = tree.getroot()

    number_by_name = {
        c.attrib["no"]: unidecode(c.attrib["name"]).upper()
        for c in network_tag.findall("./vehicleClasses/vehicleClass")
    }

    patron = re.compile(r'_Node Results_\d{3}.att') #Busca todos los archivos que acaban en 3 dígitos y se filtra el mayor. Aggregated data need to be activated.

    archivos_att = [f for f in os.listdir(directorio) if patron.search(f)]

    archivo_mas_alto = max(archivos_att,key = obtener_numero_al_final)

    path = os.path.join(directorio,archivo_mas_alto)

    #---------------------------------------------#
    # Reading result attribute data #
    #---------------------------------------------#

    with open(path,'r') as att_file:
        contador = 0
        data = []
        for line in att_file:
            line = line.strip()
            if line.startswith('$') and contador<3:
                contador += 1

            if contador == 1: #<------- PARECE QUE EN EL 24 SE GENERA UN POCO MAL Y SOLO SE GENERA UN $ A PESAR QUE SALEN DOS.
                data.append(line)

    #---------------------------------------------#
    # Obtaining dataframe from the data #
    #---------------------------------------------#
                
    data = data[:-1]
    data_str = '\n'.join(data)
    data_io = StringIO(data_str)
    df = pd.read_csv(data_io,delimiter=';')

    #---------------------------------------------#
    # Cleaning and filtering data #
    #---------------------------------------------#
    df = df.dropna(subset=['MOVEMENT\FROMLINK\ORIGEN','MOVEMENT\TOLINK\DESTINO'])
    df = df.reset_index(drop=True)
    df = df.loc[df.iloc[:,0] == 'AVG']
    df = df.reset_index(drop=True)

    #final_reading = time.perf_counter()

    #---------------------------------------------#
    # Writing results #
    #---------------------------------------------#

    #inicio_writing = time.perf_counter()
    directorio, _ = os.path.split(vissim_path)
    modelo = os.path.join(directorio,'Reporte_GEH-R2.xlsm')

    wb = xw.Book(modelo)
    ws = wb.sheets['GEH']

    num_columns = df.shape[1]

    for index, row in df.iterrows():
        ws.range(8+index, 11).value = row.iloc[2]
        ws.range(8+index, 12).value = row.iloc[3]
        ws.range(8+index, 13).value = row.iloc[4]

    for i in range(20):
        name_vehicle = ws.range(7, 55+i).value
        name_vehicle = unidecode(name_vehicle).upper()

        if name_vehicle == 'N' or len(name_vehicle) == 1:
            break

        for key, value in number_by_name.items():
            if value == name_vehicle:
                for j in range(num_columns-5):
                    name_column = df.columns[j+5]
                    pattern = r"VEHS\((\d+)\)"
                    number_class = re.search(pattern, name_column).group(1)
                    if str(number_class) == str(key):
                        for fila, elem in enumerate(df.iloc[:, j+5]):
                            ws.range(8+fila, 55+i).value = elem

    wb.save(modelo)
    #final_writing = time.perf_counter()