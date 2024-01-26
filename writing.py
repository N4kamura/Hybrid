import xml.etree.ElementTree as ET
import concurrent.futures
import os
import time
from collections import Counter
import shutil
from openpyxl import load_workbook
import pandas as pd
import re
from io import StringIO
from tools import read_one_excel, peakhourfinder, obtener_numero_al_final

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
    
    directorio_flujogramas = os.path.join(directorio_proyecto,f"7. Informacion de Campo\\{nombre_subarea}\\Vehicular\\{tipicidad}")

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
    start_time = time.perf_counter()
    print("Calculando Hora Pico")

    # Finding Peak Hour

    resultados_peakhour = []

    with concurrent.futures.ProcessPoolExecutor(
        max_workers=2
    ) as executor:
        data_peakhour = executor.map(peakhourfinder,
                                        [
                                            (turno,excel_path) for excel_path in excel_paths
                                        ])
        resultados_peakhour.append(list(data_peakhour))

    end_time = time.perf_counter()
    print(f"Tiempo usado en cáculo de hora pico: {end_time-start_time:.2f} segundos")

    horas_puntas = []

    for elements in resultados_peakhour:
        counter = Counter(item[0] for item in elements)
        moda = counter.most_common()
        valor_acompanante_maximo = max(item[1] for item in moda)

        new_moda = [item[0] for item in moda if item[1] == valor_acompanante_maximo]

        lista_modas = []
        for individual_moda in new_moda:
            maximo = 0
            for element in elements:
                if element[0] == individual_moda:
                    if element[2][0] > maximo:
                        maximo = element[2][0]
            lista_modas.append((individual_moda,maximo))

        maximo = 0
        maximo_pair = []

        for lista in lista_modas:
            if lista[1] > maximo:
                maximo = lista[1]
                maximo_pair = lista

        horas_puntas.append((maximo_pair[0],maximo_pair[0]+4))

    hours = int(horas_puntas[0][1]//4)
    minutes = int(((horas_puntas[0][1]/4)%1)*100*0.15/0.25)
    print(f'PEAK HOUR DEFINITIVE: {hours-1:02d}:{minutes:02d} - {hours:02d}:{minutes:02d}')

    #Reading Excels:

    start_time = time.perf_counter()
    print("Leyendo datos de volúmenes por cada excel")
    with concurrent.futures.ProcessPoolExecutor(
        max_workers=2 #aquí va num_process, yo lo he definido como valor 2
    ) as executor:
        data_excel = executor.map(
            read_one_excel,
            [
                excel_path
                for excel_path in excel_paths
            ],
            [11]*len(excel_paths), #Aquí yo he colocado 10 de manera general TODO: <----------------------
            [slice(horas_puntas[0][0],horas_puntas[0][1])]*len(excel_paths),
        )
    end_time = time.perf_counter()
    print(f'Tiempo usado en la lectura de excels: {end_time-start_time:.2f} segundos')

    data_excel = list(data_excel)

    #Obtaining vehicle types:
    print("Reading Vehicle Types")
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

    numero_total = len(data_excel)*len(data_excel[0])*len(data_excel[0][0][2])*len(data_excel[0][0][1])

    #----------------------------------------------------------------------------#
    # Writing field data in GEH #
    #----------------------------------------------------------------------------#

    wb = load_workbook(modelo,read_only=False,data_only=False,keep_vba=True)
    ws = wb['GEH']
    all_values = []
    insercion_data_inicio = time.perf_counter()

    for i in range(len(types)):
        ws[vehicles_names].value = types[i]
        vehicles_names = vehicles_names[0] + str(9+i)

    for l in range(len(data_excel[0][0][2])): #Aquí estan 10 tipos de vehiculos
        for i in range(len(data_excel)): #Son 3: Número de excels
            for j in range(len(data_excel[i])): #Son 4: N, S, E, O
                for k in range(len(data_excel[i][j][1])): #Aquí es por número de giros (en este caso 3)
                    all_values.append([data_excel[i][j][0],data_excel[i][j][1][k],types[l],data_excel[i][j][2][l][k]])

    print(len(all_values))
    print(numero_total)

    for index in range(numero_total):
        ws[intersection].value = all_values[index][0][7:]
        ws[od].value = all_values[index][1]
        ws[vehicle_type].value = all_values[index][2]
        ws[volumes].value = all_values[index][3]

        intersection    = intersection[0]   + str(9+index)
        od              = od[0]             + str(9+index)
        vehicle_type    = vehicle_type[0]   + str(9+index)
        volumes         = volumes[0]        + str(9+index)
    
    wb.save(modelo)
    wb.close()

    insercion_data_fin = time.perf_counter()

    print(f"Tiempo de escritura en excel = {insercion_data_fin-insercion_data_inicio:.2f} segundos.")

def writing_model(vissim_path) -> None:
    #---------------------------------------------#
    # Obtaining simulated data #
    #---------------------------------------------#

    inicio_reading = time.perf_counter()
    directorio = os.path.dirname(vissim_path)
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

            if contador == 2:
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

    final_reading = time.perf_counter()
    print(f"Tiempo de lectura de resultados = {final_reading-inicio_reading:.2f} segundos.")

    #---------------------------------------------#
    # Writing results #
    #---------------------------------------------#

    inicio_writing = time.perf_counter()
    directorio, _ = os.path.split(vissim_path)
    modelo = os.path.join(directorio,'Reporte_GEH-R2.xlsm')

    wb = load_workbook(modelo,read_only=False,data_only=False,keep_vba=True)
    ws = wb['GEH']
    
    intersection    = 'K8'
    origin          = 'L8'
    destiny         = 'M8'
    vehicle_type    = ['BC8','BD8','BE8','BF8','BG8','BH8','BI8','BJ8','BK8','BL8','BM8','BN8','BO8','BP8','BQ8','BR8','BS8','BT8','BU8','BV8']

    for index, row in df.iterrows():
        #Writing:
        ws[intersection].value  = row.iloc[2]
        ws[origin].value        = row.iloc[3]
        ws[destiny].value       = row.iloc[4]

        intersection    = intersection[0]   + str(9+index)
        origin          = origin[0]         + str(9+index)
        destiny         = destiny[0]        + str(9+index) 

        for i in range(len(row[5:])):
            ws[vehicle_type[i]].value = row.iloc[i+5]

        for j in range(len(vehicle_type)):
            vehicle_type[j] = vehicle_type[j][:2] + str(int(vehicle_type[j][2:])+1)

    wb.save(modelo)
    wb.close()

    os.startfile(modelo)

    final_writing = time.perf_counter()
    print(f"Tiempo de escritura en excel = {final_writing-inicio_writing:.2f} segundos.")