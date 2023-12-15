from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow,QLabel,QFileDialog,QButtonGroup
from PyQt5.QtGui import QPixmap
import win32com.client as com
#import pygetwindow as gw
import os
import pandas as pd
import xlwings as xw
import numpy as np
from openpyxl import load_workbook
import re
import time
import concurrent.futures
from collections import Counter
import shutil
from io import StringIO
import xml.etree.ElementTree as ET

nombres_vehiculos = []
enviado_contador = 1

def obtener_numero_al_final(archivo):
    match = re.search(r'(\d+)\.att$', archivo)
    return int(match.group(1)) if match else 0

def peakhourfinder(args): 
    turno,excel_path = args
    slices = [
            ["HR16:HR39",0+1,"MADRUGADA"], #El +1 es para que obtenga el límite superior, p.e.: [8:00 - 8:15] -> 8:15
            ["HR40:HR63",24+1,"MAÑANA"],
            ["HR64:HR83",48+1,"TARDE"],
            ["HR84:HR111",68+1,"NOCHE"],
            ]
    wb = load_workbook(excel_path, read_only=True, data_only=True)
    ws = wb["N"]

    list = [[cell.value for cell in row] for row in ws[slices[turno][0]]]
    maximum = list[0]
    maximum_cell = 0
    for i in range(1,len(list)):
        if list[i]>maximum:
            maximum         = list[i]
            maximum_cell    = i

    if maximum[0] !=0: #Acá obtengo el hora inicio. p.e.: 9*4 a 10*4 (9-10) o 9.25 a 10.25 (9:15 - 10:15)
        plus = slices[turno][1]
        peakhour = maximum_cell + plus
        hora_final, hora_inicio = peakhour,peakhour-4
        hours = int(peakhour//4)
        minutes = int(((peakhour/4)%1)*100*0.15/0.25)
        #print(f"Excel: {excel_path}")
        print(f"PEAK HOUR {slices[turno][2]}: {hours-1:02d}:{minutes:02d} - {hours:02d}:{minutes:02d} | Volumen = {maximum[0]}")
    else:
        print(f"PEAK HOUR {slices[turno][2]}: NO HAY DATOS O EL FLUJO ES NULO")
        pass
    
    return hora_inicio,hora_final, maximum

def read_one_excel(excel_path,num_veh_classes,interval):
    wb = load_workbook(excel_path,read_only=True,data_only=True)

    ws_inicio = wb['Inicio']

    num_giros = [
        [row[0].value for row in ws_inicio[s]].index(None)
        for s in [
            slice("G21", "G30"),
            slice("M21", "M30"),
            slice("G33", "G42"),
            slice("M33", "M42"),
        ]    
    ]

    name_giros = ["N", "S", "E", "O"]

    list_destination = []
    list_origin      = []

    list_slice_destination = (
        slice("F21", "F30"),
        slice("L21", "L30"),
        slice("F33", "F42"),
        slice("L33", "L42"),
    )

    list_slice_origin = (
        slice("E21", "E30"),
        slice("K21", "K30"),
        slice("E33", "E42"),
        slice("K33", "K42"),
    )
    final_result=[]
    for i_giro in range(4):
        list_destination = []
        list_origin      = []
        slice_origin = list_slice_origin[i_giro]
        slice_destination = list_slice_destination[i_giro]
        num_giro_i = num_giros[i_giro]

        list_destination.append([row[0].value for row in ws_inicio[slice_destination]][:num_giro_i])

        list_origin.append([row[0].value for row in ws_inicio[slice_origin]][:num_giro_i])

        list_od=[[str(a)+str(b) for a,b in zip(fila1,fila2)] for fila1, fila2 in zip(list_origin,list_destination)]
        list_od = list_od[0]

        ###### GET FLOWS ######
        ws = wb[name_giros[i_giro]]
        list_A = [[cell.value for cell in row] for row in ws["K16":"HB111"]]

        #Conversión a numpy.
        A = np.array(list_A, dtype="float")

        #Asignación de 0 para datos que no contienen data.
        A[np.isnan(A)] = 0

        list_sumas=[]
        matriz = np.array([A[interval, (10 * veh_type) : (10 * veh_type + num_giro_i)] for veh_type in range(num_veh_classes)])
        for mat in matriz:
            list_sumas.append([sum(mat[:,i]) for i in range(num_giro_i)]) #Saldría la suma de todos los giros en una lista y todo esto por cada tipo vehicular.
        
        _, name_excel = os.path.split(excel_path)
        final_result.append([name_excel[:-5],list_od,list_sumas])

    return final_result

param_groups = {}

class MiVentana(QMainWindow):
    def __init__(self): #Datos
        super().__init__()

        # Cargar la interfaz gráfica desde el archivo .ui
        uic.loadUi("./images/hybrid-gui.ui",self)

        #Definición de listas
        ####SEGUIMIENTO VEHICULAR####
        self.Ivehicle_type=["#Driving-Behavior"]
        self.ILookAheadDistMin=["LookAheadDistMin"]
        self.ILookAheadDistMax=["LookAheadDistMax"]
        self.ILookBackDistMin=["LookBackDistMin"]
        self.ILookBackDistMax=["LookBackDistMax"]

        ####WIEDEMANN 74####
        self.IW74ax=["W74ax"]
        self.IW74bxAdd=["W74bxAdd"]
        self.IW74bxMult=["W74bxMult"]

        ####CAMBIO DE CARRIL####
        self.IMaxDecelOwn=["MaxDecelOwn"]
        self.IMaxDecelTrail=["MaxDecelTrail"]
        self.IDecelRedDistOwn=["DecelRedDistOwn"]
        self.IDecelRedDistTrail=["DecelRedDistTrail"]
        self.IAccDecelOwn=["AccDecelOwn"]
        self.IAccDecelTrail=["AccDecelTrail"]
        self.IDiffusTm=["DiffusTm"]
        self.ISafDistFactLnChg=["SafDistFactLnChg"]
        self.ICoopDecel=["CoopDecel"]

        ####LATERAL####
        self.IDesLatPos=["DesLatPos"] #NEW
        self.IObsrvAdjLn=["ObsrvAdjLn"] #NEW
        self.IDiamQueu=["DiamQueu"] #NEW
        self.IConsNextTurn=["ConsNextTurn"] #NEW
        self.IOvtLDef=["OvtLDef"] #NEW
        self.IOvtRDef=["OvtRDef"] #NEW
        self.IMinCollTmGain=["MinCollTmGain"]
        self.IMinSpeedForLat=["MinSpeedForLat"]
        self.ILatDistStandDef=["LatDistStandDef"]
        self.ILatDistDrivDef=["LatDistDrivDef"]

        ####LABELS####
        self.label = self.findChild(QLabel,"label")
        self.label_2 = self.findChild(QLabel,"label_2")
        self.label_3 = self.findChild(QLabel,"label_3")
        self.label_5 = self.findChild(QLabel,"label_5")

        ####IMAGES####
        imagen = QPixmap('./images/car_follow.png')
        self.label.setPixmap(imagen)
        imagen_2 = QPixmap('./images/lane_change.png')
        self.label_2.setPixmap(imagen_2)
        image_3 = QPixmap('./images/lateral_behave.png')
        self.label_3.setPixmap(image_3)
        image_4 = QPixmap('./images/logo.png')
        self.label_5.setPixmap(image_4)

        #BOTONES
        self.save.clicked.connect(self.data_input)
        self.start.clicked.connect(self.ejecutar_programa)
        self.carpet.clicked.connect(self.carpeta)
        self.report.clicked.connect(self.reporte)
        self.activar.clicked.connect(self.data_campo)
        self.liviano.clicked.connect(self.livianos)
        self.menor.clicked.connect(self.menores)
        self.publico.clicked.connect(self.publicos)
        self.carga.clicked.connect(self.cargas)
        self.fijar.clicked.connect(self.fijars)
        self.exportar.clicked.connect(self.export_params_2_excel)
        
        #BOTONES PARA LOS TURNOS
        button_group = QButtonGroup(self)
        button_group.addButton(self.early)
        button_group.addButton(self.morning)
        button_group.addButton(self.evening)
        button_group.addButton(self.night)
        button_group.setExclusive(True)

        #BOTONES PARA LAS VERSIONES
        button_group_2 = QButtonGroup(self)
        button_group_2.addButton(self.checkBox)
        button_group_2.addButton(self.checkBox_2)

    def carpeta(self): #Carpeta que contiene el archivo .inpx
        self.path_file,self.inpx_name = QFileDialog.getOpenFileName(self,"Seleccionar Archivo .inpx","c:\\","Archivos .inpx (*.inpx)")

    def data_campo(self):
        global nombres_vehiculos

        while not (
            self.early.isChecked() or
            self.morning.isChecked() or
            self.evening.isChecked() or
            self.night.isChecked()
            ):
            return print("Proceso detenido: debes escoger un turno antes de utilizar este botón.")
        
        if self.early.isChecked():
            turno = 0
        elif self.morning.isChecked():
            turno = 1
        elif self.evening.isChecked():
            turno = 2
        elif self.night.isChecked():
            turno = 3
        
        ###############
        # Excel Paths #
        ###############
        #path_vissim = os.path.join(self.path_file,self.inpx_name)
        #tree = ET.parse(path_vissim)
        tree = ET.parse(self.path_file)
        network_tag = tree.getroot()

        uda_key = network_tag.find(".//userDefinedAttribute[@nameLong='Código']")
        key = uda_key.get('no')
        nodes_info = []

        for node_tag in network_tag.findall("./nodes/node"):
            number_node = node_tag.get("no")
            uda_element = node_tag.find(f"./uda[@key='{key}']")
            codigo_intersection = uda_element.get("value")
            nodes_info.append((number_node,codigo_intersection))

        #nodes_info = [[('1', 'SS-88'), ('2', 'SS-25')]]

        nombre_subarea = self.path_file
        for _ in range(4):
            nombre_subarea = os.path.dirname(nombre_subarea)
        _,nombre_subarea = os.path.split(nombre_subarea)

        tipicidad = self.path_file
        for _ in range(3):
            tipicidad = os.path.dirname(tipicidad)
        _,tipicidad = os.path.split(tipicidad)

        directorio_proyecto = self.path_file
        for _ in range(6):
            directorio_proyecto = os.path.dirname(directorio_proyecto)
        
        directorio_flujogramas = os.path.join(directorio_proyecto,f"7. Informacion de Campo\\{nombre_subarea}\\Vehicular\\{tipicidad}")

        files = os.listdir(directorio_flujogramas)

        flujogramas_ordered = []

        for node_info in nodes_info:
            for file in files:
                if file[:5] == node_info[1]:
                    flujogramas_ordered.append(file)
                    break

        excel_paths = []
        for flujograma in flujogramas_ordered:
            flujograma = os.path.join(directorio_flujogramas,flujograma)
            excel_paths.append(flujograma)

        start_time = time.perf_counter()
        print("Calculando Hora Pico")

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
                [10]*len(excel_paths), #Aquí yo he colocado 10 de manera general
                [slice(horas_puntas[0][0],horas_puntas[0][1])]*len(excel_paths),
            )
        end_time = time.perf_counter()
        print(f'Tiempo usado en la lectura de excels: {end_time-start_time:.2f} segundos')

        data_excel = list(data_excel)

        #Obtención de Tipos vehiculares
        print("DATOS DE TIPOS VEHICULARES")
        wb = load_workbook(excel_paths[0],read_only=True,data_only=True)
        ws_ma=wb['V_Ma']
        
        types = [row[0].value for s in [slice("O59","O78")] for row in ws_ma[s]]

        nombres_vehiculos = types
        wb.close()

        #Copia del modelo.
        modelo_original = "./images/Modelo.xlsm"
        directorio,_ = os.path.split(self.path_file)
        modelo = os.path.join(directorio,'Reporte_GEH-R2.xlsm')
        shutil.copy2(modelo_original,modelo)

        vehicles_names = 'H8'
        intersection = 'B8'
        od = 'C8'
        vehicle_type = 'D8'
        volumes = 'E8'

        numero_total = len(data_excel)*len(data_excel[0])*len(data_excel[0][0][2])*len(data_excel[0][0][1])

        wb = xw.Book(modelo)
        sheet = wb.sheets['GEH']
        all_values = []
        insercion_data_inicio = time.perf_counter()
        for i in range(len(types)):
            sheet.range(vehicles_names).value = types[i]
            vehicles_names=vehicles_names[0]+str(9+i)

        for l in range(len(data_excel[0][0][2])): #Aquí estan 10 tipos de vehiculos
            for i in range(len(data_excel)): #Son 3: Número de excels
                for j in range(len(data_excel[i])): #Son 4: N, S, E, O
                    for k in range(len(data_excel[i][j][1])): #Aquí es por número de giros (en este caso 3)
                        all_values.append([data_excel[i][j][0],data_excel[i][j][1][k],types[l],data_excel[i][j][2][l][k]])

        for index in range(numero_total):
            sheet.range(intersection).value=all_values[index][0][7:]
            sheet.range(od).value=all_values[index][1]
            sheet.range(vehicle_type).value=all_values[index][2]
            sheet.range(volumes).value=all_values[index][3]
            intersection=intersection[0]+str(9+index)
            od=od[0]+str(9+index)
            vehicle_type=vehicle_type[0]+str(9+index)
            volumes=volumes[0]+str(9+index)
        
        wb.save()
        wb.app.quit()
        insercion_data_fin = time.perf_counter()

        print(f"Tiempo de escritura en excel = {insercion_data_fin-insercion_data_inicio:.02f} segundos.")
        #Podría colocar un return que modifique un label acá.

    def reporte(self):
        global nombres_vehiculos

        ##############################
        # OBTENCIÓN DE DATA SIMULADA #
        ##############################
        inicio_reading = time.perf_counter()
        directorio = os.path.dirname(self.path_file)
        patron = re.compile(r'_Node Results_\d{3}.att') #Busca todos los archivos que acaban en 3 dígitos y se filtra el mayor. Aggregated data need to be activated.

        archivos_att = [f for f in os.listdir(directorio) if patron.search(f)]

        archivo_mas_alto = max(archivos_att,key = obtener_numero_al_final)

        path = os.path.join(directorio,archivo_mas_alto)

        #LECTURA DE LA DATA SIMULADA
        with open(path,'r') as att_file:
            contador = 0
            data = []

            for line in att_file:
                line = line.strip()

                if line.startswith('$') and contador<3:
                    contador += 1

                if contador == 2:
                    data.append(line)

        data = data[:-1]
        data_str = '\n'.join(data)
        data_io = StringIO(data_str)
        df = pd.read_csv(data_io,delimiter=';')

        #Tramiento del DataFrame
        df = df.dropna(subset=['MOVEMENT\FROMLINK\ORIGEN','MOVEMENT\TOLINK\DESTINO'])
        df = df.reset_index(drop=True)
        df = df.loc[df.iloc[:,0] == 'AVG']
        df = df.reset_index(drop=True)

        final_reading=time.perf_counter()

        print(f"Tiempo de lectura de resultados = {final_reading-inicio_reading:.2f} segundos.")

        #####################################
        # ESCRITURA DE INFORMACIÓN EN EXCEL #
        #####################################

        inicio_writing=time.perf_counter()

        directorio,_ = os.path.split(self.path_file)
        modelo = os.path.join(directorio,'Reporte_GEH-R2.xlsm')

        wb = xw.Book(modelo)
        sheet = wb.sheets['GEH']

        intersection    = 'K8'
        origin          = 'L8'
        destiny         = 'M8'
        vehicle_type    = ['BC8','BD8','BE8','BF8','BG8','BH8','BI8','BJ8','BK8','BL8','BM8','BN8','BO8','BP8','BQ8','BR8','BS8','BT8','BU8','BV8']

        for index,row in df.iterrows():
            #Escritura
            sheet.range(intersection).value = row.iloc[2]
            sheet.range(origin).value = row.iloc[3]
            sheet.range(destiny).value = row.iloc[4]
            #Actualización de índices en el excel
            intersection=intersection[0]+str(9+index)
            origin=origin[0]+str(9+index)
            destiny=destiny[0]+str(9+index)
            #Impresión de data pura y dura
            for i in range(len(row[5:])):
                sheet.range(vehicle_type[i]).value=row.iloc[i+5]
            for j in range(len(vehicle_type)):
                vehicle_type[j]=vehicle_type[j][:2]+str(int(vehicle_type[j][2:])+1)

        wb.save()
        
        final_writing = time.perf_counter()
        print(f"Tiempo de escritura en excel = {inicio_writing-final_writing:.2f} segundos.")

    def livianos(self):
        self.LookAheadDistMin.setText("0")
        self.LookAheadDistMax.setText("250")
        self.LookBackDistMin.setText('0')
        self.LookBackDistMax.setText('150')
        self.W74ax.setText('1.20')
        self.W74bxAdd.setText('1.50')
        self.W74bxMult.setText('2.50')
        self.MaxDecelOwn.setText('-4.00')
        self.MaxDecelTrail.setText('-3.00')
        self.DecelRedDistOwn.setText('100.00')
        self.DecelRedDistTrail.setText('100.00')
        self.AccDecelOwn.setText('-1.00')
        self.AccDecelTrail.setText('-0.25')
        self.DiffusTm.setText('210.00')
        self.SafDistFactLnChg.setText('0.50')
        self.CoopDecel.setText('-3.00')
        self.MinCollTmGain.setText('2.00')
        self.MinSpeedForLat.setText('3.60')
        self.LatDistStandDef.setText('0.20')
        self.LatDistDrivDef.setText('1.00')

    def menores(self):
        self.LookAheadDistMin.setText("30")
        self.LookAheadDistMax.setText("250")
        self.LookBackDistMin.setText('30')
        self.LookBackDistMax.setText('150')
        self.W74ax.setText('0.20')
        self.W74bxAdd.setText('1.50')
        self.W74bxMult.setText('2.50')
        self.MaxDecelOwn.setText('-4.00')
        self.MaxDecelTrail.setText('-3.00')
        self.DecelRedDistOwn.setText('100.00')
        self.DecelRedDistTrail.setText('100.00')
        self.AccDecelOwn.setText('-1.00')
        self.AccDecelTrail.setText('-0.25')
        self.DiffusTm.setText('210.00')
        self.SafDistFactLnChg.setText('0.50')
        self.CoopDecel.setText('-5.00')
        self.MinCollTmGain.setText('2.00')
        self.MinSpeedForLat.setText('3.60')
        self.LatDistStandDef.setText('0.02')
        self.LatDistDrivDef.setText('0.20')

    def publicos(self):
        self.LookAheadDistMin.setText("0")
        self.LookAheadDistMax.setText("250")
        self.LookBackDistMin.setText('0')
        self.LookBackDistMax.setText('150')
        self.W74ax.setText('1.20')
        self.W74bxAdd.setText('1.50')
        self.W74bxMult.setText('2.50')
        self.MaxDecelOwn.setText('-4.00')
        self.MaxDecelTrail.setText('-3.00')
        self.DecelRedDistOwn.setText('100.00')
        self.DecelRedDistTrail.setText('100.00')
        self.AccDecelOwn.setText('-1.00')
        self.AccDecelTrail.setText('-0.25')
        self.DiffusTm.setText('210.00')
        self.SafDistFactLnChg.setText('0.50')
        self.CoopDecel.setText('-3.00')
        self.MinCollTmGain.setText('2.00')
        self.MinSpeedForLat.setText('3.60')
        self.LatDistStandDef.setText('0.20')
        self.LatDistDrivDef.setText('1.00')

    def cargas(self):
        self.LookAheadDistMin.setText("0")
        self.LookAheadDistMax.setText("250")
        self.LookBackDistMin.setText('0')
        self.LookBackDistMax.setText('150')
        self.W74ax.setText('1.20')
        self.W74bxAdd.setText('1.50')
        self.W74bxMult.setText('2.50')
        self.MaxDecelOwn.setText('-4.00')
        self.MaxDecelTrail.setText('-3.00')
        self.DecelRedDistOwn.setText('100.00')
        self.DecelRedDistTrail.setText('100.00')
        self.AccDecelOwn.setText('-1.00')
        self.AccDecelTrail.setText('-0.25')
        self.DiffusTm.setText('210.00')
        self.SafDistFactLnChg.setText('0.50')
        self.CoopDecel.setText('-3.00')
        self.MinCollTmGain.setText('2.00')
        self.MinSpeedForLat.setText('3.60')
        self.LatDistStandDef.setText('0.20')
        self.LatDistDrivDef.setText('1.00')

    def data_input(self):
        self.version10=self.checkBox.isChecked()
        self.version23=self.checkBox_2.isChecked()
        self.Ivehicle_type.append         (self.vehicle_type.text()) #Nro. de Driving Behavior

        self.ILookAheadDistMin.append     (self.LookAheadDistMin.text())
        self.ILookAheadDistMax.append     (self.LookAheadDistMax.text())
        self.ILookBackDistMin.append      (self.LookBackDistMin.text())
        self.ILookBackDistMax.append      (self.LookBackDistMax.text())
        self.IW74ax.append                (self.W74ax.text())
        self.IW74bxAdd.append             (self.W74bxAdd.text())
        self.IW74bxMult.append            (self.W74bxMult.text())
        self.IMaxDecelOwn.append          (self.MaxDecelOwn.text())
        self.IMaxDecelTrail.append        (self.MaxDecelTrail.text())
        self.IDecelRedDistOwn.append      (self.DecelRedDistOwn.text())
        self.IDecelRedDistTrail.append    (self.DecelRedDistTrail.text())
        self.IAccDecelOwn.append          (self.AccDecelOwn.text())
        self.IAccDecelTrail.append        (self.AccDecelTrail.text())
        self.IDiffusTm.append             (self.DiffusTm.text())
        self.ISafDistFactLnChg.append     (self.SafDistFactLnChg.text())
        self.ICoopDecel.append            (self.CoopDecel.text())
        self.IMinCollTmGain.append        (self.MinCollTmGain.text())
        self.IMinSpeedForLat.append       (self.MinSpeedForLat.text())
        self.ILatDistStandDef.append      (self.LatDistStandDef.text())
        self.ILatDistDrivDef.append       (self.LatDistDrivDef.text())

        #News
        self.IDesLatPos.append(self.DesLatPos.currentText())

        if self.ObsrvAdjLn.isChecked(): self.IObsrvAdjLn.append("true")
        else: self.IObsrvAdjLn.append("false")

        if self.DiamQueu.isChecked(): self.IDiamQueu.append("true")
        else: self.IDiamQueu.append("false")

        if self.ConsNextTurn.isChecked(): self.IConsNextTurn.append("true")
        else: self.IConsNextTurn.append("false")

        if self.OvtLDef.isChecked(): self.IOvtLDef.append("true")
        else: self.IOvtLDef.append("false")

        if self.OvtRDef.isChecked(): self.IOvtRDef.append("true")
        else: self.IOvtRDef.append("false")

        #CHECK
        #self.enviado.setText(f"SAVED: {self.vehicle_type.text()}")

    def ejecutar_programa(self):
        global enviado_contador
        ####INTRODUCCION DE DATA####
        vehicle_type        =self.Ivehicle_type
        LookAheadDistMin    =self.ILookAheadDistMin
        LookAheadDistMax    =self.ILookAheadDistMax
        LookBackDistMin     =self.ILookBackDistMin
        LookBackDistMax     =self.ILookBackDistMax

        ####WIEDEMANN 74####
        W74ax               =self.IW74ax
        W74bxAdd            =self.IW74bxAdd
        W74bxMult           =self.IW74bxMult

        ####CAMBIO DE CARRIL####
        MaxDecelOwn         =self.IMaxDecelOwn
        MaxDecelTrail       =self.IMaxDecelTrail
        DecelRedDistOwn     =self.IDecelRedDistOwn
        DecelRedDistTrail   =self.IDecelRedDistTrail
        AccDecelOwn         =self.IAccDecelOwn
        AccDecelTrail       =self.IAccDecelTrail
        DiffusTm            =self.IDiffusTm
        SafDistFactLnChg    =self.ISafDistFactLnChg
        CoopDecel           =self.ICoopDecel

        ####LATERAL####
        MinCollTmGain       =self.IMinCollTmGain
        MinSpeedForLat      =self.IMinSpeedForLat
        LatDistStandDef     =self.ILatDistStandDef
        LatDistDrivDef      =self.ILatDistDrivDef

        ####NEWS####
        DesLatPos           =self.IDesLatPos
        ObsrvAdjLn          =self.IObsrvAdjLn
        DiamQueu            =self.IDiamQueu
        ConsNextTurn        =self.IConsNextTurn
        OvtLDef             =self.IOvtLDef
        OvtRDef             =self.IOvtRDef

        #INICIO DE COM
        if self.version10:
            vissim = com.Dispatch('Vissim.Vissim.10')
        elif self.version23:
            vissim = com.Dispatch('Vissim.Vissim.23')
        
        #try:
        #    vissim = com.GetObject(Class="Vissim.Vissim.10")
        #except com.pywintypes.com_error as e:
        #    print("No se pudo conectar a Vissim:",e)

        ##################################
        # ENVIO DE INFORMACIÓN AL VISSIM #
        ##################################

        if vissim.Simulation.AttValue('IsRunning'):
            vissim.Simulation.RunSingleStep() #Pausa

        key = int(vehicle_type[1])

        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(LookAheadDistMin[0],LookAheadDistMin[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(LookAheadDistMax[0],LookAheadDistMax[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(LookBackDistMin[0],LookBackDistMin[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(LookBackDistMax[0],LookBackDistMax[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(W74ax[0],W74ax[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(W74bxAdd[0],W74bxAdd[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(W74bxMult[0],W74bxMult[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(MaxDecelOwn[0],MaxDecelOwn[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(MaxDecelTrail[0],MaxDecelTrail[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(DecelRedDistOwn[0],DecelRedDistOwn[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(DecelRedDistTrail[0],DecelRedDistTrail[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(AccDecelOwn[0],AccDecelOwn[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(AccDecelTrail[0],AccDecelTrail[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(DiffusTm[0],DiffusTm[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(SafDistFactLnChg[0],SafDistFactLnChg[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(CoopDecel[0],CoopDecel[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(MinCollTmGain[0],MinCollTmGain[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(MinSpeedForLat[0],MinSpeedForLat[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(LatDistStandDef[0],LatDistStandDef[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(LatDistDrivDef[0],LatDistDrivDef[1])
        #NEWS:
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(DesLatPos[0],DesLatPos[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(ObsrvAdjLn[0],ObsrvAdjLn[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(DiamQueu[0],DiamQueu[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(ConsNextTurn[0],ConsNextTurn[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(OvtLDef[0],OvtLDef[1])
        vissim.Net.DrivingBehaviors.ItemByKey(key).SetAttValue(OvtRDef[0],OvtRDef[1])

        #CHECK
        self.enviado.setText(f"ENVIADO {enviado_contador}")
        enviado_contador += 1

        #wins = gw.getWindowsWithTitle('Vissim')
        #print(wins)
        #wins[1].activate()

        #Definición de listas
        ####SEGUIMIENTO VEHICULAR####
        self.Ivehicle_type=["#Driving-Behavior"]
        self.ILookAheadDistMin=["LookAheadDistMin"]
        self.ILookAheadDistMax=["LookAheadDistMax"]
        self.ILookBackDistMin=["LookBackDistMin"]
        self.ILookBackDistMax=["LookBackDistMax"]

        ####WIEDEMANN 74####
        self.IW74ax=["W74ax"]
        self.IW74bxAdd=["W74bxAdd"]
        self.IW74bxMult=["W74bxMult"]

        ####CAMBIO DE CARRIL####
        self.IMaxDecelOwn=["MaxDecelOwn"]
        self.IMaxDecelTrail=["MaxDecelTrail"]
        self.IDecelRedDistOwn=["DecelRedDistOwn"]
        self.IDecelRedDistTrail=["DecelRedDistTrail"]
        self.IAccDecelOwn=["AccDecelOwn"]
        self.IAccDecelTrail=["AccDecelTrail"]
        self.IDiffusTm=["DiffusTm"]
        self.ISafDistFactLnChg=["SafDistFactLnChg"]
        self.ICoopDecel=["CoopDecel"]

        ####LATERAL####
        self.IMinCollTmGain=["MinCollTmGain"]
        self.IMinSpeedForLat=["MinSpeedForLat"]
        self.ILatDistStandDef=["LatDistStandDef"]
        self.ILatDistDrivDef=["LatDistDrivDef"]
        self.IConsNextTurn=["ConsNextTurn"]

        #NEWS
        self.IDesLatPos=["DesLatPos"]
        self.IObsrvAdjLn=["ObsrvAdjLn"]
        self.IDiamQueu=["DiamQueu"]
        self.IConsNextTurn=["ConsNextTurn"]
        self.IOvtLDef=["OvtLDef"]
        self.IOvtRDef=["OvtRDef"]
        
        #Play Simulation after have changed attributes

        print("Cambios de parámetros exitoso")

    def fijars(self):
        global param_groups

        default_dictionary = {
            1: self.LookAheadDistMin.text(),
            2: self.LookAheadDistMax.text(),
            3: self.LookBackDistMin.text(),
            4: self.LookBackDistMax.text(),
            5: self.W74ax.text(),
            6: self.W74bxAdd.text(),
            7: self.W74bxMult.text(),
            8: self.MaxDecelOwn.text(),
            9: self.MaxDecelTrail.text(),
            10: self.DecelRedDistOwn.text(),
            11: self.DecelRedDistTrail.text(),
            12: self.AccDecelOwn.text(),
            13: self.AccDecelTrail.text(),
            14: self.DiffusTm.text(),
            15: self.SafDistFactLnChg.text(),
            16: self.CoopDecel.text(),
            17: self.DesLatPos.currentText(),
            18: str(self.ObsrvAdjLn.isChecked()),
            19: str(self.DiamQueu.isChecked()),
            20: str(self.ConsNextTurn.isChecked()),
            21: self.MinCollTmGain.text(),
            22: self.MinSpeedForLat.text(),
            23: str(self.OvtLDef.isChecked()),
            24: str(self.OvtRDef.isChecked()),
            25: self.LatDistStandDef.text(),
            26: self.LatDistDrivDef.text(),
        }

        param_groups[self.vehicle_type.text()] = default_dictionary

        print(param_groups)

        self.enviado.setText(f"Send {self.vehicle_type.text()}")

    def export_params_2_excel(self):
        global param_groups
        original_path = "./images/Parametros_Model.xlsx"
        directorio,_ = os.path.split(self.path_file)
        modelo = os.path.join(directorio,'Parametros_Calibracion.xlsx')

        shutil.copy2(original_path,modelo)
        workbook = load_workbook(modelo)
        worksheet = workbook['Hoja1']

        ######################
        # INGRESO DE VALORES #
        ######################
        files = [5,6,8,9,12,13,14,17,18,19,20,21,22,23,25,26,31,32,33,34,35,36,38,39,38,39]
        
        for index1,(_,group) in enumerate(param_groups.items()):
            worksheet.cell(4,6+index1*2).value = int(index1+1)
            for i,(_,value) in enumerate(group.items()):
                if i<24:
                    try:
                        worksheet.cell(files[i],6+index1*2).value=float(value)
                    except ValueError:
                        worksheet.cell(files[i],6+index1*2).value=str(value)
                elif i==24:
                    try:
                        worksheet.cell(files[i],7+index1*2).value=float(value)
                    except ValueError:
                        worksheet.cell(files[i],7+index1*2).value=str(value)
                else:
                    #print(files[i],7+index1*2)
                    try:
                        worksheet.cell(files[i],7+index1*2).value=float(value)
                    except ValueError:
                        worksheet.cell(files[i],7+index1*2).value=str(value)

        print("Reporte de parámetros generados")
        workbook.save(modelo)
        workbook.close()
        self.enviado.setText("PÁRAMETROS OK!")

def main():
    app = QApplication([])
    window = MiVentana()
    window.show()
    app.exec_()

if __name__ == "__main__": 
    main()