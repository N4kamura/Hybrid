from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow,QLabel,QFileDialog,QButtonGroup
from PyQt5.QtGui import QPixmap
import win32com.client as com
import os
from openpyxl import load_workbook
import shutil
import warnings
from writing import writing_campo, writing_model
from openpyxl import load_workbook

warnings.filterwarnings("ignore", category=DeprecationWarning)
nombres_vehiculos = []
enviado_contador = 1
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

        #Checking buttons:
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

        writing_campo(self.path_file, turno)

        print("Proceso terminado.")

    def reporte(self):
        global nombres_vehiculos
        writing_model(self.path_file)

        print("Proceso terminado.")
        
    def livianos(self):
        guide_path = "./images/Parametros_Guia.xlsx"
        wb = load_workbook(guide_path, read_only=True, data_only=True)
        ws = wb['Hoja1']

        self.LookAheadDistMin.setText   (str(ws.cell(5,6).value))
        self.LookAheadDistMax.setText   (str(ws.cell(6,6).value))
        self.LookBackDistMin.setText    (str(ws.cell(8,6).value))
        self.LookBackDistMax.setText    (str(ws.cell(9,6).value))
        self.W74ax.setText              (str(ws.cell(12,6).value))
        self.W74bxAdd.setText           (str(ws.cell(13,6).value))
        self.W74bxMult.setText          (str(ws.cell(14,6).value))
        self.MaxDecelOwn.setText        (str(ws.cell(17,6).value))
        self.MaxDecelTrail.setText      (str(ws.cell(18,6).value))
        self.DecelRedDistOwn.setText    (str(ws.cell(19,6).value))
        self.DecelRedDistTrail.setText  (str(ws.cell(20,6).value))
        self.AccDecelOwn.setText        (str(ws.cell(21,6).value))
        self.AccDecelTrail.setText      (str(ws.cell(22,6).value))
        self.DiffusTm.setText           (str(ws.cell(23,6).value))
        self.SafDistFactLnChg.setText   (str(ws.cell(25,6).value))
        self.CoopDecel.setText          (str(ws.cell(26,6).value))
        self.MinCollTmGain.setText      (str(ws.cell(35,6).value))
        self.MinSpeedForLat.setText     (str(ws.cell(36,6).value))
        self.LatDistStandDef.setText    (str(ws.cell(38,7).value))
        self.LatDistDrivDef.setText     (str(ws.cell(39,7).value))
        wb.close()

    def menores(self):
        guide_path = "./images/Parametros_Guia.xlsx"
        wb = load_workbook(guide_path, read_only=True, data_only=True)
        ws = wb['Hoja1']

        self.LookAheadDistMin.setText   (str(ws.cell(5,8).value))
        self.LookAheadDistMax.setText   (str(ws.cell(6,8).value))
        self.LookBackDistMin.setText    (str(ws.cell(8,8).value))
        self.LookBackDistMax.setText    (str(ws.cell(9,8).value))
        self.W74ax.setText              (str(ws.cell(12,8).value))
        self.W74bxAdd.setText           (str(ws.cell(13,8).value))
        self.W74bxMult.setText          (str(ws.cell(14,8).value))
        self.MaxDecelOwn.setText        (str(ws.cell(17,8).value))
        self.MaxDecelTrail.setText      (str(ws.cell(18,8).value))
        self.DecelRedDistOwn.setText    (str(ws.cell(19,8).value))
        self.DecelRedDistTrail.setText  (str(ws.cell(20,8).value))
        self.AccDecelOwn.setText        (str(ws.cell(21,8).value))
        self.AccDecelTrail.setText      (str(ws.cell(22,8).value))
        self.DiffusTm.setText           (str(ws.cell(23,8).value))
        self.SafDistFactLnChg.setText   (str(ws.cell(25,8).value))
        self.CoopDecel.setText          (str(ws.cell(26,8).value))
        self.MinCollTmGain.setText      (str(ws.cell(35,8).value))
        self.MinSpeedForLat.setText     (str(ws.cell(36,8).value))
        self.LatDistStandDef.setText    (str(ws.cell(38,9).value))
        self.LatDistDrivDef.setText     (str(ws.cell(39,9).value))
        wb.close()

    def publicos(self):
        guide_path = "./images/Parametros_Guia.xlsx"
        wb = load_workbook(guide_path, read_only=True, data_only=True)
        ws = wb['Hoja1']

        self.LookAheadDistMin.setText   (str(ws.cell(5,10).value))
        self.LookAheadDistMax.setText   (str(ws.cell(6,10).value))
        self.LookBackDistMin.setText    (str(ws.cell(8,10).value))
        self.LookBackDistMax.setText    (str(ws.cell(9,10).value))
        self.W74ax.setText              (str(ws.cell(12,10).value))
        self.W74bxAdd.setText           (str(ws.cell(13,10).value))
        self.W74bxMult.setText          (str(ws.cell(14,10).value))
        self.MaxDecelOwn.setText        (str(ws.cell(17,10).value))
        self.MaxDecelTrail.setText      (str(ws.cell(18,10).value))
        self.DecelRedDistOwn.setText    (str(ws.cell(19,10).value))
        self.DecelRedDistTrail.setText  (str(ws.cell(20,10).value))
        self.AccDecelOwn.setText        (str(ws.cell(21,10).value))
        self.AccDecelTrail.setText      (str(ws.cell(22,10).value))
        self.DiffusTm.setText           (str(ws.cell(23,10).value))
        self.SafDistFactLnChg.setText   (str(ws.cell(25,10).value))
        self.CoopDecel.setText          (str(ws.cell(26,10).value))
        self.MinCollTmGain.setText      (str(ws.cell(35,10).value))
        self.MinSpeedForLat.setText     (str(ws.cell(36,10).value))
        self.LatDistStandDef.setText    (str(ws.cell(38,11).value))
        self.LatDistDrivDef.setText     (str(ws.cell(39,11).value))
        wb.close()

    def cargas(self):
        guide_path = "./images/Parametros_Guia.xlsx"
        wb = load_workbook(guide_path, read_only=True, data_only=True)
        ws = wb['Hoja1']

        self.LookAheadDistMin.setText   (str(ws.cell(5,12).value))
        self.LookAheadDistMax.setText   (str(ws.cell(6,12).value))
        self.LookBackDistMin.setText    (str(ws.cell(8,12).value))
        self.LookBackDistMax.setText    (str(ws.cell(9,12).value))
        self.W74ax.setText              (str(ws.cell(12,12).value))
        self.W74bxAdd.setText           (str(ws.cell(13,12).value))
        self.W74bxMult.setText          (str(ws.cell(14,12).value))
        self.MaxDecelOwn.setText        (str(ws.cell(17,12).value))
        self.MaxDecelTrail.setText      (str(ws.cell(18,12).value))
        self.DecelRedDistOwn.setText    (str(ws.cell(19,12).value))
        self.DecelRedDistTrail.setText  (str(ws.cell(20,12).value))
        self.AccDecelOwn.setText        (str(ws.cell(21,12).value))
        self.AccDecelTrail.setText      (str(ws.cell(22,12).value))
        self.DiffusTm.setText           (str(ws.cell(23,12).value))
        self.SafDistFactLnChg.setText   (str(ws.cell(25,12).value))
        self.CoopDecel.setText          (str(ws.cell(26,12).value))
        self.MinCollTmGain.setText      (str(ws.cell(35,12).value))
        self.MinSpeedForLat.setText     (str(ws.cell(36,12).value))
        self.LatDistStandDef.setText    (str(ws.cell(38,13).value))
        self.LatDistDrivDef.setText     (str(ws.cell(39,13).value))
        wb.close()

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
        try:
            directorio,_ = os.path.split(self.path_file)
        except AttributeError:
            return print("Debes seleccionar la ubicación del archivo Vissim para que se guarde allí la excel.")
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