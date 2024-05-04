from vissim.clases import *

def show_info(vissim, ui, version, error_message):
    """ Extrae información y lo muestra en la interfaz. """

    key = int(ui.vehicle_type.text())

    try:
        v = vissim.Net.DrivingBehaviors.ItemByKey(key)
    except Exception as e:
        return error_message.showMessage(f"No se pudo encontrar el comportamiento vehicular: {key}")

    if version == 24:
        info = extractor_24(v)
    elif version == 10:
        info = extractor_10(v)

    #Muestra variables en la interfaz.

    ui.LookAheadDistMin.setText       (str(info.LookAheadDistMin))
    ui.LookAheadDistMax.setText       (str(info.LookAheadDistMax))
    ui.LookBackDistMin.setText        (str(info.LookBackDistMin))
    ui.LookBackDistMax.setText        (str(info.LookBackDistMax))
    ui.W74ax.setText                  (str(info.W74ax))
    ui.W74bxAdd.setText               (str(info.W74bxAdd))
    ui.W74bxMult.setText              (str(info.W74bxMult))
    ui.MaxDecelOwn.setText            (str(info.MaxDecelOwn))
    ui.MaxDecelTrail.setText          (str(info.MaxDecelTrail))
    ui.DecelRedDistOwn.setText        (str(info.DecelRedDistOwn))
    ui.DecelRedDistTrail.setText      (str(info.DecelRedDistTrail))
    ui.AccDecelOwn.setText            (str(info.AccDecelOwn))
    ui.AccDecelTrail.setText          (str(info.AccDecelTrail))
    ui.DiffusTm.setText               (str(info.DiffusTm))
    ui.SafDistFactLnChg.setText       (str(info.SafDistFactLnChg))
    ui.CoopDecel.setText              (str(info.CoopDecel))
    ui.MinCollTmGain.setText          (str(info.MinCollTmGain))
    ui.MinSpeedForLat.setText         (str(info.MinSpeedForLat))
    ui.LatDistStandDef.setText        (str(info.LatDistStandDef))
    ui.LatDistDrivDef.setText         (str(info.LatDistDrivDef))

    DesLatPos_ui = info.DesLatPos
    if DesLatPos_ui == 'MIDDLE':
        ui.DesLatPos.setCurrentIndex(0)
    elif DesLatPos_ui == 'ANY':
        ui.DesLatPos.setCurrentIndex(1)
    elif DesLatPos_ui == 'RIGHT':
        ui.DesLatPos.setCurrentIndex(2)
    elif DesLatPos_ui == 'LEFT':
        ui.DesLatPos.setCurrentIndex(3)

    if info.ObsrvAdjLn == 0:
        ui.ObsrvAdjLn.setChecked(False)
    else: ui.ObsrvAdjLn.setChecked(True)

    if info.DiamQueu == 0:
        ui.DiamQueu.setChecked(False)
    else: ui.DiamQueu.setChecked(True)

    if info.ConsNextTurn == 0:
        ui.ConsNextTurn.setChecked(False)
    else: ui.ConsNextTurn.setChecked(True)

    if info.OvtLDef == 0:
        ui.OvtLDef.setChecked(False)
    else: ui.OvtLDef.setChecked(True)
    
    if info.OvtRDef == 0:
        ui.OvtRDef.setChecked(False)
    else: ui.OvtRDef.setChecked(True)

    if ui.checkBox_2.isChecked():
        if info.Zipper == 0:
            ui.checkBox_3.setChecked(False)
        else:
            ui.checkBox_3.setChecked(True)
            ui.ZipperMinSpeed.setText(str(info.ZipperMinSpeed))

def extractor_10(v):
    """ Extrae información de vissim 10 y lo guarda en un diccionario. """
    data = vissimData10(
        LookAheadDistMin = v.AttValue('LookAheadDistMin'),
        LookAheadDistMax = v.AttValue('LookAheadDistMax'),
        LookBackDistMin =  v.AttValue('LookBackDistMin'),
        LookBackDistMax =  v.AttValue('LookBackDistMax'),
        W74ax =            v.AttValue('W74ax'),
        W74bxAdd =         v.AttValue('W74bxAdd'),
        W74bxMult =        v.AttValue('W74bxMult'),
        MaxDecelOwn =      v.AttValue('MaxDecelOwn'),
        MaxDecelTrail =    v.AttValue('MaxDecelTrail'),
        DecelRedDistOwn =  v.AttValue('DecelRedDistOwn'),
        DecelRedDistTrail =v.AttValue('DecelRedDistTrail'),
        AccDecelOwn =      v.AttValue('AccDecelOwn'),
        AccDecelTrail =    v.AttValue('AccDecelTrail'),
        DiffusTm =         v.AttValue('DiffusTm'),
        SafDistFactLnChg = v.AttValue('SafDistFactLnChg'),
        CoopDecel =        v.AttValue('CoopDecel'),
        MinCollTmGain =    v.AttValue('MinCollTmGain'),
        MinSpeedForLat =   v.AttValue('MinSpeedForLat'),
        LatDistStandDef =  v.AttValue('LatDistStandDef'),
        LatDistDrivDef =   v.AttValue('LatDistDrivDef'),
        DesLatPos =        v.AttValue('DesLatPos'),
        ObsrvAdjLn =       v.AttValue('ObsrvAdjLn'),
        DiamQueu =         v.AttValue('DiamQueu'),
        ConsNextTurn =     v.AttValue('ConsNextTurn'),
        OvtLDef =          v.AttValue('OvtLDef'),
        OvtRDef =          v.AttValue('OvtRDef'),
    )

    return data

def extractor_24(v):
    """ Extrae información de vissim 24 y lo guarda en un diccionario. """
    data = vissimData24(
        LookAheadDistMin = v.AttValue('LookAheadDistMin'),
        LookAheadDistMax = v.AttValue('LookAheadDistMax'),
        LookBackDistMin =  v.AttValue('LookBackDistMin'),
        LookBackDistMax =  v.AttValue('LookBackDistMax'),
        W74ax =            v.AttValue('W74ax'),
        W74bxAdd =         v.AttValue('W74bxAdd'),
        W74bxMult =        v.AttValue('W74bxMult'),
        MaxDecelOwn =      v.AttValue('MaxDecelOwn'),
        MaxDecelTrail =    v.AttValue('MaxDecelTrail'),
        DecelRedDistOwn =  v.AttValue('DecelRedDistOwn'),
        DecelRedDistTrail =v.AttValue('DecelRedDistTrail'),
        AccDecelOwn =      v.AttValue('AccDecelOwn'),
        AccDecelTrail =    v.AttValue('AccDecelTrail'),
        DiffusTm =         v.AttValue('DiffusTm'),
        SafDistFactLnChg = v.AttValue('SafDistFactLnChg'),
        CoopDecel =        v.AttValue('CoopDecel'),
        MinCollTmGain =    v.AttValue('MinCollTmGain'),
        MinSpeedForLat =   v.AttValue('MinSpeedForLat'),
        LatDistStandDef =  v.AttValue('LatDistStandDef'),
        LatDistDrivDef =   v.AttValue('LatDistDrivDef'),
        DesLatPos =        v.AttValue('DesLatPos'),
        ObsrvAdjLn =       v.AttValue('ObsrvAdjLn'),
        DiamQueu =         v.AttValue('DiamQueu'),
        ConsNextTurn =     v.AttValue('ConsNextTurn'),
        OvtLDef =          v.AttValue('OvtLDef'),
        OvtRDef =          v.AttValue('OvtRDef'),
        Zipper =           v.AttValue('Zipper'),
        ZipperMinSpeed =   v.AttValue('ZipperMinSpeed'),
    )

    return data

def from_ui_24(ui):
    data = vissimInfo24(
        vehicleType         = ui.vehicle_type.text(),
        LookAheadDistMin    = ui.LookAheadDistMin.text(),
        LookAheadDistMax    = ui.LookAheadDistMax.text(),
        LookBackDistMin     = ui.LookBackDistMin.text(),
        LookBackDistMax     = ui.LookBackDistMax.text(),
        W74ax               = ui.W74ax.text(),
        W74bxAdd            = ui.W74bxAdd.text(),
        W74bxMult           = ui.W74bxMult.text(),
        MaxDecelOwn         = ui.MaxDecelOwn.text(),
        MaxDecelTrail       = ui.MaxDecelTrail.text(),
        DecelRedDistOwn     = ui.DecelRedDistOwn.text(),
        DecelRedDistTrail   = ui.DecelRedDistTrail.text(),
        AccDecelOwn         = ui.AccDecelOwn.text(),
        AccDecelTrail       = ui.AccDecelTrail.text(),
        DiffusTm            = ui.DiffusTm.text(),
        SafDistFactLnChg    = ui.SafDistFactLnChg.text(),
        CoopDecel           = ui.CoopDecel.text(),
        MinCollTmGain       = ui.MinCollTmGain.text(),
        MinSpeedForLat      = ui.MinSpeedForLat.text(),
        LatDistStandDef     = ui.LatDistStandDef.text(),
        LatDistDrivDef      = ui.LatDistDrivDef.text(),
        DesLatPos           = ui.DesLatPos.currentText(),
        ObsrvAdjLn          = str(ui.ObsrvAdjLn.isChecked()).lower(),
        DiamQueu            = str(ui.DiamQueu.isChecked()).lower(),
        ConsNextTurn        = str(ui.ConsNextTurn.isChecked()).lower(),
        OvtLDef             = str(ui.OvtLDef.isChecked()).lower(),
        OvtRDef             = str(ui.OvtRDef.isChecked()).lower(),
        Zipper              = str(ui.checkBox_3.isChecked()).lower(),
        ZipperMinSpeed      = ui.ZipperMinSpeed.text(),
    )

    return data

def from_ui_10(ui):
    data = vissimInfo10(
        vehicleType         = ui.vehicle_type.text(),
        LookAheadDistMin    = ui.LookAheadDistMin.text(),
        LookAheadDistMax    = ui.LookAheadDistMax.text(),
        LookBackDistMin     = ui.LookBackDistMin.text(),
        LookBackDistMax     = ui.LookBackDistMax.text(),
        W74ax               = ui.W74ax.text(),
        W74bxAdd            = ui.W74bxAdd.text(),
        W74bxMult           = ui.W74bxMult.text(),
        MaxDecelOwn         = ui.MaxDecelOwn.text(),
        MaxDecelTrail       = ui.MaxDecelTrail.text(),
        DecelRedDistOwn     = ui.DecelRedDistOwn.text(),
        DecelRedDistTrail   = ui.DecelRedDistTrail.text(),
        AccDecelOwn         = ui.AccDecelOwn.text(),
        AccDecelTrail       = ui.AccDecelTrail.text(),
        DiffusTm            = ui.DiffusTm.text(),
        SafDistFactLnChg    = ui.SafDistFactLnChg.text(),
        CoopDecel           = ui.CoopDecel.text(),
        MinCollTmGain       = ui.MinCollTmGain.text(),
        MinSpeedForLat      = ui.MinSpeedForLat.text(),
        LatDistStandDef     = ui.LatDistStandDef.text(),
        LatDistDrivDef      = ui.LatDistDrivDef.text(),
        DesLatPos           = ui.DesLatPos.currentText(),
        ObsrvAdjLn          = str(ui.ObsrvAdjLn.isChecked()).lower(),
        DiamQueu            = str(ui.DiamQueu.isChecked()).lower(),
        ConsNextTurn        = str(ui.ConsNextTurn.isChecked()).lower(),
        OvtLDef             = str(ui.OvtLDef.isChecked()).lower(),
        OvtRDef             = str(ui.OvtRDef.isChecked()).lower(),
    )

    return data
