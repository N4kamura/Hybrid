import win32com.client as com
import numpy as np
from pathlib import Path
import xml.etree.ElementTree as ET
from unidecode import unidecode

def get_veh_classes(
        xml_path: str | Path, #Path of vissim file
        ) -> tuple[dict, int]:
    
    tree = ET.parse(xml_path)
    root = tree.getroot()
    veh_classes = {}
    vehClasses_total = []
    
    #Looking uda for code in nodes:
    for uda in root.findall("./userDefinedAttributes/userDefinedAttribute"):
        codigo = uda.attrib["nameLong"]
        if codigo == "Código":
            no_uda = uda.attrib["no"]
            break

    nodes_dict = {}
    for node in root.findall("./nodes/node"):
        number_node = int(node.attrib["no"])
        uda = node.find("./uda")
        if uda.attrib["key"] == no_uda:
            node_code = uda.attrib["value"]
        nodes_dict[number_node] = node_code

    for vehicleClass in root.findall("./vehicleClasses/vehicleClass"):
        no = int(vehicleClass.attrib["no"])
        vehClasses_total.append(no)
        name = unidecode(vehicleClass.attrib["name"]).upper()
        veh_classes[no] = name

    vehClasses_evaluation = []
    for intObjectRef in root.findall("./evaluation/vehClasses/intObjectRef"):
        vehClasses_evaluation.append(int(intObjectRef.attrib["key"]))

    no_vehClasses = len(vehClasses_evaluation)

    return veh_classes, no_vehClasses, nodes_dict

def get_results( 
        Vissim: com.Dispatch, #Vissim object
        no_vehClasses_evaluation: int, #Number of vehicle compositions used
        no_node: int, #Node number
        ) -> tuple[np.ndarray, tuple, tuple]:
    attributes = (f"VEHS(Avg,1,{i})" for i in range(1,no_vehClasses_evaluation+1))

    NODE_RES = np.nan_to_num(
        np.array(Vissim.Net.Nodes.ItemByKey(no_node).Movements.GetMultipleAttributes(
            [attr for attr in attributes]),
            dtype=int)
    )
    
    ORIGIN = Vissim.Net.Nodes.ItemByKey(no_node).Movements.GetMultipleAttributes(["FROMLINK\ORIGEN"])
    DESTINY = Vissim.Net.Nodes.ItemByKey(no_node).Movements.GetMultipleAttributes(["TOLINK\DESTINO"])

    return NODE_RES, ORIGIN, DESTINY

def writing_excel(
        NODE_RES: np.ndarray, #Node results
        ORIGIN: tuple, #Origin
        DESTINY: tuple, #Destiny
        CODE: str, #Intersection code
        veh_classes: dict, #dictionary with no & name of vehicle classes
        ws: tuple, #Path of the vissim file
        nro_row: int, #Row number
        ) -> int:

    od_row = nro_row
    for o, d in zip(ORIGIN[:-1], DESTINY[:-1]): #El número de filas aquí es igual que abajo. TODO: Se pueden fusionar
        ws.range(od_row, 11).value = CODE
        ws.range(od_row, 12).value = int(o[0])
        ws.range(od_row, 13).value = int(d[0])
        od_row += 1

    #Getting the number of vehicles compositions from excel
    veh_comp_excel = [elem for elem in ws.range('H8:H27').value]
    index = next((i for i, elem in enumerate(veh_comp_excel) if elem in ('N','n',None)), len(veh_comp_excel))

    veh_comp_excel = [elem for elem in ws.range('H8:H27').value][:index]
    
    veh_not_considered = []

    for key, value in veh_classes.items():
        check = True
        for i in range(len(veh_comp_excel)):
            if value == veh_comp_excel[i]: #Encontro que de los totales del vissim hay uno en el excel
                for index,row in enumerate(NODE_RES[:-1]):
                    ws.range(nro_row+index, 55+i).value = row[key-1]
                check = False
                break
        if check:
            veh_not_considered.append(value)
    count_row = len(NODE_RES[:-1])
        
    return count_row, veh_not_considered