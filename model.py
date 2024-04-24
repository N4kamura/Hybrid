import win32com.client as com
import numpy as np
from pathlib import Path
import xml.etree.ElementTree as ET

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
        name = vehicleClass.attrib["name"]
        veh_classes[no] = name

    vehClasses_evaluation = []
    for intObjectRef in root.findall("./evaluation/vehClasses/intObjectRef"):
        vehClasses_evaluation.append(int(intObjectRef.attrib["no"]))

    no_vehClasses = len(vehClasses_evaluation)

    if vehClasses_total != vehClasses_evaluation:
        print(f"Key vehClasses Total:\n", vehClasses_total)
        print(f"Key vehClasses Evaluation:\n", vehClasses_evaluation)
        raise ValueError("Total and evaluation vehicle classes do not match.")

    return veh_classes, no_vehClasses, nodes_dict

def get_results(
        Vissim: com.Dispatch, #Vissim object
        num_comp: int, #Number of vehicle compositions used
        no_node: int, #Node number
        ) -> tuple[np.ndarray, tuple, tuple] :
    attributes = (f"VEHS(Avg,1,{i})" for i in range(1,num_comp+1))

    NODE_RES = np.nan_to_num(
        np.array(Vissim.Net.Nodes.ItemByKey(no_node).Movements.GetMultipleAttributes(
            [attr for attr in attributes]),
            dtype=int)
    )
    
    ORIGIN = Vissim.Net.Nodes.ItemByKey(no_node).Movements.GetMultipleAttributes(["FROMLINK\ORIGEN"])
    DESTINY = Vissim.Net.Nodes.ItemByKey(no_node).Movements.GetMultipleAttributes(["TOLINK\DESTINO"])

    print(NODE_RES, type(NODE_RES))
    print(ORIGIN, type(ORIGIN))
    print(DESTINY, type(DESTINY))

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
        ws.range(od_row, 12).value = int(o)
        ws.range(od_row, 13).value = int(d)
        od_row += 1

    #Getting the number of vehicles compositions from excel
    veh_comp_excel = [elem for elem in ws.range('H8:H27').value]
    index = next((i for i, elem in enumerate(veh_comp_excel) if elem in ('N','n',None)), len(veh_comp_excel))

    veh_comp_excel = [elem for elem in ws.range('H8:H27').value][:index]

    if len(veh_comp_excel) != len(veh_classes):
        print(f"Vehicle classes excel = {len(veh_comp_excel)}")
        print(f"Vehicle classes vissim = {len(veh_classes)}")
        raise ValueError("Number of vehicle classes between excel and vissim do not match.")
    
    count_row = nro_row
    for i in range(len(veh_comp_excel)):
        check = False
        for key, value in veh_classes.items():
            if value == veh_comp_excel[i]:
                for index,row in enumerate(NODE_RES):
                    ws.range(count_row+index, 55+key-1).value = row[key-1]
                check = True
                break
        if not check:
            print("Excel vehicle class = ", veh_comp_excel[i])
            raise ValueError("Vehicle class not found in vissim.")
        count_row += 1
        
    return count_row