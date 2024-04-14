import win32com.client as com
import json
from pywintypes import com_error
import numpy as np
import xml.etree.ElementTree as ET


""" 
Sub attribute 1: SimulationRun
Sub attribute 2: TimeInterval
Sub attribute 3: Vehicle Class"""

layout_path = r"C:\Users\dacan\Downloads\layout.layx"

def get_results(vissim_file, vissim_version, num_runs, sim_res, num_cores) -> json:
    tree = ET.parse(vissim_file)
    network_tag = tree.getroot()
    
    veh_class_no = []
    for vehicleClass in network_tag.findall("./vehicleClasses/vehicleClass"):
        no = vehicleClass.attrib["no"]
        veh_class_no.append(no)

    vissim = com.Dispatch(f"Vissim.Vissim.{vissim_version}")
    
    try:
        vissim.LoadLayout(layout_path)
    except com_error as inst:
        print(f"Could not load layout file:\n{layout_path}")
        raise inst

    try:
        vissim.LoadNet(vissim_file)
    except com_error as inst:
        print(f"Could not load network file:\n{vissim_file}")
        raise inst
    
    #Simulation
    vissim.Simulation.SetAttValue("NumRuns",num_runs)
    vissim.Simulation.SetAttValue("SimRes", sim_res)
    vissim.Simulation.SetAttValue("NumCores", num_cores)
    vissim.Simulation.SetAttValue("SimPeriod", 5400)

    #Evaluation
    vissim.Evaluation.SetAttValue("NodeResCollectData", True)
    vissim.Evaluation.SetAttValue("NodeResFromTime", 1800)
    vissim.Evaluation.SetAttValue("NodeResToTime", 5400)
    vissim.Evaluation.SetAttValue("NodeResInterval", 3600)

    #Graphics
    vissim.Graphics.SetAttValue("QuickMode", True)
    
    #Start simulation
    vissim.Simulation.RunContinuous()

    NODES_TOTRES = [f"Vehs(Avg,1,{veh_type})" for veh_type in veh_class_no]

    node_totres = np.nan_to_num(
        np.array(vissim.Net.Nodes.GetMultipleAttributes(
            [fr"TotRes\{attr}" for attr in NODES_TOTRES]),
            dtype=float)
    )

    table = {"nodes_totres": tuple(map(tuple, node_totres))}

    print(table)

if __name__ == '__main__':
    vissim_path = r"C:\Users\dacan\Downloads\basic_layout.inpx"
    get_results(vissim_path, 24, 1, 8, 6)
    
