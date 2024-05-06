from dataclasses import dataclass

@dataclass
class vissimInfo24:
    vehicleType: str
    #Following
    LookAheadDistMin: str
    LookAheadDistMax: str
    LookBackDistMin: str
    LookBackDistMax: str
    #W74
    W74ax: str
    W74bxAdd: str
    W74bxMult: str
    #Lane change
    MaxDecelOwn: str
    MaxDecelTrail: str
    DecelRedDistOwn: str
    DecelRedDistTrail: str
    AccDecelOwn: str
    AccDecelTrail: str
    DiffusTm: str
    SafDistFactLnChg: str
    CoopDecel: str
    #Lateral behavior
    MinCollTmGain: str
    MinSpeedForLat: str
    LatDistStandDef: str
    LatDistDrivDef: str
    #Position
    DesLatPos: str
    ObsrvAdjLn: str
    DiamQueu: str
    ConsNextTurn: str
    OvtLDef: str
    OvtRDef: str
    #New
    Zipper: str
    ZipperMinSpeed: str

@dataclass
class vissimInfo10:
    vehicleType: str
    #Following
    LookAheadDistMin: str
    LookAheadDistMax: str
    LookBackDistMin: str
    LookBackDistMax: str
    #W74
    W74ax: str
    W74bxAdd: str
    W74bxMult: str
    #Lane change
    MaxDecelOwn: str
    MaxDecelTrail: str
    DecelRedDistOwn: str
    DecelRedDistTrail: str
    AccDecelOwn: str
    AccDecelTrail: str
    DiffusTm: str
    SafDistFactLnChg: str
    CoopDecel: str
    #Lateral behavior
    MinCollTmGain: str
    MinSpeedForLat: str
    LatDistStandDef: str
    LatDistDrivDef: str
    #Position
    DesLatPos: str
    ObsrvAdjLn: str
    DiamQueu: str
    ConsNextTurn: str
    OvtLDef: str
    OvtRDef: str

@dataclass
class vissimData24:
    LookAheadDistMin: str
    LookAheadDistMax: str
    LookBackDistMin: str
    LookBackDistMax: str
    W74ax: str
    W74bxAdd: str
    W74bxMult: str
    MaxDecelOwn: str
    MaxDecelTrail: str
    DecelRedDistOwn: str
    DecelRedDistTrail: str
    AccDecelOwn: str
    AccDecelTrail: str
    DiffusTm: str
    SafDistFactLnChg: str
    CoopDecel: str
    MinCollTmGain: str
    MinSpeedForLat: str
    LatDistStandDef: str
    LatDistDrivDef: str
    DesLatPos: str
    ObsrvAdjLn: str
    DiamQueu: str
    ConsNextTurn: str
    OvtLDef: str
    OvtRDef: str
    Zipper: str
    ZipperMinSpeed: str

@dataclass
class vissimData10:
    LookAheadDistMin: str
    LookAheadDistMax: str
    LookBackDistMin: str
    LookBackDistMax: str
    W74ax: str
    W74bxAdd: str
    W74bxMult: str
    MaxDecelOwn: str
    MaxDecelTrail: str
    DecelRedDistOwn: str
    DecelRedDistTrail: str
    AccDecelOwn: str
    AccDecelTrail: str
    DiffusTm: str
    SafDistFactLnChg: str
    CoopDecel: str
    MinCollTmGain: str
    MinSpeedForLat: str
    LatDistStandDef: str
    LatDistDrivDef: str
    DesLatPos: str
    ObsrvAdjLn: str
    DiamQueu: str
    ConsNextTurn: str
    OvtLDef: str
    OvtRDef: str
    ObsrvdVehs: str