def get_variables24(info):
    dict_info = {
        1: info.LookAheadDistMin,
        2: info.LookAheadDistMax,
        3: info.LookBackDistMin,
        4: info.LookBackDistMax,
        5: info.W74ax,
        6: info.W74bxAdd,
        7: info.W74bxMult,
        8: info.MaxDecelOwn,
        9: info.MaxDecelTrail,
        10: info.DecelRedDistOwn,
        11: info.DecelRedDistTrail,
        12: info.AccDecelOwn,
        13: info.AccDecelTrail,
        14: info.DiffusTm,
        15: info.SafDistFactLnChg,
        16: info.CoopDecel,
        17: info.DesLatPos,
        18: str(info.ObsrvAdjLn),
        19: str(info.DiamQueu),
        20: str(info.ConsNextTurn),
        21: info.MinCollTmGain,
        22: info.MinSpeedForLat,
        23: str(info.OvtLDef),
        24: str(info.OvtRDef),
        25: info.LatDistStandDef,
        26: info.LatDistDrivDef,
        #27: info.Zipper,
        #28: info.ZipperMinSpeed
    }

    return dict_info

def get_variables10(info):
    dict_info = {
        1: info.LookAheadDistMin,
        2: info.LookAheadDistMax,
        3: info.LookBackDistMin,
        4: info.LookBackDistMax,
        5: info.W74ax,
        6: info.W74bxAdd,
        7: info.W74bxMult,
        8: info.MaxDecelOwn,
        9: info.MaxDecelTrail,
        10: info.DecelRedDistOwn,
        11: info.DecelRedDistTrail,
        12: info.AccDecelOwn,
        13: info.AccDecelTrail,
        14: info.DiffusTm,
        15: info.SafDistFactLnChg,
        16: info.CoopDecel,
        17: info.DesLatPos,
        18: str(info.ObsrvAdjLn),
        19: str(info.DiamQueu),
        20: str(info.ConsNextTurn),
        21: info.MinCollTmGain,
        22: info.MinSpeedForLat,
        23: str(info.OvtLDef),
        24: str(info.OvtRDef),
        25: info.LatDistStandDef,
        26: info.LatDistDrivDef,
        #27: info.Zipper,
        #28: info.ZipperMinSpeed
    }

    return dict_info

def parse_numbers(text):
    numbers = []
    ranges = text.split(',')
    for r in ranges:
        parts = r.strip().split('-')
        if len(parts) == 1:
            numbers.append(int(parts[0]))
        elif len(parts) == 2:
            start = int(parts[0])
            end = int(parts[1])
            numbers.extend(range(start, end+1))
    return list(set(numbers))