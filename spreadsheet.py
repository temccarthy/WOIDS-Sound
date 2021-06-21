import pandas as pd
import os
import glob

class LocationInfo:
    def __init__(self, rail, location, insp_date):
        self.rail = rail
        self.location = location
        self.insp_date = insp_date


class Equipment:
    def __init__(self, discipline, num, room, equipment_id, cs, title, descr, sol_title, sol_text, image_path):
        self.discipline = discipline
        self.num = num
        self.id = discipline + str(num)
        self.room = room
        self.equipment_id = equipment_id
        self.cs = cs
        self.title = title
        self.descr = descr
        self.sol_title = sol_title
        self.sol_text = sol_text
        self.image_path = image_path  # auto calculate?


class Sheet:
    def __init__(self, path):
        self.path = path
        self.folder = self.path[:self.path.rfind("/", 0, -1)]

        fp = pd.read_excel(path)
        for tup in fp.itertuples():
            if type(tup[1]) == float:  # skip empty cells
                continue
            elif tup[1].startswith("Discipline"):  # find top of table
                skip = tup[0]+1
                break
        self.df = pd.read_excel(path, skiprows=skip)

    def check_pictures(self):
        missing_pics = []

        for row in self.df.itertuples():
            files = glob.glob(self.folder + "/" + row[3] + ".*")
            if len(files) == 0:
                missing_pics.append(row[3])

        return missing_pics

