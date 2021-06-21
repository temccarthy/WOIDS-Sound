import pandas as pd


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
        fp = pd.read_excel(path)
        print("opened")

        for tup in fp.itertuples():
            if type(tup[1]) == float:
                continue
            elif tup[1].startswith("Discipline"):
                skip = tup[0]+1
                break
        self.df = pd.read_excel(path, skiprows=skip)
        print(self.df.head(5))
