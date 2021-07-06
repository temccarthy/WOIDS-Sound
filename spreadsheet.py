import datetime
import math
import pandas as pd
import os
import glob


# Basic spreadsheet info class
class LocationInfo:
    def __init__(self, rail, location, insp_date):
        self.rail = rail
        self.location = location
        self.insp_date = insp_date.strftime("%m/%d/%Y") if isinstance(insp_date, datetime.datetime) else ""


# Piece of equipment class
class Equipment:
    def __init__(self, discipline, num, room, equipment_id, cs, title, descr, sol_title, sol_text, image_path):
        self.discipline = discipline
        self.num = num
        self.id = discipline + str(num)
        self.room = room
        self.equipment_id = equipment_id
        try:
            self.cs = min(max(int(cs), 0), 4)
        except Exception:
            self.cs = 0
        self.title = title
        self.descr = descr
        self.sol_title = sol_title
        self.sol_text = sol_text
        self.image_path = image_path

    # generates an equipment given a row in the sheet
    @staticmethod
    def generate_equip(folder, tup):
        # deal with empty cells in sheet
        tup = tuple("" if isinstance(i, float) and math.isnan(i) else i for i in tup)

        # calculate picture path
        picture_path = glob.glob(folder + "/" + tup[3][:1] + " (" + tup[3][1:] + ")" + ".*")[0]

        return Equipment(tup[1], tup[2], tup[4], tup[5], tup[6], tup[7], tup[8], tup[9], tup[10], picture_path)


# holds spreadsheet dataframe
class Sheet:
    def __init__(self, path):
        self.path = path
        self.folder = self.path[:self.path.rfind("/", 0, -1)]

        self.fp = pd.ExcelFile(path)
        loc_sheet = self.fp.parse(0)
        self.location = LocationInfo(loc_sheet.columns[1], loc_sheet.iloc[0, 1], loc_sheet.iloc[1, 1])

    def check_pictures(self):
        missing_pics = []
        for i in range(5):
            for row in self.fp.parse(i+1).itertuples():
                files = glob.glob(self.folder + "/" + row[3][:1] + " (" + row[3][1:] + ")" + ".*")
                if len(files) == 0:
                    missing_pics.append(row[3])

        return missing_pics

    @staticmethod
    def check_template_exists(path):
        template_name = "MBTA_TEMPLATE.xlsx"
        matches = glob.glob(os.path.join(path, template_name))
        return len(matches) != 0
