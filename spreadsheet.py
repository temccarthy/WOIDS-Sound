import math
import pandas as pd
import os
import glob
import PIL.Image
import shutil

template_name = "SEATTLE_TEMPLATE.xlsx"


# Basic spreadsheet info class
class LocationInfo:
    def __init__(self, insp_date):
        self.insp_date = insp_date


# Piece of equipment class
class Equipment:
    def __init__(self, id, station, room, component, notes, image_path):
        self.id = str(id)
        self.station = station
        self.room = room
        self.component = component
        self.notes = notes
        self.image_path = image_path

    # generates an equipment given a row in the sheet
    @staticmethod
    def generate_equip(folder, tup):
        # deal with empty cells in sheet
        tup = tuple("" if isinstance(i, float) and math.isnan(i) else i for i in tup)

        # calculate picture path
        image_path = glob.glob(folder + "/" + str(tup[1]) + ".*")[0]

        return Equipment(tup[1], tup[2], tup[3], tup[4], tup[5], image_path)


# holds spreadsheet dataframe
class Sheet:
    def __init__(self, path):
        self.path = path
        self.folder = self.path[:self.path.rfind("/", 0, -1)]

        self.fp = pd.ExcelFile(path)
        loc_sheet = self.fp.parse(0)
        self.location = LocationInfo(loc_sheet.columns[1])

    def check_pictures(self):
        missing_pics = []
        for row in self.fp.parse(1).itertuples():
            files = glob.glob(self.folder + "/" + str(row[1]) + ".*")
            if len(files) == 0:
                missing_pics.append(str(row[1]))

        return missing_pics

    def compress_pictures(self):
        if os.path.isdir(os.path.join(self.folder, "temp")):
            self.delete_compressed_pictures()

        os.mkdir(os.path.join(self.folder, "temp"))
        for file in os.listdir(self.folder):
            try:
                with PIL.Image.open(os.path.join(self.folder, file)) as img:
                    img.save(os.path.join(self.folder, "temp", file), optimize=True, quality=30)
            except Exception:
                pass

    def delete_compressed_pictures(self):
        shutil.rmtree(os.path.join(self.folder, "temp"))

    @staticmethod
    def check_template_exists(path):
        matches = glob.glob(os.path.join(path, template_name))
        return len(matches) != 0
