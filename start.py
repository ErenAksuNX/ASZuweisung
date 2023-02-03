from pandas import *
from os import *
from datetime import *


def nullstring(dateTime):
    if dateTime is not None and dateTime == 0:
        return "00"
    elif dateTime is not None and dateTime < 10:
        return "0" + str(dateTime)
    elif dateTime is not None:
        return str(dateTime)


def istFrist(grund):
    if "GSM-R kleine Inspek" in str(grund):
        return True
    elif "IS 39S" in str(grund):
        return True
    elif "Fahrzeuginspektion 1" in str(grund):
        return True
    elif "Fahrzeuginspektion 2" in str(grund):
        return True
    elif "Fahrzeuginspektion 3" in str(grund):
        return True
    elif "Fahrzeuginspektion 4" in str(grund):
        return True
    elif "Fahrzeuginspektion 5" in str(grund):
        return True
    elif "Br 1.1" in str(grund):
        return True
    elif "Br 1.2" in str(grund):
        return True
    elif "510" in str(grund):
        return True
    elif "520" in str(grund):
        return True
    elif "530" in str(grund):
        return True
    elif "540" in str(grund):
        return True
    elif "550" in str(grund):
        return True
    elif "Jahrespaket (3jährige Arbeit)" in str(grund):
        return True
    elif "IS 200" in str(grund):
        return True
    elif "IS 308" in str(grund):
        return True
    elif "IS 309" in str(grund):
        return True
    else:
        return False


def zgrung(grund):
    if "GSM-R kleine Inspek" in str(grund):
        return "ZF1"
    elif "IS 39S" in str(grund):
        return "PZB"
    elif "Fahrzeuginspektion 1" in str(grund):
        return "L1"
    elif "Fahrzeuginspektion 2" in str(grund):
        return "L2"
    elif "Fahrzeuginspektion 3" in str(grund):
        return "L3"
    elif "Fahrzeuginspektion 4" in str(grund):
        return "L4"
    elif "Fahrzeuginspektion 5" in str(grund):
        return "L5"
    elif "Br 1.1" in str(grund):
        return "Br1.1"
    elif "Br 1.2" in str(grund):
        return "Br1.2"
    elif "510" in str(grund):
        return "F1"
    elif "520" in str(grund):
        return "F2"
    elif "530" in str(grund):
        return "F3"
    elif "540" in str(grund):
        return "F4"
    elif "550" in str(grund):
        return "F5"
    elif "Jahrespaket (3jährige Arbeit)" in str(grund):
        return "Z1"
    elif "IS 200" in str(grund):
        return "NS"
    elif "IS 308" in str(grund):
        return "SF"
    elif "309" in str(grund):
        return "WF"


class Start:
    def __init__(self):
        self.df = None
        self.path = None
        self.destination = None

    def start(self, path, destination):
        self.path = path
        self.destination = destination

        self.df = read_excel(self.path, sheet_name="Anhang 1a_Fertigung")

        for i, row in self.df.iterrows():
            if istFrist(self.df.at[i, "Bezeichnung"]):
                self.df.loc[i, "Arbeitsscheinnr. "] = self.getAsNr(i, self.df.at[i, "Bezeichnung"])

        self.df.to_excel(self.destination + "/auswertung.xlsx")

        print(self.df)

    def getAsNr(self, i, grund):
        path = "C:/Users/eaksu/OneDrive - National Express Rail GmbH/Desktop/Arbeitsscheine Sortiert"

        path = path + "/" + zgrung(grund)
        date = datetime.strptime(str(self.df.loc[i, "Bestelldat"]), "%Y-%m-%d %H:%M:%S")
        frznr = self.df.at[i, "Fahrzeug (ext)"].split()

        frznr = frznr[0]

        dateien = list()

        for file in listdir(path):
            for i in range(-3, 3):
                file_path = nullstring(date.day + i) + "." + nullstring(date.month) + "." + str(date.year) + "_" + str(
                    frznr)
                if file_path in file:
                    dateien.append(file)

        asNr = ""

        for file in dateien:
            nr = file.split("_")
            asNr = asNr + nr[2]
            var = asNr.split(".")
            asNr = var[0] + "               "
        print(asNr)

        return asNr
