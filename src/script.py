
from dataclasses import dataclass
import os

@dataclass
class Filename :
    filename: str
    path: str
    number: int
    serialNumber: int



class Document :

    def __init__(self) -> None:
        self.Strang1Path = "../Strang1/xls/"
        self.Strang2Path = "../Strang2/xls/"
        self.Strang3Path = "../Strang3/xls/"
        
        
    def create_textfile(self):
        self.missing_ref_textfile = "../output/missing.txt"
        myFile = open(self.missing_ref_textfile, "w+")
        myFile.close()

    def write_missing_ref(self, filename):
        with open(self.missing_ref_textfile, "a") as myFile:
            myFile.write(filename + "\n")
            
            



    

    Buffer = []

    def read_dir(self):
        for file in os.listdir(self.Strang1Path):
            if file.endswith(".xls"):




