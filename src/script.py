
from dataclasses import dataclass
import glob
import os
import pandas as pd

EXIT_SUCCESS = 0
EXIT_FAILED = 1

"""

1) lire le fichier 1
2) chercher le fichier 2 et 3
3) si fichiers 2 et 3 existent : remplir le buffer avec fichier 1
4) ecrire fichiers 2 et 3
5) sinon ecire le nom dans missing file txt


"""



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
        self.Buffer = []
        
        
    def create_textfile(self) -> int:
        try:
            self.missing_ref_textfile = "../output/missing.txt"
            myFile = open(self.missing_ref_textfile, "w+")
            myFile.close()

            return EXIT_SUCCESS
        except:
            print("Error : creating text file\n")
            return EXIT_FAILED


    def write_missing_ref(self, filename) -> int:
        try:
            with open(self.missing_ref_textfile, "a") as myFile:
                myFile.write(filename + "\n")

            return EXIT_SUCCESS
        except:
            print("Error : writing in missing ref txt file\n")
            return EXIT_FAILED
            

    def parse_name(self, name) -> int:
        try:
            lst = name.split("_")
            Filename.number = lst[1]
            Filename.filename = name
            Filename.serialNumber = (lst[2])[2:-4:]
        except:
            print("Error parsing name")
            return EXIT_FAILED
        return EXIT_SUCCESS


    def display_info(self) -> None:
        print("\n________Infos________\n")
        print("|__Filename: ", Filename.filename)
        print("|__Number: ", Filename.number)
        print("|__SerialNumber: ", Filename.serialNumber)
        print("______________________\n")


    def search_file(self, chemin) -> list:
        """
        Cherche un fichier .xls avec les patterns du fichier1 dans le <chemin> précisé
        Retourne le chemin local du fichier
        """
        pattern = Filename.serialNumber
        fichiers = [f for f in glob.glob(chemin + '/*') if f.endswith('.xls') and pattern in os.path.basename(f)]
        
        if(len(fichiers) > 1):
            print(f"Multiple files found in dir : {chemin}")
        if(len(fichiers) == 0):
            return []
        return fichiers[0]
    
    
    def fill_buffer(self, chemin) -> int:
        """
        Lit le fichier 1 et remplit le buffer avec les cellules I5 à L13
        """
        # Charger le fichier XLS
        df = pd.read_excel(chemin, header=None)

        # Sélectionner les données entre les cases I5 à L13
        self.Buffer = df.iloc[4:13, 8:12]

        return EXIT_SUCCESS


    def write_buffer_to_file(self, chemin) -> int:
        # Charger le fichier XLS
        df = pd.read_excel(chemin, header=None)

        #Ecrire le Buffer sur les cases I5 à L13
        df.iloc[4:13, 8:12] = self.Buffer
        df.to_excel(chemin, index=False)

        return EXIT_SUCCESS



    def read_dir(self):
        for file in os.listdir(self.Strang1Path):
            if file.endswith(".xls"):
                pass



if(__name__ == "__main__"):

    # variable de vérification
    retour = 0

    document = Document()
    
    nameXLS = "filename_01_11BEOX12.xls"

    #récupérer les infos fichier 1
    if(document.parse_name(nameXLS)):
        print(f"Doc : {nameXLS}\n")
        document.display_info()

    #chercher fichier 2
    file2Path = document.search_file(document.Strang2Path)

    #chercher fichier 3
    file3Path = document.search_file(document.Strang3Path)

    if(file2Path != [] and file3Path != []):
        #remplir buffer avec fichier 1
        document.fill_buffer(document.Strang1Path + "/" + Filename.filename)
    else:
        print("\nErreur : fichiers 2 ou 3 non trouvé\n")
        print("|__fichier 2 : ", file2Path)
        print("|__fichier 3 : ", file3Path)
        exit()
        

    #remplir les fichiers 2 et 3 avec le buffer
    document.write_buffer_to_file(file2Path)

    
    

