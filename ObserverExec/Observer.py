#Requerir Librerías - Require Libraries
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os
import pandas as pd
from time import sleep #Para usar la función sleep

#First Class, un handler para ver los archivos que hay en el directorio
class FExcelHandler(FileSystemEventHandler):

    #Variables
    WorkDirectory = os.path.abspath('') #Manera dinamica de obtener el directorio actual
    WorkExcelFiles = os.listdir(WorkDirectory) #Lista de los documentos en el directorio

    #Funcion para los eventos de creación, modificacion, y otros

    #Función para tener los nombres
    def __init__(self):
        #Print Curren Location as Static
        print("Your Current Working Directory is:"+ self.WorkDirectory)
        print("The Current Files that your directory have are:")
        print(self.WorkExcelFiles)

        #Revisar cuantos archvios de Excel tenemos en el directorio
        def FECheck(workingFiles):
            #Variables
            xQuantity=0
            #Para cada archivo
            for excelFiles in workingFiles:
                if excelFiles.endswith('.xlsx') or excelFiles.endswith('.xls'): #Si termina con alguno de estos dos, ya que otros tipos o extensiones pueden ser complicados de manejar
                    xQuantity=xQuantity+1
            
            #Condicion dependiendo de la cantidad de Excels
            if xQuantity >= 1:
                print("You have "+ xQuantity +" file(s) available to process")
            else:
                print("You don't have any files available to process")

        #Test The File
        FECheck(self.WorkExcelFiles)


testVar = FExcelHandler()
        

