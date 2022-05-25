#Requerir Librerías - Require Libraries
import time
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
import os
import pandas as pd
from time import sleep #Para usar la función sleep

#Variables
WorkDirectory = os.path.abspath('') #Manera dinamica de obtener el directorio actual
WorkExcelFiles = os.listdir(WorkDirectory) #Lista de los documentos en el directorio


"""
print(WorkDirectory)
print(WorkExcelFiles)
Solo para Tests
"""


