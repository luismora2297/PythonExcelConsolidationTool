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
    WorkExcelFiles = [ xcelFiles for xcelFiles in os.listdir(WorkDirectory) if os.path.isfile(xcelFiles) and 'Observer.py' not in xcelFiles and 'MasterExcel.xlsx' not in xcelFiles ]  #Lista de los documentos en el directorio filtrando que los documentos no sean carpetas y sin filtrar el documento principal del programa

    #Revisar cuantos archvios de Excel tenemos en el directorio
    def FECheck(self, workingFiles):
        #Variables
        xQuantity=0
        xcelTotal = pd.DataFrame()

        #Revisar y hacer un loop en los archivos de Excel
        for excelFiles in workingFiles:
            if excelFiles.endswith('.xlsx') or excelFiles.endswith('.xls'): #Si termina con alguno de estos dos, ya que otros tipos o extensiones pueden ser complicados de manejar
                xQuantity=xQuantity+1
                print("\n You have the file "+ excelFiles +" currently processing")
                xcelCurrentFile = pd.ExcelFile(excelFiles) #Usando Panda para abrir el Archivo actual de Excel
                xcelSheets = xcelCurrentFile.sheet_names #Obtener los nombres de las hojas como llaves
                for xcelCurrentSheet in xcelSheets: #Hacer un ciclo de repeticion dentro del excel
                    getDataSheet = xcelCurrentFile.parse(sheet_name=xcelCurrentSheet) #Obtener los datos
                    xcelTotal = pd.concat([xcelTotal, getDataSheet]) #Guardar los datos en una variable

                #Parar cada 2 segundos cada proceso para no sobrecargarlo
                time.sleep(2)
                #Mover el Documento a la Carpeta de Procesado
                os.rename(self.WorkDirectory+"/"+excelFiles , self.WorkDirectory+"/Processed/"+excelFiles )
            else:
                #Mover el Documento a la Carpeta de no aplicables
                print("\n You have the file "+ excelFiles +" which is not applicable")
                #Parar cada 2 segundos cada proceso para no sobrecargarlo
                time.sleep(2)
                os.rename(self.WorkDirectory+"/"+excelFiles , self.WorkDirectory+"/Not Applicable/"+excelFiles )

        #Hoja Principal de Excel
        xcelTotal.to_excel('MasterExcel.xlsx') #Crear el documento principal que guardara dichos datos

        #Condicion dependiendo de la cantidad de Excels
        if xQuantity >= 1:
            print("\n You had "+ str(xQuantity) +" file(s) proccessed")
        else:
            print("\n You don't have any files available to be processed")

        #Mensaje Final
        print("\n Press Control + C if you want to exit this program, or just change the directory by loading other files")

    #Función inicial
    def __init__(self):
        #Print Current Location as Static
        print("Your Current Working Directory is:"+ self.WorkDirectory)
        print("\n The Current Files that your directory have are:")
        print(self.WorkExcelFiles)

        #Test The File
        self.FECheck(self.WorkExcelFiles)
    
    
    #Funciones para los eventos de creación, modificacion, y otros
    def on_any_event(self, directoryEvent): #Revisa cualquier evento no contemplado por las otras funciones en el directorio
        #print(directoryEvent)
        newWorkExcelFiles = [ xcelFiles for xcelFiles in os.listdir(self.WorkDirectory) if os.path.isfile(xcelFiles) and 'Observer.py' not in xcelFiles and 'MasterExcel.xlsx' not in xcelFiles ]  #Lista de los documentos en el directorio filtrando que los documentos recargada
        self.FECheck(newWorkExcelFiles)


#Main Executioner
if __name__ == '__main__':

    eventExecution = FExcelHandler()
    xcelObserver = Observer()
    xcelObserver.schedule(eventExecution, path='./', recursive=False) #Indicar donde hacer la observacion
    xcelObserver.start() #Comenzar el Observador
    #Try Catch para el evento de cerrar la ejecucion del programa
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        xcelObserver.stop()
    #Terminar el hilo de Trabajo
    xcelObserver.join()