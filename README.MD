#Español
Este repositorio toma dos archivos de excel en una subcarpeta de input/, los cuales son excels que contienen la informacion cruda que entrega Autodedsk Inventor sobre los ensamblajes y partes, Les realiza un prepocesado y crea un formato en el que organiza en tablas los diferentes ensamblajes en una sola hoja de excel, mientras que la segunda digita todas las partes necesarias en la carpeta de output/.

Nota: cuando la carpeta input incluye una subcarpeta llamada sample para propositos demostrativos de la funcionalidad del algoritmo, cuando el algoritmo pregunte por la ruta que contiene los archivos se debe escribir sample no importa sin en mayusculas o no, y los archivos deben ser nombrados assemblies.xlsx y parts.xlsx respectivamente

---------------------------------------------------------------------

#English
This repository requires two excel files in a subfolder inside input/, these two excel files are obtained from Autodesk Inventor with the information related to the assemblies and parts, then it preprocesses the data contained in the excels, sorts in tables each assembly with its required parts and prints the tables into a new excel file inside the folder output/.

Note: the folder named input includes a sample folder to demostrate the functionality of this algorithm, the required files must be named assemblies.xlsxx and parts.xlsx. When the program ask for the folder sample can be typed in caps or lowercase, it is case insensitive