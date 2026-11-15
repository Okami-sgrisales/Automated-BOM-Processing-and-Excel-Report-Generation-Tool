# Author: Sebastián Grisales 
# email: sgrisales@gmail.com

import pandas as pd
import numpy as np
import sys
import os

#Constants
PARTS_NAME = "parts.xlsx" #filename of the excel doc of parts saved by AUTODESK INVENTOR
ASSEMBLIES_NAME = "assemblies.xlsx" #filename of the exxcel doc of assemmblies parts saved by AUTODESK INVENTOR
REQ_COLUMNS = ["Item", "QTY", "Filename", "Description", "Comments", "Subject", "Mass", "Part Number"]
FINAL_COL = ["Item", "QTY", "Filename", "Description", "Comments", "Material", "Mass"]
RED = '\033[31m' #This makes the text red in the cml, used for warnings
RESET = '\033[0m' #resets the color of the text in the cml
REMOVE_PARTS = ["103006865.ipt", "101218025.ipt", "BALIZA_D52.ipt", "11931238.iam", "11931238.iam",]


#Loading lists into pandas dataframes
def read_excel_files():
    ###Reads the excel files and returns the dataframes, if not excel found runs file_path() again
    while True:
        try:
            folder = input("Escribe el nombre de la carpeta con los archivos: ").lower()
            folder_list = os.listdir("./Listas/")
            for item in folder_list:
                if item.lower() == folder:
                    folder = item
            file_path = f"./Listas/{folder}" 
            parts_list = pd.read_excel(f"{file_path}/{PARTS_NAME}")
            assemblies_list = pd.read_excel(f"{file_path}/{ASSEMBLIES_NAME}")
            break
        except FileNotFoundError:
            print("\nNo se encontraron los archivos de Excel. Por favor, verifica el nombre de la carpeta e inténtalo de nuevo.\n")
    return parts_list, assemblies_list


#Checks the required Columns
def checks_columns(df):
    ###Checks that the loaded dataframe has the required columns
    columns = df.columns.tolist()
    if not all(item in columns for item in REQ_COLUMNS):
        print(columns)
        print(REQ_COLUMNS)
        sys.exit("\nFaltan columnas necesarias para el funcionamiento correcto del programa\n")


#Fixing *varies* for beams
def fix_beam_filenames(df):
    ###Due to how inventor works structural elements may display an incorrect Filename
    df.loc[df["Comments"] == "TUBERIA ESTRUCTURAL", "Filename"] = df["Part Number"]


#Checking for *varies* to give a warning
def varies_alert(df):
    ###This function alerts that some elements in the lists may be filled incorrectly at the design stage
    varies_column = []
    for col in REQ_COLUMNS:
        if (df[col] == "*Varies*").any():
            varies_column.append(col)
    if varies_column:
        print(f"\n{RED}Advertencia:{RESET} las columnas {varies_column} contienen celdas rellenadas con *Varies*, se sugiere revisar projecto en AUTODESK INVENTOR\n")


#Preparing columns
def sorting_columns(df):
    ###Renames the subject column
    df = df.rename(columns = {"Subject" : "Material"}) 
    return df[FINAL_COL]


#Sorting parts
def parts_sort(df):
    ###Sorts the dataframe and fixes the  index
    df = df[~df["Filename"].isin(REMOVE_PARTS)]
    df.sort_values(by= "Filename", inplace = True)
    df["Item"] = range(1, len(df) + 1)
    df.set_index("Item", inplace = True)
    return df


#Assemblies sorting into smaller dataframes
def assemblies_hierarchy(df):
    ###This function creates organized dataframes by item, it groups the assemblies by creating a new column named Hierarchy.
    df["Item"] = df["Item"].astype(str) #Ensures that item is a string
    df["Hierarchy"] = df["Item"].str[:-2] 
    df["Hierarchy"] = df["Hierarchy"].replace("", "0")
    df["Hierarchy"] = df["Hierarchy"].str.removesuffix(".") #in some cases the item may have more than 1 number as its suffix, when the string is sliced a dot will remain
    groups = df["Hierarchy"].unique().tolist()
    assemblies_list = [] #list used to to keep track of all the assemblies
    filename = df.loc[0]["Filename"]
    filename = filename[:-9] + "1-001.iam" #This line is to try to get the name of the main assembly
    main_assembly = {"Item" : "0", "QTY" : "N/A", "Filename": filename, "Description" : "N/A", "Comments": "N/A", "Material": "N/A", "Mass" : "N/A", "Hierarchy": "0"}
    new_row = pd.DataFrame(main_assembly, index = [0])
    df = pd.concat([new_row, df], ignore_index=True)
    for item in groups:
        assembly_item = df.query("Item == @item or Hierarchy == @item")
        assembly_name = df.loc[df["Item"] == item, "Filename"].iloc[0]
        assemblies_list.append({assembly_name : assembly_item[FINAL_COL].iloc[1:].copy()})
    return assemblies_list

#Drop unnecesary columns
def drop_col(df):
    df.drop(columns = "Part Number", inplace = True)

#execute
if __name__ == "__main__":
    df_parts, df_assemblies = read_excel_files()
    df_lists = [df_parts, df_assemblies]
    for ind, df in enumerate(df_lists):
        fix_beam_filenames(df)
        checks_columns(df)
        varies_alert(df)
        df_lists[ind] = sorting_columns(df)
    df_parts = parts_sort(df_lists[0])
    df_assemblies_list = assemblies_hierarchy(df_lists[1])
    for dic in df_assemblies_list:
        print(f"\n{dic.keys()}")
        print(f"{dic.values()}\n")
