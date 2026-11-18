import xlsxwriter
import pandas as pd
import numpy as np


def get_project():
    ###Ask user for basic information
    project = input("Nombre de proyecto: ")
    client = input("Nombre de la empresa cliente: ")
    city = input("Ciudad en la que reside la empresa: ")
    return project, client, city


def create_workbook():
    ###Ask for a name for the file to create
    filename = input("Introduce el nombre que tendra el archivo de excel: ")
    #filename = "./output/" + filename + ".xlsx"
    #workbook = xlsxwriter.Workbook(filename)
    #worksheet1 = workbook.add_worksheet("Lista de Ensamblajes")
    #worksheet2 = workbook.add_worksheet("Lista de Partes(Total)")
    writer = pd.ExcelWriter(f"./output/{filename}.xlsx", engine='xlsxwriter')
    workbook  = writer.book
    workbook.add_worksheet("Lista de Ensamblajes")
    workbook.add_worksheet("Lista de Partes(Total)")
    worksheet1 = writer.sheets["Lista de Ensamblajes"]
    worksheet2 = writer.sheets["Lista de Partes(Total)"]
    return workbook, worksheet1, worksheet2, writer


def format(workbook):
    ###Generates some formatting styles
    formats = {
        "default" : workbook.add_format({
            "font_name" : "Times New Roman",
            "font_size" : 11
        }),


        "header" : workbook.add_format({
            "font_name" : "Times New Roman",
            "font_size" : 11,
            "bold" : True,
            "align" : "center"
        }),

        "top left border" : workbook.add_format({
            "font_name" : "Times New Roman",
            "font_size" : 11,
            "bold" : True,
            "top" : 2,
            "left" : 2,
        }),

        "left border" : workbook.add_format({
            "font_name" : "Times New Roman",
            "font_size" : 11,
            "bold" : True,
            "left" : 2,
            "left_color": "#000000"
        }),

        "bottom left border" : workbook.add_format({
            "font_name" : "Times New Roman",
            "font_size" : 11,
            "bold" : True,
            "bottom" : 2,
            "left" : 2
        }),

        "top border" : workbook.add_format({
            "font_name" : "Times New Roman",
            "font_size" : 11,
            "top" : 2
        }),

        "top right border" : workbook.add_format({
            "top" : 2,
            "right" : 2
        }),

        "bottom right border" : workbook.add_format({
            "bottom" : 2,
            "right" : 2
        }),

        "header_table" : workbook.add_format({
            "font_name" : "Times New Roman",
            "font_size" : 11,
            "border" : 1,
            "bold" : True,
            "align" : "center",
            "valign" : "vcenter",
        }),

        "cell" : workbook.add_format({
            "font_name" : "Times New Roman",
            "font_size" : 11,
            "border" : 1,
            "align" : "left",
            "valign" : "vcenter"
        })
    }
    return formats


def apply_border(workbook, worksheet, start_row, end_row, start_col, end_col):
    ###Creates a small box to write basic information about the project and a space for the logo
    left = workbook.add_format({"left" : 2})
    right = workbook.add_format({"right" : 2})
    top = workbook.add_format({"top" : 2})
    bottom = workbook.add_format({"bottom" : 2})

    for row in range(start_row, end_row):
        worksheet.write_blank(row, start_col, None, left)
        worksheet.write_blank(row, end_col, None, right)

    for col in range(start_col, end_col):
        worksheet.write_blank(start_row, col, None, top)
        worksheet.write_blank(end_row, col, None, bottom)


def column_width(worksheet, width_c, width_d, width_e, width_f, width_g):
    worksheet.set_column("C:C", width_c)
    worksheet.set_column("D:D", width_d)
    worksheet.set_column("E:E", width_e)
    worksheet.set_column("F:F", width_f)
    worksheet.set_column("G:G", width_g)


def insert_img(worksheet, cell, x_scale = 1, y_scale = 1, x_offset = 1, y_offset = 5):
    image_path = "./img/Sample company logo.png"
    worksheet.insert_image(cell, image_path, 
                           {"x_scale" : x_scale,
                            "y_scale" : y_scale,
                            "x_offset" : x_offset,
                            "y_offset" : y_offset})


def type_basic_info(worksheet, formats, project, client, city):
    worksheet.write("B2", "Proyeto:", formats["top left border"])
    worksheet.write("B3", "Cliente:", formats["left border"])
    worksheet.write("B4", "Ciudad:", formats["left border"])
    worksheet.write("C2", project, formats["top border"])
    worksheet.write("C3", client, formats["default"])
    worksheet.write("C4", city, formats["default"])
    worksheet.write("B5", None, formats["bottom left border"])
    worksheet.write("D2", None, formats["top right border"])
    worksheet.write("D5", None, formats["bottom right border"])

def add_table(ws, start_row, start_col, df, formats, index):
    rows, cols = df.shape
    ws.add_table(start_row, start_col, start_row + rows, start_col + cols - 1,
                 {
                     "columns" : [
                         {"header" : col, "header_format" : formats["header_table"]} for col in df.columns
                     ],
                     "style": "None",
                     "autofilter" : False
                 })
    df = df.replace([np.nan, np.inf, -np.inf], None)
    for row, values in enumerate(df.itertuples(index = index, name = None)):
        ws.write_row(start_row + row + 1, start_col, values, formats["cell"])