# Author: Sebastián Grisales 
# email: sgrisales@gmail.com

import dprep
import twriter
import os


# Constants
START_ROW = 7
BLANK_ROW = 2


# Logic
if __name__ == "__main__":
    # Data preprocessing
    df_parts, df_assemblies = dprep.read_excel_files() # reads files
    df_lists = [df_parts, df_assemblies]
    for ind, df in enumerate(df_lists):
        dprep.fix_beam_filenames(df) # fixes filenames of structured elements that are shown as *varies*
        dprep.checks_columns(df) # verify the required columns
        dprep.varies_alert(df) # shows an alert if *varies* exist in the dataframes
        df_lists[ind] = dprep.sorting_columns(df) # drops unnecesary columns
    df_parts = dprep.parts_sort(df_lists[0]) # sorts parts by name
    df_bolts = dprep.bolts(df_parts) # creates a bolt dataframe
    df_assemblies_list = dprep.assemblies_hierarchy(df_lists[1]) # creates a dataframe for each assembly


    #xlsx file processing
    out_folder = "output" #folder where the file will be created
    os.makedirs(out_folder, exist_ok = True) # verify the existence of the folder, it also creates the folder if it didn't exist
    project, client, city = twriter.get_project()
    wb, ws1, ws2, ws3, writer = twriter.create_workbook() # creates the workbook, sheets, and engine that connects pandas with xlsx
    try:
        wsl = [ws1, ws2, ws3]
        for ws in wsl:
            format = twriter.format(wb) # gets the style formats
            twriter.apply_border(wb, ws, 1, 4, 1, 3) # generates a tittle box
            twriter.column_width(ws, 44, 35, 20, 20, 12) # Adjust column width in the next order width_c, width_d, width_e, width_f, width_g
            twriter.insert_img(ws, "D2", 0.4, 0.25) # inserts image, the number arguments are the scale of x and y
            twriter.type_basic_info(ws, format, project, client, city)
        twriter.add_table(ws2, START_ROW, 0, df_parts, format, False) # adds the part list
        twriter.add_table(ws3, START_ROW, 0, df_bolts, format, False) # adds the bolt list
        current_row = START_ROW
        for assembly in df_assemblies_list:
            key = list(assembly.keys()) # gets the key from the dictionary
            current_row = twriter.add_main_assembly_row(ws1, assembly[key[0]], current_row, 0, key[0], format) # adds the header and table format
            current_row = twriter.add_table(ws1, current_row, 0, assembly[key[0]], format, False) # adds the assembly list
            current_row += BLANK_ROW
    finally:
        wb.close()