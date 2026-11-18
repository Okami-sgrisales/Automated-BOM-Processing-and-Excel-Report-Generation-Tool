# Author: Sebastián Grisales 
# email: sgrisales@gmail.com

import dprep
import twriter
import os


START_ROW = 7
BLANK_ROW = 2


if __name__ == "__main__":
    df_parts, df_assemblies = dprep.read_excel_files()
    df_lists = [df_parts, df_assemblies]
    for ind, df in enumerate(df_lists):
        dprep.fix_beam_filenames(df)
        dprep.checks_columns(df)
        dprep.varies_alert(df)
        df_lists[ind] = dprep.sorting_columns(df)
    df_parts = dprep.parts_sort(df_lists[0])
    df_assemblies_list = dprep.assemblies_hierarchy(df_lists[1])

    
    out_folder = "output"
    os.makedirs(out_folder, exist_ok = True)
    project, client, city = twriter.get_project()
    wb, ws1, ws2, writer = twriter.create_workbook()
    try:
        wsl = [ws1, ws2]
        for ws in wsl:
            format = twriter.format(wb)
            twriter.apply_border(wb, ws, 1, 4, 1, 3)
            twriter.column_width(ws, 44, 35, 20, 20, 12)
            twriter.insert_img(ws, "D2", 0.4, 0.25)
            twriter.type_basic_info(ws, format, project, client, city)
        twriter.add_table(ws2, START_ROW, 0, df_parts, format, True)
        current_row = START_ROW
        for assembly in df_assemblies_list:
            key = list(assembly.keys())
            current_row = twriter.add_main_assembly_row(ws1, assembly[key[0]], current_row, 0, key[0], format)
            current_row = twriter.add_table(ws1, current_row, 0, assembly[key[0]], format, False)
            current_row += BLANK_ROW
    finally:
        wb.close()