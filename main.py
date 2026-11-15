# Author: Sebastián Grisales 
# email: sgrisales@gmail.com

import dprep


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
    for dic in df_assemblies_list:
        print(f"\n{dic.keys()}")
        print(f"{dic.values()}\n")