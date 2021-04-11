# -------------------------------------------------------------------------------
# -- Author       Alexander J. Samo
# -- Created      03/14/2021
# -- Purpose      Automate google drive operations
# -- Copyright © 2021, Modesto City Schools, All Rights Reserved
# -------------------------------------------------------------------------------
# -- Modification History
# --
# -- 03/13/2021  Alexander J. Samo
# --      Initial write.
# -------------------------------------------------------------------------------

# ------------------------------- IMPORTS -------------------------------
import os
import sys

import pandas as pd
import gspread
import gspread_dataframe as gd

# ------------------------------ FILE PATHS ------------------------------
gregori_no_teachers_dir_path = r"\\admb-its\Shared\IS\Projects\HS Ballot Recommendations\20-21\Joseph Gregori High School\No_Teachers"
gregori_teachers_dir_path = r"\\admb-its\Shared\IS\Projects\HS Ballot Recommendations\20-21\Joseph Gregori High School\Teachers"

# -------------------------- GOOGLE FOLDER PATHS --------------------------
gregori_no_teachers_google_folder = "1QfWZx3GLmXlFAz0yt7FmugQ1-RpnOKqq"
gregori_teachers_google_folder = "1OzK0mctBmXituFJkz13rvR9UQKv71Vad"


# -------------------------- CONNECT GOOGLE DRIVE --------------------------
try:
    gc = gspread.service_account(filename='client_secret.json')
    print(gc.__dict__)
    print()
except Exception as ex:
    print(ex)


def add_sheet(filepath, file_name, google_folder_ID):
    if len(filepath) > 0:
        for filename in filepath:
            print(file_name)
            # path to file with extension
            process_files = filepath

            # setting up data frame for each file
            excel = pd.ExcelFile(process_files)
            columns = None

            # create each file in the google sheet folder
            sh = gc.create(file_name, folder_id=google_folder_ID)

            # need to share the file so others can see it
            # sh.share('samo.a@monet.k12.ca.us', perm_type='user', role='writer')

            for index, name in enumerate(excel.sheet_names):
                print(f'Reading sheet #{index}: {name}')
                sheet = excel.parse(name)
                data_frame = pd.DataFrame(sheet, index=None)
                if index == 0:
                    # Save column names from the first sheet to match for append
                    columns = sheet.columns
                sheet.columns = columns
                # Assume index of existing data frame when appended
                worksheet = sh.add_worksheet(title=name, rows=0, cols=0)
                gd.set_with_dataframe(worksheet=worksheet, dataframe=data_frame, include_index=False,
                                      include_column_header=True)
            sh.del_worksheet(sh.sheet1)

if __name__ == '__main__':
    path = r'C:\Alex\Projects\Joseph Gregori High School\Teachers'

    files = []
    # r=root, d=directories, f=files
    for r, d, f in os.walk(path):
        for file in f:
            if '.xlsx' in file:
                files.append(os.path.join(r, file))
                path_of_file = os.path.join(r, file)
    for f in files:
        pass
        add_sheet(path_of_file, file, gregori_teachers_google_folder)


