import pandas as pd
import gspread
import gspread_dataframe as gd


class Googler:

    def init(self, client_secret):
        try:
            self.gc = gspread.service_account(filename=client_secret)
            print(self.gc.__dict__)
        except Exception as ex:
            print(ex)

    def add_sheet(self, filepath, file_name, google_folder_ID):
        if len(filepath) > 0:
            for filename in filepath:
                print(file_name)
                # path to file with extension
                process_files = filepath

                # setting up data frame for each file
                excel = pd.ExcelFile(process_files)
                columns = None

                # create each file in the google sheet folder
                sh = self.gc.create(file_name, folder_id=google_folder_ID)

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
