import pandas as pd
from openpyxl import load_workbook


class CarDeduplicator:
    def __init__(self, ignore_columns: list):
        self.ignore_columns = ignore_columns
        self.df = None

    def remove_duplicates(self, df):
        print("Starting to remove duplicates...")
        try:
            self.df = df
            consider_columns = [col for col in self.df.columns if col not in self.ignore_columns]
            self.df = self.df.drop_duplicates(subset=consider_columns)
            self.df = self.df[consider_columns]  # Keeps only the columns considered for duplication
            self.df = self.df.dropna(how='all')  # Drops rows with all NaN values
            print("Done")
        except KeyError as e:
            print(f"Error: One or more columns are not found in the DataFrame: {e}")
            raise
        except Exception as e:
            print(f"An error occurred while removing duplicates: {e}")
            raise
        return self.df


class ToyotaDeduplicator(CarDeduplicator):
    def __init__(self, file, ignore_columns: list, new_sheet_name=f'UPDATED'):
        super().__init__(ignore_columns)
        self.file = file
        self.new_sheet_name = new_sheet_name

    def read_file(self):
        print("Starting to read the file...")
        try:
            if not self.file.lower().endswith(('.xls', '.xlsx', '.xlsm', '.xlsb')):
                raise ValueError("Please provide a valid Excel file.")

            df = pd.read_excel(self.file)
            print("Done")
            return df
        except FileNotFoundError:
            print(f"Error: The file {self.file} does not exist.")
            raise
        except Exception as e:
            print(f"An error occurred while reading the file: {e}")
            raise

    def save_file(self, df):
        print("Starting to save the file...")
        try:
            with pd.ExcelWriter(self.file, engine='openpyxl', mode='a') as writer:
                book = writer.book
                if self.new_sheet_name in book.sheetnames:
                    del book[self.new_sheet_name]  # Overwrites the sheet if it already exists
                df.to_excel(writer, sheet_name=self.new_sheet_name, index=False)
            print("\nSuccessful deduplication!")
        except PermissionError:
            print(f"Error: Permission denied when trying to write to {self.file}.")
            raise
        except Exception as e:
            print(f"An error occurred while saving the file: {e}")
            raise

    def deduplicate(self):
        try:
            df = self.read_file()
            deduplicated_df = self.remove_duplicates(df)
            self.save_file(deduplicated_df)
        except Exception as e:
            print(f"\nProcess failed!")
