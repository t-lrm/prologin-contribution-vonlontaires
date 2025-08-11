import os
import openpyxl
import pandas as pd


CUR_DIR = os.path.dirname(os.path.abspath(__file__))
INPUT_PATH = CUR_DIR + "/inputs"


class File:
    def __init__(self, path: str):
        self.filename = os.path.basename(path)
        self.path = path

    def get_event_name(self) -> str:
        split_filename = self.filename.split("] - ")

        if len(split_filename) > 1:
            return split_filename[1].strip()[:-5]

        return self.filename[:-5]
    
    def get_event_date(self) -> str:
        split_filename = self.filename.split("] - ")

        if len(split_filename) > 1:
            return split_filename[0].strip()[1:]

        return "???"

    def get_event_type(self) -> str:
        if "Comptabilité GCC" in self.filename:
            return "GCC!"
        elif "Comptabilité Atelier" in self.filename:
            return "Atelier"
        elif "Comptabilité ER" in self.filename:
            return "Épreuve régionale"
        else:
            return "Autre"

    def get_dataframe(self) -> pd.DataFrame:
        df = pd.read_excel(
            self.path,
            sheet_name="Organisateurs",
            engine="openpyxl",
        )
        df = df.iloc[4:, :3]
        df = df.dropna(how="all")
        return df.reset_index(drop=True)


class Main:
    def __init__(self, input_path: str):
        self.files = [
            File(os.path.join(input_path, filename))
            for filename in os.listdir(input_path)
            if filename.endswith(".xlsx")
        ]
        self.input_path = input_path

    def generate_dataframe(self, filter: list = None) -> pd.DataFrame:
        self.output_df = pd.DataFrame(
            columns=["Event name", "Event type", "Date", "Name", "Hours", "Type"]
        )

        for file in self.files:
            df = file.get_dataframe()

            # Add file information to the output dataframe
            df["Event name"] = file.get_event_name()
            df["Event type"] = file.get_event_type()

            # Rename the first three columns to match "Name", "Hours", "Type"
            df = df.rename(
                columns={
                    df.columns[0]: "Name",
                    df.columns[1]: "Hours",
                    df.columns[2]: "Type",
                }
            )

            df = df[df["Type"].isin(filter)] if filter else df  # Filter rows
            self.output_df = pd.concat([self.output_df, df], ignore_index=True)

        return self.output_df


if __name__ == "__main__":
    main = Main(INPUT_PATH)
    df = main.generate_dataframe()

    # Save the output dataframe to a csv file
    output_path = os.path.join(CUR_DIR, "output.csv")
    df.to_csv(output_path, index=False)
