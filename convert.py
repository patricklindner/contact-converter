import pandas as pd
import sys
from pandas import DataFrame
from tkinter.filedialog import askopenfilename, asksaveasfile
import sys
import os

mapping_table = {
    "Voornaam": "Given Name",
    "Achternaam": "Family Name",
    "Geb.datum": "Birthday",
    "Contact.email": "E-mail 1 - Value",
    "Mobiel privé": "Phone 1 - Value"
}


def fill_phone_numbers(df: DataFrame):
    df["Mobiel privé"] = df["Mobiel privé"].fillna(df["Tel.privé1"])
    df["Mobiel privé"] = df["Mobiel privé"].fillna(df["Tel.privé2"])


def transform_tel_nr(nr):
    if str(nr).startswith("00") or str(nr).startswith("+") or str(nr) == "nan":
        return nr
    else:
        return f"+31 {nr}"


if __name__ == "__main__":

    input_file = askopenfilename(filetypes=[("Excel Sheets", ".xlsx")])
    source = pd.read_excel(input_file, skiprows=[0, 1])

    fill_phone_numbers(source)

    source["Mobiel privé"] = source["Mobiel privé"].apply(transform_tel_nr)
    source["Tel.privé1"] = source["Tel.privé1"].apply(transform_tel_nr)

    bundle_dir = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))
    target = pd.read_csv(os.path.abspath(os.path.join(bundle_dir, "template.csv")))
    target = target.drop(target.index)

    for source_field, target_field in mapping_table.items():
        print(f"Mapping {source_field} to {target_field}")
        target[target_field] = source[source_field]

    output_file = asksaveasfile(defaultextension=".csv")
    target.to_csv(output_file, index=False)
