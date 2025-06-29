import os
import tkinter as tk
import pandas as pd
from tkinter import filedialog


def get_file():
    file_path = filedialog.askopenfilename(
        title="Select File",
        filetypes=[("All Files", "*.*"), ("All Files", "*.*")]
    )
    if file_path:
        full_path = os.path.join(os.getcwd(), file_path)
        if entry1:
            entry1.delete(0, "end")
        entry1.insert(0, full_path)
        if get_file_extension(file_path) == ".xlsm" or ".xlsx":
            file_label.config(text="Convert Excel File to Json")
        if get_file_extension(file_path) == ".json":
            file_label.config(text="Convert Json File to Excel")

def get_file_extension(file_path):
    _, extension = os.path.splitext(file_path)
    return extension

def convert_file():
    if get_file_extension(entry1.get()) == "xlsm" or ".xlsx":
        convert_excel_to_json()
    elif get_file_extension(entry1.get()) == ".json":
        convert_to_excel()
    else:
        file_label.config(text="Pls Select only a Json or Excel File")


def convert_to_excel():
    if entry1:
        df = pd.read_json(entry1.get())
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")],
                                                 title="Convert Json As", initialfile="Json_To_Excel")
        if file_path:
            with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                df.to_excel(writer, index=False)


def convert_excel_to_json():
    if entry1:
        df = pd.read_excel(entry1.get())
        json_data = df.to_json(orient="records", indent=4)

        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")],
                                                 title="Save JSON File", initialfile="Excel_To_JSON")
        if file_path:
            with open(file_path, "w") as json_file:
                json_file.write(json_data)


root = tk.Tk()
root.geometry("500x200")
root.title("Json to Excel Converter")
label1 = tk.Label(root,text="Select Json File")
label1.pack()
frame1 = tk.Frame(root)
frame1.pack()
entry1 = tk.Entry(frame1,width=60)
entry1.pack(side=tk.LEFT,padx=5, pady=5)
button1 = tk.Button(frame1, text="Select",width=10, command=get_file)
button1.pack(side=tk.LEFT, padx=5, pady=5)
file_label = tk.Label(root, text= "No Selected File")
file_label.pack(pady=5,padx=5)
convert_button = tk.Button(root,text="Convert", width=20, command=convert_file)
convert_button.pack(pady=20)

root.mainloop()