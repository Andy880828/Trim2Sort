import fnmatch
import os

import pandas as pd


def find_latest_ref_file():
    """尋找最新的 ref_ 開頭的 xlsx 檔案"""
    root_dir = os.path.dirname(os.path.abspath(__file__))
    ref_files = []
    for file in os.listdir(root_dir):
        if file.startswith("ref_") and file.endswith(".xlsx"):
            file_path = os.path.join(root_dir, file)
            ref_files.append((file_path, os.path.getmtime(file_path)))

    if ref_files:
        # 根據修改時間排序, 返回最新的
        ref_files.sort(key=lambda x: x[1], reverse=True)
        return ref_files[0][0]
    return ""


root_dir = os.path.dirname(os.path.abspath(__file__))
input_folder = os.path.join(root_dir, "outputs", "I_sorted_blasts")
output_folder = os.path.join(root_dir, "outputs", "TEST")
ref = find_latest_ref_file()


class zh_adder:
    def __init__(self, input_folder, output_folder):
        self.input_folder = input_folder
        self.output_folder = output_folder
        self.Sample_Name = fnmatch.filter(os.listdir(input_folder), "*.xlsx")
        self.Sample_Size = len(self.Sample_Name)
        self.Output_Name = os.path.join(
            output_folder, str(self.Sample_Name) + "_zh_added.xlsx")

    def zh_adding(self):
        df_ref = pd.read_excel(ref)
        os.chdir(self.input_folder)
        for Sample in self.Sample_Name:
            input_file = os.path.join(self.input_folder, Sample)
            df = pd.read_excel(input_file, engine="openpyxl")
            df_ref_merged = df.merge(df_ref, how="left", on="Scientific_name")
            output_name = os.path.splitext(Sample)[0] + "_zh_added.xlsx"
            output_path = os.path.join(self.output_folder, output_name)
            df_ref_merged.to_excel(output_path, engine="openpyxl", index=False)


zh = zh_adder(input_folder, output_folder)
zh.zh_adding()
