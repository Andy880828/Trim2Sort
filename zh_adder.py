import pandas as pd
import fnmatch
import os

input_folder = "D:\\Trim2Sort_ver3.0.0\\outputs\\I_sorted_blasts"
output_folder = "D:\\Trim2Sort_ver3.0.0\\outputs\\TEST\\"
ref = "D:\\Trim2Sort_ver3.0.0\\ref_ver3_0408-3.xlsx"

class zh_adder:
    def __init__(self, input_folder, output_folder):
        self.input_folder = input_folder
        self.output_folder = output_folder
        self.Sample_Name = fnmatch.filter(os.listdir(input_folder),'*.xlsx')
        self.Sample_Size = len(self.Sample_Name)
        self.Output_Name = os.path.join(output_folder, str(self.Sample_Name) + '_zh_added.xlsx')

    def zh_adding(self):
        df_ref = pd.read_excel(ref)
        os.chdir(self.input_folder)
        for Sample in self.Sample_Name:
                input_file = os.path.join(self.input_folder, Sample)
                df = pd.read_excel(input_file, engine='openpyxl')
                df_ref_merged = df.merge(df_ref, how='left', on='Scientific_name')
                output_name = os.path.splitext(Sample)[0] + '_zh_added.xlsx'
                output_path = os.path.join(self.output_folder, output_name)
                df_ref_merged.to_excel(output_path, engine='openpyxl', index=False)


zh = zh_adder(input_folder, output_folder)
zh.zh_adding()