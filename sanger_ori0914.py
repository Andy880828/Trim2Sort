import tkinter as tk
from tkinter import filedialog
import customtkinter
from PIL import Image
import os
import pandas as pd
from datetime import datetime
import fnmatch
import subprocess


my_image = customtkinter.CTkImage(light_image=Image.open("Trim2Sort_icon.png"),size=(128,72))
ref = "D:\\Trim2Sort_ver3.0.0\\ref_ver5_0831.xlsx"
# def Highlight(row):
#     condition = row['identity']

#     if condition < 99: 
#         css = 'background-color: #ff6b6b'
#     else:
#         css = 'background-color: transparent'

#     return [css] * len(row)

# Sanger主畫面設定
class Sanger(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("Sanger Analysis")
        self.geometry("500x400")
        self.configure(fg_color="#091235")
        self.resizable(False, False)
        self.grid_rowconfigure(0, weight=0)  # configure grid system
        self.grid_columnconfigure(0, weight=1)

        self.frame = Sanger_HeaderFrame(master=self)
        self.frame.configure(fg_color="#091235")
        self.frame.grid(row=1, column=0, padx=10, pady=(0, 15))

        self.frame_2 = Sanger_ContentFrame(master=self)
        self.frame_2.configure(fg_color="#091235",
                               height=470)
        self.frame_2.grid(row=2, column=0, padx=20, pady=0, sticky="nsew")

# Sanger HeaderFrame設定
class Sanger_HeaderFrame(customtkinter.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        self.logo = customtkinter.CTkLabel(self, image=my_image,text="")
        self.logo.grid(row=0, column=2, padx=10)
        self.button_UserGuide = customtkinter.CTkButton(self)
        self.button_UserGuide.configure(width=100,
                                        height=25,
                                        border_width=0,
                                        text="USER GUIDE", 
                                        fg_color="transparent",
                                        text_color="#88A9C3",
                                        hover_color="#2B4257",
                                        font=("Consolas", 11, "bold"))
        self.button_UserGuide.grid(row=1, column=0, padx=10)
        self.sep_label = customtkinter.CTkLabel(master=self)
        self.sep_label.configure(text="✦",
                                text_color="#88A9C3",
                                font=("Consolas", 10, "bold"))
        self.sep_label.grid(row=1, column=1, padx=10)
        self.button_Documentation = customtkinter.CTkButton(master=self)
        self.button_Documentation.configure(width=100,
                                            height=25,
                                            border_width=0,
                                            text="DOCUMENTATION", 
                                            fg_color="transparent",
                                            text_color="#88A9C3",
                                            hover_color="#2B4257",
                                            font=("Consolas", 11, "bold"))
        self.button_Documentation.grid(row=1, column=2, padx=10)
        self.sep_label = customtkinter.CTkLabel(master=self)
        self.sep_label.configure(text="✦",
                                text_color="#88A9C3",
                                font=("Consolas", 10, "bold"))
        self.sep_label.grid(row=1, column=3, padx=10)
        self.button_ContactUs = customtkinter.CTkButton(self)
        self.button_ContactUs.configure(width=100,
                                        height=25,
                                        border_width=0,
                                        text="CONTACT US", 
                                        fg_color="transparent",
                                        text_color="#88A9C3",
                                        hover_color="#2B4257",
                                        font=("Consolas", 11, "bold"))
        self.button_ContactUs.grid(row=1, column=4, padx=10)

# Sanger ContentFrame設定
class Sanger_ContentFrame(customtkinter.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        


        self.samples_path = tk.StringVar()
        self.outputs_path = tk.StringVar()

        database_info = tk.StringVar(value="Step1: Select the folder directory of Database")
        samples_info = tk.StringVar(value="Step2: Select the folder directory of Samples")
        outputs_info = tk.StringVar(value="Step3: Select the folder directory of Outputs")
        instruction = tk.StringVar(value="Start the app after everthing above is selected")

        self.database_info = customtkinter.CTkLabel(master=self)
        self.database_info.configure(textvariable=database_info,
                                    text_color="#DCF3F0",
                                    font=("Consolas", 12, "bold"))
        self.database_info.grid(row=0, column=0, padx=0, sticky="w")

    
        # 下拉式選單
        self.database_combobox = customtkinter.CTkOptionMenu(self)
        self.database_combobox.configure(values=["Teleostei-12S", "Teleostei-COI"],
                                width=370,
                                height=25,
                                corner_radius=8,
                                fg_color="#2B4257",
                                button_color="#2B4257",
                                text_color="#DCF3F0",
                                dropdown_fg_color="#4C6A78",
                                dropdown_text_color="#DCF3F0",
                                font=("Consolas", 12, "bold"),
                                dropdown_font=("Consolas", 12, "bold"))
        self.database_combobox.set("Database")
        self.database_combobox.grid(row=1, column=0, padx=0, pady=(0, 10), sticky="w")

        self.samples_info = customtkinter.CTkLabel(self)
        self.samples_info.configure(textvariable=samples_info,
                                    text_color="#DCF3F0",
                                    font=("Consolas", 12, "bold"))
        self.samples_info.grid(row=2, column=0, padx=0, sticky="w")

        self.samples_entry = customtkinter.CTkEntry(self)
        self.samples_entry.configure(textvariable=self.samples_path,
                                    text_color="#DCF3F0",
                                    width=370,
                                    height=25,
                                    border_width=2,
                                    corner_radius=8,
                                    fg_color="#2B4257",
                                    font=("Consolas", 12, "bold"))
        self.samples_entry.grid(row=3, column=0, padx=0, pady=(0, 10), sticky="w")
        self.samples_button = customtkinter.CTkButton(self)
        self.samples_button.configure(width=80,
                                    height=25,
                                    border_width=0,
                                    corner_radius=8,
                                    text="SELECT", 
                                    fg_color="#2B4257",
                                    text_color="#DCF3F0",
                                    hover_color="#4C6A78",
                                    font=("Consolas", 11, "bold"),
                                    command=self.BrowseSamples)
        self.samples_button.grid(row=3, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

        self.outputs_info = customtkinter.CTkLabel(self)
        self.outputs_info.configure(textvariable=outputs_info,
                                    text_color="#DCF3F0",
                                    font=("Consolas", 12, "bold"))
        self.outputs_info.grid(row=4, column=0, padx=0, sticky="w")

        self.outputs_entry = customtkinter.CTkEntry(self)
        self.outputs_entry.configure(textvariable=self.outputs_path,
                                    text_color="#DCF3F0",
                                    width=370,
                                    height=25,
                                    border_width=2,
                                    corner_radius=8,
                                    fg_color="#2B4257",
                                    font=("Consolas", 12, "bold"))
        self.outputs_entry.grid(row=5, column=0, padx=0, pady=(0, 10), sticky="w")
        self.outputs_button = customtkinter.CTkButton(self)
        self.outputs_button.configure(width=80,
                                    height=25,
                                    border_width=0,
                                    corner_radius=8,
                                    text="SELECT", 
                                    fg_color="#2B4257",
                                    text_color="#DCF3F0",
                                    hover_color="#4C6A78",
                                    font=("Consolas", 11, "bold"),
                                    command=self.BrowseOutputs)
        self.outputs_button.grid(row=5, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

        
        self.instruction_label = customtkinter.CTkLabel(self)
        self.instruction_label.configure(textvariable=instruction,
                                        text_color="#DCF3F0",
                                        font=("Consolas", 12, "bold"))
        self.instruction_label.grid(row=6, column=0, columnspan=2, padx=0, pady=(10, 0) ,sticky="ew")
        self.instruction_button = customtkinter.CTkButton(self)
        self.instruction_button.configure(width=80,
                                        height=25,
                                        border_width=0,
                                        corner_radius=8,
                                        text="ANALYSE", 
                                        fg_color="#2B4257",
                                        text_color="#DCF3F0",
                                        hover_color="#4C6A78",
                                        font=("Consolas", 11, "bold"),
                                        command=self.Analysis)
        self.instruction_button.grid(row=7, column=0, columnspan=2, padx=0, sticky="ew")


    def BrowseSamples(self):
        foldername = filedialog.askdirectory()
        self.samples_path.set(foldername)

    def BrowseOutputs(self):
        foldername = filedialog.askdirectory()
        self.outputs_path.set(foldername)

    def Analysis(self):
        input_folder = str(self.samples_entry.get() + '/')
        output_folder = str(self.outputs_entry.get() + '/')

        sample_names = fnmatch.filter(os.listdir(input_folder), "*.txt")
        sample_names.sort(key=lambda x: str(x.split("_")[0]))
        sample_size = len(sample_names)

        output_fasta = (output_folder +
                        str(datetime.now().month) +
                        "." +
                        str(datetime.now().day) +
                        "_N=" +
                        str(sample_size) +
                        "_output.fasta")

        os.chdir(input_folder)
        i = 0
        with open(output_fasta, "w") as outfile:
            for sample in sample_names:
                with open(sample) as infile:
                    outfile.write(">" + str(sample_names[i]).split("_")[0] + "\n")
                    for line in infile:
                        outfile.write(line)
                i += 1

        os.chdir("../")
        cutadpat_cmd = f'cutadapt.exe -u 20 -o {output_fasta}_TRIMMED.fasta {output_fasta}'
        subprocess.run(cutadpat_cmd, shell=True)

        
        if self.database_combobox.get() == "Teleostei-12S":
            os.chdir("./db-12S/")
            blast_cmd = (
                f"blastn.exe -query {output_fasta}_TRIMMED.fasta -out {output_fasta}_blasted.txt"
                f" -db 12S_Osteichthyes_db -outfmt \"6 qseqid pident qcovs sscinames sacc\" -max_target_seqs 1"
                )
            subprocess.run(blast_cmd, shell=True)
        elif self.database_combobox.get() == "Teleostei-COI":
            os.chdir("./db-CO1/")
            blast_cmd = (
                f"blastn.exe -query {output_fasta}_TRIMMED.fasta -out {output_fasta}_blasted.txt"
                f" -db CO1_Osteichthyes_db -outfmt \"6 qseqid pident qcovs sscinames sacc slen\" -max_target_seqs 1"
                )
            subprocess.run(blast_cmd, shell=True)

        blast_list = set()
        in_name = set()
        with open(output_fasta + '_blasted.txt', 'r') as blast_result:
            success_samples = blast_result.read().splitlines()
        for success_sample in success_samples:
            blast_list.add(success_sample.split()[0])
        for sample in sample_names:
            in_name.add(str(sample).split('_')[0])
        fail_lists = in_name.difference(blast_list)
        dataset = pd.DataFrame(fail_lists,columns=['No'])
        blastresult = pd.read_csv(f'{output_fasta}_blasted.txt', header=None, sep='\t', names=['No','Identity','Coverage','Scientific_name','Accession_number', 'bp'])
        concatresult = pd.concat([dataset, blastresult])
        # concatresult = concatresult.style.apply(Highlight, axis=1)
        concatresult.to_excel(f'{output_fasta}.xlsx', engine='openpyxl', index=False)


        old_file_name = f'{output_fasta}.xlsx'
        new_file_name = f'{str(old_file_name).split("_")[0]}_{str(old_file_name).split("_")[1]}.xlsx'
        os.rename(old_file_name, new_file_name)
    
        df_ref = pd.read_excel(ref)
        df_ref.set_index('Scientific_name', inplace=True)
        df_input = pd.read_excel(new_file_name)
        df_input.set_index('Scientific_name', inplace=True)
        output_file_path = new_file_name.split(".xlsx")[0] + "_zh.xlsx"
        df_ref_merged = pd.merge(df_input, df_ref, left_index=True, right_index=True, how='left')
        df_ref_merged.insert(3, 'Scientific_name', df_ref_merged.index)
        df_ref_merged = df_ref_merged.sort_values(by='No')
        df_ref_merged.to_excel(output_file_path, index=False)
    
    

        print("Produced by 高屏溪男子偶像團體")
