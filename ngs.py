import tkinter as tk
from tkinter import filedialog
import customtkinter
from PIL import Image
import os
import shutil
import pandas as pd
import numpy as np
import natsort
import gc

def Highlight(row):
    condition = row['Stat']

    if condition == '*DUPE*': 
        css = 'background-color: #00c2c7'
    elif condition == '*FILTERED*':
        css = 'background-color: #94F7B2'
    else:
        css = 'background-color: transparent'

    return [css] * len(row)

ref = "C:\\Users\\andy0\\Desktop\\Trim2Sort_ver3.0.0\\ref_ver6_0918.xlsx"

my_image = customtkinter.CTkImage(light_image=Image.open("Trim2Sort_icon.png"),size=(128,72))

# NGS主畫面設定
class NGS(customtkinter.CTkToplevel):
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.title("NGS Analysis")
        self.geometry("500x600")
        self.configure(fg_color="#091235")
        self.resizable(False, False)
        self.grid_rowconfigure(0, weight=0)  # configure grid system
        self.grid_columnconfigure(0, weight=1)

        self.frame = NGS_HeaderFrame(master=self)
        self.frame.configure(fg_color="#091235")
        self.frame.grid(row=1, column=0, padx=10, pady=(0, 15))

        self.frame_2 = NGS_ContentFrame(master=self)
        self.frame_2.configure(fg_color="#091235",
                               height=470)
        self.frame_2.grid(row=2, column=0, padx=20, pady=0, sticky="nsew")

# NGS HeaderFrame設定
class NGS_HeaderFrame(customtkinter.CTkFrame):
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

# NGS ContentFrame設定
class NGS_ContentFrame(customtkinter.CTkFrame):
    def __init__(self, master, **kwargs):
        super().__init__(master, **kwargs)

        root_dir = os.path.dirname(os.path.abspath(__file__))
        
        # 自動找出目錄中的cutadapt.exe及usearch.exe
        cutadapt_exe = os.path.join(root_dir, "cutadapt.exe")
        usearch_exe = os.path.join(root_dir, "usearch.exe")
        
        self.cutadapt_path = tk.StringVar()
        self.usearch_path = tk.StringVar()
        self.blastn_path = tk.StringVar()
        self.database_path = tk.StringVar()
        self.samples_path = tk.StringVar()
        self.outputs_path = tk.StringVar()
        
        # Auto-set paths if files exist
        if os.path.exists(cutadapt_exe):
            self.cutadapt_path.set(cutadapt_exe)
        if os.path.exists(usearch_exe):
            self.usearch_path.set(usearch_exe)

        blastn_info = tk.StringVar(value="Step3: Select the directory of Blastn.exe")
        cutadapt_info = tk.StringVar(value="Step1: Cutadapt.exe (Auto-detected from root directory)")
        usearch_info = tk.StringVar(value="Step2: Usearch.exe (Auto-detected from root directory)")
        database_info = tk.StringVar(value="Step4: Select the folder directory of Database")
        samples_info = tk.StringVar(value="Step5: Select the folder directory of Samples")
        outputs_info = tk.StringVar(value="Step6: Select the folder directory of Outputs")
        instruction = tk.StringVar(value="Start the app after everthing above is selected")

        self.cutadapt_info = customtkinter.CTkLabel(self)
        self.cutadapt_info.configure(textvariable=cutadapt_info,
                                    text_color="#DCF3F0",
                                    font=("Consolas", 12, "bold"))
        self.cutadapt_info.grid(row=0, column=0, padx=0, sticky="w")

        self.cutadapt_entry = customtkinter.CTkEntry(self)
        self.cutadapt_entry.configure(textvariable=self.cutadapt_path,
                                    text_color="#DCF3F0",
                                    width=370,
                                    height=25,
                                    border_width=2,
                                    corner_radius=8,
                                    fg_color="#2B4257",
                                    font=("Consolas", 12, "bold"))
        self.cutadapt_entry.grid(row=1, column=0, padx=0, pady=(0, 10), sticky="w")

        self.cutadapt_button = customtkinter.CTkButton(self)
        self.cutadapt_button.configure(width=80,
                                 height=25,
                                 border_width=0,
                                 corner_radius=8,
                                 text="SELECT", 
                                 fg_color="#2B4257",
                                 text_color="#DCF3F0",
                                 hover_color="#4C6A78",
                                 font=("Consolas", 11, "bold"),
                                 command=self.BrowseCutadapt)
        self.cutadapt_button.grid(row=1, column=1, padx=(10, 0), pady=(0, 10), sticky="w")


        self.usearch_info = customtkinter.CTkLabel(self)
        self.usearch_info.configure(textvariable=usearch_info,
                                    text_color="#DCF3F0",
                                    font=("Consolas", 12, "bold"))
        self.usearch_info.grid(row=2, column=0, padx=0, sticky="w")

        self.usearch_entry = customtkinter.CTkEntry(self)
        self.usearch_entry.configure(textvariable=self.usearch_path,
                                    text_color="#DCF3F0",
                                    width=370,
                                    height=25,
                                    border_width=2,
                                    corner_radius=8,
                                    fg_color="#2B4257",
                                    font=("Consolas", 12, "bold"))
        self.usearch_entry.grid(row=3, column=0, padx=0, pady=(0, 10), sticky="w")
        self.usearch_button = customtkinter.CTkButton(self)
        self.usearch_button.configure(width=80,
                                    height=25,
                                    border_width=0,
                                    corner_radius=8,
                                    text="SELECT", 
                                    fg_color="#2B4257",
                                    text_color="#DCF3F0",
                                    hover_color="#4C6A78",
                                    font=("Consolas", 11, "bold"),
                                    command=self.BrowseUsearch)
        self.usearch_button.grid(row=3, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

        self.blastn_info = customtkinter.CTkLabel(master=self)
        self.blastn_info.configure(textvariable=blastn_info,
                                    text_color="#DCF3F0",
                                    font=("Consolas", 12, "bold"))
        self.blastn_info.grid(row=4, column=0, padx=0, sticky="w")

        self.blastn_entry = customtkinter.CTkEntry(self)
        self.blastn_entry.configure(textvariable=self.blastn_path,
                                    text_color="#DCF3F0",
                                    width=370,
                                    height=25,
                                    border_width=2,
                                    corner_radius=8,
                                    fg_color="#2B4257",
                                    font=("Consolas", 12, "bold"))
        self.blastn_entry.grid(row=5, column=0, padx=0, pady=(0, 10), sticky="w")
        self.blastn_button = customtkinter.CTkButton(master=self)
        self.blastn_button.configure(width=80,
                                    height=25,
                                    border_width=0,
                                    corner_radius=8,
                                    text="SELECT", 
                                    fg_color="#2B4257",
                                    text_color="#DCF3F0",
                                    hover_color="#4C6A78",
                                    font=("Consolas", 11, "bold"),
                                    command=self.BrowseBlastn)
        self.blastn_button.grid(row=5, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

        
        self.database_info = customtkinter.CTkLabel(master=self)
        self.database_info.configure(textvariable=database_info,
                                    text_color="#DCF3F0",
                                    font=("Consolas", 12, "bold"))
        self.database_info.grid(row=6, column=0, padx=0, sticky="w")

        self.database_entry = customtkinter.CTkEntry(self)
        self.database_entry.configure(textvariable=self.database_path,
                                    text_color="#DCF3F0",
                                    width=370,
                                    height=25,
                                    border_width=2,
                                    corner_radius=8,
                                    fg_color="#2B4257",
                                    font=("Consolas", 12, "bold"))
        self.database_entry.grid(row=7, column=0, padx=0, pady=(0, 10), sticky="w")
        self.database_button = customtkinter.CTkButton(master=self)
        self.database_button.configure(width=80,
                                    height=25,
                                    border_width=0,
                                    corner_radius=8,
                                    text="SELECT", 
                                    fg_color="#2B4257",
                                    text_color="#DCF3F0",
                                    hover_color="#4C6A78",
                                    font=("Consolas", 11, "bold"),
                                    command=self.BrowseDatabase)
        self.database_button.grid(row=7, column=1, padx=(10, 0), pady=(0, 10), sticky="w")


        self.samples_info = customtkinter.CTkLabel(self)
        self.samples_info.configure(textvariable=samples_info,
                                    text_color="#DCF3F0",
                                    font=("Consolas", 12, "bold"))
        self.samples_info.grid(row=8, column=0, padx=0, sticky="w")

        self.samples_entry = customtkinter.CTkEntry(self)
        self.samples_entry.configure(textvariable=self.samples_path,
                                    text_color="#DCF3F0",
                                    width=370,
                                    height=25,
                                    border_width=2,
                                    corner_radius=8,
                                    fg_color="#2B4257",
                                    font=("Consolas", 12, "bold"))
        self.samples_entry.grid(row=9, column=0, padx=0, pady=(0, 10), sticky="w")
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
        self.samples_button.grid(row=9, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

        self.outputs_info = customtkinter.CTkLabel(self)
        self.outputs_info.configure(textvariable=outputs_info,
                                    text_color="#DCF3F0",
                                    font=("Consolas", 12, "bold"))
        self.outputs_info.grid(row=10, column=0, padx=0, sticky="w")

        self.outputs_entry = customtkinter.CTkEntry(self)
        self.outputs_entry.configure(textvariable=self.outputs_path,
                                    text_color="#DCF3F0",
                                    width=370,
                                    height=25,
                                    border_width=2,
                                    corner_radius=8,
                                    fg_color="#2B4257",
                                    font=("Consolas", 12, "bold"))
        self.outputs_entry.grid(row=11, column=0, padx=0, pady=(0, 10), sticky="w")
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
        self.outputs_button.grid(row=11, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

        
        self.instruction_label = customtkinter.CTkLabel(self)
        self.instruction_label.configure(textvariable=instruction,
                                        text_color="#DCF3F0",
                                        font=("Consolas", 12, "bold"))
        self.instruction_label.grid(row=12, column=0, columnspan=2, padx=0, pady=(10, 0) ,sticky="ew")
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
                                        command=self.Analyse)
        self.instruction_button.grid(row=13, column=0, columnspan=2, padx=0, sticky="ew")

    def TrimPrimers(self,R1,R2):
        trimming = "{exe_path}\
        -a GTCGGTAAAACTCGTGCCAGC...CAAACTGGGATTAGATACCCCACTATG \
        -A CATAGTGGGGTATCTAATCCCAGTTTG...GCTGGCACGAGTTTTACCGAC \
        --discard-untrimmed \
        -o {outpath}/{R1}_TRIMMED_R1.fastq \
        -p {outpath}/{R2}_TRIMMED_R2.fastq {R1} {R2}".format(exe_path=self.cutadapt_entry.get(), outpath=A_primer_trimming, R1=R1, R2=R2)
        os.system(trimming)

    def MergePairs(self,R1,R2):
        merging = "{exe_path}\
        -fastq_mergepairs {inpath}/{R1}\
        -reverse {inpath}/{R2}\
        -fastqout {outpath}/{R1}_merged.fastq".format(exe_path=self.usearch_entry.get(), inpath=A_primer_trimming, R1=R1, R2=R2, outpath=B_merged)
        os.system(merging)

    def QualityControl(self,file):
        qualifying = "{exe_path}\
        -fastq_filter {inpath}/{file} \
        -fastq_truncqual 20 \
        -fastqout {outpath}/{file}_QUAL.fastq".format(exe_path=self.usearch_entry.get(), inpath=B_merged, file=file, outpath=C_quality)
        os.system(qualifying)

    def FilterLength(self,file):
        D_length_filtering = "{exe_path}\
        -fastq_filter {inpath}/{file} \
        -fastq_minlen 150 \
        -fastaout {outpath}/{file}_LENG.fasta".format(exe_path=self.usearch_entry.get(), inpath=C_quality, file=file, outpath=D_length)
        os.system(D_length_filtering)

    def Cluster(self,file):
        clustering = "{exe_path}\
        -fastx_uniques {inpath}/{file} \
        -fastaout {outpath}/{file}_UNIQ.fasta \
        -sizeout -relabel Uniq".format(exe_path=self.usearch_entry.get(), inpath=D_length, file=file, outpath=E_uniques)
        os.system(clustering)

    def OTU(self,file):
        OTU_making = "{exe_path}\
        -cluster_otus {inpath}/{file} \
        -otus {outpath1}/{file}_OTU.fasta \
        -relabel Otu".format(exe_path=self.usearch_entry.get(), inpath=E_uniques, file=file, outpath1=F_OTUs)
        os.system(OTU_making)    

    def OTUtable(self,file1,file2):
        OTUtab_making = "{exe_path}\
        -otutab {inpath1}/{file1} \
        -otus {inpath2}/{file2} \
        -otutabout {outpath}/{file2}_table.txt \
        -mapout {outpath}/{file2}_map.txt".format(exe_path=self.usearch_entry.get(), inpath1=B_merged, inpath2=F_OTUs, file1=file1, file2=file2, outpath=G_OTUtable)
        os.system(OTUtab_making)

    def RenameOTUtable(self,file):
        oldfile = "{}/{}".format(G_OTUtable,file)
        newname = "{}/{}.txt".format(G_OTUtable,str(file).split('.')[0])
        os.rename(oldfile, newname)

    def Blast(self,file):
        result = "{blastn} -query {inpath}/{file} -out \
        {outpath}/{file}_blasted.txt -db COI_Osteichthyes \
        -outfmt \"6 qseqid pident qcovs sscinames sacc\" -max_target_seqs 3".format(blastn=self.blastn_entry.get(),file=file, inpath=F_OTUs, outpath=H_blasts)
        os.system(result)   

    def CleanUpFilename(self,file):
        oldfile = "{}/{}".format(I_sorted_blasts,file)
        newname = "{}/{}.xlsx".format(I_sorted_blasts,str(file).split('.')[0])
        os.rename(oldfile, newname)

    def BrowseCutadapt(self):
        filename = filedialog.askopenfilenames()
        self.cutadapt_path.set(filename)

    def BrowseUsearch(self):
        filename = filedialog.askopenfilenames()
        self.usearch_path.set(filename)

    def BrowseBlastn(self):
        filename = filedialog.askopenfilenames()
        self.blastn_path.set(filename)

    def BrowseDatabase(self):
        foldername = filedialog.askdirectory()
        self.database_path.set(foldername)

    def BrowseSamples(self):
        foldername = filedialog.askdirectory()
        self.samples_path.set(foldername)

    def BrowseOutputs(self):
        foldername = filedialog.askdirectory()
        self.outputs_path.set(foldername)

    def Analyse(self):
        #Get input files (no renaming)
        input_list = os.listdir(str(self.samples_entry.get() + '/'))
        input_list = natsort.natsorted(input_list)

        #Create output folders (auto cleanup)
        # 先強制垃圾回收，確保所有檔案句柄都已釋放
        gc.collect()
        
        folders = ["A_primer_trimming", "B_merged", "C_quality", "D_length", "E_uniques", "F_OTUs", "G_OTUtable", "H_blasts", "I_sorted_blasts"]
        for folder in folders:
            path = os.path.join(str(self.outputs_entry.get()), folder)
            if os.path.exists(path):
                try:
                    shutil.rmtree(path)
                except PermissionError:
                    # 如果無法刪除，嘗試清空內容
                    for item in os.listdir(path):
                        item_path = os.path.join(path, item)
                        try:
                            if os.path.isdir(item_path):
                                shutil.rmtree(item_path, ignore_errors=True)
                            else:
                                os.remove(item_path)
                        except:
                            pass
            os.makedirs(path, exist_ok=True)
            globals()['{}'.format(folder)] = path    

        Sample_Size = len(input_list)//2
        os.chdir(str(self.samples_entry.get() + '/'))

        #Trim primers
        for i in range(0, Sample_Size*2, 2):
            self.TrimPrimers(input_list[i],input_list[i+1])
        A_primer_trimming_folder_forsort = os.listdir(A_primer_trimming)
        A_primer_trimming_folder_forsort = natsort.natsorted(A_primer_trimming_folder_forsort)   

        #Merge pair-ends
        for i in range(0, Sample_Size*2, 2):
            self.MergePairs(A_primer_trimming_folder_forsort[i],A_primer_trimming_folder_forsort[i+1])
        B_merged_folder_forsort = os.listdir(B_merged)
        B_merged_folder_forsort = natsort.natsorted(B_merged_folder_forsort)

        #Trim unqualified nucleotides
        for i in range(0,Sample_Size):
            self.QualityControl(B_merged_folder_forsort[i])
        qualified_folder_forsort = os.listdir(C_quality)
        qualified_folder_forsort = natsort.natsorted(qualified_folder_forsort)
        
        #Adjust lengh of sequence & Delete 0 byte file
        for i in range(0,Sample_Size):
            self.FilterLength(qualified_folder_forsort[i])
        
        # f_size = os.listdir(D_length)
        # f_other = os.listdir(B_merged)
        # print(f_size)
        # for n in len(f_size):
        #     file_size = f_size[n]
        #     f_other_for_del = f_other[n]
        #     size = os.stat(file_size)
        # if size.st_size == 0:
        #     shutil.rmtree(file_size,ignore_errors=True)
        #     shutil.rmtree(f_other_for_del,ignore_errors=True)

        D_length_folder_forsort = os.listdir(D_length)
        Sample_Size_1 = len(D_length_folder_forsort)
        D_length_folder_forsort = natsort.natsorted(D_length_folder_forsort)
        
        #Cluster sequence into OTU
        for i in range(0,Sample_Size_1):
            self.Cluster(D_length_folder_forsort[i])
        E_uniques_folder_forsort = os.listdir(E_uniques)
        E_uniques_folder_forsort = natsort.natsorted(E_uniques_folder_forsort)
        for i in range(0,Sample_Size_1):
            self.OTU(E_uniques_folder_forsort[i])
        OTU_folder_forsort = os.listdir(F_OTUs)
        OTU_folder_forsort = natsort.natsorted(OTU_folder_forsort)
        
        #Make OTUtable
        for i in range(0,Sample_Size_1):
            self.OTUtable(B_merged_folder_forsort[i], OTU_folder_forsort[i])
        OTUtab_folder_forsort = os.listdir(G_OTUtable)
        OTUtab_folder_forsort = natsort.natsorted(OTUtab_folder_forsort) 
        
        #Rename OTUtable
        for i in range(1, Sample_Size_1*2+1, 2):
            self.RenameOTUtable(OTUtab_folder_forsort[i])
        OTU_folder_forsort = os.listdir(F_OTUs)
        OTU_folder_forsort = natsort.natsorted(OTU_folder_forsort)
        
        #Blast OTU with database
        os.chdir(str(self.database_entry.get()))
        for i in range(0,Sample_Size_1):
            self.Blast(OTU_folder_forsort[i]) 
        
        # Combine blast results with OTUtable
        os.chdir(H_blasts)
        txts = sorted(os.listdir(H_blasts))
        os.chdir(G_OTUtable)
        reads = sorted(os.listdir(G_OTUtable))
        df_ref = pd.read_excel(ref)

        for n, i in zip(range(0, len(txts)), range(1, len(reads), 2)):
            if txts[n].endswith('.txt'):
                original_data = pd.read_csv(
                    H_blasts + '/' + txts[n], 
                    engine='python', 
                    header=None, 
                    sep='\t', 
                    encoding='utf-8', 
                    names=['OTU', 'Identity', 'Coverage', 'Scientific_name', 'Accession_number']
                )
                
                # 讀取 reads 資料
                reads_data = pd.read_csv(
                    reads[i], 
                    engine='python', 
                    header=None, 
                    sep='\t', 
                    encoding='utf-8', 
                    names=['OTU', 'Reads'], 
                    skiprows=1
                )
                total_reads = reads_data.Reads.sum()
                
                # 分離符合條件和不符合條件的資料
                grouped_otu = (original_data[original_data.Identity >= 97]).reset_index(drop=True)
                drop_filtered = original_data[original_data.Identity < 97].copy()
                drop_filtered['Stat'] = '*FILTERED*'
                drop_filtered = drop_filtered.drop_duplicates(subset=['OTU', 'Scientific_name'], keep='first')
                
                # 處理符合條件的資料（Identity >= 97）
                grouped_otu['Stat'] = np.where(grouped_otu['OTU'].duplicated(), '*DUPE*', '')
                drop_dupe = grouped_otu[grouped_otu.Stat == '*DUPE*']
                drop_dupe = drop_dupe.drop_duplicates(subset=['OTU', 'Scientific_name'], keep='first')
                clean_data = grouped_otu.drop_duplicates(subset=['OTU'], keep='first')
                
                # 合併 reads 到符合條件的資料
                csv_data = pd.merge(clean_data, reads_data)
                ratio = csv_data.Reads / float(total_reads)
                csv_data.insert(7, column='Ratio', value=ratio)
                
                # 組合主要資料（包含 clean_data 和 drop_dupe）
                main_data = pd.concat([csv_data, drop_dupe], axis=0, ignore_index=True)
                main_data = main_data.sort_values(['OTU', 'Stat'], ignore_index=True)
                main_data = main_data.groupby('OTU', include_groups=False).apply(
                    lambda x: x.drop_duplicates(subset=['Scientific_name'], keep='first')
                ).reset_index(drop=True)
                
                # 找出哪些 OTU 有符合條件的資料
                otus_with_valid_data = set(main_data['OTU'].unique())
                
                # 處理 FILTERED 資料
                # 1. 找出完全沒有符合條件的 OTU（所有 row 都是 FILTERED）
                all_filtered_otus = set(drop_filtered['OTU'].unique()) - otus_with_valid_data
                
                # 2. 對於完全 FILTERED 的 OTU，保留第一個 row 並加上 reads/ratio
                filtered_with_reads = drop_filtered[drop_filtered['OTU'].isin(all_filtered_otus)].copy()
                filtered_with_reads = filtered_with_reads.drop_duplicates(subset=['OTU'], keep='first')
                filtered_with_reads = pd.merge(filtered_with_reads, reads_data, on='OTU')
                ratio_filtered = filtered_with_reads.Reads / float(total_reads)
                filtered_with_reads['Ratio'] = ratio_filtered
                
                # 3. 對於有符合條件資料的 OTU，其 FILTERED rows 不顯示 reads/ratio
                filtered_without_reads = drop_filtered[~drop_filtered['OTU'].isin(all_filtered_otus)].copy()
                filtered_without_reads['Reads'] = np.nan
                filtered_without_reads['Ratio'] = np.nan
                
                # 合併所有 FILTERED 資料
                all_filtered = pd.concat([filtered_with_reads, filtered_without_reads], axis=0, ignore_index=True)
                
                # 合併所有資料
                final_csv = pd.concat([main_data, all_filtered], axis=0, ignore_index=True)
                final_csv = final_csv[['OTU', 'Identity', 'Coverage', 'Scientific_name', 'Stat', 'Reads', 'Ratio', 'Accession_number']]
                final_csv = final_csv.sort_values('OTU', key=natsort.natsort_keygen())
                
                # 合併中文名稱
                final_zh_csv = final_csv.merge(df_ref, how='left', on='Scientific_name')
                
                # 套用樣式
                final_csv_styled = final_csv.style.apply(Highlight, axis=1)
                final_zh_csv_styled = final_zh_csv.style.apply(Highlight, axis=1)
                
                # 使用 ExcelWriter context manager 確保檔案正確關閉
                excel_path1 = os.path.join(I_sorted_blasts, '{}'.format(reads[i].split('.')[0]) + '.xlsx')
                excel_path2 = os.path.join(I_sorted_blasts, '{}_zh_added'.format(reads[i].split('.')[0]) + '.xlsx')
                
                with pd.ExcelWriter(excel_path1, engine='openpyxl') as writer:
                    final_csv_styled.to_excel(writer, index=False, sheet_name='Sheet1')
                
                with pd.ExcelWriter(excel_path2, engine='openpyxl') as writer:
                    final_zh_csv_styled.to_excel(writer, index=False, sheet_name='Sheet1')
                
                # 強制垃圾回收以釋放檔案句柄
                del final_csv_styled, final_zh_csv_styled
                gc.collect()
        
        # os.chdir(I_sorted_blasts)
        # excel_lists = os.listdir(I_sorted_blasts)
        # for excel_list in excel_lists:
        #     species_list = pd.read_csv(excel_list, encoding='utf_8_sig')
        #     with pd.ExcelWriter(excel_list) as writer:
        #         species_list.to_excel(writer, sheet_name='Sheet1', index=False, encoding='utf_8_sig')
        #         for column in species_list:
        #             column_length = max(species_list[column].astype(str).map(len).max(), len(column))
        #             col_idx = species_list.columns.get_loc(column)
        #             writer.sheets['Sheet1'].set_column(col_idx, col_idx, column_length+2)


        # writer.save()

        #Cleanup final filename
        os.chdir(I_sorted_blasts)
        files = sorted(os.listdir(I_sorted_blasts))
        for file in files:
            self.CleanUpFilename(file) 

        print("Produced by 高屏溪男子偶像團體")