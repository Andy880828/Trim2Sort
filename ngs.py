import tkinter as tk
from tkinter import filedialog
import customtkinter
from PIL import Image
import os
import shutil
import pandas as pd
import numpy as np
import natsort

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

        self.cutadapt_path = tk.StringVar()
        self.usearch_path = tk.StringVar()
        self.blastn_path = tk.StringVar()
        self.database_path = tk.StringVar()
        self.samples_path = tk.StringVar()
        self.outputs_path = tk.StringVar()

        blastn_info = tk.StringVar(value="Step3: Select the directory of Blastn.exe")
        cutadapt_info = tk.StringVar(value="Step1: Select the directory of Cutadapt.exe")
        usearch_info = tk.StringVar(value="Step2: Select the directory of Usearch.exe")
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
        #Remane samples
        input_list = os.listdir(str(self.samples_entry.get() + '/'))
        for num in range(len(input_list)):
            label_num = str(num).rjust(3,'0')
            oldname = input_list[num]
            newname = label_num + '_' + oldname
            os.rename(self.samples_entry.get() + '/' + oldname, self.samples_entry.get() + '/' + newname)
        
        input_list = os.listdir(str(self.samples_entry.get() + '/' ))
        input_list.sort(key=lambda x: int(x.split('_')[0]))

        #Create output folders
        folders = ["A_primer_trimming", "B_merged", "C_quality", "D_length", "E_uniques", "F_OTUs", "G_OTUtable", "H_blasts", "I_sorted_blasts"]
        for folder in folders:
            path = os.path.join(str(self.outputs_entry.get()), folder)
            if not os.path.exists(path):
                os.makedirs(path)
                globals()['{}'.format(folder)] = path
        else:
            shutil.rmtree(path)
            os.makedirs(path)
            globals()['{}'.format(folder)] = path    

        Sample_Size = len(input_list)//2
        os.chdir(str(self.samples_entry.get() + '/'))

        #Trim primers
        for i in range(0, Sample_Size*2, 2):
            self.TrimPrimers(input_list[i],input_list[i+1])
        A_primer_trimming_folder_forsort = os.listdir(A_primer_trimming)
        A_primer_trimming_folder_forsort.sort(key=lambda x: int(x.split('_')[0]))   

        #Merge pair-ends
        for i in range(0, Sample_Size*2, 2):
            self.MergePairs(A_primer_trimming_folder_forsort[i],A_primer_trimming_folder_forsort[i+1])
        B_merged_folder_forsort = os.listdir(B_merged)
        B_merged_folder_forsort.sort(key=lambda x: int(x.split('_')[0]))

        #Trim unqualified nucleotides
        for i in range(0,Sample_Size):
            self.QualityControl(B_merged_folder_forsort[i])
        qualified_folder_forsort = os.listdir(C_quality)
        qualified_folder_forsort.sort(key=lambda x: int(x.split('_')[0]))
        
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
        D_length_folder_forsort.sort(key=lambda x: int(x.split('_')[0]))
        
        #Cluster sequence into OTU
        for i in range(0,Sample_Size_1):
            self.Cluster(D_length_folder_forsort[i])
        E_uniques_folder_forsort = os.listdir(E_uniques)
        E_uniques_folder_forsort.sort(key=lambda x: int(x.split('_')[0]))
        for i in range(0,Sample_Size_1):
            self.OTU(E_uniques_folder_forsort[i])
        OTU_folder_forsort = os.listdir(F_OTUs)
        OTU_folder_forsort.sort(key=lambda x: int(x.split('_')[0]))
        
        #Make OTUtable
        for i in range(0,Sample_Size_1):
            self.OTUtable(B_merged_folder_forsort[i], OTU_folder_forsort[i])
        OTUtab_folder_forsort = os.listdir(G_OTUtable)
        OTUtab_folder_forsort.sort(key=lambda x: int(x.split('_')[0])) 
        
        #Rename OTUtable
        for i in range(1, Sample_Size_1*2+1, 2):
            self.RenameOTUtable(OTUtab_folder_forsort[i])
        OTU_folder_forsort = os.listdir(F_OTUs)
        OTU_folder_forsort.sort(key=lambda x: int(x.split('_')[0]))
        
        #Blast OTU with database
        os.chdir(str(self.database_entry.get()))
        for i in range(0,Sample_Size_1):
            self.Blast(OTU_folder_forsort[i]) 
        
        #Combine blast results with OTUtable
        os.chdir(H_blasts)
        txts = sorted(os.listdir(H_blasts))
        os.chdir(G_OTUtable)
        reads = sorted(os.listdir(G_OTUtable))
        df_ref = pd.read_excel(ref)
        for n, i  in zip(range(0, len(txts)), range(1, len(reads), 2)):
            if txts[n].endswith('.txt'):
                original_data = pd.read_csv(H_blasts + '/' + txts[n], engine='python', header=None, sep='\t', encoding='utf-8', names=['OTU', 'Identity', 'Coverage', 'Scientific_name', 'Accession_number'])
                reads_data = pd.read_csv(reads[i], engine='python', header=None, sep='\t', encoding='utf-8', names=['OTU', 'Reads'], skiprows=1)
                total_reads = reads_data.Reads.sum()
                csv_data = pd.merge(original_data, reads_data)
                ratio = csv_data.Reads / float(total_reads)
                csv_data.insert(6, column='Ratio', value=ratio)
                csv_data.loc[csv_data['Identity'] >= 97, 'Stat'] = '*DUPE*'
                csv_data.loc[csv_data['Identity'] < 97, 'Stat'] = '*FILTERED*'
                # if (csv_data['Identity'] >= 97).any():
                #     csv_data['Stat'] = '*FILTERED*'
                # elif (csv_data['Identity'] < 97).any():
                #     csv_data['Stat'] = '*DUPE*'

                drop_duplicate_data = csv_data.drop_duplicates(subset=['OTU', 'Scientific_name'], keep='first').copy()
                drop_duplicate_data = drop_duplicate_data.sort_values(by='Identity', ascending=False)
                drop_duplicate_data.loc[drop_duplicate_data.duplicated(subset=['OTU'], keep='first'), ['Reads', 'Ratio']] = ''
                clean_data = drop_duplicate_data[drop_duplicate_data['Stat'] == '*DUPE*'].groupby('OTU')['Identity'].idxmax()
                drop_duplicate_data.loc[clean_data, 'Stat'] = ''
              
                final_csv = drop_duplicate_data[['OTU', 'Identity', 'Coverage', 'Scientific_name', 'Stat', 'Reads', 'Ratio', 'Accession_number']]
                final_csv = final_csv.sort_values('OTU', key=natsort.natsort_keygen())
                final_zh_csv = final_csv.merge(df_ref, how='left', on='Scientific_name')
                final_csv = final_csv.style.apply(Highlight, axis=1)
                final_zh_csv = final_zh_csv.style.apply(Highlight, axis=1)
                final_csv.to_excel(I_sorted_blasts + '/' + '{}'.format(reads[i].split('.')[0]) + '.xlsx', engine='openpyxl', index=False, encoding='utf-8')
                final_zh_csv.to_excel(I_sorted_blasts + '/' + '{}_zh_added'.format(reads[i].split('.')[0]) + '.xlsx', engine='openpyxl', index=False, encoding='utf-8')
        
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