import pathlib
from pathlib import Path
import datetime
import os
import pandas as pd
import csv
import re


#main source
blastn = "D:/Trim2Sort_ver3.0.0/db-12S/blastn.exe" #測試用
#blastn = input(f'Please input path of blastn.exe') #實際用
db_folder = "D:/Trim2Sort_ver3.0.0/db-12S" #測試用
#db_folder = input(f'Please input path of database') #實際用
cutadapt = "D:/Trim2Sort_ver3.0.0/cutadapt.exe" #測試用
#cutadapt = input(f'Please input path of cutadapt.exe') #實際用
rscript = "rscript"
dada2_seqdivider = "D:/DaDa2life/dada2_seqdivider.R"
ref_path = "D:/Trim2Sort_ver3.0.0/ref_ver5_0831.xlsx"
today = str(datetime.date.today())

# function def
# create all step folder and store the path
def folder_creater(foldername):
    new_folder = Path(pre_route + "/" + foldername + today + "/")
    try:
        new_folder.mkdir()
    except FileExistsError as exist_already:
        print(foldername + "exist already")
    new_folder_route = str(pre_route + "/" + foldername + today + "/")
    return new_folder_route

# make a list of number and help sorting
def extract_number(string):
    matches = re.findall(r'\d+', string)
    if matches:
        return int(matches[0])
    else:
        return 0



pre_route = input(f'Please input path of your folder:') #實際用
pre_route = pre_route.replace('\\', '/')
#run
#dada2 analysis
dada2 = "{rscript} {dada2_seqdivider}\
    --inpath {pre_route}\
    --cutadapt {cutadapt}\
    --FWD GTCGGTAAAACTCGTGCCAGC\
    --REV CATAGTGGGGTATCTAATCCCAGTTTG".format(rscript=rscript, dada2_seqdivider=dada2_seqdivider, pre_route=pre_route, cutadapt=cutadapt)
os.system(dada2)


# create all folders and store all routes
inroute = pre_route+"/dada2/"
blasted_folder_name = "blasted"
blasted_folder_route = folder_creater(blasted_folder_name)
final_report_name = "final_report"
final_report_route = folder_creater(final_report_name)
chinese_report_name = "ch_report"
chinese_report_route = folder_creater(chinese_report_name)

# Samplename detection
samples = os.listdir(inroute)
samplenames = []
for sample in samples:
    if not sample.endswith(".fasta"):
        continue
    samplename = sample.split("_dada2")[0]
    samplenames.append(samplename)

# blast
os.chdir(db_folder)
for samplename in samplenames:
    result = "{blastn} -query {inpath} -out \
    {outpath} -db 12S_Osteichthyes_db \
    -outfmt \"6 qseqid pident qcovs sscinames sallacc\" -max_target_seqs 2".format(blastn=blastn, inpath=inroute+samplename+"_dada2.fasta", outpath=blasted_folder_route+samplename+"_blasted.txt")
    os.system(result)   

# add index in blasted_file
for samplename in samplenames:
    with open(blasted_folder_route+samplename+"_blasted.txt", 'r') as blasted_file:
        reader = csv.reader(blasted_file, delimiter='\t')        
        origin_blasted_data = [row for row in reader]
    col_name = ['ASV', 'Identity', 'Coverage', 'Scientific_name', 'Accession_number']
    origin_blasted_data.insert(0, col_name)
    with open(blasted_folder_route+samplename+"_blasted.txt", 'w') as blasted_file:
        writer = csv.writer(blasted_file, delimiter='\t')
        writer.writerows(origin_blasted_data)

# mix readnum & blasted data
for samplename in samplenames:
    try:
        track_path = inroute+samplename+"_dada2track.csv"
        blasted_path = blasted_folder_route+samplename+"_blasted.txt"
        track_data = pd.read_csv(track_path)
        blasted_data = pd.read_csv(blasted_path, sep="\t")
        track_data.set_index('ASV', inplace=True)
        blasted_data.set_index('ASV', inplace=True)
        merged_data = pd.merge(track_data, blasted_data, left_index=True, right_index=True, how='outer')
        merged_data = merged_data.iloc[merged_data.index.map(extract_number).argsort()]
        merged_data.to_csv(final_report_route+samplename+"_report.csv", index=True, index_label='ASV')

    except:
        print(samplename+"error")

# filter duplicate data
for samplename in samplenames:
    origin_report_file = pd.read_csv(final_report_route + samplename + "_report.csv", header=0)
    drop_duplicate_data = origin_report_file.drop_duplicates(subset=['ASV', 'Scientific_name'], keep='first').copy()
    drop_duplicate_data.loc[:, 'Identity'] = drop_duplicate_data['Identity'].astype(float)
    max_identities = drop_duplicate_data.groupby(['ASV'])['Identity'].transform(max)
    filtered_blasted_data = drop_duplicate_data[drop_duplicate_data['Identity'] == max_identities]
    filtered_blasted_data.loc[filtered_blasted_data.duplicated(subset=['ASV'], keep='first'), ['Value', 'Ratio']] = ''
    filtered_blasted_data.to_csv(final_report_route + samplename + "_report.csv", index=False)

# add chinese name
for samplename in samplenames:
    try:
        ref_path = ref_path
        report_path = final_report_route+samplename+"_report.csv"
        ref_data = pd.read_excel(ref_path)
        report_data = pd.read_csv(report_path)
        ref_data.set_index('Scientific_name', inplace=True)
        report_data.set_index('Scientific_name', inplace=True)
        merged_data = pd.merge(report_data, ref_data, left_index=True, right_index=True, how='left')
        merged_data.insert(5, 'Scientific_name', merged_data.index)
        merged_data = merged_data.sort_values(by='ASV', key=lambda x: x.map(extract_number))
        merged_data.to_excel(chinese_report_route+samplename+"_chreport.xlsx",index=False)
    
    except:
        print(samplename+"error")




# cheers
from turtle import *
color("red","yellow")
begin_fill()
while True:
    forward(100)
    left(165)
    if abs(pos())<1:
        break
end_fill()
done()
