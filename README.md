# Trim2Sort

一個用於 DNA 序列分析的圖形化使用者介面（GUI）應用程式，支援 Sanger 定序與次世代定序（NGS）資料處理與分析。

## 功能特色

- **雙模式支援**：同時支援 Sanger 定序與 NGS 資料分析
- **自動化流程**：整合多個生物資訊工具，提供一鍵式分析流程
- **圖形化介面**：使用 CustomTkinter 建立現代化的使用者介面
- **結果視覺化**：自動標記重複序列與過濾序列，並以 Excel 格式輸出結果

## 系統需求

- **作業系統**：Windows
- **Python 版本**：>= 3.12
- **必要工具**：
  - `cutadapt.exe`（已包含在專案中）
  - `usearch.exe`（已包含在專案中）
  - `blastn.exe`（位於 `blast+/bin/` 目錄中）

## 安裝步驟

1. **複製專案**

   ```bash
   git clone https://github.com/Andy880828/Trim2Sort.git
   cd Trim2Sort
   ```

2. **安裝 Python 依賴套件**

   ```bash
   uv sync
   ```

3. **確認必要工具位置**

   - `cutadapt.exe` 應位於專案根目錄
   - `usearch.exe` 應位於專案根目錄
   - `blastn.exe` 位於 `blast+/bin/` 目錄中

   視情況可更新，
   cutadapt 後期版本就不提供 .exe ，目前的 4.0 應為最後版本，故不用更新。
   usearch 可至官網 https://www.drive5.com/usearch/ 下載新版本後丟進根目錄請改名為 usearch.exe。
   blastn 為 NCBI blast+ 軟體的一部分，可至 https://ftp.ncbi.nlm.nih.gov/blast/executables/blast+/LATEST/ 下載 windows 可安裝檔，安裝後將整個資料夾(名稱如 blast-2.17.0+ )拉入根目錄後改名為 blast+即可。

## 使用方法

### 啟動程式

執行 `execute(開啟請按我).bat` 或直接執行：

```bash
uv run main.py
```

### NGS 分析流程

1. **選擇分析模式**：在主畫面選擇「NGS」

2. **設定工具路徑**（依序完成）：

   - **Step 1**：選擇 `cutadapt.exe`（若位於根目錄會自動偵測）
   - **Step 2**：選擇 `usearch.exe`（若位於根目錄會自動偵測）
   - **Step 3**：選擇 `blastn.exe` 所在目錄（若位於 `blast+/bin/`會自動偵測）
   - **Step 4**：選擇資料庫目錄（例如：`db-CO1/` 或 `db-12S/`）
   - **Step 5**：選擇樣本資料夾（包含配對端 FASTQ 檔案）
   - **Step 6**：選擇輸出資料夾

3. **開始分析**：點擊「ANALYSE」按鈕

### NGS 分析流程說明

程式會自動執行以下步驟：

1. **引子修剪**（A_primer_trimming）：使用 cutadapt 移除引子序列
2. **配對端合併**（B_merged）：合併正向與反向讀序
3. **品質控制**（C_quality）：過濾低品質序列（品質分數 < 20）
4. **長度過濾**（D_length）：保留長度 ≥ 150 bp 的序列
5. **序列去重**（E_uniques）：識別並合併重複序列
6. **OTU 生成**（F_OTUs）：建立操作分類單元（OTU）
7. **OTU 表格**（G_OTUtable）：生成 OTU 豐度表格
8. **BLAST 比對**（H_blasts）：與資料庫進行序列比對
9. **結果整理**（I_sorted_blasts）：合併比對結果與 OTU 表格，輸出 Excel 檔案

### 輸出結果

分析完成後，會在輸出資料夾的 `I_sorted_blasts` 目錄中生成以下檔案：

- `[樣本名稱].xlsx`：包含比對結果、讀序數與比例
- `[樣本名稱]_zh_added.xlsx`：額外包含中文物種名稱

**結果標記說明**：

- `*DUPE*`：標記為重複序列（藍色背景）
- `*FILTERED*`：標記為過濾序列（綠色背景，Identity < 97%）

## 專案結構

```
Trim2Sort/
├── main.py                 # 主程式入口
├── ngs.py                  # NGS 分析模組
├── sanger.py               # Sanger 分析模組
├── zh_adder.py             # 中文名稱添加工具
├── cutadapt.exe            # 引子修剪工具
├── usearch.exe             # 序列分析工具
├── blast+/                 # BLAST+ 工具套件
│   └── bin/
│       └── blastn.exe
├── ref_ver6_0918.xlsx      # 中文物種名稱參考表
├── samples/                # 樣本資料夾（使用者提供）
├── outputs/                # 輸出資料夾（分析結果）
└── README.md               # 本檔案
```

## 依賴套件

- `customtkinter >= 5.2.2`：現代化 GUI 框架
- `pandas >= 2.3.3`：資料處理
- `numpy >= 2.3.4`：數值運算
- `openpyxl >= 3.1.5`：Excel 檔案處理
- `biopython >= 1.86`：生物資訊工具
- `natsort >= 8.4.0`：自然排序
- `pillow >= 12.0.0`：影像處理
- `jinja2 >= 3.1.6`：模板引擎

## 注意事項

1. **檔案格式**：NGS 分析需要配對端 FASTQ 檔案（R1 與 R2）
2. **檔案命名**：樣本檔案應成對命名（例如：`sample1_R1.fastq` 與 `sample1_R2.fastq`）
3. **資料庫設定**：使用前請確認 BLAST 資料庫已正確建立
4. **路徑設定**：建議使用絕對路徑以避免路徑相關問題

## 版本歷史

**2025/11 以前的更新紀錄可參考"更新日誌(read me).txt"**

### 主要更新（v3.0.0）

- 修復 Dupe 及 Filter 的 reads 與 ratio 顯示問題
- 更新 COI_Osteichthyes 及 12S_Osteichthyes 序列資料庫
- 改進 Sanger 分析中的 Trimmomatic 引子修剪功能
- 優化結果輸出格式與視覺化

## 未來規劃

- [ ] 修正中文資料庫中的 valid name 錯誤
- [ ] 改進 Sanger 分析的排序方式
- [ ] 優化 Sanger 分析執行時間
- [ ] 在 NGS 分析中加入 DADA2 去噪方法

**Produced by 高屏溪男子偶像團體**
