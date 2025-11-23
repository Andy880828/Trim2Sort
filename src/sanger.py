"""Sanger 定序分析模組

Sanger sequencing analysis module.
"""

from collections.abc import Sequence
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import filedialog

from Bio import SeqIO
import customtkinter
import pandas as pd
from PIL import Image

from src.utils.logger_utils import get_logger
from src.utils.path_utils import (
    find_latest_ref_file,
    get_blastn_path,
    get_icon_path,
    get_primer_path,
    get_project_root,
    get_trimmomatic_path,
)
from src.utils.subprocess_utils import run_command
from src.utils.ui_config import COLORS, FONTS, LAYOUT

logger = get_logger(__name__)


def load_app_image() -> customtkinter.CTkImage:
    """載入應用程式圖示

    Load application icon.

    Returns:
        customtkinter.CTkImage: 應用程式圖示 / Application icon.
    """
    icon_path = get_icon_path()
    return customtkinter.CTkImage(light_image=Image.open(icon_path), size=(128, 72))


class Sanger(customtkinter.CTkToplevel):
    """Sanger 分析主視窗

    Sanger analysis main window.
    """

    def __init__(self, *args, **kwargs) -> None:
        """初始化 Sanger 分析視窗

        Initialize Sanger analysis window.

        Args:
            *args: 位置參數 / Positional arguments.
            **kwargs: 關鍵字參數 / Keyword arguments.
        """
        super().__init__(*args, **kwargs)
        self.title("Sanger Analysis")
        self.geometry("500x400")
        self.configure(fg_color=COLORS.PRIMARY_BG)
        self.resizable(False, False)
        self.grid_rowconfigure(0, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self.frame = SangerHeaderFrame(master=self)
        self.frame.configure(fg_color=COLORS.PRIMARY_BG)
        self.frame.grid(row=1, column=0, padx=10, pady=(0, 15))

        self.frame_2 = SangerContentFrame(master=self)
        self.frame_2.configure(fg_color=COLORS.PRIMARY_BG, height=470)
        self.frame_2.grid(row=2, column=0, padx=20, pady=0, sticky="nsew")


class SangerHeaderFrame(customtkinter.CTkFrame):
    """Sanger 標題框架

    Sanger header frame.
    """

    def __init__(self, master: customtkinter.CTk | customtkinter.CTkFrame, **kwargs) -> None:
        """初始化標題框架

        Initialize header frame.

        Args:
            master: 父視窗 / Parent window.
            **kwargs: 關鍵字參數 / Keyword arguments.
        """
        super().__init__(master, **kwargs)

        my_image = load_app_image()

        logo = customtkinter.CTkLabel(self, image=my_image, text="")
        logo.grid(row=0, column=2, padx=10)

        buttons_config = [
            ("USER GUIDE", 0),
            ("DOCUMENTATION", 2),
            ("CONTACT US", 4),
        ]

        for text, col in buttons_config:
            button = customtkinter.CTkButton(
                self,
                width=100,
                height=LAYOUT.BUTTON_HEIGHT,
                border_width=0,
                text=text,
                fg_color="transparent",
                text_color=COLORS.TEXT_SECONDARY,
                hover_color=COLORS.SECONDARY_BG,
                font=(FONTS.FAMILY, FONTS.SIZE_NORMAL, FONTS.STYLE_BOLD),
            )
            button.grid(row=1, column=col, padx=10)

            if col < 4:
                sep_label = customtkinter.CTkLabel(
                    self,
                    text="✦",
                    text_color=COLORS.TEXT_SECONDARY,
                    font=(FONTS.FAMILY, 10, FONTS.STYLE_BOLD),
                )
                sep_label.grid(row=1, column=col + 1, padx=10)


class SangerContentFrame(customtkinter.CTkFrame):
    """Sanger 內容框架

    Sanger content frame.
    """

    def __init__(self, master: customtkinter.CTk | customtkinter.CTkFrame, **kwargs) -> None:
        """初始化內容框架

        Initialize content frame.

        Args:
            master: 父視窗 / Parent window.
            **kwargs: 關鍵字參數 / Keyword arguments.
        """
        super().__init__(master, **kwargs)

        self.samples_path = tk.StringVar()
        self.outputs_path = tk.StringVar()

        self._setup_ui()

    def _setup_ui(self) -> None:
        """設定使用者介面

        Setup user interface.
        """
        database_info = tk.StringVar(value="Step1: Select the folder directory of Database")
        samples_info = tk.StringVar(value="Step2: Select the folder directory of Samples")
        outputs_info = tk.StringVar(value="Step3: Select the folder directory of Outputs")
        instruction = tk.StringVar(value="Start the app after everthing above is selected")

        database_info_label = customtkinter.CTkLabel(
            self,
            textvariable=database_info,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        database_info_label.grid(row=0, column=0, padx=0, sticky="w")

        self.database_combobox = customtkinter.CTkOptionMenu(
            self,
            values=["Teleostei-12S", "Teleostei-COI"],
            width=370,
            height=LAYOUT.ENTRY_HEIGHT,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            button_color=COLORS.SECONDARY_BG,
            text_color=COLORS.TEXT_PRIMARY,
            dropdown_fg_color=COLORS.DROPDOWN_BG,
            dropdown_text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            dropdown_font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        self.database_combobox.set("Database")
        self.database_combobox.grid(row=1, column=0, padx=0, pady=(0, 10), sticky="w")

        samples_info_label = customtkinter.CTkLabel(
            self,
            textvariable=samples_info,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        samples_info_label.grid(row=2, column=0, padx=0, sticky="w")

        samples_entry = customtkinter.CTkEntry(
            self,
            textvariable=self.samples_path,
            text_color=COLORS.TEXT_PRIMARY,
            width=370,
            height=LAYOUT.ENTRY_HEIGHT,
            border_width=LAYOUT.BORDER_WIDTH,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        samples_entry.grid(row=3, column=0, padx=0, pady=(0, 10), sticky="w")

        samples_button = customtkinter.CTkButton(
            self,
            width=80,
            height=LAYOUT.BUTTON_HEIGHT,
            border_width=0,
            corner_radius=LAYOUT.CORNER_RADIUS,
            text="SELECT",
            fg_color=COLORS.SECONDARY_BG,
            text_color=COLORS.TEXT_PRIMARY,
            hover_color=COLORS.HOVER_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_NORMAL, FONTS.STYLE_BOLD),
            command=self.browse_samples,
        )
        samples_button.grid(row=3, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

        outputs_info_label = customtkinter.CTkLabel(
            self,
            textvariable=outputs_info,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        outputs_info_label.grid(row=4, column=0, padx=0, sticky="w")

        outputs_entry = customtkinter.CTkEntry(
            self,
            textvariable=self.outputs_path,
            text_color=COLORS.TEXT_PRIMARY,
            width=370,
            height=LAYOUT.ENTRY_HEIGHT,
            border_width=LAYOUT.BORDER_WIDTH,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        outputs_entry.grid(row=5, column=0, padx=0, pady=(0, 10), sticky="w")

        outputs_button = customtkinter.CTkButton(
            self,
            width=80,
            height=LAYOUT.BUTTON_HEIGHT,
            border_width=0,
            corner_radius=LAYOUT.CORNER_RADIUS,
            text="SELECT",
            fg_color=COLORS.SECONDARY_BG,
            text_color=COLORS.TEXT_PRIMARY,
            hover_color=COLORS.HOVER_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_NORMAL, FONTS.STYLE_BOLD),
            command=self.browse_outputs,
        )
        outputs_button.grid(row=5, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

        instruction_label = customtkinter.CTkLabel(
            self,
            textvariable=instruction,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        instruction_label.grid(row=6, column=0, columnspan=2, padx=0, pady=(10, 0), sticky="ew")

        self.analyse_button = customtkinter.CTkButton(
            self,
            width=80,
            height=LAYOUT.BUTTON_HEIGHT,
            border_width=0,
            corner_radius=LAYOUT.CORNER_RADIUS,
            text="ANALYSE",
            fg_color=COLORS.SECONDARY_BG,
            text_color=COLORS.TEXT_PRIMARY,
            hover_color=COLORS.HOVER_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_NORMAL, FONTS.STYLE_BOLD),
            command=self.analysis,
            state="disabled",
        )
        self.analyse_button.grid(row=7, column=0, columnspan=2, padx=0, sticky="ew")

        self._setup_field_validation()

    def _setup_field_validation(self) -> None:
        """設定欄位驗證

        Setup field validation.
        """
        self.database_combobox.configure(command=lambda _choice: self._check_fields())
        self.samples_path.trace_add("write", lambda *_args: self._check_fields())
        self.outputs_path.trace_add("write", lambda *_args: self._check_fields())
        self._check_fields()

    def _check_fields(self) -> None:
        """檢查所有必要欄位是否已填寫

        Check if all required fields are filled.
        """
        database_choice = self.database_combobox.get()
        all_filled = (
            database_choice != "Database"
            and bool(self.samples_path.get())
            and bool(self.outputs_path.get())
        )

        if all_filled:
            self.analyse_button.configure(state="normal")
        else:
            self.analyse_button.configure(state="disabled")

    def browse_samples(self) -> None:
        """瀏覽選擇樣本資料夾

        Browse to select samples folder.
        """
        foldername = filedialog.askdirectory()
        if foldername:
            self.samples_path.set(foldername)
        self._check_fields()

    def browse_outputs(self) -> None:
        """瀏覽選擇輸出資料夾

        Browse to select outputs folder.
        """
        foldername = filedialog.askdirectory()
        if foldername:
            self.outputs_path.set(foldername)
        self._check_fields()

    def analysis(self) -> None:
        """執行 Sanger 分析

        Execute Sanger analysis.
        """
        input_folder = Path(self.samples_path.get())
        output_folder = Path(self.outputs_path.get())

        if not input_folder.exists() or not output_folder.exists():
            logger.error("輸入或輸出資料夾不存在")
            return

        sample_files = sorted(
            input_folder.glob("*.ab1"), key=lambda x: x.stem.split("_Primer-Added")[0]
        )
        sample_size = len(sample_files)

        if sample_size == 0:
            logger.error("未找到任何 .ab1 檔案")
            return

        now = datetime.now()
        output_name = output_folder / f"{now.month}.{now.day}_N={sample_size}"

        merged_fastq = Path(f"{output_name}merged.fastq")
        trimmed_fastq = Path(f"{output_name}_trimmed.fastq")
        trimmed_fasta = Path(f"{output_name}_trimmed.fasta")
        blasted_txt = Path(f"{output_name}_blasted.txt")
        trimlog = Path(f"{output_name}_trimlog.log")

        primer = self._get_primer_file()
        if not primer:
            logger.error("未選擇資料庫")
            return

        self._merge_ab1_to_fastq(sample_files, merged_fastq)
        self._trim_sequences(merged_fastq, trimmed_fastq, trimlog, primer)
        self._convert_fastq_to_fasta(trimmed_fastq, trimmed_fasta)
        self._run_blast(trimmed_fasta, blasted_txt)
        self._process_blast_results(blasted_txt, sample_files, output_name)

        logger.info("分析完成")

    def _get_primer_file(self) -> str | None:
        """取得 Primer 檔案名稱

        Get primer file name.

        Returns:
            str | None: Primer 檔案名稱或 None / Primer file name or None.
        """
        database_choice = self.database_combobox.get()

        match database_choice:
            case "Teleostei-12S":
                return "MiFish.fa"
            case "Teleostei-COI":
                return "COIF1R1.fa"
            case _:
                return None

    def _merge_ab1_to_fastq(self, ab1_files: Sequence[Path], output_fastq: Path) -> None:
        """合併所有 AB1 檔案為單一 FASTQ 檔案

        Merge all AB1 files into a single FASTQ file.

        Args:
            ab1_files (Sequence[Path]): AB1 檔案列表 / List of AB1 files.
            output_fastq (Path): 輸出 FASTQ 檔案路徑 / Output FASTQ file path.
        """
        with output_fastq.open("w") as output_handle:
            for ab1_file in ab1_files:
                records = SeqIO.parse(ab1_file, "abi")
                filename = ab1_file.stem
                for record in records:
                    record.id = f"{filename} {record.id}"
                    SeqIO.write(record, output_handle, "fastq")

    def _trim_sequences(
        self, input_fastq: Path, output_fastq: Path, trimlog: Path, primer: str
    ) -> None:
        """使用 Trimmomatic 修剪序列

        Trim sequences using Trimmomatic.

        Args:
            input_fastq (Path): 輸入 FASTQ 檔案路徑 / Input FASTQ file path.
            output_fastq (Path): 輸出 FASTQ 檔案路徑 / Output FASTQ file path.
            trimlog (Path): Trim log 檔案路徑 / Trim log file path.
            primer (str): Primer 檔案名稱 / Primer file name.
        """
        trimmomatic_path = get_trimmomatic_path()
        primer_path = get_primer_path(primer)

        trimming_cmd = [
            "java",
            "-jar",
            str(trimmomatic_path),
            "SE",
            "-phred33",
            "-trimlog",
            str(trimlog),
            str(input_fastq),
            str(output_fastq),
            f"ILLUMINACLIP:{primer_path}:3:30:10",
            "LEADING:30",
            "TRAILING:30",
        ]
        run_command(trimming_cmd)

    def _convert_fastq_to_fasta(self, input_fastq: Path, output_fasta: Path) -> None:
        """將 FASTQ 轉換為 FASTA 格式

        Convert FASTQ to FASTA format.

        Args:
            input_fastq (Path): 輸入 FASTQ 檔案路徑 / Input FASTQ file path.
            output_fasta (Path): 輸出 FASTA 檔案路徑 / Output FASTA file path.
        """
        with output_fasta.open("w") as output_handle:
            for record in SeqIO.parse(input_fastq, "fastq"):
                SeqIO.write(record, output_handle, "fasta")

    def _run_blast(self, query_fasta: Path, output_txt: Path) -> None:
        """執行 BLASTN 比對

        Run BLASTN alignment.

        Args:
            query_fasta (Path): 查詢 FASTA 檔案路徑 / Query FASTA file path.
            output_txt (Path): 輸出文字檔案路徑 / Output text file path.
        """
        blastn_exe = get_blastn_path()
        database_choice = self.database_combobox.get()
        root_dir = get_project_root()

        match database_choice:
            case "Teleostei-12S":
                db_dir = root_dir / "db-12S"
                db_name = "12S_Osteichthyes_db"
            case "Teleostei-COI":
                db_dir = root_dir / "db-CO1"
                db_name = "CO1_Osteichthyes_db"
            case _:
                logger.error("未知的資料庫類型")
                return

        blast_cmd = [
            str(blastn_exe),
            "-query",
            str(query_fasta),
            "-out",
            str(output_txt),
            "-db",
            str(db_dir / db_name),
            "-outfmt",
            "6 qseqid pident qcovs sscinames sacc qlen",
            "-max_target_seqs",
            "1",
        ]
        run_command(blast_cmd, cwd=db_dir)

    def _process_blast_results(
        self, blast_txt: Path, sample_files: Sequence[Path], output_name: Path
    ) -> None:
        """處理 BLAST 結果並產生 Excel 檔案

        Process BLAST results and generate Excel files.

        Args:
            blast_txt (Path): BLAST 結果文字檔案路徑 / BLAST result text file path.
            sample_files (Sequence[Path]): 樣本檔案列表 / List of sample files.
            output_name (Path): 輸出檔案名稱前綴 / Output file name prefix.
        """
        lines = blast_txt.read_text().splitlines()
        processed_lines = []

        for line in lines:
            parts = line.split("\t")
            parts[0] = parts[0].replace("_Primer-Added", "")
            processed_lines.append("\t".join(parts))

        blast_txt.write_text("\n".join(processed_lines))

        blast_set = set()
        for line in processed_lines:
            blast_set.add(line.split()[0])

        in_name_set = {sample_file.stem.split("_Primer-Added")[0] for sample_file in sample_files}
        fail_lists = in_name_set.difference(blast_set)

        dataset = pd.DataFrame(list(fail_lists), columns=["No"])
        blastresult = pd.read_csv(
            blast_txt,
            header=None,
            sep="\t",
            names=["No", "Identity", "Coverage", "Scientific_name", "Accession_number", "bp"],
        )
        concat_result = pd.concat([dataset, blastresult], ignore_index=True)

        result_file = Path(f"{output_name}_result.xlsx")
        concat_result.to_excel(result_file, engine="openpyxl", index=False)

        parts = result_file.stem.split("_")
        new_file_name = result_file.parent / f"{parts[0]}_{parts[1]}.xlsx"
        result_file.rename(new_file_name)

        self._add_chinese_names(new_file_name)

    def _add_chinese_names(self, input_file: Path) -> None:
        """添加中文名稱到結果檔案

        Add Chinese names to result file.

        Args:
            input_file (Path): 輸入 Excel 檔案路徑 / Input Excel file path.
        """
        ref_path = find_latest_ref_file()
        if not ref_path:
            logger.warning("未找到參考檔案, 無法添加中文名稱")
            return

        try:
            df_ref = pd.read_excel(ref_path, engine="openpyxl")
            df_ref.set_index("Scientific_name", inplace=True)

            df_input = pd.read_excel(input_file, engine="openpyxl")
            df_input.set_index("Scientific_name", inplace=True)

            df_ref_merged = pd.merge(
                df_input, df_ref, left_index=True, right_index=True, how="left"
            )
            df_ref_merged.insert(3, "Scientific_name", df_ref_merged.index)
            df_ref_merged = df_ref_merged.sort_values(by="No")

            output_file_path = Path(str(input_file).replace(".xlsx", "_zh.xlsx"))
            df_ref_merged.to_excel(output_file_path, engine="openpyxl", index=False)

            logger.info(f"已添加中文名稱: {output_file_path}")
        except Exception as e:
            logger.error(f"添加中文名稱失敗: {e}")
