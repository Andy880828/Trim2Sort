"""NGS(次世代定序)分析模組

NGS (Next Generation Sequencing) analysis module.
"""

from collections.abc import Sequence
import gc
from pathlib import Path
import re
import shutil
import tkinter as tk
from tkinter import filedialog, messagebox
import warnings

import customtkinter
import natsort
import numpy as np
import pandas as pd
from PIL import Image

from src.utils.excel_utils import highlight_row
from src.utils.logger_utils import get_logger
from src.utils.path_utils import (
    find_latest_ref_file,
    get_blastn_path,
    get_cutadapt_path,
    get_icon_path,
    get_usearch_path,
)
from src.utils.sequence_utils import reverse_complement
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


class NGSConfig:
    """NGS 分析設定

    NGS analysis configuration.
    """

    def __init__(self) -> None:
        """初始化 NGS 設定

        Initialize NGS configuration.
        """
        self.primer_mode: str = "MiFish-U"
        self.forward_primer: str = "GTCGGTAAAACTCGTGCCAGC"
        self.reverse_primer: str = "CAAACTGGGATTAGATACCCCACTATG"
        self.quality_threshold: int = 20
        self.length_threshold: int = 150
        self.ref_path: str = find_latest_ref_file()


class ConfigWindow(customtkinter.CTkToplevel):
    """設定視窗

    Configuration window.
    """

    def __init__(self, parent: customtkinter.CTkFrame, *args, **kwargs) -> None:
        """初始化設定視窗

        Initialize configuration window.

        Args:
            parent: 父視窗 / Parent window.
            *args: 位置參數 / Positional arguments.
            **kwargs: 關鍵字參數 / Keyword arguments.
        """
        super().__init__(parent, *args, **kwargs)
        self.parent = parent
        self.title("Configuration")
        self.geometry("480x480")
        self.configure(fg_color=COLORS.PRIMARY_BG)
        self.resizable(False, False)

        self._setup_ui()

    def _setup_ui(self) -> None:
        """設定使用者介面

        Setup user interface.
        """
        primer_label = customtkinter.CTkLabel(
            self,
            text="Primers:",
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        primer_label.grid(row=0, column=0, padx=20, pady=(15, 3), sticky="w")

        self.primer_mode = tk.StringVar(value=self.parent.config.primer_mode)
        primer_combo = customtkinter.CTkComboBox(
            self,
            variable=self.primer_mode,
            values=["MiFish-U", "Manual"],
            width=370,
            height=LAYOUT.ENTRY_HEIGHT,
            border_width=LAYOUT.BORDER_WIDTH,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            text_color=COLORS.TEXT_PRIMARY,
            button_color=COLORS.HOVER_BG,
            button_hover_color=COLORS.BUTTON_HOVER,
            dropdown_fg_color=COLORS.SECONDARY_BG,
            dropdown_text_color=COLORS.TEXT_PRIMARY,
            dropdown_hover_color=COLORS.HOVER_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            command=self.on_primer_mode_change,
        )
        primer_combo.grid(row=1, column=0, columnspan=2, padx=20, pady=(0, 5), sticky="w")

        forward_label = customtkinter.CTkLabel(
            self,
            text="Forward Primer:",
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        forward_label.grid(row=2, column=0, padx=20, pady=(8, 3), sticky="w")

        self.forward_primer = tk.StringVar(value=self.parent.config.forward_primer)
        self.forward_entry = customtkinter.CTkEntry(
            self,
            textvariable=self.forward_primer,
            text_color=COLORS.TEXT_PRIMARY,
            width=440,
            height=LAYOUT.ENTRY_HEIGHT,
            border_width=LAYOUT.BORDER_WIDTH,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        self.forward_entry.grid(row=3, column=0, columnspan=2, padx=20, pady=(0, 5), sticky="w")

        reverse_label = customtkinter.CTkLabel(
            self,
            text="Reverse Primer:",
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        reverse_label.grid(row=4, column=0, padx=20, pady=(8, 3), sticky="w")

        self.reverse_primer = tk.StringVar(value=self.parent.config.reverse_primer)
        self.reverse_entry = customtkinter.CTkEntry(
            self,
            textvariable=self.reverse_primer,
            text_color=COLORS.TEXT_PRIMARY,
            width=440,
            height=LAYOUT.ENTRY_HEIGHT,
            border_width=LAYOUT.BORDER_WIDTH,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        self.reverse_entry.grid(row=5, column=0, columnspan=2, padx=20, pady=(0, 5), sticky="w")

        quality_label = customtkinter.CTkLabel(
            self,
            text="Quality Threshold:",
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        quality_label.grid(row=6, column=0, padx=20, pady=(8, 3), sticky="w")

        self.quality_threshold = tk.StringVar(value=str(self.parent.config.quality_threshold))
        quality_entry = customtkinter.CTkEntry(
            self,
            textvariable=self.quality_threshold,
            text_color=COLORS.TEXT_PRIMARY,
            width=440,
            height=LAYOUT.ENTRY_HEIGHT,
            border_width=LAYOUT.BORDER_WIDTH,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        quality_entry.grid(row=7, column=0, columnspan=2, padx=20, pady=(0, 5), sticky="w")

        length_label = customtkinter.CTkLabel(
            self,
            text="Length Threshold:",
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        length_label.grid(row=8, column=0, padx=20, pady=(8, 3), sticky="w")

        self.length_threshold = tk.StringVar(value=str(self.parent.config.length_threshold))
        length_entry = customtkinter.CTkEntry(
            self,
            textvariable=self.length_threshold,
            text_color=COLORS.TEXT_PRIMARY,
            width=440,
            height=LAYOUT.ENTRY_HEIGHT,
            border_width=LAYOUT.BORDER_WIDTH,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        length_entry.grid(row=9, column=0, columnspan=2, padx=20, pady=(0, 5), sticky="w")

        ref_label = customtkinter.CTkLabel(
            self,
            text="Ref Path (Auto-detected):",
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        ref_label.grid(row=10, column=0, padx=20, pady=(8, 3), sticky="w")

        self.ref_path = tk.StringVar(value=self.parent.config.ref_path)
        self.ref_entry = customtkinter.CTkEntry(
            self,
            textvariable=self.ref_path,
            text_color=COLORS.TEXT_PRIMARY,
            width=350,
            height=LAYOUT.ENTRY_HEIGHT,
            border_width=LAYOUT.BORDER_WIDTH,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        self.ref_entry.grid(row=11, column=0, padx=20, pady=(0, 5), sticky="w")

        ref_button = customtkinter.CTkButton(
            self,
            width=70,
            height=LAYOUT.BUTTON_HEIGHT,
            border_width=0,
            corner_radius=LAYOUT.CORNER_RADIUS,
            text="SELECT",
            fg_color=COLORS.SECONDARY_BG,
            text_color=COLORS.TEXT_PRIMARY,
            hover_color=COLORS.HOVER_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_NORMAL, FONTS.STYLE_BOLD),
            command=self.browse_ref,
        )
        ref_button.grid(row=11, column=1, padx=(0, 20), pady=(0, 5), sticky="w")

        button_frame = customtkinter.CTkFrame(self, fg_color="transparent")
        button_frame.grid(row=12, column=0, columnspan=2, padx=20, pady=(15, 15), sticky="ew")

        ok_button = customtkinter.CTkButton(
            button_frame,
            width=220,
            height=30,
            border_width=0,
            corner_radius=LAYOUT.CORNER_RADIUS,
            text="OK",
            fg_color=COLORS.SECONDARY_BG,
            text_color=COLORS.TEXT_PRIMARY,
            hover_color=COLORS.HOVER_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            command=self.on_ok,
        )
        ok_button.pack(side="left", padx=(0, 10), expand=True, fill="x")

        cancel_button = customtkinter.CTkButton(
            button_frame,
            width=220,
            height=30,
            border_width=0,
            corner_radius=LAYOUT.CORNER_RADIUS,
            text="CANCEL",
            fg_color=COLORS.SECONDARY_BG,
            text_color=COLORS.TEXT_PRIMARY,
            hover_color=COLORS.HOVER_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            command=self.on_cancel,
        )
        cancel_button.pack(side="left", expand=True, fill="x")

        self.on_primer_mode_change(self.primer_mode.get())

    def on_primer_mode_change(self, choice: str) -> None:
        """當 Primer 模式改變時調整欄位狀態

        Adjust field states when primer mode changes.

        Args:
            choice (str): 選擇的模式 / Selected mode.
        """
        match choice:
            case "MiFish-U":
                self.forward_primer.set("GTCGGTAAAACTCGTGCCAGC")
                self.reverse_primer.set("CAAACTGGGATTAGATACCCCACTATG")
                self.forward_entry.configure(state="disabled")
                self.reverse_entry.configure(state="disabled")
            case "Manual":
                self.forward_entry.configure(state="normal")
                self.reverse_entry.configure(state="normal")

    def browse_ref(self) -> None:
        """瀏覽選擇 Ref 檔案

        Browse to select ref file.
        """
        filename = filedialog.askopenfilename(
            title="Select Reference File",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if filename:
            self.ref_path.set(filename)

    def on_ok(self) -> None:
        """儲存配置並關閉視窗

        Save configuration and close window.
        """
        self.parent.config.primer_mode = self.primer_mode.get()
        self.parent.config.forward_primer = self.forward_primer.get()
        self.parent.config.reverse_primer = self.reverse_primer.get()

        try:
            self.parent.config.quality_threshold = int(self.quality_threshold.get())
        except ValueError:
            self.parent.config.quality_threshold = 20
            logger.warning("品質閾值設定無效, 使用預設值 20")

        try:
            self.parent.config.length_threshold = int(self.length_threshold.get())
        except ValueError:
            self.parent.config.length_threshold = 150
            logger.warning("長度閾值設定無效, 使用預設值 150")

        self.parent.config.ref_path = self.ref_path.get()
        self.destroy()

    def on_cancel(self) -> None:
        """取消並關閉視窗

        Cancel and close window.
        """
        self.destroy()


class NGS(customtkinter.CTkToplevel):
    """NGS 分析主視窗

    NGS analysis main window.
    """

    def __init__(self, *args, **kwargs) -> None:
        """初始化 NGS 分析視窗

        Initialize NGS analysis window.

        Args:
            *args: 位置參數 / Positional arguments.
            **kwargs: 關鍵字參數 / Keyword arguments.
        """
        super().__init__(*args, **kwargs)
        self.title("NGS Analysis")
        self.geometry("500x650")
        self.configure(fg_color=COLORS.PRIMARY_BG)
        self.resizable(False, False)
        self.grid_rowconfigure(0, weight=0)
        self.grid_columnconfigure(0, weight=1)

        self.frame = NGSHeaderFrame(master=self)
        self.frame.configure(fg_color=COLORS.PRIMARY_BG)
        self.frame.grid(row=1, column=0, padx=10, pady=(0, 15))

        self.frame_2 = NGSContentFrame(master=self)
        self.frame_2.configure(fg_color=COLORS.PRIMARY_BG, height=520)
        self.frame_2.grid(row=2, column=0, padx=20, pady=0, sticky="nsew")


class NGSHeaderFrame(customtkinter.CTkFrame):
    """NGS 標題框架

    NGS header frame.
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


class NGSContentFrame(customtkinter.CTkFrame):
    """NGS 內容框架

    NGS content frame.
    """

    def __init__(self, master: customtkinter.CTk | customtkinter.CTkFrame, **kwargs) -> None:
        """初始化內容框架

        Initialize content frame.

        Args:
            master: 父視窗 / Parent window.
            **kwargs: 關鍵字參數 / Keyword arguments.
        """
        super().__init__(master, **kwargs)

        self.cutadapt_path = tk.StringVar()
        self.usearch_path = tk.StringVar()
        self.blastn_path = tk.StringVar()
        self.database_path = tk.StringVar()
        self.database_selector = tk.StringVar()
        self.samples_path = tk.StringVar()
        self.outputs_path = tk.StringVar()

        self.config = NGSConfig()

        self._setup_ui()
        self._auto_detect_tools()

    def _auto_detect_tools(self) -> None:
        """自動偵測工具路徑

        Auto-detect tool paths.
        """
        cutadapt_exe = get_cutadapt_path()
        usearch_exe = get_usearch_path()
        blastn_exe = get_blastn_path()

        if cutadapt_exe.exists():
            self.cutadapt_path.set(str(cutadapt_exe))
        if usearch_exe.exists():
            self.usearch_path.set(str(usearch_exe))
        if blastn_exe.exists():
            self.blastn_path.set(str(blastn_exe))
        self._check_fields()

    def _setup_ui(self) -> None:
        """設定使用者介面

        Setup user interface.
        """
        cutadapt_info = tk.StringVar(
            value="Step1: Cutadapt.exe (Auto-detected from root directory)"
        )
        usearch_info = tk.StringVar(value="Step2: Usearch.exe (Auto-detected from root directory)")
        blastn_info = tk.StringVar(value="Step3: Blastn.exe (Auto-detected from root directory)")
        database_info = tk.StringVar(value="Step4: Select the folder directory of Database")
        samples_info = tk.StringVar(value="Step5: Select the folder directory of Samples")
        outputs_info = tk.StringVar(value="Step6: Select the folder directory of Outputs")
        instruction = tk.StringVar(value="Start the app after everthing above is selected")

        row = 0
        entries = [
            (cutadapt_info, self.cutadapt_path, self.browse_cutadapt),
            (usearch_info, self.usearch_path, self.browse_usearch),
            (blastn_info, self.blastn_path, self.browse_blastn),
        ]

        for info_var, path_var, command in entries:
            info_label = customtkinter.CTkLabel(
                self,
                textvariable=info_var,
                text_color=COLORS.TEXT_PRIMARY,
                font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            )
            info_label.grid(row=row, column=0, padx=0, sticky="w")

            entry = customtkinter.CTkEntry(
                self,
                textvariable=path_var,
                text_color=COLORS.TEXT_PRIMARY,
                width=370,
                height=LAYOUT.ENTRY_HEIGHT,
                border_width=LAYOUT.BORDER_WIDTH,
                corner_radius=LAYOUT.CORNER_RADIUS,
                fg_color=COLORS.SECONDARY_BG,
                font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            )
            entry.grid(row=row + 1, column=0, padx=0, pady=(0, 10), sticky="w")

            button = customtkinter.CTkButton(
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
                command=command,
            )
            button.grid(row=row + 1, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

            row += 2

        database_info_label = customtkinter.CTkLabel(
            self,
            textvariable=database_info,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        database_info_label.grid(row=row, column=0, padx=0, sticky="w")

        database_entry = customtkinter.CTkEntry(
            self,
            textvariable=self.database_path,
            text_color=COLORS.TEXT_PRIMARY,
            width=370,
            height=LAYOUT.ENTRY_HEIGHT,
            border_width=LAYOUT.BORDER_WIDTH,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        database_entry.grid(row=row + 1, column=0, padx=0, pady=(0, 10), sticky="w")

        database_button = customtkinter.CTkButton(
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
            command=self.browse_database,
        )
        database_button.grid(row=row + 1, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

        row += 2

        database_selector_info = tk.StringVar(value="Select Database File (.nal/.ndb)")
        database_selector_info_label = customtkinter.CTkLabel(
            self,
            textvariable=database_selector_info,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        database_selector_info_label.grid(row=row, column=0, padx=0, sticky="w")

        self.database_selector_combo = customtkinter.CTkComboBox(
            self,
            variable=self.database_selector,
            values=[],
            width=370,
            height=LAYOUT.ENTRY_HEIGHT,
            border_width=LAYOUT.BORDER_WIDTH,
            corner_radius=LAYOUT.CORNER_RADIUS,
            fg_color=COLORS.SECONDARY_BG,
            text_color=COLORS.TEXT_PRIMARY,
            button_color=COLORS.HOVER_BG,
            button_hover_color=COLORS.BUTTON_HOVER,
            dropdown_fg_color=COLORS.SECONDARY_BG,
            dropdown_text_color=COLORS.TEXT_PRIMARY,
            dropdown_hover_color=COLORS.HOVER_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        self.database_selector_combo.grid(row=row + 1, column=0, padx=0, pady=(0, 10), sticky="w")

        row += 2

        folder_entries = [
            (samples_info, self.samples_path, self.browse_samples),
            (outputs_info, self.outputs_path, self.browse_outputs),
        ]

        for info_var, path_var, command in folder_entries:
            info_label = customtkinter.CTkLabel(
                self,
                textvariable=info_var,
                text_color=COLORS.TEXT_PRIMARY,
                font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            )
            info_label.grid(row=row, column=0, padx=0, sticky="w")

            entry = customtkinter.CTkEntry(
                self,
                textvariable=path_var,
                text_color=COLORS.TEXT_PRIMARY,
                width=370,
                height=LAYOUT.ENTRY_HEIGHT,
                border_width=LAYOUT.BORDER_WIDTH,
                corner_radius=LAYOUT.CORNER_RADIUS,
                fg_color=COLORS.SECONDARY_BG,
                font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            )
            entry.grid(row=row + 1, column=0, padx=0, pady=(0, 10), sticky="w")

            button = customtkinter.CTkButton(
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
                command=command,
            )
            button.grid(row=row + 1, column=1, padx=(10, 0), pady=(0, 10), sticky="w")

            row += 2

        instruction_label = customtkinter.CTkLabel(
            self,
            textvariable=instruction,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        instruction_label.grid(row=row, column=0, columnspan=2, padx=0, pady=(10, 0), sticky="ew")

        row += 1

        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        config_button = customtkinter.CTkButton(
            self,
            height=LAYOUT.BUTTON_HEIGHT,
            border_width=0,
            corner_radius=LAYOUT.CORNER_RADIUS,
            text="CONFIG",
            fg_color=COLORS.SECONDARY_BG,
            text_color=COLORS.TEXT_PRIMARY,
            hover_color=COLORS.HOVER_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_NORMAL, FONTS.STYLE_BOLD),
            command=self.open_config,
        )
        config_button.grid(row=row, column=0, padx=(0, 5), sticky="ew")

        self.analyse_button = customtkinter.CTkButton(
            self,
            height=LAYOUT.BUTTON_HEIGHT,
            border_width=0,
            corner_radius=LAYOUT.CORNER_RADIUS,
            text="ANALYSE",
            fg_color=COLORS.ACCENT,
            text_color=COLORS.PRIMARY_BG,
            hover_color=COLORS.ACCENT_HOVER,
            font=(FONTS.FAMILY, FONTS.SIZE_NORMAL, FONTS.STYLE_BOLD),
            command=self.analyse,
            state="disabled",
        )
        self.analyse_button.grid(row=row, column=1, padx=(5, 0), sticky="ew")

        self._setup_field_validation()

    def _setup_field_validation(self) -> None:
        """設定欄位驗證

        Setup field validation.
        """
        self.cutadapt_path.trace_add("write", lambda *_args: self._check_fields())
        self.usearch_path.trace_add("write", lambda *_args: self._check_fields())
        self.blastn_path.trace_add("write", lambda *_args: self._check_fields())
        self.database_path.trace_add("write", lambda *_args: self._check_fields())
        self.database_selector.trace_add("write", lambda *_args: self._check_fields())
        self.samples_path.trace_add("write", lambda *_args: self._check_fields())
        self.outputs_path.trace_add("write", lambda *_args: self._check_fields())
        self._check_fields()

    def _check_fields(self) -> None:
        """檢查所有必要欄位是否已填寫

        Check if all required fields are filled.
        """
        all_filled = (
            bool(self.cutadapt_path.get())
            and bool(self.usearch_path.get())
            and bool(self.blastn_path.get())
            and bool(self.database_path.get())
            and bool(self.database_selector.get())
            and bool(self.samples_path.get())
            and bool(self.outputs_path.get())
        )

        if all_filled:
            self.analyse_button.configure(state="normal")
        else:
            self.analyse_button.configure(state="disabled")

    def open_config(self) -> None:
        """開啟配置視窗

        Open configuration window.
        """
        config_window = ConfigWindow(self)
        config_window.focus()

    def browse_cutadapt(self) -> None:
        """瀏覽選擇 Cutadapt 執行檔

        Browse to select Cutadapt executable.
        """
        filename = filedialog.askopenfilenames()
        if filename:
            self.cutadapt_path.set(filename[0] if isinstance(filename, tuple) else filename)
        self._check_fields()

    def browse_usearch(self) -> None:
        """瀏覽選擇 Usearch 執行檔

        Browse to select Usearch executable.
        """
        filename = filedialog.askopenfilenames()
        if filename:
            self.usearch_path.set(filename[0] if isinstance(filename, tuple) else filename)
        self._check_fields()

    def browse_blastn(self) -> None:
        """瀏覽選擇 Blastn 執行檔

        Browse to select Blastn executable.
        """
        filename = filedialog.askopenfilenames()
        if filename:
            self.blastn_path.set(filename[0] if isinstance(filename, tuple) else filename)
        self._check_fields()

    def browse_database(self) -> None:
        """瀏覽選擇資料庫資料夾

        Browse to select database folder.
        """
        foldername = filedialog.askdirectory()
        if not foldername:
            return

        self.database_path.set(foldername)

        db_path = Path(foldername)
        taxdb_btd_path = db_path / "taxdb.btd"
        taxdb_bti_path = db_path / "taxdb.bti"

        if not taxdb_btd_path.exists() or not taxdb_bti_path.exists():
            messagebox.showwarning(
                "Warning", "no taxdb.btd & taxdb.bti files found, scientific name will be N/A"
            )

        database_files = []
        for file_path in db_path.iterdir():
            if file_path.is_file():
                if file_path.suffix == ".nal":
                    database_files.append(f"(Combined) {file_path.stem}")
                elif file_path.suffix == ".ndb":
                    # 過濾掉附屬檔案 (如 .00.ndb, .01.ndb 等)
                    # 只保留純正的 .ndb 檔案 (不以 .XX.ndb 結尾)
                    file_name = file_path.name
                    if not re.search(r"\.\d{2}\.ndb$", file_name):
                        database_files.append(file_path.stem)

        if database_files:
            self.database_selector_combo.configure(values=database_files)
            self.database_selector.set(database_files[0])
        else:
            self.database_selector_combo.configure(values=[])
            self.database_selector.set("")
        self._check_fields()

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

    def analyse(self) -> None:
        """執行 NGS 分析

        Execute NGS analysis.
        """
        samples_dir = Path(self.samples_path.get())
        outputs_dir = Path(self.outputs_path.get())

        if not samples_dir.exists() or not outputs_dir.exists():
            logger.error("樣本或輸出資料夾不存在")
            messagebox.showerror("Error", "Samples or outputs folder does not exist")
            return

        input_files = natsort.natsorted(
            [f for f in samples_dir.iterdir() if f.is_file() and f.suffix.lower() == ".fastq"]
        )
        sample_size = len(input_files) // 2

        if sample_size == 0:
            logger.error("未找到任何樣本檔案")
            messagebox.showerror("Error", "No sample files found")
            return

        gc.collect()

        folder_names = [
            "A_primer_trimming",
            "B_merged",
            "C_quality",
            "D_length",
            "E_uniques",
            "F_OTUs",
            "G_OTUtable",
            "H_blasts",
            "I_sorted_blasts",
        ]

        folders: dict[str, Path] = {}
        for folder_name in folder_names:
            folder_path = outputs_dir / folder_name
            if folder_path.exists():
                try:
                    shutil.rmtree(folder_path)
                except PermissionError:
                    for item in folder_path.iterdir():
                        try:
                            if item.is_dir():
                                shutil.rmtree(item, ignore_errors=True)
                            else:
                                item.unlink()
                        except Exception:
                            pass
            folder_path.mkdir(parents=True, exist_ok=True)
            folders[folder_name] = folder_path

        processor = NGSProcessor(
            cutadapt_path=self.cutadapt_path.get(),
            usearch_path=self.usearch_path.get(),
            blastn_path=self.blastn_path.get(),
            database_path=self.database_path.get(),
            database_selector=self.database_selector.get(),
            config=self.config,
            folders=folders,
        )

        try:
            processor.run_analysis(samples_dir, input_files, sample_size)
            logger.info("分析完成")
            messagebox.showinfo("Success", "Analysis completed successfully")
        except Exception as e:
            logger.error(f"分析過程發生錯誤: {e}")
            messagebox.showerror("Error", f"Analysis failed: {e}")


class NGSProcessor:
    """NGS 分析處理器

    NGS analysis processor.
    """

    def __init__(
        self,
        cutadapt_path: str,
        usearch_path: str,
        blastn_path: str,
        database_path: str,
        database_selector: str,
        config: NGSConfig,
        folders: dict[str, Path],
    ) -> None:
        """初始化 NGS 處理器

        Initialize NGS processor.

        Args:
            cutadapt_path (str): Cutadapt 執行檔路徑 / Cutadapt executable path.
            usearch_path (str): Usearch 執行檔路徑 / Usearch executable path.
            blastn_path (str): Blastn 執行檔路徑 / Blastn executable path.
            database_path (str): 資料庫路徑 / Database path.
            database_selector (str): 資料庫選擇器 / Database selector.
            config (NGSConfig): NGS 設定 / NGS configuration.
            folders (dict[str, Path]): 資料夾路徑字典 / Folder paths dictionary.
        """
        self.cutadapt_path = cutadapt_path
        self.usearch_path = usearch_path
        self.blastn_path = blastn_path
        self.database_path = database_path
        self.database_selector = database_selector
        self.config = config
        self.folders = folders

    def run_analysis(
        self, samples_dir: Path, input_files: Sequence[Path], sample_size: int
    ) -> None:
        """執行完整分析流程

        Execute complete analysis workflow.

        Args:
            samples_dir (Path): 樣本資料夾路徑 / Samples directory path.
            input_files (Sequence[Path]): 輸入檔案列表 / Input files list.
            sample_size (int): 樣本數量 / Sample size.
        """
        logger.info("開始 NGS 分析")

        logger.info(f"步驟 1/9: 修剪 Primers (共 {sample_size} 個樣本)")
        for i in range(0, sample_size * 2, 2):
            self._trim_primers(samples_dir, input_files[i].name, input_files[i + 1].name)
        gc.collect()

        primer_trimming_files = natsort.natsorted(
            [f for f in self.folders["A_primer_trimming"].iterdir() if f.is_file()]
        )

        logger.info(f"步驟 2/9: 合併配對序列 (共 {sample_size} 個樣本)")
        for i in range(0, sample_size * 2, 2):
            self._merge_pairs(primer_trimming_files[i].name, primer_trimming_files[i + 1].name)
        gc.collect()

        merged_files = natsort.natsorted(list(self.folders["B_merged"].iterdir()))

        logger.info(f"步驟 3/9: 品質控制 (共 {len(merged_files)} 個檔案)")
        for merged_file in merged_files:
            self._quality_control(merged_file.name)
        gc.collect()

        qualified_files = natsort.natsorted(list(self.folders["C_quality"].iterdir()))

        logger.info(f"步驟 4/9: 過濾長度 (共 {len(qualified_files)} 個檔案)")
        for qualified_file in qualified_files:
            self._filter_length(qualified_file.name)
        gc.collect()

        length_files = natsort.natsorted(list(self.folders["D_length"].iterdir()))
        adjusted_sample_size = len(length_files)

        logger.info(f"步驟 5/9: 聚類序列 (共 {adjusted_sample_size} 個檔案)")
        for length_file in length_files:
            self._cluster(length_file.name)
        gc.collect()

        uniques_files = natsort.natsorted(list(self.folders["E_uniques"].iterdir()))

        logger.info(f"步驟 6/9: 建立 OTU (共 {len(uniques_files)} 個檔案)")
        for uniques_file in uniques_files:
            self._create_otu(uniques_file.name)
        gc.collect()

        otu_files = natsort.natsorted(list(self.folders["F_OTUs"].iterdir()))

        logger.info(f"步驟 7/9: 建立 OTU 表格 (共 {len(otu_files)} 個檔案)")
        for merged_file, otu_file in zip(merged_files, otu_files, strict=False):
            self._create_otu_table(merged_file.name, otu_file.name)
        gc.collect()

        otu_table_files = natsort.natsorted(list(self.folders["G_OTUtable"].iterdir()))

        logger.info(f"步驟 8/9: 重新命名 OTU 表格 (共 {len(otu_table_files) // 2} 個檔案)")
        for i in range(1, adjusted_sample_size * 2 + 1, 2):
            self._rename_otu_table(otu_table_files[i].name)
        gc.collect()

        otu_files = natsort.natsorted(list(self.folders["F_OTUs"].iterdir()))

        logger.info(f"步驟 9/9: 執行 BLAST (共 {len(otu_files)} 個檔案)")
        for otu_file in otu_files:
            self._run_blast(otu_file.name)
        gc.collect()

        logger.info("合併 BLAST 結果")
        self._combine_blast_results()
        gc.collect()

        logger.info("NGS 分析完成")

    def _trim_primers(self, samples_dir: Path, r1: str, r2: str) -> None:
        """修剪 Primers

        Trim primers.

        Args:
            samples_dir (Path): 樣本資料夾路徑 / Samples directory path.
            r1 (str): R1 檔案名稱 / R1 file name.
            r2 (str): R2 檔案名稱 / R2 file name.
        """
        logger.info(f"修剪 Primers: {r1} / {r2}")
        forward = self.config.forward_primer
        reverse = self.config.reverse_primer
        rev_comp_forward = reverse_complement(forward)
        rev_comp_reverse = reverse_complement(reverse)

        primer_trimming_dir = self.folders["A_primer_trimming"]
        report_dir = primer_trimming_dir / "report"
        report_dir.mkdir(parents=True, exist_ok=True)

        filename_base = Path(r1).stem
        json_output = report_dir / f"{filename_base}.cutadapt.json"

        trimming_cmd = [
            self.cutadapt_path,
            "-a",
            f"{forward}...{reverse}",
            "-A",
            f"{rev_comp_reverse}...{rev_comp_forward}",
            "--discard-untrimmed",
            "--json",
            str(json_output),
            "-o",
            str(primer_trimming_dir / f"{r1}_TRIMMED_R1.fastq"),
            "-p",
            str(primer_trimming_dir / f"{r2}_TRIMMED_R2.fastq"),
            str(samples_dir / r1),
            str(samples_dir / r2),
        ]
        run_command(trimming_cmd)
        logger.info(f"完成修剪 Primers: {r1} / {r2}")

    def _merge_pairs(self, r1: str, r2: str) -> None:
        """合併配對序列

        Merge paired sequences.

        Args:
            r1 (str): R1 檔案名稱 / R1 file name.
            r2 (str): R2 檔案名稱 / R2 file name.
        """
        logger.info(f"合併配對序列: {r1} / {r2}")
        primer_trimming_dir = self.folders["A_primer_trimming"]
        merged_dir = self.folders["B_merged"]

        merging_cmd = [
            self.usearch_path,
            "-fastq_mergepairs",
            str(primer_trimming_dir / r1),
            "-reverse",
            str(primer_trimming_dir / r2),
            "-fastqout",
            str(merged_dir / f"{r1}_merged.fastq"),
        ]
        run_command(merging_cmd)
        logger.info(f"完成合併配對序列: {r1}")

    def _quality_control(self, file: str) -> None:
        """品質控制

        Quality control.

        Args:
            file (str): 檔案名稱 / File name.
        """
        logger.info(f"品質控制: {file}")
        merged_dir = self.folders["B_merged"]
        quality_dir = self.folders["C_quality"]
        quality_threshold = self.config.quality_threshold

        qualifying_cmd = [
            self.usearch_path,
            "-fastq_filter",
            str(merged_dir / file),
            "-fastq_truncqual",
            str(quality_threshold),
            "-fastqout",
            str(quality_dir / f"{file}_QUAL.fastq"),
        ]
        run_command(qualifying_cmd)
        logger.info(f"完成品質控制: {file}")

    def _filter_length(self, file: str) -> None:
        """過濾長度

        Filter length.

        Args:
            file (str): 檔案名稱 / File name.
        """
        logger.info(f"過濾長度: {file}")
        quality_dir = self.folders["C_quality"]
        length_dir = self.folders["D_length"]
        length_threshold = self.config.length_threshold

        length_filtering_cmd = [
            self.usearch_path,
            "-fastq_filter",
            str(quality_dir / file),
            "-fastq_minlen",
            str(length_threshold),
            "-fastaout",
            str(length_dir / f"{file}_LENG.fasta"),
        ]
        run_command(length_filtering_cmd)
        logger.info(f"完成過濾長度: {file}")

    def _cluster(self, file: str) -> None:
        """聚類序列

        Cluster sequences.

        Args:
            file (str): 檔案名稱 / File name.
        """
        logger.info(f"聚類序列: {file}")
        length_dir = self.folders["D_length"]
        uniques_dir = self.folders["E_uniques"]

        clustering_cmd = [
            self.usearch_path,
            "-fastx_uniques",
            str(length_dir / file),
            "-fastaout",
            str(uniques_dir / f"{file}_UNIQ.fasta"),
            "-sizeout",
            "-relabel",
            "Uniq",
        ]
        run_command(clustering_cmd)
        logger.info(f"完成聚類序列: {file}")

    def _create_otu(self, file: str) -> None:
        """建立 OTU

        Create OTU.

        Args:
            file (str): 檔案名稱 / File name.
        """
        logger.info(f"建立 OTU: {file}")
        uniques_dir = self.folders["E_uniques"]
        otu_dir = self.folders["F_OTUs"]

        otu_making_cmd = [
            self.usearch_path,
            "-unoise3",
            str(uniques_dir / file),
            "-zotus",
            str(otu_dir / f"{file}_ZOTU.fasta"),
        ]
        run_command(otu_making_cmd)
        logger.info(f"完成建立 OTU: {file}")

    def _create_otu_table(self, merged_file: str, otu_file: str) -> None:
        """建立 OTU 表格

        Create OTU table.

        Args:
            merged_file (str): 合併檔案名稱 / Merged file name.
            otu_file (str): OTU 檔案名稱 / OTU file name.
        """
        logger.info(f"建立 OTU 表格: {otu_file}")
        merged_dir = self.folders["B_merged"]
        otu_dir = self.folders["F_OTUs"]
        otu_table_dir = self.folders["G_OTUtable"]

        otu_tab_making_cmd = [
            self.usearch_path,
            "-otutab",
            str(merged_dir / merged_file),
            "-otus",
            str(otu_dir / otu_file),
            "-otutabout",
            str(otu_table_dir / f"{otu_file}_table.txt"),
            "-mapout",
            str(otu_table_dir / f"{otu_file}_map.txt"),
        ]
        run_command(otu_tab_making_cmd)
        logger.info(f"完成建立 OTU 表格: {otu_file}")

    def _rename_otu_table(self, file: str) -> None:
        """重新命名 OTU 表格

        Rename OTU table.

        Args:
            file (str): 檔案名稱 / File name.
        """
        logger.info(f"重新命名 OTU 表格: {file}")
        otu_table_dir = self.folders["G_OTUtable"]
        old_file = otu_table_dir / file
        new_name = otu_table_dir / f"{Path(file).stem}.txt"
        old_file.rename(new_name)
        logger.info(f"完成重新命名 OTU 表格: {file} -> {new_name.name}")

    def _run_blast(self, file: str) -> None:
        """執行 BLAST

        Run BLAST.

        Args:
            file (str): 檔案名稱 / File name.
        """
        logger.info(f"執行 BLAST: {file}")
        otu_dir = self.folders["F_OTUs"]
        blast_dir = self.folders["H_blasts"]
        db_display_name = self.database_selector
        # 如果資料庫名稱以 "(Combined) " 開頭, 則去掉前綴以獲取實際資料庫名稱
        if db_display_name.startswith("(Combined) "):
            db_name = db_display_name.replace("(Combined) ", "", 1)
        else:
            db_name = db_display_name

        blast_cmd = [
            self.blastn_path,
            "-query",
            str(otu_dir / file),
            "-out",
            str(blast_dir / f"{file}_blasted.txt"),
            "-db",
            db_name,
            "-outfmt",
            "6 qseqid pident qcovs sscinames sacc",
            "-max_target_seqs",
            "3",
        ]
        run_command(blast_cmd, cwd=self.database_path)
        logger.info(f"完成執行 BLAST: {file}")

    def _combine_blast_results(self) -> None:
        """合併 BLAST 結果

        Combine BLAST results.
        """
        logger.info("開始合併 BLAST 結果")
        blast_dir = self.folders["H_blasts"]
        otu_table_dir = self.folders["G_OTUtable"]
        sorted_blasts_dir = self.folders["I_sorted_blasts"]

        blast_files = sorted(blast_dir.glob("*.txt"))
        otu_table_files = sorted(otu_table_dir.glob("*.txt"))

        ref_path = self.config.ref_path
        if not ref_path:
            logger.warning("未找到參考檔案")
            return

        try:
            df_ref = pd.read_excel(ref_path, engine="openpyxl")
        except Exception as e:
            logger.error(f"讀取參考檔案失敗: {e}")
            return

        logger.info(f"開始處理 {len(blast_files)} 個 BLAST 結果檔案")
        for blast_file, otu_table_file in zip(blast_files, otu_table_files[1::2], strict=False):
            self._process_single_blast_result(blast_file, otu_table_file, df_ref, sorted_blasts_dir)
            gc.collect()
        logger.info(f"完成合併 BLAST 結果 (共處理 {len(blast_files)} 個檔案)")
        del df_ref
        gc.collect()

    def _process_single_blast_result(
        self,
        blast_file: Path,
        otu_table_file: Path,
        df_ref: pd.DataFrame,
        sorted_blasts_dir: Path,
    ) -> None:
        """處理單一 BLAST 結果

        Process single BLAST result.

        Args:
            blast_file (Path): BLAST 檔案路徑 / BLAST file path.
            otu_table_file (Path): OTU 表格檔案路徑 / OTU table file path.
            df_ref (pd.DataFrame): 參考資料 DataFrame / Reference data DataFrame.
            sorted_blasts_dir (Path): 排序後 BLAST 結果目錄 / Sorted BLAST results directory.
        """
        try:
            original_data = pd.read_csv(
                blast_file,
                engine="python",
                header=None,
                sep="\t",
                encoding="utf-8",
                names=["OTU", "Identity", "Coverage", "Scientific_name", "Accession_number"],
            )

            reads_data = pd.read_csv(
                otu_table_file,
                engine="python",
                header=None,
                sep="\t",
                encoding="utf-8",
                names=["OTU", "Reads"],
                skiprows=1,
            )
            total_reads = reads_data.Reads.sum()

            grouped_otu = original_data[original_data.Identity >= 97].reset_index(drop=True)
            drop_filtered = original_data[original_data.Identity < 97].copy()
            drop_filtered["Stat"] = "*FILTERED*"
            drop_filtered = drop_filtered.drop_duplicates(
                subset=["OTU", "Scientific_name"], keep="first"
            )

            grouped_otu["Stat"] = np.where(grouped_otu["OTU"].duplicated(), "*DUPE*", "")
            drop_dupe = grouped_otu[grouped_otu.Stat == "*DUPE*"]
            drop_dupe = drop_dupe.drop_duplicates(subset=["OTU", "Scientific_name"], keep="first")
            clean_data = grouped_otu.drop_duplicates(subset=["OTU"], keep="first")

            csv_data = pd.merge(clean_data, reads_data, on="OTU")
            ratio = csv_data.Reads / float(total_reads)
            csv_data.insert(7, column="Ratio", value=ratio)

            main_data = pd.concat([csv_data, drop_dupe], axis=0, ignore_index=True)
            main_data = main_data.sort_values(["OTU", "Stat"], ignore_index=True)

            with warnings.catch_warnings():
                warnings.filterwarnings("ignore", category=FutureWarning)
                main_data = (
                    main_data.groupby("OTU")
                    .apply(lambda x: x.drop_duplicates(subset=["Scientific_name"], keep="first"))
                    .reset_index(drop=True)
                )

            otus_with_valid_data = set(main_data["OTU"].unique())
            all_filtered_otus = set(drop_filtered["OTU"].unique()) - otus_with_valid_data

            filtered_with_reads = drop_filtered[drop_filtered["OTU"].isin(all_filtered_otus)].copy()
            filtered_with_reads = filtered_with_reads.drop_duplicates(subset=["OTU"], keep="first")
            filtered_with_reads = pd.merge(filtered_with_reads, reads_data, on="OTU")
            ratio_filtered = filtered_with_reads.Reads / float(total_reads)
            filtered_with_reads["Ratio"] = ratio_filtered

            filtered_without_reads = drop_filtered[
                ~drop_filtered["OTU"].isin(all_filtered_otus)
            ].copy()
            filtered_without_reads["Reads"] = np.nan
            filtered_without_reads["Ratio"] = np.nan

            all_filtered = pd.concat(
                [filtered_with_reads, filtered_without_reads], axis=0, ignore_index=True
            )

            final_csv = pd.concat([main_data, all_filtered], axis=0, ignore_index=True)
            final_csv = final_csv[
                [
                    "OTU",
                    "Identity",
                    "Coverage",
                    "Scientific_name",
                    "Stat",
                    "Reads",
                    "Ratio",
                    "Accession_number",
                ]
            ]
            final_csv = final_csv.sort_values("OTU", key=natsort.natsort_keygen())

            final_csv["Scientific_name"] = final_csv["Scientific_name"].astype(str)
            if "Scientific_name" in df_ref.columns:
                df_ref["Scientific_name"] = df_ref["Scientific_name"].astype(str)

            final_zh_csv = final_csv.merge(df_ref, how="left", on="Scientific_name")

            final_csv_styled = final_csv.style.apply(highlight_row, axis=1)
            final_zh_csv_styled = final_zh_csv.style.apply(highlight_row, axis=1)

            excel_path1 = sorted_blasts_dir / f"{otu_table_file.stem}.xlsx"
            excel_path2 = sorted_blasts_dir / f"{otu_table_file.stem}_zh_added.xlsx"

            with pd.ExcelWriter(excel_path1, engine="openpyxl") as writer:
                final_csv_styled.to_excel(writer, index=False, sheet_name="Sheet1")

            with pd.ExcelWriter(excel_path2, engine="openpyxl") as writer:
                final_zh_csv_styled.to_excel(writer, index=False, sheet_name="Sheet1")

            del (
                original_data,
                reads_data,
                grouped_otu,
                drop_filtered,
                drop_dupe,
                clean_data,
                csv_data,
                main_data,
                filtered_with_reads,
                filtered_without_reads,
                all_filtered,
                final_csv,
                final_zh_csv,
                final_csv_styled,
                final_zh_csv_styled,
            )
            gc.collect()

            logger.info(f"已處理 BLAST 結果: {blast_file.name}")

        except Exception as e:
            logger.error(f"處理 BLAST 結果失敗 {blast_file.name}: {e}")
            gc.collect()
