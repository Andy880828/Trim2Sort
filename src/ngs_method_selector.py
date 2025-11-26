"""NGS 分析方法選擇器

NGS analysis method selector.
"""

import customtkinter
from PIL import Image

from src.ngs import NGS
from src.ngs_dada2 import NGS_DADA2
from src.utils.logger_utils import get_logger
from src.utils.path_utils import get_icon_path
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


class NGSMethodSelector(customtkinter.CTkToplevel):
    """NGS 分析方法選擇視窗

    NGS analysis method selector window.
    """

    def __init__(self, *args, **kwargs) -> None:
        """初始化 NGS 方法選擇視窗

        Initialize NGS method selector window.

        Args:
            *args: 位置參數 / Positional arguments.
            **kwargs: 關鍵字參數 / Keyword arguments.
        """
        super().__init__(*args, **kwargs)
        self.title("Choose NGS Method")
        self.geometry("320x280")
        self.configure(fg_color=COLORS.PRIMARY_BG)
        self.resizable(False, False)

        self.ngs_window: customtkinter.CTkToplevel | None = None

        self._setup_ui()

    def _setup_ui(self) -> None:
        """設定使用者介面

        Setup user interface.
        """
        my_image = load_app_image()

        logo = customtkinter.CTkLabel(self, image=my_image, text="")
        logo.pack(pady=(20, 10))

        label = customtkinter.CTkLabel(
            self,
            text="Choose Analysis Method:",
            fg_color="transparent",
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        label.pack(pady=(0, 20))

        button_usearch = customtkinter.CTkButton(
            self,
            text="USEARCH",
            command=self.open_usearch,
            fg_color=COLORS.SECONDARY_BG,
            hover_color=COLORS.HOVER_BG,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            corner_radius=LAYOUT.CORNER_RADIUS,
            width=250,
            height=LAYOUT.BUTTON_HEIGHT,
        )
        button_usearch.pack(pady=5)

        button_dada2 = customtkinter.CTkButton(
            self,
            text="DADA2",
            command=self.open_dada2,
            fg_color=COLORS.SECONDARY_BG,
            hover_color=COLORS.HOVER_BG,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            corner_radius=LAYOUT.CORNER_RADIUS,
            width=250,
            height=LAYOUT.BUTTON_HEIGHT,
        )
        button_dada2.pack(pady=5)

        button_cancel = customtkinter.CTkButton(
            self,
            text="CANCEL",
            command=self.destroy,
            fg_color=COLORS.ACCENT,
            hover_color=COLORS.ACCENT_HOVER,
            text_color=COLORS.PRIMARY_BG,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            corner_radius=LAYOUT.CORNER_RADIUS,
            width=250,
            height=LAYOUT.BUTTON_HEIGHT,
        )
        button_cancel.pack(pady=5)

    def open_usearch(self) -> None:
        """開啟 USEARCH 分析視窗

        Open USEARCH analysis window.
        """
        logger.info("選擇 USEARCH 分析方法")
        if self.ngs_window is None or not self.ngs_window.winfo_exists():
            self.ngs_window = NGS(self.master)
            self.destroy()
        else:
            self.ngs_window.focus()
            self.destroy()

    def open_dada2(self) -> None:
        """開啟 DADA2 分析視窗

        Open DADA2 analysis window.
        """
        logger.info("選擇 DADA2 分析方法")
        if self.ngs_window is None or not self.ngs_window.winfo_exists():
            self.ngs_window = NGS_DADA2(self.master)
            self.destroy()
        else:
            self.ngs_window.focus()
            self.destroy()
