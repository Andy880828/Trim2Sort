"""Trim2Sort 主程式

Trim2Sort main application.
"""

import customtkinter
from PIL import Image

from src.ngs_method_selector import NGSMethodSelector
from src.sanger import Sanger
from src.utils.logger_utils import init_logger
from src.utils.path_utils import get_icon_path, get_project_root
from src.utils.ui_config import COLORS, FONTS, LAYOUT

# 初始化根 Logger(僅在應用程式啟動時執行一次)
LOG_FILE = get_project_root() / "logs" / "trim2sort.log"
init_logger(level="INFO", log_file=LOG_FILE, use_color=True)


def load_app_image() -> customtkinter.CTkImage:
    """載入應用程式圖示

    Load application icon.

    Returns:
        customtkinter.CTkImage: 應用程式圖示 / Application icon.
    """
    icon_path = get_icon_path()
    return customtkinter.CTkImage(light_image=Image.open(icon_path), size=(128, 72))


class App(customtkinter.CTk):
    """Trim2Sort 主應用程式視窗

    Trim2Sort main application window.
    """

    def __init__(self, *args, **kwargs) -> None:
        """初始化主應用程式視窗

        Initialize main application window.

        Args:
            *args: 位置參數 / Positional arguments.
            **kwargs: 關鍵字參數 / Keyword arguments.
        """
        super().__init__(*args, **kwargs)
        self.geometry("280x170")
        self.resizable(False, False)
        self.title("Trim2Sort")
        self.configure(fg_color=COLORS.PRIMARY_BG)

        self.grid_rowconfigure(0, weight=0)
        self.grid_columnconfigure((0, 1), weight=1)

        self.toplevel_window: customtkinter.CTkToplevel | None = None

        self._setup_ui()

    def _setup_ui(self) -> None:
        """設定使用者介面

        Setup user interface.
        """
        my_image = load_app_image()

        logo = customtkinter.CTkLabel(self, image=my_image, text="")
        logo.grid(row=0, column=0, columnspan=2, padx=20)

        label = customtkinter.CTkLabel(
            self,
            text="Step0: Choose the sequence format",
            fg_color="transparent",
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
        )
        label.grid(row=1, column=0, columnspan=2, padx=20, pady=(10, 5))

        button_sanger = customtkinter.CTkButton(
            self,
            text="Sanger",
            command=self.open_sanger_analysis,
            fg_color=COLORS.SECONDARY_BG,
            hover_color=COLORS.HOVER_BG,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            corner_radius=LAYOUT.CORNER_RADIUS,
        )
        button_sanger.grid(row=2, column=0, padx=(10, 10), pady=(10, 10), sticky="ew")

        button_ngs = customtkinter.CTkButton(
            self,
            text="NGS",
            command=self.open_ngs_analysis,
            fg_color=COLORS.SECONDARY_BG,
            hover_color=COLORS.HOVER_BG,
            text_color=COLORS.TEXT_PRIMARY,
            font=(FONTS.FAMILY, FONTS.SIZE_MEDIUM, FONTS.STYLE_BOLD),
            corner_radius=LAYOUT.CORNER_RADIUS,
        )
        button_ngs.grid(row=2, column=1, padx=(10, 10), pady=(10, 10), sticky="ew")

    def open_sanger_analysis(self) -> None:
        """開啟 Sanger 分析視窗

        Open Sanger analysis window.
        """
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = Sanger(self)
        else:
            self.toplevel_window.focus()

    def open_ngs_analysis(self) -> None:
        """開啟 NGS 分析視窗

        Open NGS analysis window.
        """
        if self.toplevel_window is None or not self.toplevel_window.winfo_exists():
            self.toplevel_window = NGSMethodSelector(self)
        else:
            self.toplevel_window.focus()


if __name__ == "__main__":
    app = App()
    app.mainloop()
