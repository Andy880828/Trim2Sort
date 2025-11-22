"""UI 相關設定常數

UI-related configuration constants.
"""

from dataclasses import dataclass


@dataclass(slots=True, frozen=True)
class UIColors:
    """UI 顏色設定

    UI color configurations.
    """

    PRIMARY_BG: str = "#091235"
    SECONDARY_BG: str = "#2B4257"
    HOVER_BG: str = "#4C6A78"
    TEXT_PRIMARY: str = "#DCF3F0"
    TEXT_SECONDARY: str = "#88A9C3"
    ACCENT: str = "#D4AF37"
    ACCENT_HOVER: str = "#F4CF5A"
    BUTTON_HOVER: str = "#5D7B8A"
    DROPDOWN_BG: str = "#4C6A78"


@dataclass(slots=True, frozen=True)
class UIFonts:
    """UI 字型設定

    UI font configurations.
    """

    FAMILY: str = "Consolas"
    SIZE_NORMAL: int = 11
    SIZE_MEDIUM: int = 12
    STYLE_BOLD: str = "bold"


@dataclass(slots=True, frozen=True)
class UILayout:
    """UI 佈局設定

    UI layout configurations.
    """

    BUTTON_HEIGHT: int = 25
    ENTRY_HEIGHT: int = 25
    BORDER_WIDTH: int = 2
    CORNER_RADIUS: int = 8


# 建立全域實例
COLORS = UIColors()
FONTS = UIFonts()
LAYOUT = UILayout()
