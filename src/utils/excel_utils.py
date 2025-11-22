"""Excel 相關工具函式

Excel-related utility functions.
"""

import pandas as pd


def highlight_row(row: pd.Series) -> list[str]:
    """根據狀態欄位標記該列的背景顏色

    Highlight row background color based on Stat column.

    Args:
        row (pd.Series): DataFrame 的一列資料 / A row of DataFrame.

    Returns:
        list[str]: CSS 樣式列表 / List of CSS styles.
    """
    condition = row.get("Stat", "")

    match condition:
        case "*DUPE*":
            css = "background-color: #00c2c7"
        case "*FILTERED*":
            css = "background-color: #94F7B2"
        case _:
            css = "background-color: transparent"

    return [css] * len(row)
