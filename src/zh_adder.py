"""中文名稱添加工具

Chinese name addition utility.
"""

from pathlib import Path

import pandas as pd

from src.utils.logger_utils import get_logger
from src.utils.path_utils import find_latest_ref_file

logger = get_logger(__name__)


class ZhAdder:
    """中文名稱添加器

    Chinese name adder for adding Chinese names to Excel files.
    """

    def __init__(self, input_folder: Path | str, output_folder: Path | str) -> None:
        """初始化中文名稱添加器

        Initialize Chinese name adder.

        Args:
            input_folder (Path | str): 輸入資料夾路徑 / Input folder path.
            output_folder (Path | str): 輸出資料夾路徑 / Output folder path.
        """
        self.input_folder = Path(input_folder)
        self.output_folder = Path(output_folder)
        self.sample_files = list(self.input_folder.glob("*.xlsx"))
        self.sample_size = len(self.sample_files)

    def add_chinese_names(self) -> None:
        """為所有 Excel 檔案添加中文名稱

        Add Chinese names to all Excel files.
        """
        ref_path = find_latest_ref_file()
        if not ref_path:
            logger.error("未找到參考檔案")
            return

        try:
            df_ref = pd.read_excel(ref_path, engine="openpyxl")
        except Exception as e:
            logger.error(f"讀取參考檔案失敗: {e}")
            return

        self.output_folder.mkdir(parents=True, exist_ok=True)

        for sample_file in self.sample_files:
            try:
                df = pd.read_excel(sample_file, engine="openpyxl")
                df_ref_merged = df.merge(df_ref, how="left", on="Scientific_name")

                output_name = f"{sample_file.stem}_zh_added.xlsx"
                output_path = self.output_folder / output_name

                df_ref_merged.to_excel(output_path, engine="openpyxl", index=False)
                logger.info(f"已處理: {sample_file.name} -> {output_name}")
            except Exception as e:
                logger.error(f"處理檔案 {sample_file.name} 失敗: {e}")


def main() -> None:
    """主程式入口

    Main program entry point.
    """
    from src.utils.path_utils import get_project_root

    root_dir = get_project_root()
    input_folder = root_dir / "outputs" / "I_sorted_blasts"
    output_folder = root_dir / "outputs" / "TEST"

    zh_adder = ZhAdder(input_folder, output_folder)
    zh_adder.add_chinese_names()


if __name__ == "__main__":
    main()
