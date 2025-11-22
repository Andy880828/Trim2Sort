"""路徑相關工具函式

Path-related utility functions.
"""

from pathlib import Path

from src.utils.logger_utils import get_logger

logger = get_logger(__name__)


def get_project_root() -> Path:
    """取得專案根目錄

    Get the project root directory.

    Returns:
        Path: 專案根目錄路徑 / Project root directory path.
    """
    current_file = Path(__file__).resolve()
    return current_file.parent.parent.parent


def get_icon_path() -> Path:
    """取得圖示檔案路徑

    Get the icon file path.

    Returns:
        Path: 圖示檔案路徑 / Icon file path.
    """
    return get_project_root() / "static" / "Trim2Sort_icon.png"


def find_latest_ref_file() -> str:
    """尋找最新的 ref_ 開頭的 xlsx 檔案

    Find the latest xlsx file starting with ref_.

    Returns:
        str: 最新 ref 檔案的完整路徑, 若無則回傳空字串 / Full path of the latest ref file, or empty string if none found.
    """
    docs_dir = get_project_root() / "docs"
    ref_files: list[tuple[Path, float]] = []

    if docs_dir.exists():
        for file_path in docs_dir.iterdir():
            if (
                file_path.is_file()
                and file_path.name.startswith("ref_")
                and file_path.suffix == ".xlsx"
            ):
                ref_files.append((file_path, file_path.stat().st_mtime))

    if ref_files:
        ref_files.sort(key=lambda x: x[1], reverse=True)
        logger.info(f"找到最新的 ref 檔案: {ref_files[0][0]}")
        return str(ref_files[0][0])

    logger.warning("未找到任何 ref 檔案")
    return ""


def get_cutadapt_path() -> Path:
    """取得 cutadapt.exe 路徑

    Get the cutadapt.exe path.

    Returns:
        Path: cutadapt.exe 路徑 / cutadapt.exe path.
    """
    return get_project_root() / "dependencies" / "cutadapt.exe"


def get_usearch_path() -> Path:
    """取得 usearch.exe 路徑

    Get the usearch.exe path.

    Returns:
        Path: usearch.exe 路徑 / usearch.exe path.
    """
    return get_project_root() / "dependencies" / "usearch.exe"


def get_blastn_path() -> Path:
    """取得 blastn.exe 路徑

    Get the blastn.exe path.

    Returns:
        Path: blastn.exe 路徑 / blastn.exe path.
    """
    return get_project_root() / "dependencies" / "blast+" / "bin" / "blastn.exe"


def get_trimmomatic_path() -> Path:
    """取得 Trimmomatic jar 檔案路徑

    Get the Trimmomatic jar file path.

    Returns:
        Path: Trimmomatic jar 檔案路徑 / Trimmomatic jar file path.
    """
    return (
        get_project_root()
        / "dependencies"
        / "Trimmomatic-0.39"
        / "Trimmomatic-0.39"
        / "trimmomatic-0.39.jar"
    )


def get_primer_path(primer_name: str) -> Path:
    """取得 Primer 檔案路徑

    Get the primer file path.

    Args:
        primer_name (str): Primer 檔案名稱 / Primer file name.

    Returns:
        Path: Primer 檔案路徑 / Primer file path.
    """
    return (
        get_project_root()
        / "dependencies"
        / "Trimmomatic-0.39"
        / "Trimmomatic-0.39"
        / "adapters"
        / primer_name
    )
