"""BLAST 核酸資料庫建置輔助腳本

BLAST nucleotide database construction helper script.

請在下方 def main() 函式中修改參數, 並將檔案路徑改為實際路徑。

please modify the parameters in the def main() function and change the file paths to actual paths.
"""

import gzip
import logging
from pathlib import Path
import shutil
import sys
import tarfile

# 將專案根目錄添加到 sys.path, 以便導入 src 模組(不然現在是獨立於 src 外的腳本)
_script_dir = Path(__file__).resolve().parent
_project_root = _script_dir.parent
if str(_project_root) not in sys.path:
    sys.path.insert(0, str(_project_root))

from src.utils.path_utils import get_project_root  # noqa: E402
from src.utils.subprocess_utils import run_command  # noqa: E402

# 初始化獨立的 logger(不影響主程式)
logger = logging.getLogger("make_blastndb_helper")
logger.setLevel(logging.INFO)

# 清除這個 logger 的 handlers(如果有的話)
if logger.handlers:
    logger.handlers.clear()

# 設定 console handler
console_handler = logging.StreamHandler(sys.stdout)
console_handler.setLevel(logging.INFO)
console_formatter = logging.Formatter(
    "%(asctime)s | %(levelname)-8s | %(filename)s:%(lineno)d | %(funcName)s() | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
console_handler.setFormatter(console_formatter)
logger.addHandler(console_handler)

# 設定 file handler
log_file = get_project_root() / "logs" / "make_blastndb_helper.log"
log_file.parent.mkdir(parents=True, exist_ok=True)
file_handler = logging.FileHandler(log_file, encoding="utf-8", mode="a")
file_handler.setLevel(logging.INFO)
file_formatter = logging.Formatter(
    "%(asctime)s | %(levelname)-8s | %(filename)s:%(lineno)d | %(message)s",
    datefmt="%Y-%m-%d %H:%M:%S",
)
file_handler.setFormatter(file_formatter)
logger.addHandler(file_handler)

# 防止傳播到根 logger, 避免影響主程式
logger.propagate = False


def extract_gzip_file(gzip_path: Path, output_path: Path) -> None:
    """解壓縮 gzip 檔案

    Extract gzip file.

    Args:
        gzip_path (Path): gzip 檔案路徑 / Gzip file path.
        output_path (Path): 輸出檔案路徑 / Output file path.
    """
    logger.info(f"解壓縮 gzip 檔案: {gzip_path} -> {output_path}")
    with gzip.open(gzip_path, "rb") as f_in, open(output_path, "wb") as f_out:
        shutil.copyfileobj(f_in, f_out)
    logger.info(f"解壓縮完成: {output_path}")


def extract_tar_gz_file(tar_gz_path: Path, output_dir: Path) -> None:
    """解壓縮 tar.gz 檔案(可能需解兩層)

    Extract tar.gz file (may need to extract twice).

    Args:
        tar_gz_path (Path): tar.gz 檔案路徑 / Tar.gz file path.
        output_dir (Path): 輸出目錄 / Output directory.
    """
    logger.info(f"解壓縮 tar.gz 檔案: {tar_gz_path} -> {output_dir}")
    with tarfile.open(tar_gz_path, "r:gz") as tar:
        tar.extractall(path=output_dir, filter="data")
    logger.info(f"第一層解壓縮完成, 檔案已解壓到: {output_dir}")

    # 檢查是否需要解第二層(尋找解壓後的 .tar.gz 檔案)
    nested_tar_gz = list(output_dir.glob("*.tar.gz"))
    if nested_tar_gz:
        logger.info("發現嵌套的 tar.gz 檔案, 進行第二層解壓縮")
        for nested_tar in nested_tar_gz:
            with tarfile.open(nested_tar, "r:gz") as tar:
                tar.extractall(path=output_dir, filter="data")
            logger.info(f"第二層解壓縮完成: {nested_tar}")
    logger.info(f"解壓縮完成, 檔案已解壓到: {output_dir}")


def parse_accession2taxid(accession2taxid_path: Path, output_path: Path) -> None:
    """解析 accession2taxid 檔案

    Parse accession2taxid file (equivalent to: sed '1d' | awk '{print $2" "$3}').

    Args:
        accession2taxid_path (Path): accession2taxid 檔案路徑 / Accession2taxid file path.
        output_path (Path): 輸出檔案路徑 / Output file path.
    """
    logger.info(f"解析 accession2taxid 檔案: {accession2taxid_path}")

    buffer_size = 10000  # 一次寫 1 萬行
    buffer = []

    with (
        open(accession2taxid_path, encoding="utf-8") as f_in,
        open(output_path, "w", encoding="utf-8") as f_out,
    ):
        next(f_in)  # skip header

        for line in f_in:
            parts = line.rstrip("\n").split("\t")
            if len(parts) >= 3:
                buffer.append(f"{parts[1]} {parts[2]}\n")

            if len(buffer) >= buffer_size:
                f_out.writelines(buffer)
                buffer.clear()

        # 最後 flush
        if buffer:
            f_out.writelines(buffer)
    logger.info(f"解析完成, 輸出到: {output_path}")


def check_database_files(output_path: Path) -> bool:
    """檢查資料庫檔案是否存在

    Check if database files (.ndb and .ntf) exist.

    Args:
        output_path (Path): 輸出目錄 / Output directory.

    Returns:
        bool: 如果 .ndb 和 .ntf 檔案都存在則回傳 True / True if both .ndb and .ntf files exist.
    """
    ndb_files = list(output_path.glob("*.ndb"))
    ntf_files = list(output_path.glob("*.ntf"))
    has_ndb = len(ndb_files) > 0
    has_ntf = len(ntf_files) > 0

    if has_ndb and has_ntf:
        logger.info("資料庫建置成功, 找到 .ndb 和 .ntf 檔案")
        return True
    else:
        logger.error(f"資料庫建置失敗, 缺少必要檔案: .ndb={has_ndb}, .ntf={has_ntf}")
        return False


def main() -> None:
    """主函式

    Main function.
    """
    # 輸入參數(檔案路徑需全英文)
    # Input parameters (file paths must be in English only)
    accession2taxid_path = Path(
        r"C:\Users\Andy\Desktop\Data\database\course\nucl_gb.accession2taxid.gz"
    )  # 請修改為實際路徑 / Please modify to actual path
    sequence_fasta_path = Path(
        r"C:\Users\Andy\Desktop\Data\database\course\sequence.fasta"
    )  # 請修改為實際路徑 / Please modify to actual path
    taxdb_path = Path(
        r"C:\Users\Andy\Desktop\Data\database\course\taxdb.tar.gz"
    )  # 請修改為實際路徑 / Please modify to actual path
    title = "12S_Osteichthyes"  # 資料庫標題 / Database title
    output_path = Path(
        r"C:\Users\Andy\Desktop\Data\database"
    )  # 請修改為實際路徑 / Please modify to actual path

    logger.info("開始建置 BLAST 核酸資料庫")
    logger.info(f"accession2taxid_path: {accession2taxid_path}")
    logger.info(f"sequence_fasta_path: {sequence_fasta_path}")
    logger.info(f"taxdb_path: {taxdb_path}")
    logger.info(f"title: {title}")
    logger.info(f"output_path: {output_path}")

    # 檢查輸入檔案是否存在
    if not accession2taxid_path.exists():
        logger.error(f"accession2taxid_path 不存在: {accession2taxid_path}")
        return

    if not sequence_fasta_path.exists():
        logger.error(f"sequence_fasta_path 不存在: {sequence_fasta_path}")
        return

    if not taxdb_path.exists():
        logger.error(f"taxdb_path 不存在: {taxdb_path}")
        return

    # 建立輸出目錄
    output_path.mkdir(parents=True, exist_ok=True)
    temp_dir = output_path / "temp"
    temp_dir.mkdir(parents=True, exist_ok=True)

    try:
        # 步驟 1: 解壓縮
        logger.info("=== 步驟 1: 解壓縮 ===")
        # 判斷 accession2taxid_path 是檔案還是目錄
        if accession2taxid_path.is_file():
            nucl_gz_path = accession2taxid_path
        elif accession2taxid_path.is_dir():
            nucl_gz_path = accession2taxid_path / "nucl_gb.accession2taxid.gz"
        else:
            logger.error(f"accession2taxid_path 不存在或格式錯誤: {accession2taxid_path}")
            return

        if not nucl_gz_path.exists():
            logger.error(f"找不到 nucl_gb.accession2taxid.gz: {nucl_gz_path}")
            return

        nucl_accession2taxid_path = temp_dir / "nucl_gb.accession2taxid"
        extract_gzip_file(nucl_gz_path, nucl_accession2taxid_path)

        extract_tar_gz_file(taxdb_path, output_path)

        # 檢查解壓縮後的檔案
        expected_files = [
            "taxdb.btd",
            "taxdb.bti",
            "taxonomy4blast.sqlite3",
        ]
        for filename in expected_files:
            file_path = output_path / filename
            if not file_path.exists():
                logger.warning(f"預期的檔案不存在: {file_path}")

        # 步驟 2: Parsing
        logger.info("=== 步驟 2: Parsing ===")
        taxidmapfile_path = temp_dir / "taxidmapfile.txt"
        parse_accession2taxid(nucl_accession2taxid_path, taxidmapfile_path)

        # 步驟 3: 建立資料庫
        logger.info("=== 步驟 3: 建立資料庫 ===")
        makeblastdb_exe = get_project_root() / "dependencies" / "blast+" / "bin" / "makeblastdb.exe"

        if not makeblastdb_exe.exists():
            logger.error(f"找不到 makeblastdb.exe: {makeblastdb_exe}")
            return

        # makeblastdb 的 -out 參數是資料庫的基礎名稱(不含副檔名)
        db_name = "database"
        cmd = [
            str(makeblastdb_exe),
            "-in",
            str(sequence_fasta_path),
            "-dbtype",
            "nucl",
            "-title",
            title,
            "-parse_seqids",
            "-taxid_map",
            str(taxidmapfile_path),
            "-out",
            db_name,
        ]

        logger.info(f"執行命令: {' '.join(cmd)}")
        logger.info(f"執行目錄: {output_path}")
        run_command(cmd, cwd=output_path, check=True, capture_output=True)
        logger.info("資料庫建置命令執行完成")

        # 步驟 4: 確認建庫成功並刪除 temp
        logger.info("=== 步驟 4: 確認建庫成功 ===")
        db_output_path = output_path
        if check_database_files(db_output_path):
            logger.info("刪除臨時目錄")
            if temp_dir.exists():
                shutil.rmtree(temp_dir, ignore_errors=True)
                logger.info(f"已刪除臨時目錄: {temp_dir}")
        else:
            logger.error("資料庫建置失敗, 保留臨時目錄以供檢查")

        logger.info("BLAST 核酸資料庫建置流程完成")

    except Exception as e:
        logger.error(f"執行過程中發生錯誤: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()
