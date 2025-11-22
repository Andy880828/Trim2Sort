"""日誌工具模組

Logger utility module.
"""

import logging
from pathlib import Path
import sys
from typing import ClassVar, Literal


class ColoredFormatter(logging.Formatter):
    """彩色日誌格式化器

    Colored log formatter for terminal output.
    """

    COLORS: ClassVar[dict[str, str]] = {
        "DEBUG": "\033[36m",
        "INFO": "\033[32m",
        "WARNING": "\033[33m",
        "ERROR": "\033[31m",
        "CRITICAL": "\033[35m",
    }
    RESET: ClassVar[str] = "\033[0m"

    def format(self, record: logging.LogRecord) -> str:
        """格式化日誌記錄

        Format log record with colors.

        Args:
            record (logging.LogRecord): 日誌記錄 / Log record.

        Returns:
            str: 格式化後的日誌訊息 / Formatted log message.
        """
        log_color = self.COLORS.get(record.levelname, "")
        levelname = record.levelname
        record.levelname = f"{log_color}{levelname}{self.RESET}"
        result = super().format(record)
        record.levelname = levelname
        return result


def init_logger(
    level: Literal["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"] = "INFO",
    log_file: Path | str | None = None,
    use_color: bool = True,
) -> None:
    """初始化根 Logger

    Initialize root logger (should be called once at application startup).

    Args:
        level (Literal["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]): 日誌級別 / Log level.
        log_file (Path | str | None): 日誌檔案路徑, None 表示不寫入檔案 / Log file path, None means no file output.
        use_color (bool): 是否在終端使用彩色輸出 / Whether to use colored output in terminal.
    """
    root_logger = logging.getLogger()
    root_logger.setLevel(getattr(logging, level))

    if root_logger.handlers:
        root_logger.handlers.clear()

    detailed_format = (
        "%(asctime)s | %(levelname)-8s | %(filename)s:%(lineno)d | %(funcName)s() | %(message)s"
    )
    simple_format = "%(asctime)s | %(levelname)-8s | %(filename)s:%(lineno)d | %(message)s"
    date_format = "%Y-%m-%d %H:%M:%S"

    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setLevel(getattr(logging, level))

    if use_color and sys.stdout.isatty():
        console_formatter = ColoredFormatter(detailed_format, datefmt=date_format)
    else:
        console_formatter = logging.Formatter(detailed_format, datefmt=date_format)

    console_handler.setFormatter(console_formatter)
    root_logger.addHandler(console_handler)

    if log_file:
        log_path = Path(log_file)
        log_path.parent.mkdir(parents=True, exist_ok=True)
        file_handler = logging.FileHandler(log_path, encoding="utf-8", mode="a")
        file_handler.setLevel(getattr(logging, level))
        file_formatter = logging.Formatter(simple_format, datefmt=date_format)
        file_handler.setFormatter(file_formatter)
        root_logger.addHandler(file_handler)


def get_logger(name: str) -> logging.Logger:
    """取得指定名稱的 Logger

    Get a logger with the specified name.

    Args:
        name (str): Logger 名稱, 通常使用 __name__ / Logger name, usually __name__.

    Returns:
        logging.Logger: Logger 實例 / Logger instance.
    """
    return logging.getLogger(name)
