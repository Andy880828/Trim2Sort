"""子進程執行工具模組

Subprocess execution utility module.
"""

import os
from pathlib import Path
import subprocess

from src.utils.logger_utils import get_logger

logger = get_logger(__name__)


def run_command(
    cmd: list[str] | str,
    cwd: Path | str | None = None,
    check: bool = True,
    capture_output: bool = True,
) -> subprocess.CompletedProcess[str]:
    """執行外部命令 (處理編碼問題)

    Execute external command with encoding handling.

    Args:
        cmd (list[str] | str): 命令列表或字串 / Command list or string.
        cwd (Path | str | None): 工作目錄 / Working directory.
        check (bool): 是否在命令失敗時拋出異常 / Whether to raise exception on command failure.
        capture_output (bool): 是否捕獲輸出 / Whether to capture output.

    Returns:
        subprocess.CompletedProcess[str]: 命令執行結果 / Command execution result.

    Raises:
        subprocess.CalledProcessError: 當 check=True 且命令失敗時 / When check=True and command fails.
    """
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    if isinstance(cwd, Path):
        cwd = str(cwd)

    try:
        result = subprocess.run(
            cmd,
            check=check,
            capture_output=capture_output,
            text=True,
            encoding="utf-8",
            errors="replace",
            cwd=cwd,
            env=env,
        )
        return result
    except subprocess.CalledProcessError as e:
        error_msg = e.stderr if hasattr(e, "stderr") and e.stderr else str(e)
        cmd_str = " ".join(cmd) if isinstance(cmd, list) else str(cmd)
        logger.error(f"命令執行失敗: {cmd_str}")
        if error_msg:
            logger.error(f"錯誤訊息: {error_msg}")
        raise
    except Exception as e:
        logger.error(f"執行命令時發生未預期錯誤: {e}")
        raise
