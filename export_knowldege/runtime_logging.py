"""运行日志配置：同时输出到终端和 UTF-8 日志文件。"""

from __future__ import annotations

import logging
import sys
from datetime import datetime
from pathlib import Path


def configure_logging(log_dir: Path, logger_name: str, *additional_logger_names: str) -> tuple[logging.Logger, Path]:
    """为一次运行创建独立日志文件，并返回 logger 与日志路径。"""
    try:
        log_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        raise RuntimeError(f"无法创建日志目录: {log_dir} ({exc})") from exc
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_path = log_dir / f"{logger_name}_{timestamp}.log"
    suffix = 1
    while log_path.exists():
        log_path = log_dir / f"{logger_name}_{timestamp}_{suffix}.log"
        suffix += 1

    logger_names = tuple(dict.fromkeys((logger_name, *additional_logger_names)))
    loggers = [logging.getLogger(name) for name in logger_names]
    for configured_logger in loggers:
        configured_logger.setLevel(logging.INFO)
        configured_logger.propagate = False
        for handler in configured_logger.handlers[:]:
            handler.close()
            configured_logger.removeHandler(handler)

    formatter = logging.Formatter(
        fmt="%(asctime)s | %(levelname)s | %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
    )
    try:
        file_handler = logging.FileHandler(log_path, encoding="utf-8")
    except OSError as exc:
        raise RuntimeError(f"无法创建日志文件: {log_path} ({exc})") from exc
    file_handler.setFormatter(formatter)
    stream_handler = logging.StreamHandler(sys.stdout)
    stream_handler.setFormatter(formatter)
    for configured_logger in loggers:
        configured_logger.addHandler(file_handler)
        configured_logger.addHandler(stream_handler)
    loggers[0].info("日志开始记录: %s", log_path)
    return loggers[0], log_path
