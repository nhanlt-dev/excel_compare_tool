import logging
import sys
import traceback
from pathlib import Path
from datetime import datetime

def setup_logger(name="excel_compare", log_level=logging.INFO):
    """Configure application-wide logger with file and console handlers."""
    logger = logging.getLogger(name)
    logger.setLevel(log_level)

    # Avoid duplicate handlers
    if logger.handlers:
        return logger

    # Formatter
    formatter = logging.Formatter(
        '%(asctime)s [%(levelname)s] %(name)s: %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Console handler
    console = logging.StreamHandler(sys.stdout)
    console.setFormatter(formatter)
    logger.addHandler(console)

    # File handler - create logs directory in project root
    try:
        project_root = Path(__file__).parent.parent.parent  # utils/ -> project root
        log_dir = project_root / "logs"
        log_dir.mkdir(exist_ok=True)

        log_file = log_dir / f"excel_compare_{datetime.now().strftime('%Y%m%d')}.log"
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setFormatter(formatter)
        logger.addHandler(file_handler)
    except Exception as e:
        # If we can't create log file, still have console logging
        console.handle(logging.LogRecord(
            name=name, level=logging.WARNING,
            pathname='', lineno=0,
            msg=f"Failed to create log file: {e}",
            args=(), exc_info=None
        ))

    return logger

def get_logger():
    """Get or create the application logger."""
    return logging.getLogger("excel_compare")

def log_exception(exc_type, exc_value, exc_traceback):
    """Log unhandled exceptions with full traceback."""
    logger = get_logger()
    logger.error("Unhandled exception:", exc_info=(exc_type, exc_value, exc_traceback))
