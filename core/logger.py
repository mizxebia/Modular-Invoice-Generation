import logging

from core.config import config

log_cfg = config["logging"]

logging.basicConfig(
    filename=log_cfg["log_file"],
    level=getattr(logging, log_cfg["log_level"]),
    format="%(asctime)s | %(levelname)s | %(message)s"
)

logger = logging.getLogger(__name__)
