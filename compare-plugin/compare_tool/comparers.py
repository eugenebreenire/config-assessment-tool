# compare_tool/comparers.py

"""
Domain-aware comparison dispatcher.

Keeps your old API:
    compare_files_other_sheets(prev, curr, out)

…while also supporting:
    compare_files_other_sheets(prev, curr, out, domain="BRUM")
    compare_files_other_sheets(prev, curr, out, domain="MRUM")
"""

import logging
from typing import Optional

from .comparers_apm import compare_files_other_sheets_apm
from .comparers_brum import compare_files_other_sheets_brum
from .comparers_mrum import compare_files_other_sheets_mrum

logger = logging.getLogger(__name__)

__all__ = [
    "compare_files_other_sheets",
    "compare_files_other_sheets_apm",
    "compare_files_other_sheets_brum",
    "compare_files_other_sheets_mrum",
]


def compare_files_other_sheets(
    previous_file_path: str,
    current_file_path: str,
    output_file_path: str,
    domain: Optional[str] = "APM",
) -> None:
    """
    Unified dispatcher for APM / BRUM / MRUM.

    - Backwards compatible for existing APM code:
        compare_files_other_sheets(prev, curr, out)

    - Extended usage for BRUM / MRUM:
        compare_files_other_sheets(prev, curr, out, domain="BRUM")
        compare_files_other_sheets(prev, curr, out, domain="MRUM")
    """
    dom = (domain or "APM").upper()

    if dom == "APM":
        return compare_files_other_sheets_apm(
            previous_file_path, current_file_path, output_file_path
        )
    elif dom == "BRUM":
        return compare_files_other_sheets_brum(
            previous_file_path, current_file_path, output_file_path
        )
    elif dom == "MRUM":
        return compare_files_other_sheets_mrum(
            previous_file_path, current_file_path, output_file_path
        )
    else:
        logger.warning(
            "Unknown domain '%s' passed to compare_files_other_sheets; defaulting to APM.",
            dom,
        )
        return compare_files_other_sheets_apm(
            previous_file_path, current_file_path, output_file_path
        )
