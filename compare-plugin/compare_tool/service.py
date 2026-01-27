# compare_tool/service.py

import os
import logging
from typing import Dict, Tuple, Optional, Any
from pathlib import Path

from .excel_io import save_workbook, check_controllers_match
from .summary import (
    create_summary_workbooks,
    compare_files_summary,
    copy_summary_to_result,
)
from .comparers import compare_files_other_sheets
from .insights import build_comparison_json

from compare_tool.powerpoint.apm import generate_powerpoint_from_apm as generate_apm_ppt
from compare_tool.powerpoint.brum import generate_powerpoint_from_brum
from compare_tool.powerpoint.mrum import generate_powerpoint_from_mrum



logger = logging.getLogger(__name__)

# Base directory of the project (points at compare-plugin root)
BASE_DIR = Path(__file__).resolve().parent.parent


def _resolve_template_path(config: Dict, domain_key: str, default_name: str) -> Optional[str]:
    """
    Build an absolute template path from config.json settings.

    domain_key examples:
      - "apm_template_file"
      - "brum_template_file"
      - "mrum_template_file"

    We expect config.json to contain something like:
      "TEMPLATE_FOLDER": "templates",
      "apm_template_file": "template.pptx",
      "brum_template_file": "template_brum.pptx",
      "mrum_template_file": "template_mrum.pptx"
    """
    folder_name = config.get("TEMPLATE_FOLDER") or config.get("template_folder")
    if not folder_name:
        logger.warning("No TEMPLATE_FOLDER/template_folder defined in config.json")
        return None

    filename = config.get(domain_key, default_name)
    template_path = BASE_DIR / folder_name / filename
    if template_path.exists():
        return str(template_path)

    logger.warning("Template not found at %s", template_path)
    return None


# ---------------------------------------------------------------------------
# APM
# ---------------------------------------------------------------------------
def run_comparison(
    previous_file_path: str,
    current_file_path: str,
    config: Dict,
) -> Tuple[str, str]:
    """
    High-level comparison pipeline for APM.
    Returns (output_file_path, powerpoint_output_path).
    """

    upload_folder = config["upload_folder"]
    result_folder = config["result_folder"]

    # Resolve APM template path using config + BASE_DIR
    template_path = _resolve_template_path(
        config=config,
        domain_key="apm_template_file",
        default_name="template.pptx",
    )

    os.makedirs(upload_folder, exist_ok=True)
    os.makedirs(result_folder, exist_ok=True)

    # Use names from config.json where possible
    output_file_name = config.get("output_file", "comparison_result.xlsx")
    previous_sum_name = config.get("previous_sum_file", "previous_sum.xlsx")
    current_sum_name = config.get("current_sum_file", "current_sum.xlsx")
    comparison_sum_name = config.get("comparison_sum_file", "comparison_sum.xlsx")

    output_file_path = os.path.join(result_folder, output_file_name)
    previous_sum_path = os.path.join(upload_folder, previous_sum_name)
    current_sum_path = os.path.join(upload_folder, current_sum_name)
    comparison_sum_path = os.path.join(result_folder, comparison_sum_name)

    powerpoint_output_path = os.path.join(result_folder, "Analysis_Summary_APM.pptx")

    # 1. Recalculate formulas in both input workbooks
    save_workbook(previous_file_path)
    save_workbook(current_file_path)

    # 2. Check controllers
    if not check_controllers_match(previous_file_path, current_file_path):
        raise ValueError("Controllers do not match between previous and current files.")

    # 3. Summary extraction & comparison
    create_summary_workbooks(
        previous_file_path, current_file_path, previous_sum_path, current_sum_path
    )
    compare_files_summary(previous_sum_path, current_sum_path, comparison_sum_path)

    # 4. Per-sheet comparisons -> main comparison_result.xlsx (APM domain)
    compare_files_other_sheets(
        previous_file_path,
        current_file_path,
        output_file_path,
        domain="APM",
    )

    # 5. Copy final summary into result workbook
    copy_summary_to_result(comparison_sum_path, output_file_path)

    # 6. PowerPoint (APM-specific generator)
    generate_apm_ppt(
        comparison_result_path=output_file_path,
        powerpoint_output_path=powerpoint_output_path,
        current_file_path=current_file_path,
        previous_file_path=previous_file_path,
        template_path=template_path,
        domain="APM",
        config=config,
    )

    # 7. Insights JSON (APM)
    try:
        build_comparison_json(
            domain="APM",
            comparison_result_path=output_file_path,
            current_file_path=current_file_path,
            previous_file_path=previous_file_path,
            result_folder=result_folder,
            meta={"domain": "APM"},
        )
    except Exception as e:
        logger.warning("Failed to build APM Insights JSON: %s", e, exc_info=True)

    logger.info("APM comparison pipeline completed successfully.")
    return output_file_path, powerpoint_output_path


# ---------------------------------------------------------------------------
# BRUM
# ---------------------------------------------------------------------------
def run_comparison_brum(
    previous_file_path: str,
    current_file_path: str,
    config: Dict,
) -> Tuple[str, str]:
    """
    BRUM comparison pipeline.
    Uses BRUM-specific template + filenames and BRUM comparers.
    """

    upload_folder = config["upload_folder"]
    result_folder = config["result_folder"]

    os.makedirs(upload_folder, exist_ok=True)
    os.makedirs(result_folder, exist_ok=True)

    output_file_name = config.get("output_file_brum", "comparison_result_brum.xlsx")
    previous_sum_name = config.get("previous_sum_file_brum", "previous_sum_brum.xlsx")
    current_sum_name = config.get("current_sum_file_brum", "current_sum_brum.xlsx")
    comparison_sum_name = config.get(
        "comparison_sum_file_brum", "comparison_sum_brum.xlsx"
    )

    output_file_path = os.path.join(result_folder, output_file_name)
    previous_sum_path = os.path.join(upload_folder, previous_sum_name)
    current_sum_path = os.path.join(upload_folder, current_sum_name)
    comparison_sum_path = os.path.join(result_folder, comparison_sum_name)
    powerpoint_output_path = os.path.join(result_folder, "Analysis_Summary_BRUM.pptx")

    # 1. Recalculate formulas
    save_workbook(previous_file_path)
    save_workbook(current_file_path)

    # 2. Controllers must match
    if not check_controllers_match(previous_file_path, current_file_path):
        raise ValueError(
            "Controllers do not match between previous and current files (BRUM)."
        )

    # 3. Summary extraction & comparison
    create_summary_workbooks(
        previous_file_path, current_file_path, previous_sum_path, current_sum_path
    )
    compare_files_summary(previous_sum_path, current_sum_path, comparison_sum_path)

    # 4. Per-sheet comparisons (BRUM domain)
    compare_files_other_sheets(
        previous_file_path,
        current_file_path,
        output_file_path,
        domain="BRUM",
    )

    # 5. Copy summary into result workbook
    copy_summary_to_result(comparison_sum_path, output_file_path)

    # 6. PowerPoint – now use BRUM-specific generator
    generate_powerpoint_from_brum(
        comparison_result_path=output_file_path,
        powerpoint_output_path=powerpoint_output_path,
        current_file_path=current_file_path,
        previous_file_path=previous_file_path,
        config=config,
    )

    # 7. Insights JSON (BRUM)
    try:
        build_comparison_json(
            domain="BRUM",
            comparison_result_path=output_file_path,
            current_file_path=current_file_path,
            previous_file_path=previous_file_path,
            result_folder=result_folder,
            meta={"domain": "BRUM"},
        )
    except Exception as e:
        logger.warning("Failed to build BRUM Insights JSON: %s", e, exc_info=True)

    logger.info("BRUM comparison pipeline completed successfully.")
    return output_file_path, powerpoint_output_path


# ---------------------------------------------------------------------------
# MRUM
# ---------------------------------------------------------------------------
def run_comparison_mrum(
    previous_file_path: str,
    current_file_path: str,
    config: Dict,
) -> Tuple[str, str]:
    """
    MRUM comparison pipeline.
    Uses MRUM-specific template + filenames and MRUM comparers.
    """

    upload_folder = config["upload_folder"]
    result_folder = config["result_folder"]

    os.makedirs(upload_folder, exist_ok=True)
    os.makedirs(result_folder, exist_ok=True)

    output_file_name = config.get("output_file_mrum", "comparison_result_mrum.xlsx")
    previous_sum_name = config.get("previous_sum_file_mrum", "previous_sum_mrum.xlsx")
    current_sum_name = config.get("current_sum_file_mrum", "current_sum_mrum.xlsx")
    comparison_sum_name = config.get(
        "comparison_sum_file_mrum", "comparison_sum_mrum.xlsx"
    )

    output_file_path = os.path.join(result_folder, output_file_name)
    previous_sum_path = os.path.join(upload_folder, previous_sum_name)
    current_sum_path = os.path.join(upload_folder, current_sum_name)
    comparison_sum_path = os.path.join(result_folder, comparison_sum_name)
    powerpoint_output_path = os.path.join(result_folder, "Analysis_Summary_MRUM.pptx")

    # 1. Recalculate formulas
    save_workbook(previous_file_path)
    save_workbook(current_file_path)

    # 2. Controllers must match
    if not check_controllers_match(previous_file_path, current_file_path):
        raise ValueError(
            "Controllers do not match between previous and current files (MRUM)."
        )

    # 3. Summary extraction & comparison
    create_summary_workbooks(
        previous_file_path, current_file_path, previous_sum_path, current_sum_path
    )
    compare_files_summary(previous_sum_path, current_sum_path, comparison_sum_path)

    # 4. Per-sheet comparisons (MRUM domain)
    compare_files_other_sheets(
        previous_file_path,
        current_file_path,
        output_file_path,
        domain="MRUM",
    )

    # 5. Copy summary into result workbook
    copy_summary_to_result(comparison_sum_path, output_file_path)

    # 6. PowerPoint – MRUM-specific generator
    generate_powerpoint_from_mrum(
        comparison_result_path=output_file_path,
        powerpoint_output_path=powerpoint_output_path,
        current_file_path=current_file_path,
        previous_file_path=previous_file_path,
        config=config,
    )

    # 7. Insights JSON (MRUM)
    try:
        build_comparison_json(
            domain="MRUM",
            comparison_result_path=output_file_path,
            current_file_path=current_file_path,
            previous_file_path=previous_file_path,
            result_folder=result_folder,
            meta={"domain": "MRUM"},
        )
    except Exception as e:
        logger.warning("Failed to build MRUM Insights JSON: %s", e, exc_info=True)

    logger.info("MRUM comparison pipeline completed successfully.")
    return output_file_path, powerpoint_output_path

