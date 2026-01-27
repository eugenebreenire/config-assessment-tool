import pytest
import shutil
from pathlib import Path
from compare_tool.powerpoint.apm import generate_powerpoint_from_apm

def test_generate_powerpoint_from_apm(tmp_path):
    # Setup
    # Path to the existing template
    valid_template_path = Path(__file__).parent.parent / "templates" / "template.pptx"

    # Copy the valid template to the temporary directory
    template_path = tmp_path / "template.pptx"
    shutil.copy(valid_template_path, template_path)

    # Create mock input files in the temporary directory
    comparison_result_path = tmp_path / "comparison.json"
    comparison_result_path.write_text("{}")  # Simulate an empty JSON file

    current_file_path = tmp_path / "current.xlsx"
    current_file_path.touch()  # Simulate an empty Excel file

    previous_file_path = tmp_path / "previous.xlsx"
    previous_file_path.touch()  # Simulate an empty Excel file

    powerpoint_output_path = tmp_path / "output.pptx"

    # Act
    generate_powerpoint_from_apm(
        comparison_result_path=str(comparison_result_path),
        powerpoint_output_path=str(powerpoint_output_path),
        current_file_path=str(current_file_path),
        previous_file_path=str(previous_file_path),
        template_path=str(template_path),
    )

    # Assert
    assert powerpoint_output_path.exists(), "PowerPoint output file was not created."