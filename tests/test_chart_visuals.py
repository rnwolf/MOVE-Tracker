import pytest
import os
import sys
import subprocess
from datetime import date
import numpy as np
from PIL import Image, ImageChops  # Pillow for image comparison
from conftest import create_test_excel_input

# Path to the main script
SCRIPT_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "move_tracker_report.py")
)
REFERENCE_IMAGES_DIR = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "reference_images")
)

# Ensure the reference images directory exists
os.makedirs(REFERENCE_IMAGES_DIR, exist_ok=True)


def compare_images(img1_path, img2_path, diff_output_path=None, threshold=0.01):
    """
    Comparisons two images and returns True if they are similar enough, False otherwise.
    Optionally saves a diff image.
    """
    try:
        img1 = Image.open(img1_path).convert("RGB")
        img2 = Image.open(img2_path).convert("RGB")

        if img1.size != img2.size:
            print(f"Image sizes differ: {img1.size} vs {img2.size}")
            return False

        diff = ImageChops.difference(img1, img2)
        diff_array = np.array(diff)
        # Calculate the sum of absolute differences across all channels
        total_diff = np.sum(np.abs(diff_array))
        max_diff = (
            np.prod(img1.size) * 255 * 3
        )  # Max possible diff (width * height * 255 * channels)
        normalized_diff = total_diff / max_diff

        if diff_output_path:
            # Create a visual diff image
            diff.save(diff_output_path)

        print(
            f"Normalized image difference: {normalized_diff:.4f} (Threshold: {threshold})"
        )
        return normalized_diff <= threshold

    except FileNotFoundError as e:
        print(f"Error: Image file not found - {e}")
        return False
    except Exception as e:
        print(f"An error occurred during image comparison: {e}")
        return False


@pytest.mark.parametrize(
    "snapshot_date_str, test_scenario_name",
    [
        ("2025-01-15", "basic_scenario"),
        # Add more scenarios as needed
        # ("2025-01-21", "withdrawal_and_scope_change_scenario"),
    ],
)
def test_chart_visual_consistency(create_test_excel_input, tmp_path, snapshot_date_str, test_scenario_name):
    """
    Tests the visual consistency of generated charts against reference images.
    """
    # Define paths for generated and reference charts
    generated_we_chart_path = tmp_path / f"{snapshot_date_str}_work_execution_chart.png"
    generated_fever_chart_path = tmp_path / f"{snapshot_date_str}_fever_chart.png"

    reference_we_chart_path = os.path.join(
        REFERENCE_IMAGES_DIR,
        f"{test_scenario_name}_{snapshot_date_str}_work_execution_chart.png",
    )
    reference_fever_chart_path = os.path.join(
        REFERENCE_IMAGES_DIR,
        f"{test_scenario_name}_{snapshot_date_str}_fever_chart.png",
    )

    excel_input_path, _ = create_test_excel_input

    # Run the script to generate charts
    command = [
        sys.executable,
        SCRIPT_PATH,
        "--excel-path",
        str(excel_input_path),  # Use the proper test fixture Excel file
        "--snapshot-date",
        snapshot_date_str,
        "--log-level",
        "INFO",
        "--save-charts-only",
        str(tmp_path),  # Use new option and save to tmp_path
    ]
    result = subprocess.run(command, cwd=tmp_path, capture_output=True, text=True)

    assert (
        result.returncode == 0
    ), f"Script failed with error:\n{result.stderr}\n{result.stdout}"

    # --- Work Execution Chart Comparison ---
    if not os.path.exists(reference_we_chart_path):
        print(
            f"WARNING: Reference image for Work Execution Chart not found: {reference_we_chart_path}. Generating it."
        )
        # If reference doesn't exist, create it from the generated one
        Image.open(generated_we_chart_path).save(reference_we_chart_path)
        # Fail the test on first run to ensure manual review
        pytest.fail(
            f"Reference image created for {test_scenario_name} Work Execution Chart. Please review and commit."
        )

    diff_we_path = (
        tmp_path
        / f"diff_{test_scenario_name}_{snapshot_date_str}_work_execution_chart.png"
    )
    assert compare_images(
        generated_we_chart_path, reference_we_chart_path, diff_we_path
    ), f"Work Execution Chart visual mismatch for {test_scenario_name} on {snapshot_date_str}. Diff saved to {diff_we_path}"

    # --- Fever Chart Comparison ---
    if not os.path.exists(reference_fever_chart_path):
        print(
            f"WARNING: Reference image for Fever Chart not found: {reference_fever_chart_path}. Generating it."
        )
        Image.open(generated_fever_chart_path).save(reference_fever_chart_path)
        pytest.fail(
            f"Reference image created for {test_scenario_name} Fever Chart. Please review and commit."
        )

    diff_fever_path = (
        tmp_path / f"diff_{test_scenario_name}_{snapshot_date_str}_fever_chart.png"
    )
    assert compare_images(
        generated_fever_chart_path, reference_fever_chart_path, diff_fever_path
    ), f"Fever Chart visual mismatch for {test_scenario_name} on {snapshot_date_str}. Diff saved to {diff_fever_path}"
