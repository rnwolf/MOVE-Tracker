import pytest
import os
import sys
import subprocess
from datetime import date
import numpy as np
import pandas as pd
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


def get_snapshot_dates_from_progress_log(excel_path):
    """
    Extract all snapshot dates from the Progress_Log sheet after running the script.
    Returns a list of date strings in YYYY-MM-DD format.
    """
    try:
        # First, run the script to generate the full Progress_Log
        temp_snapshot = "2025-01-21"  # Use a date that should capture all events
        command = [
            sys.executable,
            SCRIPT_PATH,
            "--excel-path",
            str(excel_path),
            "--snapshot-date",
            temp_snapshot,
            "--log-level",
            "ERROR",  # Minimize output
            "--no-chart-insertion",  # Don't generate charts yet
        ]
        
        result = subprocess.run(command, capture_output=True, text=True)
        if result.returncode != 0:
            print(f"Warning: Failed to generate Progress_Log: {result.stderr}")
            return []
        
        # Read the Progress_Log sheet
        df_progress = pd.read_excel(excel_path, sheet_name="Progress_Log")
        
        # Convert Snapshot_Date to string format and return unique dates
        snapshot_dates = df_progress["Snapshot_Date"].dt.strftime("%Y-%m-%d").unique().tolist()
        return sorted(snapshot_dates)
        
    except Exception as e:
        print(f"Error extracting snapshot dates: {e}")
        return []


def generate_all_reference_charts(excel_input_path, test_scenario_name="basic_scenario"):
    """
    Generate reference charts for all snapshot dates in the Progress_Log.
    This function should be run manually when you want to create/update reference images.
    """
    import tempfile
    from pathlib import Path
    
    print(f"Generating reference charts for scenario: {test_scenario_name}")
    
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        
        # Get all snapshot dates from the Progress_Log
        snapshot_dates = get_snapshot_dates_from_progress_log(excel_input_path)
        
        if not snapshot_dates:
            print("No snapshot dates found. Cannot generate reference charts.")
            return False
        
        print(f"Found {len(snapshot_dates)} snapshot dates: {snapshot_dates}")
        
        success_count = 0
        for snapshot_date_str in snapshot_dates:
            print(f"\nGenerating reference charts for {snapshot_date_str}...")
            
            # Generate charts for this snapshot date
            command = [
                sys.executable,
                SCRIPT_PATH,
                "--excel-path",
                str(excel_input_path),
                "--snapshot-date",
                snapshot_date_str,
                "--log-level",
                "ERROR",  # Minimize output
                "--save-charts-only",
                str(tmp_path),
            ]
            
            result = subprocess.run(command, capture_output=True, text=True)
            
            if result.returncode == 0:
                # Move generated charts to reference directory
                work_exec_chart = tmp_path / f"{snapshot_date_str}_work_execution_chart.png"
                fever_chart = tmp_path / f"{snapshot_date_str}_fever_chart.png"
                
                ref_work_exec = os.path.join(
                    REFERENCE_IMAGES_DIR,
                    f"{test_scenario_name}_{snapshot_date_str}_work_execution_chart.png"
                )
                ref_fever = os.path.join(
                    REFERENCE_IMAGES_DIR,
                    f"{test_scenario_name}_{snapshot_date_str}_fever_chart.png"
                )
                
                if work_exec_chart.exists():
                    Image.open(work_exec_chart).save(ref_work_exec)
                    print(f"  [OK] Work Execution Chart: {ref_work_exec}")
                else:
                    print(f"  [FAIL] Work Execution Chart not found: {work_exec_chart}")
                
                if fever_chart.exists():
                    Image.open(fever_chart).save(ref_fever)
                    print(f"  [OK] Fever Chart: {ref_fever}")
                else:
                    print(f"  [FAIL] Fever Chart not found: {fever_chart}")
                
                success_count += 1
            else:
                print(f"  [FAIL] Failed to generate charts: {result.stderr}")
        
        print(f"\nGenerated reference charts for {success_count}/{len(snapshot_dates)} snapshot dates")
        return success_count == len(snapshot_dates)


def test_generate_all_reference_charts(create_test_excel_input):
    """
    Test function to generate all reference charts for the basic scenario.
    Run this test manually when you want to create/update all reference images.
    
    Usage: pytest tests/test_chart_visuals_enhanced.py::test_generate_all_reference_charts -s
    """
    excel_input_path, _ = create_test_excel_input
    success = generate_all_reference_charts(excel_input_path, "basic_scenario")
    assert success, "Failed to generate all reference charts"


@pytest.mark.parametrize("test_scenario_name", ["basic_scenario"])
def test_comprehensive_chart_visual_consistency(create_test_excel_input, tmp_path, test_scenario_name):
    """
    Tests the visual consistency of generated charts against reference images for ALL snapshot dates.
    This test automatically discovers all snapshot dates from the Progress_Log and tests each one.
    
    Usage: pytest tests/test_chart_visuals_enhanced.py::test_comprehensive_chart_visual_consistency -s
    """
    excel_input_path, _ = create_test_excel_input
    
    # Get all snapshot dates from the Progress_Log
    snapshot_dates = get_snapshot_dates_from_progress_log(excel_input_path)
    
    if not snapshot_dates:
        pytest.skip("No snapshot dates found in Progress_Log")
    
    print(f"Testing {len(snapshot_dates)} snapshot dates: {snapshot_dates}")
    
    failed_dates = []
    
    for snapshot_date_str in snapshot_dates:
        print(f"\nTesting charts for {snapshot_date_str}...")
        
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

        # Run the script to generate charts
        command = [
            sys.executable,
            SCRIPT_PATH,
            "--excel-path",
            str(excel_input_path),
            "--snapshot-date",
            snapshot_date_str,
            "--log-level",
            "ERROR",  # Minimize output
            "--save-charts-only",
            str(tmp_path),
        ]
        result = subprocess.run(command, capture_output=True, text=True)

        if result.returncode != 0:
            failed_dates.append(f"{snapshot_date_str}: Script failed - {result.stderr}")
            continue

        # Check Work Execution Chart
        if not os.path.exists(reference_we_chart_path):
            print(f"  WARNING: Reference Work Execution Chart not found: {reference_we_chart_path}")
            # Auto-generate reference if it doesn't exist
            if generated_we_chart_path.exists():
                Image.open(generated_we_chart_path).save(reference_we_chart_path)
                print(f"  Generated reference: {reference_we_chart_path}")
            else:
                failed_dates.append(f"{snapshot_date_str}: Generated Work Execution Chart not found")
                continue
        else:
            # Compare with existing reference
            diff_we_path = tmp_path / f"diff_{test_scenario_name}_{snapshot_date_str}_work_execution_chart.png"
            if not compare_images(generated_we_chart_path, reference_we_chart_path, diff_we_path):
                failed_dates.append(f"{snapshot_date_str}: Work Execution Chart visual mismatch")

        # Check Fever Chart
        if not os.path.exists(reference_fever_chart_path):
            print(f"  WARNING: Reference Fever Chart not found: {reference_fever_chart_path}")
            # Auto-generate reference if it doesn't exist
            if generated_fever_chart_path.exists():
                Image.open(generated_fever_chart_path).save(reference_fever_chart_path)
                print(f"  Generated reference: {reference_fever_chart_path}")
            else:
                failed_dates.append(f"{snapshot_date_str}: Generated Fever Chart not found")
                continue
        else:
            # Compare with existing reference
            diff_fever_path = tmp_path / f"diff_{test_scenario_name}_{snapshot_date_str}_fever_chart.png"
            if not compare_images(generated_fever_chart_path, reference_fever_chart_path, diff_fever_path):
                failed_dates.append(f"{snapshot_date_str}: Fever Chart visual mismatch")
    
    # Report results
    if failed_dates:
        failure_msg = f"Chart visual consistency failed for {len(failed_dates)} dates:\n" + "\n".join(failed_dates)
        pytest.fail(failure_msg)
    else:
        print(f"\n[SUCCESS] All {len(snapshot_dates)} snapshot dates passed visual consistency tests!")


# Convenience function for manual testing
if __name__ == "__main__":
    import tempfile
    from pathlib import Path
    
    print("Generating all reference charts...")
    
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        
        # Create test Excel input manually (since we can't use fixtures here)
        excel_input_path, _ = create_test_excel_input(tmp_path)
        success = generate_all_reference_charts(excel_input_path, "basic_scenario")
        
        if success:
            print("All reference charts generated successfully!")
        else:
            print("Failed to generate some reference charts.")