from conftest import create_test_excel_input
import pandas as pd
import numpy as np
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from datetime import date, datetime, timedelta
import os
import sys
import subprocess
import random
from dateutil.relativedelta import relativedelta
from matplotlib import pyplot as plt

# Path to the main script (assuming it's in the parent directory)
SCRIPT_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "move_tracker_report.py")
)





def test_progress_log_generation_basic(create_test_excel_input, tmp_path):
    """
    Tests if the Progress_Log sheet is correctly generated for a basic scenario.
    """
    input_excel_path, planned_start_date_fixture = create_test_excel_input
    # snapshot_date = planned_start_date_fixture + timedelta(
    #     days=15
    # )  # Example snapshot date
    snapshot_date = date(2025, 1, 21)
    snapshot_date_str = snapshot_date.strftime("%Y-%m-%d")

    command = [
        sys.executable,
        SCRIPT_PATH,
        "--excel-path",
        str(input_excel_path),
        "--snapshot-date",
        snapshot_date_str,
        "--log-level",
        "DEBUG",
        "--overwrite",
    ]

    result = subprocess.run(command, cwd=tmp_path, capture_output=True, text=True)

    assert (
        result.returncode == 0
    ), f"Script failed with error:\n{result.stderr}\n{result.stdout}"

    # Read the generated Progress_Log sheet
    df_progress_log = pd.read_excel(input_excel_path, sheet_name="Progress_Log")
    # Convert Snapshot_Date column to datetime.date objects to match expected_df
    df_progress_log["Snapshot_Date"] = df_progress_log["Snapshot_Date"].dt.date

    # --- EXPECTED PROGRESS_LOG DATA FOR THIS SCENARIO ---
    # Note: The expected data is based on the fixed historic and current data in the test input.
    expected_data = {
        "Snapshot_Date": [
            date(2025, 1, 1),
            date(2025, 1, 5),
            date(2025, 1, 6),
            date(2025, 1, 10),
            date(2025, 1, 11),
            date(2025, 1, 12),
            date(2025, 1, 13),
            date(2025, 1, 14),
            date(2025, 1, 17),
            date(2025, 1, 18),
            date(2025, 1, 21),
        ],
        "Scope_At_Snapshot": [4, 4, 4, 4, 4, 6, 5, 5, 5, 4, 4],
        "Actual_Work_Completed": [0, 1, 1, 2, 2, 2, 2, 2, 2, 3, 4],
        "Elapsed_Time_Days": [1, 5, 6, 10, 11, 12, 13, 14, 17, 18, 21],
        "Actual_Operational_Throughput": [
            0.0,
            0.2,
            0.16666666666666666,
            0.2,
            0.18181818181818182,
            0.16666666666666666,
            0.15384615384615385,
            0.142857142857143,
            0.117647058823529,
            0.166666666666667,
            0.19047619047619,
        ],
        "Current_50th_Percentile_Flow_Time": [
            4.0,
            4.0,
            4.0,
            5.0,
            5.0,
            5.0,
            5.0,
            5.0,
            5.0,
            5.0,
            5.0,
        ],
        "Forecasted_Delivery_Date": [
            None,
            date(2025, 1, 18),
            date(2025, 1, 20),
            date(2025, 1, 20),
            date(2025, 1, 21),
            date(2025, 2, 2),
            date(2025, 1, 29),
            date(2025, 1, 31),
            date(2025, 2, 5),
            date(2025, 1, 27),
            date(2025, 1, 25),
        ],
        "Buffer_Consumption_Percentage": [
            0,
            0.111111111,
            0.333333333,
            0.333333333,
            0.444444444,
            1.777777778,
            1.333333333,
            1.555555556,
            2.111111111,
            1.111111111,
            0.888888889,
        ],
        "Work_Done_Percentage": [
            0,
            0.25,
            0.25,
            0.5,
            0.5,
            0.333333333,
            0.4,
            0.4,
            0.4,
            0.75,
            1,
        ],
        "Fever_chart_signal": [
            "Green",
            "Green",
            "Yellow",
            "Green",
            "Yellow",
            "Beyond Red",
            "Beyond Red",
            "Beyond Red",
            "Beyond Red",
            "Beyond Red",
            "Red",
        ],
    }
    expected_df = pd.DataFrame(expected_data)
    expected_df["Snapshot_Date"] = pd.to_datetime(expected_df["Snapshot_Date"]).dt.date
    # Convert Forecasted_Delivery_Date in df_progress_log to datetime.date, handling NaT
    df_progress_log["Forecasted_Delivery_Date"] = df_progress_log[
        "Forecasted_Delivery_Date"
    ].apply(lambda x: x.date() if pd.notna(x) else pd.NaT)

    expected_df["Forecasted_Delivery_Date"] = expected_df[
        "Forecasted_Delivery_Date"
    ].apply(
        lambda x: x
        if pd.notna(x)
        else pd.NaT  # Ensure expected_df also has NaT for comparison
    )

    pd.testing.assert_frame_equal(df_progress_log, expected_df, check_dtype=False)

    # Verify chart files are created
    work_exec_chart_path = tmp_path / f"{snapshot_date_str}_work_execution_chart.png"
    fever_chart_path = tmp_path / f"{snapshot_date_str}_fever_chart.png"

    # assert work_exec_chart_path.exists()
    # assert work_exec_chart_path.stat().st_size > 0

    # assert fever_chart_path.exists()
    # assert fever_chart_path.stat().st_size > 0

    # Optional: Verify images are embedded in the Excel file
    wb_output = load_workbook(input_excel_path)
    assert "Work_Execution_Chart" in wb_output.sheetnames
    assert "Fever_Chart" in wb_output.sheetnames
    assert len(wb_output["Work_Execution_Chart"]._images) > 0
    assert len(wb_output["Fever_Chart"]._images) > 0
