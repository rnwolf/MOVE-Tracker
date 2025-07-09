import pytest
import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from datetime import date, datetime, timedelta
import os
import sys
import subprocess
import random
from dateutil.relativedelta import relativedelta

# Path to the main script (assuming it's in the parent directory)
SCRIPT_PATH = os.path.abspath(
    os.path.join(os.path.dirname(__file__), "..", "move_tracker_report.py")
)


@pytest.fixture
def create_test_excel_input(tmp_path):
    """
    Fixture to create a temporary Excel input file for testing,
    mimicking the output of _create_excel_template.
    """
    excel_file = tmp_path / "test_input.xlsx"
    wb = Workbook()

    # --- Fixed dates for reproducible tests ---
    planned_start_date = date(2025, 1, 1)
    planned_delivery_date = date(2025, 3, 31)

    # --- MOVE_Configuration Sheet ---
    ws_config = wb.create_sheet("MOVE_Configuration")
    ws_config.append(["Parameter", "Value"])

    config_data = [
        ("Planned_Start_Date", planned_start_date),
        ("Planned_Delivery_Date", planned_delivery_date),
        (
            "Historic_50th_Percentile_Flow_Time",
            "=_xlfn.PERCENTILE.INC('Historic_Work_Items'!E:E, 0.5)",
        ),
        ("Historic_50th_Percentile_Flow_Time_Override", ""),
        (
            "Initial_Scope",
            "=SUMPRODUCT( --(Current_Work_Items!C$2:INDEX(Current_Work_Items!C:C, MATCH(1E+100, Current_Work_Items!C:C)) <= $B$2), --((Current_Work_Items!G$2:INDEX(Current_Work_Items!G:G, MATCH(1E+100, Current_Work_Items!C:C)) > $B$2) + ISBLANK(Current_Work_Items!G$2:INDEX(Current_Work_Items!G:G, MATCH(1E+100,Current_Work_Items!C:C)))))",
        ),
        ("Initial_Ideal_Completion_Flow_Time", "=B6*B4"),
        ("Buffer_Green_Date", "=B2+B7"),
        ("Buffer_Yellow_Date", "=B8+(0.2*B7)"),
        ("Buffer_Red_Date", "=B9+(0.2*B7)"),
        ("Buffer_Beyond_Red_Date", "=B10+(0.2*B7)"),
        ("Fever_Green_Yellow_Left_Y", 0.2),
        ("Fever_Green_Yellow_Right_Y", 0.5),
        ("Fever_Yellow_Red_Left_Y", 0.5),
        ("Fever_Yellow_Red_Right_Y", 0.8),
    ]

    for idx, (param, value) in enumerate(config_data):
        row_num = idx + 2
        ws_config.cell(row=row_num, column=1, value=param)
        if isinstance(value, str) and value.startswith("="):
            ws_config.cell(row=row_num, column=2, value=value)
        else:
            ws_config.cell(row=row_num, column=2, value=value)

    # Convert date strings to actual dates
    ws_config["B2"].number_format = "YYYY-MM-DD"
    ws_config["B3"].number_format = "YYYY-MM-DD"
    # Apply date formatting to buffer dates
    ws_config["B8"].number_format = "YYYY-MM-DD"
    ws_config["B9"].number_format = "YYYY-MM-DD"
    ws_config["B10"].number_format = "YYYY-MM-DD"
    ws_config["B11"].number_format = "YYYY-MM-DD"

    # --- Historic_Work_Items Sheet ---
    ws_historic = wb.create_sheet("Historic_Work_Items")
    historic_headers = [
        "Historical_WI_ID",
        "Description",
        "Actual_Start_Date",
        "Actual_Completion_Date",
        "Flow_Time_Days",
    ]
    ws_historic.append(historic_headers)

    # Fixed historic data for reproducible tests
    ws_historic.append(
        [
            "HIST-001",
            "Sample 1",
            date(2024, 1, 1),
            date(2024, 1, 5),
            "=INT(D2)-INT(C2)+1",
        ]
    )  # 4 days
    ws_historic.append(
        [
            "HIST-002",
            "Sample 2",
            date(2024, 1, 10),
            date(2024, 1, 17),
            "=INT(D3)-INT(C3)+1",
        ]
    )  # 7 days
    ws_historic.append(
        [
            "HIST-003",
            "Sample 3",
            date(2024, 2, 1),
            date(2024, 2, 3),
            "=INT(D4)-INT(C4)+1",
        ]
    )  # 2 days

    # --- Current_Work_Items Sheet ---
    ws_current = wb.create_sheet("Current_Work_Items")
    current_headers = [
        "Work_Item_ID",
        "Description",
        "Commitment_Date",
        "Status",
        "Actual_Start_Date",
        "Actual_Completion_Date",
        "Date_Withdrawn",
    ]
    ws_current.append(current_headers)

    # Fixed current data for reproducible tests
    ws_current.append(
        [
            "WI-001",
            "Current Item 1",
            date(2025, 1, 1),
            "Completed",
            date(2025, 1, 1),
            date(2025, 1, 5),
            None,
        ]
    )
    ws_current.append(
        [
            "WI-002",
            "Current Item 2",
            date(2025, 1, 1),
            "Completed",
            date(2025, 1, 6),
            date(2025, 1, 10),
            None,
        ]
    )
    ws_current.append(
        [
            "WI-003",
            "Current Item 3",
            date(2025, 1, 1),
            "In Progress",
            date(2025, 1, 11),
            None,
            None,
        ]
    )
    ws_current.append(
        ["WI-004", "Current Item 4", date(2025, 1, 1), "Not Started", None, None, None]
    )

    # --- Other Sheets (empty for now) ---
    wb.create_sheet("Instructions")
    wb.create_sheet("Progress_Log")
    wb.create_sheet("Work_Execution_Chart")
    wb.create_sheet("Fever_Chart")

    # Remove default sheet
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    wb.save(excel_file)
    return excel_file, planned_start_date


def test_progress_log_generation_basic(create_test_excel_input, tmp_path):
    """
    Tests if the Progress_Log sheet is correctly generated for a basic scenario.
    """
    input_excel_path, planned_start_date_fixture = create_test_excel_input
    snapshot_date = planned_start_date_fixture + timedelta(
        days=15
    )  # Example snapshot date
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

    # --- EXPECTED PROGRESS_LOG DATA FOR THIS SCENARIO ---
    # This needs to be manually calculated based on the fixture's generated data
    # and the chosen snapshot_date.
    expected_data = {
        "Snapshot_Date": [
            date(2025, 1, 1),
            date(2025, 1, 5),
            date(2025, 1, 10),
            date(2025, 1, 15),
        ],
        "Scope_At_Snapshot": [1, 1, 2, 3],
        "Actual_Work_Completed": [0, 1, 1, 2],
        "Elapsed_Time_Days": [1, 5, 10, 15],
        "Actual_Operational_Throughput": [
            0.0,
            0.2,
            0.1,
            0.13333333333333333,
        ],  # Placeholder, needs exact calculation
        "Current_50th_Percentile_Flow_Time": [
            0.0,
            4.0,
            4.0,
            4.0,
        ],  # Placeholder, needs exact calculation
        "Forecasted_Delivery_Date": [
            date(2025, 3, 31),
            date(2025, 3, 31),
            date(2025, 3, 31),
            date(2025, 3, 31),
        ],  # Placeholder
        "Buffer_Consumption_Percentage": [0.0, 0.0, 0.0, 0.0],  # Placeholder
        "Work_Done_Percentage": [0.0, 1.0, 0.5, 0.6666666666666666],  # Placeholder
        "Current_Buffer_Signal": ["Green", "Green", "Green", "Green"],  # Placeholder
    }
    expected_df = pd.DataFrame(expected_data)
    expected_df["Snapshot_Date"] = pd.to_datetime(expected_df["Snapshot_Date"]).dt.date
    expected_df["Forecasted_Delivery_Date"] = pd.to_datetime(
        expected_df["Forecasted_Delivery_Date"]
    ).dt.date

    pd.testing.assert_frame_equal(df_progress_log, expected_df, check_dtype=False)

    # Verify chart files are created
    work_exec_chart_path = tmp_path / f"{snapshot_date_str}_work_execution_chart.png"
    fever_chart_path = tmp_path / f"{snapshot_date_str}_fever_chart.png"

    assert work_exec_chart_path.exists()
    assert work_exec_chart_path.stat().st_size > 0

    assert fever_chart_path.exists()
    assert fever_chart_path.stat().st_size > 0

    # Optional: Verify images are embedded in the Excel file
    wb_output = load_workbook(input_excel_path)
    assert "Work_Execution_Chart" in wb_output.sheetnames
    assert "Fever_Chart" in wb_output.sheetnames
    assert len(wb_output["Work_Execution_Chart"]._images) > 0
    assert len(wb_output["Fever_Chart"]._images) > 0
