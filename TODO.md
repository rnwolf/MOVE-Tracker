# TODO's

## Problem Statement

    Reference. Tameflow, TOC and CCPM


## Incorporate the flow overide in calculations (Ignored at present, inspite of what docs say.)

##  Update docs with

    Initial_Ideal_Completion_Flow_Time=Target_Scope * Historic_50th_Percentile_Flow_Time
    Green_Buffer_Start_Date=Planned_Start_Date + Initial_Ideal_Completion_Flow_Time
    Yellow_Buffer_Start_Date=Green_Buffer_Start_Date + (20% * Initial_Ideal_Completion_Flow_Time)
    Red_Buffer_Start_Date=Yellow_Buffer_Start_Date + (20% * Initial_Ideal_Completion_Flow_Time)
    Beyond_Red_Date=Red_Buffer_Start_Date + (20% * Initial_Ideal_Completion_Flow_Time)

## Correct spec for the following


    4. Input Excel File Structure

    Current_Work_Items

    Needs Committment_Date

    MOVE_Configuration

    Expected Parameters
    Add
    Buffer_Beyond_Red_Date


    The difficulty of having to evaluate the calculated cells in Excel and hence the testing strategy of calculating Parameters and values within the test script.



    6. Core Logic / Workflow

    4. Calculate Derived Configuration & Flow Times:

    Initial_Scope (at Planned_Start_Date) should be derived from Committment_Date  <= Planned_Start_Date NOT Actual_Start_Date <= Planned_Start_Date

    the scope we need to count the work items that have been committed to and exclude those that we now know are withdrawn as at the Planned_Start_Date.


    7 Generate Charts (Matplotlib, PNG Output):

    Work Execution Chart:

    Add vertical lines at Planned_Delivery_Date, Green_Buffer_Start_Date, Yellow_Buffer_Start_Date, Red_Buffer_Start_Date >>> and Buffer_Beyond_Red_Date <<<< using the dynamically calculated buffer dates.

    Fever Chart:

    Annotate the line markers with associated Snapshot_Dates so that we can relate Fever Chart plot with Progress_Log.




    7. Data Structures

    dataclass for MOVEConfiguration

    buffer_beyond_red_date: date
    What about "historic_50th_percentile_flow_time"?


    11. Testing Strategy

    "A single test file, test_excel_integration.py"

    tests\test_chart_visuals.py

    Actually multiple files.
