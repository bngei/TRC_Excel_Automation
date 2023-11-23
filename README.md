# TRC_Excel_Automation

### Data Extraction Functions:
Created a set of functions, each responsible for extracting specific data from designated sheets within Excel files:

- general_extraction: Extracts general data from the "1.General" sheet.
- enclosure_extraction: Focuses on data extraction from the "2.Enclosure" sheet.
- power_extraction: Extracts relevant information from the "2.Power" sheet.
- instrumentation_extraction: Focuses on data extraction from the "3.Instrumentation" sheet.
- plc_extraction: Extracts data specifically from the "4.PLC.RTU" sheet.
- comms_extraction: Deals with data extraction from the "5.Comms" sheet.
- retrieve_asset_ID: Extracts and returns the asset ID from the specified Excel file path.


### Consolidation and File Management:
- create_excel_file: Generates an empty Excel file at the given path if one does not exist already.
- create_duplicate_file: Creates a duplicate of the consolidated Excel file with a timestamp in the filename.
- recieve_date_and_time: Generates a timestamp in a specific format.
- return_list_of_paths: Gathers and returns a list of file paths within a designated folder, specifically targeting Excel files.
- main_function: The primary function orchestrating the consolidation process.
    - It creates a consolidated Excel file (Wireframe_Consolidated_Data.xlsx) if it doesn't exist.
    - Iterates through a list of file paths, invoking extraction functions to compile data into the consolidated file.
    - Creates a duplicate of the consolidated file with a timestamp.


### Workflow Overview:
- Input: A list of file paths pointing to Excel files structured similarly.
- Output: A master Excel file (Wireframe_Consolidated_Data.xlsx) containing data aggregated from various sheets of input files. Additionally, a timestamped duplicate of this master file.


### Purpose:
This work simplifies and automates the process of consolidating structured data from multiple Excel files into a single comprehensive file, aiding in efficient data analysis and management.

This tool streamlines the extraction and combination of specific information from designated sheets in Excel files, enhancing the accessibility and usability of data gathered from disparate sources.
