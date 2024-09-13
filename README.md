# TST Optibus Apps Ops

This repository contains scripts created for a client called TST, intended to support their operations on the Ops Go Live in 2022. These scripts automate critical reporting tasks, ensuring that TST can operate efficiently and effectively.

## Folder Structure

The repository follows a structured layout:
1. **SMS Report NOS**
2. **Weekly Schedule Report**
3. **Vehicle Distribution Report**

This organization helps maintain a strict structure, enabling users to easily navigate and contribute to the repository.

---

## 1. SMS Report NOS 📱

### Overview

The SMS Report NOS script automates the transformation of daily reports and driver CSVs into an .exe file, following a specific format for sending automatic SMS messages to drivers with their tasks for the following day.

### Motivation

TST needed an efficient way to notify drivers of their next day's tasks. Manually preparing and sending these notifications was error-prone and time-consuming. This script automates the process, ensuring timely and accurate communication.

### Features

- **Automated Data Processing**: Reads and processes daily reports and driver CSVs.
- **SMS Formatting**: Formats the information according to the specific requirements for SMS notifications.
- **User-friendly Interface**: Simple Streamlit app for easy interaction.

### Installation

1. **Clone the repository**:
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2. **Install the required dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

### Usage

1. **Run the Streamlit app**:
    ```bash
    streamlit run sms_report_nos.py
    ```

2. **Generate SMS Report**:
    - Open the Streamlit interface in your browser.
    - Upload the daily report and driver CSV files.
    - Download the generated SMS report.

### Script Overview

#### Flow of the Script

1. **File Upload**:
   - The user uploads the daily report and driver CSV files through the Streamlit interface.

2. **Data Processing**:
   - The script reads and processes the input files, translating headers and preparing the data for SMS formatting.

3. **SMS Formatting**:
   - The script formats the information according to TST's SMS notification requirements.

4. **Report Generation**:
   - The script generates an Excel file with the formatted SMS messages and provides a download link.

#### Key Functions

- **get_telephone_file(input_file)**:
  - Reads the driver CSV file and extracts telephone numbers.
- **get_output_path(inputFile, sheet_name)**:
  - Generates a unique output path for the report.
- **get_sheet_list(input_file)**:
  - Retrieves the list of sheets in the daily report.
- **translate_headers(df)**:
  - Translates the headers from Portuguese to English.
- **get_sheet_data(input_file, sheet_name, drivers_phones)**:
  - Extracts and processes data from the specified sheet.
- **print_sheet_data(data_list, df)**:
  - Appends data rows to the DataFrame.
- **print_header(df)**:
  - Initializes the DataFrame with the appropriate header.
- **build_sms_report(daily_report, drivers_csv)**:
  - Coordinates the entire process of building the SMS report.

---

## 2. Weekly Schedule Report 🗓️

### Overview

The Weekly Schedule Report script generates a table that can be displayed in depots, detailing the tasks and duties of each driver for the week.

### Motivation

TST needed a clear and organized way to communicate weekly schedules to drivers. This script automates the creation of a comprehensive schedule, reducing manual effort and minimizing errors.

### Features

- **Automated Schedule Generation**: Reads daily reports and compiles weekly schedules.
- **Detailed Information**: Includes driver names, duties, vehicles, and timings.
- **User-friendly Interface**: Simple Streamlit app for easy interaction.

### Installation

1. **Clone the repository**:
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2. **Install the required dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

### Usage

1. **Run the Streamlit app**:
    ```bash
    streamlit run weekly_schedule_report.py
    ```

2. **Generate Schedule Report**:
    - Open the Streamlit interface in your browser.
    - Upload the daily report file.
    - Download the generated weekly schedule report.

### Script Overview

#### Flow of the Script

1. **File Upload**:
   - The user uploads the daily report file through the Streamlit interface.

2. **Data Processing**:
   - The script reads and processes the daily report, translating headers and validating the data.

3. **Schedule Generation**:
   - The script compiles the weekly schedule, ensuring all required information is included.

4. **Report Generation**:
   - The script generates an Excel file with the weekly schedule and provides a download link.

#### Key Functions

- **validate_daily_report(df)**:
  - Validates the structure and content of the daily report.
- **translate_headers(df)**:
  - Translates the headers from Portuguese to English.
- **get_weekday_name(day)**:
  - Returns the Portuguese name for a given weekday.
- **get_month_name_pt(month)**:
  - Returns the Portuguese name for a given month.
- **get_sheet_list(input_file)**:
  - Retrieves the list of sheets in the daily report.
- **get_sheet_data(input_file, sheet_name)**:
  - Extracts and processes data from the specified sheet.
- **print_sheet_data(data_list, output_file)**:
  - Appends data rows to the output file.
- **print_header(output_file)**:
  - Initializes the output file with the appropriate header.
- **style_sheet(output_file, sheet)**:
  - Styles the output file for better readability.
- **build_weekly_schedule_report(daily_report)**:
  - Coordinates the entire process of building the weekly schedule report.

---

## 3. Vehicle Distribution Report 🚍

### Overview

The Vehicle Distribution Report script combines data from the full schedule and the daily report to show the planned allocation of each vehicle throughout the blocks and duties for the week. It also provides space for dispatchers to fill in actual information.

### Motivation

TST needed a detailed view of vehicle allocations to manage operations effectively. This script automates the creation of a comprehensive vehicle distribution report, providing both planned and actual data.

### Features

- **Automated Data Integration**: Combines data from full schedules and daily reports.
- **Detailed Allocation Information**: Includes duty IDs, vehicles, drivers, and timings.
- **User-friendly Interface**: Simple Streamlit app for easy interaction.

### Installation

1. **Clone the repository**:
    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2. **Install the required dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

### Usage

1. **Run the Streamlit app**:
    ```bash
    streamlit run vehicle_distribution_report.py
    ```

2. **Generate Vehicle Distribution Report**:
    - Open the Streamlit interface in your browser.
    - Upload the full schedule and daily report files.
    - Download the generated vehicle distribution report.

### Script Overview

#### Flow of the Script

1. **File Upload**:
   - The user uploads the full schedule and daily report files through the Streamlit interface.

2. **Data Processing**:
   - The script reads and processes the input files, translating headers and preparing the data for report generation.

3. **Report Generation**:
   - The script generates an Excel file with the vehicle distribution report and provides a download link.

#### Key Functions

- **get_sheet_names(daily_report)**:
  - Retrieves the list of sheets in the daily report.
- **create_full_schedule_df(full_schedule)**:
  - Processes the full schedule data into a DataFrame.
- **create_full_schedule_dict(full_schedule)**:
  - Converts the full schedule DataFrame into a dictionary.
- **create_daily_report_dict(daily_report, sheet_name)**:
  - Processes the daily report data into a dictionary.
- **create_table_lines_list(daily_blocks_dict, fs_blocks_dict)**:
  - Creates a list of table lines from the processed data.
- **create_block_lines_dict(daily_blocks_dict, fs_blocks_dict)**:
  - Creates a dictionary of block lines for the report.
- **get_max_count(blockLinesDictionary)**:
  - Determines the maximum count of block occurrences.
- **create_output_directory(daily_report)**:
  - Creates an output directory for the report files.
- **create_output_file(newDirpath, sheet_name)**:
  - Creates an output file for the report.
- **create_header(blockLinesDictionary)**:
  - Creates the header for the report.
- **print_header(wb, blockLinesDictionary)**:
  - Prints the header in the output file.
- **print_table_lines(wb, blockLinesDictionary)**:
  - Prints the table lines in the output file.
- **style_output_report(wb)**:
  - Styles the output report for better readability.
- **build_vehicle_reports(full_schedule, daily_report)**:
  - Coordinates the entire process of building the vehicle distribution report.
