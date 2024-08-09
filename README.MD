# Excel Comparison Tool

This tool compares two Excel files (`originSource.xlsx` and `newSource.xlsx`) and generates a new Excel file (`output.xlsx`) with the differences, new data, and errors.

## Prerequisites

- Go 1.16 or later
- `github.com/tealeg/xlsx` library

## Installation

1. **Install Go**: Follow the instructions on the [official Go website](https://golang.org/doc/install) to install Go.
2. **Get the xlsx library**: Run the following command to install the `xlsx` library:
    ```sh
    go get github.com/tealeg/xlsx
    ```

## Usage

1. **Prepare the Excel files**: Ensure you have `originSource.xlsx` and `newSource.xlsx` in the same directory as the Go program.
2. **Run the program**: Execute the following command in your terminal:
    ```sh
    go run main.go
    ```
3. **Check the output**: The program will generate `output.xlsx` in the same directory.

## Explanation of the Code

- **main**: The entry point of the program. It opens the source files, creates a new Excel file, compares the sheets, and saves the output.
- **openExcelFile**: Opens an Excel file and returns the file object.
- **createNewExcelFile**: Creates a new Excel file with three sheets: Difference, New Data, and Data Error. It also adds headers to these sheets.
- **addHeaders**: Adds headers to a given sheet.
- **compareSheets**: Compares the sheets from the two source files and populates the Difference, New Data, and Data Error sheets.
- **compareSheet**: Compares rows from two sheets and categorizes them into differences, new data, or errors.
- **compareRows**: Compares two rows and returns `true` if they are identical.
- **addRowToSheet**: Adds a row to a given sheet.
- **saveExcelFile**: Saves the Excel file to disk.
- **validateRow**: Validates a row based on specific criteria.

## Example

After running the program, you will get an `output.xlsx` file with the following sheets:

- **Difference**: Rows that are different between the two source files.
- **New Data**: Rows that are present in the new source file but not in the original source file.
- **Data Error**: Rows that have validation errors.

```sh
$ go run main.go
2023/10/01 12:00:00 Opening originSource.xlsx...
2023/10/01 12:00:00 Opening newSource.xlsx...
2023/10/01 12:00:00 Creating new Excel file...
2023/10/01 12:00:00 Comparing sheets...
2023/10/01 12:00:00 Saving output.xlsx...
Data saved to output.xlsx