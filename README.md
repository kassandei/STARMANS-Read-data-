# STARMANS Microcontroller Data to Excel Processor

This script reads data from a microcontroller and populates it into specific cells of an Excel file. The Excel file must be empty and present in the same directory as the script each time the script is run.

## Requirements

Ensure you have Python installed on your system. Additionally, install the required libraries:

- `pyserial`
- `openpyxl`

You can install these libraries using pip:

```sh
pip install pyserial openpyxl
```

## Setup Instructions

Download the Script: Download the script_name.py script and save it in a directory of your choice.

Prepare the Excel Template:

Create an Excel file (data_template.xlsx) and leave it empty.
Place this Excel file in the same directory as the script. This file will serve as the template where data will be populated.

## Usage

- Connect Microcontroller:

- Ensure your microcontroller is connected to the computer via USB.

- Modify USB Port (if needed): The USB port may vary from computer to computer. If necessary, modify the serial port path in line 59 of script_name.py to match your system.

- Run the Script: Open a terminal or command prompt.
Navigate to the directory where script_name.py and data_template.xlsx are located.
Execute the script using Python with the following command:

```sh
python script_name.py
```
- Monitor Data Transfer: The script will continuously read data from the microcontroller until it has received 20 lines of data or you interrupt it with Ctrl+C.

## Data Processing:

Data received is parsed into blocks and then saved to specific cells in the Excel file (data_template.xlsx).
Excel Output:

Once data is successfully saved, the script will print a message confirming the Excel file has been updated.

## Notes

Ensure that each time you run the script, data_template.xlsx is empty and placed in the script's directory.
The script assumes a specific layout for data placement in the Excel file. Modify the script's cell_positions if your Excel layout differs.
If multiple Excel files (*.xlsx) are present in the script's directory, the script will use the first one found.
