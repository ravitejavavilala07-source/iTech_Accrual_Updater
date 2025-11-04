# How to Run and Test PAYSHEET_SPLIT_CELL_VALIDATOR.py

This comprehensive guide will help you manually run and test the `PAYSHEET_SPLIT_CELL_VALIDATOR.py` script. Follow the step-by-step instructions along with examples and expected outputs to ensure successful execution of the validator script.

## Table of Contents
1. [Prerequisites](#prerequisites)
2. [Script Overview](#script-overview)
3. [Running the Script](#running-the-script)
   - [Step 1: Open Terminal](#step-1-open-terminal)
   - [Step 2: Navigate to Script Location](#step-2-navigate-to-script-location)
   - [Step 3: Execute the Script](#step-3-execute-the-script)
4. [Testing the Script](#testing-the-script)
   - [Example Test Cases](#example-test-cases)
   - [Expected Outputs](#expected-outputs)
5. [Troubleshooting](#troubleshooting)
6. [Conclusion](#conclusion)

## Prerequisites
Before running the script, ensure that you have:
- Python 3.x installed on your machine.
- Necessary dependencies installed (usually can be installed via `pip`).

## Script Overview
`PAYSHEET_SPLIT_CELL_VALIDATOR.py` is designed to validate the contents of paysheet cell data to ensure data integrity before processing.

## Running the Script

### Step 1: Open Terminal
- On Windows, you can use Command Prompt or PowerShell.
- On macOS or Linux, open the Terminal application.

### Step 2: Navigate to Script Location
Use the `cd` command to change the directory to where the script is located. For example:
```bash
cd /path/to/your/script/
```

### Step 3: Execute the Script
Run the script using Python by executing the following command:
```bash
python PAYSHEET_SPLIT_CELL_VALIDATOR.py
```

## Testing the Script
To ensure that the script works as expected, you can create some test cases.

### Example Test Cases
1. **Valid Input**:
   - Input: `valid_data.txt`
   - Command:
   ```bash
   python PAYSHEET_SPLIT_CELL_VALIDATOR.py valid_data.txt
   ```
   - Expected Output: `Validation successful: All cells are valid.`

2. **Invalid Input**:
   - Input: `invalid_data.txt`
   - Command:
   ```bash
   python PAYSHEET_SPLIT_CELL_VALIDATOR.py invalid_data.txt
   ```
   - Expected Output: `Validation failed: Invalid data found in cells.`

### Expected Outputs
- Successful validation should print a success message to the console.
- If validation fails, the script should output error messages indicating which cells are invalid.

## Troubleshooting
- **Error Message**: "File not found."
  - **Solution**: Ensure that the file path is correct and the file exists.
- **Error Message**: "Invalid format."
  - **Solution**: Check that the input file follows the expected format.

## Conclusion
By following this guide, you’ll be able to manually run and test the `PAYSHEET_SPLIT_CELL_VALIDATOR.py` script effectively. Ensure to revisit the test cases to validate changes made to the script in the future.