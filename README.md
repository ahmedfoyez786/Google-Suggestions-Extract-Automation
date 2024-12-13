# Google Search Suggestions Extract Automation || QUPS_Test Automation_Q1_Assignment

This project automates the process of fetching Google search suggestions for a list of keywords stored in an Excel file. It utilizes Selenium for browser automation and OpenPyXL for reading and writing Excel files.

## Features
- Reads keywords from an Excel file based on the current day of the week.
- Performs Google searches for each keyword.
- Fetches the longest and shortest suggestions for each keyword.
- Writes the results back into the Excel file.

## Requirements
### Prerequisites
- Python 3.8 or higher
- Microsoft Edge Browser
- Microsoft Edge WebDriver

### Python Libraries
Install the required Python libraries using:
```bash
pip install selenium openpyxl
```

## File Structure
### Input Excel File
- The Excel file should have multiple sheets named after days of the week (e.g., Saturday, Sunday, Monday, etc.).
- Each sheet should have keywords listed in column **C** starting from row 3.

### Output
- The longest and shortest suggestions for each keyword are written into columns **D** and **E**, respectively.

## Usage
1. Place the Excel file named `Excel.xlsx` in the root directory of the project.
2. Ensure the Edge WebDriver is installed and the path is correctly set in the script.
3. Run the script:
   ```bash
   python main.py
   ```

## Script Behavior
1. Identifies the current day and selects the corresponding sheet from the Excel file.
2. Reads keywords from column **C** of the selected sheet.
3. Performs a Google search for each keyword.
4. Fetches suggestions from Google and identifies the longest and shortest options.
5. Writes the results back into columns **D** and **E** of the same sheet.

## Example Excel Structure
| **A** | **B**       | **C**       | **D**               | **E**              |
|-------|-------------|-------------|---------------------|--------------------|
|       |             | Keywords    | Longest Option      | Shortest Option    |
|       | Keyword1    | Dhaka       | Dhaka Education Board | Dhaka             |
|       | Keyword2    | Saturday    | Saturday Night Live | Saturday          |

## Important Notes
- Ensure that the Excel file has sheets named exactly as the days of the week.
- The WebDriver should be compatible with your version of Microsoft Edge.
- The script simulates human typing and interaction to avoid detection as a bot.




## Acknowledgments
- [Selenium Documentation](https://www.selenium.dev/documentation/)
- [OpenPyXL Documentation](https://openpyxl.readthedocs.io/en/stable/)
