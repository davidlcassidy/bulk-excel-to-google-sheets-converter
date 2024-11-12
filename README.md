# Excel to Google Sheets Converter

This Google Apps Script automates the conversion of Excel files stored in a Google Drive folder to Google Sheets. It then saves these Google Sheets files in a separate specified folder within Google Drive.

## Table of Contents
1. [Overview](#overview)
2. [Requirements](#requirements)
3. [Installation](#installation)
4. [Configuration](#configuration)
5. [Usage](#usage)
6. [Code Explanation](#code-explanation)
7. [Troubleshooting](#troubleshooting)
8. [License](#license)

## Overview

This script performs the following tasks:
1. Finds all Excel files in a specified Google Drive folder.
2. Checks if a corresponding Google Sheets file already exists in the target folder.
   - If the file exists, the script updates it.
   - If the file does not exist, the script creates a new Google Sheets file with the same name.
3. Saves the Google Sheets files in a target folder for easy access.

## Requirements

- **Google Workspace Account**: You must have access to Google Drive and Google Apps Script.
- **OAuth Permissions**: The script requires authorization to manage files on Google Drive.
- **Excel Files in Google Drive**: Ensure you have a folder with Excel files (.xls or .xlsx) ready for conversion.

## Installation

1. **Create a New Google Apps Script Project**:
   - In Google Drive, click on **New > Google Apps Script** to create a new Apps Script project.
   
2. **Copy the Script**:
   - Copy the entire script code provided and paste it into the Apps Script editor.

3. **Save and Name the Project**:
   - Click **File > Save**, then provide a name, e.g., `Excel to Google Sheets Converter`.

4. **Authorize the Script**:
   - When you first run the script, you will be prompted to authorize it. Follow the on-screen instructions to provide the necessary permissions.

## Configuration

Before running the script, you need to specify the folder IDs for both the source (Excel files) and the target (Google Sheets) folders.

1. **Get Folder IDs**:
   - Go to your Google Drive and navigate to the Excel files folder. In the URL, find the folder ID (the part after `folders/`).
   - Do the same for the folder where you want the converted Google Sheets files to be saved.
   
2. **Set Folder IDs in the Script**:
   - Update the following variables in the script with your folder IDs:
     ```javascript
     var excelFolderId = 'YOUR_EXCEL_FOLDER_ID';
     var sheetsFolderId = 'YOUR_SHEETS_FOLDER_ID';
     ```

## Usage

1. **Run the `main` Function**:
   - In the Apps Script editor, select `main` from the function dropdown and click the **Run** button. 
   
2. **Check the Target Folder**:
   - After running the script, check your target folder in Google Drive to see the converted Google Sheets files. The script will update existing files or create new ones as needed.

## Code Overview

Here is an overview of the key parts of the script:

- **`main()`**:
  - This is the main entry point that initiates the conversion process by calling `convertExcelFiles2Sheets_` with the source and target folder IDs.

- **`convertExcelFiles2Sheets_` Function**:
  - **Parameters**: `excelFolderId` (source folder) and `sheetsFolderId` (target folder).
  - **Logic**:
    - Loops through each file in the source folder.
    - Checks if a Google Sheets version of the Excel file already exists in the target folder.
    - Uses Google Drive API functions (`APIGoogleDriveUpload_` and `APIGoogleDriveUpdate_`) to upload or update files.

- **Helper Functions**:
  - `APIGoogleDriveUpload_`: Uploads an Excel file and converts it to Google Sheets format.
  - `APIGoogleDriveUpdate_`: Updates an existing Google Sheets file with new data from the Excel file.

## Troubleshooting

- **Permissions**: Ensure you grant the required permissions when prompted.
- **API Limitations**: Large Excel files may cause timeouts. Try breaking them into smaller files or processing fewer files at once.
- **Quota Limits**: Google Drive API usage may be subject to daily limits. Monitor your quota in the Google Cloud Console if needed.

## License

This project is licensed under the GNU GPLv3 License - see the [LICENSE](LICENSE) file for details.
