# Google Sheets API Integration
This project aims to read a spreadsheet in Google Sheets and make changes based on the data found in it. In this case, it is calculating the status.

## Requirements
Before using this code, follow the steps below to set up your credentials and activate the Google Sheets API:

- **Step 1**: Access Google Cloud Console, Link: [Google Cloud Console](https://www.google.com/aclk?sa=l&ai=DChcSEwjp38z90byEAxU2YUgAHd9mB3UYABAAGgJjZQ&gclid=EAIaIQobChMI6d_M_dG8hAMVNmFIAB3fZgd1EAAYASAAEgLBg_D_BwE&sig=AOD64_00JYK1PVOod-wWc5tqJ_uYVTyoag&q&adurl&ved=2ahUKEwix3Mb90byEAxWHppUCHacwBWgQ0Qx6BAgFEAE)
- **Step 2**: Search for **"Google Sheets API"** and activate it for your project.
- **Step 3**: Search for **"Credentials"** and create or select a service credential. Download the JSON file containing your credentials.

## Configuration
- Clone this repository to your local machine.
- Ensure you have the .NET Core SDK installed on your machine.
- Add the JSON credential file downloaded from the Google Cloud Console to the project's Credentials directory.
- In the code, **update the value of the credPath variable** to point to the path of the credential file on your local machine.
- Update the value of the spreadsheetId variable to the ID of the Google Sheets spreadsheet you want to modify.

## How to Use
After setting up your credentials and adjusting the variables in the code, you can run the project. It will read the specified spreadsheet, calculate the students' grades, and update the spreadsheet with the results.

## Notes
Ensure that the Google Sheets spreadsheet has the correct structure and that the ranges specified in the code correspond to the correct data in the spreadsheet.
Link to the spreadsheet used in the project: [Google Sheets Spreadsheet](https://docs.google.com/spreadsheets/d/1-iz4w7WxPP5Dj6I-UjL5t1Y9WDrXqgg_g7B2csJVRbo/edit#gid=0)

## License
This project was made by Caio Braga (github.com/caiotbraga)
