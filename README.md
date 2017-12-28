# College-Organization-Manager
A semi-autonomous Python3 and Google Sheets based college organization management tool designed for easy deployment and use. Meant for groups of 10-500 due to current speed limitations of the API.

Features:

* Automatic interpretation of data provided
* Designed around University of Cincinnati student identification (but can be easily modified)
* Automatic roster updates from member sign ins including additions of new members
* Ability to copy and paste in multiple rosters and condense them down to one large roster for carrying over large amounts of previously gathered data
* Control panel located within the spreadsheet document for easy settings updates during runtime
* Automatic processing of semester change

## Requirements
1. A Google account (free)
2. A system with python3 installed and internet connectivity

## Setup
1. Create a Google Oauth2 client_key.json file from [here](https://console.cloud.google.com/apis/dashboard) by following the instructions [here](https://www.youtube.com/watch?v=vISRn5qFrkM)
2. (Optional) Create a python virtual encironment to keep the dependencies seperated from other projects
    
    ```bash
    python3 -m venv Name-Of-Virtual-Environment-Folder
    source Name-Of-Virtual-Environment-Folder/bin/activate
    ```

3. Install package dependencies
    
    ```bash
    pip3 install -r requirements.txt
    ```
    
4. Run the program
    
    ```bash
    python3 Manager.py
    ```
    
5. Watch the program fail as I have yet to release the segment that creates the Google sheet automatically
6. Connect a google sign in sheet...

## Configuration
1. TBD

## Usage and Capabilities
1. TBD

## Documentation
1. TBD, probably just screenshots and arrows with a sprinkle of graphs

## Future Development plans
- Instructions on how to cleanly integrate a google forms based sign in sheet
- Cleaner code, ways to rename things, and add-on implementation examples
- Available update notification
- Feedback from instances of use
- Automatic generation of the google sheet that is required to use the program
- Statistics of the internal roster
- Models instead of ad-hoc lists of lists
- Emailconnectivity to alert system administrator of major issues
- Subroster generation to quickly determine subgroups of the organization's members

## Notes
- I think I'm obligated in some way by the MIT to mention gspread which functions as the API for connecting to the spreadsheet
- I think I'm also obligated to mention the other open source projects I used, they are all listed in the requirements.txt file but I'm not distributing them so I don't think I need to include their licenses here, you can look them up yourself
