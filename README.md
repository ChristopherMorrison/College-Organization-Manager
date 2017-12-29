# College-Organization-Manager
A semi-autonomous Python3 and Google Sheets based college organization management tool designed for easy deployment and use. Meant for groups of up to 500 due to current speed limitations of the API.

Features:

* Designed around University of Cincinnati student identification (but can be easily modified)
* Automatic roster updates from member sign ins, including automatic additions of new members
* Ability to copy and paste in multiple rosters and condense them down to one large roster for carrying over large amounts of previously gathered data
* Control panel located within the spreadsheet document for easy settings updates during runtime
* Automatic processing of semester change
* Easy to modify roster filters and settings
* Detection and consolidation of duplicate member entries

## Requirements
1. A Google account
2. A system with python3 with pip/pip3 and internet connectivity

## Setup
1. Create a Google Oauth2 client_key.json file from [the Google Developers Console](https://console.cloud.google.com/apis/dashboard) by following the instructions [from this YouTube Video](https://www.youtube.com/watch?v=vISRn5qFrkM)

2. (Optional) Create a [python virtual environment](https://docs.python.org/3/tutorial/venv.html) to keep the dependencies seperate from other projects

    For Linux or Mac:

    python3 -m venv Name-Of-Virtual-Environment-Folder
    source Name-Of-Virtual-Environment-Folder/bin/activate

    For Windows:

    python3 -m venv Name-Of-Virtual-Environment-Folder
    Name-Of-Virtual-Environment-Folder\Scripts\activate.bat

3. Install package dependencies
    
    ```bash
    pip3 install -r requirements.txt
    ```
    
4. Run the program
    
    ```bash
    python3 Manager.py
    ```
    
5. If there is no settings.cfg or there is an issue with the current settings.cfg file the program will go through a first time setup and you will need to enter your email and a few other things when prompted
6. If the sheet cannot be connected to you will be prompted to either create a new one or remediate the issue by altering settings.cfg

## Documentation
1. Screenshots to come
2. Examples to come

## Future Development plans
- Automatic generation of the google sheet that is required to use the program
- Instructions on how to cleanly integrate a google forms based sign in sheet or do it automatically
- Statistics of the internal roster
- Cleaner code, ways to rename things, and add-on implementation examples
- Models instead of ad-hoc lists of lists
- Available updated version notification
- Changes based on feedback from instances of use
- Email connectivity to alert system administrator of major issues
- Subroster generation to quickly determine subgroups of the organization's members

## Notes
- [Anton Burnashev's gspread](https://github.com/burnash/gspread) is the API used to communicate with the google sheet.
