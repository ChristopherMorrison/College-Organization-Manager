# Imports
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import datetime
import os
import httplib2


# Very important variables
settings_filename = 'settings.cfg'

# TODO: Investigate replacing settings_file class with a dictionary
class settings_file:
    Google_Sheet_Name = None
    Google_Sheet_key = None
    Google_Sheet_URL = None
    Administrator_Email = None

# debug variables
dbo = None

# API details
scope = ["https://spreadsheets.google.com/feeds"]
creds = ServiceAccountCredentials.from_json_keyfile_name("client_secret.json",scope)

# TODO: Scrape all up the global variables to be within a single object and populate them from settings.cfg
# Global object variables
client      = None
spreadsheet = None

# Global State variables
client_authorized  = False
spreadsheet_opened = False

# Global Control Panel Variables
last_roster_aggregation_time     = '-'
last_sign_in_processing_time     = '-'
last_sign_in_processed_timestamp = None
last_subroster_generation_time   = '-'
agent_start_time = None

post_interval_sleep_time   = 1 
roster_aggregation_period  = 10
check_signin_period        = 10
generate_subroster_period  = 10

current_semester = None
fall_semester_start_date   = None
spring_semester_start_date = None
summer_semester_start_date = None

# Colors and print functions via ANSI excape sequences
# src: https://stackoverflow.com/questions/287871/print-in-terminal-with-colors
class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'
def printError(str_message):
    print('[' + bcolors.FAIL + 'FAIL' + bcolors.ENDC + '] ' + str_message + '')
def printSuccess(str_message):
    print('[' + bcolors.OKGREEN + 'GOOD' + bcolors.ENDC + '] ' + str_message + '')
def printWarning(str_message):
    print('[' + bcolors.WARNING + 'WARN' + bcolors.ENDC + '] ' + str_message + '')
def printInfo(str_message):
    print('[' + bcolors.OKBLUE + 'INFO' + bcolors.ENDC + '] ' + str_message + '')
def printMessage(str_message):
    print('----->' + bcolors.HEADER + str_message + bcolors.ENDC)


# Helper functions
def start_client():
    # Clear any previous client instance
    global client
    client = None
    
    global client_authorized    
    client_authorized = False
    
    # Open new client instance
    try:
        client = gspread.authorize(creds)
        printSuccess('Authentication token accepted')
        client_authorized = True
        return client
    except:
        printError('Could not authorize the client instance, try checking the Oauth2 credential json.')
        raise
def open_spreadsheet(key_or_name):
    # Make sure the client has been started
    global client
    global client_authorized
    
    if not client_authorized:
        printError('The client must be started before the client can open a sheet')
        assert client_authorized, "The client is not authorized"
    
    # attempt to open the sheet
    printInfo('Attempting to open sheet ' + key_or_name)
    try:
        document = client.open(key_or_name)
        printSuccess('Spreadsheet ' + key_or_name + ' successfully opened')
        global spreadsheet_opened
        spreadsheet_opened = True
        return document
    except:
        printError('Could not open the sheet ' + key_or_name + ', please ensure that the provided Oauth2 credentials are allowed to access the document and that the correct document is selected.')
        raise
def open_worksheet(worksheet_name):
    # Make the spreadsheet is opened
    global spreadsheet
    global spreadsheet_opened
    
    if not spreadsheet_opened:
        printError('The spreadsheet must be opened before worksheets can be open')
        assert spreadsheet_opened
    
    # Open the worksheet
    printInfo('Attempting to open worksheet ' + worksheet_name)
    try:
        opened_sheet = spreadsheet.worksheet(worksheet_name)
        printSuccess('Worksheet "' + worksheet_name + '" opened')
        return opened_sheet
    except:
        printError('Could not open worksheet ' + worksheet_name)
        raise
def str2dt(string):
    return datetime.datetime.strptime(string,'%m/%d/%Y %H:%M:%S')
def debug(object_to_debug):
    global dbo
    dbo = object_to_debug
    exit()
def writeback(worksheet, values):
    '''
        Please note that values is the list of list of values and not the cell values
        If you had the cell values you would just use worksheet.update_cells(cell_list)
    '''
    # Print message
    printWarning('Now undergoing worksheet writeback, this may take a long time')
    
    # get cell range from worksheet
    all_values = worksheet.get_all_values()
    start_range = worksheet.get_addr_int(1, 1)
    end_range = worksheet.get_addr_int(len(all_values), len(all_values[0]))
    active_cells = worksheet.range(start_range + ':' + end_range)
    
    # Add depth to active_cells to make it easier to traverse
    active_cells = [active_cells[x:x+len(values[0])] for x in range(0, len(active_cells), len(values[0]))]
    
    # Update cell values
    for row in range(len(values)):
        for col in range(len(values[0])):
            active_cells[row][col].value = values[row][col]
    
    # clear remaining cells
    if len(active_cells) > len(values):
        offset = len(values)
        for row in range(len(active_cells) - offset):
            for col in range(values[0]):
                active_cells[row + offset][col].value = ''
    
    # flatten active_cells list
    temp = []
    for entry in active_cells:
        temp+=entry
    active_cells = temp
    
    # writeback to the worksheet
    worksheet.update_cells(active_cells)
    printSuccess('Writeback complete')
    
    return

# Setup functions
def First_Time_Setup(settings_filename = 'settings.cfg'):
    #TODO
    # Make sure settings.cfg exists or (generate a new one and exit)
    if not os.path.isfile(settings_filename):
        Generate_Settings_File(settings_filename)
    
    # Load in values from settings.cfg
    Load_Settings_File(settings_filename)

    # make sure we have an admin email
    if settings_file.Administrator_Email == '':
        printWarning('There is no Administrator_Email provided in the settings file "'+settings_filename+'". Please provide an email before continuing')
        # TODO prompt user for email
    
    # Make sure the provided credentials work/exist
    '''
    try:
        if os.path.isfile(credentials_file):
            global client
            client = start_client(///credentials///)
        else:
            printWarning('There are no Oauth2 credentials with the file name "' + credentials_file +'". Please refer to README to find information on how to generate these.')
            raise #TODO add custom errors
    except:
        printWarning('The Oauth2 credentials are invalid. Please check your credentials file and refer to README.md.')
        raise
    '''
    # Make sure the spreadsheet exists
    '''
    if settings.sheet(name/key/url) == '':
        printWarning('No name, key, or url have been provided')
        exit()
    try: # cascading open by name, url, key
        # TODO: open_worksheet_url
        # TODO: open_worksheet_key
    except:
        # TODO: prompt user to either check the settings file or have the sheet generated
    '''

    # check worksheets or generate
    '''
    try:
        open_worksheet control panel
    except:
        would you like to generate the control panel?

    try:
        open_worksheet roster
    except:
        would you like to generate a new roster?

    try:
        open_worksheet signins
    except:
        you may want to add a sign in sheet, this isn't auto yet
    
    # TODO: the subroster rules sheet
    '''
    
    # Add the Administrator_Email as a collaborator
    '''
    There is a gspread call for this
    something like spreadsheet.add_collaborator(email)
    '''
    
    # Update settings.cfg
    Write_Settings_File(settings_filename)
    
    # run first time updates on worksheet
    
    return
def Generate_Settings_File(filename = 'settings.cfg'):
    settings = open(filename, 'w')
    settings.write("# Only one of the following needs to be defined\n")
    settings.write("# However, defining multiple creates redundancy\n")
    settings.write("# Failing to set any of these will prompt a first time setup to be run\n")
    settings.write("Google_Sheet_Name=\n")
    settings.write("Google_Sheet_key=\n")
    settings.write("Google_Sheet_URL=\n\n")
    
    settings.write("# Email account to invite to the sheet if it is generated by the program\n")
    settings.write("# Google Email (gmail) accounts are recommended for this\n")
    settings.write("# This value is required for first time setup to be run\n")
    settings.write("Administrator_Email=\n")
    settings.close()
    printSuccess('A new settings file has been generated')
    return
def Load_Settings_File(filename = 'settings.cfg'):
    loaded_settings = [line.split('=') for line in open(filename,'r').read().split('\n')]
    for setting in loaded_settings:
        if setting[0] == 'Google_Sheet_Name':
            settings_file.Google_Sheet_Name = setting[1]
        elif setting[0] == 'Google_Sheet_key':
            settings_file.Google_Sheet_key = setting[1]
        elif setting[0] == 'Google_Sheet_URL':
            settings_file.Google_Sheet_URL = setting[1]
        
        elif setting[0] == 'Administrator_Email':
            settings_file.Administrator_Email = setting[1]
        
    return
def Write_Settings_File(filename = 'settings.cfg'):
    # TODO:
    # open the old settings file
    # replace the values
    # write out
    return

# control panel functions
def Control_Value(cp_worksheet, str_setting):
    active_cell = cp_worksheet.find(str_setting)
    active_cell = cp_worksheet.cell(active_cell.row, active_cell.col +1)
    return active_cell
def sync_control_panel():
    control_panel = open_worksheet('Control Panel')
    
    # UPDATE TIMES
    
    # update check-in time
    active_cell = Control_Value(control_panel, 'Last Agent Check-in')
    control_panel.update_cell(active_cell.row, active_cell.col, time.asctime())
    
    # Last Roster Aggregation
    active_cell = Control_Value(control_panel, 'Last Roster Aggregation')
    if not active_cell.value == last_roster_aggregation_time:
        control_panel.update_cell(active_cell.row, active_cell.col, last_roster_aggregation_time)
    
    # Last Sign-in Processing
    active_cell = Control_Value(control_panel, 'Last Sign-in Processing')
    if not active_cell.value == last_sign_in_processing_time:
        control_panel.update_cell(active_cell.row, active_cell.col, last_sign_in_processing_time)
    
    # Last Processed sign in (last_sign_in_processed_timestamp)
    global last_sign_in_processed_timestamp
    active_cell = Control_Value(control_panel, 'Last Processed Sign in TS')
    if active_cell.value == '-':
        control_panel.update_cell(active_cell.row, active_cell.col, '01/01/0001 00:00:00')
        active_cell = Control_Value(control_panel, 'Last Processed Sign in TS')
    if last_sign_in_processed_timestamp != None and str2dt(active_cell.value) < last_sign_in_processed_timestamp:
            control_panel.update_cell(active_cell.row, active_cell.col, last_sign_in_processed_timestamp)
    else:
        last_sign_in_processed_timestamp = active_cell.value
        last_sign_in_processed_timestamp = str2dt(last_sign_in_processed_timestamp)
    
    # Last SubRoster Generation
    active_cell = Control_Value(control_panel, 'Last SubRoster Generation')
    if not active_cell.value == last_subroster_generation_time:
        control_panel.update_cell(active_cell.row, active_cell.col, last_subroster_generation_time)
    
    
    # update startup time
    active_cell = Control_Value(control_panel, 'Last Agent Startup Time')
    if not active_cell.value == agent_start_time:
        control_panel.update_cell(active_cell.row, active_cell.col, agent_start_time)

    # Current Semester
    global current_semester
    active_cell = Control_Value(control_panel, 'Current Semester')
    if active_cell.value == '-':
        control_panel.update_cell(active_cell.row, active_cell.col, 'spring 0001')
        active_cell = Control_Value(control_panel, 'Current Semester')
    if current_semester == None:
        current_semester = active_cell.value
    if current_semester != active_cell.value:
        control_panel.update_cell(active_cell.row, active_cell.col, current_semester)
    
    # AGENT CONFIG
    
    # Post-Interval sleep time
    global post_interval_sleep_time
    post_interval_sleep_time = Control_Value(control_panel, 'Post-Interval sleep time').value
    
    # Roster aggregation period
    global roster_aggregation_period
    active_cell = Control_Value(control_panel, 'Roster Aggregation Period')
    if not active_cell.value == roster_aggregation_period:
        roster_aggregation_period = active_cell.value
    
    # Check Sign-n period
    global check_signin_period
    active_cell = Control_Value(control_panel, 'Check Sign-in Period')
    if not active_cell.value == check_signin_period:
        check_signin_period = active_cell.value
    
    # Generate subroster period
    global generate_subroster_period
    active_cell = Control_Value(control_panel, 'Generate Subroster Period')
    if not active_cell.value == generate_subroster_period:
        generate_subroster_period = active_cell.value
    
    # Shut down command
    active_cell = Control_Value(control_panel, 'Shutdown next interval?')
    if active_cell.value.lower() == 'yes':
        printWarning('Shutdown command has been recieved from the control panel.')
        exit()
    
    # Fall Semester start date (if not already read)
    global fall_semester_start_date
    if fall_semester_start_date is None:
        active_cell = Control_Value(control_panel, 'Fall Semester Start (MM/DD)')
        fall_semester_start_date = datetime.datetime.strptime(active_cell.value, '%m/%d')
        fall_semester_start_date = fall_semester_start_date.replace(year = datetime.datetime.now().year)
        
    # Spring semester start date (if not already read)
    global spring_semester_start_date
    if spring_semester_start_date is None:
        active_cell = Control_Value(control_panel, 'Spring Semester Start (MM/DD)')
        spring_semester_start_date = datetime.datetime.strptime(active_cell.value, '%m/%d')
        spring_semester_start_date = spring_semester_start_date.replace(year = datetime.datetime.now().year)
    
    # Summer semester start date (if not already read)
    global summer_semester_start_date
    if summer_semester_start_date is None:
        active_cell = Control_Value(control_panel, 'Summer Semester Start (MM/DD)')
        summer_semester_start_date = datetime.datetime.strptime(active_cell.value, '%m/%d')
        summer_semester_start_date = summer_semester_start_date.replace(year = datetime.datetime.now().year)

# Roster processing
def process_signins():
    printInfo('Processing recent sign ins')
    global last_sign_in_processing_time
    last_sign_in_processing_time = time.asctime()
    
    # load sign in sheet
    signin_sheet = open_worksheet('Sign ins')
    signins = signin_sheet.get_all_values()
    signins_timestamp = 0
    signins_m_num = 1
    signins_6_2 = 2
    signins_name = 3
    
    # skip the header to prevent errors
    signins = signins[1:]
        
    # filter out entries that have been processed
    recent_signins = []
    global last_sign_in_processed_timestamp
    for entry in signins:
        if entry[signins_timestamp] != '' and str2dt(entry[signins_timestamp]) > last_sign_in_processed_timestamp:
            recent_signins += [entry]
    
    # if there are no recent sign ins, stop
    if len(recent_signins) == 0:
        printSuccess('There are no new signins to process.')
        return
    
    # filter user input
    for entry in recent_signins:
        #name->title
        entry[3] = entry[3].title()
        
        #6+2->lower
        entry[2] = entry[2].lower()
        
        #M#->Upper and add M if not there
        entry[1] = entry[1].upper()
        if entry[1][0] != 'M':
            entry[1] = 'M' + entry[1]
    
    # Load in main roster
    ws_roster = open_worksheet('Full Roster')
    ws_roster_values_original = ws_roster.get_all_values()
    ws_roster_values_updated = ws_roster.get_all_values()
    ws_roster_header = ws_roster.get_all_values()[0]
    
    # Find index values of used fields 
    id_index = ws_roster_header.index('ID')
    name_index = ws_roster_header.index('Full Name')
    m_number_index = ws_roster_header.index('ID Num')
    join_date_index = ws_roster_header.index('Join Date')
    semester_attendance_index = ws_roster_header.index('Attendance This Semester')
    last_meeting_date_index = ws_roster_header.index('Last Meeting Date')
    total_attendance_index = ws_roster_header.index('Total Attendance')
    roster_size = len(ws_roster.get_all_values())
    
    # for each signin entry
    for entry in recent_signins:
        try:
            print('Looking for ' + str(entry[signins_name]))
            # look for existing entry (account for index diff)
            try: # by 6-2
                entry_index = ws_roster.find(entry[signins_6_2]).row - 1 
            except:
                try: # by M#
                    entry_index = ws_roster.find(entry[signins_m_num]).row - 1 
                except:
                    try: # by university email
                        entry_index = ws_roster.find(entry[signins_6_2] + '@mail.uc.edu').row - 1 
                    except:
                        raise
            
            print('Found')
            # add M num if we don't have it
            if ws_roster_values_updated[entry_index][m_number_index] == '':
                ws_roster_values_updated[entry_index][m_number_index] == entry[signins_m_num]
            
            # add name if we don't have it
            if ws_roster_values_updated[entry_index][name_index] == '':
                ws_roster_values_updated[entry_index][name_index] == entry[signins_name]
                
        except:
            print('not found')
            # new sign ins -> create new entry
            # Don't worry about alphabetic position for now, that is what process_roster() is for
            new_entry = ['' for i in range(len(ws_roster_header))]
            new_entry[id_index] = entry[signins_6_2]
            new_entry[m_number_index] = entry[signins_m_num]
            new_entry[name_index] = entry[signins_name]
            new_entry[join_date_index] = entry[signins_timestamp]
            ws_roster_values_updated.append(new_entry)
            entry_index = len(ws_roster_values_updated) - 1
            
        finally:
            # old and new sign ins -> update number of meetings attended and last meeting date
            previous_attendance = int(str(0) + str(ws_roster_values_updated[entry_index][semester_attendance_index]))
            total_attendance = int(str(0) + str(ws_roster_values_updated[entry_index][total_attendance_index]))
            
            print(previous_attendance)
            print(total_attendance)
            
            ws_roster_values_updated[entry_index][semester_attendance_index] = previous_attendance + 1
            ws_roster_values_updated[entry_index][total_attendance_index] = total_attendance + 1
            ws_roster_values_updated[entry_index][last_meeting_date_index] = entry[signins_timestamp]
    
    # Update google sheet
    writeback(ws_roster, ws_roster_values_updated)
    
    printSuccess('Signins successfully proccesed.')
    
    # update the last processed sign-in time
    last_sign_in_processed_timestamp = str2dt(recent_signins[-1][signins_timestamp])
def process_roster():
    printInfo('Cleaning up the roster')

    # Load roster
    ws_roster = open_worksheet('Full Roster')
    ws_roster_values = ws_roster.get_all_values()
    ws_roster_values_original = list(ws_roster_values) # makes a copy vs referencing the same object
    ws_roster_header = ws_roster_values[0]
    ws_roster_values = ws_roster_values[1:]
    
    # Indexes of values from header
    ID = ws_roster_header.index('ID')
    Full_Name = ws_roster_header.index('Full Name')
    First_Name = ws_roster_header.index('First Name')
    Last_Name = ws_roster_header.index('Last Name')
    ID_Number = ws_roster_header.index('ID Num')
    Major = ws_roster_header.index('Major')
    Graduation_Year = ws_roster_header.index('Graduation Year')
    Email_UC = ws_roster_header.index('University Email')
    Email_Pref = ws_roster_header.index('Preferred Email')
    
    # for entry in roster:
    for entry in ws_roster_values:
        if entry != ['']:
            # Standardize fields
            entry[ID] = entry[ID].lower()
            
            entry[Full_Name] = entry [Full_Name].title()
            entry[First_Name] = entry [First_Name].title()
            entry[Last_Name] = entry [Last_Name].title()
            
            entry[Email_UC] = entry[Email_UC].lower()
            entry[Email_Pref] = entry[Email_Pref].lower()
            
            # If they have a full name but not a first and last, break into first and last
            if entry[First_Name] == '' and entry[Last_Name] == '' and entry[Full_Name] !='':
                name = entry[Full_Name].split()
                if len(name) == 2:
                    entry[First_Name] = name[0]
                    entry[Last_Name] = name[1]

            # if they do not have a full name, but do have seperate names, put those together
            if entry[Full_Name] == '' and (entry[First_Name] != '' or entry[Last_Name] != ''):
                entry[Full_Name] == entry[First_Name] + entry[Last_Name]

            # Add leading M to M number if they don't have it
            if entry[ID_Number] != '' and entry[ID_Number][0] != 'M':
                entry[ID_Number] = 'M' + entry[ID_Number]

            # If they have a preffered email that is UC domain but not a uc email, copy to uc email
            if 'uc.edu' in entry[Email_Pref] and entry[Email_UC] == '':
                entry[Email_UC] = entry[Email_Pref]

            # If they have a UC email but not a UC ID, remove '@mail.uc.edu' and set as UC ID
            if entry[ID] == '' and entry[Email_UC] != '':
                entry[ID] = entry[Email_UC][0:entry[Email_UC].index('@')]

            # If they have a UC id but not a UC email, append '@mail.uc.edu' and set as UC email
            if entry[ID] != '' and entry[Email_UC] == '':
                entry[Email_UC] = entry[ID] + '@mail.uc.edu'
            
            # If the preffered email address is their uc address, drop it
            if entry[Email_Pref] != '' and entry[Email_UC] == entry[Email_Pref]:
                entry[Email_Pref] = ''
    
    # Compile duplicates
    for entry in ws_roster_values:
        if entry[ID] != '':            
            # Find other entries with the same ID
            like_entries = [lentry for lentry in ws_roster_values if lentry != [''] and lentry[ID] == entry[ID]]
            
            # If all like entries have one common value for a field, share that
            for i in range(len(entry)):
                value_range = []
                [value_range.append(lentry[i]) for lentry in like_entries if lentry[i] != '' and lentry[i] not in value_range]
                
                if len(value_range) == 1:
                    entry[i] = value_range[0]

    # Remove duplicates and sort
    ws_roster_values_unique = []
    [ws_roster_values_unique.append(entry) for entry in ws_roster_values if not entry in ws_roster_values_unique]
    ws_roster_values_unique = [('\t').join(entry) for entry in ws_roster_values_unique]
    ws_roster_values_unique.sort()
    ws_roster_values_unique = [entry.split('\t') for entry in ws_roster_values_unique]
    
    
    # Writeback to the gsheet roster (only cells that have changed)
    ws_roster_values_unique = [ws_roster_header] + ws_roster_values_unique
    if not ws_roster_values_unique == ws_roster_values_original:
        writeback(ws_roster, ws_roster_values_unique)
    else:
        printSuccess('There is no change in the roster')
    
    
    # update last_roster_aggregation_time
    global last_roster_aggregation_time
    last_roster_aggregation_time = time.asctime()
    
    printSuccess('Roster aggreagtion complete, there are ' + str(len(ws_roster_values_unique)) + ' entries.')
    return
def generate_subrosters():
    printInfo('Generating subrosters')
    printWarning('Subroster generation has not been implemented yet')
    # Make the subroster worksheets if they don't exist
    
    # people who haven't attended any meetings in the last 2 semesters are considered inactive
    # mark as inactive in roster
    return
def update_semester():
    printInfo('The semester is being updated.')
    
    # load roster
    ws_roster = open_worksheet('Full Roster')
    ws_roster_values_original = ws_roster.get_all_values()
    ws_roster_values_updated = ws_roster.get_all_values()
    ws_roster_header = ws_roster_values_original[0]
    ws_roster_values_updated = ws_roster_values_updated[1:]
    
    # Roster field index values
    this_semester_index = ws_roster_header.index('Attendance This Semester')
    last_semester_index = ws_roster_header.index('Attendance Last Semester')
    
    for entry in ws_roster_values_updated:
        # move over current semester sign in counts to last semester
        entry[last_semester_index] = entry[this_semester_index]
        
        # reset current semester sign in
        entry[this_semester_index] = '0'
    
    # Writeback data
    writeback(ws_roster, ws_roster_values_updated)
    
    # update current semester
    global current_semester
    global spring_semester_start_date
    global summer_semester_start_date
    global fall_semester_start_date
    
    today = datetime.datetime.now()
    if spring_semester_start_date < today and today < summer_semester_start_date:
        current_semester = 'spring ' + str(today.year)
    elif summer_semester_start_date < today and today < fall_semester_start_date:
        current_semester = 'summer ' + str(today.year)
    elif fall_semester_start_date < today:
        current_semester = 'fall ' + str(today.year)
    printMessage('Current semester is now ' + current_semester)
    
    return

# Main Loop
def main():
    # pull in global variables
    global check_signin_period
    global roster_aggregation_period
    global generate_subroster_period
    global post_interval_sleep_time
    global agent_start_time
    global client
    global spreadsheet
    global current_semester
    global spring_semester_start_date
    global summer_semester_start_date
    global fall_semester_start_date
    
    # TODO: Add ability to parse command line input
    # --Client_Key to use different or multiple Oauth2 creds within the same dir 
    # --Settings_File to use different settings files
    # --Administrator_Email to provide email for first time setup so prompt isn't required
    # --fastSetup to skip prompting options on first time setup
    # --help -h -?
    # --dryRun to make sure the current settings are valid and the document can be connected

    # TODO: Check for updates and notify admin?
    # is there a python module for checking git versions?
    # if not should I try to parse 'git status' -> what if in a fork?
    # what about git status upstream or git status /url/

    # Take note of the starting time
    agent_start_time = time.asctime()
    printInfo('Starting College-Organization-Manager agent at ' + agent_start_time)
    
    # TODO Try to open the settings and worksheet
    #      except: run first time setup or another smaller setup segment
    '''
    try:
        Load_Settings_File(settings_filename)
        # client = start_client()
    except:
        First_Time_Setup(settings_filename)
    #exit()
    '''

    #Loop through tasks
    current_cycle = 0
    while True:
        printInfo('Current Cycle is ' + str(current_cycle))
        
        # Start the client (refreshes Authentication after long sleep intervals)
        client = start_client()
        spreadsheet = open_spreadsheet('Python Agent Master Sheet')
        
        # read/update the control panel to decide what to do next
        sync_control_panel()
        
        # Process all the roster data
        if current_cycle % int(roster_aggregation_period) == 0:
            process_roster()
        
        # look for sign ins 
        if current_cycle % int(check_signin_period) == 0:
            process_signins()
        
        # Generate subrosters
        if current_cycle % int(generate_subroster_period) == 0:
            pass
            #TODO: generate_subrosters()
            # generate subrosters based on the subroster definition worksheet
            
        # if today is a semester switch day and the semester has not been switched yet, switch it
        today = datetime.datetime.now()
        season = current_semester.split(' ')[0].lower()
        
        # if not season x and between season x dates
        if (
                (season != 'spring'   and spring_semester_start_date < today and today < summer_semester_start_date) or
                (season != 'summmer' and summer_semester_start_date < today and today < fall_semester_start_date) or
                (season != 'fall' and fall_semester_start_date < today)
            ):
            update_semester()
        else:
            printSuccess('The Semester has not changed')
        
        
        printInfo('Going to sleep for ' + post_interval_sleep_time + ' minute(s)')
        time.sleep(60 * float(post_interval_sleep_time))
        current_cycle = current_cycle + 1
    
    printInfo('End of program has been reached')
    return
    
# This prevents declaration order errors because all other functions/objects have been entered into the interpreter memory already
if __name__ == "__main__":
    try:
        main()
    except httplib2.ServerNotFoundError:
        printError('Could not connect to the specified server, try checking your internet connection.')
        raise
    except SystemExit:
        printWarning('Program is terminating.')
    except:
        printError('An handled error has occured.')
        raise
 

 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
