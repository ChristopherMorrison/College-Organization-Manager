# Imports
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import datetime

# debug variables
dbo = None

# API details
scope = ["https://spreadsheets.google.com/feeds"]
creds = ServiceAccountCredentials.from_json_keyfile_name("calico_client_secret.json",scope)


# Global object variables
client      = None
spreadsheet = None


# State variables
client_authorized  = False
spreadsheet_opened = False


# Control Panel Variables
last_roster_aggregation_time     = 'Not run yet'
last_sign_in_processing_time     = 'Not run yet'
last_sign_in_processed_timestamp = None
last_subroster_generation_time   = 'Not run yet'
agent_start_time = None

post_interval_sleep_time   = 1 
roster_aggregation_period  = 10
check_signin_period        = 10
generate_subroster_period  = 10
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
    print('------>' + bcolors.HEADER + str_message + bcolors.ENDC)


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
        printError('Could not authorize the client instance, try checking the Oauth2 credential json.\n')
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
        printSuccess('Worksheet ' + worksheet_name + ' opened')
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


# control panel functions
def Control_Value(cp_worksheet, str_setting):
    active_cell = cp_worksheet.find(str_setting)
    active_cell = cp_worksheet.cell(active_cell.row, active_cell.col +1)
    return active_cell
def sync_control_panel():
    control_panel = open_worksheet('Control Panel')
    
    # UPDATE TIMES
    
    # update check-in time
    active_cell = control_panel.find('Last Agent Check-in')
    control_panel.update_cell(active_cell.row, active_cell.col + 1, time.asctime())
    
    # Last Roster Aggregation
    active_cell = control_panel.find('Last Roster Aggregation')
    if not active_cell.value == last_roster_aggregation_time:
        control_panel.update_cell(active_cell.row, active_cell.col + 1, last_roster_aggregation_time)
    
    # Last Sign-in Processing
    active_cell = control_panel.find('Last Sign-in Processing')
    if not active_cell.value == last_sign_in_processing_time:
        control_panel.update_cell(active_cell.row, active_cell.col + 1, last_sign_in_processing_time)
    
    # Last Processed sign in (last_sign_in_processed_timestamp)
    global last_sign_in_processed_timestamp
    active_cell = control_panel.find('Last Processed Sign in TS')
    active_cell = control_panel.cell(active_cell.row, active_cell.col + 1)
    if last_sign_in_processed_timestamp != None and str2dt(active_cell.value) < last_sign_in_processed_timestamp:
        control_panel.update_cell(active_cell.row, active_cell.col, last_sign_in_processed_timestamp)
    else:
        last_sign_in_processed_timestamp = active_cell.value
        last_sign_in_processed_timestamp = str2dt(last_sign_in_processed_timestamp)
    
    # Last SubRoster Generation
    active_cell = control_panel.find('Last SubRoster Generation')
    if not active_cell.value == last_subroster_generation_time:
        control_panel.update_cell(active_cell.row, active_cell.col + 1, last_subroster_generation_time)
    
    
    # update startup time
    active_cell = control_panel.find('Last Agent Startup Time')
    if not active_cell.value == agent_start_time:
        control_panel.update_cell(active_cell.row, active_cell.col + 1, agent_start_time)

    
    # AGENT CONFIG
    
    # Post-Interval sleep time
    global post_interval_sleep_time
    post_interval_sleep_time = Control_Value(control_panel, 'Post-Interval sleep time').value
    
    # Roster aggregation period
    global roster_aggregation_period
    active_cell = control_panel.find('Roster Aggregation Period')
    active_cell = control_panel.cell(active_cell.row, active_cell.col + 1)
    if not active_cell.value == roster_aggregation_period:
        roster_aggregation_period = active_cell.value
    
    # Check Sign-n period
    global check_signin_period
    active_cell = control_panel.find('Check Sign-in Period')
    active_cell = control_panel.cell(active_cell.row, active_cell.col + 1)
    if not active_cell.value == check_signin_period:
        check_signin_period = active_cell.value
    
    # Generate subroster period
    global generate_subroster_period
    active_cell = control_panel.find('Generate Subroster Period')
    active_cell = control_panel.cell(active_cell.row, active_cell.col + 1)
    if not active_cell.value == generate_subroster_period:
        generate_subroster_period = active_cell.value
    
    # Shut down command
    active_cell = control_panel.find('Shutdown next interval?')
    active_cell = control_panel.cell(active_cell.row, active_cell.col +1)
    if active_cell.value.lower() == 'yes':
        printWarning('Shutdown command has been recieved from the control panel.')
        exit()
    
    # Fall Semester start date (if not already read)
    global fall_semester_start_date
    if fall_semester_start_date is None:
        active_cell = control_panel.find('Fall Semester Start Date')
        active_cell = control_panel.cell(active_cell.row, active_cell.col + 1)
        fall_semester_start_date = active_cell.value
        
    # Spring semester start date (if not already read)
    global spring_semester_start_date
    if spring_semester_start_date is None:
        active_cell = control_panel.find('Spring Semester Start Date')
        active_cell = control_panel.cell(active_cell.row, active_cell.col + 1)
        spring_semester_start_date = active_cell.value
    
    # Summer semester start date (if not already read)
    global summer_semester_start_date
    if summer_semester_start_date is None:
        active_cell = control_panel.find('Summer Semester Start Date')
        active_cell = control_panel.cell(active_cell.row, active_cell.col + 1)
        summer_semester_start_date = active_cell.value


# Roster processing
def process_signins():
    printInfo('Processing recent sign ins')
    global last_sign_in_processing_time
    last_sign_in_processing_time = time.asctime()
    
    # load sign in sheet into memory
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
        if entry[signins_timestamp] != '' and str2dt(entry[signin_timestamp]) > last_sign_in_processed_timestamp:
            recent_signins += [entry]
    
    # if there are no recent sign ins, stop
    if len(recent_signins) == 0:
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
    
    # update Roster with sign in data
    ws_roster = open_worksheet('Full Roster')
    ws_roster_header = ws_roster.get_all_values()[0]
    
    # THESE ARE 1 INDEXED BECAUSE OF THE API
    id_index = 1 + ws_roster_header.index('ID')
    name_index = 1 + ws_roster_header.index('Full Name')
    m_number_index = 1 + ws_roster_header.index('ID Num')
    join_date_index = 1 + ws_roster_header.index('Join Date')
    semester_attendance_index = 1 + ws_roster_header.index('Attendance This Semester')
    last_meeting_date_index = 1 + ws_roster_header.index('Last Meeting Date')
    total_attendance_index = 1 + ws_roster_header.index('Total Attendance')
    roster_size = len(ws_roster.get_all_values())
    
    # for each signin entry
    for signin in recent_signins:
        try:
            # look for existing entry
            entry_index = ws_roster.find(entry[signins_6_2]).row
            if ws_roster.cell(entry_index, m_number_index).value == '':
                ws_roster.update_cell(entry_index, m_number_index, entry[1])
            if ws_roster.cell(entry_index, name_index).value == '':
                ws_roster.update_cell(entry_index, name_index, entry[3])
        except:
            # new sign ins -> create new entry
            # using .insert_row is NOT the way to do this
            roster_size = roster_size + 1
            entry_index = roster_size
            ws_roster.update_cell(entry_index, id_index, entry[signins_6_2])
            ws_roster.update_cell(entry_index, m_number_index, entry[signins_m_num])
            ws_roster.update_cell(entry_index, name_index, entry[signins_name])
            ws_roster.update_cell(entry_index, join_date_index, entry[signins_timestamp])
        finally:
            # old and new sign ins -> update number of meetings attended and last meeting date
            previous_attendance = int(str(0) + ws_roster.cell(entry_index, semester_attendance_index).value)
            total_attendance = int(str(0) + ws_roster.cell(entry_index, total_attendance_index).value)
            
            ws_roster.update_cell(entry_index, semester_attendance_index, previous_attendance + 1 )
            ws_roster.update_cell(entry_index, total_attendance_index, total_attendance + 1 )
            ws_roster.update_cell(entry_index, last_meeting_date_index, entry[singins_timestamp])
            
    
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
    
    # The row and cell indexing will be tricky here because of the combination of 0 and 1 indexing
    # for the maximum number of rows we have
    
    printWarning('Now undergoing roster writeback, this may take a long time')
    
    for row in range(min([len(ws_roster_values_original), len(ws_roster_values_unique)])):
        printInfo('Updating row ' + str(row + 1))
        # if the cells are not equal, writeback to the roster, elif will not work here becuase of the indexing
        if ws_roster_values_original[row] != ws_roster_values_unique[row]:
            #update cell for cell in range if unique!=original
            for cell in range(len(ws_roster_values_unique[row])):
                printInfo('comparing cell ' + str(cell+1) + ' of ' + str(len(ws_roster_values_unique[row])))
                if ws_roster_values_unique[row][cell] != ws_roster_values_original[row][cell]:
                    printInfo('Inserting ' + ws_roster_values_unique[row][cell])
                    ws_roster.update_cell(row + 1, cell + 1, ws_roster_values_unique[row][cell])
            
            #[ws_roster.update_cell(row + 1, cell + 1, ws_roster_values_unique[row][cell]) for cell in range(len(ws_roster_values_unique[row])) if ws_roster_values_unique[row][cell] != ws_roster_values_original[row][cell]]
    
    # if we are out of unique values, clear the remaining rows
    if len(ws_roster_values_original) > len(ws_roster_values_unique):
        offset = len(ws_roster_values_unique)
        for row in range(len(ws_roster_values_original) - offset):
            printInfo('Clearing row ' + str(row + offset))
            [ws_roster.update_cell(row + offset + 1 , cell + 1, '') for cell in 
            range(len(ws_roster_header))]
    
    # update last_roster_aggregation_time
    global last_roster_aggregation_time
    last_roster_aggregation_time = time.asctime()
    
    return
def generate_subrosters():
    printInfo('Generating subrosters')
    printWarning('Subroster generation has not been implemented yet')
    # Make the subroster worksheets if they don't exist
    
    # people who haven't attended any meetings in the last 2 semesters are considered inactive
    # mark as inactive in roster
    return
def update_semester():
    return
# API instance
def main():
    # pull in global variables
    global check_signin_period
    global roster_aggregation_period
    global generate_subroster_period
    global post_interval_sleep_time
    global agent_start_time
    global client
    global spreadsheet
    
    # Take note of the starting time
    agent_start_time = time.asctime()
    printInfo('Starting CALICO agent at ' + agent_start_time)
    
    
    #Loop through tasks
    current_cycle = 0
    while True:
        printInfo('Current Cycle is ' + str(current_cycle))
        
        # Start the client (refreshes Authentication after long sleep intervals)
        client = start_client()
        spreadsheet = open_spreadsheet('Python Agent Master Sheet')
        
        # read/update the control panel to decide what to do next
        sync_control_panel()
        
        # look for sign ins 
        if current_cycle % int(check_signin_period) == 0:
            process_signins()
        
        # Process all the roster data
        if current_cycle % int(roster_aggregation_period) == 0:
            process_roster()
            
        # Generate subrosters
        if current_cycle % int(generate_subroster_period) == 0:
            generate_subrosters()
            
        # if today is a semester switch day and the semester has not been switched yet, switch it
        # if is fall and in spring range
        # if is spring and in summer range
        # if is summer and in fall range
        
        
        printInfo('Going to sleep for ' + post_interval_sleep_time + ' minute(s)')
        time.sleep(60 * float(post_interval_sleep_time))
        current_cycle = current_cycle + 1
    
    printInfo('End of program has been reached')
    return
    
# This prevents prototyping errors because all other functions/objects have been entered into the interpreter memory
if __name__ == "__main__":
    try:
        main()
    except:
        print('')
        printWarning('Program is terminating')
        raise
 

 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
