# Imports
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import time
import datetime

# debug variables
dbo = None

# API details
scope = ["https://spreadsheets.google.com/feeds"]
creds = ServiceAccountCredentials.from_json_keyfile_name("client_secret.json",scope)

# Global object variables
client      = None
spreadsheet = None

# State variables
client_authorized  = False
spreadsheet_opened = False

# Control Panel Variables
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
            # TODO look for miss by preffered email in a series of try catch
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
    
    # writeback to the roster
    
    printWarning('Now undergoing roster writeback, this may take a long time')
    
    # Select Cells covered by the original data
    start_range = ws_roster.get_addr_int(1,1)
    end_range = ws_roster.get_addr_int(len(ws_roster_values_updated), len(ws_roster_values_updated[0]))
    active_cells = ws_roster.range(start_range + ':' + end_range)
    
    # Break active cells from flat list to list list
    active_cells = [active_cells[x:x+len(ws_roster_header)] for x in range(0, len(active_cells), len(ws_roster_header))]
    
    # Update values from updated data
    for row in range(len(ws_roster_values_updated)):
        for col in range(len(ws_roster_header)):
            active_cells[row][col].value = ws_roster_values_updated[row][col]
        
    # flatten active_cells list
    temp = []
    for entry in active_cells:
        temp+=entry
    active_cells = temp
    
    # .update_cells batch call
    ws_roster.update_cells(active_cells)
    
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
    
    # The row and cell indexing will be tricky here because of the combination of 0 and 1 indexing
    # for the maximum number of rows we have
    
    printWarning('Now undergoing roster writeback, this may take a long time')
    
    # Select Cells covered by the original data
    start_range = ws_roster.get_addr_int(1,1)
    end_range = ws_roster.get_addr_int(len(ws_roster_values_original), len(ws_roster_values_original[0]))
    active_cells = ws_roster.range(start_range + ':' + end_range)
    active_cells = [active_cells[x:x+len(ws_roster_header)] for x in range(0, len(active_cells), len(ws_roster_header))]
    
    # Update values from unique data
    for row in range(len(ws_roster_values_unique)):
        for col in range(len(ws_roster_header)):
            active_cells[row][col].value = ws_roster_values_unique[row][col]
    
    
    # Clear hanging original data values
    offset = len(ws_roster_values_unique)
    for row in range(len(ws_roster_values_original) - offset):
        for col in range(len(ws_roster_header)):
            active_cells[row + offset][col].value = ''
    
    # flatten active_cells list
    temp = []
    for entry in active_cells:
        temp+=entry
    active_cells = temp
    
    # .update_cells batch call
    ws_roster.update_cells(active_cells)
    
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
    
    
    this_semester_index = ws_roster_header.index('Attendance This Semester')
    last_semester_index = ws_roster_header.index('Attendance Last Semester')
    
    for entry in ws_roster_values_updated:
        # move over current semester sign in counts to last semester
        entry[last_semester_index] = entry[this_semester_index]
        
        # reset current semester sign in
        entry[this_semester_index] = '0'
    
    # writeback roster    
    ws_roster_values_updated = [ws_roster_header] + ws_roster_values_updated    
    
    printWarning('Now undergoing roster writeback, this may take a long time')
    
    # Select Cells covered by the original data
    start_range = ws_roster.get_addr_int(1,1)
    end_range = ws_roster.get_addr_int(len(ws_roster_values_updated), len(ws_roster_values_updated[0]))
    active_cells = ws_roster.range(start_range + ':' + end_range)
    
    # Break active cells from flat list to list list
    active_cells = [active_cells[x:x+len(ws_roster_header)] for x in range(0, len(active_cells), len(ws_roster_header))]
    
    # Update values from updated data
    for row in range(len(ws_roster_values_updated)):
        for col in range(len(ws_roster_header)):
            active_cells[row][col].value = ws_roster_values_updated[row][col]
        
    # flatten active_cells list
    temp = []
    for entry in active_cells:
        temp+=entry
    active_cells = temp
    
    # .update_cells batch call
    #ws_roster.update_cells(active_cells)
    
    printSuccess('Signins successfully proccesed.')
    '''====================================='''
    
    # update current semester
    global current_semester
    global spring_semester_start_date
    global summer_semester_start_date
    global fall_semester_start_date
    
    today = datetime.datetime.now()
    if spring_semester_start_date < today:#TODO and today < summer_semester_start_date:
        current_semester = 'spring ' + str(today.year)
    elif summer_semester_start_date < today and today < fall_semester_start_date:
        current_semester = 'summer ' + str(today.year)
    elif fall_semester_start_date < today:
        current_semester = 'fall ' + str(today.year)
    printMessage('Current semester should be ' + current_semester)
    
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
            
        # if today is a semester switch day and the semester has not been switched yet, switch it
        today = datetime.datetime.now()
        season = current_semester.split(' ')[0].lower()
        
        # if not season x and between season x dates
        if (
                #(season != 'spring'   and spring_semester_start_date < today and today < summer_semester_start_date) or
                (season != 'spring'   and spring_semester_start_date < today) or
                (season != 'summmer' and summer_semester_start_date < today and today < fall_semester_start_date) or
                (season != 'fall' and fall_semester_start_date < today)
            ):
            update_semester()
        
        
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
 

 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
 
