import os
import glob
import shutil
import subprocess
import time
import datetime
import re
import pandas as pd
import win32com.client
import threading
import pythoncom

#==============================
# Constants
#==============================
staff_id = os.getlogin()
cUser = os.path.expanduser('~')
current_directory = os.path.dirname(os.path.realpath(__file__))
previous_directory = os.path.abspath(os.path.join(current_directory, os.pardir))
temp_directory = "C:\\temp"

SAP_path = 'C:\\Program Files\\SAP\\FrontEnd\\SAPgui\\sapshcut.exe'
file_path_publicHolidays = os.path.join(current_directory, "publicHoliday.csv")
df_publicHolidays = pd.read_csv(file_path_publicHolidays, parse_dates=['Date'], date_format='%d-%m-%y')

#==============================
# R
#==============================
def robust_get_Rscript_path(base_path = r"C:\\Program Files\\", search_pattern = r"R-*\\bin\\Rscript.exe"):
    # Common base path for R installation
    search_pattern_full = os.path.join(base_path, search_pattern)  # Search for Rscript.exe in directories that match 'R-*'
    r_executable_paths = glob.glob(search_pattern_full)  # Use glob to find the executable
    
    # If multiple versions are found, sort to get the latest version
    if r_executable_paths:
        r_executable_paths.sort(reverse=True)
        return r_executable_paths[0]
    else:
        raise FileNotFoundError("Rscript.exe not found in expected directories.")

#==============================
# SAP
#==============================
def launch_SAP_session(SAP_path, staff_id, leave_initial_bool = False, UQ_mode = False, Anti_SAP_idle_test = True, Accept_SNC_logon = True, utilize_idle_session = True,
                       UP_name = "CCMS UP2  Production IS-UT 6.08", UQ_name = "CCMS UQ4  Quality Assurance IS-UT 6.08"):
    # Features:
    # 1. Through this function, you can achieve concurrent control among SAP session
    # 2. Dynamically launch idle SAP session, not limited to use the 0-th SAP session only, not affect the opening SAP sesssion you are using
 
    # Notes:
    # 1. leave_initial_bool is for leaving the first session to safely open new session, though other sessions is utilized if first session is busy
    # 2. UQ_mode is for testing, if it set to be true, SAP UQ Quality Assurance is launched instead  (False - PRD, True - UQ)
    # 3. Maximum number of SAP session opened concurrently is 6
    # 4. When Anti_SAP_idle test is on, the function shall automatically task kill SAP if the opening SAP get stuck before launching a new SAP session
    # 5. When Accept_SNC_logon is on, the function shall automatically accept SNC logon if the opening SAP requires SNC logon before launching a new SAP session
    # 6. When utilize_idle_session is on, the function shall utilize idle session if launching a new SAP session is not successful due to the maximum number of SAP session is reached
    # 7. This function requires users have logged in to SAP at least once before using this function

    # Launch SAP
    SAP_in_use_bool = False   # Check if user is using SAP
    initial_session = None
    try:  # Attempt to check if any SAP is opening
        SAP_GUI_AUTO = win32com.client.GetObject('SAPGUI')
        application = SAP_GUI_AUTO.GetScriptingEngine
        connection = application.Children(0)
        initial_session = connection.Children(0)
        SAP_in_use_bool = True  # User is using SAP beforehand
    except:
        SAP_in_use_bool = False
    time.sleep(0.01)

    # In case the SAP has timed out
    if Anti_SAP_idle_test and initial_session != None:
        try:  # Do some actions to ensure that the initial_session is in stable status
            print("The current initial session: " + initial_session.FindById("wnd[0]").Text)
            if "SAP Easy Access" in initial_session.FindById("wnd[0]").Text:
                initial_session.findById("wnd[0]").setFocus()
                initial_session.findById("wnd[0]/tbar[0]/okcd").text = "SAP IDLE TESTING 1"  # Do not change this value to any potential transaction code
                initial_session.findById("wnd[0]").sendVKey(0)
                initial_session.findById("wnd[0]/tbar[0]/okcd").text = ""
                if "SAP IDLE TESTING 1 does not exist" not in initial_session.findById("wnd[0]/sbar").text:
                    raise AssertionError("Anti SAP idle test does not pass.")
                initial_session.findById("wnd[0]/tbar[0]/okcd").text = "SAP IDLE TESTING 2"  # Do not change this value to any potential transaction code
                initial_session.findById("wnd[0]").sendVKey(0)
                initial_session.findById("wnd[0]/tbar[0]/okcd").text = ""
                if "SAP IDLE TESTING 2 does not exist" not in initial_session.findById("wnd[0]/sbar").text:
                    raise AssertionError("Anti SAP idle test does not pass.")
            else:
                memory_switch_on = False
                initial_session.findById("wnd[0]/mbar/menu[3]/menu[4]/menu[3]").select()  # Switch memory usage to another value (True / False)
                if "Memory consumption" not in initial_session.findById("wnd[0]/sbar").text:
                    raise AssertionError("Anti SAP idle test does not pass.")
                else:
                    if "switched on" in initial_session.findById("wnd[0]/sbar").text:
                        memory_switch_on = True
                initial_session.findById("wnd[0]/mbar/menu[3]/menu[4]/menu[3]").select()  # Switch memory usage to original value (False / True)
                if memory_switch_on:
                    if "switched off" not in initial_session.findById("wnd[0]/sbar").text:
                        raise AssertionError("Anti SAP idle test does not pass.")
                else:
                    if "switched on" not in initial_session.findById("wnd[0]/sbar").text:
                        raise AssertionError("Anti SAP idle test does not pass.")
        except:
            time.sleep(1)
            os.system('cmd /c "taskkill /f /im saplogon.exe"')
            SAP_in_use_bool = False
            time.sleep(1)

    time.sleep(0.01)
    if not SAP_in_use_bool:  # If opening SAP terminal cannot be found, launch a new SAP
        if UQ_mode:
            subprocess.call('"' + SAP_path + '" -desc=' + UQ_name + ' -client=100 -user=' + staff_id)
        else:
            subprocess.call('"' + SAP_path + '" -desc=' + UP_name + ' -client=100 -user=' + staff_id)
        # Dynamically wait for SAP to be launched
        for i in range(30):  # Maximum wait for launching a SAP object would be 30 * 1 = 30 seconds
            time.sleep(1)  # A short delay for session to get the SAP object
            try:
                SAP_GUI_AUTO = win32com.client.GetObject('SAPGUI')
                application = SAP_GUI_AUTO.GetScriptingEngine
                connection = application.Children(0)
                initial_session = connection.Children(0)
                SAP_in_use_bool = False  # User is not using SAP beforehand
                time.sleep(0.01)
                break
            except:
                pass

    # Term of Uses Check
    if Accept_SNC_logon and initial_session != None and initial_session.FindById("wnd[0]/sbar").text.startswith("SNC logon by 100"):
        initial_session.findById("wnd[0]").sendVKey(0)

    time.sleep(0.01)
    # Session Assignment
    target_session = None
    current_children_count = connection.Children.Count  # For example, if session(0) and session(1) are in opening, connection.Children.Count returns 2, so current_children_count = 2
    create_session_success = False
    if current_children_count == 0:
        print("Warning: SAP object did not launch successfully.")
    if current_children_count == 1 and not SAP_in_use_bool and not leave_initial_bool:  # If user is using SAP, don't use the sessions they are using
        target_session = initial_session
    else:
        # Dynamically launch new session
        if current_children_count < 6:  # Maximum number of SAP session opened concurrently is 6
            for i in range(6):
                last_children_count = current_children_count
                try:
                    connection.Children(i).createsession()  # We attempt to use any non-busy and current opening session to launch SAP
                except:
                    pass
                # Dynamically wait for new session to be launched
                for m in range(500):  # Maximum wait for a new session to be launched would be 500 * 0.01 = 5 seconds
                    current_children_count = connection.Children.Count
                    if current_children_count == last_children_count + 1:  # If we need to ensure the new session(2) for example has opened, we need check if connection.Children.Count returns 3 (current_children_count + 1)
                        break
                    else:
                        last_children_count = current_children_count  # In case in meanwhile, a session has closed, so we need to trace the number of session opened
                        time.sleep(0.01)  # A short delay for session to open
                try:
                    target_session = connection.Children(last_children_count)  # connection.children(n) refers to the session(n), if session(0), session(1) is opening, connection.Children.Count returns 2, so we need to put last_children_count
                    create_session_success = True
                    break
                except:
                    print("Warning: Failed to refer to newly created SAP session by SAP session (" + str(i) + ")")
        else:
            print("Warning: Maximum number of SAP session opened concurrently is 6, cannot open new SAP session.")
        if not create_session_success:
            print("Warning: Create new SAP session without success.")
    time.sleep(0.01)

    # Utilize Idle Session
    if not create_session_success and utilize_idle_session:  # If new SAP cannot be launched, it might be due to the maximum number of SAP session is reached, so we shall utilize idle session
        for i in range(1,6):  # We shall not use the 0-th session, as it might be used by user
            try:
                if "SAP Easy Access" in connection.Children(i).FindById("wnd[0]").Text:
                    target_session = connection.Children(i)
                    try:
                        if target_session.FindById("wnd[1]").text.endswith("100 Information"):
                            target_session.FindById("wnd[1]").sendVKey(0)
                    except:
                        pass
                    print("Warning: As new session cannot be created, utilize idle session (" + str(i) + ") instead.")
                    break
            except:
                pass
    time.sleep(0.01)

    return target_session, connection.Children.Count

def concurrent_SAP_session_iterator(SAP_path, staff_id, target_func, iterator, max_num_session = 6, leave_initial_bool = True, UQ_mode = False, Anti_SAP_idle_test = True, Accept_SNC_logon = True, utilize_idle_session = True):
    threads = []
    for params in iterator:
        session, n = launch_SAP_session(SAP_path, staff_id, leave_initial_bool, UQ_mode, Anti_SAP_idle_test, Accept_SNC_logon, utilize_idle_session)
        t = threading.Thread(target=target_func, args=(pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, session), *params))
        threads.append(t)
        t.start()
        if n >= max_num_session:
            for t in threads:
                t.join()
        time.sleep(1)
    for t in threads:
        t.join()

#==============================
# Encryption
#==============================
def encrypt(target_string, key_1, key_2, key_3):
    encrypted_string = ""
    alphas = "ABCDEFGHIJKLNMOPQRSTUVWXYZabcdefghijklnmopqrstuvwxyz0123456789~!@#$%^&*()_+`-={}|[]\:;<>?,./"
    shift_position = 0
    previous_position_culminative = 0
    target_string_extend = str(hash(os.getlogin()) % 90 + 10) + target_string
    for n in range(len(target_string_extend)):
        for i in range(len(alphas)):
            if target_string_extend[n] == alphas[i]:
                shift_position = (i + key_1 * previous_position_culminative + key_2 * n + key_3) % len(alphas)
                previous_position_culminative += shift_position
                break
        encrypted_string += alphas[shift_position]
    return encrypted_string

def decrypt(target_string, key_1, key_2, key_3):
    decrypted_string = ""
    alphas = "ABCDEFGHIJKLNMOPQRSTUVWXYZabcdefghijklnmopqrstuvwxyz0123456789~!@#$%^&*()_+`-={}|[]\:;<>?,./"
    shift_position = 0
    previous_position_culminative = 0
    for n in range(len(target_string)):
        for i in range(len(alphas)):
            shift_position = (i + key_1 * previous_position_culminative + key_2 * n + key_3) % len(alphas)
            if target_string[n] == alphas[shift_position]:    
                previous_position_culminative += shift_position
                break
        decrypted_string += alphas[i]
    return decrypted_string[2:]

#==============================
# File Manipulations
#==============================
def Excel_SpreadSheet_SaveAsExcel(output_file_path, target_workbook_name):
    excel = win32com.client.Dispatch("Excel.Application")
    for workbook in excel.Workbooks:
        if target_workbook_name in workbook.Name:
            break
    if os.path.exists(output_file_path):
        os.remove(output_file_path)  # Remove the existing file to avoid overwrite prompt
    workbook.SaveAs(output_file_path)
    workbook.Close(SaveChanges=True)
    print("Exported the Excel SpreadSheet to Excel file: ", output_file_path)

def excel_to_html(xl_path, sheet, text_range):
    # xl_path：excel file path    # xl_path：excel文件路径
    # sheet: excel file sheet     # sheet: 要复制的sheet
    # text_range: copy range      # text_range: 复制范围
    html_path = r'C:\\Temp\\tmp.html'    # Temporary HTML file   # 临时HTML文件
    xl_temp_path = r'C:\\Temp\\tmp.xlsx'
    print('html_path: '+ html_path)
    if os.path.exists(html_path):
        os.remove(html_path)
    if os.path.exists(xl_temp_path):
        os.remove(xl_temp_path)
    shutil.copy(xl_path, xl_temp_path)

    ExcelAPP = win32com.client.DispatchEx('Excel.Application')
    WordApp = win32com.client.DispatchEx("Word.Application")
    ExcelAPP.Visible = False
    ExcelAPP.DisplayAlerts = False
    WordApp.Visible = False
    WordApp.DisplayAlerts = False

    doc = WordApp.Documents.Add()
    book = ExcelAPP.Workbooks.Open(xl_temp_path)
    sht = book.Worksheets(sheet)
    sht.Range(text_range).Copy()
    # Copy to word  # 先贴到Word
    doc.Content.PasteExcelTable(False, False, False)
    # Word to HML   # 再把Word存为HTML格式
    doc.SaveAs(html_path, FileFormat=10)
    ExcelAPP.Workbooks.Close()
    ExcelAPP.Application.Quit()
    WordApp.Documents.Close()
    WordApp.Application.Quit()

    f = open(html_path, "r")
    text_html = f.read()
    f.close()
    return text_html

def colnum_to_excel_col(col_num):
    # Convert column number to Excel letter (e.g., 1 -> 'A', 27 -> 'AA')
    col_str = ''
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        col_str = chr(65 + remainder) + col_str
    return col_str

def get_latest_file(file_path_format):
    list_of_files = glob.glob(file_path_format)
    if list_of_files:
        latest_file = max(list_of_files, key=os.path.getctime)
        return latest_file
    return None

def clear_target_files_in_folder(target_folder, start_match="", end_match=""):
    print("Clearing files in the target folder: " + target_folder)
    if os.path.exists(target_folder):
        for file_name in [file_name for file_name in os.listdir(target_folder) if file_name.startswith(start_match) and file_name.endswith(end_match)]:
            file_path = os.path.join(target_folder, file_name)
            try:
                os.remove(file_path)
                print('Deleted file: ' + file_path)
            except FileNotFoundError:
                print(file_path + ' not found or has already been deleted. Check the file path.')
                continue
            except Exception as e:
                print(e)
    else:
        print('The target folder does not exist: ' + target_folder)

def split_input_list(input_list_path, queue_folder, num_device_each = 1999, file_name = "inputlist", file_format = ".txt"):
    print("Splitting the input list file...")
    df_input_list = pd.read_csv(input_list_path, header=None)
    z = []
    a = 1
    if len(df_input_list) < num_device_each:
        df_input_list.to_csv(os.path.join(queue_folder, file_name + "_1" + file_format), header=False, index=False, quoting=False)
    else:
        while len(df_input_list) >= num_device_each:
            z.append(df_input_list.iloc[0:num_device_each, :])
            z[a-1].to_csv(os.path.join(queue_folder, file_name + "_" + str(a) + file_format), header=False, index=False, quoting=False)
            a += 1
            df_input_list = df_input_list.iloc[num_device_each:, :]
            if len(df_input_list) < num_device_each:
                df_input_list.to_csv(os.path.join(queue_folder, file_name + "_" + str(a) + file_format), header=False, index=False, quoting=False)
            else:
                continue
    print("The input list file has been splitted into " + str(a) + " files.")

def rename_existing_file_to_contain_postfix(file_path, staff_id_indicator = True):
    if os.path.exists(file_path):
        placing_folder = os.path.dirname(file_path)
        file_name = os.path.splitext(os.path.basename(file_path))[0]
        file_extension = os.path.splitext(file_path)[1]
        if staff_id_indicator:
            staff_id = os.getlogin()
            previous_file_with_staff_id = glob.glob(os.path.join(placing_folder, file_name + "_" + staff_id + "*"))
            seq = len(previous_file_with_staff_id)
            new_file_path = os.path.join(placing_folder, file_name + "_" + staff_id + "_" + str(seq) + file_extension)
        else:
            previous_file = glob.glob(os.path.join(placing_folder, file_name + "*"))
            seq = len(previous_file)
            new_file_path = os.path.join(placing_folder, file_name + "_" + str(seq) + file_extension)
        os.rename(file_path, new_file_path)
        print("Renamed " + file_path + " to " + new_file_path)

def strip_whitespace(cell):
    return cell.strip() if isinstance(cell, str) else cell

def get_latest_file_by_creation_time(path):    
    files = os.listdir(path)
    paths = [os.path.join(path,basename) for basename in files]
    return max(paths, key=os.path.getctime)

def get_latest_file_by_modified_time(path):    
    files = os.listdir(path)
    paths = [os.path.join(path,basename) for basename in files]
    return max(paths, key=os.path.getmtime)

#==============================
# Date Time Calculation
#==============================
def next_working_date(base_day, df_publicHolidays):
    result_day = base_day
    while True:
        result_day += datetime.timedelta(days=1)
        if result_day in df_publicHolidays or result_day.weekday() == 6 or result_day.weekday() == 0:  # Weekday is 0 for Monday, 6 for Sunday
            pass
        else:
            break
    return result_day

def last_working_date(base_day, df_publicHolidays):
    result_day = base_day
    while True:
        result_day -= datetime.timedelta(days=1)
        if result_day in df_publicHolidays or result_day.weekday() == 6 or result_day.weekday() == 0:  # Weekday is 0 for Monday, 6 for Sunday
            pass
        else:
            break
    return result_day

def get_working_date_difference(start_day, end_day, df_publicHolidays):
    if start_day > end_day:
        start_day, end_day = end_day, start_day
    n = 0
    loop_day = start_day
    while loop_day < end_day:
        loop_day += datetime.timedelta(days=1)
        if loop_day in df_publicHolidays or loop_day.weekday() == 6 or loop_day.weekday() == 0:  # Weekday is 0 for Monday, 6 for Sunday
            pass
        else:
            n += 1
    return n

def is_valid_date_yyyymmdd(date_str):
    if len(date_str) != 8 or not date_str.isdigit():
        return False

    year = int(date_str[:4])
    month = int(date_str[4:6])
    day = int(date_str[6:])

    if month < 1 or month > 12:
        return False

    if day < 1 or day > 31:
        return False

    if year < 1:
        return False

    if month in [4, 6, 9, 11] and day > 30:
        return False

    if month == 2:
        if (year % 4 == 0 and year % 100 != 0) or year % 400 == 0:
            if day > 29:
                return False
        elif day > 28:
            return False

    return True

def is_valid_date_ddmmyyyy_with_dot(date_str):
    if not re.match(r'\d{2}\.\d{2}\.\d{4}', date_str):
        return False

    day, month, year = map(int, date_str.split('.'))

    if month < 1 or month > 12:
        return False

    if day < 1 or day > 31:
        return False

    if year < 1:
        return False

    if month in [4, 6, 9, 11] and day > 30:
        return False

    if month == 2:
        if (year % 4 == 0 and year % 100 != 0) or year % 400 == 0:
            if day > 29:
                return False
        elif day > 28:
            return False

    return True

#==============================
# Integers
#==============================
def is_integer(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

# string to integer
def S2I(s):
    try:
        return int(s)
    except Exception as e:
        print("Warning: Input is not a valid integer. Exception: " + str(e))
        return s
    
#========================
# Outlook Email
#========================
def search_outlook_emails(subject_keyword, folder_id=6, message_alert=True, return_single_result=False, loop_from_latest=False):
    # Search for emails in the specified Outlook folder with a subject containing the given keyword.
    # subject_keyword: string to search for in the email subject
    # folder_id: Outlook folder ID (default 6 = Inbox); please refer to: https://learn.microsoft.com/en-us/office/vba/api/outlook.oldefaultfolders
    # message_alert: whether to print messages when emails are found or not found
    # return_single_result: if True, return only the first found email
    # loop_from_latest: if True, search emails from latest to oldest

    pythoncom.CoInitialize()
    outlook = win32com.client.Dispatch('outlook.application').GetNamespace("MAPI")
    folder = outlook.GetDefaultFolder(folder_id)
    found_records = []

    if loop_from_latest:
        items = reversed(folder.Items)  # Search from latest to oldest
    else:
        items = folder.Items  # Search from oldest to latest

    for item in items:
        if subject_keyword in item.Subject:
            sent_time = item.SentOn
            if message_alert:
                print("Email found: " + item.Subject + " (sent at " + sent_time.strftime('%Y-%m-%d %H:%M:%S') + ")")
            found_records.append(item)
            if return_single_result:
                break
    if message_alert and not found_records:
        print("No email with the specified subject found.")
    return found_records

def extract_replied_content(email_body):
    # Extract the replied content from an email body, excluding the original message.
    # This function assumes that the original message starts with "From:" or similar markers.
    if "From:" in email_body:
        return email_body.split("From:", 1)[0].strip()
    else:
        return email_body.strip()

SENTIMENT_ANALYSIS_ACCEPT = 0
SENTIMENT_ANALYSIS_REJECT = 1
SENTIMENT_ANALYSIS_UNCERTAIN = 2
SENTIMENT_ANALYSIS_UNKNOWN = 3
def sentiment_analysis_text_acceptance(text, message_alert=True):
    # Simple sentiment analysis to check for acceptance/approval in the text
    accept_keywords = [
        "approve", "approved", "approval", "go", "go ahead", "agreed", "consent", "support", "yes", "ok", "okay", "confirmed", "confirm", "accepted", "accept", "fine", "sure", "sounds good",
        "noted", "not a problem", "no objection", "proceed", "endorsed", "endorse", "alright", "all right", "looks good", "looks fine", "looks ok", "looks okay", "works for me", "good to go", "clear",
        "no issues", "no problem", "no concerns", "no objections", "green light", "go for it", "sounds fine", "sounds ok", "sounds okay", "sounds acceptable", "all set", "all clear",
        "happy with this", "happy to proceed", "happy to approve", "agrees", "agreement", "in favor", "in agreement", "positive", "positively", "endorses", "endorsing", "endorsed", "will do",
        "done", "not opposed", "no opposition", "no reason not to", "no reason against", "no reason to object", "no reason to reject", "no reason to decline", "no reason to disagree", "no reason to withhold",
        "but", "however", "and", "also"
    ]
    reject_keywords = [
        "reject", "rejected", "not approve", "decline", "disagree", "no", "denied", "deny", "not agreed", "not consent", "not support", "no ok", "not ok", "not okay", "no okay",
        "not confirmed", "not accept", "not fine", "not sure", "object", "objection", "cannot", "can't", "refuse", "refused", "withhold", "withheld",
        "not endorsed", "not endorse", "no endorse", "noted with concern", 
        "concerned", "concerns", "issue", "issues", "problem", "problems", "not happy", "not satisfied", "not acceptable", "not possible", "not recommended", "not advisable", "not agreed", "not in favor",
        "not in agreement", "negative", "negatively", "opposed", "opposes", "opposing", "opposition", "against", "disagreement", "disagrees", "disapproves", "disapproved", "disapproving", "not allowed",
        "not permitted", "not authorized", "not authorized to proceed", "not authorized to approve",
        "not authorized to accept", "not authorized to endorse", "not authorized to support", "not authorized to consent",
        "not authorized to agree", "not authorized to confirm", "not authorized to proceed", "not authorized to go ahead",
        "but", "however", "and", "also"
    ]
    text_lower = text.lower()
    is_approved = any(keyword in text_lower for keyword in accept_keywords)
    is_rejected = any(keyword in text_lower for keyword in reject_keywords)
    if is_approved and not is_rejected:
        if message_alert: print("Sentiment Analysis: Approved")
        return SENTIMENT_ANALYSIS_ACCEPT
    elif is_rejected and not is_approved:
        if message_alert: print("Sentiment Analysis: Not Approved")
        return SENTIMENT_ANALYSIS_REJECT
    elif is_approved and is_rejected:
        if message_alert: print("Sentiment Analysis: Mixed/Unclear")
        return SENTIMENT_ANALYSIS_UNCERTAIN
    else:
        if message_alert: print("Sentiment Analysis: Unable to determine approval status")
        return SENTIMENT_ANALYSIS_UNKNOWN

#========================
# Custom
#========================
def df_to_list(df, file_path):
    with open(file_path, 'w') as f:
        df_string = df.to_string(header=False, index=False)
        f.write(df_string)