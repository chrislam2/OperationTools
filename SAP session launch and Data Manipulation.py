import os
import glob
import shutil
import time
import datetime
import dateutil.relativedelta
import pandas as pd
import subprocess
import win32com.client as win32
import threading
import pythoncom

##############################
# Global Variables
##############################
login_name = os.getlogin()
cUser = os.path.expanduser('~')
current_directory = os.path.dirname(os.path.realpath(__file__))
previous_directory = os.path.abspath(os.path.join(current_directory, os.pardir))
temp_directory = "C:\\temp"

now = datetime.datetime.now()
today = datetime.datetime.today()
today_date_string = today.strftime("%Y%m%d")
now_string = now.strftime("%Y%m%d %H:%M:%S")

###############################
# Functions
###############################
def robust_get_Rscript_path():
    # Common base path for R installation
    base_path = r"C:\\Program Files\\"
    search_pattern = os.path.join(base_path, r"R-*\\bin\\Rscript.exe")  # Search for Rscript.exe in directories that match 'R-*'
    r_executable_paths = glob.glob(search_pattern)  # Use glob to find the executable
    
    # If multiple versions are found, sort to get the latest version
    if r_executable_paths:
        r_executable_paths.sort(reverse=True)
        return r_executable_paths[0]
    else:
        raise FileNotFoundError("Rscript.exe not found in expected directories.")

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

def is_integer(s):
    try:
        int(s)
        return True
    except ValueError:
        return False

def strip_whitespace(cell):
    return cell.strip() if isinstance(cell, str) else cell

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

def clear_target_files_in_folder(target_folder, start_match="", end_match=""):
    print("Clearing files in the target folder: " + target_folder)
    if os.path.exists(target_folder):
        for file_name in [file_name for file_name in os.listdir(target_folder) if file_name.startswith(start_match) and file_name.endswith(end_match)]:
            try:
                file_path = os.path.join(target_folder, file_name)
                os.remove(file_path)
                print('Deleted file: ' + file_path)
            except FileNotFoundError:
                print(file_path + ' not found or has already been deleted. Check the file path.')
                continue
            except Exception as e:
                print(e)
    else:
        print('The target folder does not exist: ' + target_folder)

def split_dataset_to_multiple(target_file_path, queue_folder, num_item_each = 1999):
    print("Splitting the file: " + target_file_path)
    df_target = pd.read_csv(target_file_path, header=None)
    z = []
    a = 1
    if len(df_target) < num_item_each:
        df_target.to_csv(os.path.join(queue_folder, "list_1.txt"), header=False, index=False, quoting=False)
    else:
        while len(df_target) >= num_item_each:
            z.append(df_target.iloc[0:num_item_each, :])
            z[a-1].to_csv(os.path.join(queue_folder, "list_" + str(a) + ".txt"), header=False, index=False, quoting=False)
            a += 1
            df_target = df_target.iloc[num_item_each:, :]
            if len(df_target) < num_item_each:
                df_target.to_csv(os.path.join(queue_folder, "list_" + str(a) + ".txt"), header=False, index=False, quoting=False)
            else:
                continue
    print("The file has been splitted into " + str(a) + " files.")

def excel_to_html(xl_path, sheet, text_range):
    # xl_path：excel file path
    # sheet: excel file sheet
    # text_range: copy range
    html_path = r'C:\\Temp\\tmp.html'    # Temporary HTML file
    print('html_path: '+ html_path)
    ExcelAPP = win32.DispatchEx('Excel.Application')
    WordApp = win32.DispatchEx("Word.Application")
    ExcelAPP.Visible = False
    ExcelAPP.DisplayAlerts = False
    WordApp.Visible = False
    WordApp.DisplayAlerts = False

    doc = WordApp.Documents.Add()
    book = ExcelAPP.Workbooks.Open(xl_path)
    sht = book.Worksheets(sheet)
    sht.Range(text_range).Copy()
    # Copy to word
    doc.Content.PasteExcelTable(False, False, False)
    # Word to HML
    doc.SaveAs(html_path, FileFormat=10)
    ExcelAPP.Workbooks.Close()
    ExcelAPP.Application.Quit()
    WordApp.Documents.Close()
    WordApp.Application.Quit()

    f = open(html_path, "r")
    text_html = f.read()
    f.close()
    return text_html

def launch_SAP_session(SAP_path, login_name, leave_initial_bool = False, UQ_mode = False, Anti_SAP_idle_test = True):
    # Features:
    # 1. Through this function, you can achieve concurrent control among SAP session
    # 2. Dynamically launch idle SAP session, not limited to use the 0-th SAP session only, not affect the opening SAP sesssion you are using
 
    # Notes:
    # 1. leave_initial_bool is for leaving the first session to safely open new session, though other sessions is utilized if first session is busy
    # 2. UQ_mode is for testing, if it set to be true, SAP UQ Quality Assurance is launched instead  (False - PRD, True - UQ)
    # 3. Maximum number of SAP session opened concurrently is 6
    # 4. When Anti SAP idle test is on, the function shall automatically task kill SAP if the opening SAP get stuck before launching a new SAP session
 
    # Import Libraries
    import os
    import time
    import subprocess
    import win32com.client

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
                    raise("Anti SAP idle test does not pass.")
                initial_session.findById("wnd[0]/tbar[0]/okcd").text = "SAP IDLE TESTING 2"  # Do not change this value to any potential transaction code
                initial_session.findById("wnd[0]").sendVKey(0)
                initial_session.findById("wnd[0]/tbar[0]/okcd").text = ""
                if "SAP IDLE TESTING 2 does not exist" not in initial_session.findById("wnd[0]/sbar").text:
                    raise("Anti SAP idle test does not pass.")
            else:
                memory_switch_on = False
                initial_session.findById("wnd[0]/mbar/menu[3]/menu[4]/menu[3]").select()  # Switch memory usage to another value (True / False)
                if "Memory consumption" not in initial_session.findById("wnd[0]/sbar").text:
                    raise("Anti SAP idle test does not pass.")
                else:
                    if "switched on" in initial_session.findById("wnd[0]/sbar").text:
                        memory_switch_on = True
                initial_session.findById("wnd[0]/mbar/menu[3]/menu[4]/menu[3]").select()  # Switch memory usage to original value (False / True)
                if memory_switch_on:
                    if "switched off" not in initial_session.findById("wnd[0]/sbar").text:
                        raise("Anti SAP idle test does not pass.")
                else:
                    if "switched on" not in initial_session.findById("wnd[0]/sbar").text:
                        raise("Anti SAP idle test does not pass.")
        except:
            os.system('cmd /c "taskkill /f /im saplogon.exe"')
            SAP_in_use_bool = False

    time.sleep(1)
    if not SAP_in_use_bool:  # If opening SAP terminal cannot be found, launch a new SAP
        if UQ_mode:
            subprocess.call('"' + SAP_path + '" -desc=CCMS UQ4  Quality Assurance IS-UT 6.08 -client=100 -user=' + login_name)
        else:
            subprocess.call('"' + SAP_path + '" -desc=CCMS UP2  Production IS-UT 6.08 -client=100 -user=' + login_name)
        # Dynamically wait for SAP to be launched
        for i in range(30):  # Maximum wait for launching a SAP object would be 30 * 1 = 30 seconds
            time.sleep(1)  # A short delay for session to get the SAP object
            try:
                SAP_GUI_AUTO = win32com.client.GetObject('SAPGUI')
                application = SAP_GUI_AUTO.GetScriptingEngine
                connection = application.Children(0)
                initial_session = connection.Children(0)
                SAP_in_use_bool = False  # User is not using SAP beforehand
                break
            except:
                pass

    # Session Assignment
    target_session = None
    current_children_count = connection.Children.Count  # For example, if session(0) and session(1) are in opening, connection.Children.Count returns 2, so current_children_count = 2
    if current_children_count == 0:
        print("Warning: SAP object did not launch successfully.")
    if current_children_count == 1 and not SAP_in_use_bool and not leave_initial_bool:  # If user is using SAP, don't use the sessions they are using
        target_session = initial_session
    else:
        # Dynamically launch new session
        create_session_success = False
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
        if not create_session_success:
            print("Warning: SAP sessions are busy. Create new session without success.")
    time.sleep(0.1)
   
    return target_session, connection.Children.Count

def concurrent_SAP_session_iterator(SAP_path, login_name, target_func, max_num_session, iterator):
    threads = []
    for params in iterator:
        session, n = launch_SAP_session(SAP_path, login_name, True, False)
        t = threading.Thread(target=target_func, args=(pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch, session), *params))
        threads.append(t)
        t.start()
        if n >= max_num_session:
            for t in threads:
                t.join()
        time.sleep(1)
    for t in threads:
        t.join()

