import os

###############################
# Purpose
# To search any file names that contains searching_string with specific file extensions
#
# You can click the path in the Visual Studio Code terminal to access the path directly
# Tips: Holding "Ctrl" and left click the path
###############################

###############################
# Global Variables
###############################
login_name = os.getlogin()
current_directory = os.path.dirname(os.path.realpath(__file__))

directory_path = "C:\\Users\\"
search_string = 'Test'
file_extensions = ['.csv', 'xlsm', 'xlsx']

###############################
# Functions
###############################
def find_file_name_with_string(directory, search_string, file_extensions):
    matching_files = []
    for root, _, files in os.walk(directory):
        for file in files:
            if any(file.endswith(ext) for ext in file_extensions) and search_string in file:
                file_path = os.path.join(root, file)
                matching_files.append(file_path)
    
    return matching_files

###############################
# Main Process
###############################
files = find_file_name_with_string(directory_path, search_string, file_extensions)
if len(files) > 0:
    print("List of files with filenames containing " + search_string + ":")
    for file in files:
        print(file)
else:
    print("No files with filenames containing " + search_string + " found.")