import os

#****************************************************
# Action Value
# 1 - search: search any files with desired file extensions containing the search string in the specific directory
# 2 - replace: replace all files with desired file extensions containing the search string with replacement string in the specific directory
#
# Action variable must input integer value (1, 2)
#****************************************************

###############################
# Global Variables
###############################
login_name = os.getlogin()
current_directory = os.path.dirname(os.path.realpath(__file__))

directory_path = "C:\\Users\\" + login_name + "\\"
search_string = 'Test'
replacement_string = 'Test'
file_extensions = ['.py', '.R', '.cpp', '.java', '.vbs', '.ps']  # Specify desired file extensions
action = 1

###############################
# Functions
###############################
def find_files_with_string(directory, search_string, file_extensions):
    matching_files = []
    
    for root, _, files in os.walk(directory):
        for file in files:
            if any(file.endswith(ext) for ext in file_extensions):
                file_path = os.path.join(root, file)
                try:
                    with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                        if search_string in f.read():
                            matching_files.append(file_path)
                except Exception as e:
                    print("Error reading " + file_path + ": " + str(e))
    
    return matching_files

def replace_strings_in_files(directory, search_string, replacement_string, file_extensions):
    files = find_files_with_string(directory, search_string, file_extensions)
    replaced_files = []
    
    for file in files:
        try:
            with open(file, 'r', encoding='utf-8', errors='ignore') as f:
                content = f.read()
            new_content = content.replace(search_string, replacement_string)
            with open(file, 'w', encoding='utf-8') as f:
                f.write(new_content)
            replaced_files.append(file)
        except Exception as e:
            print("Error replacing string in " + file + ": " + str(e))
    
    return replaced_files

###############################
# Main Process
###############################
if action == 1:
    files = find_files_with_string(directory_path, search_string, file_extensions)
    if len(files) > 0:
        print("List of files that contain string with " + search_string + ":")
        for file in files:
            print(file)
    else:
        print("No files that contain string with " + search_string + " is found.")
elif action == 2:
    files = replace_strings_in_files(directory_path, search_string, replacement_string, file_extensions)
    if len(files) > 0:
        print("String " + search_string + " replaced in files with " + replacement_string + ":")
        for file in files:
            print(file)
    else:
        print("No files that contain string with " + search_string + " is found.")
else:
    print("Do nothing.")