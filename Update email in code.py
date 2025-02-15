import os
import re

#****************************************************
# Purpose
# Modify email list in all code files under the target directory
#
# Action Value
# 1 - add: add new email next to the target email
# 2 - replace: replace the target email with new email
# (if target email duplicate in a email list, it replaces all instances of the target emails with new email)
# 3 - delete: delete the target email
# (if target email duplicate in a email list, it deletes all instances of the target emails)
# 4 - remove duplicate: remove the duplicate instances of the target email, the same target email will only be seen once in each email list
#
# Action variable must be input a integer value (1, 2, 3, 4)
#****************************************************

#******************************
# Global Variables
#******************************
ADD_ACTION = 1
REPLACE_ACTION = 2
DELETE_ACTION = 3
REMOVE_DUPLICATE_ACTION = 4
file_extensions = ['.R', '.py', '.c', '.java', '.vbs', '.bat', '.ps1']

#******************************
# Functions
#******************************
def add_email_to_files(root_directory, target_email, new_email):
    pattern = re.compile(r'(mail\.(?:to|cc)\s*=\s*[\'"])(.*?)([\'"])', re.IGNORECASE)

    for subdir, _, files in os.walk(root_directory):
        for file in files:
            if any(file.endswith(ext) for ext in file_extensions):
                try:
                    file_path = os.path.join(subdir, file)
                    print('Processing file: ' + file_path)

                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    def replace_emails(match):
                        emails = match.group(2)
                        email_list = [e.strip() for e in emails.split(';')]
                        if target_email in email_list and new_email not in email_list:
                            email_list.insert(email_list.index(target_email) + 1, new_email)
                        return match.group(1) + ';'.join(email_list) + match.group(3)

                    new_content = pattern.sub(replace_emails, content)

                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(new_content)
                except Exception as e:
                    print(e)

def replace_email_in_files(root_directory, target_email, new_email):
    pattern = re.compile(r'(mail\.(?:to|cc)\s*=\s*[\'"])(.*?)([\'"])', re.IGNORECASE)

    for subdir, _, files in os.walk(root_directory):
        for file in files:
            if any(file.endswith(ext) for ext in file_extensions):
                try:
                    file_path = os.path.join(subdir, file)
                    print('Processing file: ' + file_path)

                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    def replace_emails(match):
                        emails = match.group(2)
                        email_list = [e.strip() for e in emails.split(';')]
                        email_list = [new_email if e == target_email else e for e in email_list]
                        return match.group(1) + ';'.join(email_list) + match.group(3)

                    new_content = pattern.sub(replace_emails, content)

                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(new_content)
                except Exception as e:
                    print(e)

def remove_email_from_files(root_directory, target_email):
    pattern = re.compile(r'(mail\.(?:to|cc)\s*=\s*[\'"])(.*?)([\'"])', re.IGNORECASE)

    for subdir, _, files in os.walk(root_directory):
        for file in files:
            if any(file.endswith(ext) for ext in file_extensions):
                try:
                    file_path = os.path.join(subdir, file)
                    print('Processing file: ' + file_path)

                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    def remove_emails(match):
                        emails = match.group(2)
                        email_list = [e.strip() for e in emails.split(';')]
                        email_list = [e for e in email_list if e != target_email]
                        return match.group(1) + ';'.join(email_list) + match.group(3)

                    new_content = pattern.sub(remove_emails, content)

                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(new_content)
                except Exception as e:
                    print(e)

def remove_specific_email_duplicates(root_directory, target_email):
    pattern = re.compile(r'(mail\.(?:to|cc)\s*=\s*[\'"])(.*?)([\'"])', re.IGNORECASE)

    for subdir, _, files in os.walk(root_directory):
        for file in files:
            if any(file.endswith(ext) for ext in file_extensions):
                try:
                    file_path = os.path.join(subdir, file)
                    print('Processing file: ' + file_path)

                    with open(file_path, 'r', encoding='utf-8') as f:
                        content = f.read()

                    def remove_target_duplicates(match):
                        emails = match.group(2)
                        email_list = [e.strip() for e in emails.split(';')]
                        seen = set()
                        new_list = []
                        for email in email_list:
                            if email == target_email:
                                if email not in seen:
                                    new_list.append(email)
                                    seen.add(email)
                            else:
                                new_list.append(email)
                        return match.group(1) + ';'.join(new_list) + match.group(3)

                    new_content = pattern.sub(remove_target_duplicates, content)

                    with open(file_path, 'w', encoding='utf-8') as f:
                        f.write(new_content)
                except Exception as e:
                    print(e)

#******************************
# Main Process
#******************************
root_directory = input("Enter the target directory: ")
target_email = input("Enter the target email: ")
new_email = input("Enter the new email: ")
action = int(input("Please indicate the action. (1 - add, 2 - replace, 3 - delete, 4 - remove duplicate, other input - do nothing): "))
if action == ADD_ACTION:
   add_email_to_files(root_directory, target_email, new_email)
elif action == REPLACE_ACTION:
    replace_email_in_files(root_directory, target_email, new_email)
elif action == DELETE_ACTION:
    confirm = input('Confirm to delete the email ' + target_email + '? (Y - yes, other input - no)')
    if confirm == 'Y':
        remove_email_from_files(root_directory, target_email)
    else:
        print("Do nothing.")
elif action == REMOVE_DUPLICATE_ACTION:
    remove_specific_email_duplicates(root_directory, target_email)
else:
    print("Do nothing.")