import msvcrt
import time
import requests
import os
import subprocess

### fontend menu ###
def Frontend_menu():
    print()
    print()
    print('       __________________________________________________________________________________________________________________       ')

    print("       Welcome to 'TH for distribution' Accounting Assistant")
    print('       made by Alfred Tharwat')
    print('       __________________________________________________________________________________________________________________       ')
    print('       [0] check for updates')
    print('       [1] Sales Assistant')
    print('       [2] Storage Assistant')
    print()
    print()
    print()
    print('       __________________________________________________________________________________________________________________       ')
    print('       Enter a menu option in the Keyboard [0,1,2,3,4] :')
    print()
    print()


### check for updates

# Define the GitHub repository URL and file path in the repository
github_repo_url = "https://github.com/asyoty/TH_for_Distribution_Accounting_assistant"
file_path_in_repo = "dist/TH for Distribution Accounting assistant.py"
# Function to download and replace the program
def update_program():
    try:
        # Download the file from the GitHub repository
        response = requests.get(f"{github_repo_url}/raw/main/{file_path_in_repo}")
        if response.status_code == 200:
            # Save the downloaded file with the same name as the current script
            with open(os.path.basename(__file__), "wb") as file:
                file.write(response.content)
            print("Program updated successfully. Please restart the program.")
        else:
            print("Update failed. Please contact admin.")
    except Exception as e:
        print(f"Update failed: {str(e)}")



### calling other programs ###
Frontend_menu()

while True:
    # calling the updater
    if msvcrt.getch().decode('utf-8') ==  '0':
        print('checking for updates.....')
        update_program()
        break
    # calling the Acconting Assistant
    elif msvcrt.getch().decode('utf-8') ==  '1':
        print('Initializing Sales Asisstant.....')
        import Sales_Assistant
        print('Task completed successfully ')
        break
    # calling the Storage Assistant
    elif msvcrt.getch().decode('utf-8') ==  '2':
        print('Initializing Storage Asisstant.....')
        import Storage_Assistant
        print('Task completed successfully ')
        break

while True:
        user_input = input("Enter 'exit' to quit: ")
        if user_input == 'exit':
            print("Exiting the program.")
            time.sleep(3)
            break