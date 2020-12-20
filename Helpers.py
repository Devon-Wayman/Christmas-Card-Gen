from os import system, name 
from datetime import datetime
import os

class bcolors:
    HEADER = '\033[95m'
    OKBLUE = '\033[94m'
    OKCYAN = '\033[96m'
    OKGREEN = '\033[92m'
    WARNING = '\033[93m'
    FAIL = '\033[91m'
    ENDC = '\033[0m'
    BOLD = '\033[1m'
    UNDERLINE = '\033[4m'

def ClearScreen(): 
    if name == 'nt': 
        _ = system('cls') 
    else: 
        _ = system('clear') 

def VerbosePrint(message):
    dt_string = datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S.%f')[:-3]
    print (f"[{dt_string}: {message}")

def PrintError(message):
    print(f"{bcolors.WARNING}{message}{bcolors.ENDC}")

def RemoveAllSpaces(string):
    return string.replace(" ", "")


# Return string value from user input
def GetUserInput(requestMessage, printAsError):
    if printAsError:
        userIn = input(f"{bcolors.WARNING}{requestMessage}{bcolors.ENDC}")
    else:
        userIn = input(requestMessage)
    return userIn.strip()
    
# Ask user to set a bool value to true or false
def SetBool(message):
    userInput = input(message)
    userInput = userInput.strip().lower()
    if userInput == 'y':
        return True
    elif userInput == 'n':
        return False
    else:
        print("Illegal entry. We will go with No as a default")
        return False

# Ask a user to set an int value
def SetInt(message):
    userInput = input(message)
    userInput = userInput.strip().lower()
    
    if userInput == "":
        return 11

    return int(userInput)

def CheckForGenDocsFolder():
    if os.path.isdir('generated_docs'):
        return
    else:
        PrintError("Main folder for generated docs not found. Creating now...")
        os.mkdir('generated_docs')
        return False