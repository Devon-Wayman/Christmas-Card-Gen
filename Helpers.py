from os import system, name 
from datetime import datetime

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

# Used to return user input. Can also be passed in a message to print before entry region renders
def GetUserInput(requestMessage, printAsError):
    if printAsError:
        userIn = input(f"{bcolors.WARNING}{requestMessage}{bcolors.ENDC}")
    else:
        userIn = input(requestMessage)
    return userIn.strip()