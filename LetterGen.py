import os.path
from excel2json import convert_from_file
import json
import docx
from time import sleep
import Helpers

word_template_name = ""
excel_doc_name = ""
enableVerbose = False # Enable/disable verbose console output
	
# Check if all entered file names can be found
def AllFilesExist():
    if os.path.isfile(word_template_name) and os.path.isfile(excel_doc_name):
        return True
    else:
        return False

def AskVerbose():
    response = input ("Would you like to enable verbose console output (y/n): ").strip().lower()
    if response == 'y':
        return True
    elif response == 'n':
        return False
    else:
        print("Illegal entry. We will go with No to keep things clean")
        return False

# Parse Excel sheet to JSON
def ParseToJson():
    print("Parsing Excel sheet contents and generating Word docs from template")

    if enableVerbose == True: Helpers.VerbosePrint("Attempting to convert Word doc to JSON...")

    try:
        convert_from_file(excel_doc_name)
    except Exception as err:
        Helpers.PrintError(f"Error parsing content: {err}.\nExiting program...")
        exit(0)

    if enableVerbose == True: Helpers.VerbosePrint("Opening Sheet1.json (default Excel workbook name)")

    input_file = open('Sheet1.json') # Use first (and only) json file to load data from
    contactDict = json.load(input_file)
    GenerateDocuments(contactDict)

# Load Word doc template and apply changes from contacts dictionary
def GenerateDocuments(contactDict):
    if enableVerbose == True: Helpers.VerbosePrint("Beginning Word doc generation...")

    # { "find this text": "replace with this text" }
    ReplacementDictionary = {"Family Name Here": "FamilyName", "Address Line 1": "AddressLine1", "Address Line 2": "AddressLine2"}

    familyIndex = 0

    for contact in contactDict:
        doc = docx.Document(word_template_name)
        run = doc.add_paragraph().add_run() # Allow a run to be used on modified text
        style = doc.styles['Normal'] # Normal styling applied
        font = style.font # Accessor to font face
        font.bold = True
        font.name = 'Cavolini' # Set font face
        font.size = docx.shared.Pt(11) # Set font size

        if enableVerbose == True: Helpers.VerbosePrint("Formatting " + contactDict[familyIndex]["FamilyName"])
        
        for i in ReplacementDictionary:
            for p in doc.paragraphs:
                if p.text.find(i) >= 0:
                    p.text = p.text.replace(i, contact[ReplacementDictionary[i]])
                    p.add_run()

        familyIndex += 1
        fileName = Helpers.RemoveAllSpaces(str(contact["FamilyName"]))

        if enableVerbose == True:
            Helpers.VerbosePrint(f"Attempting to save {fileName}.docx")

        try:
            doc.save(f"generated_docs/{fileName}.docx")
        except IOError as ex:
            Helpers.PrintError(f"Error saving generated document for {fileName}: {ex}")

    print("Process completed. Check generated_docs folder for final outputs.")

if __name__ == "__main__":
    Helpers.ClearScreen()
    word_template_name = Helpers.GetUserInput("Please enter name of Word doc template: ", False)
    excel_doc_name = Helpers.GetUserInput("Please enter name of Excel doc to rip data from: ", False)

    if AllFilesExist() == False:
        Helpers.PrintError("Could not find excel doc with matching name. Please run the script and try again...")
        sleep(5)
        exit(0)

    enableVerbose = AskVerbose() 
    sleep(1)
    ParseToJson()

