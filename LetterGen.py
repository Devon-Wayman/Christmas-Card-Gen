import os.path
from excel2json import convert_from_file
import json
import docx
from time import sleep
import Helpers

word_template_name = ""
excel_doc_name = ""
fontName = ""

enableVerbose = False # Enable/disable verbose console output
useBoldFont = False # Set whether or not to apply bold factor to new text

fontSize = 11 # Font size to set

# Check if all entered file names can be found
def AllFilesExist():
    if os.path.isfile(word_template_name) and os.path.isfile(excel_doc_name):
        return True
    else:
        return False

# Parse Excel sheet to JSON file
def ParseToJson():
    print("Parsing Excel sheet contents and generating Word docs from template")

    if enableVerbose: Helpers.VerbosePrint("Attempting to convert Word doc to JSON...")

    try:
        convert_from_file(excel_doc_name)
    except Exception as err:
        Helpers.PrintError(f"Error parsing content: {err}.\nExiting program...")
        exit(0)

    if enableVerbose: Helpers.VerbosePrint("Opening Sheet1.json (default Excel workbook name)")

    input_file = open('Sheet1.json') # Use first (and only) json file to load data from

    if enableVerbose: Helpers.VerbosePrint("Loading JSON file into dictionary")
    contactDict = json.load(input_file) # Load the JSON data into a formatted dictionary
    GenerateDocuments(contactDict) # Begin generating new Word docs using the passed in data

# Load Word doc template and apply changes from contacts dictionary
def GenerateDocuments(contactDict):
    if enableVerbose: Helpers.VerbosePrint("Beginning Word doc generation...")

    # { "find this text": "replace with this text" }
    ReplacementDictionary = {"Family Name Here": "FamilyName", "Address Line 1": "AddressLine1", "Address Line 2": "AddressLine2"}

    familyIndex = 0 # Index number to grab cooresponding family's data from the contacts dictionary when replacing text

    for contact in contactDict:
        doc = docx.Document(word_template_name) # Open the UNMODIFIED Word doc template
        run = doc.add_paragraph().add_run() # Allow a run to be used on modified text
        style = doc.styles['Normal'] # Normal styling applied
        font = style.font # Accessor to font face
        font.bold = useBoldFont # Generate bold font if useBoldFont is set to true

        if not fontName == "":
            font.name = fontName # Set font face to custom by user
        else:
            font.name = 'Cavolini' # Use this default if none was given

        font.size = docx.shared.Pt(fontSize) # Set font size to one passed in by user

        if enableVerbose: Helpers.VerbosePrint("Formatting " + contactDict[familyIndex]["FamilyName"])
        
        for i in ReplacementDictionary:
            for p in doc.paragraphs:
                if p.text.find(i) >= 0:
                    p.text = p.text.replace(i, contact[ReplacementDictionary[i]]) # replace detected text accordingly
                    p.add_run() # run text formatting on the newly replaced document text

        familyIndex += 1
        fileName = Helpers.RemoveAllSpaces(str(contact["FamilyName"]))

        if enableVerbose: Helpers.VerbosePrint(f"Attempting to save {fileName}.docx")

        try:
            doc.save(f"generated_docs/{fileName}.docx")
        except IOError as ex:
            Helpers.PrintError(f"Error saving {fileName}.docx: {ex}. Please ensure an older copy of the file is not currently open!")

    print("\nProcess completed. Check generated_docs folder for final outputs.\n")

if __name__ == "__main__":
    Helpers.ClearScreen()
    word_template_name = Helpers.GetUserInput("Please enter name of Word doc template: ", False)
    excel_doc_name = Helpers.GetUserInput("Please enter name of Excel doc to rip data from: ", False)

    if not AllFilesExist():
        Helpers.PrintError("Could not find excel doc with matching name. Please re-run the script and try again...")
        sleep(5)
        exit(0)

    enableVerbose = Helpers.SetBool("Would you like to enable verbose console output (y/n)? ")
    useBoldFont = Helpers.SetBool("Would you like the replaced font to be bold (y/n)? ") 
    fontName = Helpers.GetUserInput("Please enter font family name to use (leave blank for default; Cavolini): ", False)
    fontSize = Helpers.SetInt("Please enter a font size to use (leave blank for 11): ")

    sleep(1)
    ParseToJson()

