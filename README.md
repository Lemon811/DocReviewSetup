# DOCREV
### Chris Lemley | CS50P 2025 via EdX | Submitted 2/2/2025

# PURPOSE
In my line of work, we often have to review a large amount of documents to gather information
for our reports. Typically, a large folder of files is provided and the file are unorganized
and scattered. Over the course of the project, multiple people will be reviewing the documents.
Additional documents may be added to the folder as they are found or created.

The purpose of my applicaiton DocRev is to create an excel summary table of all the documents in the folder that includes the following information:
- Path to the file : the path to the location of the file on your computer
- Link to the file location : a link directly to the file
- File name : the name of the file, without its extension
- File type : the file extension
- File size (KB) : the size of the file in KB
- Date added : the date the file was noticed by the applicaiton, new files will have a date
differing from those added during the initialization

If new documents have been added to the folder, the user can select rerun and a list of the
documents added to the folder will be written to excel. This will avoid confusion when 
multiple people are adding documents in nested folders. The resulting excel sheet of a rerun
will have the same information listed above

# REQUIREMENTS
DocRev uses the open source Python libraries listed in the file "requirement.txt" 
Before running DocRev, ensure all dependencies are installed using thefollowing command: 
`pip install -r requirements.txt`

# USE
When the application is run, a main GUI box will appear with two file/folder selectors and three
checkboxes. 

## Initialization
For an initializaiton (i.e. the first time a document summary is generated) select the
folder where all the documents are located, check the "intialization" check box, and hit "Run".
Do not select the check box for "rerun" during a initialization. An excel file with all the file
information should appear in the the selected folder.

## Rerun
A rerun is used when additional documents have been added to the folder, or you want to determine
if additional documents have been added. Select the folder where all the documents are located,
check the "rerun" check box, select the document review excel file already in the folder, and 
hit "run". Do not select the check box for "initialization" during a rerun. An excel file with
all the file information of the new documents should appear in selected folder.

## Show Summary
Show summary is an optional check box to display the number of files already tabulated in the 
original excel sheet and the number of new files added to the folder.

# FUNCTIONS
The helper and fuctions used in the DocRev application are as follows:

### pull_file_info
Given a folder path, return a list of dictionaries with the file properties of each file.

### write_to_excel
Given a path to where the excel document should export and a list of dictionaries, write an
excel file with all the file information (no return value).

### find_new_docs
Given a folder path and an excel document with a list of files, return a dataframe of all
the documents in the folder that are not listed in the excel file. The return value is a
DataFrame with all the file information gathered using the function `pull_file_info`

### write_new_docs
Given a path to where the excel document should export and a DataFrame of new documents, write an
excel file with all the file information (no return value). If the passed DataFrame is empty, return nothing.

### main
Constructs the GUI window with all the user input fields. Main then runs continuosly until the
user exits the program.
