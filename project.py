
# Import the necessary modules
import arrow
import os
import pandas as pd
import PySimpleGUI as sg

# -------------------------------------------------------------------------------------------------
# H E L P E R   F U N C T I O N S
# -------------------------------------------------------------------------------------------------

def pull_file_info(folder_path: str) -> list:
    '''Pull file information from file in the specified path, return a list of dictionaries of attributes'''
    
    attribute_list = []

    # Loop through all directories (folders)
    for root, _, files in os.walk(folder_path):
        for file in files:
            file_path = os.path.join(root, file)
            if os.path.isfile(file_path): # If the item is a file, not another folder
                file_info = {
                    "File Path": str(file_path),
                    "File Name": str(os.path.splitext(os.path.basename(file_path))[0]),
                    "File Type": str(os.path.splitext(file_path)[1]),
                    "File Size (KB)": str(os.path.getsize(file_path) / (10**3))
                }
                attribute_list.append(file_info)

    return attribute_list


def write_to_excel(out_folder: str, info: list):
    '''Given the file information as a list of dictionaries, write to an excel file in the out folder'''

    global data
    data = pd.DataFrame(info) # dataframe of the file info passed in

    # Create columns for hyperlink and date added
    data['Link'] = data['File Path'].apply(lambda x: f'=HYPERLINK("{x}", "Link")')
    data['Date Added'] = arrow.now().format("MM/DD/YYYY")
    data = data[["File Path", "Link", "File Name", "File Type", "File Size (KB)", 'Date Added']] # Re-ordering
    
    # Write to excel
    date = arrow.now().format("YYYY.MM.DD-HH.mm")
    output_file = os.path.join(out_folder, f"####.#### - Document Summary_{date}.xlsx")
    data.to_excel(output_file, index=True)


def find_new_docs(folder_path: str, excel_path: str) -> pd.DataFrame:
    '''Add documents to an existing excel list and the date the doc was added'''
    
    global to_add
    try:
        existing_df = pd.read_excel(excel_path)
    except FileNotFoundError:
        sg.popup_error("Excel File Not Found")
        return pd.DataFrame
    
    existing_files = set(existing_df['File Name'].tolist()) # List of existing file names

    all_files = pd.DataFrame(pull_file_info(folder_path)) # dataframe of all files in the folder
    
    # Use a boolean mask to filter for files not in the existing list 
    mask = ~all_files['File Name'].isin(existing_files)
    to_add = all_files[mask].copy()
    
    if not to_add.empty:
        to_add['Date Added'] = arrow.now().format("MM/DD/YYYY")

    return to_add


def write_new_docs(folder_path: str, to_add_df: pd.DataFrame):
    '''Write a new excel file with only the documents not in excel_path'''

    if to_add_df.empty:
        print('Empty dataframe passed to write new docs')
        return

    # Create a column for the hyperlink and the date added
    to_add_df['Link'] = to_add_df['File Path'].apply(lambda x: f'=HYPERLINK("{x}", "Link")')
    to_add_df['Date Added'] = arrow.now().format("MM/DD/YYYY")
    to_add_df = to_add_df[["File Path", "Link", "File Name", "File Type", "File Size (KB)", "Date Added"]] # Re-ordering

    # Output
    date = arrow.now().format("YYYY.MM.DD-HH.mm")
    output_file = os.path.join(folder_path, f"Added Docs_{date}.xlsx")
    to_add_df.to_excel(output_file, index=True)


# -------------------------------------------------------------------------------------------------
# M A I N   F U N C T I O N
# -------------------------------------------------------------------------------------------------

def main():
    sg.theme('LightGrey1') # set the theme of the window

    # Layout the GUI
    layout = [
        [sg.Text("Select Folder:"), sg.InputText(), sg.FolderBrowse()],
        [sg.Checkbox("Initialization", default=False, key="initial")],
        [sg.Checkbox("Rerun", default=False, key="rerun")],
        [sg.Text("If rerunning: Select Excel Sheet:"), sg.InputText(), sg.FileBrowse()],
        [sg.Checkbox("Show Summary", default=False, key="showSummary")],
        [sg.Button("Run"), sg.Button("Exit")]
    ]

    # Create the window and continuously listen for events
    window = sg.Window("DocRev", layout)
    while True:
        event, values = window.read()

        # End the loop if the window is closed for the user presses exit
        if event == sg.WIN_CLOSED or event == "Exit":
            break

        # When the user clicks run
        if event == "Run":

            # Values from the user's input to the GUI
            folder = values[0] 
            excel_sheet = values[1]          
            process_initial = values["initial"]
            process_rerun = values["rerun"]
            show_summary = values["showSummary"]

            if process_initial:
                info = pull_file_info(folder)
                write_to_excel(folder, info)
                sg.popup("Initialization Run Complete")

            if process_rerun: 
                to_add = find_new_docs(folder, excel_sheet)
                if to_add.empty:
                    sg.popup("Rerun Complete, no new docs found")
                else:
                    write_new_docs(folder, to_add)
                    sg.popup("Rerun Complete")

            if show_summary: 
                sg.popup(f"Summary:\nOriginal Docs: {data.shape[0]}\nAdded Docs: {to_add.shape[0]}")            

    window.close()

if __name__ == "__main__":
    main()

# Made with the generous and eternally patient helpers Copilot, Claude, and Gemini