from fileinput import filename
import glob
import os
from typing import Tuple
from datetime import datetime
from openpyxl import load_workbook
# import Tkinter, tkFileDialog

def extract_all_xlsx_files_in_directory(path) -> list[str]:
    xlsx_found = []
    for file in glob.glob(path + "**/*.xlsx", recursive=True):
        xlsx_found.append(file)
    return xlsx_found


def filter_for_reports(list_of_files: list) -> list[Tuple[str, str]]:
    """Filters a list of file addresses for FCR reports, then
        returns those reports only.

    Args:
        list_of_files (list): A list of file addresses of .xlsx files

    Returns:
        list: A list of FCR reports
    """
    to_return = []
    for file in list_of_files:
        filename_actual = file.lower().split("\\")[-1]
        presumed_dates = [int(s) for s in filename_actual.split() if s.isdigit() and len(s) == 4]
        
        if len(presumed_dates) >= 1:
            date = str(max(presumed_dates))
        else:
            date = "N/A"
        
        keyword_set = {'FCR', 'financial', 'financials', 'financial report', 'financials report', 'fcr'}
        if any(keyword in filename_actual for keyword in keyword_set) and "final" in filename_actual:
            to_return.append((date, file))
    
    return to_return

def filter_for_risk_reports(list_of_files: list) -> list[Tuple[str, str]]:
    to_return = []
    for file in list_of_files:
        filename_actual = file.lower().split("\\")[-1]
        presumed_dates = [int(s) for s in filename_actual.split() if s.isdigit() and len(s) == 4]
        
        if len(presumed_dates) >= 1:
            date = str(max(presumed_dates))
        else:
            date = "N/A"
        
        keyword_set = {'scorecard', 'risk rating', 'borrower risk rating scorecard'}
        if any(keyword in filename_actual for keyword in keyword_set) and "final" in filename_actual:
            to_return.append((date, file))

def return_newest_file(list_of_files: list) -> str:
    """Returns the newest file in a list of files. Heuristic, not final.
    """
    latest_file = max(list_of_files, key=os.path.getmtime)
    return latest_file
    
#def main():
    # root = Tkinter.Tk()
    # dirname = tkFileDialog.askdirectory(parent=root, initialdir="/", title='Please select a directory')
    
    # TODO: Add a check to make sure the user selected a directory
    
    # TODO: Implement navigation to directory and file selection process (manual is better?)
