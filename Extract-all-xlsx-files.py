from fileinput import filename
import glob
import os

def extract_all_xlsx_files_in_directory(path) -> list:
    xlsx_found = []
    for file in glob.glob(path + "**/*.xlsx", recursive=True):
        xlsx_found.append(file)
    return xlsx_found

def filter_for_reports(list_of_files: list) -> list:
    """Filters a list of file addresses for FCR reports, then
        returns those reports only.

    Args:
        list_of_files (list): A list of file addresses of .xlsx files

    Returns:
        list: A list of FCR reports
    """
    to_return = []
    for file in list_of_files:
        filename_actual = file.split("\\")[-1].lower()
        keyword_set = {'report', 'financial', 'financials', 'financial report', 'financials report', 'credit'}
        if any(keyword in filename_actual for keyword in keyword_set):
            to_return.append(file)
    
    return to_return

def return_newest_file(list_of_files: list) -> str:
    """Returns the newest file in a list of files. Heuristic, not final.
    """
    latest_file = max(list_of_files, key=os.path.getmtime)
    return latest_file
    
    