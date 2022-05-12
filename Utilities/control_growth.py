import os
import glob

def control_growth_of_docx():
    """
    This function controls the number of docx files in the DOWNLOAD_FOLDER
    """
    # Get a list of all the file paths that ends with .docx from in specified directory
    fileList = glob.glob('./DOWNLOAD_FOLDER/*.docx')
    growth = len(fileList)
    limit = 5

    if growth > limit:
        leave_behind = 0
        # Iterate over the list of filepaths & remove each file.
        for filePath in fileList:
            try:
                if leave_behind <= limit - 2:
                    os.remove(filePath)
                    leave_behind += 1
            except:
                print("Error while deleting file : ", filePath)

def control_growth_of_xlsx(location):
    """
    This function controls the number of docx files in the DOWNLOAD_FOLDER
    """
    # Get a list of all the file paths that ends with .docx from in specified directory
    fileList = glob.glob(location)
    growth = len(fileList)
    limit = 3

    if growth > limit:
        leave_behind = 0
        # Iterate over the list of filepaths & remove each file.
        for filePath in fileList:
            try:
                if leave_behind <= limit - 2:
                    os.remove(filePath)
                    leave_behind += 1
            except:
                print("Error while deleting file : ", filePath)

