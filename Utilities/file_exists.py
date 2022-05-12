import os
ALLOWED_EXTENSIONS = {'xlsx'}

def file_exists(file_path):
    """
    This function checks if a file exists
    """
    return os.path.exists(file_path)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
