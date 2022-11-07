import shutil
import os
path=r"C:\Solution"
if os.path.isdir(path)==False:
    os.makedirs(path,exist_ok=True)

source = "Suggestion.xlsx"
destination = r"C:\Solution"

shutil.copy(source, destination)
source = "TOLERANCE.xlsx"
shutil.copy(source, destination)

source = "Print_Template.xlsx"
shutil.copy(source, destination)

