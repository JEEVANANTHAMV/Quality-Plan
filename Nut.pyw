import shutil
import os
path=r"C:\Solution"
if os.path.isdir(path)==False:
    os.makedirs(path,exist_ok=True)

source = "Nut\Temp1.xlsx"
destination = r"C:\Solution"

shutil.copy(source, destination)
source = "Nut\Temp2.xlsx"
shutil.copy(source, destination)

source = "Nut\Temp3.xlsx"
shutil.copy(source, destination)
