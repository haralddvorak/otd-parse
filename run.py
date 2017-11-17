from OfficeTime import OfficeTimeFile
import os

current_path = os.path.curdir

# scan through current directory for every .otd-file available
with os.scandir(current_path) as it:
    for entry in it:
        if entry.name.endswith('.otd') and entry.is_file():
            # parse the file
            my_file = OfficeTimeFile(entry.name)
            # export the file to Excel readable csv
            my_file.csv_export()

# Alternative run on one specific file:
# myfile = OfficeTimeFile("OfficeTimeTest.otd")
# myfile.csv_export()
