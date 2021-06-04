# Renaming the file names in the folder code  - Tested and working

import os

path = os.chdir("C:\\Users\\VAS\Desktop\\Test\\Name_change_experiment")

i = 1

for file in os.listdir(path):

    new_file_name = "Jan_2020_{}.xlsx".format(i)
    os.rename(file, new_file_name)


    i = i + 1