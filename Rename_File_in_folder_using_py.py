>import os

>>> path = os.chdir("C:\\Users\\VAS\\Desktop\\Test\\Name_change_experiment")
>>> i = 0
>>> for file in os.listdir(path):

    new_file_name = "2020_01().xlsx".format(i)
    os.rename(file, new_file_name)

    i = i + 1
