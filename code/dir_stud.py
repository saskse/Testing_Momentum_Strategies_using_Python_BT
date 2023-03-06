# Python library imports
import os
import glob

# Import correction_stud.py
import correction_stud

#set path
path = 'C:\\Users\\senns\\Documents\\Uni_Stuff\\2022\\Bachelorarbeit\\Final_Take\\Code_Stud\\Testing_Momentum_Strategies_using_Python_BT\\data\\input\\files_for_correction_stud\\ita_IA_3_2022-12-23T20-12-29_674'
stud_number = 0

#set up directory
for dir in os.listdir(path): 
    filenames = glob.glob(os.path.join(path + '\\' + dir + '\\2_submissions\\' , "*.xlsx"))
    stud_number +=1
    if filenames != []:
        correction_stud.correction(filenames[0])