#Python library imports
import os
import glob

#Import correction_stud.py
import correction_wb

#Set path
path = 'C:\\Users\\senns\\Documents\\Uni_Stuff\\2022\\Bachelorarbeit\\Final_Take\\Testing_Momentum_Strategies_using_Python_BT\\files_for_correction'
stud_number = 0

#Set up directory
for dir in os.listdir(path): 
    filenames = glob.glob(os.path.join(path + '\\' + dir + '\\2_submissions\\' , "*.xlsx"))
    stud_number +=1
    if filenames != []:
        correction_wb.correction(filenames[0])
