import pandas as pd
import pywinauto                                        # bringing an .exe to the foreground
import time                                             # call time
import win32con
import win32gui                                         # bring apps to front foreground
from pathlib import PureWindowsPath                     # library that cleans up windows path extensions
import tkinter as tk                                    # Tkinter's Tk class
from tkinter import messagebox                          # standard message box  

main = tk.Tk()
main.geometry('100x100+5+5')                                 # Set the geometry of the GUI (LxH+posX+posY)
main.title("Main Window")

samp_arr_raw1 = []
samp_arr_raw2 = []
samp_arr_raw3 = []
samp_arr_raw = []

for a in range(1,25):
    samp_arr_raw1.append(a)
    samp_arr_raw.append(a)
for b in range(26,50):
    samp_arr_raw2.append(b)
    samp_arr_raw.append(b)
for c in range(51,75):
    samp_arr_raw3.append(c)
    samp_arr_raw.append(c)

def console():                                                                  # place python console in foreground and 1/4 screen
    #py_title = r'C:\Users\oqpathscope\AppData\Local\Programs\Python\Python311\python.exe'   # the title of the console
    py_title = r'C:\Program Files\Python311\python.exe'                         # the title of the console
    py_app = pywinauto.Application().connect(title=py_title)                    # connect to the app with the title
    py_win = py_app[py_title]                                                   # assign the app to a variable
    py_win.set_focus()                                                          # set the focus to the excel file to the foreground
    print('Python is in the foreground now.')
    time.sleep(1)                                                               # pause to allow to come to foreground

    full_py = win32gui.GetForegroundWindow()                                    # grab the window in the foreground
    py_rect = win32gui.GetWindowRect(full_py)                                   # assign window rectangle coordinates to an array
    a = py_rect[0]                                                              # a=upper left corner positon of the Kinesis window in the X coordinates of the screen
    b = py_rect[1]                                                              # b=upper left corner positon of the Kinesis window in the Y coordinates of the screen
    c = py_rect[2] - a                                                          # c is the length of the kinesis window, should be half the length of the screen 1920/2=960
    d = py_rect[3] - b                                                          # d is the height of the kinesis window, should be the entire height of the screen 1080
    if b != 0:                                                                  # if b is not equal to 0 (Y in the 0 location)
        win32gui.SetWindowPos(full_py, win32con.HWND_TOP, 0, 0, 960, 1000, 0)  # set to upper left hand corner (0,600) 1/4 screen (960x600)
        print('Python Console is now 1/4 screen.')                              # X,Y,L,H. X&Y are top left corner position. L&W of the GUI window
        time.sleep(1)
    else:                                                                   
        print('Python is already 1/4 screen.')

console()
#path_xl= r'C:\Users\Public\Documents\Lumedica\OctEngine\Data\13DTEST18UM-01-A 13D/'
#file_xl= r'data_20230714_122858.xlsx'

path_xl= r'\\RXS-FS-02\userdocs\shaynes\My Documents\R&D - Software\Python\Excel_File_2023-07-23/'
file_xl= r'data_20210502_054539.xlsx'

loc_xl = path_xl + "/" + file_xl
win_loc_xl = PureWindowsPath(loc_xl)

df_xl = pd.read_excel(win_loc_xl, usecols="A:I", index_col=None, engine='openpyxl', keep_default_na=True, na_values='np.nan')   # keep_default_na=True returns nan, keep_default_na=False returns N/A
print("Dataframe df_xl:\n", df_xl)   

data_length = len(df_xl)
print("Data Length is: ", data_length)

ending_row = int(data_length)+1                                             # get last row of data
print("Therefore, the last row with data (in Excel) is: ", ending_row, " (", ending_row, "+1)")                                              

col_file_name = df_xl.iloc[0:ending_row, 3]                                 # reads all mold numbers to ending row, in column 3=D
print("col_file_name: \n", col_file_name)
mold_data_temp = col_file_name.str.split()
print("mold_data_temp: \n", mold_data_temp)

df_mold_text = []
df_mold_num  = []
df_meas_pos  = []
for i in range(0,data_length):                                              # read rows 0:sample_length in column 3=D, with Mold, 1, 015 into different variable lists
    df_mold_text.append(mold_data_temp[i][0])
    df_mold_num.append(int(mold_data_temp[i][1]))                          # datatype list of integers
    df_meas_pos.append(int(mold_data_temp[i][2]))
print("Column Mold Text is: \n", df_mold_text)
print("Column Sample Numbers are: \n", df_mold_num)
print("Column Measurement Positions are: \n", df_meas_pos)

samp_num_015 = []
samp_num_040 = []
for j in range(0,data_length):
    if df_meas_pos[j] == 15:
        samp_num_015.append(df_mold_num[j])
    elif df_meas_pos[j] == 40:          
        samp_num_040.append(df_mold_num[j])
    else:
        print("else, pass")
        pass
print("samp_num_015: ", samp_num_015)
print("samp_num_040: ", samp_num_040)

count_015 = len(samp_num_015)
print("Number of 015 samples tested: ", count_015)
count_040 = len(samp_num_040)
print("Number of 040 samples tested: ", count_040)

duplic_015 = samp_num_015.iloc[0:count_015]
duplic_040 =samp_num_040.iloc[0:count_040]
#duplic_015.duplicated(keep='last')                          # Mark 'True' for first duplicate. We will delete that row.
#duplic_040.duplicated(keep='last')                          # Mark 'True' for first duplicate. We will delete that row.
duplic_015.duplicated()                                     # Mark 'True' for first duplicate. We will delete that row.
duplic_040.duplicated()                                     # Mark 'True' for first duplicate. We will delete that row.
print("duplic_015: \n", duplic_015)
print("duplic_040: \n", duplic_040)

col_10A = df_xl.iloc[0:count_015, 4]                                      # read rows 0:samp len 015, in column 4=E
col_10B = df_xl.iloc[0:count_015, 5]                                      # read rows 0:samp len 015, in column 5=F
col_45A = df_xl.iloc[count_015:ending_row, 6]                             # read rows samp len 015:total samp len, in column 6=G
col_45B = df_xl.iloc[count_015:ending_row, 7]                             # read rows samp len 015:total samp len, in column 7=H
col_45C = df_xl.iloc[count_015:ending_row, 8]                             # read rows samp len 015:total samp len, in column 8=I

print("Col_10A is: \n", col_10A)
print("Col_10B is: \n", col_10B)
print("Col_45A is: \n", col_45A)
print("Col_45B is: \n", col_45B)
print("Col_45C is: \n", col_45C)

temp015_1 = []
temp015_2 = []
temp040_1 = []
temp040_2 = []
temp040_3 = []
for k in range(0,count_015):                           # do we want the blank rows included in the count? (full datalength?)
    temp015_1.append(col_10A[k])
    temp015_2.append(col_10B[k])
print("temp015_1: \n", temp015_1)
print("temp015_2: \n", temp015_2)

for m in range(count_015,data_length):       
    temp040_1.append(col_45A[m])
    temp040_2.append(col_45B[m])
    temp040_3.append(col_45C[m])
print("temp040_1: \n", temp040_1)
print("temp040_2: \n", temp040_2)
print("temp040_3: \n", temp040_3)

data_set = [df_mold_num, df_meas_pos, temp015_1, temp015_2, temp040_1, temp040_2, temp040_3]
print("data_set: \n", data_set)

# Check for N/A values
samp_num_fail_015 = []
samp_arr_fail_015 = []
samp_num_fail_040 = []
samp_arr_fail_040 = []
df_mold_015  = []
df_count_015 = []

data_set_fail = pd.isna(data_set)
print("data_set_fail: \n", data_set_fail )    # Not sure why its printing N/A when there are no N/A's in the dataset

if count_015 == count_040:
    print("Sample lengths are equal, potentially no duplicates.")
else:
    print("Sample lengths are NOT equal, therefore there are duplicates")

#for k in range(0,data_length):
    #if data_set[1][k] == 15 and (int(data_set_fail[2][k]) == 1 or int(data_set_fail[3][k]) == 1): # if the scan was of a 15 position, and it failed position 1 or 2
    #    print("Notice: 015 measurement failed at Mold ", data_set[0][k])
    #    df_mold_015.append(data_set[0][k])
    #    df_count_015.append[k]

    #    if len_check_015 != list_check_015:                                                 # if not equal, then there are duplicates for that mold number
    #        samp_num_fail_015.append(data_set[0][k])
    #    else:                                                                               # Equal, no duplicates
    #        samp_num_fail_015.append(data_set[0][k])
    #else:
    #    pass

    #if data_set[1][k] == 40 and (int(data_set_fail[4][k]) == 1 or int(data_set_fail[5][k]) == 1 or int(data_set_fail[6][k]) == 1): #if the scan was of a 40 position, and it failed position 3, 4, or 5
    #    print("Notice: 040 measurement failed at Mold ", data_set[0][k])
    #    samp_num_fail_040.append(data_set[k])
    #else:
    #    pass

#samp_num_fail_015.append(data_set[0][max(dup_array_015)])
#print("samp_num_fail_015: ", samp_num_fail_015)

#samp_arr_fail_015 = [*set(samp_num_fail_015)]                                               # remove duplicates 
#samp_arr_fail_015.sort(reverse=False)                                                       # sort in ascending order
#print("samp_arr_fail_015: ", samp_arr_fail_015)

#samp_arr_fail_040 = [*set(samp_num_fail_040)]                                               # remove duplicates 
#samp_arr_fail_040.sort(reverse=False)                                                       # sort in ascending order
#print("samp_arr_fail_040: ", samp_arr_fail_040)

#samp_fail_len = (len(samp_arr_fail_015) + len(samp_arr_fail_040))
#print("samp_fail_len: ", samp_fail_len)

#failed_samps_015 = []
#failed_samps_015.clear()
#for i in samp_arr_fail_015:
#    if i <= 25:
#        failed_samps_015.append(str(samp_arr_raw1[i]))
#        print("failed_samps_015 = str(samp_arr_raw1[i]) ", failed_samps_015)
#    elif i > 25 and i < 50:
#        failed_samps_015.append(str(samp_arr_raw2[i]))
#        print("failed_samps_015 = str(samp_arr_raw2[i]) ", failed_samps_015)
#    elif i > 50 and i < 75:
#        failed_samps_015.append(str(samp_arr_raw3[i]))
#        print("failed_samps_015 = str(samp_arr_raw3[i]) ", failed_samps_015)
#    else:
#        print("failed_samps_015 = str(samp_arr_raw[i]) ", str(samp_arr_raw[i])) 

#failed_samps_040 = []
#for i in samp_arr_fail_040:  
#    if i <= 25:
#        failed_samps_040.append(str(samp_arr_raw1[i]))
#        print("failed_samps_040 = str(samp_arr_raw1[i]) ", failed_samps_040)
#    elif i > 25 and i < 50:
#        failed_samps_040.append(str(samp_arr_raw2[i]))
#        print("failed_samps_040 = str(samp_arr_raw2[i]) ", failed_samps_040)
#    elif i > 50 and i < 75:
#        failed_samps_040.append(str(samp_arr_raw3[i]))
#        print("failed_samps_040 = str(samp_arr_raw3[i]) ", failed_samps_040)
#    else:
#        print("failed_samps_040 = str(samp_arr_raw[i]) ", str(samp_arr_raw[i]))

#if samp_fail_len > 0:
#    msg_box1 = tk.messagebox.askquestion('Retest Samples', f'Sample positions 015: {", " . join(map(str, failed_samps_015))} failed.\nSample positions 040: {", " . join(map(str, failed_samps_040))} failed.\nWould you like to retest samples?', icon='question', type='yesno')
#    if msg_box1 == 'yes':
#        print("Yes selected, samples need to be retested")
#        if samp_fail_len > 0:
#            msg_box2 = tk.messagebox.askquestion('Retest Mode', 'Rotate sample 180 degrees.\nYES for manual retest.\nNO for automated test.', icon='question', type='yesno')
#            if msg_box2== 'yes':
#                print("Yes was selected, operator will test manually.")
#                if samp_fail_len > 0:
#                    msg_box3 = tk.messagebox.askquestion('Manual Test Mode', 'Click OK when finished testing manually', icon='info', type='ok')
#                    if msg_box3 == 'ok':
#                        print("Ok was selected, operator finished manually testing samples.")
#                    else:
#                        pass
#            else:
#                print("No was selected, therefore retest will be automated.")
#                retest(failed_samps_015, failed_samps_040)
#                #failed_samps_015.clear()
#                #failed_samps_040.clear()
#    else:
#        print("No was selected, no samples will be retested.")

main.mainloop()
