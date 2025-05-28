#from mimetypes import init                 # Python path      # Allowing 015 040 buttons to equal specific string when clicked
from functools import partial               # Python path      # Partial is used for using two commands on one button
from pathlib   import PureWindowsPath       # pathlib  0.1.0   # Cleans up windows path extensions
from openpyxl  import *                     # openpyxl 3.1.2   # Write to excel, alternatively could use import openpyxl as xl
from tkinter import messagebox              # tk       0.1.0   # Standard windows messageboxes, for exiting the program, warnings, etc..
from PIL import ImageTk, Image              # PILLOW   9.5.0   # Importing a .jpg, .png etc.. for a background in a GUI
import tkinter.filedialog as tfd            # tk       0.1.0   # Not sure what this library is for.
import tkinter as tk                        # tk       0.1.0   # Tkinters Tk class for GUI
from tkinter import ttk                     # tk       0.1.0   # Tkinters Tkk class for GUI
import subprocess                           # Python path      # Opening an executable like calculator
import time                                 # Python path      # For pausing for 5 seconds --> time.sleep(5)
import datetime as dt                       # Python path      # datetime now function
import os                                   # Python path      # psutil 5.9.4  # kill/close an executable like calculator

####################################################################################################################################################
#################################################       1st GUI SCREEN           ###################################################################
####################################################################################################################################################
########################################         INITIALIZING STANDARD DISPLAY          ############################################################
GUI = tk.Tk()                                       # This creates a graphical user interface
GUI.title("All That Glitters is GOLD")              # String of the title
GUI.geometry('800x520')                             # GUI size. Largest is 1920x1080 (full screen)
GUI.configure(background = 'white')                 # Set background color
GUI.option_add('*Font', 'Helvetica 10 italic')      # Set Font for label outputs from user inputs.
GUI.option_add('*foreground', 'black')              # Set text color for label outputs from user inputs.
GUI.option_add('*background', 'white')              # Set background color for label outputs from user inputs.init
########################################               BUTTON PRESS STYLE            ###############################################################
style = ttk.Style()
style.theme_use('default')                           # alt, default, clam and classic
style.map('T.TButton',background=[('active', 'pressed','white'),('!active', 'white'),('active', '!pressed','grey')])  # active, not active, not pressed
style.map('T.TButton',relief    =[('pressed','raised'),('!pressed','raised')])                                          # pressed not pressed
style.configure('T.TButton',font=('Helvetica', '10'))                                                                   # Ttk buttons (ALL)

style.map('B.TButton',background=[('active', 'pressed','white'),('!active', 'white'),('active', '!pressed','grey')])  # active, not active, not pressed
style.map('B.TButton',relief    =[('pressed','raised'),('!pressed','raised')])                                          # pressed not pressed
style.configure('B.TButton',font=('Helvetica', '10'))                                                           # BOLD buttons, save, exit

style.map('P.TButton',background=[('active', 'pressed','#FF69B4'),('!active', 'white'),('active', '!pressed','grey')])# active, not active, not pressed
style.map('P.TButton',relief    =[('pressed','raised'),('!pressed','raised')])                                          # pressed not pressed
style.configure('P.TButton',font=('Helvetica', '10'))                                                           # PINK button

style.configure('W.TLabel',font=('Helvetica', '10'), foreground='black', background='white')  

# Python is serial, so each widget will output in the order listed below;
########################################               GOLD BACKGROUND            ###############################################################
def resize_image(event):
    new_width = event.width
    new_height = event.height
    bg_img = copy_img.resize((new_width, new_height))
    new_img = ImageTk.PhotoImage(bg_img)
    lbl_img.config(image = new_img)
    lbl_img.bg_img = new_img #avoid garbage collection

bg_img = Image.open(r'C:\Users\Susan\OneDrive\Documents\Visual Studio 2022\Python\TkinterGUI_2023-04-07\gold.png')
copy_img = bg_img.copy()
new_img = ImageTk.PhotoImage(bg_img)
lbl_img = ttk.Label(GUI, image = new_img, background='white')
lbl_img.bind('<Configure>', resize_image)
lbl_img.pack(fill=tk.BOTH, expand = True)

#############################################            EXCEL FILE            ###############################################################
wb = Workbook()
ws = wb.active
ws['A1'] = "Credentials"        # Column 1
ws['B1'] = "subWO"              # Column 2
ws['C1'] = "Sample"             # Column 3
ws['D1'] = "Measurements"       # Column 4
ws['E1'] = "Date&Time"          # Column 5
ws['F1'] = "Oct Eq #"           # Column 6
ws['G1'] = "Prod/R&D"           # Column 7

new_line = ws.max_row + 1

#############################################            MAIN BODY            ###############################################################
# Display the command label before the entry box to indicate what information the Operator is to type in
lbl_cmd_date = ttk.Label(GUI, text = "Todays Date:"               , style= 'W.TLabel').place(x=20,y=20)
lbl_cmd_fold = ttk.Label(GUI, text = "Folder Name:"               , style= 'W.TLabel').place(x=20,y=50)
lbl_cmd_file = ttk.Label(GUI, text = "File Name:"                 , style= 'W.TLabel').place(x=20,y=80)
lbl_cmd_cred = ttk.Label(GUI, text = "Enter Operator Credentials:", style= 'W.TLabel').place(x=20,y=110)
lbl_cmd_WO   = ttk.Label(GUI, text = "Enter Work Order Number:"   , style= 'W.TLabel').place(x=20,y=140)
lbl_cmd_samp = ttk.Label(GUI, text = "Enter Sample Sizes:"        , style= 'W.TLabel').place(x=20,y=170)
lbl_cmd_meas = ttk.Label(GUI, text = "Select Measurement Size:"   , style= 'W.TLabel').place(x=20,y=200)
lbl_cmd_oct  = ttk.Label(GUI, text = "Select OCT Equipment:"      , style= 'W.TLabel').place(x=20,y=230)
lbl_cmd_prd  = ttk.Label(GUI, text = "Production:"                , style= 'W.TLabel').place(x=20,y=260)
lbl_cmd_rnd  = ttk.Label(GUI, text = "R&D:"                       , style= 'W.TLabel').place(x=375,y=260)

# Entry boxes to take information from operator
entry_cred = tk.Entry(GUI, width= 10)
entry_cred.focus_set()                                              # Place curosr in the first entry box
entry_cred.place(x=200,y=110)
entry_WO = tk.Entry(GUI, width= 10)                                  
entry_WO.place(x=200,y=140)

# Display the label of what to expect in the following output box
lbl_disp_cred = ttk.Label(GUI, text = "Credentials:"        , style= 'W.TLabel').place(x=20,y=320)
lbl_disp_WO   = ttk.Label(GUI, text = "Work Order Number:"  , style= 'W.TLabel').place(x=20,y=350)
lbl_disp_samp = ttk.Label(GUI, text = "Sample Sizes:"       , style= 'W.TLabel').place(x=20,y=380)
lbl_disp_meas = ttk.Label(GUI, text = "Measurement Size:"   , style= 'W.TLabel').place(x=20,y=410)
lbl_disp_oct  = ttk.Label(GUI, text = "OCT Equipment:"      , style= 'W.TLabel').place(x=20,y=440)
lbl_disp_pr   = ttk.Label(GUI, text = "Production/R&D:"     , style= 'W.TLabel').place(x=20,y=470)

# Display the user inputs as output
lbl_out_date = ttk.Label(GUI, text = f'{dt.datetime.now():%b %d, %Y}' , style= 'W.TLabel').place(x=200,y=20)
lbl_out_cred = ttk.Label(GUI, text = '', width= 4                     , style= 'W.TLabel')
lbl_out_cred.place(x=200,y=320)
lbl_out_WO   = ttk.Label(GUI, text = '', width= 6                     , style= 'W.TLabel')
lbl_out_WO.place(x=200,y=350)

# Functions to display the user inputs as output
def fun_cred():
    global entry 
    cred = entry_cred.get()[:3]                                     # entry_cred is the variable we are passing. Limit 3 characters
    lbl_out_cred.configure(text = cred)                             # Display entry_cred from user on GUI
    ws.cell(column=1, row=new_line, value= entry_cred.get()[:3])
    print(entry_cred.get()[:3])                                     # Print can be removed after development

def fun_WO():
    global entry 
    WO = entry_WO.get()[:6]                                         # entry_WO is the variable we are passing. Limit 3 characters
    lbl_out_WO.configure(text = WO)                                 # Display entry_WO from user on GUI
    ws.cell(column=2, row=new_line, value= entry_WO.get()[:6])
    print(entry_WO.get()[:6])                                      # Print can be removed after development

smpl_sz_1= 'smpl_sz_1'
dio_sz_1 = 'dio_sz_1'
smpl_sz_2= 'smpl_sz_2'
dio_sz_2 = 'dio_sz_2'
smpl_sz_3= 'smpl_sz_3'
dio_sz_3 = 'dio_sz_3'
smpl_sz_4= 'smpl_sz_4'
dio_sz_4 = 'dio_sz_4'
smpl_sz_5= 'smpl_sz_5'
dio_sz_5 = 'dio_sz_5'
smpl_sz_6= 'smpl_sz_6'
dio_sz_6 = 'dio_sz_6'
smpl_sz_7= 'smpl_sz_7'
dio_sz_7 = 'dio_sz_7'
smpl_sz_8= 'smpl_sz_8'
dio_sz_8 = 'dio_sz_8'
smpl_sz_9= 'smpl_sz_9'
dio_sz_9 = 'dio_sz_9'
smpl_sz_10= 'smpl_sz_10'
dio_sz_10 = 'dio_sz_10'
def fun_samp(entry_samp):
    global smpl_sz_1, dio_sz_1, smpl_sz_2, dio_sz_2, smpl_sz_3, dio_sz_3, smpl_sz_4, dio_sz_4, smpl_sz_5, dio_sz_5
    global smpl_sz_6, dio_sz_6, smpl_sz_7, dio_sz_7, smpl_sz_8, dio_sz_8, smpl_sz_9, dio_sz_9, smpl_sz_10, dio_sz_10
    btn_dio = tk.Entry(GUI, width= 10)
    btn_dio.insert(0, smpl_sz_1)
    btn_dio.place(x=200,y=380)
    ws.cell(column=3, row=new_line).value= entry_samp

excel_meas = 'excel_meas_empty'
file_meas  = 'file_meas_empty'
def fun_meas(entry_meas):
    global excel_meas, file_meas
    if entry_meas == '-B':
        excel_meas = 'Posterior'
        file_meas  = '-B'
        btn_015 = tk.Entry(GUI, width= 10)
        btn_015.insert(0,excel_meas+file_meas)
        btn_015.place(x=200,y=410)
    elif entry_meas== '-A':
        excel_meas  = 'Anterior'
        file_meas   = '-A'
        btn_040 = tk.Entry(GUI, width= 10)
        btn_040.insert(0,excel_meas+file_meas)
        btn_040.place(x=200,y=410)
    elif entry_meas== '':
        excel_meas  = 'Full Lens'
        file_meas   = ''
        btn_040 = tk.Entry(GUI, width= 10)
        btn_040.insert(0,excel_meas+file_meas)
        btn_040.place(x=200,y=410)
    fun_cred()
    fun_WO()
    ws.cell(column=4, row=new_line).value= excel_meas 
    print("entry_meas is: ", entry_meas)
    print("excel_meas is: ", excel_meas)
    print("file_meas is: " , file_meas)

excel_pr = 'excel_pr_empty'
file_pr  = 'file_pr_empty'
def fun_pr(entry_pr):
    global excel_pr, file_pr, eq_num_pr
    if entry_pr == '02':                            # Haptics
        excel_pr = 'Haptics'
        file_pr  = 'L02-'
        btn_pr02 = tk.Entry(GUI, width= 17)
        btn_pr02.insert(0,file_pr+' '+excel_pr)
        btn_pr02.place(x=200,y=470)
    elif entry_pr == '06':                          # R&D
        excel_pr   = 'LAL'
        file_pr    = 'L06-'
        btn_rd06 = tk.Entry(GUI, width= 17)
        btn_rd06.insert(0,file_pr+' '+excel_pr)
        btn_rd06.place(x=200,y=470)
    elif entry_pr == '07':                          # Standard Production
        excel_pr   = 'Std Production'
        file_pr    = 'L07-'
        btn_pr07 = tk.Entry(GUI, width= 17)
        btn_pr07.insert(0,file_pr+' '+excel_pr)
        btn_pr07.place(x=200,y=470)
    elif entry_pr == '08':                          # Next Gen LAL+ R&D
        excel_pr   = 'LAL+ R&D'
        file_pr    = 'L08-'
        btn_rd08 = tk.Entry(GUI, width= 17)
        btn_rd08.insert(0,file_pr+' '+excel_pr)
        btn_rd08.place(x=200,y=470)
    elif entry_pr == '00':                          # Calibration
        excel_pr   = 'Calibration'
        file_pr    = 'data'
        btn_cal = tk.Entry(GUI, width= 17)
        btn_cal.insert(0,file_pr+' '+excel_pr)
        btn_cal.place(x=200,y=470)
    ws.cell(column=7, row=new_line).value= excel_pr
    print("entry_pr is: ", entry_pr)
    print("excel_pr is: ", excel_pr)
    print("file_pr is: " , file_pr)

excel_oct = 'excel_oct_empty'
fold_oct  = 'fold_oct_empty'
eq_num_oct  = 'eq_num_empty'
def fun_oct(entry_oct):
    global excel_oct, fold_oct, eq_num_oct
    if entry_oct == 'OCT 1':
        excel_oct = 'EQ# 1364'
        fold_oct  = 'OCT 1'
        eq_num_oct= '1364'
        btn_oct1 = tk.Entry(GUI, width= 15)
        btn_oct1.insert(0,fold_oct+ ' '+excel_oct)
        btn_oct1.place(x=200,y=440)
    elif entry_oct == 'OCT 2':
        excel_oct = 'EQ# 2104'
        fold_oct  = 'OCT 2'
        eq_num_oct= '2104'
        btn_oct2 = tk.Entry(GUI, width= 15)
        btn_oct2.insert(0,fold_oct+ ' '+excel_oct)
        btn_oct2.place(x=200,y=440)
    ws.cell(column=6, row=new_line).value= excel_oct
    print("entry_oct is: ", entry_oct)
    print("excel_oct is: ", excel_oct)
    print("fold_oct is: " , fold_oct)
    print("eq_num_oct is: ", eq_num_oct)

def fun_save():
    ws.cell(column=5, row=new_line).value= dt.datetime.now().strftime("%Y-%m-%d%H:%M:%S")
    if excel_pr== 'Calibration':
        filepath = r"C:\Users\Susan\OneDrive\Documents\Visual Studio 2022\Python\TkinterGUI_2023-04-07/" + excel_oct + ' (Lumedica ' + fold_oct + ')' + '/' + dt.datetime.now().strftime("%m-%d-%y") + 'EQ' + eq_oct_num + 'CALIBRATION'
        filename = 'data_' + dt.datetime.now().strftime("%m-%d-%y") + '_' + entry_WO.get()[:6] + '.xlsx'
    else:
        filepath = r"C:\Users\Susan\OneDrive\Documents\Visual Studio 2022\Python\TkinterGUI_2023-04-07/" + excel_oct + ' (Lumedica ' + fold_oct + ')'
        filename = file_pr + '_' + entry_WO.get()[:6] + file_meas + '.xlsx'

    win_filepath = PureWindowsPath(filepath)
    if not os.path.exists(win_filepath):
        os.makedirs(win_filepath)

    loc = filepath + '/' + filename
    win_loc = PureWindowsPath(loc)
    wb.save(win_loc)

    lbl_out_fil_pat = ttk.Label(GUI, text = win_filepath, style= 'W.TLabel')
    lbl_out_fil_pat.place(x=200,y=50)
    lbl_out_fil_nam = ttk.Label(GUI, text = filename    , style= 'W.TLabel')
    lbl_out_fil_nam.place(x=200,y=80)

    print("Filepath - RAW: \n", filepath)
    print("Filepath - WIN: \n", win_filepath)
    print("Filename:       \n", filename)
    print("Location - RAW: \n", loc)
    print("Location - WIN: \n", win_loc)

    os.startfile(win_loc)                               # Open Excel Workbook
    time.sleep(15)                                      # Wait 15 seconds to allow to open before force closing
    os.system('TaskKill /F /IM Excel.exe')      # Force close Excel Workbook

def open_lum():                                         # Function to open the calculator app
    subprocess.Popen('C:\\Windows\\System32\\calc.exe') # Open windows calculator
    print("Open!")
    time.sleep(10)                                      # Wait 10 seconds
    os.system('TaskKill /F /IM CalculatorApp.exe')      # Task Kill the calculator app
    print("Close!")

def exit_app():                                         # Function send a messagebox when exit is clicked.
    msg_box = tk.messagebox.askquestion('Exit', "Are you sure you want to exit the application?", icon='warning')
    if msg_box == 'yes':
        GUI.destroy()
    else:
        tk.messagebox.showinfo('Exit', "Thanks for staying, please continue.")

btn_pres_cnt = 1
def pink(event):                                             # Function to make the buttons clicked pink every 5 clicks
    global btn_pres_cnt
    if(btn_pres_cnt==5 or btn_pres_cnt==10 or btn_pres_cnt==15 or btn_pres_cnt==20 or btn_pres_cnt==25 or btn_pres_cnt==30):
        style.map('T.TButton',background=[('active', 'pressed','#FF69B4'),('!active', 'white'),('active', '!pressed','grey')])
        style.configure('T.TButton',font=('Helvetica', '10'))  
    else:           # else is normal style
        style.map('T.TButton',background=[('active', 'pressed','whtie'),('!active', 'white'),('active', '!pressed','grey')])
        style.configure('T.TButton',font=('Helvetica', '10'))
    print("btn_pres_cnt = ", btn_pres_cnt)
    btn_pres_cnt +=1

#############################################     BUTTONS TO BE CLICKED            ###############################################################
btn_dio = ttk.Button(GUI, text = 'Dioptics', style= 'T.TButton', command=partial(fun_samp, 'Dioptics'))
btn_dio.bind('<Button-1>', pink)
btn_dio.place(x=200,y=168)

btn_015 = ttk.Button(GUI, text = 'Posterior', style= 'T.TButton', command=partial(fun_meas, '-B'))
btn_015.bind('<Button-1>', pink)
btn_015.place(x=200,y=198)

btn_040 = ttk.Button(GUI, text = 'Anterior', style= 'T.TButton', command=(fun_meas, '-A'))
btn_040.bind('<Button-1>', pink)
btn_040.place(x=280,y=198)

btn_100 = ttk.Button(GUI, text = 'Full Lens', style= 'T.TButton', command=partial(fun_meas, ''))
btn_100.bind('<Button-1>', pink)
btn_100.place(x=360,y=198)

btn_oct1 = ttk.Button(GUI, text = 'OCT 1 EQ# 1364', style= 'T.TButton', command=partial(fun_oct, 'OCT 1'))
btn_oct1.bind('<Button-1>', pink)
btn_oct1.place(x=200,y=228)

btn_oct2 = ttk.Button(GUI, text = 'OCT 2 EQ# 2104', style= 'T.TButton', command=partial(fun_oct, 'OCT 2'))
btn_oct2.bind('<Button-1>', pink)
btn_oct2.place(x=320,y=228)

btn_pr02 = ttk.Checkbutton(GUI, text = 'Haptics', onvalue= 1, offvalue= 0, style= 'T.TButton', command=partial(fun_samp, '02'))
btn_pr02.bind('<Button-1>', pink)
btn_pr02.place(x=200,y=258)

btn_pr07 = ttk.Checkbutton(GUI, text = 'Production', onvalue= 1, offvalue= 0, style= 'T.TButton', command=partial(fun_pr, '07'))
btn_pr07.bind('<Button-1>', pink)
btn_pr07.place(x=280,y=258)

btn_rd06 = ttk.Checkbutton(GUI, text = 'LAL', onvalue= 1, offvalue= 0, style= 'T.TButton', command=partial(fun_pr, '06'))
btn_rd06.bind('<Button-1>', pink)
btn_rd06.place(x=420,y=258)

btn_rd08 = ttk.Checkbutton(GUI, text = 'LAL+', onvalue= 1, offvalue= 0, style= 'T.TButton', command=partial(fun_pr, '08'))
btn_rd08.bind('<Button-1>', pink)
btn_rd08.place(x=500,y=258)

btn_cal = ttk.Checkbutton(GUI, text = 'Calibration', onvalue= 1, offvalue= 0, style= 'T.TButton', command=partial(fun_pr, '00'))
btn_cal.bind('<Button-1>', pink)
btn_cal.place(x=580,y=258)

btn_sav = ttk.Button(GUI, text = 'Save' , style= 'B.TButton', command=partial(fun_save))
btn_sav.bind('<Button-1>', pink)
btn_sav.place(x=500,y=470)

btn_lum = ttk.Button(GUI, text = 'Open' , style= 'B.TButton', command=open_lum)
btn_lum.bind('<Button-1>', pink)
btn_lum.place(x=600,y=470)

btn_exit = ttk.Button(GUI, text = 'Exit', style= 'B.TButton', command=exit_app)
btn_exit.bind('<Button-1>', pink)
btn_exit.place(x=700,y=470)

GUI.mainloop()                                  # Must be at the end of a GUI in order for the app to work b/c windows refreshes constantly.