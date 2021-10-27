# rheysys assistant developed by Vishal Patel (vishalresearch.com)
# vishal@vishalresearch.com
import ctypes
import sys
import time
import tkinter as tk
from tkinter import filedialog as of

import keyboard
import pyautogui as ui
import keyboard as kb # for the keypress use 'send' function otherwise u need to use press and release in pair
import os
import pandas as pd



#variables
title='Oracle Fusion Middleware Forms Services - WebUtil - I' # full name is internet explorer, but when
w_title= 'RHESYS Assist v1.0' # This variable contains windows title
uid=pwd=''
rnum=0 # record number
delay=3
app_icon='img/icon.ico'
# citix is ued full name is not shown in citrix Ie. To support citrix title needs to be till I.
# if we end title earlier then it can take any other browser as main browser
# to make sure we are opening Internet explorer, title has to finish on I.

# get screen size
s_info= tk.Tk()
s_info.update_idletasks()
s_width=s_info.winfo_screenwidth()
s_height=s_info.winfo_screenheight()
s_info.destroy()



# ----------------------------------------------
# main ui functions
# ----------------------------------------------
def shut():
    sys.exit()
    exit(0)
def launch_ie(): #function to launch internet explorer
    # note: if internet explorer is launched using this function then it will be closed when clos this script/program
    # currently it is unused
    url = "https://researchuat.uws.edu.au/forms/frmservlet?config=rmsmenu"
    ppath = os.environ['PROGRAMFILES'].split('\\')
    ppath = ppath[0] + '\\' + ppath[1] + '\Internet Explorer\iexplore.exe'

    command = "\"%s\" " % ppath + ' %s' % url
    os.system(command)

def login_(uid,pwd): # this function is to perfrom login
    # get cursor to first field (safety)
    #ui.hotkey('ctrl', 'up')
    kb.send('ctrl+up')
    # type user id
    ui.typewrite(uid)
    ui.sleep(0.1)
    # move to next field
    kb.send('tab')
    ui.sleep(0.1)
    # type password
    ui.typewrite(pwd)
    ui.sleep(0.1)
    # login
    kb.send('tab')

def check(win):
    print('Checking windows in all window list')
    titles = ui.getAllTitles()
    # check if orcale fussion is opened or not
    err=True # assume that our app is not opened
    for t_list in titles:
        if title in t_list:
            err=False
    if err:
        print(titles)
        ui.alert('Please try again after opening RHESYS page in internet explorer')

def show_win(title): # function to activate ie window
    print('Show win')
    # this function needs to be called everytime after clicking on any button
    # it will activate ie window
    # note: make sure automator and Ie are on same screen

    '''x=ui.getWindowsWithTitle(title)[0].centerx
    y=ui.getWindowsWithTitle(title)[0].centery
    if s_width-x>s_width or s_height-y>s_height or s_width-x<0 or s_height-y<0:
        ui.alert('Move IE and Automator to main window',title=w_title)'''

    # else:
    # ------------------------------------
    # need to work on moving window back to main screen till then make sure that window and automator are in
    # same screen
    #ui.getWindowsWithTitle(title)[0].moveTo(s_width,
    #                                        s_height)  # move ie window to main screen to avoid problem finding in second screen
    # -------------------------------------
    ui.getWindowsWithTitle(title)[0].activate()
    ui.getWindowsWithTitle(title)[0].maximize()
    ui.getWindowsWithTitle(title)[0].activate()
    ui.sleep(0.1)
    ui.moveTo(s_width/2,s_height/2) # move mouse to center of the screen
    ui.sleep(0.3)

def switch(title,uid,pwd): # this function is to bring RHESYS windows in forground
    ui.getWindowsWithTitle(title)[0].activate()
    ui.getWindowsWithTitle(title)[0].maximize()
    # wait for 1 second
    time.sleep(2)
    # call login function
    login_(uid,pwd)

# ----------------------------------------------
# navigation function
# ----------------------------------------------

def new_project(): # thi function is to create new project entry
    show_win(title) # get ie window
    kb.send('alt')
    kb.send('p')
    ui.sleep(0.3)
    kb.send('p')
def get_data(window,f_path,fn): # This function reads excel file data
    if fn=='data':
        data=pd.read_excel(f_path) # open excel file
    else:
        try:
            data = pd.read_excel(f_path,'Grant Tracker')  # open excel file
        except:
            ui.alert('Wrong Grant Tracker File, Try again', title=w_title)
            window.destroy()  # destroy main window
            main_win()  # reopen main window
    cols=data.shape[1] # get number of columns
    # this feature will prevent user from selecting wrong file
    if cols!=81 and fn=='data':
        ui.alert('Wrong Data File, Try again',title=w_title)
        window.destroy() # destroy main window
        main_win() # reopen main window
    elif cols!=83 and fn=='gt':
        ui.alert('Wrong Grant Tracker File, Try again', title=w_title)
        window.destroy()  # destroy main window
        main_win()  # reopen main window
    return data
def erase_txt():
    ui.sleep(2)
    kb.send('end')
    ui.sleep(0.5)
    kb.send('shift+home')
    ui.sleep(0.2)
    kb.send('backspace')
    ui.sleep(0.2)
    #kb.send('backspace')
def error_data(): # show error when data is missing in file
    ui.alert("Missing Information Check Data File\n or Fill Manually.",title=w_title)
def error_loc(): #show error when unable to locate text on screen (dev error)
    ui.alert(" Unable to locate element on screen\n Contact developer.",title=w_title)
def fill_prj_tile(data): # this function fills information in project sub tab
    # ********
    # REQ
    # ********
    print(data['Project title'][rnum])
    show_win(title)
    ui.scroll(1000) # positive value to scroll up
    loc=ui.locateOnScreen('img/pr_title.png',grayscale=False,confidence=0.8)
    if loc==None:
        error_loc()
    else:
        ui.click(loc[0]+loc[0]/4,loc[1]+loc[1]/8)
        kb.write(data['Project title'][rnum])
    ui.sleep(delay)
def fill_prj_des(data): # this function fills information in project sub tab
    show_win(title)
    print('Writing description.')
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/pr_des.png',grayscale=False,confidence=0.8)
    if loc==None:
        error_loc()
    else:
        ui.click(loc[0] + loc[0]*2, loc[1] + loc[1]/8)
        kb.write(data['Project description'][rnum])
    ui.sleep(delay)
def fill_start(data): # this function fills information in project sub tab
    print('Writing Start date.')
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/pr_start.png',grayscale=False,confidence=0.8)
    if loc==None:
        error_loc()
    else:
        ui.click(loc[0] + 100, loc[1] + 10)
        sd=str(data['Start date'][rnum]).split(' ')[0]

        # convert date from yymmdd to ddmmyy
        sd=sd.split('-')
        # check whether data is there in the data or not
        if len(sd)<=1:
            ui.alert('Project Start Date not found fill it manually.',title=w_title)
        else:
            sd=sd[2]+'/'+sd[1]+'/'+sd[0]
            kb.write(sd)
    ui.sleep(delay)
def fill_fin(data): # this function fills information in project sub tab
    print('Writing finish date.')
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/pr_fin.png',grayscale=False,confidence=0.8)
    if loc==None:
        error_loc()
    else:
        ui.click(loc[0] + 100, loc[1] + 10)
        ed = str(data['End date'][rnum]).split(' ')[0]
        # convert date from yymmdd to ddmmyy
        ed = ed.split('-')
        # check whether data was there in the data or not
        if len(ed)<=1:
            ui.alert('Project Finish Date not found fill it manually.',title=w_title)
        else:
            ed = ed[2] + '/' + ed[1] + '/' + ed[0]
            kb.write(ed)
    ui.sleep(delay)
def fill_res(data):
    print('Writing research%.')
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/res_per.png',grayscale=False,confidence=0.9)
    if loc==None:
        error_loc()
    else:
        ui.click(loc[0] + 100, loc[1] + 10)
        r_per=str(data['% Research'][rnum])
        # check if it is a number or not
        print('rpr:',r_per)
        if not (r_per=='nan'):
            kb.write(r_per)
        else:
            ui.alert('Research Percentage not found fill it manually.',w_title)
    ui.sleep(delay)
def fill_overheads(data):
    print('Writing levy.')
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/levy.png',grayscale=False,confidence=0.8)
    if loc==None:
        error_loc()
    else:
        ui.click(loc[0] + 100, loc[1] + 10)
        kb.write(str(data['Overheads'][rnum]))
        kb.send('enter')
    ui.sleep(delay)
def fill_ip_owner(data):
    print('Writing Ip owner.')
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/ip_owner.png',grayscale=False,confidence=0.8)
    if loc==None:
        error_loc()
    else:
        ui.click(loc[0] + 100, loc[1] + 10)
        kb.write(str(data['Project IP arrangement'][rnum]))
    ui.sleep(delay)
def fill_rdo_bdo(data): #need to work (drop down)
    print('Writing rdo bdo.')
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/rdo_bdo.png',grayscale=True,confidence=0.5)
    if loc==None:
        error_loc()
    else:
        kb.write(str(data['BD contact'][rnum]))
        kb.send('tab')
    ui.sleep(delay)
def fill_school(data):
    # ********
    # REQ
    # ********
    print('Writing school.')
    # switch to ie window
    show_win(title)
    # scroll down
    ui.scroll(-1000)
    loc = ui.locateOnScreen('img/nos.png',grayscale=False,confidence=0.8)  # find field
    if loc==None:
        error_loc()
    else:
        ui.click(loc[0] + loc[0]/8, loc[1] + loc[1]/8)  # click on it
        school_name=str(data['School/Institute/Centre'][rnum])
        print(school_name)
        kb.write(school_name)
        kb.send('enter')
    ui.sleep(delay)
def fill_cheif_in(data,btn): # double check everyting (it takes more button as argument)
    # *********
    # REQ
    # *********
    # note: there is a bug in RHESYS
    # when you press clear all menu in RHESYS it also removes label of chief investigator
    # this may lead to error finding chief investigator text field
    # **change its searching style**
    # urgent (6/10/21)
    # **make mechanism to deal with multiple researchers with same name**
    print('Writing CI data.')

    # disable fill more button (in case someone fill information in more form manually)
    btn['state']='disabled'

    r_names = str(data['Researcher name'][rnum]).split(';')
    r_names = ''.join(r_names).split('#')
    r_name=[]
    for i in r_names:
        if not (i.isnumeric()):
            r_name.append(i) # add all researcher names to r_name
    # check if there are more than one researcher
    multi=False # flag for multiple researcher (initially we assume that there is only one researcher)
    if len(r_name)>1:
        multi=True
    # add first researcher
    show_win(title)
    ui.scroll(1000)
    ui.sleep(delay)
    loc = ui.locateOnScreen('img/c_cin.png',grayscale=False,confidence=0.8) # find field
    if loc==None:
        error_loc()
    else:
        ui.click(loc[0] + 15, loc[1] + 25) # click on it
        # erase if anything is written (here we can not use erase text function this field is different)
        ui.sleep(2)
        kb.write('123')
        kb.send('enter')
        # now everything is erased
        kb.write('%')
        kb.send('enter') # search in full list
        ui.sleep(delay+1)
        loc = ui.locateOnScreen('img/find.png',grayscale=False,confidence=0.8)  # find field
        if loc==None:
            error_loc()
        else:
            ui.click(loc[0] + 100, loc[1] + 10)  # click on it
            erase_txt()

            name=r_name[0].split(' ') # divide first name and last name
            kb.write(name[1]+'%'+name[0]) # write lastname % first name
            kb.send('enter')
            ui.sleep(1)
            kb.send('enter')
            ui.sleep(delay)
            # fill school data first
            # because to goto more we need to save this record and to save this record
            # school information is required
            fill_school(data)
            ui.sleep(delay)
            # save record
            ui_save()
            # Click on more if more researchers exits
            if multi:
                btn['state']='normal'
                ui.sleep(delay)
                # click on more
                ui.scroll(1000)
                loc = ui.locateOnScreen('img/btn_more.png',grayscale=False,confidence=0.8)  # find field
                if loc == None:
                    error_loc()
                else:
                    ui.click(loc[0] + 10, loc[1] + 10)  # click on it
                    ui.alert('There are more than one researcher!!\nClick on Fill More once more researcher window is opened.',w_title)
    ui.sleep(delay)

def fill_more(data,btn):
    print('Fill more')
    show_win(title)
    r_names = str(data['Researcher name'][rnum]).split(';')
    r_names = ''.join(r_names).split('#')
    r_name = []
    for i in r_names:
        if not (i.isnumeric()):
            r_name.append(i)  # add all researcher names to r_name
    btn['state']='disabled' # once fill more is used disable it again
    for res_index in range(1,len(r_name)):
        kb.send('ctrl+down')
        # erase if anything is written (here we can not use erase text function this field is different)
        ui.sleep(2)
        kb.write('123')
        kb.send('enter')
        # now everything is erased
        kb.write('%')
        kb.send('enter')  # search in full list
        ui.sleep(delay + 1)
        show_win(title) # we are activating the ie again in case other applications pop up
        loc = ui.locateOnScreen('img/find.png', grayscale=False, confidence=0.8)  # find field
        if loc == None:
            error_loc()
        else:
            ui.click(loc[0] + 100, loc[1] + 10)  # click on it
            erase_txt()

            name = r_name[res_index].split(' ')  # divide first name and last name
            kb.write(name[1] + '%' + name[0])  # write lastname % first name
            # pressing enter to search
            kb.send('enter')
            # checking if there are more than one name
            ui.sleep(delay)
            loc = ui.locateOnScreen('img/inv_win.png', grayscale=False, confidence=0.8)  # find field
            if not(loc == None):
                ui.alert('There are multiple names, select then name manually then click on ok.')
                kb.send('enter')
                ui.sleep(delay)
            else:
                ui.sleep(delay)
                # keeping fixed position of Investigator for now
                # change it on further instruction
            kb.write('Investigator')
    ui.sleep(delay)
    # save record
    ui_save()
    # exit current window
    #ui_exit()
def fill_project(data,btn):
    print('-----------------------')
    fill_prj_tile(data) # checked
    fill_prj_des(data) # chekced
    fill_start(data) # checked
    fill_fin(data) # checked
    fill_res(data) # checked
    #fill_overheads(data) # in meeting with Alison we decided not to fill it
    #fill_ip_owner(data) # no data (blank meeting)
    #fill_rdo_bdo(data) # no data (blank meeting)
    fill_cheif_in(data,btn) # checked
    #fill_more(data,btn) # only for testing
    #fill_start_date()
    #fill_school(data)
#--------------------------------------------------------------
def next_record(lbl): # This function will read next record
    global rnum
    rnum+=1
    lbl.update()
def ui_more(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(1000)
    loc = ui.locateOnScreen('img/btn_more.png',grayscale=True,confidence=0.5)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+3,loc[1]+3)
def ui_c_cmnt(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(1000)
    loc=ui.locateOnScreen('img/btn_c_comments.png',grayscale=True,confidence=0.5)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+3,loc[1]+3)
def ui_grants(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_grants.png',grayscale=True,confidence=0.5)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+3,loc[1]+3)
def ui_comments(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_comments.png',grayscale=True,confidence=0.5)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+3,loc[1]+3)
def ui_ua(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_user_action.png',grayscale=True,confidence=0.5)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+3,loc[1]+3)
def ui_status(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_status.png',grayscale=True,confidence=0.5)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+3,loc[1]+3)
def ui_forms(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_forms.png',grayscale=True,confidence=0.5)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+3,loc[1]+3)
def ui_keywords(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_keyword.png',grayscale=True,confidence=0.5)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+3,loc[1]+3)
def ui_links(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_links.png',grayscale=True,confidence=0.5)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+3,loc[1]+3)
def ui_save(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_save.png',grayscale=False,confidence=0.8)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+6,loc[1]+6)
def ui_exit(): # this function finds the location of button more and click on it
    show_win(title)
    ui.scroll(-1000)
    loc=ui.locateOnScreen('img/btn_exit.png',grayscale=False,confidence=0.8)
    if loc == None:
        error_loc()
    else:
        ui.click(loc[0]+6,loc[1]+6)

# ----------------------------------------------
# windows/ gui
# ----------------------------------------------

def login_win():
    global uid,pwd
    # login window creatio
    login=tk.Tk()
    login.title('Login. . .')
    login.resizable(0,0)


    login.geometry("+%s"%(s_width-110)+"-%s"%(s_height-130))
    login.protocol('WM_DELETE_WINDOW',lambda : ui.alert('Click on exit'))
    #tk variables
    u_id=tk.StringVar()
    p_wd=tk.StringVar()

    uid=tk.Entry(login,textvariable=u_id,width=15).grid(row=0,column=1)
    pwd=tk.Entry(login,textvariable=p_wd,show='*',width=15).grid(row=1,column=1)
    login_btn=tk.Button(login,text='Login',command=lambda :check(login)).grid(row=2,column=1)
    exit_btn=tk.Button(login,text='Exit',command=lambda :exit(0)).grid(row=3,column=1)
    uid = u_id.get()
    pwd = p_wd.get()
    switch(title, uid, pwd)
    login.mainloop() # main loop for login window



#-------------------------------------
def main_win():
    global rnum
    # check if RHESYS is opend or not
    check(title)
    # creation of control window
    control=tk.Tk()
    control.title(w_title)
    control.geometry("+%s"%(s_width-390)+"-%s"%(s_height-680))
    control.protocol('WM_DELETE_WINDOW',lambda : ui.alert('Click on exit'))
    control.iconbitmap(app_icon)
    control.resizable(0,0)

    # ask for data the file name
    ftypes=(
        ('Excel Workbook', '*.xlsx'),
        ('Excel 97- Excel 2003 Workbook', '*.xls')
    )
    file=of.askopenfilename(title='Select Share point Excel File',
                            filetypes=ftypes
                            )
    if file=='':
        ui.alert('You didn\'t selecte a file, exiting...', title=w_title)
        exit(0)
    data=get_data(control,file,'data')
    # storing ids into list_ids it is used to link index number
    list_ids = data['ID'].tolist()
    #-------------------------------------------------------------------------

    # ask for grant tracker file
    ftypes = (
        ('Excel Workbook', '*.xlsx'),
        ('Excel 97- Excel 2003 Workbook', '*.xls')
    )
    file = of.askopenfilename(title='Select Grant Tracker Excel File',
                              filetypes=ftypes
                              )
    if file == '':
        ui.alert('You didn\'t selecte a file, exiting...', title=w_title)
        exit(0)
    data_gt = get_data(control, file,'gt')
    # storing ids into list_ids it is used to link index number
    gt_list_ids = data_gt['ID'].tolist()


    # function to copy data to clip board
    def copydata(data):
        control.clipboard_clear()
        control.clipboard_append(data.get())
        # ui.alert('Copied!!!',w_title)

    # ------------------------------------------------------------------------
    # function to search for id and update content
    def search_update():
        global rnum # to change main rnum variable
        try:
            id_number = int(search_id_number.get())
            rnum=list_ids.index(id_number) # record number for data
            rnum_gt=gt_list_ids.index(id_number) # record number for grant tracker

            # project title
            record=str(data['Project title'][rnum])
            if record!='nan':
                lbl_title = tk.Label(text='Project title').grid(row=1, column=0,sticky=tk.W)
                pr_title_var = tk.StringVar()  # project title
                pr_title_var.set(record)
                entry_title = tk.Entry(textvariable=pr_title_var)
                entry_title.grid(row=2,columnspan=2,sticky=tk.E+tk.W)
                entry_title['state'] = 'disabled'
                # function biding
                entry_title.bind("<Button-1>", lambda a: copydata(pr_title_var))
                # scrollbar
                sb_title = tk.Scrollbar(orient='horizontal', width=10)
                sb_title.grid(row=3, columnspan=2, sticky=tk.E + tk.W)
                entry_title.config(xscrollcommand=sb_title.set)
                sb_title.config(command=entry_title.xview)

            # project description label
            record=str(data['Project description'][rnum])
            if record!='nan':
                lbl_des = tk.Label(text='Project Description:').grid(row=4, column=0,sticky=tk.W)
                pr_des_var = tk.StringVar()  # project description
                pr_des_var.set(record)
                entry_des = tk.Entry(textvariable=pr_des_var)
                entry_des.grid(row=5,columnspan=2,sticky=tk.E+tk.W)
                entry_des['state'] = 'disabled'
                # function bidning
                entry_des.bind("<Button-1>", lambda a: copydata(pr_des_var))
                # scrollbar
                sb_des = tk.Scrollbar(orient='horizontal', width=10)
                sb_des.grid(row=6, columnspan=2, sticky=tk.E + tk.W)
                entry_des.config(xscrollcommand=sb_des.set)
                sb_des.config(command=entry_des.xview)
            # update project start label
            record=str(data['Start date'][rnum])
            if record!='nan':
                record=record.split(' ')[0]
                record=record.split('-')
                record = record[2] + '/' + record[1] + '/' + record[0]
                lbl_pstart = tk.Label(text='Project Start:').grid(row=7, column=0,sticky=tk.W)
                pr_start_var = tk.StringVar()  # project start date
                pr_start_var.set(record)
                entry_start = tk.Entry(textvariable=pr_start_var)
                entry_start.grid(row=8,columnspan=2,sticky=tk.E+tk.W)
                entry_start['state']='disabled'
                # function binding
                entry_start.bind("<Button-1>", lambda a: copydata(pr_start_var))

            # update project end label
            record=str(data['End date'][rnum])
            if record!='nan':
                record = record.split(' ')[0]
                record = record.split('-')
                record = record[2] + '/' + record[1] + '/' + record[0]
                lbl_pend = tk.Label(text='Project End:').grid(row=9,sticky=tk.W)
                pr_end_var = tk.StringVar()  # project end date
                pr_end_var.set(record)
                entry_end = tk.Entry(textvariable=pr_end_var)
                entry_end.grid(row=10,columnspan=2,sticky=tk.E+tk.W)
                entry_end['state'] = 'disabled'
                # function binding
                entry_end.bind("<Button-1>", lambda a: copydata(pr_end_var))

            # update % research label
            record=str(data['% Research'][rnum])
            if record!='nan':
                lbl_rpercent = tk.Label(text='% Research:').grid(row=11,sticky=tk.W)
                pr_rpercent_var = tk.StringVar()  # % research
                pr_rpercent_var.set(record)
                entry_rpercent = tk.Entry(textvariable=pr_rpercent_var)
                entry_rpercent.grid(row=12,columnspan=2,sticky=tk.E+tk.W)
                entry_rpercent['state'] = 'disabled'  # entry box must be enabled to insert data
                # function binding
                entry_rpercent.bind("<Button-1>", lambda a: copydata(pr_rpercent_var))

            # update levy label
            record=str(data['Overheads'][rnum])
            if record!='nan':
                lbl_levy = tk.Label(text='Levy:').grid(row=13,sticky=tk.W)
                pr_levy_var = tk.StringVar()  # levy
                pr_levy_var.set(record)
                entry_levy = tk.Entry(textvariable=pr_levy_var)
                entry_levy.grid(row=14,columnspan=2,sticky=tk.E+tk.W)
                entry_levy['state'] = 'disabled'
                # function binding
                entry_levy.bind("<Button-1>", lambda a: copydata(pr_levy_var))

            # update rdo/bdo label
            record=str(data['BD contact'][rnum])
            if record!='nan':
                lbl_rdo_bdo = tk.Label(text='RDO/BDO:').grid(row=15,sticky=tk.W)
                pr_rdobdo_var = tk.StringVar()  # project rdo bdo
                pr_rdobdo_var.set(record)
                entry_rdobdo = tk.Entry(textvariable=pr_rdobdo_var)
                entry_rdobdo.grid(row=16,columnspan=2,sticky=tk.E+tk.W)
                entry_rdobdo['state'] = 'disabled'
                # function binding
                entry_rdobdo.bind("<Button-1>", lambda a: copydata(pr_rdobdo_var))

            # update ip ownership label
            record=str(data['Project IP arrangement'][rnum])
            if record != 'nan':
                lbl_ipowner = tk.Label(text='IP ownership:').grid(row=17,sticky=tk.W)
                pr_ipowner_var = tk.StringVar()  # ip owner
                pr_ipowner_var.set(record)
                entry_ipowner = tk.Entry(textvariable=pr_ipowner_var)
                entry_ipowner.grid(row=18,columnspan=2,sticky=tk.E+tk.W)
                entry_ipowner['state'] = 'disabled'
                # function binding
                entry_ipowner.bind("<Button-1>", lambda a: copydata(pr_ipowner_var))

            # update chief investigator label
            record=str(data['Researcher name'][rnum])
            if record!='nan':
                record=record.split(';')
                record = ''.join(record).split('#')
                r_name = []
                for i in record:
                    if not (i.isnumeric()):
                        r_name.append(i)  # add all researcher names to r_name
                researchers=''
                for i in r_name:
                    researchers=researchers+i+','
                researchers=researchers.rstrip(',')
                lbl_ciname = tk.Label(text='Chief Investigator:').grid(row=19,sticky=tk.W)
                pr_ci_var = tk.StringVar()  # project chief investigator
                pr_ci_var.set(researchers)
                entry_ci = tk.Entry(textvariable=pr_ci_var)
                entry_ci.grid(row=20,columnspan=2,sticky=tk.E+tk.W)
                entry_ci['state'] = 'disabled'
                # function binding
                entry_ci.bind("<Button-1>", lambda a: copydata(pr_ci_var))

            # update name of school label
            record=str(data['School/Institute/Centre'][rnum])
            if record!='nan':
                lbl_sch_cen = tk.Label(text='Name School or Center:').grid(row=21,sticky=tk.W)
                pr_schcen_var = tk.StringVar()  # project school/center
                pr_schcen_var.set(record)
                entry_schcen = tk.Entry(textvariable=pr_schcen_var)
                entry_schcen.grid(row=22,columnspan=2,sticky=tk.E+tk.W)
                entry_schcen['state'] = 'disabled'
                # function binding
                entry_schcen.bind("<Button-1>", lambda a: copydata(pr_schcen_var))
        # ------------------------------------------------------------------------
            # grant  information
            print(data_gt.info())
            # ----------------------------------------
            record=str(data_gt['Project description']) # --> Grant Description
            if record!='nan':
                lbl_gt_des=tk.Label(text='Grant Description:').grid(row=1,column=1,sticky=tk.W)
                gt_des_var=tk.StringVar()
                gt_des_var.set(record)
                entry_gt_des=tk.Entry(textvariable=gt_des_var)
                entry_gt_des.grid(row=2,column=1)

            'Partner organisation'  # --> Funder
            ''
        # ----------------------------------------
        except Exception as x:
            print(x)
            ui.alert('Record Id not found!',w_title)

    # search entry box
    lbl_search=tk.Label(text='Search Id:').grid(row=0,column=0,sticky=tk.W)
    search_id_number = tk.StringVar()
    search_text = tk.Entry(text='Id:', textvariable=search_id_number).grid(row=0, column=1)
    search_btn=tk.Button(text='Search',command=search_update).grid(row=0,column=3,sticky=tk.E)
    exit_btn=tk.Button(text='Close',command=shut).grid(row=23,column=1)
    # configuration to keep control window always on top
    control.attributes('-topmost',True)
    control.mainloop()


main_win()