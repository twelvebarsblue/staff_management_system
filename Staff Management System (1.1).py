from tkinter import *
from tkinter import ttk
from tkcalendar import Calendar
import csv
import pandas as pd
from datetime import datetime
import re
import os

# Theme

import sv_ttk

sms_version = "1.1"

# Create The Gui Object
tk = Tk()
tk.wm_title(f"Staff Management System (Version {sms_version}) (April 2023)")
 
# Set the geometry of the GUI Interface
tk.geometry("700x700")
tk.maxsize(700, 700)

# Add the Calendar module
cal = Calendar(tk, selectmode='day',
               year=2023, month=1,
               day=1)
cal.place(x=0,y=0)

cal2 = Calendar(tk, selectmode='day',
               year=2023, month=1,
               day=1)
cal2.place(x=250,y=0)

# # Dark and Light Mode

# sv_ttk.use_light_theme()

# button = ttk.Button(tk, text="Toggle theme", command=sv_ttk.toggle_theme)
# button.pack()

# Popup message

def popupmsg(title,msg):
    popup = Tk()
    popup.wm_title(title)
    label = Label(popup, text=msg)
    label.pack(side="top", fill="x", pady=10)
    B1 = Button(popup, text="Okay", command = popup.destroy)
    B1.pack()
    popup.mainloop()

    
# Files location

staff_csv = 'Staff/staff.csv'
database_csv = 'Database/database.csv'

# Staff

Label(tk, text="Staff: ").place(x=0, y=200)
staff_menu = StringVar()

try:
    staff_list = pd.read_csv(staff_csv, header = None).iloc[:,0].dropna().unique()
except Exception:
    staff_list = [""]

staff_message = ttk.OptionMenu(tk, staff_menu,"", *staff_list)
staff_message.config(width=20)
staff_message.place(x=0,y=220)

# Movement

Label(tk, text="Movement: ").place(x=0, y=260)
movement_menu = StringVar()

movement_option = ["On Leave", "Time off", "Attend Courses", 
                   "Others"]
movement_message = ttk.OptionMenu(tk, movement_menu, "", *movement_option)
movement_message.config(width=20)
movement_message.place(x=0, y=280)

Label(tk, text="Others/Courses: ").place(x=0, y=310)
others = ttk.Entry(tk)
others.place(x=0, y=330)

# Start Time

Label(tk, text="Start Time: ").place(x=0, y=360)

start_time_menu = StringVar()
start_time_option = ("7:00", "8:00", "9:00", "10:00", "11:00", "12:00",
                     "13:00", "14:00", "15:00", "16:00", "17:00")
start_time_message = ttk.OptionMenu(tk, start_time_menu, "", *start_time_option)
start_time_message.config(width = 20)
start_time_message.place(x=0, y=380)

# End Time

Label(tk, text="End Time: ").place(x=0, y=420)

end_time_menu = StringVar()
end_time_option = ("8:00", "9:00", "10:00", "11:00", "12:00",
                     "13:00", "14:00", "15:00", "16:00", "17:00", "18:00")
end_time_message = ttk.OptionMenu(tk, end_time_menu, "", *end_time_option)
end_time_message.config(width = 20)
end_time_message.place(x=0, y=440)

# Save to Database 

def save_file():
    start_date = pd.to_datetime(cal.get_date(), format = "%m/%d/%y").date() 
    end_date = pd.to_datetime(cal2.get_date(), format = "%m/%d/%y").date()
    total_date = pd.date_range(start_date, end_date)
    name = staff_menu.get() 
    movement = movement_menu.get()
    others_entry = others.get()
    start_time = start_time_menu.get()
    end_time = end_time_menu.get()
    
    # Do some checks first
    
    if total_date.empty:
        popupmsg("Warning","End date cannot be earlier than start date")
    
    if start_time == "" or end_time == "":
        popupmsg("Warning","Please select a time")
        
    else: 
        start_time_convert = pd.to_datetime(start_time).time()
        end_time_convert = pd.to_datetime(end_time).time()
        time_difference = end_time_convert.hour-start_time_convert.hour
        
    if  time_difference <= 0:
        popupmsg("Warning","End time cannot be earlier than start time")
        
    if name == "":
        popupmsg("Warning","Please select a staff")

    if movement == "":
        popupmsg("Warning","Please select movement") 
    
    if movement in ['Others', 'Attend Courses']:
        if others_entry == "":
            popupmsg("Warning","Please specify a course or movement") 

    for d in total_date:
        date = d.date() 

        if movement == "Others":
            with open(database_csv, mode='a') as database:
                database_writer = csv.writer(database, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
                database_writer.writerow([date,start_time_convert,end_time_convert,name,others_entry])
                database.close()

        if movement == "Attend Courses":
            with open(database_csv, mode='a') as database:
                database_writer = csv.writer(database, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
                database_writer.writerow([date,start_time_convert,end_time_convert,name,movement,others_entry])
                database.close()

        if movement not in ["Others", "Attend Courses"]:        
            if date.weekday() < 5: # Exclude weekends
                with open(database_csv, mode='a') as database:
                    database_writer = csv.writer(database, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
                    database_writer.writerow([date,start_time_convert,end_time_convert,name,movement])
                    database.close()
        
        

ttk.Button(tk, text="Save To Database", command=save_file).place(x=0, y=490)

# Clear date by staff

def clear_date_by_staff():
    staff = staff_menu.get()
    start_date = pd.to_datetime(cal.get_date(), format = "%m/%d/%y")
    end_date = pd.to_datetime(cal2.get_date(), format = "%m/%d/%y")
    total_date = pd.date_range(start_date, end_date)
    if staff == "":
        popupmsg("Warning","Please select a staff")
    if total_date.empty:
        popupmsg("Warning","End date cannot be earlier than start date")
    else:
        database = pd.read_csv(database_csv)
        database['Date'] =  pd.to_datetime(database['Date'], format="%Y-%m-%d")
        database = database[~((database['Date'] >= start_date) & (database['Date'] <= end_date) & (database['Staff'] == staff))]
        database.to_csv(database_csv, index=False)


ttk.Button(tk, text="Clear Date (Staff)", command=clear_date_by_staff).place(x=0, y=520)

# Clear all by staff

def staff_clear():
    staff = staff_menu.get()
    if staff == "":
        popupmsg("Warning","Please select a staff")
    else:
        database = pd.read_csv(database_csv)
        database = database[~(database['Staff'] == staff)]
        database.to_csv(database_csv, index=False)
    
ttk.Button(tk, text="Clear All (Staff)", command=staff_clear).place(x=0, y=550)


# Export staff to csv

def export_staff_csv():
    staff = staff_menu.get()
    if staff == "":
        popupmsg("Warning","Please select a staff")
    else:
        database = pd.read_csv(database_csv)
        database = database[database['Staff'] == staff]
        database.to_csv(f"Export/{staff}.csv", index=False)


ttk.Button(tk, text="Export CSV (Staff)", command=export_staff_csv).place(x=0, y=580)

# Export date to csv

def export_date_csv():
    start_date = pd.to_datetime(cal.get_date(), format = "%m/%d/%y")
    end_date = pd.to_datetime(cal2.get_date(), format = "%m/%d/%y")
    total_date = pd.date_range(start_date, end_date)
    if total_date.empty:
        popupmsg("Warning","End date cannot be earlier than start date")
    else:
        database = pd.read_csv(database_csv)
        database['Date'] =  pd.to_datetime(database['Date'], format="%Y-%m-%d")
        database = database[((database['Date'] >= start_date) & (database['Date'] <= end_date))]
        database.to_csv(f"Export/movement_on_{start_date.date()}_to_{end_date.date()}.csv", index=False)


ttk.Button(tk, text="Export CSV (Date)", command=export_date_csv).place(x=0, y=610)

# Export staff date to csv

def export_staff_date_csv():
    start_date = pd.to_datetime(cal.get_date(), format = "%m/%d/%y")
    end_date = pd.to_datetime(cal2.get_date(), format = "%m/%d/%y")
    total_date = pd.date_range(start_date, end_date)
    staff = staff_menu.get()
    if total_date.empty:
        popupmsg("Warning","End date cannot be earlier than start date")
    if staff == "":
        popupmsg("Warning","Please select a staff")
    else:
        database = pd.read_csv(database_csv)
        database['Date'] =  pd.to_datetime(database['Date'], format="%Y-%m-%d")
        database = database[((database['Date'] >= start_date) & (database['Date'] <= end_date) & (database['Staff'] == staff))]
        database.to_csv(f"Export/{staff}_movement_on_{start_date.date()}_to_{end_date.date()}.csv", index=False)


ttk.Button(tk, text="Export CSV (Staff + Date)", command=export_staff_date_csv).place(x=0, y=640)

          
# Add or remove staff

Label(tk, text="Add/Remove Staff: ").place(x=580, y=200)
name_message = ttk.Entry(tk)
name_message.place(x=560, y=220)

def save_new_staff():
    global staff_message, staff_list
    name = name_message.get()
    if name == "":
        popupmsg("Warning","Please input staff name")
    if name in staff_list:
        popupmsg("Warning","Staff already exists")
    else:
        with open(staff_csv, mode='a') as staff_db:
            database_writer = csv.writer(staff_db, delimiter=',', quotechar='"', quoting=csv.QUOTE_MINIMAL, lineterminator='\n')
            database_writer.writerow([name])
            staff_db.close()
            # Refresh
            staff_list = pd.read_csv(staff_csv, header = None).iloc[:,0].dropna().unique()
            staff_message.destroy()
            staff_message = ttk.OptionMenu(tk, staff_menu,"", *staff_list)
            staff_message.config(width=20)
            staff_message.place(x=0,y=220)
            

ttk.Button(tk, text="Add", command=save_new_staff).place(x=535, y=250)

def delete_staff():
    global staff_message, staff_list
    name = name_message.get()
    if name == "":
        popupmsg("Warning","Please input staff name")
    if name in staff_list:    
        database = pd.read_csv(database_csv)
        list_of_staff = pd.read_csv(staff_csv, header = None)
        database = database[~(database['Staff'] == name)]
        database.to_csv(database_csv, index=False)  
        list_of_staff = list_of_staff[~(list_of_staff == name)]
        list_of_staff.to_csv(staff_csv, index=False, header = False)
        # Refresh
        staff_list = pd.read_csv(staff_csv, header = None).iloc[:,0].dropna().unique()
        staff_message.destroy()
        staff_message = ttk.OptionMenu(tk, staff_menu,"", *staff_list)
        staff_message.config(width=20)
        staff_message.place(x=0,y=220)
        
    else:
        popupmsg("Warning",f"{name} is not in staff list")


ttk.Button(tk, text="Remove", command=delete_staff).place(x=611, y=250)


# Backup

backup_list = []

for filename in os.scandir('Backup/Database'):
    backup_list.append(filename.name)
    
backup_date = []

for value in backup_list:
    backup_date.append(re.findall("\d{4}[/-]\d{2}[/-]\d{2}", value)[0])

ttk.Label(tk, text="Backup Database: ").place(x=590, y=490)

backup_menu = StringVar()
backup_menu.set("")
backup_message = ttk.OptionMenu(tk, backup_menu,"", *backup_date)
backup_message.config(width = 20)
backup_message.place(x=535,y=520)                       
    
def load_backup():
    date = backup_menu.get()
    
    if date == '':
        popupmsg("Warning","Please select a backup")
        
    else:
        global staff_csv
        global database_csv
    
        staff_csv = f'Backup/Staff/staff_{date}.csv'
        database_csv = f'Backup/Database/database_{date}.csv'


ttk.Button(tk, text="Load", command=load_backup).place(x=520, y=550)

def backup_database():
    date = datetime.now()
    database = pd.read_csv(database_csv)
    staff = pd.read_csv(staff_csv)
    database.to_csv(f"Backup/Database/database_{date.date()}.csv", index=False)
    staff.to_csv(f"Backup/Staff/staff_{date.date()}.csv", index=False)
    # Refresh
    global backup_list, backup_date, backup_menu, backup_message
    
    backup_list = []

    for filename in os.scandir('Backup/Database'):
        backup_list.append(filename.name)

    backup_date = []

    for value in backup_list:
        backup_date.append(re.findall("\d{4}[/-]\d{2}[/-]\d{2}", value)[0])
        
    backup_message.destroy()
    backup_menu = StringVar()
    backup_menu.set("")
    backup_message = ttk.OptionMenu(tk, backup_menu,"", *backup_date)
    backup_message.config(width = 20)
    backup_message.place(x=535,y=520)  



ttk.Button(tk, text="Backup Current", command=backup_database).place(x=595, y=550)


# Output



def view_date():
    # Init text box
    global scroll_v, scroll_h, text
    scroll_v = Scrollbar(tk)
    scroll_h = Scrollbar(tk, orient= HORIZONTAL)
    text = Text(tk, height= 100, width= 100, yscrollcommand= scroll_v.set, xscrollcommand = scroll_h.set, wrap= NONE, font= ('Helvetica 8'))  
    text.delete('1.0', END)
    
    start_date = pd.to_datetime(cal.get_date(), format = "%m/%d/%y")
    end_date = pd.to_datetime(cal2.get_date(), format = "%m/%d/%y")
    total_date = pd.date_range(start_date, end_date)
    staff = staff_menu.get()
    if total_date.empty:
        popupmsg("Warning","End date cannot be earlier than start date")
    else:
        database = pd.read_csv(database_csv)
        database['Date'] =  pd.to_datetime(database['Date'], format="%Y-%m-%d")
        date_movement = database[((database['Date'] >= start_date) & (database['Date'] <= end_date))]

        #Add a Vertical Scrollbar

        scroll_v.place(x = 509, y = 190, height = 480, width = 20)

        #Add a Horizontal Scrollbar

        scroll_h.place(x = 160, y = 670, height = 20, width = 350)
        #Add a Text widget

        text.place(x = 160, y = 190, height = 480, width = 350)
        text.insert(END, date_movement)

        #Attact the scrollbar with the text widget
        scroll_h.config(command = text.xview)
        scroll_v.config(command = text.yview)


    
Button(tk, text="By Date", command=view_date).pack(anchor = 's', side=LEFT)

def view_staff():
    global scroll_v, scroll_h, text
    scroll_v = Scrollbar(tk)
    scroll_h = Scrollbar(tk, orient= HORIZONTAL)
    text = Text(tk, height= 500, width= 350, yscrollcommand= scroll_v.set, xscrollcommand = scroll_h.set, wrap= NONE, font= ('Helvetica 8'))  
    text.delete('1.0', END)
    
    start_date = pd.to_datetime(cal.get_date(), format = "%m/%d/%y")
    end_date = pd.to_datetime(cal2.get_date(), format = "%m/%d/%y")
    total_date = pd.date_range(start_date, end_date)
    staff = staff_menu.get()
    if total_date.empty:
        popupmsg("Warning","End date cannot be earlier than start date")
    if staff == "":
        popupmsg("Warning","Please select a staff")
    else:
        database = pd.read_csv(database_csv)
        database['Date'] =  pd.to_datetime(database['Date'], format="%Y-%m-%d")
        staff_movement = database[((database['Date'] >= start_date) & (database['Date'] <= end_date) & (database['Staff'] == staff))]

    #Add a Vertical Scrollbar

    scroll_v.place(x = 509, y = 190, height = 480, width = 20)

    #Add a Horizontal Scrollbar

    scroll_h.place(x = 160, y = 670, height = 20, width = 350)
    #Add a Text widget

    text.place(x = 160, y = 190, height = 480, width = 350)
    text.insert(END, staff_movement)

    #Attact the scrollbar with the text widget
    scroll_h.config(command = text.xview)
    scroll_v.config(command = text.yview)

Button(tk, text="By Staff", command=view_staff).pack(anchor = "s", side=LEFT)


# About
    
Button(tk, text="About", command=lambda: popupmsg("About", f"Staff Management System\nAuthor: Ten Yi Yang, under the supervision of Dr. Low Ee Vien\nEmail: tenyiyang.psh@moh.gov.my\nVersion {sms_version} (April 2023)")).pack(anchor = "s", side=RIGHT)


# Close Output Screen

def close_output_screen():
    global scroll_v, scroll_h, text
    scroll_v.destroy()
    scroll_h.destroy()
    text.destroy()


Button(tk, text="Close Screen", command = close_output_screen).pack(anchor = "s", side=RIGHT)

# Execute Tkinter
tk.mainloop()
