from openpyxl import load_workbook
from tkinter import *
from tkinter import ttk, filedialog
import os
from datetime import datetime

#


##########################################################################################################
##################### STOPWATCH ####################################################################
#################################################################################################################

counter = 21600
running = False


def counter_label(label):
    def count():
        if running:
            global counter

            # To manage the initial delay.
            if counter == 21600:
                display = "Starting..."
            else:
                tt = datetime.fromtimestamp(counter)
                string = tt.strftime("%H:%M:%S")
                display = string

            label['text'] = display  # Or label.config(text=display)

            # label.after(arg1, arg2) delays by
            # first argument given in milliseconds
            # and then calls the function given as second argument.
            # Generally like here we need to call the
            # function in which it is present repeatedly.
            # Delays by 1000ms=1 seconds and call count again.
            label.after(1000, count)
            counter += 1

    # Triggering the start of the counter.
    count()


# start function of the stopwatch
def Start(label):
    global running
    running = True
    counter_label(label)
    start['state'] = 'disabled'
    stop['state'] = 'normal'
    reset['state'] = 'normal'


# Stop function of the stopwatch
def Stop():
    global running
    start['state'] = 'normal'
    stop['state'] = 'disabled'
    reset['state'] = 'normal'
    running = False


# Reset function of the stopwatch
def Reset(label):
    global counter
    counter = 21600

    # If rest is pressed after pressing stop.
    if running == False:
        reset['state'] = 'disabled'
        label['text'] = 'Welcome!'

    # If reset is pressed while the stopwatch is running.
    else:
        label['text'] = 'Starting...'


def open_stopwatch():
    root = Tk()
    root.title("Stopwatch")

    # Fixing the window size.
    root.minsize(width=250, height=70)
    label = Label(root, text="Welcome!", fg="black", font="Verdana 30 bold")
    label.pack()
    f = Frame(root)
    global start, stop, reset
    start = Button(f, text='Start', width=6, command=lambda: Start(label))
    stop = Button(f, text='Stop', width=6, state='disabled', command=Stop)
    reset = Button(f, text='Reset', width=6, state='disabled', command=lambda: Reset(label))
    f.pack(anchor='center', pady=5)
    start.pack(side="left")
    stop.pack(side="left")
    reset.pack(side="left")
    root.mainloop()


###############################################################################################################
##################### Functions #########################################################

def openfile():
    file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '*.')])
    if file:
        global filename
        filename = os.path.abspath(file.name)
    os.system(filename)


def input(cell, i):  # input into spreadsheet
    file = "spreadsheet.xlsx"
    workbook = load_workbook(filename=file)
    # open workbook
    sheet = workbook.active
    # modify the desired cell
    sheet[str(cell)] = str(i)
    # save the file
    workbook.save(filename=file)


def browse():
    file = filedialog.askopenfile(mode='r', filetypes=[('Excel Files', '.*')])
    if file:
        global filename
        filename = os.path.abspath(file.name)


def get_data():
    E1 = str(int(p1_1.get()) + reverse(p1_6.get()))
    E2 = str(int(p2_1.get()) + reverse(p2_6.get()))
    print("E1: " + p1_1.get() + "+" + str(reverse(p1_6.get())) + "=" + E1)
    print("E2: " + p2_1.get() + "+" + str(reverse(p2_6.get())) + "=" + E2)

    A1 = str(int(p1_7.get()) + reverse(p1_2.get()))
    A2 = str(int(p2_7.get()) + reverse(p2_2.get()))
    print("E1: " + p1_7.get() + "+" + str(reverse(p1_2.get())) + "=" + A1)
    print("E2: " + p2_7.get() + "+" + str(reverse(p2_2.get())) + "=" + A2)

    C1 = str(int(p1_3.get()) + reverse(p1_8.get()))
    C2 = str(int(p2_3.get()) + reverse(p2_8.get()))
    print("E1: " + p1_3.get() + "+" + str(reverse(p1_8.get())) + "=" + C1)
    print("E2: " + p2_3.get() + "+" + str(reverse(p2_8.get())) + "=" + C2)

    N1 = str(int(p1_4.get()) + reverse(p1_9.get()))
    N2 = str(int(p2_4.get()) + reverse(p2_9.get()))
    print("E1: " + p1_4.get() + "+" + str(reverse(p1_9.get())) + "=" + N1)
    print("E2: " + p2_4.get() + "+" + str(reverse(p2_9.get())) + "=" + N2)

    O1 = str(int(p1_5.get()) + reverse(p1_10.get()))
    O2 = str(int(p2_5.get()) + reverse(p2_10.get()))
    print("E1: " + p1_5.get() + "+" + str(reverse(p1_10.get())) + "=" + N1)
    print("E2: " + p2_5.get() + "+" + str(reverse(p2_10.get())) + "=" + N2)

    #  Order: Time, Trial Number, Time, P1_E, P1_A,	P1_C, P1_N, P1_O, P2_E, P2_A, P2_C, P2_N, P2_O, P1_ACQ, P2_ACQ,
    #  P1_SAT, P2_SAT, P1_AGE, P2_AGE, P1_GPA, P2_GPA, P1_RACE, P2_RACE, P1_GENDER, P2_GENDER

    data = [trial_num_entry.get(), time.get(), E1, A1, C1, N1, O1, E2, A2, C2, N2, O2, p1_acq.get(), p2_acq.get(),
            p1_sat.get(), p2_sat.get(), p1_age.get(), p2_age.get(), p1_gpa.get(), p2_gpa.get(), p1_race.get(),
            p2_race.get(), p1_gender.get(), p2_gender.get()]

    print(data)

    return data

def set_row(trial_num):
    cells = ["A", "B", 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M',  'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U',
             'V', 'W', 'X']
    for i in range(len(cells)):
        cells[i] = str(cells[i]) + str(int(trial_num) + 1)
    return cells


def reverse(n):
    if n == "7":
        return 1
    elif n == "6":
        return 2
    elif n == "5":
        return 3
    elif n == "4":
        return 4
    elif n == "3":
        return 5
    elif n == "2":
        return 6
    elif n == "1":
        return 7


def insert_data(trial):
    cells = set_row(trial)  # adds the row to the cells according to the trial number
    d = get_data()
    for i in range(len(d)):
        input(cells[i], d[i])

###############################Tk Main ##################################################
win = Tk()
win.title("Data Entry Tool")

global filename

genders = [
    "Male",
    "Female",
    "Other",
    "Prefer not to Answer"
]  # etc

races = [
    "Black or African American",
    "American Indian",
    "Middle Eastern or North African",
    "White",
    "Hispanic, Latino, or Spanish origin",
    "Native Hawaiian or Other Pacific Islander",
    "Asian",
    "Other"
]  # etc

p1_gender = StringVar(win)
p1_gender.set(genders[0])

p2_gender = StringVar(win)
p2_gender.set(genders[0])

p1_race = StringVar(win)
p1_race.set(races[0])

p2_race = StringVar(win)
p2_race.set(races[0])

trial_num_entry = Entry(win, font='Georgia 10')
trial_num_entry.insert(0, "Trial Number")
trial_num_entry.grid(row=4, column=0, padx=10, pady=5)

time = Entry(win, font='Georgia 10')
time.insert(0, "Time")
time.grid(row=5, column=0, padx=10, pady=5)

input_button = Button(win, text="Insert Data", bg="red", fg="white", font='Georgia 13',
                      command=lambda: insert_data(trial_num_entry.get()))
input_button.grid(row=6, column=0, padx=10, pady=5)

stopwatch = Button(win, text="Open Stopwatch", bg="green", fg="white", font='Georgia 13', command=open_stopwatch)
stopwatch.grid(row=7, column=0, padx=10, pady=5)

# browse = ttk.Button(win, text="Browse", command=browse)
# browse.grid(row = 0, column = 3)

# ttk.Button(win, text="Edit File", command=openfile).pack()#row = 0, column = 0, sticky = W, pady = 2)


ypadding = 2

###########################participant 1
p1 = Label(win, text="Participant 1:", font='Georgia 13')
p1.grid(column=1, row=0, pady=ypadding)

p1_1 = Entry(win)
p1_1.insert(0, "E1")
p1_1.grid(column=1, row=1, pady=ypadding)

p1_2 = Entry(win)
p1_2.insert(0, "A2R")
p1_2.grid(column=1, row=2, pady=ypadding)

p1_3 = Entry(win)
p1_3.insert(0, "C3")
p1_3.grid(column=1, row=3, pady=ypadding)

p1_4 = Entry(win)
p1_4.insert(0, "N4")
p1_4.grid(column=1, row=4, pady=ypadding)

p1_5 = Entry(win)
p1_5.insert(0, "05")
p1_5.grid(column=1, row=5, pady=ypadding)

p1_6 = Entry(win)
p1_6.insert(0, "E6R")
p1_6.grid(column=2, row=1, pady=ypadding)

p1_7 = Entry(win)
p1_7.insert(0, "A7")
p1_7.grid(column=2, row=2, pady=ypadding)

p1_8 = Entry(win)
p1_8.insert(0, "C8R")
p1_8.grid(column=2, row=3, pady=ypadding)

p1_9 = Entry(win)
p1_9.insert(0, "N9R")
p1_9.grid(column=2, row=4, pady=ypadding)

p1_10 = Entry(win)
p1_10.insert(0, "010R")
p1_10.grid(column=2, row=5, pady=ypadding)

p1_acq = Entry(win)
p1_acq.insert(0, "Acquaintanceship")
p1_acq.grid(column=3, row=2, pady=ypadding, padx=5)

p1_sat = Entry(win)
p1_sat.insert(0, "Satisfaction")
p1_sat.grid(column=3, row=4, pady=ypadding, padx=5)

p1_age = Entry(win)
p1_age.insert(0, "Age")
p1_age.grid(column=4, row=2, pady=ypadding)

p1_gpa = Entry(win)
p1_gpa.insert(0, "GPA")
p1_gpa.grid(column=5, row=2, pady=ypadding)

p1_gender_OM = OptionMenu(win, p1_gender, *genders)  # OM = OptionMenu
p1_gender_OM.grid(column=5, row=4, pady=ypadding)  # OM = OptionMenu

p1_race_OM = OptionMenu(win, p1_race, *races)
p1_race_OM.grid(column=4, row=4, pady=ypadding)

#################@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@@######################participant 2
p2 = Label(win, text="Participant 2:", font='Georgia 13')
p2.grid(column=1, row=6, pady=ypadding)

p2_1 = Entry(win)
p2_1.insert(0, "E1")
p2_1.grid(column=1, row=7, pady=ypadding)

p2_2 = Entry(win)
p2_2.insert(0, "A2R")
p2_2.grid(column=1, row=8, pady=ypadding)

p2_3 = Entry(win)
p2_3.insert(0, "C3")
p2_3.grid(column=1, row=9, pady=ypadding)

p2_4 = Entry(win)
p2_4.insert(0, "N4")
p2_4.grid(column=1, row=10, pady=ypadding)

p2_5 = Entry(win)
p2_5.insert(0, "05")
p2_5.grid(column=1, row=11, pady=ypadding)

p2_6 = Entry(win)
p2_6.insert(0, "E6R")
p2_6.grid(column=2, row=7, pady=ypadding)

p2_7 = Entry(win)
p2_7.insert(0, "A7")
p2_7.grid(column=2, row=8, pady=ypadding)

p2_8 = Entry(win)
p2_8.insert(0, "C8R")
p2_8.grid(column=2, row=9, pady=ypadding)

p2_9 = Entry(win)
p2_9.insert(0, "N9R")
p2_9.grid(column=2, row=10, pady=ypadding)

p2_10 = Entry(win)
p2_10.insert(0, "010R")
p2_10.grid(column=2, row=11, pady=ypadding)

p2_acq = Entry(win)
p2_acq.insert(0, "Acquaintanceship")
p2_acq.grid(column=3, row=8, pady=ypadding)

p2_sat = Entry(win)
p2_sat.insert(0, "Satisfaction")
p2_sat.grid(column=3, row=10, pady=ypadding)

p2_age = Entry(win)
p2_age.insert(0, "Age")
p2_age.grid(column=4, row=8, pady=ypadding)

p2_gpa = Entry(win)
p2_gpa.insert(0, "GPA")
p2_gpa.grid(column=5, row=8, pady=ypadding)

p2_gender_OM = OptionMenu(win, p2_gender, *genders)
p2_gender_OM.grid(column=5, row=10, pady=ypadding)

p2_race_OM = OptionMenu(win, p2_race, *races)
p2_race_OM.grid(column=4, row=10, pady=ypadding)

win.mainloop()
