from winshell import desktop
from os import system, path as ospath
from win32gui import GetWindowText, EnumWindows, ShowWindow, SetForegroundWindow
from keyboard import press, release, press_and_release
from webbrowser import get as wbget
from time import sleep
from datetime import datetime
from win32com.client import Dispatch

weekdays = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")

# helper function for find_starting_with
def handler(handle, window_list):
    name = GetWindowText(handle)
    if str(name) != "":
        # populate the list
        window_list.append((handle, name))

# finds a window, which name starts with the given String
def find_starting_with(window_name):
    window_list = []
    # process all windows by the handler
    EnumWindows(handler, window_list)

    for w in window_list:
        if str(w[1]).startswith(window_name):
            # found, matching start
            return w[0]
    # not found
    return 0

# function for setting focus on a window
def focus_on(name):
    count = 0
    # search until found
    while True:
        count += 1
        w = find_starting_with(name)
        press('alt')
        sleep(0.2)
        press_and_release('tab')
        release('alt')
        if w != 0:
            try:
                ShowWindow(w, 6)
                ShowWindow(w, 9)
                sleep(4)
                SetForegroundWindow(w)
                return 1
            except Exception as e:
                print(e)
                sleep(0.5)
        if count > 20:
            print("ERROR: Couldn't find the browser tab.")
            return 0
        sleep(0.5)

# sorts the data file by day / time
def sort_file():
    f = open("./zoom-joiner-data.txt", "r")
    lines = f.readlines()
    f.close()
    lines.sort(key=lambda x: x.split(',')[2] + x.split(',')[1])
    f = open("zoom-joiner-data.txt", "w")
    for l in lines:
        if l != "\n":
            f.write(l)
    f.close()

# function to add lectures to the file
def add_lecture():
    global weekdays
    while True:
        system('cls')
        join_name = input("Name of the lecture/meeting:\n>")
        if len(join_name) > 30:
            print("Name is too long.\n")
            continue
        elif len(join_name) == 0:
            continue
        else:
            break

    while True:
        join_time = input("Time (\"hh:mm\", 24h clock) to join the zoom call:\n>")
        if len(join_time) != 5:
            print("Time has wrong format.\n")
            continue
        break

    while True:
        join_day = int(input(f"Select weekday: ({weekdays[0]}:0, {weekdays[1]}:1, {weekdays[2]}:2, {weekdays[3]}:3, {weekdays[4]}:4, {weekdays[5]}:5, {weekdays[6]}:6)\n>"))
        if not (join_day in range(7)):
            print("Select the day using the given numbers.\n")
            continue
        break

    join_link = input("Paste the Zoom link here:\n>")
    file = open("./zoom-joiner-data.txt", "a")
    file.write(f"{join_name},{join_time},{join_day},{join_link}\n")
    file.close()
    sort_file()
    print(f"Added zoom call \"{join_name}\".")
    return

# lists data and returns a list with the lines
def list_data():
    system('cls')
    f = open("./zoom-joiner-data.txt", "r")
    lines = f.readlines()
    f.close()
    i = 0
    # display line by line
    for l in lines:
        split_l = l.split(',')
        padding = ""
        for j in range(30 - len(split_l[0])):
            # add padding
            padding += " "
        print(f"{i}:\t{split_l[0] + padding}{split_l[1]}\t{weekdays[int(split_l[2])]}")
        i += 1
    return lines

# function to remove calls
def remove_calls():
    global weekdays
    while True:
        system('cls')
        lines = list_data()
        choice = int(input("\nChoose a call to delete:\n>"))
        # choose line to be removed
        if choice > len(lines) - 1 or choice < 0:
            print("This choice is not in range.")
            continue
        else:
            break

    f = open("./zoom-joiner-data.txt", "w")
    for i in range(len(lines)):
        # write back to file
        if i == choice:
            continue
        f.write(lines[i])
    f.close()
    print(f"Removed zoom call.")
    return 

# let user select a call from the list and returns that call as a list of it's properties
def select_call():
    global weekdays
    while True:
        system('cls')
        lines = list_data()
        # choose line to join call
        choice = int(input("\nChoose a call to join:\n>"))
        if choice > len(lines) - 1 or choice < 0:
            print("This choice is not in range.")
            continue
        return lines[choice].split(',')

# function returning the next upcoming call as a list with the properties of that call
def next_call():
    file = open("./zoom-joiner-data.txt", "r")
    calls = file.readlines()
    file.close()

    current_time = datetime.now().strftime("%H:%M")
    current_day = int(datetime.today().weekday())

    calls_today = list()
    # sort out calls that are coming up today
    for c in calls:
        if c == "\n":
            continue
        this_call = c.split(",")
        if int(this_call[2]) == current_day and \
                (int(this_call[1][:2]) > int(current_time[:2]) or
                (int(this_call[1][:2]) == int(current_time[:2]) and
                 int(this_call[1][3:]) >= int(current_time[3:]))):
            calls_today.append(this_call)

    if len(calls_today) == 0:
        print("No more calls today, restart tomorrow.")
        sleep(20)
        return None

    next_c = calls_today[0]
    for c in calls_today:
        if int(c[1][:2]) < int(next_c[1][:2]):
            # earlier hour
            next_c = c
        elif int(c[1][:2]) == int(next_c[1][:2]) and \
                int(c[1][3:]) < int(next_c[1][3:]):
            # same hour, earlier min
            next_c = c

    return next_c

# calculates how long to sleep until the next meeting starts
def to_sleep(next_time):
    now = datetime.now().strftime("%H:%M:%S")
    h_to_sleep = int(next_time[:2]) - int(now[:2])
    if h_to_sleep < 0:
        # late entry, hour missed
        return 0
    if int(next_time[3:]) >= int(now[3:5]):
        m_to_sleep = int(next_time[3:]) - int(now[3:5])
    else:
        if h_to_sleep == 0:
            # late entry, same hour, minutes missed
            return 0
        m_to_sleep = int(next_time[3:]) + 60 - int(now[3:5])
    s_to_sleep = 60 - int(now[6:])
    # -1 because of the s_to_sleep
    return s_to_sleep + 60 * (60 * h_to_sleep + m_to_sleep - 1) 

# function for joining the zoom calls, takes argument: either the exact call to join now, or None
def join_calls(call=None):
    join_now = False
    if call is not None:
        join_now = True
    while True:
        system('cls')
        # join given one the first time
        if join_now:
            next_c = call
        else:
            next_c = next_call()
            if next is None:
                return

        if not join_now:
            print(f"Next call is:\t{next_c[0]} at {next_c[1]}")
            # sleep until then in intervals in case the program got interrupted
            while True:
                sl = to_sleep(next_c[1])
                if sl > 15:
                    sleep(15)
                else:
                    if sl < 1:
                        break
                    sleep(sl)
                    break

        print("Joining... (Intervention with keyboard/mouse can lead to problems.)\n\a")
        wb = wbget()
        wb.open_new(next_c[3])
        res = focus_on("Launch Meeting - Zoom")
        if res == 0:
            sleep(61)
            continue
        # accept to open in zoom
        press_and_release('tab')
        press_and_release('tab')
        press_and_release('enter')
        # no immediate join after the first one
        join_now = False
        # so the same meeting isn't joined again
        sleep(57)

# create desktop link with icon
def link():
    running_dir = ospath.dirname(ospath.abspath(__file__))
    desk = desktop()
    path = ospath.join(desk, "Zoom Joiner.lnk")
    target = running_dir + "\\Zoom-Joiner.py"
    wDir = running_dir
    icon = running_dir + "\\icon.ico"
    shell = Dispatch('WScript.Shell')
    shortcut = shell.CreateShortCut(path)
    shortcut.Targetpath = target
    shortcut.WorkingDirectory = wDir
    shortcut.IconLocation = icon
    shortcut.save()
    print("Shortcut created.\n")

if __name__ == "__main__":
    print("---------------------------- Zoom Joiner ----------------------------\n")

    # start the program
    while True:
        start = str(input("ENTER : start the program\na : add zoom calls\nr : remove zoom calls\nl : list zoom calls\nd : create desktop link\nj : join specific call\n>")).lower()
        if start == "a":
            add_lecture()
            system('cls')
        elif start == "r":
            remove_calls()
            system('cls')
        elif start == "l":
            list_data()
            print("")
        elif start == "d":
            system('cls')
            link()
        elif start == "j":
            join_calls(select_call())
            break
        elif start == "exit":
            break
        elif start == "":
            join_calls()
            break
        else:
            system('cls')
