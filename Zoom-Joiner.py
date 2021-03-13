import webbrowser
from time import sleep
from datetime import datetime
import keyboard
import win32gui


weekdays = ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday")


# helper function for find_starting_with
def handler(handle, window_list):
    name = win32gui.GetWindowText(handle)
    if str(name) != "":                                                  # populate the list
        window_list.append((handle, name))


# finds a window, which name starts with the given String
def find_starting_with(window_name):
    window_list = []
    win32gui.EnumWindows(handler, window_list)                      # process all windows by the handler

    for w in window_list:
        if str(w[1]).startswith(window_name):                       # found, matching start
            return w[0]

    return 0                                                        # not found


# function for setting focus on a window
def focus_on(name):
    count = 0
    while True:                                                     # search until found
        count += 1
        w = find_starting_with(name)
        keyboard.press('alt')
        sleep(0.2)
        keyboard.press_and_release('tab')
        keyboard.release('alt')
        if w != 0:
            try:
                win32gui.ShowWindow(w, 6)
                win32gui.ShowWindow(w, 9)
                win32gui.SetForegroundWindow(w)
                return 1
            except Exception as e:
                print(e)
                sleep(0.5)
        if count > 20:
            print("ERROR: Couldn't find the browser tab.")
            return 0
        sleep(0.5)


# function to add lectures to the file
def add_lecture():
    global weekdays
    while True:
        join_name = input("Name of the lecture/meeting:\n>")
        if len(join_name) > 30:
            print("Name is too long.\n")
            continue
        join_time = input("Time (\"hh:mm\", 24h clock) to join the zoom call:\n>")
        join_date = input(f"Select weekday: ({weekdays[0]}:0, {weekdays[1]}:1, {weekdays[2]}:2, {weekdays[3]}:3, {weekdays[4]}:4, {weekdays[5]}:5, {weekdays[6]}:6)\n>")
        join_link = input("Paste the Zoom link here:\n>")
        file = open("./data.txt", "a")
        file.write(f"{join_name},{join_time},{join_date},{join_link}\n")
        file.close()
        print(f"Added zoom call \"{join_name}\".")
        more = int(input("\nAdd more zoom calls? (1/0)\n>"))
        if more == 0:
            return


# lists data and returns a list with the lines
def list_data():
    f = open("./data.txt", "r")
    lines = f.readlines()
    f.close()
    i = 0
    for l in lines:                                                             # display line by line
        split_l = l.split(',')
        padding = ""
        for j in range(30 - len(split_l[0])):                                   # add padding
            padding += " "
        print(f"{i}:\t{split_l[0] + padding}{split_l[1]}\t{weekdays[int(split_l[2])]}")
        i += 1
    return lines


# function to remove calls
def remove_calls():
    global weekdays
    while True:
        lines = list_data()
        choice = int(input("\nChoose a call to delete:\n>"))                    # choose line to be removed
        if choice > len(lines) - 1 or choice < 0:
            print("This choice is not in range.")
            continue
        f = open("./data.txt", "w")
        for i in range(len(lines)):                                             # write back to file
            if i == choice:
                continue
            f.write(lines[i])
        f.close()
        print(f"Removed zoom call.")
        more = int(input("\nRemove more zoom calls? (1/0)\n>"))
        if more == 0:
            return


# function returning the next upcoming call as a list with the properties of that call
def next_call():
    file = open("./data.txt", "r")
    calls = file.readlines()
    file.close()

    current_time = datetime.now().strftime("%H:%M")
    current_day = int(datetime.today().weekday())

    calls_today = list()
    for c in calls:                                             # sort out calls that are coming up today
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
        if int(c[1][:2]) < int(next_c[1][:2]):           # earlier hour
            next_c = c
        elif int(c[1][:2]) == int(next_c[1][:2]) and \
                int(c[1][3:]) < int(next_c[1][3:]):        # same hour, earlier min
            next_c = c

    return next_c


# calculates how long to sleep until the next meeting starts
def to_sleep(next_time):
    now = datetime.now().strftime("%H:%M:%S")
    h_to_sleep = int(next_time[:2]) - int(now[:2])
    if int(next_time[3:]) >= int(now[3:5]):
        m_to_sleep = int(next_time[3:]) - int(now[3:5])
    else:
        m_to_sleep = int(next_time[3:]) + 60 - int(now[3:5])
    s_to_sleep = 60 - int(now[6:])
    return s_to_sleep + 60 * (60 * h_to_sleep + m_to_sleep - 1)  # -1 because of the s_to_sleep


# function for joining the zoom calls
def join_calls():
    while True:
        next = next_call()
        if next is None:
            return

        print(f"Next call is:\t{next[0]} at {next[1]}")

        # sleep until then in intervals in case the program got interrupted
        while True:
            sl = to_sleep(next[1])
            if sl > 15:
                sleep(15)
            else:
                if sl < 1:
                    break
                sleep(sl)
                break

        print("Joining... (Intervention with keyboard/mouse can lead to problems.)\n\a")
        wb = webbrowser.get()
        wb.open_new(next[3])
        sleep(2)
        res = focus_on("Launch Meeting - Zoom")
        if res == 0:
            sleep(61)
            continue
        keyboard.press_and_release('tab')                                       # accept to open in zoom
        keyboard.press_and_release('tab')
        keyboard.press_and_release('enter')
        sleep(61)                                                               # so the same meeting isn't joined again

        
print("---------------------------- Zoom Joiner ----------------------------\n")

# start the program
while True:
    start = str(input("ENTER : start the program\na : add zoom calls\tr : remove zoom calls\tl : list zoom calls\n>")).lower()
    if start == "a":
        add_lecture()
        print("\n")
    elif start == "r":
        remove_calls()
        print("\n")
    elif start == "l":
        list_data()
        print("")
    elif start == "":
        join_calls()
        break
