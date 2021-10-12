import xlrd, time, datetime, os
import pyautogui as auto

time_slots = 8


# __precode__
wb1 = xlrd.open_workbook(
    r"D:\OneDrive\Repositories\Zoom-Automatic-Login-v2\Credentials.xls"
)
wb2 = xlrd.open_workbook(r"D:\OneDrive\Repositories\Zoom-Automatic-Login-v2\Mode.xls")
wb3 = xlrd.open_workbook(
    r"D:\OneDrive\Repositories\Zoom-Automatic-Login-v2\Timetable.xls"
)
sheet_credential = wb1.sheet_by_index(0)
sheet_mode = wb2.sheet_by_index(0)
sheet_timetable = wb3.sheet_by_index(0)
sheet_credential.cell_value(0, 0)
sheet_mode.cell_value(0, 0)
sheet_timetable.cell_value(0, 0)


# __functions__
def start_zoom(x=0):
    auto.FAILSAFE = False
    if x == 1:
        time.sleep(0.5)
        auto.hotkey("win", "d")
    time.sleep(0.5)
    auto.press("win")
    time.sleep(0.5)
    auto.write(message="start zoom")
    time.sleep(0.5)
    auto.press("enter")
    time.sleep(10)
    auto.hotkey("win", "up")


def auto_join(id, passw):
    time.sleep(1)
    auto.click(x=776, y=433)
    time.sleep(4)
    auto.click(x=769, y=466)
    time.sleep(1)
    auto.write(message=id)
    auto.press("enter")
    time.sleep(4)
    auto.click(x=769, y=466)
    auto.write(message=passw)
    auto.press("enter")


def kill_zoom(x=0):
    os.system("taskkill /IM zoom.exe /T /F")
    auto.moveTo(1410, 1057)
    auto.moveTo(1700, 1057)
    auto.moveTo(1920 / 2, 1080 / 2)
    if x == 1:
        auto.hotkey("ctrl", "win", "f4")


def wait_till(timea, subject):
    today = datetime.date.today()
    dt = datetime.datetime.strptime(timea, "%H:%M")
    when = datetime.datetime(*today.timetuple()[:3], *dt.timetuple()[3:6])
    wait_time = (when - datetime.datetime.now()).total_seconds()
    if wait_time < 0:
        print(f"Time to join the {subject} Lecture: {when} has already passed")
        q = str(input("Do you still want to join this lecture?(Yes/No) >>"))
        if q.lower() == "yes":
            print(f"Joining {subject}'s class")
            return 1
        elif q.lower() == "no":
            return 0
    else:
        print(
            f"Waiting {int(wait_time/60)} minutes until {when} to join "
            + subject
            + " class"
        )
        time.sleep(wait_time)
        print(f"Joining {subject}'s class")
        return 1


time_start = []
time_end = []


def time_list():
    global time_start, time_end
    for x in range(sheet_timetable.ncols - 1):
        time_range = sheet_timetable.cell_value(0, x + 1)
        time_start.append(time_range[0:5])
        time_end.append(time_range[6:11])
    return time_start, time_end


subject_list = []
day_index = 0
day = datetime.datetime.today().strftime("%A")


def subjects_today(today_day=day):
    global subject_list, day_index
    # today_day = datetime.datetime.today().strftime("%A")
    for x in range(sheet_timetable.nrows):
        if today_day == sheet_timetable.cell_value(x, 0):
            list = sheet_timetable.row_values(x)[1:]
            for y in list:
                if list.count(y) != 1 and y != "" and y != "R":
                    list[list.index(y) + 1] = "R"
            subject_list.extend(list)
            day_index = x
            break
    return list


subject_mode = []


def mode():
    global subject_list, subject_mode
    subject_mode = list(subject_list)
    for z in subject_mode:
        if z != "" and z != "R":
            subject_mode[subject_mode.index(z)] = [
                z,
                sheet_mode.cell_value(day_index, subject_mode.index(z) + 1),
            ]
    return subject_mode


credentials_list = []


def credentials():
    global credentials_list, subject_mode
    for v in subject_mode:
        if v != "" and v != "R":
            for i in range(sheet_credential.nrows):
                if v[0] == sheet_credential.cell_value(i, 0):
                    row_number = i
                    if v[1] == 1:
                        column_number = 1
                    elif v[1] == 2:
                        column_number = 3
                    credentials_list.append(
                        [
                            sheet_credential.cell_value(row_number, column_number),
                            str(
                                int(
                                    sheet_credential.cell_value(
                                        row_number, column_number + 1
                                    )
                                )
                            ),
                        ]
                    )
        elif v == "":
            credentials_list.append("")
        elif v == "R":
            credentials_list.append("R")
    return credentials_list


# __main__


def start_sequence():
    time_list()
    subjects_today()
    mode()
    credentials()


days = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def join_meeting():
    for x in range(8):
        if subject_list[x] != "R" and subject_list[x] != "":
            if wait_till(time_start[x], subject_list[x]) == 1:
                kill_zoom()
                start_zoom(1)
                auto_join(credentials_list[x][0], credentials_list[x][1])
        ...


start_sequence()

print("Running " + days[day_index - 1] + "'s Schedule")
print(time_start)
print(subject_list)
print(subject_mode)
print(credentials_list)
join_meeting()
