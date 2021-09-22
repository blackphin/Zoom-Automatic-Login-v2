import xlrd, time, datetime, os, csv
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
        auto.hotkey("ctrl", "win", "d")
    time.sleep(0.5)
    auto.press("win")
    time.sleep(0.5)
    auto.write(message="start zoom")
    time.sleep(0.5)
    auto.press("enter")
    time.sleep(5)
    auto.hotkey("win", "up")


def auto_join(id="977 0864 8127", passw="261893"):
    auto.click(x=776, y=433)
    time.sleep(0.5)
    auto.click(x=769, y=466)
    auto.write(message=id)
    auto.press("enter")
    time.sleep(2)
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


def wait_till(timea):
    today = datetime.date.today()
    dt = datetime.datetime.strptime(timea, "%H:%M")
    when = datetime.datetime(*today.timetuple()[:3], *dt.timetuple()[3:6])
    wait_time = (when - datetime.datetime.now()).total_seconds()
    if wait_time < 0:
        print(f"Time {when} has already passed")
    else:
        print(f"Waiting {wait_time} seconds until {when}")
        time.sleep(wait_time)


time_start = []
time_end = []


def time_list():
    for x in range(sheet_timetable.ncols - 1):
        time_range = sheet_timetable.cell_value(0, x + 1)
        time_start.append(time_range[0:5])
        time_end.append(time_range[6:11])


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


# __main__
time_list()
print(time_start)
print(time_end)

subjects_today()
print(subject_list)
print(day_index)

mode()
print(subject_mode)

credentials()
print(credentials_list)
