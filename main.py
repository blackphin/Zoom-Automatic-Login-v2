import xlrd, pyautogui, time, datetime, os, csv

time_slots = 8


# __precode__
wb1 = xlrd.open_workbook(
    r"D:\OneDrive\Repositories\Zoom-Automatic-Login-v2\Credentials.xls"
)
wb2 = xlrd.open_workbook(
    r"D:\OneDrive\Repositories\Zoom-Automatic-Login-v2\Timetable.xls"
)
sheet_credential = wb1.sheet_by_index(0)
sheet_timetable = wb2.sheet_by_index(0)
sheet_credential.cell_value(0, 0)
sheet_timetable.cell_value(0, 0)


# __functions__
def launch():
    ...


def auto_join(id, passw):
    pyautogui.FAILSAFE = False
    pyautogui.hotkey("win", "d")
    time.sleep(0.5)
    pyautogui.press("win")
    time.sleep(0.5)
    pyautogui.write(message="start zoom")
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(1)
    pyautogui.click(x=956, y=573)
    time.sleep(0.5)
    pyautogui.click(x=1155, y=518)
    time.sleep(5)
    pyautogui.hotkey("win", "up")
    time.sleep(0.5)
    pyautogui.click(x=717, y=57)
    time.sleep(0.5)
    pyautogui.click(x=774, y=425)
    time.sleep(1)
    pyautogui.write(message=id)
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(2)
    pyautogui.write(message=passw)
    time.sleep(0.5)
    pyautogui.press("enter")
    time.sleep(20)
    pyautogui.hotkey("win", "up")
    time.sleep(0.5)
    pyautogui.hotkey("win", "d")


time_start = []
time_end = []


def time_list():
    for x in range(sheet_timetable.ncols - 1):
        time_range = sheet_timetable.cell_value(0, x + 1)
        time_start.append(time_range[0:5])
        time_end.append(time_range[6:11])


subject_list = []


def subjects_today():
    today_day = datetime.datetime.today().strftime("%A")
    for x in range(sheet_timetable.nrows):
        if today_day == sheet_timetable.cell_value(x, 0):
            list = sheet_timetable.row_values(x)[1:]
            for x in list:
                if list.count(x) != 1 and x != "":
                    list[list.index(x) + 1] = ""
            subject_list.extend(list)


# __main__
time_list()
print(time_start, time_end)

subjects_today()
print(subject_list)


# datetime.datetime.strptime(time, "%H%M")
