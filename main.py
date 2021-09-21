import xlrd, pyautogui, time, datetime, os

time_start = ["09:00", "10:05", "11:10", "12:15", "13:20", "14:25", "15:30"]
time_end = ["09:45", "10:50", "11:55", "13:00", "14:05", "15:10", "16:15"]

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


def timelist():
    ...
