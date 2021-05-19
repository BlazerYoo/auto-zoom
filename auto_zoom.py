import pandas as pd, webbrowser as wb, time, pyautogui as pag, ctypes
from datetime import datetime as dt


# a or b day
day = 'c'


# spreadsheet with zoom info
df = pd.DataFrame(pd.read_excel('zoom.xlsx'))


# print schedule for the day
print('_________________________________________')
if day.lower() == 'a':
    print(df.iloc[:4])
elif day.lower() == 'b':
    print(df.iloc[4:])
else:
    print(df.iloc[0:])
print('_________________________________________')


# start end class
def check_start_class(period):
    # check if class started and check time down to minute
    global class_started
    if not(class_started) and dt.now().strftime('%H:%M:%S')[:5] == df.iloc[period]['start_time_24'][:5]:
        print('Starting Zoom for period', df.iloc[period]['period'])
        zoom_link = df.iloc[period]['meeting_link']
        print(zoom_link)
        wb.open(zoom_link)
        # click on 'Open Zoom Meetings'
        open_zoom_meeting = pag.locateCenterOnScreen('./images/open_zoom_meetings.png', confidence=0.8)
        pause = 0
        if open_zoom_meeting == None:
            while open_zoom_meeting == None:
                time.sleep(1)
                pause += 1
                if pause == 5:
                    break
                open_zoom_meeting = pag.locateCenterOnScreen('./images/open_zoom_meetings.png', confidence=0.8)
                if open_zoom_meeting != None:
                    pag.click(x=open_zoom_meeting[0], y=open_zoom_meeting[1])
                    print()
                    break
                print('\'Open Zoom Meetings\' button not found yet. Trying again for another',5-pause,'seconds...' )
        else:
            pag.click(x=open_zoom_meeting[0], y=open_zoom_meeting[1])
        class_started = True

class_started = False

if day.lower() == 'a':
    # run while time is not yet 3 pm
    while int(dt.now().strftime('%H:%M:%S')[:2]) < 15:
        for period in range(4):
            check_start_class(period)
elif day.lower() == 'b':
    # run while time is not yet 3 pm
    while int(dt.now().strftime('%H:%M:%S')[:2]) < 15:
        for period in range(4,8):
            check_start_class(period)
else:
    # run while time is not yet 3 pm
    while int(dt.now().strftime('%H:%M:%S')[:2]) < 15:
        for period in range(9):
            check_start_class(period)