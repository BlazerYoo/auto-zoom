import pandas as pd, webbrowser as wb, time, pyautogui as pag, ctypes
from datetime import datetime as dt


# a or b day
day = 'a'


# spreadsheet with zoom info
df = pd.DataFrame(pd.read_excel("zoom.xlsx"))


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
        print('Starting Zoom for period', str(df.iloc[period]['period']) + '...')
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
                    print('Could not open period', df.iloc[period]['period'], 'Zoom meeting')
                    break
                open_zoom_meeting = pag.locateCenterOnScreen('./images/open_zoom_meetings.png', confidence=0.8)
                if open_zoom_meeting != None:
                    pag.click(x=open_zoom_meeting[0], y=open_zoom_meeting[1])
                    print('Successfully opened period', df.iloc[period]['period'], 'Zoom meeting')
                    class_started = True
                    break
                print('\'Open Zoom Meetings\' button not found yet. Trying again for another', 5-pause, 'seconds...' )
        else:
            pag.click(x=open_zoom_meeting[0], y=open_zoom_meeting[1])
            print('Successfully opened period', df.iloc[period]['period'], 'Zoom meeting')
            class_started = True
        # maximize zoom window
        #fg_win = user32.GetForegroundWindow()
        #user32.ShowWindow(fg_win, SW_MAXIMISE)





def check_end_class(period):
    # check if class started and check time down to minute
    global class_started
    if class_started and dt.now().strftime('%H:%M:%S')[:5] == df.iloc[period]['end_time_24'][:5]:
        print('Ending Zoom for period', df.iloc[period]['period'])
        # click on 'Leave Room'
        leave_room = pag.locateCenterOnScreen('./images/leave_room.png', confidence=0.8)
        pause = 0
        if leave_room == None:
            while leave_room == None:
                time.sleep(1)
                pause += 1
                if pause == 3:
                    print('Could not leave breakout room pt1 in period', df.iloc[period]['period'], 'Zoom meeting')
                    break
                leave_room = pag.locateCenterOnScreen('./images/leave_room.png', confidence=0.8)
                if leave_room != None:
                    pag.click(x=leave_room[0], y=leave_room[1])
                    print('Successfully left breakout room pt1 in period', df.iloc[period]['period'], 'Zoom meeting')
                    break
                print('\'Leave Room\' button not found yet. Trying again for another', 3-pause, 'seconds...' )
        else:
            pag.click(x=leave_room[0], y=leave_room[1])
            print('Successfully left breakout room pt1 in period', df.iloc[period]['period'], 'Zoom meeting')

        # click on 'Leave Breakout Room'
        leave_breakout_room = pag.locateCenterOnScreen('./images/leave_breakout_room.png', confidence=0.8)
        pause = 0
        if leave_breakout_room == None:
            while leave_breakout_room == None:
                time.sleep(1)
                pause += 1
                if pause == 3:
                    print('Could not leave breakout room in period', df.iloc[period]['period'], 'Zoom meeting')
                    break
                leave_breakout_room = pag.locateCenterOnScreen('./images/leave_breakout_room.png', confidence=0.8)
                if leave_breakout_room != None:
                    pag.click(x=leave_breakout_room[0], y=leave_breakout_room[1])
                    print('Successfully left breakout room in period', df.iloc[period]['period'], 'Zoom meeting')
                    break
                print('\'Leave Breakout Room\' button not found yet. Trying again for another', 3-pause, 'seconds...' )
        else:
            pag.click(x=leave_breakout_room[0], y=leave_breakout_room[1])
            print('Successfully left breakout room in period', df.iloc[period]['period'], 'Zoom meeting')

        # click on 'Leave'
        leave = pag.locateCenterOnScreen('./images/leave.png', confidence=0.8)
        pause = 0
        if leave == None:
            while leave == None:
                time.sleep(1)
                pause += 1
                if pause == 3:
                    print('Could not leave pt1 in period', df.iloc[period]['period'], 'Zoom meeting')
                    break
                leave = pag.locateCenterOnScreen('./images/leave.png', confidence=0.8)
                if leave != None:
                    pag.click(x=leave[0], y=leave[1])
                    print('Successfully left pt1 in period', df.iloc[period]['period'], 'Zoom meeting')
                    break
                print('\'Leave\' button not found yet. Trying again for another', 3-pause, 'seconds...' )
        else:
            pag.click(x=leave[0], y=leave[1])
            print('Successfully left pt1 in period', df.iloc[period]['period'], 'Zoom meeting')

        # click on 'Leave Meeting'
        leave_meeting = pag.locateCenterOnScreen('./images/leave_meeting.png', confidence=0.8)
        pause = 0
        if leave_meeting == None:
            while leave_meeting == None:
                time.sleep(1)
                pause += 1
                if pause == 3:
                    print('Could not leave period', df.iloc[period]['period'], 'Zoom meeting')
                    break
                leave_meeting = pag.locateCenterOnScreen('./images/leave_meeting.png', confidence=0.8)
                if leave_meeting != None:
                    pag.click(x=leave_meeting[0], y=leave_meeting[1])
                    print('Successfully left period', df.iloc[period]['period'], 'Zoom meeting')
                    class_started = False
                    break
                print('\'Leave Meeting\' button not found yet. Trying again for another', 3-pause, 'seconds...' )
        else:
            pag.click(x=leave_meeting[0], y=leave_meeting[1])
            print('Successfully left period', df.iloc[period]['period'], 'Zoom meeting')
            class_started = False





class_started = False

if day.lower() == 'a':
    # run while time is not yet 3 pm
    while int(dt.now().strftime('%H:%M:%S')[:2]) < 15:
        for period in range(4):
            check_start_class(period)
            check_end_class(period)
elif day.lower() == 'b':
    # run while time is not yet 3 pm
    while int(dt.now().strftime('%H:%M:%S')[:2]) < 15:
        for period in range(4,8):
            check_start_class(period)
            check_end_class(period)
else:
    # run while time is not yet 3 pm
    while int(dt.now().strftime('%H:%M:%S')[:2]) < 15:
        for period in range(9):
            check_start_class(period)
            check_end_class(period)