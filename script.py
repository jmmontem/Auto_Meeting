import pyautogui as agui, datetime, time, subprocess, csv, os, webbrowser, openpyxl


def manualJoin(meeting: list):
    ''' Open Manual Zoom Meeting '''
    subfolders = [ f.path for f in os.scandir("C:\\Users") if f.is_dir() ]

    ## Open Zoom if it's not open or it's hidden
    for i in subfolders:
        if os.path.isfile(i + "\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe"):
            subprocess.Popen(i + "\\AppData\\Roaming\\Zoom\\bin\\Zoom.exe")

    time.sleep(2)
    while True:
        cords = agui.locateCenterOnScreen('img/manual_join.png')
        if cords != None:
            agui.click(cords)
            break;
        else:
            print("Could not fint the Manual Join Picture...")
            
    time.sleep(2)
    agui.typewrite(meeting[0])
    cords = agui.locateCenterOnScreen("img/videooff.png")
    agui.click(cords)
    time.sleep(1)
    cords = agui.locateCenterOnScreen("img/join.png")
    agui.click(cords)

    if len(meeting) == 2:
        time.sleep(1)
        agui.typewrite(meeting[1])
        cords = agui.locateCenterOnScreen("img/joinmeeting.png")
        agui.click(cords)

    return
    
def linkJoin(link):
    webbrowser.open(link)
    start = time.time()
    time.sleep(1)
    while True:
        cords = agui.locateCenterOnScreen('img/link_join.png')
        if cords != None:
            print("Found the button...")
            agui.click(cords)
            break;
        else:
            print("Could not fint the Link Join Picture...")
            if (time.time() - start >= 120):
                print("Didn't click anything... breaking off")
                break;
            
    return


def fillMeetings():
    result = []
    wb = openpyxl.load_workbook('List.xlsx')
    sheet = wb['Sheet1']
    for i in sheet.iter_rows(values_only = True):
        if i[0] != None:   
            result.append(i)
    return sorted(result[1:])

def run_script(meetings: list):
    for i in range(len(meetings)):
        cur_meeting = meetings[i]

        cur = round(time.time(), 0)
        temp = cur_meeting[0].timestamp()
        print(cur, temp, cur - temp)

        if (cur < temp - 60):
            print("next class in ", end ="")
            print(datetime.timedelta(seconds = (temp - cur) - 60))
            print("current time: ", end="")
            print(time.asctime( time.localtime(time.time())))
            time.sleep(temp - cur - 60)
        elif (cur - temp) > 600:
            print("skipped meeting " + str(i + 1))
            continue

        var = os.system("taskkill /f /im Zoom.exe")

        print("Meeting at this time will start: " + str(cur_meeting[0]))

        if cur_meeting[1] != None:
            linkJoin(str(cur_meeting[1]))

        if cur_meeting[2] != None:
            meeting_l = []
            if cur_meeting[3] != None:
                meeting_l = [str(cur_meeting[2]), str(cur_meeting[3])]
            else:
                meeting_l = [str(cur_meeting[2])]
                
            manualJoin(meeting_l)

        time.sleep(5)


    print("Done, press anything to exit")
    input()
    var = os.system("taskkill /f /im Zoom.exe")

        
if __name__ == '__main__':
    meetings = fillMeetings()
    run_script(meetings)

