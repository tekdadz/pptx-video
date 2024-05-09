import win32com.client
import time

powerpoint = win32com.client.Dispatch("Powerpoint.Application")
try:
    # Attempt to open file
    presentation = powerpoint.Presentations.Open(FileName=r'E:\python\pptx-video\test.pptx')
except:
    # If file cannot be found
    print('File cannot be found')
    exit

try:
    # May need a few other parameters as well
    presentation.CreateVideo(r'E:\python\pptx-video\test.mp4')
    while presentation.CreateVideoStatus == 1:
        time.sleep(1)
    presentation.Close()
    print('Done')
except:
    print('Unable to export to video')
