# Reference for using WMI to get a list currently running services
#https://www.geeksforgeeks.org/python-get-list-of-running-processes/
# Reference for getting elapsed time
# https://www.geeksforgeeks.org/how-to-measure-elapsed-time-in-python/
# Reference for the GUI
# PySimpleGUI
from tabnanny import check
from time import strftime, time
import wmi
import PySimpleGUI as sg
from datetime import datetime
import sys
import os
from win10toast import ToastNotifier
from playsound import playsound

word_process_name = 'WINWORD.EXE'
time_to_show_notification = 30

def main():
    start_time = check_word()
    toaster = ToastNotifier()
    create_window(start_time,toaster)

# Checks for Microsoft Word's status
# If Word is running quit the loop
# Then Keep track of the time that Word started
def check_word():
    # Initializing the wmi constructor
    is_word_running = False
    print("Currently checking for Microsoft Word")
    # Iterating through all the running processes
    while True:
        output = os.popen('wmic process get description, processid').read()
        if word_process_name in output:
            print("Microsoft Word is running")
            break
    start_time_word = datetime.now()
    
    return start_time_word

# Makes a window using PySimpleGUI that shows the clock and elapsed time since Word was started
# The window is Always On Top and Asynchronous (no user input needed)
def create_window(start_time,toaster):
    
    #sg.theme('#211AE2')   # Add a little color to your windows
    # All the stuff inside your window. This is the PSG magic code compactor...
    layout = [  [sg.Text('Time:', font='Arial, 17', text_color='#3A3ABC', background_color = '#A0AE73',key='-OUTPUT_TIME-')],
                [sg.Text('Elapsed Time:', font='Arial 17',text_color='#3A3ABC',background_color = '#A0AE73',key='-ELAPSED_TIME-')],
                [sg.OK(), sg.Cancel()]
            ]

    # Create the Window
    window = sg.Window('Current Time', layout, keep_on_top = True, grab_anywhere=True,location=(550,50),background_color = '#A0AE73')
    # Event Loop to process "events"
    while True:             
        event, values = window.read(timeout=100)
        if event in (sg.WIN_CLOSED, 'Cancel'):
           break

        update_time(window, start_time, toaster)

        check_word_stop()

        

    window.close()

# Combines the time with the context
def combine_strings(str1,str2):
    return str1 + " "+ str2

# Checks to see if user quits Word then quits the program
def check_word_stop():
    output = os.popen('wmic process get description, processid').read()

    if word_process_name not in output:
        sys.exit()

# Continously updates time to the popup window
# Update elapsed time as well
def update_time(window, start_time,toaster):
    current_time = datetime.now()
    cur_time = combine_strings('Time:', str(current_time.strftime("%H:%M")))
    window['-OUTPUT_TIME-'].update(cur_time)
    elapsed_time = str(current_time-start_time).split(':')
    print(elapsed_time)
    elap_time = combine_strings(elapsed_time[0] + ':' + elapsed_time[1], 'has passed (in H:M)')
    show_notification(*(elapsed_time[1],elapsed_time[2]), toaster)
    window['-ELAPSED_TIME-'].update(elap_time)

# Uses Win10Toast to show a notification with ## minutes has passed
# I put it to 30 minutes
def show_notification(time_mins, time_ms, toaster):
    time_mins = int(time_mins)
    time_ms = float(time_ms)
    if time_mins >= time_to_show_notification and time_mins % time_to_show_notification == 0:
        if  1>= time_ms >= 0:
            toaster.show_toast("Notification","{} minutes has passed".format(time_mins),duration=5,threaded=True)


if __name__ == '__main__':
    main()
