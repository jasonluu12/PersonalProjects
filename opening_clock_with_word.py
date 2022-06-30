# Reference for using WMI to get a list currently running services (Method 2 using OS)
#https://www.geeksforgeeks.org/python-get-list-of-running-processes/
# Reference for getting elapsed time (Using DateTime module)
# https://www.geeksforgeeks.org/how-to-measure-elapsed-time-in-python/
# Reference for the GUI
# PySimpleGUI.readthedocs.io
# Reference for https://docs.python.org/3/library/datetime.html#timedelta-objects

from tabnanny import check
from time import strftime
import wmi
import PySimpleGUI as sg
from datetime import datetime
import sys
import os
word_process_name = 'WINWORD.EXE'

def main():
    start_time = check_word()
    create_window(start_time)

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
    # while True:
    #     f = wmi.WMI()
    #     for process in f.Win32_Process():
    #         # Check if Microsoft Word has started
    #         if process.Name == word_process_name:
    #             print("Microsoft Word is running")
    #             is_word_running = True

    #     if is_word_running:
    #         start_time_word = datetime.now()
    #         break
    return start_time_word

# Makes a window using PySimpleGUI that shows the clock and elapsed time since Word was started
# The window is Always On Top and Asynchronous (no user input needed)
def create_window(start_time):
    #sg.theme('#211AE2')   # Add a little color to your windows
    # All the stuff inside your window. This is the PSG magic code compactor...
    layout = [  [sg.Text('Time:', font='Arial, 17', text_color='#3A3ABC', background_color = '#A0AE73',key='-OUTPUT_TIME-')],
                [sg.Text('Elapsed Time:', font='Arial 17',text_color='#3A3ABC',background_color = '#A0AE73',key='-ELAPSED_TIME-')],
                [sg.OK(), sg.Cancel()]
            ]

    # Create the Window
    window = sg.Window('Current Time', layout, keep_on_top = True, grab_anywhere=True,location=(500,50),background_color = '#A0AE73')
    # Event Loop to process "events"
    while True:             
        event, values = window.read(timeout=100)
        if event in (sg.WIN_CLOSED, 'Cancel'):
           break

        check_word_stop()

        current_time = datetime.now()
        cur_time = combine_strings('Time:', str(current_time.strftime("%H:%M")))
        window['-OUTPUT_TIME-'].update(cur_time)
        elapsed_time = str(current_time-start_time).split(':')
        elap_time = combine_strings(elapsed_time[0] + ':' + elapsed_time[1], 'has passed (in H:M)')
        window['-ELAPSED_TIME-'].update(elap_time)

    window.close()

# Combines the time with the context
def combine_strings(str1,str2):
    return str1 + " "+ str2

# Checks to see if user quits Word then quits the program
def check_word_stop():
    output = os.popen('wmic process get description, processid').read()

    if word_process_name not in output:
        sys.exit()


if __name__ == '__main__':
    main()
