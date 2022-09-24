import os
import sys
import threading
from queue import Queue

queue = Queue()

import PySimpleGUI as sg
import ctypes

from mivisor.utils import load_excel_data
from mivisor.windows import create_data_table, create_annotate_column_window

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(True)
except:
    pass


def load_excel_data_thread(filepath, window, queue):
    load_excel_data(filepath, queue)
    window.write_event_value('-THREAD-DONE-', 'load excel')


def main():
    sg.theme('BlueMono')
    sg.set_options(font=('Helvetica', 12))

    # Initialize important values for downstream analyses
    data_frame = None
    annotation = None

    menu_def = [
        ['Registry', ['Drugs']],
        ['Tools', ['Scan data']],
        ['About', ['Program']],
    ]

    layout = [
        [sg.Menu(menu_def)],
        [sg.Text('Mivisor 2022.1', font=('Helvetica', 28))],
        [sg.Text('Analytical Tools for Microbiology', font=('Helvetica', 20))],
        [sg.Frame(title='Quick Fix', expand_y=True, expand_x=True,
                  element_justification='center',
                  layout=[[sg.Button('Load data', key='-LOAD-')],
                          [sg.Button('Annotate columns', key='-ANNOTATE-')],
                          [sg.Button('Generate antibiogram', key='-GENERATE-')],
                          [sg.Exit(button_color='white on red')]]
                  )
         ]
    ]

    window = sg.Window('Mivisor v.2022.1',
                       layout=layout,
                       size=(800, 300),
                       element_justification='center',
                       resizable=True).finalize()

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'Exit':
            break
        elif event == '-LOAD-':
            filepath = sg.popup_get_file('Data file', title='Import Data', file_types=[('Excel', '*.xlsx *.xls')])
            if filepath:
                try:
                    sg.popup_quick_message('Loading data...', background_color='green')
                    thread_id = threading.Thread(
                        target=load_excel_data_thread,
                        args=(filepath, window, queue),
                        daemon=True,
                    )
                    thread_id.start()
                except:
                    sg.popup_error('Failed to open the file.', title='File Error')
        elif event == '-ANNOTATE-':
            if data_frame is not None:
                annotation = create_annotate_column_window(data_frame)
            else:
                sg.popup_quick_message('Load data first.', background_color='red')
        elif event == '-GENERATE-':
            if not annotation or data_frame is None:
                sg.popup_quick_message('Not enough data to generate an antibiogram.', background_color='red')
        elif event == '-THREAD-DONE-':
            df, data, headers = queue.get()
            data_frame = df
            if data and headers:
                create_data_table(data[:100], headers)
            else:
                sg.popup_error('Failed to open the file.', title='File Error')
    window.close()


if __name__ == '__main__':
    main()
