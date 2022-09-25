import os
import sys
from tomlkit import parse, dump
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
    annotation = {}

    menu_def = [
        ['Registry', ['Drugs']],
        ['Tools', ['Scan data']],
        ['About', ['Program']],
    ]

    layout = [
        [sg.Menu(menu_def)],
        [sg.Text('Mivisor 2022.1', font=('Helvetica', 28))],
        [sg.Text('Analytical Tools for Microbiology', font=('Helvetica', 20))],
        [sg.Text('โดยคณะเทคนิคการแพทย์ มหาวิทยาลัยมหิดล', font=('Helvetica', 14))],
        [sg.Frame(title='Steps', expand_y=True, expand_x=True,
                  element_justification='center',
                  layout=[[sg.Button('Load data', key='-LOAD-', size=(20, 1))],
                          [sg.Button('Preview data', disabled=True, key='-PREVIEW-', size=(20, 1))],
                          [sg.Button('Annotate columns', key='-ANNOTATE-', size=(20, 1))],
                          [sg.Button('Generate antibiogram', key='-GENERATE-', size=(20, 1))],
                  ]
                  )
         ],
         [sg.Exit(button_color='white on red', size=(20, 1))],
    ]

    window = sg.Window('Mivisor v.2022.1',
                       layout=layout,
                       size=(800, 350),
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
                if os.path.exists('annotation.toml'):
                    annotation = parse(open('annotation.toml').read())
                annotation = create_annotate_column_window(data_frame, annotation)
                if annotation:
                    dump(annotation, open('annotation.toml', 'w'))
                    sg.popup_quick_message('Saved annotation successfully', background_color='green')
            else:
                sg.popup_quick_message('Load data first.', background_color='red')
        elif event == '-GENERATE-':
            if not annotation or data_frame is None:
                sg.popup_quick_message('Not enough data to generate an antibiogram.', background_color='red')
        elif event == '-THREAD-DONE-':
            df, data, headers = queue.get()
            data_frame = df
            if data and headers:
                window.find_element('-PREVIEW-').update(disabled=False)
                create_data_table(data[:100], len(df), headers)
            else:
                sg.popup_error('Failed to open the file.', title='File Error')
        elif event == '-PREVIEW-':
            if data and headers:
                create_data_table(data[:100], len(df), headers)
            else:
                sg.popup_error('ไม่สามารถเปิดข้อมูลได้ กรุณาลองเปิดไฟล์ข้อมูลอีกครั้ง', title='Error')
    window.close()


if __name__ == '__main__':
    main()
