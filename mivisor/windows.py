import PySimpleGUI as sg


def create_data_table(data, headers):
    layout = [
        [sg.Text('Data Preview', font=('Helvetica', 18))],
        [sg.Table(values=data,
                  key='-TABLE-',
                  headings=headers,
                  auto_size_columns=True,
                  display_row_numbers=True,
                  num_rows=20,
                  vertical_scroll_only=False,
                  expand_y=True,
                  expand_x=False)],
        [sg.Frame(title='Data Format',
                  expand_x=True,
                  layout=[[sg.Checkbox('ข้อมูลมีชื่อคอลัมน์เป็นชื่อย่อของยา หรือชื่อยา', default=True, key='-NOT-FLAT-TABLE-')]])],
        [sg.Frame(title='Summary',
                  expand_x=True,
                  layout=[
                      [sg.Text('Total rows:'), sg.Text(len(data))],
                      [sg.Text('Total columns:'), sg.Text(len(headers))],
                  ]
                  )],
        [sg.CloseButton('Close')]
    ]

    window = sg.Window('Imported Data Preview', layout=layout, resizable=True).finalize()

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'Exit':
            break
    window.close()


def create_annotate_column_window(df, annot):
    headers = list(df)
    id_ = annot.get('-ID-')
    date_ = annot.get('-DATE-')
    drugs_ = annot.get('-DRUGS-')
    layout = [
        [sg.Text('Identifier:', size=(10, 1)), sg.Combo(values=headers, key='-ID-', expand_x=True, default_value=id_)],
        [sg.Text('Date:', size=(10, 1)), sg.Combo(values=headers, key='-DATE-', expand_x=True, default_value=date_)],
        [sg.Text('Drugs:', size=(10, 1)), sg.Listbox(values=headers,
                                                     default_values=drugs_,
                                                     highlight_background_color='blue',
                                                     key='-DRUGS-',
                                                     select_mode=sg.LISTBOX_SELECT_MODE_EXTENDED,
                                                     size=(20, 20))],
        [sg.OK(), sg.CloseButton('Close'), sg.Help()]
    ]

    window = sg.Window('Column Annotation', layout=layout, resizable=True).finalize()

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'Exit':
            values = None
            break
        elif event == 'Help':
            sg.popup_ok('ในการสร้าง antibiogram โปรแกรมต้องทราบคอลัมน์ที่จะเป็นตัวแทนของผู้ป่วย (identifier)',
                        'และต้องทราบวันที่เพื่อใช้ในการ deduplicate และกรองข้อมูล',
                        'นอกจากนั้นโปรแกรมต้องทราบคอลัมน์ที่เป็นชื่อยาเพื่อใช้ในการวิเคราะห์',
                        background_color='white')
        elif event == 'OK':
            error = False
            if not values['-ID-']:
                sg.popup_quick_message('Identifier is not speficied.', background_color='red')
                error = True

            if values['-DATE-']:
                try:
                    dates = df[values['-DATE-']].dt.date
                except:
                    sg.popup_quick_message('Date column is not valid.', background_color='red')
                    error = True
            else:
                sg.popup_quick_message('Date column is not specified.', background_color='red')
                error = True

            if not values['-DRUGS-']:
                sg.popup_quick_message('No drugs specified.', background_color='red')
                error = True

            if not error:
                if sg.popup_ok_cancel('คุณได้ระบุคอลัมน์เสร็จสิ้นแล้วหรือไม่') == 'OK':
                    break
    window.close()
    return values
