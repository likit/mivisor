import PySimpleGUI as sg


def create_data_table(data, headers):
    layout = [
        [sg.Table(values=data,
                  key='-TABLE-',
                  headings=headers,
                  auto_size_columns=True,
                  display_row_numbers=True,
                  num_rows=20,
                  vertical_scroll_only=False,
                  expand_y=True,
                  expand_x=False)],
        [sg.Text('Total rows:'), sg.Text(len(data))],
        [sg.Text('Total columns:'), sg.Text(len(headers))],
        [sg.CloseButton('Close')]
    ]

    window = sg.Window('Imported Data Preview', layout=layout, resizable=True).finalize()

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == 'Exit':
            break
    window.close()
