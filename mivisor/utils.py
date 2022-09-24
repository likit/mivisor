import pandas as pd


def load_excel_data(filepath, queue):
    data = []
    try:
        df = pd.read_excel(filepath)
    except ValueError:
        queue.put((None, None, None))
    else:
        for idx, row in df.iterrows():
            data.append(row.to_list())
            queue.put((df, data, list(df.columns)))