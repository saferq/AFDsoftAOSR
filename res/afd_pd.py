import pandas as pd
# inp = [{'c1':10, 'c2':100}, {'c1':11,'c2':110}, {'c1':12,'c2':120}]
# df = pd.DataFrame(inp)

# for index, row in df.iterrows():
    # print(row["c2"])

class PandasWork():
    """ Работа с pandas """
    def create_df(self, date_df, col_name=1, col_start_content=2):
        """ Создание DataFrame """
        df = pd.DataFrame(date_df[col_start_content:], columns=date_df[col_name])
        return df