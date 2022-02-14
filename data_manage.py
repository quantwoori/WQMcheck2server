from dbms.DBmssql import MSSQL
from settings.Sstk import KOSPI

import openpyxl as op
import pandas as pd

from datetime import datetime, timedelta


class XLClean:
    def __init__(self, excel_file:str):
        self.data = pd.read_excel(excel_file)
        self.server = MSSQL.instance()
        self.server.login(
            id='wsol1',
            pw='wsol1'
        )
        self.kospi = KOSPI()

    @staticmethod
    def clean_date(excel_date:int, dfmt:str="%Y%m%d") -> str:
        dt = datetime.fromordinal(
            int(
                datetime(1900, 1, 1).toordinal() +
                excel_date - 2
            )
        )
        return dt.strftime(dfmt)

    def clean_stock(self, stock_name:str):
        return self.kospi.k100['STK_CODE'][stock_name]

    @staticmethod
    def clean_column(df:pd.DataFrame, num_col:int) -> pd.DataFrame:
        """
        CLEAN CHECK DATA MESS INTO SOMETHING USABLE
        :param num_col:
        How many parameters did you requested?
        :return:
        pd.DataFrame
        """
        GROUP_COL = 2 + num_col
        SEG_COL = ['date', 'borrowed', 'd_borrowed', 'r_borrowed']

        result = pd.DataFrame(None)
        for i in range(0, len(df.columns), GROUP_COL):
            name = df.columns[i : (i + GROUP_COL)][1]
            seg = df[df.columns[i : (i + GROUP_COL)]][1:]  # First Column is a Bogus
            seg.columns = SEG_COL
            seg['stock'] = [name] * len(seg)

            result = pd.concat([result, seg])
        result = result.reset_index(drop=True)
        result = result.dropna()
        return result


class CheckData:
    def __init__(self, path:str="test.xlsx", result_path:str='result.xlsx'):
        # CONSTANT
        self.COST = 4
        self.ROW_START, self.COL_START = 1, 1

        # Excel Workfile
        self.wf = op.Workbook()
        self.dpath = path
        self.rpath = result_path

        # Target Stocks
        self.kospi = KOSPI()

    def xl_func_writer(self, start_date:str, end_date:str, stock_code:str) -> str:
        func = f'=CH("{start_date}", "{end_date}", -1, "D", FALSE, "{stock_code}", "14212,14214,14216", "ASC", "withtable=true;")'
        return func

    def xl_cell_input(self, start_date:str, end_date:str, stk_codes:set) -> None:
        ws = self.wf.active
        r = 0
        for s in stk_codes:
            val = self.xl_func_writer(start_date, end_date, stock_code=s)
            ws.cell(row=self.ROW_START,
                    column=self.COL_START + r * self.COST).value = val
            r += 1

        self.wf.save(self.dpath)

    def process_rpa_res(self, loc=r'C:\Users\Check\Documents\result.xlsx'):
        d = pd.read_excel(loc)


def main_database():
    # Data
    xl = XLClean(r'C:/Users/Check/Documents/result.xlsx')
    d = xl.clean_column(xl.data, 2)
    d.date = d.date.apply(xl.clean_date)
    d.stock = d.stock.apply(xl.clean_stock)

    # Database
    DB_COL = ['date', 'borrow', 'd_borrow', 'r_borrow', 'stkcode']
    d_db = d.to_numpy().tolist()
    d_db = [tuple(_) for _ in d_db]
    xl.server.insert_row(
        table_name='RAWborrow',
        schema='dbo',
        database='WSOL',
        col_=DB_COL,
        rows_=d_db
    )


class SigBand:
    def __init__(self):
        self.kospi = KOSPI()
        self.server = MSSQL().instance()
        self.server.login(
            id='wsol1',
            pw='wsol1'
        )

    def get_data(self, standard_date:datetime, calc_dates:int=20, dfmt='%Y%m%d') -> pd.DataFrame:
        # DATE FOR CALCULATION. CONSTANT
        START_DATE = (standard_date - timedelta(days=1)).strftime(dfmt)

        stocks = self.kospi.k100['STK_CODE'].values()
        result = pd.DataFrame(None)
        for s in stocks:
            # COLUMN
            COL = ['date', 'stkcode', 'r_borrow']

            # CONDITION
            cond1 = f"stkcode = '{s}'"
            cond2 = f"date <= '{START_DATE}'"
            cond = ' and '.join([cond1, cond2])

            seg = self.server.select_recent(
                database='WSOL',
                schema='dbo',
                table='RAWborrow',
                column=COL,
                recent=calc_dates,
                condition=cond
            )
            seg = pd.DataFrame(seg, columns=COL)
            result = pd.concat([result, seg])

        return result

    def get_band(self, data:pd.DataFrame, target:str='r_borrow', bandwidth:float=2):
        """
        :param data: DataFrame for single stkcode
        :param target: Name of the target column
        :return:
        """
        ub = data[target].mean() + bandwidth * data[target].std()
        lb = data[target].mean() - bandwidth * data[target].std()
        return ub, lb

    def signal(self, boundary:tuple, value:float):
        u, l = boundary
        if value > u:
            return 'abnormal_high'
        elif value < l:
            return 'abnormal_low'
        else:
            return 'normal'


def main_signal(dt:datetime=datetime.today()):
    sig = SigBand()
    d = sig.get_data(standard_date=dt)
    g = d.groupby('stkcode')
    result = list()
    for name, seg in g:
        up, down = sig.get_band(seg)
        date = datetime.strptime(
            seg['date'].max(),
            "%Y%m%d"
        ).strftime('%Y-%m-%d')

        value = seg['r_borrow'][0]
        signal = sig.signal(
            boundary=(up, down),
            value=value
        )

        if signal == 'abnormal_high':
            sig_stren = (value - up) * 10 ** 5
        elif signal == 'abnormal_low':
            sig_stren = (value - down) * 10 ** 5
        else:
            # Normal
            sig_stren = 0 * 10 ** 5
        result.append(
            (date, int(sig_stren), signal, name, 'shortsig_v2_10_5')
        )
    return result


if __name__ == "__main__":
    sig = SigBand()
    m = main_signal()
    sig.server.insert_row(
        table_name='sig',
        schema='dbo',
        database='WSOL',
        col_=['date','sigstren', 'sig','stk','sigtyp'],
        rows_=m
    )