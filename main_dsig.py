from data_manage import *


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