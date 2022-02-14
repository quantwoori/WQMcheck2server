from data_manage import *
from datetime import datetime, timedelta


if __name__ == "__main__":
    cd = CheckData()
    dt = (datetime.today() - timedelta(days=1)).strftime('%Y%m%d')

    cd.xl_cell_input(
        start_date="20220101",
        end_date=dt,
        stk_codes=cd.kospi.k100['STK_CODE'].values()
    )
