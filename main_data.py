from data_manage import *
from datetime import datetime, timedelta


if __name__ == "__main__":
    cd = CheckData()
    if datetime.now().strftime('%H%M%S') > '170000':
        dt= datetime.today().strftime('%Y%m%d')
    else:
        dt = (datetime.today() - timedelta(days=1)).strftime('%Y%m%d')

    holiday = 0
    if datetime.now().strftime('%A') == 'Monday':
        # Find Friday
        dt = (datetime.today() - timedelta(days=3)).strftime('%Y%m%d')
    elif holiday > 0:
        # Find Day Before the Holiday
        dt = (datetime.today() - timedelta(days=(holiday + 1))).strftime('%Y%m%d')
    print(dt)
    cd.xl_cell_input(
        start_date=dt,
        end_date=dt,
        stk_codes=cd.kospi.k100['STK_CODE'].values()
    )
