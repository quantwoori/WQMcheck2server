from data_manage import *
from datetime import datetime, timedelta


if __name__ == "__main__":
    cd = CheckData()
    if datetime.now().strftime('%H%M%S') > '170000':
        dt= datetime.today().strftime('%Y%m%d')
    else:
        dt = (datetime.today() - timedelta(days=1)).strftime('%Y%m%d')

    print("###### 날짜 세팅 ######")
    print("다음 중 해당 사항 체크")
    print("1. 어제가 공유일이었다면 휴일 횟수를 적으세요.(주말 포함)")
    print("  -> 만약 오늘이 수요일이고 어제(화요일) 하루만 공휴일이라면")
    print("     1 이라고 적으세요.")
    print("     ex) 1")
    print("  -> 만약 오늘이 월요일이고 전거래일(금요일) 하루가 공휴일이었다면")
    print("     1(금요일) + 2(토, 일) 해서 3 이라고 적으세요")
    print("     ex) 3")
    print("  -> 휴일이 아니었다면 그냥 0이라고 적으세요")
    print("     ex) 0")



    while True:
        manual = input()
        try:
            m = int(manual)
            print(manual, "이라고 답변하셨습니다.")
        except ValueError:
            print(manual, "이라고 답변하셨습니다. 이는 정수(integer)가 아닙니다.")
            print("정수로 입력 바랍니다. Again")
        else:
            break

    holiday = m
    if datetime.now().strftime('%A') == 'Monday':
        # Find Friday
        dt = (datetime.today() - timedelta(days=3)).strftime('%Y%m%d')

    if holiday > 0:
        # Find Day Before the Holiday
        dt = (datetime.today() - timedelta(days=(holiday + 1))).strftime('%Y%m%d')
    print(dt, "날짜로 데이터를 불러옵니다.")
    cd.xl_cell_input(
        start_date=dt,
        end_date=dt,
        stk_codes=cd.kospi.k100['STK_CODE'].values()
    )

    print("%%%%%% 주의 사항 %%%%%%")
    print("어제 기업 이름 변경 공시가 떴다면 --> settings /Sstk.py 에서 기업을 찾아서 이름을 바꿔주세요.")
    print('"POSCO": "005490"  -->  "POSCO홀딩스": "005490"')
