# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.


def analysis(name):
    import xlwings as xw
    import pandas as pd

    # 엑셀 인스턴스 생성
    app = xw.App(visible=False)
    # 파일 상장법인목록
    wb = xw.Book('프로그램목록.xlsx')
    # 첫번째 시트 읽어오기
    sheet = wb.sheets[0]
    # 데이터프레임 형태로 엑셀 시트 읽어오기
    df = sheet.range('A1').options(pd.DataFrame, index=False, expand='table').value

    print(df)
    print('최대값->', df['상세기능'].value_counts())

    # 인스턴스 종료
    app.kill()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    analysis('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
