# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

def draw_plot(df):
    import matplotlib.pyplot as plt
    from matplotlib import font_manager, rc

    font_path = "C:/Windows/Fonts/NGULIM.TTF"
    font = font_manager.FontProperties(fname=font_path).get_name()
    rc('font', family=font)

    plot_df = df.groupby(["담당자"])[["담당자"]].count()
    plot_df.plot.bar()
    plt.title("개발자별 할당 프로그램")
    plt.xlabel("담당자")
    plt.ylabel("count")
    plt.show()

def analysis(name):
    import xlwings as xw
    import pandas as pd
    import numpy as np
    import os

    dir_name = "산출물폴더"
    sheet_name = "Sheet1"
    # 엑셀 인스턴스 생성
    app = xw.App(visible=False)
    # 파일 상장법인목록
    file_list = os.listdir(dir_name)
    file_list = [f for f in file_list if f[0].isalnum()]
    df = None

    '''
    지정폴더 내에서 필요한 파일 리스트를 읽어서 데이터프레임으로 합친다
    '''
    for f in file_list :
        wb = xw.Book(os.path.join(dir_name, f))
        sheet = wb.sheets[sheet_name]
        d = sheet.range('A1').options(pd.DataFrame, index=False, expand='table').value
        df = pd.DataFrame(d) if df is None else pd.concat([df, pd.DataFrame(d)], ignore_index=True)
        wb.close()

    '''
    담당자별 개발 프로그램 수 분석
    '''
    #날짜 전처리
    df["시작일"].astype('datetime64[ns]')
    df["종료일"].astype('datetime64[ns]')
    df.insert(4, "시간", df["종료일"] - df["시작일"])

    #담당자별 개발 기능 시리즈
    pfuncs = df["담당자"].value_counts()
    work_load = df.groupby("담당자")["시간"]

    # 총 개발 기능 수
    sum_funcs = sum(df["담당자"].value_counts())
    # 개인별 평균 개발 기능수
    avg_funcs = np.mean(df["담당자"].value_counts())
    # 개인별 평균 개발 기간
    # avg_funcs = np.mean(df["담당자"].value_counts())
    # 개발물량 제일 적은 사람
    p_min_funcs = (pfuncs.loc[pfuncs == df["담당자"].value_counts().min()])
    # 개발물량 제일 많은 사람
    p_max_funcs = (pfuncs.loc[pfuncs == df["담당자"].value_counts().max()])

    '''
    분석 summary 출력
    '''
    op1 = f"Summary{'*'*20}"
    op2 = f"물량 제일 많은 사람은 \n{'*'*20}\n {p_max_funcs}"
    op3 = f"물량 제일 적은 사람은 \n{'*' * 20}\n {p_min_funcs}"

    '''
    분석 Summary 파일로 출력
    '''
    draw_plot(df)

    bk = xw.Book()
    sh1 = bk.sheets(1)
    sh1['B2'].value = op1
    sh1['B3'].value = op2
    sh1['B4'].value = op3

    bk.save('리포트.xlsx')
    bk.close()

    # 인스턴스 종료
    app.kill()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    analysis('PyCharm')

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
