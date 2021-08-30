# -*- coding: utf-8 -*-
"""특정 폴더의 산출물을 읽어서 분석하고 xlsx 리포트를 생성
    * 자동화를 위해  S/W개발4팀 WIN 과제로 진행
    * 멤버 : 안창선, 박성민, 김한영, 박근희
    * 설치가 필요한 Library : pandas, matplotlib, xlwings, numpy

ToDo:
    * 산출물 종류별로 풍성한 분석 기능을 추가
    * Output으로 그래프 포함 다양한 측면에서 시각화 보완
"""

from matplotlib import font_manager, rc
import matplotlib.pyplot as plt
import seaborn as sns
import xlwings as xw
import pandas as pd
import numpy as np
import os


# 그래프의 한글 깨짐 방지
font_path = "C:/Windows/Fonts/NGULIM.TTF"
font = font_manager.FontProperties(fname=font_path).get_name()
rc('font', family=font)

"""설정정보 지정
    * 산출물 폴더 위치 및 엑셀파일내 sheet명
    * 데이터 컬럼 리스트 및 실제이터 시작 위치 지정
"""
target_folder = "산출물폴더"
sheet_name = "Sheet1"
data_columns = ["업무", "상세기능", "단위기능", "담당자", "시작일", "종료일"]
start_row = 'A2'

def get_frame_data():
    """분석데이터를 읽어들임

    산출물 폴더 내의 엑셀 파일을 DataFrame Data로 읽어 들임

    Args:
        None

    Returns:
        df (pandas.DataFrame) : Pandas DataFrame Object
    Note:
        해당 폴더 내의 모든 xlsx파일을 읽어들여서 concat하여 리턴함
    """

    # 엑셀 파일 => frame_data로 변환
    file_list = os.listdir(target_folder)
    file_list = [f for f in file_list if f[0].isalnum()]

    df = None

    '''
    지정폴더 내에서 필요한 파일 리스트를 읽어서 데이터프레임으로 합친다
    '''
    for f in file_list:
        wb = xw.Book(os.path.join(target_folder, f))
        sheet = wb.sheets[sheet_name]
        d = sheet.range(start_row).options(pd.DataFrame, index=False, expand='table').value
        df = pd.DataFrame(d) if df is None else pd.concat([df, pd.DataFrame(d)], ignore_index=True)
        wb.close()
    df.columns = data_columns

    return df


def get_pre_processed(df):
    """DataFrame 화된 자료를 전처리

    DataFrame내의 자료를 분석하기 쉽게 전처리

    Args:
        df (DataFrame) : Pandas DataFrame Object (처리전)

    Returns:
        df (DataFrame) : Pandas DataFrame Object (처리후)


    Note:
        다양한 전처리기법이나 행/열의 추가/변경이 필요
    """

    '''
    담당자별 개발 프로그램 수 분석
    '''
    # 날짜 전처리
    df["시작일"].astype('datetime64[ns]')
    df["종료일"].astype('datetime64[ns]')
    df.insert(4, "소요시간", (df["종료일"] - df["시작일"]).dt.days)

    return df


def get_people_analysis(df):
    """DataFrame 화된 자료를 개인별 관점으로 분석

    DataFrame내의 자료를 개인별 측면에서 분석(개인별 물량 등)

    Args:
        df (DataFrame) : Pandas DataFrame Object

    Returns:
        p_min_funcs (Series ): 가장 적은 개발물량을 가지고 있는 개발자
        p_max_funcs (Series ) : 가장 많은 개발물량을 가지고 있는 개발자

    Note:
        가장 많은 노력이 들어가야 할 Function으로 산출물 유형에 따라 다양한 분석이 필요 함
    """

    '''
    담당자별 개발 프로그램 수 분석
    '''

    # 담당자별 개발 기능 시리즈
    pfuncs = df["담당자"].value_counts()
    # 총 개발 기능 수
    sum_funcs = sum(df["담당자"].value_counts())
    # 개인별 평균 개발 기능수
    avg_funcs = np.mean(df["담당자"].value_counts())
    # 개발물량 제일 적은 사람
    p_min_funcs = (pfuncs.loc[pfuncs == df["담당자"].value_counts().min()])
    # 개발물량 제일 많은 사람
    p_max_funcs = (pfuncs.loc[pfuncs == df["담당자"].value_counts().max()])
    # 상위 10%와 하위 10%간의 개발물량 배수 계산
    df_g = df['소요시간'].groupby(df['담당자']).mean()
    up_f = (df_g.sort_values().head(round(df.shape[0] / 10)))
    lo_f = (df_g.sort_values().tail(round(df.shape[0] / 10)))
    amt_ratio = (lo_f.mean() / up_f.mean())

    psummary = df.groupby(['담당자']).agg(소요시간=('소요시간', 'sum'), 기능수=('상세기능', 'count'))
    # ax = sns.scatterplot(x='MySum', y='MyCount', hue="담당자", data=-gdf)

    # plot_df = df.groupby(["담당자","업무명"])[["담당자", "업무명"]].count()

    plt.figure()
    ax = psummary.plot.bar()
    plt.title("개발자별 할당 프로그램")
    plt.xlabel("담당자")
    plt.ylabel("count")

    return p_min_funcs, p_max_funcs, amt_ratio, ax.get_figure(), psummary


def get_func_analysis(df):
    """DataFrame 화된 자료를 분석하기 위한 분석 Core 모듈

    DataFrame내의 자료를 순석대상 산출물 유형에 따라 다양하게 분석

    Args:
        df (DataFrame) : Pandas DataFrame Object

    Returns:
        per_func_ratio (Series): 상세기능 개발기간 최저/최다 비율
        ax (Plot Figure) : Plot 이미지

    Note:
        프로그램 기능별 다양한 분석 및 Plot Draw
    """
    import matplotlib.pyplot as plt
    # 업무별 개발 총 시간 분석
    df_g = df['소요시간'].groupby(df['상세기능']).mean()

    up_f = (df_g.sort_values().head(round(df.shape[0] / 10)))
    lo_f = (df_g.sort_values().tail(round(df.shape[0] / 10)))
    per_func_ratio = lo_f.mean() / up_f.mean()

    df_g = df['소요시간'].groupby(df['상세기능']).sum().reset_index()
    plt.figure()
    plt.xlabel("소요시간")
    plt.title("상세기능별 소요시간")
    ax = sns.barplot(x=df_g['상세기능'], y=df_g['소요시간'])
    return per_func_ratio, ax.get_figure()


def analysis(df):
    """산출물 분석 및 레포트를 출력하기 위한 Main 모듈

    분석은 분석 Core모듈에 대행하고, 분석 후 해당 데이터를 기준으로 그래프 Object를 생성하며,
    이를 결과 리포트로 내보낸다.

    Args:
        df (DataFrame) : Pandas DataFrame Object

    Returns:
        None

    Note:
        분석레포트 파일의 커스터마이징을 위해 해당 부분은 별도 Function으로 분리도 필요 함
    """

    a_data = dict()
    a_data['raw_data'] = df
    a_data['p_min_funcs'], a_data['p_max_funcs'], a_data['amt_ratio'], a_data['p_plot'], \
            a_data['psummary'] = get_people_analysis(df)
    a_data['f_anal'], a_data['f_plot'] = get_func_analysis(df)

    return a_data


def save_report(a_data):
    """산출물 분석 및 레포트를 출력하기 위한 Main 모듈

    분석은 분석 Core모듈에 대행하고, 분석 후 해당 데이터를 기준으로 그래프 Object를 생성하며,
    이를 결과 리포트로 내보낸다.

    Args:
        df (DataFrame) : Pandas DataFrame Object

    Returns:
        None

    Note:
        분석레포트 파일의 커스터마이징을 위해 해당 부분은 별도 Function으로 분리도 필요 함
    """

    bk = xw.Book()
    sh1 = bk.sheets(1)
    sh1.range("B1:B4").column_width = 15
    sh1['B2'].value = "Summary"
    sh1['B3'].value = "Max Amount"
    sh1['C3'].value = f"{a_data['p_max_funcs'].index[0]}"
    sh1['D3'].value = f"{a_data['p_max_funcs'][0]}"

    sh1['B4'].value = "Min Amount"
    sh1['C4'].value = f"{a_data['p_min_funcs'].index[0]}"
    sh1['D4'].value = f"{a_data['p_min_funcs'][0]}"

    sh1["A50"].options(pd.DataFrame, header=1, index=True, expand='table').value = a_data["psummary"]
    table_range = sh1.range('C5').expand()
    left = table_range.left + table_range.width + 1
    top = table_range.top
    sh1.pictures.add(a_data['p_plot'], name="개인별 그래프", update=True, left=left, top=top)

    table_range = sh1.range('C29').expand()
    left = table_range.left + table_range.width + 1
    top = table_range.top
    sh1.pictures.add(a_data['f_plot'], name="기능별 그래프", update=True, left=left, top=top)

    bk.save('리포트.xlsx')
    bk.close()


def main(name):
    """산출물 분석 Entry Function

    분석 데이터 생성, 분석, 결과리포트 저장 등의 Function을 호출

    Args:
        name (String) : Function Name이나, 기본 파이참 모듈 생성시 주는 값이라 걍 뒀음

    Returns:
        None

    Note:
        더이상 복잡하게 여기에 뭘 넣지는 않는게 좋겠음
    """
    try:
        app = xw.App(visible=False)
        # 분석 대상 데이터 수집
        df = get_frame_data()
        df = get_pre_processed(df)
        # 분석 & 리포트 저장
        save_report(analysis(df))
    except Exception as e:
        print("Exception:", e)
    finally:
        app.kill()


if __name__ == '__main__':
    main('WIN')
