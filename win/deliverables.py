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

def get_file_list(read_type=1):
    """폴더 내에서 특정 버전의 산출물만 추려 내는 기능

    read_type에 따라서 특정 버전을 가져오거나, 업무영역별로 최신 버전을 읽어 들인다
    (현재 특정 버전만 가져옴)

    Args:
        read_type (int) :   0 == 특정 버전(V1.0)
                            1 == 최대 버전

    Returns:
        new_files

    Note:
        해당 업무영역의 최신 버전을 읽어 오도록 개선 필요
    """
    file_list = os.listdir(target_folder)
    file_list = [f.split('_') for f in file_list]

    dfs = pd.DataFrame(file_list, columns= ['DIV', 'PRJ_NM', 'STEPS','DELIVERABLES','DIV_NAME','YYMMDD','VERSION'])
    dfs['VERSION'] = dfs['VERSION'].apply(lambda x : x[0:4])
    if read_type == 0:
        dfs = dfs.query("VERSION.str.startswith('V1.0')")
    else:
        df_max_version = dfs.groupby(['DIV'])['VERSION'].max()
        for k, v in df_max_version.iteritems():
            dfs = dfs.drop(dfs[(dfs['DIV'] == k) & (dfs['VERSION'] != v)].index)

    dfs['VERSION'] = dfs['VERSION'].apply(lambda x: x+'.xlsx')
    new_files = ["_".join(row) for _, row in dfs.iterrows()]

    return new_files


if __name__ == '__main__':
    get_file_list(read_type=1)
