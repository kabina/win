import re
import sys
import io

#아톰 에디터에서 한글 출력 위해 선언
sys.stdout = io.TextIOWrapper(sys.stdout.detach(), encoding = 'utf-8')
sys.stderr = io.TextIOWrapper(sys.stderr.detach(), encoding = 'utf-8')

#텍스트 정제(전처리)
def cleanText(readData):
    #텍스트에 포함되어 있는 특수 문자 제거
    text = re.sub('[-=+,#/\?:^$.@*\"※~&%ㆍ!』\\‘|\(\)\[\]\<\>`\'…》]', '', readData)
    return text

if __name__ == "__main__":
    oriText = "!@#abcdefghijklmn!@#$%^&*()"
    print('+++ 텍스트 특수문자 제거 전 +++')
    print('Original : ', oriText)
    print('+++ 텍스트 특수문자 제거 후 +++')
    print('After Text : ', cleanText(oriText))
