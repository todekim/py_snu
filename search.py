from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from datetime import datetime
from dateutil.relativedelta import relativedelta

import urllib
import pandas as pd
import time
import random
import chromedriver_autoinstaller

# 구글 검색 url
# cd_min: 검색 시작 날짜
# cd_max: 검색 종료 날짜
# 날짜 형식: 월/일/년
now = datetime.now()
startM = now - relativedelta(months=6) #6개월 전 데이터 부터 검색
end   = now.strftime("%m/%d/%Y") #오늘 날짜까지 검색 
start = startM.strftime("%m/%d/%Y") 
url = "https://www.google.com/search?tbm=nws&tbs=cdr:1,cd_min:"+start+",cd_max:"+end+"&&q="

f_name = "in.xlsx"
# 가공 후, 파일명 변경!!!

# 드라이버 실행
driver = webdriver.Chrome()

data = pd.read_excel(f_name, "Sheet1")
data.fillna('0')

# 결과 저장 list
result_list = list()

for row_num in range(len(data.index)):
    name = data.iloc[row_num]["이름"]
    comp = data.iloc[row_num]["소속"]

    ############ 엑셀 내부 기업명 정제 ############
    if 'A' in name:
        name = name.split('A')[0]

    if 'B' in name:
        name = name.split('B')[0]

    if "사명변경" in comp:
        comp = comp.split(":")[1].strip()

    if "\n" in comp:
        comp = comp.replace("\n", " ").strip()

    if "→" in comp:
        comp = comp.split('→')[1]

    if "주식회사" in comp:
        comp = comp.replace("주식회사", "").strip()

    if "특허법인" in comp:
        comp = comp.replace("특허법인", "").strip()

    if "법무법인" in comp:
        comp = comp.replace("법무법인", "").strip()

    if "특허사무소" in comp:
        comp = comp.replace("특허사무소", "").strip()

    if "협동조합" in comp:
        comp = comp.replace("협동조합", "").strip()

    if "㈜" in comp:
        comp = comp.replace("㈜", "").strip()

    if "(주)" in comp:
        comp = comp.replace("(주)", "").strip()

    if "(유)" in comp:
        comp = comp.replace("(유)", "").strip()

    if "/" in comp:
        comp = comp.split("/")[0]

    if "(" in comp:
        if comp.index("(") == 0:
            comp = comp.split(")")[1]

        if comp.index("(") > 1:
            comp = comp.split("(")[0]

        if "(" in comp:
            comp = comp.replace("(", "").strip()

        if ")" in comp:
            comp = comp.replace(")", "").strip()

    if "," in comp:
        comp = comp.split(",")[1].strip()
    ############ 엑셀 내부 기업명 정제 ############



    # 검색어 설정: 이름, 소속 필수로 포함, 투자/인수 둘 중 하나 포함
    search_word = f'"{name}"+"{comp}"+("투자"|"인수")'

    # 브라우저에 url 입력
    driver.get(url)

    # 검색어 입력창 엘리먼트 
    search_el = "/html/body/div[1]/div[3]/form/div[1]/div[1]/div[1]/div/div[2]/textarea"

    # 검색어 입력창 클릭
    driver.find_element(By.XPATH, search_el).click()

    # 검색어 한글자씩 입력
    for sep_word in list(search_word):
        # 구글 captcha 방지용 딜레이
        time_to_sleep = random.randint(0, 0)
        time.sleep(time_to_sleep)

        driver.find_element(By.XPATH, search_el).send_keys(sep_word)


    # 검색 버튼 클릭
    driver.find_element(By.XPATH, search_el).send_keys(Keys.ENTER)

    # 구글 captcha 방지용 딜레이
    time_to_sleep = random.randint(1, 3)
    time.sleep(time_to_sleep)

    # 검색 결과 파싱
    html = driver.page_source
    soup = BeautifulSoup(html, 'html.parser')

    ########## 검색 결과 선택자는 구글 자체에서 변경될 수 있어 확인 후 바뀌었으면 변경 필요 ##########
    # 검색 결과 링크 선택
    link = soup.select_one("#rso > div > div > div:nth-child(1) > div > div > a")
    # 검색 결과 제목 선택
    title = soup.select_one(
        "#rso > div > div > div:nth-child(1) > div > div > a > div > div.iRPxbe > div.n0jPhd.ynAwRc.MBeuO.nDgy9d")
    # 검색 결과 내용 선택
    content = soup.select_one(
        "#rso > div > div > div:nth-child(1) > div > a > div > div.iRPxbe > div.GI74Re.nDgy9d")
    # 검색 결과가 없습니다. 문구
    noSearch = soup.select_one("#topstuff > div > div > p:nth-child(1)")

    # 대체 검색어 추천
    replaceSearch = soup.select_one("#topstuff > div > div > div > a")

    # 검색 결과가 있는지 저장하는 변수
    flag = False

    # 검색 결과가 없거나 대체 검색어가 뜨지 않는 경우(검색 결과가 있음.)
    if noSearch is None and replaceSearch is None:
        print("in no search if")
        flag = True

    # 검색 결과 여부에 따라 진행
    if flag:
        print("in flag if")
        # 기사 링크, 제목이 있는 경우 list에 저장 후 콘솔에 출력
        if link is not None and title is not None:
            print("in flag in link if")
            result_list.append([urllib.parse.unquote(name), urllib.parse.unquote(comp), title.text,
                                link['href']])
            print(urllib.parse.unquote(name), urllib.parse.unquote(comp), title.text,
                  link['href'], sep=" ", end='\n')

    else:
        print("in flag else")
        print(urllib.parse.unquote(name), urllib.parse.unquote(comp), '',
              '', sep=" ", end='\n')

    # 구글 captcha 방지용 딜레이
    time_to_sleep = random.randint(2, 5)
    time.sleep(time_to_sleep)

# 검색 결과 list를 엑셀로 변환하기 위해 dataframe으로 변환
result_df = pd.DataFrame(result_list, columns=[['이름', '소속', '제목', '링크']])
print(result_df)

# 엑셀로 저장.(파일명의 날짜 수정 후 전달)
result_df.to_excel('out.xlsx', sheet_name="서치 결과")
