import sys
from bs4 import BeautifulSoup
import requests
import urllib.request
import urllib.parse
import pandas as pd
import xlwings as xw
import re

# 명령줄 인자로 회사명 입력 받기
if len(sys.argv) < 2:
    print("회사명을 입력해야 합니다.")
    sys.exit()

company = sys.argv[1]  # 회사명
encoded_company = urllib.parse.quote_plus(company)

# 검색 URL 생성
base_url = 'https://allthatcompany.com/search?q='
search_url = base_url + encoded_company

# 첫 페이지 요청 및 파싱
r = requests.get(search_url)
soup = BeautifulSoup(r.text, "html.parser")

# 검색 결과 테이블과 페이지 링크 파싱
items = soup.select("#search-table")
pages = soup.select(".page-link")

# 페이지 수 계산
pageNum = 1
a = len(pages)

if a == 0:
    lPage = 1
else:
    lastPage = pages[len(pages)-2].text
    lPage = int(lastPage)

# 빈 데이터프레임 초기화
df = pd.DataFrame()

# 모든 페이지에 대해 반복하여 데이터 수집
while pageNum <= lPage:
    url = f"https://allthatcompany.com/search?query={encoded_company}&q={encoded_company}&page={pageNum}"

    html = urllib.request.urlopen(url).read()
    soup2 = BeautifulSoup(html, 'html.parser')

    items = soup2.find(class_='table table-bordered')
    a_tags = items.find_all('a')

    # href와 텍스트를 추출하여 리스트로 저장
    result_data = [("https://allthatcompany.com" + a.get('href'), a.get_text(strip=True)) for a in a_tags]

    # 새로운 데이터를 데이터프레임으로 변환합니다.
    new_df = pd.DataFrame.from_records(result_data, columns=['주소', '회사명'])

    # 기존 데이터프레임에 새로운 데이터를 추가합니다.
    df = pd.concat([df, new_df], ignore_index=True)

    pageNum += 1

# 시트명 설정 및 특수 문자 제거 함수
def sanitize_sheet_name(name):
    # 길이 제한
    name = name[:31]
    # 특수 문자 제거
    return re.sub(r'[\\/*?:[\]<>|]', "", name)

# 입력된 회사명으로 시트명 설정 (특수 문자 제거 후)
sheet_name = sanitize_sheet_name(company)

# 현재 열려 있는 엑셀 파일에 접근하여 데이터를 추가
wb = xw.Book.caller()  # 현재 열려있는 엑셀 파일
try:
    # 시트가 존재하는지 확인하고, 존재하면 그 시트를 사용
    sheet = wb.sheets[sheet_name]
except:
    # 존재하지 않으면 새 시트를 추가
    sheet = wb.sheets.add(sheet_name)

# 데이터프레임을 엑셀 시트에 기록
sheet.clear()  # 기존 내용을 지우고 새로 쓰기
sheet.range("A1").value = df

print(f"엑셀 문서에 {sheet_name} 시트가 생성되거나 업데이트되었습니다.")
