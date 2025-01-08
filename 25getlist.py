from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import load_workbook
from datetime import datetime
from urllib.parse import urlparse

# WebDriver 초기화
def init_driver():
    options = webdriver.ChromeOptions()
    options.add_argument("--headless")  # 백그라운드 실행
    options.add_argument("--disable-gpu")
    return webdriver.Chrome(options=options)

# URL 유효성 검증
def is_valid_url(url):
    try:
        result = urlparse(url)
        return all([result.scheme, result.netloc])
    except ValueError:
        return False

def get_list(driver, url):
    try:
        driver.get(url)

        # 데이터 추출
        data = {}

        def extract_value(xpath, method='attribute', attribute_name='value'):
            try:
                element = WebDriverWait(driver, 20).until(
                    EC.presence_of_element_located((By.XPATH, xpath))
                )
                if method == 'attribute':
                    value = element.get_attribute(attribute_name)
                else:
                    value = element.text
                
                print(f"XPath: {xpath} -> Value: {value if value else 'N/A'}")
                return value if value else "N/A"
            except Exception as e:
                print(f"Error while extracting value for XPath {xpath}: {e}")
                return "N/A"

        # 각 항목을 추출하여 사전 값에 저장
        data["사전규격등록번호"] = extract_value("//th[label[text()='사전규격등록번호']]/following-sibling::td/input")
        data["발주계획통합번호"] = extract_value("//th[label[text()='발주계획통합번호']]/following-sibling::td/a", method='text')
        data["사전규격명"] = extract_value("//th[label[text()='사전규격명']]/following-sibling::td/label", method='text')
        data["수요기관"] = extract_value("//th[label[text()='수요기관']]/following-sibling::td/input")
        #data["배정예산액"] = extract_value("//th[label[text()='배정예산액(부가세포함)']]/following-sibling::td/input")
        data["배정예산액"] = extract_value("//th[label[contains(text(),'배정예산액')]]/following-sibling::td/input")
        data["의견등록마감일시"] = extract_value("//th[label[text()='의견등록마감일시']]/following-sibling::td/input")
        data["사전규격 공개일시"] = extract_value("//th[label[text()='사전규격 공개일시']]/following-sibling::td/input")

        return data

    except Exception as e:
        print(f"Error while extracting data from URL: {url} | Error: {e}")
        return None


def process_row(row, list_row, worksheet, driver):
    url_value = worksheet.cell(row=row, column=11).value
    url = f"https://www.g2b.go.kr/link/PRVA004_02/?bfSpecRegNo={url_value}"
    if not is_valid_url(url):
        print(f"Invalid URL at row {row}: {url}")
        return

    data = get_list(driver, url)

    if data:
        # 데이터 기록
        worksheet.cell(row=list_row, column=1).value = list_row - 1
        worksheet.cell(row=list_row, column=3).value = datetime.now().date()
        worksheet.cell(row=list_row, column=5).value = data.get("사전규격 공개일시", "N/A")[:10]
        worksheet.cell(row=list_row, column=6).value = data.get("의견등록마감일시", "N/A")[:10]
        worksheet.cell(row=list_row, column=10).value = data.get("사전규격등록번호", "N/A")
        worksheet.cell(row=list_row, column=11).value = data.get("발주계획통합번호", "N/A")
        worksheet.cell(row=list_row, column=12).value = data.get("사전규격명", "N/A")
        worksheet.cell(row=list_row, column=13).value = data.get("수요기관", "N/A")
        worksheet.cell(row=list_row, column=15).value = data.get("사전규격명", "N/A")
        worksheet.cell(row=list_row, column=17).value = data.get("배정예산액", "N/A")

        print(f"Row {list_row}: Data written successfully -> {data}")
    else:
        print(f"Row {list_row}: No data extracted for URL {url}")

def main():
    excel_file_path = input("파일명+확장자: ")
    workbook = load_workbook(excel_file_path)
    worksheet = workbook['상세정보_작업']

    row = 76  # 데이터 시작 행
    list_row = 2

    driver = init_driver()  # 하나의 WebDriver만 사용

    while True:
        cell_value = worksheet.cell(row=row, column=6).value
        if cell_value != 1:  # 특정 값이 1인 셀만 처리
            break
        process_row(row, list_row, worksheet, driver)
        row += 1
        list_row += 1

    # 엑셀 저장
    workbook.save(excel_file_path)
    driver.quit()  # WebDriver 종료
    print("\n완료")

if __name__ == "__main__":
    main()
