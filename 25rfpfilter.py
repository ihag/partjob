import requests
import html
from openpyxl import load_workbook
from datetime import datetime
import os
import xlwings as xw

def fetch_data(bid_pbanc_no, bid_pbanc_ord):
    url = "https://www.g2b.go.kr/pn/pnp/pnpe/ItemBidPbac/selectItemAnncMngV.do"
    
    payload = {
        "dmItemMap":{
            "bidPbancNo": bid_pbanc_no,
            "bidPbancOrd": bid_pbanc_ord
        }
    }

    headers = {
        'Cookie': 'WHATAP=z1c19kr5h17kv5; XTVID=A2501060840080519; JSESSIONID=Mzg0MDJkOGEtMjAwNC00YTAyLTg1MDItMGUyNTBmZjIyOGMw;',
        'Content-Type': 'application/json;charset=UTF-8',
        'Accept': 'application/json',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36',
        'Referer': 'https://www.g2b.go.kr/',
        'Origin': 'https://www.g2b.go.kr',
        'Menu-Info': '{"menuNo":"01196","menuCangVal":"PNPE027_01","bsneClsfCd":"%EC%97%85130026","scrnNo":"06085"}',
        'SubmissionId': 'mf_wfm_container_mainWframe_selectItemAnncMngV',
        'Sec-CH-UA': '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
        'Sec-CH-UA-Mobile': '?0',
        'Sec-CH-UA-Platform': '"Windows"',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
    }

    response = requests.post(url, headers=headers, json=payload)

    if response.status_code == 200:
        return response.json()
    else:
        print(f"Error: {response.status_code} {response.text}")
        return None

def process_row(row, worksheet):
    bid_pbanc_no = worksheet.range(f'K{row}').value[:13]  # 11번째 열은 K열
    bid_pbanc_ord = worksheet.range(f'K{row}').value[-3:]  # 11번째 열은 K열
    data = fetch_data(bid_pbanc_no, bid_pbanc_ord)

    if data:
        dmItemMap = data.get("dmItemMap", {})
        worksheet.range(f'W{row}').value = dmItemMap.get("bfSpecRegNo", "N/A")
        dmst_unty_groupname = dmItemMap.get("dmstUntyGrpNm", "N/A")
        worksheet.range(f'X{row}').value = html.unescape(dmst_unty_groupname)
    else:
        print(f"Row {row}: No data extracted for bidPbancNo {bid_pbanc_no}")

def collect_j_column_values(worksheet):
    max_rows = 12000
    print(f"Scanning J column from row 2 to {max_rows}")
    
    j_values = []

    for row in range(2, max_rows + 1):
        j_value = worksheet.range(f'J{row}').value
        if isinstance(j_value, (int, float)):
            j_values.append(str(j_value).strip()) 
        elif isinstance(j_value, str) and j_value.startswith('R25'):
            j_values.append(j_value.strip())

    return set(j_values)

def open_excel_file(file_path):
    try:
        app = xw.App(visible=False)
        wb = app.books.open(file_path)
        return wb, app
    except FileNotFoundError:
        print(f"File {file_path} not found.")
        return None, None

def process_excel(file_path):
    wb, app = open_excel_file(file_path)
    
    if wb is None:
        print("Workbook could not be opened. Exiting.")
        return
    
    worksheet = wb.sheets['상세정보_작업']
    print(f"Processing sheet: {worksheet.name}")

    j_values_set = collect_j_column_values(worksheet)

    row = 2
    while True:
        v_value = worksheet.range(f'V{row}').value
        t_value = worksheet.range(f'T{row}').value

        print(f"Processing row {row} - T column value: {t_value}, V column value: {v_value}")

        if v_value is None:
            if not t_value:
                print(f"T column is empty in row {row}. Stopping process.")
                break

            if t_value == '??':
                print(f"Skipping row {row} because T column has '??'")
                row += 1
                continue

            if isinstance(t_value, str) and t_value.startswith("http"):
                process_row(row, worksheet)

                number = worksheet.range(f'W{row}').value  # 현재 행에서 값을 가져옴
                if number and number != '0':
                    if number in j_values_set:
                        print(f"Found matching number {number} in J column set. Updating row {row}.")
                        worksheet.range(f'V{row}').value = 0
                    else:
                        print(f"No match for registration number {number} in J column set.")
                else:
                    print(f"Invalid or no registration number found for row {row}.")
            else:
                print(f"No valid hyperlink found in T column for row {row}.")
        else:
            print(f"Skipping row {row} because V column already has a value.")
        
        row += 1

    print("Saving changes to the Excel file.")
    wb.save()
    wb.close()
    app.quit()
    print(f"File saved: {file_path}")

def main():
    current_folder = os.getcwd()
    file_name = input("Enter the Excel file name (with extension): ")
    file_path = os.path.join(current_folder, file_name)

    process_excel(file_path)

if __name__ == "__main__":
    main()
