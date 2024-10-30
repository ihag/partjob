import os
import xlwings as xw
from bs4 import BeautifulSoup
import requests

# Step 1: Open the Excel file using xlwings
def open_excel_file(file_path):
    try:
        app = xw.App(visible=False)  # Open Excel in the background
        wb = app.books.open(file_path)
        return wb, app
    except FileNotFoundError:
        print(f"File {file_path} not found.")
        return None, None

# Step 2: Collect all numeric values from J column directly, even with gaps
def collect_j_column_values(sheet):
    max_rows = 12000  # Set a high row limit to ensure we capture all data in column J
    print(f"Scanning J column from row 2 to {max_rows}")
    
    j_values = []

    # Loop through all rows in column J from row 2 to max_rows
    for row in range(2, max_rows + 1):
        j_value = sheet.range(f'J{row}').value  # Get the value from column J

        # Check if the value is a number (int or float), and not None
        if isinstance(j_value, (int, float)):
            j_values.append(int(j_value))  # Convert to integer and store in list

    return set(j_values)  # Convert list to set to remove duplicates

# Step 3: Process the Excel file and check reg_no against the set of J values
def process_excel(file_path):
    wb, app = open_excel_file(file_path)
    
    if wb is None:
        print("Workbook could not be opened. Exiting.")
        return  # Exit if the file couldn't be opened
    
    sheet = wb.sheets['상세정보_작업']  # Open your specific sheet
    print(f"Processing sheet: {sheet.name}")

    # Collect J column values into a set
    j_values_set = collect_j_column_values(sheet)

    # Loop through rows, starting from row 2
    row = 2
    while True:
        v_value = sheet.range(f'V{row}').value  # V column (22nd)
        t_value = sheet.range(f'T{row}').value  # T column (20th)

        # Debugging: Print current row processing status
        print(f"Processing row {row} - T column value: {t_value}, V column value: {v_value}")

        # If V column is empty, process the row
        if v_value is None:
            # If T column is empty, stop processing
            if not t_value:
                print(f"T column is empty in row {row}. Stopping process.")
                break  # Stop processing when T column is empty

            # Skip rows where T column contains '??'
            if t_value == '??':
                print(f"Skipping row {row} because T column has '??'")
                row += 1
                continue

            # Proceed if T column contains a valid hyperlink
            if isinstance(t_value, str) and t_value.startswith("http"):
                url = t_value
                print(f"Processing URL in row {row}: {url}")
                
                # Try to fetch registration number from URL
                number = get_registration_number(url)

                # If we get a valid registration number, check if it exists in the set
                if number and number != '0':
                    if number in j_values_set:  # Check against the pre-loaded set of J column values
                        print(f"Found matching number {number} in J column set. Updating row {row}.")
                        sheet.range(f'V{row}').value = 0  # Set V column to 0
                    else:
                        print(f"No match for registration number {number} in J column set.")
                    
                    # Write the registration number in W column and add debug output
                    print(f"Writing registration number {number} to W column in row {row}.")
                    sheet.range(f'W{row}').value = number  # Write the registration number to W column
                else:
                    print(f"Invalid or no registration number found for row {row}.")
            else:
                print(f"No valid hyperlink found in T column for row {row}.")
        else:
            print(f"Skipping row {row} because V column already has a value.")
        
        row += 1

    print("Saving changes to the Excel file.")
    wb.save()  # Save the workbook after processing
    wb.close()
    app.quit()  # Close the Excel application
    print(f"File saved: {file_path}")

# Step 4: Get registration number from the URL
def get_registration_number(url):
    try:
        print(f"Fetching registration number from URL: {url}")
        response = requests.get(url)
        response.raise_for_status()  # Raise an exception for HTTP errors
        soup = BeautifulSoup(response.content, 'html.parser')
        
        reg_no = soup.find(string="사전규격등록번호").find_next("td").text.strip()  # Get registration number
        print(f"Registration number found: {reg_no}")
        return int(reg_no)
    except Exception as e:
        print(f"Error fetching registration number from {url}: {e}")
        return None

# Step 5: Main function to handle the file processing
def main():
    current_folder = os.getcwd()  # Get current working directory
    file_name = input("Enter the Excel file name (with extension): ")
    file_path = os.path.join(current_folder, file_name)
    
    process_excel(file_path)
    print("Processing completed.")

if __name__ == "__main__":
    main()
