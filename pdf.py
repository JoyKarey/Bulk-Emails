import os
import re
import win32com.client as win32

def clean_file_name(file_name):
    """
    Clean the file name by removing leading and trailing spaces
    and special characters, but keep spaces between words.
    """
    # Strip leading and trailing spaces
    cleaned_name = file_name.strip()
    # Remove characters that are not alphanumeric or spaces
    cleaned_name = re.sub(r'[^a-zA-Z0-9\s]', '', cleaned_name)
    return cleaned_name

def save_excel_sheets_as_pdf(file_path, output_dir):
    # Create output directory if it doesn't exist
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    # Start Excel application
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False  # Run Excel in the background
    excel.DisplayAlerts = False  # Suppress Excel alerts

    try:
        # Disable automation security prompts
        excel.AutomationSecurity = 3

        print(f"Opening Excel file: {file_path}")
        workbook = excel.Workbooks.Open(file_path, UpdateLinks=False, ReadOnly=True)
        print(f"Successfully opened workbook: {workbook.Name}")

        # Loop through each sheet in the workbook
        for sheet in workbook.Sheets:
            try:
                # Clean the sheet name to be used as a file name
                cleaned_sheet_name = clean_file_name(sheet.Name)
                print(f"Processing sheet: {sheet.Name} as {cleaned_sheet_name}")
                
                # Construct the output PDF file path
                output_pdf = os.path.join(output_dir, f"{cleaned_sheet_name}.pdf")
                
                # Save the current sheet as PDF
                sheet.ExportAsFixedFormat(0, output_pdf)
                print(f"Saved: {output_pdf}")
            except Exception as e:
                print(f"Error processing sheet {sheet.Name}: {e}")
        
    except Exception as e:
        print(f"An error occurred: {e}")
    finally:
        # Close the workbook and Excel application
        try:
            if workbook:
                workbook.Close(SaveChanges=False)
        except Exception as e:
            print(f"Error closing workbook: {e}")
        excel.Quit()

# Usage example
file_path = r'C:\Users\JWANGARI\Desktop\PYTHON\PRIVATE AUGUST 2024 DATA.xlsx'  # Update with your file path
output_dir = r'C:\Users\JWANGARI\Desktop\PYTHON'  # Update with your output directory
save_excel_sheets_as_pdf(file_path, output_dir)
