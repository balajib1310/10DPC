
import pyautogui
import pyperclip
import time
import re
import os
import subprocess
from datetime import datetime

# ============================================================
# EXCEL COM SUPPORT
# ============================================================

try:
    import win32com.client
except ImportError:
    raise RuntimeError(
        "\npywin32 is not installed.\n\n"
        "Run:\n"
        'python -m pip install pywin32\n'
    )


# ============================================================
# CONFIGURATION
# ============================================================

pyautogui.FAILSAFE = True
pyautogui.PAUSE = 0.3

PROJECT_DIR = os.path.dirname(os.path.abspath(__file__))

today = datetime.now().strftime("%Y-%m-%d")
current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

excel_filename = f"daily_report_{today}.xlsx"
screenshot_filename = f"daily_report_screenshot_{today}.png"

excel_path = os.path.join(PROJECT_DIR, excel_filename)
screenshot_path = os.path.join(PROJECT_DIR, screenshot_filename)

chrome_path = r"C:\Program Files\Google\Chrome\Application\chrome.exe"
chrome_profile = "Default"

website_url = "https://quotes.toscrape.com/"


# ============================================================
# HELPER
# ============================================================

def wait(seconds, message=None):
    if message:
        print(message)

    time.sleep(seconds)


def fail(message):
    print("\n" + "=" * 60)
    print("ERROR")
    print("=" * 60)
    print(message)
    print("=" * 60)

    raise RuntimeError(message)


# ============================================================
# STEP 1 - OPEN CHROME
# ============================================================

print("\n" + "=" * 60)
print("STEP 1 - OPENING CHROME")
print("=" * 60)

print("Chrome profile :", chrome_profile)
print("Website        :", website_url)

if not os.path.exists(chrome_path):
    fail(
        f"Chrome was not found at:\n{chrome_path}"
    )

print("Launching Chrome...")

subprocess.Popen([
    chrome_path,
    "--profile-directory=Default",
    "--new-window",
    website_url
])

wait(8, "Waiting for Chrome...")


# ============================================================
# STEP 2 - OPEN WEBSITE
# ============================================================

print("\n" + "=" * 60)
print("STEP 2 - OPENING WEBSITE")
print("=" * 60)

print("Making sure Chrome is displaying the website...")

pyautogui.hotkey("ctrl", "l")
wait(0.5)

pyperclip.copy(website_url)
pyautogui.hotkey("ctrl", "v")
pyautogui.press("enter")

wait(8, "Waiting for webpage to load...")

print("Website loaded.")


# ============================================================
# STEP 3 - COPY WEBPAGE
# ============================================================

print("\n" + "=" * 60)
print("STEP 3 - COPYING WEBPAGE")
print("=" * 60)

screen_width, screen_height = pyautogui.size()

print(
    f"Screen resolution detected: "
    f"{screen_width} x {screen_height}"
)

pyautogui.click(
    screen_width // 2,
    screen_height // 2
)

wait(1)

print("Selecting webpage content...")

pyautogui.hotkey("ctrl", "a")

wait(1)

print("Copying webpage content...")

pyautogui.hotkey("ctrl", "c")

wait(2)


# ============================================================
# CHECK CLIPBOARD
# ============================================================

page_content = pyperclip.paste()

if not page_content.strip():

    fail(
        "Could not copy webpage content from Chrome."
    )

print("Webpage content copied successfully.")

print("\nFirst 500 characters:")
print("-" * 60)
print(page_content[:500])
print("-" * 60)


# ============================================================
# STEP 4 - EXTRACT FIRST QUOTE
# ============================================================

print("\n" + "=" * 60)
print("STEP 4 - EXTRACTING QUOTE")
print("=" * 60)

quote_match = None

quote_match = re.search(
    r'“(.*?)”',
    page_content,
    re.DOTALL
)

if not quote_match:

    quote_match = re.search(
        r'"(.*?)"',
        page_content,
        re.DOTALL
    )


if quote_match:

    fetched_data = quote_match.group(1).strip()

    fetched_data = re.sub(
        r"\s+",
        " ",
        fetched_data
    )

else:

    print("Quote pattern not found.")
    print("Using fallback extraction.")

    lines = [
        line.strip()
        for line in page_content.splitlines()
        if line.strip()
    ]

    fetched_data = next(
        (
            line
            for line in lines
            if len(line) > 20
            and "Quotes to Scrape" not in line
        ),
        "No quotation could be extracted."
    )


print("\nFetched information:")
print(fetched_data)


# ============================================================
# STEP 5 - PREPARE EXCEL DATA
# ============================================================

print("\n" + "=" * 60)
print("STEP 5 - PREPARING REPORT DATA")
print("=" * 60)

headers = [
    "Date & Time",
    "Fetched Information",
    "Comment"
]

comment = "Daily update captured successfully."

row_data = [
    current_time,
    fetched_data,
    comment
]

print("Report data prepared.")


# ============================================================
# STEP 6 - OPEN EXCEL
# ============================================================

print("\n" + "=" * 60)
print("STEP 6 - OPENING MICROSOFT EXCEL")
print("=" * 60)

print("Starting Microsoft Excel 2024...")

# Close existing Excel processes
try:
    subprocess.run(
        ["taskkill", "/F", "/IM", "EXCEL.EXE"],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL
    )
except Exception:
    pass

wait(3)

# ------------------------------------------------------------
# Find Excel executable
# ------------------------------------------------------------

possible_excel_paths = [
    r"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
    r"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
    r"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE",
    r"C:\Program Files (x86)\Microsoft Office\Office16\EXCEL.EXE",
]

excel_exe = None

for path in possible_excel_paths:
    print("Checking:", path)

    if os.path.exists(path):
        excel_exe = path
        break


# ------------------------------------------------------------
# If Excel was found
# ------------------------------------------------------------

if excel_exe:

    print()
    print("Excel found:")
    print(excel_exe)
    print()
    print("Launching Excel...")

    subprocess.Popen([excel_exe])

else:

    print()
    print("Excel was not found in standard Office locations.")
    print()
    print("Trying Windows Excel launcher...")

    try:

        subprocess.Popen(
            ["cmd", "/c", "start", "", "excel"]
        )

    except Exception as e:

        fail(
            "Could not start Microsoft Excel.\n\n"
            "Please verify that Office Home 2024 is installed.\n\n"
            f"Details: {e}"
        )


wait(10, "Waiting for Microsoft Excel 2024...")


# ============================================================
# STEP 7 - CREATE BLANK WORKBOOK USING EXCEL COM
# ============================================================

print("\n" + "=" * 60)
print("STEP 7 - CREATING BLANK WORKBOOK")
print("=" * 60)

print("Connecting directly to Microsoft Excel...")

try:

    excel = win32com.client.Dispatch("Excel.Application")

except Exception as e:

    fail(
        "Excel started, but Python could not connect to "
        "Microsoft Excel through COM.\n\n"
        f"Details:\n{e}"
    )


print("Excel COM connection successful.")

excel.Visible = True
excel.DisplayAlerts = False

print("Creating new blank workbook...")

try:

    workbook = excel.Workbooks.Add()

except Exception as e:

    try:
        excel.Quit()
    except Exception:
        pass

    fail(
        "Excel opened, but a blank workbook could not be created.\n\n"
        f"Details:\n{e}"
    )


worksheet = workbook.Worksheets(1)

print("Blank workbook created successfully.")

wait(3)

# ============================================================
# STEP 8 - ENTER DATA DIRECTLY INTO EXCEL
# ============================================================

print("\n" + "=" * 60)
print("STEP 8 - ENTERING DATA")
print("=" * 60)

print("Writing report data into worksheet...")

# Header row
worksheet.Cells(1, 1).Value = headers[0]
worksheet.Cells(1, 2).Value = headers[1]
worksheet.Cells(1, 3).Value = headers[2]

# Data row
worksheet.Cells(2, 1).Value = row_data[0]
worksheet.Cells(2, 2).Value = row_data[1]
worksheet.Cells(2, 3).Value = row_data[2]

print("Data entered successfully.")


# ============================================================
# STEP 9 - FORMAT EXCEL
# ============================================================

print("\n" + "=" * 60)
print("STEP 9 - FORMATTING EXCEL")
print("=" * 60)

print("Auto-fitting columns...")

worksheet.Columns("A:C").AutoFit()

# Make headers bold
worksheet.Range("A1:C1").Font.Bold = True

# Freeze top row
try:
    worksheet.Application.ActiveWindow.SplitRow = 1
    worksheet.Application.ActiveWindow.FreezePanes = True
except Exception:
    pass

print("Columns auto-fitted.")

wait(2)


# ============================================================
# STEP 10 - SAVE EXCEL LOCALLY FIRST
# ============================================================

print("\n" + "=" * 60)
print("STEP 10 - SAVING EXCEL")
print("=" * 60)

print("Saving Excel workbook to a local temporary file first...")

# ------------------------------------------------------------
# IMPORTANT:
# Do NOT save directly to OneDrive.
# Save locally first, then copy the finished XLSX.
# ------------------------------------------------------------

temp_excel_path = os.path.join(
    os.environ.get("TEMP", r"C:\Temp"),
    f"daily_report_{today}_temp.xlsx"
)

print()
print("Temporary Excel file:")
print(temp_excel_path)

# Remove old temporary file if it exists
if os.path.exists(temp_excel_path):

    print("Removing old temporary file...")

    try:
        os.remove(temp_excel_path)
    except Exception as e:
        fail(
            "Could not remove old temporary Excel file.\n\n"
            f"{temp_excel_path}\n\n"
            f"Details: {e}"
        )


# ------------------------------------------------------------
# SAVE LOCALLY
# ------------------------------------------------------------

try:

    XLSX_FORMAT = 51

    print("Saving workbook locally...")

    workbook.SaveAs(
        Filename=temp_excel_path,
        FileFormat=XLSX_FORMAT
    )

    print("Local Excel SaveAs completed.")

except Exception as e:

    print()
    print("Excel local SaveAs failed.")
    print("Error:")
    print(e)

    try:
        workbook.Close(SaveChanges=False)
    except Exception:
        pass

    try:
        excel.Quit()
    except Exception:
        pass

    fail(
        "Excel could not save the temporary workbook.\n\n"
        f"Temporary path:\n{temp_excel_path}\n\n"
        f"Details:\n{e}"
    )


# ------------------------------------------------------------
# VERIFY TEMPORARY FILE
# ------------------------------------------------------------

print()
print("Checking temporary Excel file...")

for i in range(10):

    if os.path.exists(temp_excel_path):

        print(
            f"Temporary file detected after "
            f"{i + 1} second(s)."
        )

        break

    time.sleep(1)

else:

    try:
        workbook.Close(SaveChanges=False)
    except Exception:
        pass

    try:
        excel.Quit()
    except Exception:
        pass

    fail(
        "Excel reported that SaveAs completed, "
        "but the temporary file was not found.\n\n"
        f"Expected:\n{temp_excel_path}"
    )


temp_size = os.path.getsize(temp_excel_path)

print("Temporary file size:", temp_size, "bytes")

if temp_size == 0:

    try:
        workbook.Close(SaveChanges=False)
    except Exception:
        pass

    try:
        excel.Quit()
    except Exception:
        pass

    fail(
        "Temporary Excel file was created but is 0 bytes."
    )


# ============================================================
# STEP 11 - CLOSE EXCEL
# ============================================================

print("\n" + "=" * 60)
print("STEP 11 - CLOSING EXCEL")
print("=" * 60)

try:

    workbook.Close(SaveChanges=False)

except Exception as e:

    print("Workbook close warning:")
    print(e)


try:

    excel.Quit()

except Exception as e:

    print("Excel close warning:")
    print(e)


print("Excel closed.")

wait(3)


# ============================================================
# STEP 12 - COPY FILE TO ONEDRIVE
# ============================================================

print("\n" + "=" * 60)
print("STEP 12 - COPYING REPORT TO ONEDRIVE")
print("=" * 60)

print("Destination:")
print(excel_path)

# Make sure destination directory exists
if not os.path.exists(PROJECT_DIR):

    fail(
        "Destination directory does not exist:\n"
        f"{PROJECT_DIR}"
    )


# ------------------------------------------------------------
# Remove existing destination file
# ------------------------------------------------------------

if os.path.exists(excel_path):

    print("Existing destination file found.")

    try:

        os.remove(excel_path)

        print("Existing destination file removed.")

    except PermissionError:

        fail(
            "The existing Excel file is locked and cannot be "
            "replaced.\n\n"
            f"File:\n{excel_path}\n\n"
            "Close the file in Excel/OneDrive and run again."
        )

    except Exception as e:

        fail(
            "Could not remove existing destination file.\n\n"
            f"File:\n{excel_path}\n\n"
            f"Details:\n{e}"
        )


# ------------------------------------------------------------
# Copy temporary file to final location
# ------------------------------------------------------------

print("Copying temporary file to final location...")

try:

    import shutil

    shutil.copy2(
        temp_excel_path,
        excel_path
    )

except Exception as e:

    fail(
        "Could not copy the completed Excel file "
        "to the OneDrive folder.\n\n"
        f"Source:\n{temp_excel_path}\n\n"
        f"Destination:\n{excel_path}\n\n"
        f"Details:\n{e}"
    )


print("File copied successfully.")


# ============================================================
# STEP 13 - VERIFY FINAL EXCEL FILE
# ============================================================

print("\n" + "=" * 60)
print("STEP 13 - VERIFYING EXCEL FILE")
print("=" * 60)

print("Expected:")
print(excel_path)

# Wait for filesystem/OneDrive
for i in range(15):

    if os.path.exists(excel_path):

        print()
        print(
            f"Final file detected after "
            f"{i + 1} second(s)."
        )

        break

    time.sleep(1)

else:

    fail(
        "Final Excel file was not found after copying.\n\n"
        f"Expected:\n{excel_path}"
    )


file_size = os.path.getsize(excel_path)

if file_size == 0:

    fail(
        "Final Excel file exists but is 0 bytes.\n\n"
        f"File:\n{excel_path}"
    )


print()
print("SUCCESS!")
print()
print("Excel file saved successfully.")
print()
print("File:")
print(excel_path)
print()
print("File size:", file_size, "bytes")


# ============================================================
# STEP 14 - OPEN SAVED REPORT FOR SCREENSHOT
# ============================================================

print("\n" + "=" * 60)
print("STEP 14 - OPENING SAVED REPORT")
print("=" * 60)

print("Opening the actual saved Excel file...")

try:

    excel = win32com.client.Dispatch("Excel.Application")

    excel.Visible = True
    excel.DisplayAlerts = False

    saved_workbook = excel.Workbooks.Open(
        excel_path
    )

    saved_worksheet = saved_workbook.Worksheets(1)

    saved_workbook.Activate()
    saved_worksheet.Activate()

    # Select A1
    saved_worksheet.Range("A1").Select()

    # AutoFit
    saved_worksheet.Columns("A:C").AutoFit()

    # Bold header
    saved_worksheet.Range("A1:C1").Font.Bold = True

    # Zoom
    try:
        excel.ActiveWindow.Zoom = 100
    except Exception:
        pass

    # Maximize Excel
    try:
        excel.WindowState = -4137
    except Exception:
        pass

    print("Saved report opened successfully.")

except Exception as e:

    print()
    print("Could not reopen saved report for screenshot.")
    print(e)

    # The Excel file itself is already successfully saved.
    saved_workbook = None
    excel = None


# ============================================================
# STEP 15 - SCREENSHOT
# ============================================================

print("\n" + "=" * 60)
print("STEP 15 - CAPTURING REPORT SCREENSHOT")
print("=" * 60)

wait(5, "Preparing Excel report for screenshot...")

try:

    pyautogui.screenshot().save(
        screenshot_path
    )

    print()
    print("Report screenshot saved:")
    print(screenshot_path)

except Exception as e:

    print("Screenshot error:")
    print(e)


# ============================================================
# STEP 16 - CLOSE EXCEL
# ============================================================

print("\n" + "=" * 60)
print("STEP 16 - CLOSING EXCEL")
print("=" * 60)

if saved_workbook is not None:

    try:
        saved_workbook.Close(
            SaveChanges=False
        )
    except Exception as e:
        print("Workbook close warning:", e)


if excel is not None:

    try:
        excel.Quit()
    except Exception as e:
        print("Excel close warning:", e)


# ============================================================
# CLEAN TEMP FILE
# ============================================================

print("\nCleaning temporary file...")

try:

    if os.path.exists(temp_excel_path):

        os.remove(temp_excel_path)

        print("Temporary file removed.")

except Exception as e:

    print(
        "Temporary file cleanup warning:",
        e
    )


# ============================================================
# COMPLETION
# ============================================================

print("\n" + "=" * 60)
print("DAILY REPORT BOT COMPLETED")
print("=" * 60)

print()

print("Excel File:")
print(excel_path)

print()

print("Screenshot:")
print(screenshot_path)

print()

print("Everything completed successfully!")