import time
import os
import glob
import pandas as pd
import xlwings as xw

from datetime import datetime as dt
from xlwings.constants import SortOrder

from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


# ====== MANAGERS PASSINGS DOWNLOAD LINKS
eccd = f"<REPORT LINK>"
passing = f"<REPORT LINK>"
design = f"<REPORT LINK>"
task = f"<REPORT LINK>"
closed = f"<REPORT LINK>"
presale = f"<REPORT LINK>"
bpart = f"<REPORT LINK>"
MPS_links = [eccd, passing, design, task, closed, presale, bpart]
start_time = time.time()

a
def sfDownload(export_url, export_name):

    download_dir = f"{os.getcwd()}"
    export_name = f"{export_name}.csv"
    edge_options = Options()
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True,
        "profile.default_content_setting_values.automatic_downloads": 1
    }
    edge_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Edge(options=edge_options)

    driver.get(export_url)
    WebDriverWait(driver, 30).until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Next']"))).click()

    time.sleep(30)
    downloaded = max(glob.glob(os.path.join(download_dir, "*")), key=os.path.getctime)

    if downloaded:
        new_path = os.path.join(download_dir, export_name)
        if os.path.exists(new_path): os.remove(new_path)
        os.rename(downloaded, new_path)
        print(f"Renamed to {new_path}")
    else:
        print("Download failed or timed out")
    
    driver.quit()


def delRep(csv_df, mainwb, sheet, clrCol):

    time.sleep(2)
    try: sheet.api.ShowAllData()
    except:pass
    time.sleep(2)
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    clear_range = sheet.range(f'A1:{clrCol}{last_row}')
    time.sleep(2)
    clear_range.clear_contents()
    sheet.range('A1').options(index=False, header=True).value = csv_df
    time.sleep(2)
    # Deletes P2 copyright crap at the end
    last_row = sheet.range('A' + str(sheet.cells.last_cell.row)).end('up').row
    start_delete = last_row - 4
    sheet.api.Rows(f"{start_delete}:{last_row}").Delete()


#============================= MAIN PASSINGS =================================================
print("MAIN PASSINGS REPORT")
#=============================================================================================

sfDownload(eccd, "eccd")
sfDownload(passing, "passing")

#============ ECCD PASSINGS
eccd_df = pd.read_csv('eccd.csv',encoding='windows-1252')
cols = eccd_df.columns[9]
eccd_df[cols] = pd.to_datetime(eccd_df[cols], errors='coerce').dt.normalize()
#eccd_df[cols] = eccd_df[cols].apply(lambda x: pd.to_datetime(x, errors='coerce').dt.date)
currentYear = 2026

statFilter = ["Cancelled", "Closed", "Pending Cancelled"]
eccd_df["Construction Milestone Complete Date"] = pd.to_datetime(eccd_df["Construction Milestone Complete Date"], errors="coerce").dt.normalize()
status_filter = (eccd_df["Status Picklist"].astype(str).str.strip().isin(statFilter))
older_than_currYear_filter = (eccd_df["Construction Milestone Complete Date"].notna() & (eccd_df["Construction Milestone Complete Date"].dt.year < currentYear-1))

eccd_df = eccd_df[~(status_filter & older_than_currYear_filter)].reset_index(drop=True)
eccd_df = eccd_df.sort_values(by='Construction Milestone Complete Date')
eccd_df = eccd_df.sort_values(by='Status Picklist')

eccd_df.to_csv('eccd_filtered.csv', index=False)
len(eccd_df)

#====== PASSINGS
passing_df = pd.read_csv('passing.csv',encoding='windows-1252')
cols = passing_df.columns[7]
passing_df[cols] = pd.to_datetime(passing_df[cols], errors='coerce').dt.normalize()
#passing_df[cols] = passing_df[cols].apply(lambda x: pd.to_datetime(x, errors='coerce').dt.date)
passing_df.to_csv('passing_filtered.csv', index=False)

#====== MAIN PASSINGS REPORT
app = xw.App()
app.display_alerts = False
mainwb = app.books.open('MAIN PASSINGS REPORT.xlsm')

#====== Lets you search for workbooks using the explorer
# Tk().withdraw()
# file_path = askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm;*.xlsb")])
# time.sleep(2)
# mainwb = xw.Book(file_path)

#====== Uses workbooks that are currently open
#mainwb = xw.apps.active.books.active

eccd_df = pd.read_csv('eccd_filtered.csv',encoding='windows-1252')
passing_df = pd.read_csv('passing_filtered.csv',encoding='windows-1252')
time.sleep(10)

#====== MAIN PASSINGS REPORT: ECCD PASSINGS
print("ECCD PASSINGS")
eccdSheet = mainwb.sheets['ECCD PASSINGS']
try: eccdSheet.api.ShowAllData()
except:pass

last_job_row = eccdSheet.range('A' + str(eccdSheet.cells.last_cell.row)).end('up').row
clear_range = eccdSheet.range(f'A3:M{last_job_row}')
time.sleep(2)
clear_range.clear_contents()
eccdSheet.range('A3').options(index=False, header=False).value = eccd_df
time.sleep(2)

last_formula_row = eccdSheet.range('N' + str(eccdSheet.cells.last_cell.row)).end('up').row
time.sleep(2)
last_row = eccdSheet.used_range.last_cell.row
fill_range = eccdSheet.range(f'N{last_formula_row}:AF{last_row}')

try:eccdSheet.range(f'N{last_formula_row}:AF{last_formula_row}').api.AutoFill(Destination=fill_range.api,Type=0)
except: print("Fill Failed at Eccd")

# Deletes P2 copyright crap at the end
last_row = eccdSheet.range('A' + str(eccdSheet.cells.last_cell.row)).end('up').row
start_delete = last_row - 4
eccdSheet.api.Rows(f"{start_delete}:{last_row}").Delete()

#====== MAIN PASSINGS REPORT: PASSING DATA
print("Passings")
passSheet = mainwb.sheets['PASSING DATA']
time.sleep(2)
try: passSheet.api.ShowAllData()
except:pass
time.sleep(2)

last_row = passSheet.range('A' + str(passSheet.cells.last_cell.row)).end('up').row
clear_range = passSheet.range(f'A2:H{last_row}')
time.sleep(2)
clear_range.clear_contents()
passSheet.range('A2').options(index=False, header=False).value = passing_df
time.sleep(2)

last_formula_row = passSheet.range('I' + str(passSheet.cells.last_cell.row)).end('up').row
last_row = passSheet.used_range.last_cell.row
fill_range = passSheet.range(f'I{last_formula_row}:P{last_row}')
try:passSheet.range(f'I{last_formula_row}:P{last_formula_row}').api.AutoFill(Destination=fill_range.api,Type=0)
except: print("Fill Failed at Passing Data")

# Deletes P2 copyright crap at the end
last_row = passSheet.range('A' + str(passSheet.cells.last_cell.row)).end('up').row
start_delete = last_row - 4
passSheet.api.Rows(f"{start_delete}:{last_row}").Delete()

last_row = passSheet.range("L" + str(passSheet.cells.last_cell.row)).end("up").row
data_range = passSheet.range(f"A1:P{last_row}")
col_L = passSheet.range(f"L1:L{last_row}")
print("FILTERING PASSING DATA this will take some time!!!")

passSheet.api.Sort.SortFields.Add(Key=col_L.api, SortOn=0, Order=SortOrder.xlAscending, DataOption=0)
passSheet.api.Sort.SetRange(data_range.api)
passSheet.api.Sort.Header = 1
passSheet.api.Sort.Apply()
passSheet.api.Range(data_range.address).AutoFilter(Field=12, Criteria1="#N/A")

time.sleep(2)
visible_rows = data_range.offset(1, 0).resize(data_range.rows.count - 1).api.SpecialCells(12)
visible_rows.EntireRow.Delete()

time.sleep(2)
try: passSheet.api.ShowAllData()
except:pass

app.api.Application.Run("'MAIN PASSINGS REPORT.xlsm'!UPDATEPIVOTS")

#============================= MANAGERS PASSINGS =============================================
print("MANAGERS PASSINGS REPORT")
#=============================================================================================

eccdSheet = mainwb.sheets['ECCD PASSINGS']
passSheet = mainwb.sheets['PASSING DATA']
app.display_alerts = False
manwb = app.books.open('MANAGER PASSINGS today.xlsb')
time.sleep(5)

try: eccdSheet.api.ShowAllData()
except:pass
eccd_last_row = eccdSheet.range("A" + str(eccdSheet.cells.last_cell.row)).end("up").row
eccd_data = eccdSheet.range(f"A3:AB{eccd_last_row}").value
eccd_source_range = eccdSheet.range(f"A3:AB{eccd_last_row}")
time.sleep(2)

try: passSheet.api.ShowAllData()
except:pass
pass_last_row = passSheet.range("A" + str(passSheet.cells.last_cell.row)).end("up").row
pass_data = passSheet.range(f"A3:H{pass_last_row}").value
pass_source_range = eccdSheet.range(f"A3:AB{pass_last_row}")
time.sleep(2)

    #====== ALL JOBS
print("All Jobs")
allSheet = manwb.sheets['ALL JOBS']
try: allSheet.api.ShowAllData()
except:pass

curr_last_row = allSheet.range("A" + str(eccdSheet.cells.last_cell.row)).end("up").row
clear_range = allSheet.range(f'A3:AB{curr_last_row}')
clear_range.clear_contents()
time.sleep(2)
#allSheet.range('A3').options(index=False, header=False).value = eccd_data

dest_range = allSheet.range("A3")
eccd_source_range.api.Copy()
dest_range.api.PasteSpecial(-4163)

    #====== JP MORGAN
print("JP Morgan")
jpSheet = manwb.sheets['JP MORGAN']
try: jpSheet.api.ShowAllData()
except:pass
time.sleep(2)
curr_last_row = jpSheet.range("A" + str(jpSheet.cells.last_cell.row)).end("up").row
clear_range = jpSheet.range(f'A3:AB{curr_last_row}')
clear_range.clear_contents()
time.sleep(2)
#jpSheet.range('A3').options(index=False, header=False).value = eccd_data
dest_range = jpSheet.range("A3")
eccd_source_range.api.Copy()
dest_range.api.PasteSpecial(-4163)

curr_last_row = jpSheet.range("A" + str(jpSheet.cells.last_cell.row)).end("up").row
data_range = jpSheet.range(f"A2:AB{curr_last_row}")
col_u = jpSheet.range(f"U3:U{curr_last_row}")
time.sleep(2)
data_range.api.AutoFilter(Field=21, Criteria1="<>JPMORGAN")
last_job_row = jpSheet.range("A" + str(jpSheet.cells.last_cell.row)).end("up").row
clear_range = jpSheet.range(f'A3:AB{last_job_row}')
clear_range.delete()
jpSheet.api.AutoFilterMode = False

    #====== Zero Passings
print("Zero Passings")
zpSheet = manwb.sheets['ZERO PASSINGS PAST DUE ECCD']
try: zpSheet.api.ShowAllData()
except:pass
time.sleep(2)
curr_last_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
clear_range = zpSheet.range(f'A3:AB{curr_last_row}')
clear_range.clear_contents()
time.sleep(2)
#zpSheet.range('A3').options(index=False, header=False).value = eccd_data
dest_range = zpSheet.range("A3")
eccd_source_range.api.Copy()
dest_range.api.PasteSpecial(-4163)

    #====== Column H
curr_last_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
data_range = zpSheet.range(f"A2:AB{curr_last_row}")
col_h = zpSheet.range(f"H3:H{curr_last_row}")
time.sleep(2)
data_range.api.AutoFilter(Field=8, Criteria1=["Cancelled", "Closed", "Pending Cancel"],Operator=7)
last_job_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
clear_range = zpSheet.range(f'A3:AB{last_job_row}')
time.sleep(2)
clear_range.delete()
zpSheet.api.AutoFilterMode = False

    #====== Column K
curr_last_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
col_k = zpSheet.range(f"K3:K{curr_last_row}")
time.sleep(2)
data_range.api.AutoFilter(Field=11,Criteria1=[f">{dt.today().strftime("%m/%d/%Y")}"])
last_job_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
clear_range = zpSheet.range(f'A3:AB{last_job_row}')
time.sleep(2)
clear_range.delete()
zpSheet.api.AutoFilterMode = False

curr_last_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
col_k = zpSheet.range(f"K3:K{curr_last_row}")
time.sleep(2)
data_range.api.AutoFilter(Field=11,Criteria1= "=")
last_job_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
clear_range = zpSheet.range(f'A3:AB{last_job_row}')
time.sleep(2)
clear_range.delete()
zpSheet.api.AutoFilterMode = False

    #====== Column X
curr_last_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
col_x = zpSheet.range(f"X3:X{curr_last_row}")
time.sleep(2)
data_range.api.AutoFilter(Field=24, Criteria1=">0")
last_job_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
clear_range = zpSheet.range(f'A3:AB{last_job_row}')
time.sleep(2)
clear_range.delete()
zpSheet.api.AutoFilterMode = False

    #====== Column Y
curr_last_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
col_y = zpSheet.range(f"Y3:Y{curr_last_row}")
time.sleep(2)
data_range.api.AutoFilter(Field=25, Criteria1="=0")
last_job_row = zpSheet.range("A" + str(zpSheet.cells.last_cell.row)).end("up").row
clear_range = zpSheet.range(f'A3:AB{last_job_row}')
time.sleep(2)
clear_range.delete()
zpSheet.api.AutoFilterMode = False

    #====== Address Release
print("Address Release")
arSheet = manwb.sheets['ADDRESS RELEASE']
try: arSheet.api.ShowAllData()
except:pass
time.sleep(2)
last_job_row = arSheet.range("A" + str(arSheet.cells.last_cell.row)).end("up").row
clear_range = arSheet.range(f'A2:H{last_job_row}')
clear_range.clear_contents()
time.sleep(2)
arSheet.range('A2').options(index=False, header=False).value = pass_data

last_formula_row = arSheet.range('I' + str(arSheet.cells.last_cell.row)).end('up').row
last_row = arSheet.used_range.last_cell.row
fill_range = arSheet.range(f'I{last_formula_row}:S{last_row}')
try:arSheet.range(f'I{last_formula_row}:S{last_formula_row}').api.AutoFill(Destination=fill_range.api,Type=0)
except: print("Fill Failed at Address Release")

elapsed = time.time() - start_time
minutes = int(elapsed // 60)
seconds = int(elapsed % 60)
print(f"Managers Passings Report Runtime: {minutes} min {seconds} sec\n Please Update yearly Summary and make sure Release Calandar filters work\n. ALSO make sure you save as todays date AND reduce file size!!!")


#===================== FINISH MAIN PASSINGS REPORT ============================================
print("Running rest of Main Passings Report")
#==============================================================================================

sfDownload(design, "design")
sfDownload(task, "task")
sfDownload(closed, "closed")
sfDownload(presale, "presale")
sfDownload(bpart, "bpart")

design_df = pd.read_csv('design.csv',encoding='windows-1252')
task_df = pd.read_csv('task.csv',encoding='windows-1252')
close_df = pd.read_csv('closed.csv',encoding='windows-1252')
presale_df = pd.read_csv('presale.csv',encoding='windows-1252')
bp_df = pd.read_csv('bpart.csv',encoding='windows-1252')
time.sleep(10)

    #====== DESIGN PASSING DATA
print("Design Passings")
designSheet = mainwb.sheets['DESIGN PASSINGS']
delRep(design_df, mainwb, designSheet, 'B')
time.sleep(2)

    #====== TASK DATA
print("Task Report")
taskSheet = mainwb.sheets['TASK REPORT']
delRep(task_df, mainwb, taskSheet, 'I')
time.sleep(2)

    #====== CLOSED DATE
print("Closed Date")
closeSheet = mainwb.sheets['CLOSED DATE']
delRep(close_df, mainwb, closeSheet, 'D')
time.sleep(2)

last_formula_row = closeSheet.range('E' + str(closeSheet.cells.last_cell.row)).end('up').row
time.sleep(2)
last_row = closeSheet.used_range.last_cell.row
fill_range = closeSheet.range(f'E{last_formula_row}:E{last_row}')
try: closeSheet.range(f'E{last_formula_row}:E{last_formula_row}').api.AutoFill(Destination=fill_range.api,Type=0)
except: print("Fill Failed at Closed Date")

    #====== PRE SALES PASSINGS
print("Pre Sales")
presaleSheet = mainwb.sheets['PRE SALES PASSINGS']
delRep(presale_df, mainwb, presaleSheet, 'C')
time.sleep(2)

    #====== JB BUSINESS PARTNER PASSINGS
print("Business Partner")
bpartSheet = mainwb.sheets['BUSINESS PARTNER']
delRep(bp_df, mainwb, bpartSheet, 'C')
time.sleep(2)

app.api.Application.Run(f"'{mainwb.name}'!DESIGNDUPLICATES")
time.sleep(2)
app.api.Application.Run(f"'{mainwb.name}'!UPDATEPIVOTS")

elapsed = time.time() - start_time
minutes = int(elapsed // 60)
seconds = int(elapsed % 60)
print(f"Managers Passings Report Runtime: {minutes} min {seconds} sec\n Finished")

# Use if the last part of the Report throws Errors
app = xw.App()
app.display_alerts = False
mainwb = app.books.open('MAIN PASSINGS REPORT.xlsm')
try: eccdSheet.api.ShowAllData()
except:pass

design_df = pd.read_csv('design.csv',encoding='windows-1252')
task_df = pd.read_csv('task.csv',encoding='windows-1252')
close_df = pd.read_csv('closed.csv',encoding='windows-1252')
presale_df = pd.read_csv('presale.csv',encoding='windows-1252')
bp_df = pd.read_csv('bpart.csv',encoding='windows-1252')
time.sleep(10)

    #====== DESIGN PASSING DATA
print("Design Passings")
designSheet = mainwb.sheets['DESIGN PASSINGS']
delRep(design_df, mainwb, designSheet, 'B')
time.sleep(2)

    #====== TASK DATA
print("Task Report")
taskSheet = mainwb.sheets['TASK REPORT']
delRep(task_df, mainwb, taskSheet, 'I')
time.sleep(2)

    #====== CLOSED DATE
print("Closed Date")
closeSheet = mainwb.sheets['CLOSED DATE']
delRep(close_df, mainwb, closeSheet, 'D')
time.sleep(2)

last_formula_row = closeSheet.range('E' + str(closeSheet.cells.last_cell.row)).end('up').row
time.sleep(2)
last_row = closeSheet.used_range.last_cell.row
fill_range = closeSheet.range(f'E{last_formula_row}:E{last_row}')
try: closeSheet.range(f'E{last_formula_row}:E{last_formula_row}').api.AutoFill(Destination=fill_range.api,Type=0)
except: print("Fill Failed at Closed Date")

    #====== PRE SALES PASSINGS
print("Pre Sales")
presaleSheet = mainwb.sheets['PRE SALES PASSINGS']
delRep(presale_df, mainwb, presaleSheet, 'C')
time.sleep(2)

    #====== JB BUSINESS PARTNER PASSINGS
print("Business Partner")
bpartSheet = mainwb.sheets['BUSINESS PARTNER']
delRep(bp_df, mainwb, bpartSheet, 'C')
time.sleep(2)

app.api.Application.Run(f"'{mainwb.name}'!DESIGNDUPLICATES")
time.sleep(2)
app.api.Application.Run(f"'{mainwb.name}'!UPDATEPIVOTS")

elapsed = time.time() - start_time
minutes = int(elapsed // 60)
seconds = int(elapsed % 60)
print(f"Managers Passings Report Runtime: {minutes} min {seconds} sec\n Finished")