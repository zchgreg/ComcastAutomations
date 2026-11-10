import tableauserverclient as TSC
import pandas as pd
import io

SERVER = <Sever>
OUTPUT_FILE = "chicago_funnel.csv"
OUTPUT_DIR = "csv_exports"
api_version = "3.25"

fullSite = "https://data.comcast.com/#/views/Funnel_Forecast/ConstructionFunnelDashboard?:iid=1"
TARGET_DASHBOARD = fullSite.split("views/")[1].split("/")[0].strip()
TARGET_DASHBOARD = TARGET_DASHBOARD.split('?', 1)[0]
SHEET = fullSite.split("views/")[1].split("/")[1].strip()
SHEET = SHEET.split('?', 1)[0]

# pat_name = input("Enter Tableau PAT name: ").strip()
# pat_secret = input("Enter Tableau PAT secret: ").strip()
pat_name = "ConstFunnelPAT"
pat_secret = <insertAPIkeyhere>
site = ""   # empty string means default site
tableau_auth = TSC.PersonalAccessTokenAuth(pat_name, pat_secret, site_id="")
server = TSC.Server(SERVER, use_server_version=True)
with server.auth.sign_in(tableau_auth):
    all_workbooks = list(TSC.Pager(server.workbooks))
    print(f"Total workbooks retrieved: {len(all_workbooks)}")

    found_workbook = None
    found_view = None


    for wb in all_workbooks:

        server.workbooks.populate_views(wb)
        if str(wb.content_url).strip() == TARGET_DASHBOARD:
            found_workbook = wb
        
        if found_workbook:
            server.workbooks.populate_views(found_workbook)
            for view in found_workbook.views:
                if view.content_url.split("/")[-1].strip() == SHEET:
                    found_workbook = wb
                    found_view = view
                    break
            if found_view: break

    if not found_workbook: raise Exception(f"Dashboard '{TARGET_DASHBOARD}' not found")
    if not found_view: raise Exception(f"Dashboard '{SHEET}' not found")

    if found_workbook:
        print("\n=== Dashboard Found ===")
        print("Dashboard Name:", TARGET_DASHBOARD)
        print("Workbook ID:", found_workbook.id)
        print("Worksheet Name:", found_view.name)

    server.views.populate_csv(found_view)
    csv_bytes = b"".join(found_view.csv)
    df_chicago = pd.read_csv(io.BytesIO(csv_bytes))
    try: df_chicago = df_chicago.loc[df_chicago["Region"] == "CHICAGO"]
    except: pass
    #df_chicago = df_chicago.loc[df_chicago["Measure Names"] == "Remaining Passings (Unagg)"]
    df_chicago.to_csv(f'{SHEET}.csv', index=False)
