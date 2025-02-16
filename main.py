import os
from urllib import request
from warnings import catch_warnings
import gspread
import datetime
from fastapi import FastAPI, Request
from starlette.middleware.cors import CORSMiddleware # CORS
from oauth2client.service_account import ServiceAccountCredentials

# CORS
app = FastAPI()
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"]
)

Auth = "./norse-lotus-423606-i2-353a26d9cd49.json"

os.environ['GOOGLE_APPLICATION_CREDENTIALS'] = Auth
scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
credentials = ServiceAccountCredentials.from_json_keyfile_name(Auth, scope)
Client = gspread.authorize(credentials)

SpreadSheet = Client.open_by_key("1sLIWIP7_7m8FOZm7dfsTudvYfDVHal-T11mzRG2TBsQ")

mainSheet = SpreadSheet.worksheet("企業データ一覧")

@app.get("/getcompnames")
async def get_comp_names():
    cols = mainSheet.col_values(5)
    return cols


def get_comp_data(compRow):
    header_data = mainSheet.row_values(1)
    comp_data = mainSheet.row_values(compRow)
    return_dict = dict()
    for h, c in zip(header_data,comp_data):
        return_dict[h] = c
    
    return return_dict

def get_comp_row(compName):
    objComp = mainSheet.find(compName,in_column=5)
    if objComp is None:
        return False
    else:
        objRow = objComp.row
        return objRow

@app.get("/getcompdata")
async def get_comp_datas(compName: str):
    comp_row = get_comp_row(compName)
    if comp_row != False:
        return get_comp_data(comp_row)
    else:
        return False

@app.post("/updatedata/")
async def update_data(request: Request):
    json_data = await request.json
    try:
        comp_name = json_data.get("company-name")
        comp_row = get_comp_row(comp_name)

        header_data = mainSheet.row_values(1)
        for key, val in json_data.items():
            print(key)
            if key in header_data:
                indx_of_header = header_data.index(key) + 1
                mainSheet.update_cell(comp_row,indx_of_header,val)
        return True
    except:
        print("ERR: Cannot find a company name")
        return False