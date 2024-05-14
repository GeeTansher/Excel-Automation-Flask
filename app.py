from flask import Flask, render_template, request, Response, jsonify
import pandas as pd
import os
import time
import win32com.client as win32
import numpy as np
from pathlib import Path
import sys
import re
import pythoncom
import logging
from openpyxl.descriptors import (
    Convertible,
)

pythoncom.CoInitialize()
win32c = win32.constants
logging.basicConfig(level=logging.DEBUG)

app = Flask(__name__, static_url_path='/static')

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
master_file_path = Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'] + "\\" + "master.csv")
total_file_path = Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'] + "\\" + "total.csv")
final_file_path = Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'] + "\\" + "final.xlsx")

# List of registered users
users = [
    {'username': 'hello', 'password': 'hello'},
    {'username': 'user2', 'password': 'pass2'},
    # Add more users as needed
]

@app.route('/')
def login():
    return render_template('login.html')

@app.route('/login', methods=['POST'])
def check_login():
    username = request.form['username']
    password = request.form['password']

    for user in users:
        if user['username'] == username and user['password'] == password:
            return render_template('file_upload.html')

    return "Invalid username or password."

@app.route('/get_column_names', methods=['POST'])
def get_column_names():
    file1 = request.files['file1']
    df1 = pd.read_excel(file1)
    column_names = df1.columns.tolist()
    return jsonify({'column_names': column_names})

@app.route('/upload', methods=['POST'])
def process_files():
    file1 = request.files['file1']
    file2 = request.files['file2']
    
    
    # file1.save(Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'] + "\\" + file1.filename))
    # file2.save(Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'] + "\\" + file2.filename))
    file1.save(master_file_path)
    file2.save(total_file_path)
    
    
    
    # convert_xls_to_xlsx_for_master(Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'] + "\\" + file1.filename),
    #                             Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'] + "\\" + "master.xlsx"))
    
    # convert_xls_to_xlsx_for_total(Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'] + "\\" + file2.filename),
    #                             Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'] + "\\" + "total.xlsx"))
            
    process_data()

    return jsonify({'status': 'success'})
    
@app.route('/process')
def rendering_processed_data():
    return render_template('processed_data.html')

@app.route('/download')
def download_data():
    output_path = final_file_path

    def generate():
        with open(output_path, 'rb') as f:
            while True:
                data = f.read(1024)
                if not data:
                    break
                yield data
        # os.remove(output_path)
        for file in os.listdir(Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'])):
            os.remove(Path(os.getcwd() + "\\" + app.config['UPLOAD_FOLDER'] + "\\" + file))

    return Response(generate(), content_type='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


# def convert_xls_to_xlsx_for_master(inP, outP) -> None:
#     df = pd.read_excel(inP, sheet_name='MD_REPORT_0623_BD')
#     df.to_excel(outP, index=False, header=True)
    
# def convert_xls_to_xlsx_for_total(inP, outP) -> None:
#     df = pd.read_excel(inP)
#     df.to_excel(outP, index=False, header=True)
    
def process_data():
    # total data
    # df = pd.read_excel(total_file_path, skiprows=4)
    df = pd.read_csv(total_file_path, skiprows=4)
    df = df.drop(df.columns[[0, 1]], axis=1)
    df = df.dropna(how='all', axis=1)
    df = df.dropna(how='any', axis=0)
    df['Account No'] = df['Account No'].map(int)
    df["Sub Station"] = ""

    # master data
    df1 = pd.read_csv(master_file_path, usecols=["ACCOUNTNO","SUBSTATION"])
    df1["Account No"] = df1["ACCOUNTNO"]
    df1["Sub Station"] = df1["SUBSTATION"]
    df1 = df1.drop(df1.columns[[0,1]], axis=1)
    df1['Account No'] = df1['Account No'].map(int)
    
    # process
    Left_join = pd.merge(df, df1, on ='Account No', how ='left')
    Left_join["Sub Station"] = Left_join["Sub Station_y"]
    Left_join = Left_join.drop('Sub Station_x', axis=1)
    Left_join = Left_join.drop('Sub Station_y', axis=1)
    Left_join.to_excel(final_file_path, sheet_name = 'main', index = False, header=True)

    
    # sheet name for data
    sheet_name = 'main'
    
    # file path with file name
    # f_path = Path(r"C:\Data\Python and OpenCV\Excel pivot table automation\Website\uploads\final.xlsx")
    f_path = final_file_path
    
    # excel file
    # f_name = 'final.xlsx' # change to your Excel file name
    
    # function calls
    # run_excel(f_path, f_name, sheet_name)
    run_excel(f_path, sheet_name)
    

def pivot_table(wb: object, ws1: object, pt_ws: object, ws_name: str, pt_name: str, pt_rows: list, pt_cols: list, pt_filters: list, pt_fields: list):
    """
    wb = workbook1 reference
    ws1 = worksheet1 that contain the data
    pt_ws = pivot table worksheet number
    ws_name = pivot table worksheet name
    pt_name = name given to pivot table
    pt_rows, pt_cols, pt_filters, pt_fields: values selected for filling the pivot tables
    """

    # pivot table location
    pt_loc = len(pt_filters) + 2
    
    # grab the pivot table source data
    pc = wb.PivotCaches().Create(SourceType=win32c.xlDatabase, SourceData=ws1.UsedRange)
    
    # create the pivot table object
    pc.CreatePivotTable(TableDestination=f'{ws_name}!R{pt_loc}C1', TableName=pt_name)

    # selecte the pivot table work sheet and location to create the pivot table
    pt_ws.Select()
    pt_ws.Cells(pt_loc, 1).Select()

    # Sets the rows, columns and filters of the pivot table
    for field_list, field_r in ((pt_filters, win32c.xlPageField), 
                                (pt_rows, win32c.xlRowField),
                                (pt_cols, win32c.xlColumnField)):
        for i, value in enumerate(field_list):
            pt_ws.PivotTables(pt_name).PivotFields(value).Orientation = field_r
            pt_ws.PivotTables(pt_name).PivotFields(value).Position = i + 1

    # Sets the Values of the pivot table
    for field in pt_fields:
        pt_ws.PivotTables(pt_name).AddDataField(pt_ws.PivotTables(pt_name).PivotFields(field[0]), field[1], field[2]).NumberFormat = field[3]

    # Visiblity True or Valse
    pt_ws.PivotTables(pt_name).ShowValuesRow = True
    pt_ws.PivotTables(pt_name).ColumnGrand = True
    


def run_excel(f_path: Path, sheet_name: str):

    # filename = f_path / f_name
    filename = f_path
    # create excel object
    excel = win32.gencache.EnsureDispatch('Excel.Application', pythoncom.CoInitialize())

    # excel can be visible or not
    excel.Visible = True  # False
    
    # try except for file / path
    # wb = excel.Workbooks.Open(filename)
    try:
        wb = excel.Workbooks.Open(filename)
    except pythoncom.com_error as e:
        if e.excepinfo[5] == -2146827284:
            print(f'Failed to open spreadsheet.  Invalid filename or location: {filename}')
        else:
            raise e
        sys.exit(1)

    # set worksheet
    ws1 = wb.Sheets(sheet_name)
    
    # Setup and call pivot_table
    pivot_table_name = 'pivot_table' + ((str)(time.time())).split('.')[0]
    ws2_name = pivot_table_name
    wb.Sheets.Add().Name = ws2_name
    ws2 = wb.Sheets(ws2_name)
    
    # update the pt_name, pt_rows, pt_cols, pt_filters, pt_fields at your preference
    pt_name = pivot_table_name  # pivot table name, must be a string
    pt_rows = ['Sub Station']  # rows of pivot table, must be a list
    pt_cols = ['Collection Date']  # columns of pivot table, must be a list
    pt_filters = []  # filter to be applied on pivot table, must be a list
    # [0]: field name [1]: pivot table column name [3]: calulation method [4]: number format (explain the list item of pt_fields below)
    # pt_fields = [['Sale OfEC / ED', 'Sum of Sale OfEC / ED', win32c.xlSum, '0'],  # must be a list of lists
                #  ['Europe', 'Total Sales in Europe', win32c.xlSum, '0'],
                #  ['Japan', 'Total Sales in Japan', win32c.xlSum, '0'],
                #  ['Rest of World', 'Total Sales in Rest of World', win32c.xlSum, '0'],
                #  ['Global', 'Total Global Sales', win32c.xlSum, '0']]
                
    pt_fields = [['Sale Of\nEC / ED', 'Sum of Sale OfEC / ED', win32c.xlSum, '0']]
    # calculation method: xlAverage, xlSum, xlCount
    pivot_table(wb, ws1, ws2, ws2_name, pt_name, pt_rows, pt_cols, pt_filters, pt_fields)
    wb.Save() # save the pivot table created
    wb.Close(True)
    excel.Quit()
    excel = None
    del excel
    

if __name__ == '__main__':
    app.debug = True
    app.run(host='127.0.0.1', port=5000)