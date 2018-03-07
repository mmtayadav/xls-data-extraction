import sys
import xlrd,xlwt
import datetime
import re
def open_file(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
 
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
    now = datetime.datetime.now()
    cur_yr,cur_mon  = now.year,now.month
    print first_sheet.ncols
    print first_sheet.nrows
    first_col_val = {}
    col_no = []
    row_no = 0
    sec_col_months  = []
    sec_col_days  = []
    latest_yr = ''
    latest_month = ''
    row_no = 0
    for row in range(1,first_sheet.nrows):
        first_col_val[first_sheet.cell(row,0)]=row

        print "first_sheet.cell(row,0).value" + str(first_sheet.cell(row,0).value)
        if str(first_sheet.cell(row,0).value) == "Memo:":
            col_no.append(row)
        if first_sheet.cell(row,0).value == cur_yr:
            col_no.append(row)
            latest_yr = cur_yr
        if first_sheet.cell(row,0).value == cur_yr-1:
            col_no.append(row)
            latest_yr = cur_yr
            
    print col_no
    for data in range(col_no[0],col_no[1]):
        if re.search(r'\d+',str(first_sheet.cell(data,1).value)):
            sec_col_days.append(first_sheet.cell(data,1).value)
        else:
            if str(first_sheet.cell(data,1).value) == "Year":
                row_no = int(data)-4
                break
            sec_col_months.append(first_sheet.cell(data,1).value)
    latest_month = sec_col_months[-1:]
    print latest_yr
    print latest_month
    print row_no
    data_header = ['Date','BCB_Commercial_Exports_Total','BCB_Commercial_Exports_Advances_on_Contracts','BCB_Commercial_Exports_Payment_Advance',\
                   'BCB_Commercial_Exports_Others','BCB_Commercial_Imports','BCB_Commercial_Balance','BCB_Financial_Purchases',\
                   'BCB_Financial_Sales','BCB_Financial_Balance','BCB_Balance']
    book = xlwt.Workbook()
    sh = book.add_sheet('sheet')
    row_count = 0
    
    for row in range(row_no,row_no+4):
        col_count = 0
        if row_count ==0:
            for header in data_header:
                sh.write(row_count,col_count,header)
                col_count+=1
        else:
            for col in range(1,first_sheet.ncols):
                if col ==1:
                    sh_date = str(latest_month[0]) + " " + str(int(first_sheet.cell(row,1).value)) + "," + str(cur_yr-1)                  
                    date = datetime.datetime.strptime(sh_date, '%b %d,%Y').strftime('%m/%d/%Y')
                    #datetime_object = datetime.strptime(date, '%b %d,%Y')
                    sh.write(row_count,col_count,date)
                else:
                    sh.write(row_count,col_count,first_sheet.cell(row,col).value)
                col_count+=1
        row_count+=1

    book.save("bcb_output_1.xls")

    
if __name__ == "__main__":

    #get_latest_file_date()
    path = sys.argv[1]
    open_file(path)
