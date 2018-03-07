import sys
import xlrd,xlwt
import datetime
import re
import sys

def bcb_out2(path):
    """
    Open and read an Excel file
    """
    book = xlrd.open_workbook(path)
 
    # get the first worksheet
    first_sheet = book.sheet_by_index(0)
    now = datetime.datetime.now()
    cur_yr,cur_mon  = now.year,now.month
    print cur_yr
    print type(cur_yr) 
    col_no = []
    row_no = 0
    sec_col_months  = []
    latest_yr = ''
    latest_month = ''
    row_no = 0
    yr_row = 0

    
    for row in range(1,first_sheet.nrows):
        if row >=12 and isinstance(first_sheet.cell(row,0).value, float):
            if int(first_sheet.cell(row,0).value) == float(cur_yr):                
                latest_yr = str(cur_yr)
                yr_row = row
            elif int(first_sheet.cell(row,0).value) == float(cur_yr-1):
                print "check1"
                latest_yr = str(cur_yr)
                yr_row = row

            
    print yr_row
    for data in range(yr_row,first_sheet.nrows):
        if re.search(r'\w{3}',str(first_sheet.cell(data,1).value)):
            sec_col_months.append(first_sheet.cell(data,1).value)
        data_row = yr_row + len(sec_col_months)
        print "data_row" + str(data_row)
            
    latest_month = sec_col_months[-1:]
    print latest_yr
    print latest_month
    print data_row

    print "data" +  str(first_sheet.cell(data_row,2).value)
    data = str(first_sheet.cell(int(data_row)-1,2).value)

    book = xlwt.Workbook()
    sh = book.add_sheet('sheet')
    
    data_header = ['Date','BCB_FX_Position']
    for c,header in enumerate(data_header):
        sh.write(0,c,header)
        

    row_count = 0   

    sh_date = str(latest_month[0]) + " " + str(1) + "," + str(int(latest_yr)-1)                  
    date = datetime.datetime.strptime(sh_date, '%b %d,%Y').strftime('%m/%d/%Y')
    sh.write(1,0,date)
    sh.write(1,1,data)


    book.save("bcb_output_2.xls")
    

    
if __name__ == "__main__":

    path = sys.argv[1]
    bcb_out2(path)
