import openpyxl as xl

def find_cell(sheet , max_row , max_col , heading):
    for i in range(1 , max_row + 1):
        for j in range(1 , max_col + 1):
            if sheet.cell(row = i , column = j).value == heading:
                return i , j
            
def process_files(files):
    result = {monthly : [0 , 0] , quarterly : [0 , 0]}

    for file in files:
        sheet = xl.load_workbook(file).active
        max_row , max_col = sheet.max_row , sheet.max_column
        mru_row , mru_col = find_cell(sheet , max_row , max_col , "Billing Portion")
        con_row , con_col = find_cell(sheet , max_row , max_col , "Live & Temp Disc Con")
        inv_row , inv_col = find_cell(sheet , max_row , max_col , "Invoice Generated")

        for i in range(1 , max_row):
            if sheet.cell(row = i , column = mru_col) is not None:
                mru = sheet.cell(row = i , column = mru_col)
                con = sheet.cell(row = i , column = con_col)
                inv = sheet.cell(row = i , column = inv_col)
                
        print(sheet , mru_col , con_col , inv_col)
        
def main():
    print("Hello . Rename CCC wise excels as 103.xlsx , 201.xlsx etc.")
    files = ['103.xlsx' , '201.xlsx' , '202.xlsx' , '208.xlsx' , \
             '300.xlsx' , '301.xlsx' , '401.xlsx' , '402.xlsx']
    process_files(files)
    
if __name__ == '__main__':
    main()
