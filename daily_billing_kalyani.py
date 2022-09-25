import openpyxl as xl

def find_cell(sheet , max_row , max_col , heading):
    for i in range(1 , max_row + 1):
        for j in range(1 , max_col + 1):
            if sheet.cell(row = i , column = j).value == heading:
                return i , j
            
def process_files(files):
    for file in files:
        #result = {"MPR" : [0 , 0]  , "QPR" : [0 , 0]}
        result = {"A" : [0 , 0] , "I" : [0 , 0] , "OTH_MONTHLY" : [0 , 0]  , "QPR" : [0 , 0]}
        sheet = xl.load_workbook(file).active
        max_row , max_col = sheet.max_row , sheet.max_column
        mru_row , mru_col = find_cell(sheet , max_row , max_col , "Billing Portion")
        con_row , con_col = find_cell(sheet , max_row , max_col , "Live & Temp Disc Con")
        inv_row , inv_col = find_cell(sheet , max_row , max_col , "Invoice Generated")

        for i in range(1 , max_row + 1):
            mru = sheet.cell(row = i , column = mru_col).value
            if mru is not None and mru.strip()[-3:] in ('QPR' , 'MPR'):
                con = sheet.cell(row = i , column = con_col).value
                inv = sheet.cell(row = i , column = inv_col).value
                if con is None : con = 0   #where con count value is blank , consider it as ZERO
                if inv is None : inv = 0   #where invoice count value is blank , consider it as ZERO
                result[mru.strip()[-3:]][0] += con
                result[mru.strip()[-3:]][1] += inv
                
        print(sheet , result)
        
def main():
    print("Hello . Rename CCC wise excels as 103.xlsx , 201.xlsx etc.")
    bill_files = ['103.xlsx' , '201.xlsx' , '202.xlsx' , '208.xlsx' , \
             '300.xlsx' , '301.xlsx' , '401.xlsx' , '402.xlsx']
    mru_files = 'mru_wise_class.xlsx'
    process_files(bill_files)
    
if __name__ == '__main__':
    main()
