

def process_files(files):
    for file in files:
        print(file)
        
def main():
    print("Hello . Rename CCC wise excels as 103.xlsx , 201.xlsx etc.")
    files = ['103.xlsx' , '201.xlsx' , '202.xlsx' , '208.xlsx' , \
             '300.xlsx' , '301.xlsx' , '401.xlsx' , '402.xlsx']
    process_files(files)
    
if __name__ == '__main__':
    main()
