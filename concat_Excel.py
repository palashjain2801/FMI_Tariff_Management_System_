import os
import glob
import pandas as pd
def edit_Excel():
    os.chdir(r'C:\Users\Palash\Desktop\FMI_Tariff_Management_System'+r'\FMI_DATA\excel_for_database')
    extension = 'xlsx'
    all_filenames = []
    colnames = ['Company Name', 'Pick-up(Country)', 'Pick-up(State)', 'Pick-up(District)','Pick-up(Port Name)','Pick-up-ZipCode','Destination(Country)','Destination(District)','Destination(State)','Destination(Port Name)','Destination-ZipCode','Miles','Base-Rate' ]
    print ("working")
    for i in glob.glob('*.{}'.format(extension)):
        print(i)
        all_filenames.append(i)
    print (all_filenames)

    # os.remove(r"C:\Users\Palash\Desktop\FMI_DATA\out\all_data.xlsx")
    join = pd.concat([pd.read_excel(f) for f in all_filenames ])
    # print(join)
    join.to_excel(r'C:\Users\Palash\Desktop\FMI_Tariff_Management_System'+r"\FMI_DATA\out\all_data.xlsx")

edit_Excel()