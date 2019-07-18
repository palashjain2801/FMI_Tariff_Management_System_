import pyodbc
from prettytable import PrettyTable
import xlrd
from xlrd import open_workbook, cellname

def  Databasework(source,destination):
    # ******* Connection with SQL Server ***********
    server = 'DESKTOP-K0J9CFC\SQLSERVER'
    database = 'FMIDB'
    username = 'myusername'
    password = 'mypassword'
    # t = PrettyTable(['Company Name', 'Pick-up(Country)', 'Pick-up(State)', 'Pick-up(District)','Pick-up(Port Name)','Pick-up-ZipCode','Destination(Country)','Destination(District)','Destination(State)','Destination(Port Name)','Destination-ZipCode','Miles','Base-Rate'])
    # t = PrettyTable(['Com', 'Pic1', 'Pic2', 'Pic3','Pic4','Pic6','Des1','Des2','Des3','Des4','Des5','Miles','Ba'])
    # t = PrettyTable(['Company', 'Pick-up(Country)', 'P_St', 'P_District','P_PortName','PZipCode','Dest(Country)','D(St)','D(District)','DPort Name','DZipCode','Miles','BaseRate'])
    t = PrettyTable(['CompanyName', 'Pick-up(Country)', 'PState', 'PDistrict','PPort Name','PZip','Destination(Country)','DState','DDistrict','DPort Name','DZip','Miles','BaseRate'])

    cnxn = pyodbc.connect('DRIVER={ODBC Driver 17 for SQL Server};SERVER='+server+';DATABASE='+database+'; Trusted_Connection=yes;')
    cursor = cnxn.cursor()

    # *******Connection with SQL Server Finished **************

    # ****** Running Query to Create Table in the Proper Format
    # cursor.execute("Drop Table TestTable1;")
    cursor.execute("Drop Table TestTable1;")
    cursor.execute("CREATE TABLE TestTable1(CompanyName  varchar(40),PickupCountry varchar (25),PickupState varchar(25) ,PickupDistrict varchar(25), PickupPortName varchar(25),PickupZipCode varchar(25),DestinationCountry varchar(25),DestinationState varchar(25),DestinationDistrict varchar(25),DestinationPortName varchar(25),DestinationZipCode varchar(25) , Miles varchar(25), BaseRate varchar(25));")
    # cursor.execute("INSERT INTO TestTable1 VALUES ('FMI','USA','LA','New Orleans','NILL','123123','USA','AL','Pennington','something', '987','950','123');")
    #cursor.execute("select * from dbo.combined_csv where Source = 'Seattle' and Destination = 'Aberdeen';")
    print("Table Created")


    # ********Connecting to Excel  ***********
    book = xlrd.open_workbook(r'C:\Users\Palash\Desktop\FMI_Tariff_Management_System'+r"\FMI_DATA\out\all_data.xlsx")
    sheet = book.sheet_by_index(0)
    print('inserting in number of rows:',sheet.nrows)  
    print('inserting in number of columns:',sheet.ncols)
    print(sheet.cell_value(1,0))
    # query = """INSERT INTO TestTable1  VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s.%s)"""
    query = """INSERT INTO TestTable1  VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"""


    for row in range(1 , sheet.nrows):
        # for col in range(1 , sheet.ncols):
        #     print(sheet.cell_value(row,col))
        #     print(type(str(sheet.cell_value(row,col))))
        a=str(sheet.cell_value(row,1))
        b=str(sheet.cell_value(row,2))
        c=str(sheet.cell_value(row,3))
        d=str(sheet.cell_value(row,4))
        e=str(sheet.cell_value(row,5))
        f=str(sheet.cell_value(row,6)) 
        g=str(sheet.cell_value(row,7))
        h=str(sheet.cell_value(row,8))
        i=str(sheet.cell_value(row,9))
        j=str(sheet.cell_value(row,10))
        k=str(sheet.cell_value(row,11))
        l=str(sheet.cell_value(row,12))
        m=str(sheet.cell_value(row,13))
        # print(row)

        value = (a,b,c,d,e,f,g,h,i,j,k,l,m)
        cursor.execute(query,value)
    
    pickup = source
    drop = destination 
    # DestinationDistrict='ALBANY'
    # cursor.execute("select * from TestTable1 where DestinationDistrict='ALBANY' ;")
    if pickup != "NONE" and drop != "NONE":
        print("select * from TestTable1 where "+pickup+"' and " + drop+ " ;")
        a = "select * from TestTable1 where "+pickup+" and " + drop+ " ;"
    elif pickup == "NONE":
        print("select * from TestTable1 where ",drop,";")
        a = "select * from TestTable1 where "+ drop+ " ;"
    elif drop == "NONE":
        print("select * from TestTable1 where ",pickup,";")
        a = "select * from TestTable1 where "+ pickup+ " ;"
    print (a)
    cursor.execute(a)

    # cursor.execute("select * from TestTable1 ;")

    for row in cursor:
        t.add_row(row)
    # print (t)
    #cursor.commit()
    # ********End : Connecting to Excel  ***********
    print("finished")
    return t 


# Databasework("PickupCountry='USA'","DestinationDistrict='ALBANY'")
# Databasework("PickupDistrict='Los Angeles'" , "DestinationDistrict='Chandler'")