import pypyodbc # i use this library to deal with sql server
import xlrd # for reading excel files


## connecting to the database 
database =pypyodbc.connect('DRIVER={SQL Server};SERVER=.\SQLEXPRESS;DATABASE=Cars_DB;')

## opening the excel file
book = xlrd.open_workbook('cars_data.xlsx')

## opening sheet one in the excel file
sheet = book.sheet_by_name("Sheet1")
#sheet = book.sheet_by_index(0)


## creating cursor for the database 
cursor = database.cursor()

## creating the insert query
query = """ INSERT INTO tbl_Car(Car_Num,Owner_Company,Branch,Service_Mode,Shaceh_Number,Motor_Number,Fuel_Type,Car_Type,Car_Model,Car_Load,Car_Weight,Shape,Color)"
                " VALUES(%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s))"""

## our script loop to iterate over the sheet rows
for i in range(1, sheet.nrows): #nrows = number of rows in the excel file
    Car_Num         = sheet.cell(r,).value
    Owner_Company   = sheet.cell(r,7).value
    Branch          = sheet.cell(r,1).value
    Service_Mode    = sheet.cell(r,).value
    Shaceh_Number   = sheet.cell(r,11).value
    Motor_Number    = sheet.cell(r,12).value
    Fuel_Type       = sheet.cell(r,).value
    Car_Type        = sheet.cell(r,).value
    Car_Model       = sheet.cell(r,10).value
    Car_Load        = sheet.cell(r,14).value
    Car_Weight      = sheet.cell(r,13).value
    Shape           = sheet.cell(r,9).value
    Color           = sheet.cell(r,15).value
    
    values = (Car_Num,Owner_Company,Branch,Service_Mode,Shaceh_Number,Motor_Number,Fuel_Type,Car_Type,Car_Model,Car_Load,Car_Weight,Shape,Color)

    cursor.execute(query , values)  ## executing the query

cursor.close()    # closing the connection
database.commit() # commiting the changes to the database
database.close()  # close the database 

    



