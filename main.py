#pip install gspread
#pip install oauth2client
#pip install openpyxl
import random
from flask import Flask, render_template
import pandas as pd
from datetime import datetime
# read the excel file in pandas
df = pd.read_excel ('KQnet_Data_Info.xlsx', sheet_name='Dummy Ratings')
df1 = df.loc[1:32, "#":"Reuse"]

# extract the sheet1
kq1 = df1.drop(columns=["Elective"])
kq1.to_csv('KQ_Data_1.csv',index=False)

# extract the sheets2
df2 = df.loc[df.Rating_Category_ID ==2 , :]
kq2 = df2[['#', 'User ID', 'Rating_Category_ID', 'Semester_year_id', 'Semester_id', 'Parent_id', 'Subject_id',
 "Created_at", "Updated_at", "Quality of Program", "School Support Systems", "Facilities", "Transportantion Options",
 "Library", "Book Store", "Cafeteria/Food Options","Job Prospects/Co-op","Overall", "Recommend?"]]

kq2.to_csv('KQ_Data_2.csv',index=False)

# importing services to save it into google drive
import gspread
from oauth2client.service_account import ServiceAccountCredentials

## file1
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(credentials)

spreadsheet = client.open('CSV-to-G-Sheet01')

with open('KQ_Data_1.csv', 'r') as file_obj:
    content = file_obj.read()
    client.import_csv(spreadsheet.id, data=content)


## file2
scope = ["https://spreadsheets.google.com/feeds", 'https://www.googleapis.com/auth/spreadsheets',
         "https://www.googleapis.com/auth/drive.file", "https://www.googleapis.com/auth/drive"]

credentials = ServiceAccountCredentials.from_json_keyfile_name('client_secret.json', scope)
client = gspread.authorize(credentials)

spreadsheet = client.open('CSV-to-G-Sheet02')

with open('KQ_Data_2.csv', 'r') as file_obj:
    content = file_obj.read()
    client.import_csv(spreadsheet.id, data=content)



# datetime object containing current date and time
now = datetime.now()
# dd/mm/YY H:M:S
date_time = now.strftime("%d/%m/%Y %H:%M:%S")
#print("date and time =", dt_string)	

# showing a the temporary host of a website

app = Flask(  
	__name__,
	template_folder='templates',  # Name of html file folder
	static_folder='static'  # Name of directory for static files
)

@app.route('/', methods=['GET', 'POST'])
def index():
  return render_template('index.html', date_time=date_time)

if __name__ == "__main__":  # Makes sure this is the main process
	app.run( # Starts the site
		host='0.0.0.0',  # EStablishes the host, required for repl to detect the site
		port=random.randint(2000, 9000),  # Randomly select the port the machine hosts on.
        #Debug=True # Run webapp in debug mode
	)