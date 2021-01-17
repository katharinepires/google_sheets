# Using Python to get modifications in a Google Spreadsheet

You have two ways to do this. Here I will explain both but in the Notebook is only one way.

Fist you have to install the dependences:

`pip install gspread oauth2client` 

After yoy have to create a credential on Google to manage spreadsheets. To do this you have go to https://console.developers.google.com/project and click on 'Create Project', fill in the project name.
Then go to 'API's and services' -> 'Credentials' -> Create credentials. After creating the credential, fill in the 'Service account name' field and then the role for that account, I selected 'Project -> Owner' below the 'Key type' option, select `.JSON` and click 'Create' , a file will appear to be downloaded, save this file in the directory where the script will be.
Finally, you will need to activate the Google Sheets API. After that, you can easily go to Python.

Now with the spreadsheet, empty or not, let's go to the steps for access:

`import gspread` 

`from oauth2client.service_account import ServiceAccountCredentials`

`escopo = ['https://spreadsheets.google.com/feeds']`

`cred = ServiceAccountCredentials.from_json_keyfile_name('yourfile.json', escopo)`

`gc = gspread.authorize(cred)`

`wks = gc.open_by_key('KEY')`

`worksheet = wks.get_worksheet(0)` 

In place of KEY you can put both the name of the file, for that you will have to get the customer's email and send the spreadsheet to that emiail is in the `.JSON` file that was downloaded, just open, copy and paste when sharing the spreadsheet. The other option is to use the id:

![=planilha](https://user-images.githubusercontent.com/67076633/104828743-3e8ab000-584b-11eb-9272-ea9a9e95f2c4.png)

Test spreadsheet, using the customer's email in the file, just press share and use the email. To use the key, just take the link, copy and paste it into some notepad and get the key contained in that link:

![key](https://user-images.githubusercontent.com/67076633/104828759-798ce380-584b-11eb-8f56-aa35b1714b9d.png)

On the notebook that comes with this repository, there are some codes to modify the spreadsheet !!
