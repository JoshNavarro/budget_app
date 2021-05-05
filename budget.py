# An app to help you save yo money
# Created by Joshua Navarro
# Started 10/03/19

#!/usr/local/bin/python3

from datetime import datetime
from googleapiclient.discovery import build
from oauth2client.service_account import ServiceAccountCredentials
from operator import itemgetter
import os
import pandas
from pathlib import Path
from pprint import pprint
import re
from sortedcontainers import SortedDict

def get_spreadsheet_id(creds):
    '''
    Retrieve spreadsheet ID from google drive

    Args: creds=SpecialCredentials
    Return: spreadsheet ID
    '''
    service = build('drive', 'v3', credentials=creds)

    # Call the Drive v3 files.list API
    results = service.files().list(
        corpora='user',
        q="name contains 'Budget'",
        pageSize=10,
        spaces='drive', 
        fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])

    if not items:
        print('No files found.')
    else:
        for item in items:
            # # TESTING spreadsheet ID
            # if item['name'] == 'Testing':
            # REAL BUDGET spreadsheet ID
            if item['name'] == '2021 Budget':
                return item['id']

    # The spreadsheet to request.
    # Real Budget Sheet
    #spreadsheet_id = '1aakSKaeEc1Mo8cw1MGOrHzZdLvXp8RE4D8c7Ja-R01k'
    # Test budget sheet
    #spreadsheet_id = '1yRTLNA43UoU7vrWZEJ9PoK2ZH1AzxkqyBAqL--vuK9Y'

def get_sheet_id(creds,spreadsheet_id,sheet_name):
    """
    Retrieve google sheet id (spreadsheet tabs) based on tab name
    Args:   creds: credentials for account
            spreadsheet_id: google spreadsheet id
            sheet_name: sheet title to get id for
    Return: sheet id for sheet name
    """
    # Define sheets service
    service = build('sheets', 'v4', credentials=creds)
    
    # Attributes to retrieve spreadsheet data
    ranges = []
    include_grid_data = False

    # Call spreadsheets.get API
    request = service.spreadsheets().get(spreadsheetId=spreadsheet_id,ranges=ranges,includeGridData=include_grid_data)
    response = request.execute()

    # Set current sheet id
    current_sheet_id = None

    # Search all sheets in tab for matching sheet name
    for sheet in response['sheets']:
        if sheet['properties']['title'] == sheet_name:
            current_sheet_id = sheet['properties']['sheetId']
            break

    return current_sheet_id
    
def get_csv_files():
    """
    Search downloads folder for csv files
    Args: None
    Return: List of csv files sorted by date last modified
    """
    input_files = SortedDict()
    # Get path to home directory
    home = str(Path.home())
    # Create path to be searched from home (downloads)
    download_dir = '/Downloads/'
    # Add search path to home path to get FULL PATH
    path_to_search = home + download_dir
    path_to_search_str_length = len(path_to_search)
    # Search all files in PATH
    for file in os.listdir(path_to_search):
        # Match CSV files in PATH
        if file.lower().endswith(".csv"):
            # Full file path
            filename = path_to_search + file
            # Date last modified of file
            last_modified = os.stat(filename).st_mtime
            # Sorting files by date last modified using sorteddict
            if filename not in input_files:
                input_files[last_modified] = filename

    # Return sorted file paths
    return input_files.values()[:]

def csv_to_list(inputFile):
    """
    Convert input csv file into transactions list labeled Debit/Credit
    Args: input csv file
    Return: list of transactions
    """
    #inputFile = 'Statement closed Oct 16, 2019.CSV' # CITI TEST DATA
    #inputFile = 'export_20191018.csv' # PSCU TEST DATA (SMALL SET)
    #inputFile = 'export_20191002.csv' # PSCU TEST DATA (INCLUDES PENDING TRANSACTIONS)

    df = pandas.read_csv(inputFile)

    # csv file not a bank statement
    if 'Date' not in df.columns and 'Description' not in df.columns:
        return None

    # PSCU csv to list
    if 'Check Number' in df.columns:
        transactions = ['Debit',df[['Description','Comments','Date','Amount']].values.tolist()]

        # Remove pending transactions from list (TO ELIMINATE DUPLICATES LATER)
        while True:
            if transactions[1][0] and ('Pending' in transactions[1][0][0]):
                transactions[1].pop(0)
            else:
                break
        
        # Format amount from string to float (FOR INSERTING/READING TO GOOGLE SHEETS)
        for item in transactions[1]:
            # Limiting character count for transaction name
            #item[0] = item[0][:21]
            # Inserting placeholder for notes
            item[1] = ''
            # PSCU list negative amounts as ($xx.xx)
            if item[3][0] == '(':
                # Convert string to negative float
                item[3] = -(float(item[3].strip('($)').replace(',','')))
            else:
                # Convert string to positive float
                item[3] = float(item[3].strip('($)').replace(',',''))
            
    # CITI csv to list
    elif 'Member Name' in df.columns:
        df['Debit'] = df['Debit'].fillna(df['Credit'])
        transactions = df[['Description','Member Name','Date','Debit']].values.tolist()
        # Inserting a placeholder for notes
        for item in transactions:
            item[1] = ''
        # CITI lists transactions in DESC order BY date
        transactions.reverse()
        transactions = ['Credit', transactions]

    # CHASE csv to list
    elif 'Posting Date' in df.columns:
        transactions = df[['Description','Details','Posting Date','Amount']]
        # Inserting a placeholder for notes
        for item in transactions:
            item[1] = ''
        transactions = ['Debit', transactions]

    else:
        return None

    return transactions

def compare_sheet_data_to_csv_data(trans_type,transactions,spreadsheet_id,creds):
    """
    Edit transaction list to contain new data only
    Args:   transactions: list of transactions
            spreadsheet_id: spreadsheet id of google spreadsheet to read
            creds: credentials for google account
    """
    # Define sheets service
    service = build('sheets', 'v4', credentials=creds)
    # Read date from google sheet
    if trans_type == 'Debit':
        read_range = 'Debit!A2:D500'
    elif trans_type == 'Credit':
        read_range = 'Credit!A2:D500'

    # Value retrieves amount as number
    value_render_option = 'FORMULA'
    # Value retrieves date as serial number (days since 12/30/1899 as integer)
    date_time_render_option = 'SERIAL_NUMBER'
    # Calling spreadsheets.values.get api
    request = service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,range=read_range,valueRenderOption=value_render_option, dateTimeRenderOption=date_time_render_option)
    response = request.execute()

    # Saving data from response
    if response.get('values'):
        sheet_data = response['values']
    else:
        sheet_data = []

    # Converting date --> datetime then datetime --> excel date (TO COMPARE RESPONSE FROM GOOGLE API)
    for transaction in transactions:
        transaction[2] = datetime.strptime(transaction[2], '%m/%d/%Y').date()
        temp = datetime(1899, 12, 30).date()
        transaction[2] = (transaction[2] - temp).days

    ### DEBUG ###
    #   print(trans_list[2])
    #   print(sheet_data[26])

    # If csv element in sheet_data, remove first element of csv list
    while True:
        # nothing to read from range, add all transactions
        if len(sheet_data) == 0:
            break
        # Whole transaction list is already in the sheet
        if transactions[-1] in sheet_data:
            transactions = []
            break
        # Transaction list is not empty and transaction is already present in google sheet
        ##########   NOTE: COMPARISON HAPPENS HERE. PLEASE FINE TUNE COMPARISONS #############
        if transactions and (transactions[0] in sheet_data):
            transactions.pop(0)
        else:
            break
    # Loop until False (csv element not in sheet_data)
    return transactions

def csv_list_to_sheet(trans_type,transactions,spreadsheet_id,creds):
    """
    Insert transaction list to google spreadsheet
    Args:   transactions: list of transactions
            spreadsheet_id: spreadsheet id of google spreadsheet to update
            creds: credentials for google account 
    Return: Range of values updated
    """
    # Define sheets service
    service = build('sheets', 'v4', credentials=creds)
    # Range to enter new items
    if trans_type == 'Debit':
        range_ = 'Debit!A1:F6'
    elif trans_type == 'Credit':
        range_ = 'Credit!A1:F6'

    # Value to enter values as if doing it on google sheets UI
    value_input_option = 'USER_ENTERED'

    # Value to append and not overwrite
    insert_data_option = 'INSERT_ROWS'

    # Setting transaction list as values to be added to sheet
    value_range_body = {"values": transactions}

    # Calling spreadsheets.values.append api
    request = service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range_, valueInputOption=value_input_option, insertDataOption=insert_data_option, body=value_range_body)
    response = request.execute()

    return response['updates']['updatedRange']

def input_file_to_sheet(input_files,spreadsheet_id,creds):
    """
    Extract transactions from input files, edit transaction list for new data only, add data to google spreadsheet
    Args:   input_files: list of input csv files
            spreadsheet_id: spreadsheet id of google spreadsheet to update
            creds: credentials for google account 
    Return: Tuple of range of values updated
    """
    # Create master debit/credit transaction lists
    debit_transactions = []
    credit_transactions = []

    # get transactions list for each input file
    for file in input_files:
        transactions = csv_to_list(file)

        # transactions don't exist for input file
        if not transactions:
            continue

        # populate master debit transactions list
        if transactions[0] == 'Debit':
            debit_transactions += transactions[1]
        # populate master credit transactions list
        elif transactions[0] == 'Credit':
            credit_transactions += transactions[1]

    # sort master debit/credit transaction lists
    debit_transactions.sort(key=itemgetter(2))
    credit_transactions.sort(key=itemgetter(2))

    for transaction in debit_transactions:
        transaction[0] = extract_useful_string(transaction[0])
    #   print(transaction[0])
    for transaction in credit_transactions:
        transaction[0] = extract_useful_string(transaction[0])

    debit_range = None
    credit_range = None

    # Edit debit transaction list to contain only new data, add list to google spreadsheet
    debit_transactions = compare_sheet_data_to_csv_data('Debit',debit_transactions,spreadsheet_id,creds)
    if(len(debit_transactions) != 0):
        debit_range = csv_list_to_sheet('Debit',debit_transactions,spreadsheet_id,creds)
        #pass

    # Edit credit transaction list to contain only new data, add list to google spreadsheet
    credit_transactions = compare_sheet_data_to_csv_data('Credit',credit_transactions,spreadsheet_id,creds)
    if(len(credit_transactions) != 0):
        credit_range = csv_list_to_sheet('Credit',credit_transactions,spreadsheet_id,creds)
        #pass

    return (debit_range,credit_range)

def extract_useful_string(transaction):
    # Cut string after first special character except (,.'*&/)
    # Cut string after patterns with letters&numbers i.e. xxx478, F1567, 12AM
    # if string starts with non letter, keep first pattern
    # PSCU specific: cut string after "CO:", ZEL *(KEEP 2 WORDS AFTER)
    # Exceptions: BIG 5 SPORTING GOODS, To Share 00 WORDS

    # Regular expression to match above criteria
    pattern = re.compile(r'([^-\w.,\'*&/# ]|, | [a-zA-Z]*\d| #\d*| \d+)', re.I)

    # Finds the matches of above pattern in each transaction
    matches = pattern.search(transaction)

    # Return full transaction name if match not found
    if matches == None:
        return transaction

    # Return the string up to the first match
    return transaction[:matches.span()[0]]
    
def update_balance_column(updated_range,spreadsheet_id,creds):
    """
    Update the balance column of newly appended transactions
    Args:   updated_range: range of newly appended transactions
            creds: credentials to perform api call
    Return: None
    """
    (debit_range,credit_range) = updated_range

    # Nothing to be updated
    if not (debit_range or credit_range):
        return

    debit_balance_values = []
    credit_balance_values = []

    service = build('sheets','v4',credentials=creds)

    # If debit sheet updated, get range for debit sheet to be updated
    if debit_range:
        debit_range = debit_range.replace("A","E")
        debit_range = debit_range.replace(":D",":E")
        debit_range_values = re.findall("\d+",debit_range)
        # Generate list of formulas to be applied to sheet
        for i in range(int(debit_range_values[0]),int(debit_range_values[1])+1):
            debit_balance_values.append('=E'+str(i-1)+'+D'+str(i))

    # If credit sheet updated, get range for credit sheet to be updated
    if credit_range:
        credit_range = credit_range.replace("A","E")
        credit_range = credit_range.replace(":D",":E")
        credit_range_values = re.findall("\d+",credit_range)
        for i in range(int(credit_range_values[0]),int(credit_range_values[1])+1):
            credit_balance_values.append('=E'+str(i-1)+'+D'+str(i))

    # Determine request body data depending on data updated
    if debit_range and credit_range:
        data = [
            {
                'range':debit_range,
                'majorDimension':'COLUMNS',
                'values':[debit_balance_values]
            },
            {
                'range':credit_range,
                'majorDimension':'COLUMNS',
                'values':[credit_balance_values]
            }
        ]
    elif debit_range:
        data = [
            {
                'range':debit_range,
                'majorDimension':'COLUMNS',
                'values':[debit_balance_values]
            }
        ]
    elif credit_range:
        data = [
            {
                'range':credit_range,
                'majorDimension':'COLUMNS',
                'values':[credit_balance_values]
            }
        ]
    # Request body
    batch_update_values_request_body = {
        # How the data should be interprested
        'value_input_option': 'USER_ENTERED',
        'data': data
    }

    # Calling spreadsheets.values.batchUpdate api
    request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=batch_update_values_request_body)
    response = request.execute()

def categorize(updated_range,spreadsheet_id,creds):
    """
    Add a category to all new transactions added based off of categories for previous transacations
    Args:   updated_range: range of new transactions to add categories to
            spreadsheet_id: id of spreadsheet to be updated
            creds: credentials to perform api call
    Return: None
    """
    # Setting debit/credit range to batch get
    (debit_range,credit_range) = updated_range

    # Nothing to be updated
    if not (debit_range or credit_range):
        return

    if debit_range:
        debit_range_values = re.findall('\d+',debit_range)
        debit_range = 'Debit!A2:F' + debit_range_values[1]
        debit_categories = []
        ranges = [debit_range]

    if credit_range:
        credit_range_values = re.findall('\d+',credit_range)
        credit_range = 'Credit!A2:F' + credit_range_values[1]
        credit_categories = []
        ranges = [credit_range]

    if debit_range and credit_range:
        ranges = [debit_range,credit_range]
    
    
    # Create dictionary
    categories = {}

    # Define sheets service
    service = build('sheets', 'v4', credentials=creds)
    # Value retrieves amount as number
    value_render_option = 'FORMULA'
    # Value retrieves date as serial number (days since 12/30/1899 as integer)
    date_time_render_option = 'SERIAL_NUMBER'

    major_dimension = 'ROWS'

    # Calling spreadsheets.values.get api
    request = service.spreadsheets().values().batchGet(spreadsheetId=spreadsheet_id,ranges=ranges,valueRenderOption=value_render_option, dateTimeRenderOption=date_time_render_option, majorDimension=major_dimension)
    response = request.execute()

    # Populate dictionary with existing categories
    for values in response['valueRanges']:
        for value in values['values']:
            try:
                categories[value[0]] = value[5]
            except IndexError:
                pass

    # Resetting debit/credit range to update range to batch update
    (debit_range,credit_range) = updated_range

    # Update newest data category column
    if debit_range and credit_range:
        # Update range to batch update (Column F = Categories column)
        debit_range = debit_range.replace("A","F")
        debit_range = debit_range.replace(":D",":F")
        credit_range = credit_range.replace("A","F")
        credit_range = credit_range.replace(":D",":F")

        # Update category list values of newly appended debit transactions
        for value in response['valueRanges'][0]['values'][-(int(debit_range_values[1])-int(debit_range_values[0]))-1:]:
            # Update list to corresponding transaction category if it exists
            if(value[0] in categories):
                debit_categories.append(categories[value[0]])
            else:
                debit_categories.append('')

        # Update category list values of newly appended credit transactions
        for value in response['valueRanges'][1]['values'][-(int(credit_range_values[1])-int(credit_range_values[0]))-1:]:
            if(value[0] in categories):
                credit_categories.append(categories[value[0]])
            else:
                credit_categories.append('')

        # Update data to send to batch update API
        data = [
            {
                'range':debit_range,
                'majorDimension':'COLUMNS',
                'values':[debit_categories]
            },
            {
                'range':credit_range,
                'majorDimension':'COLUMNS',
                'values':[credit_categories]
            }
        ]

    elif debit_range:
        # Update range to batch update (Column F = Categories column)
        debit_range = debit_range.replace("A","F")
        debit_range = debit_range.replace(":D",":F")
        
        # Update category list values of newly appended debit transactions
        for value in response['valueRanges'][0]['values'][-(int(debit_range_values[1])-int(debit_range_values[0]))-1:]:
            if(value[0] in categories):
                debit_categories.append(categories[value[0]])
            else:
                debit_categories.append('')

        data = [
            {
                'range':debit_range,
                'majorDimension':'COLUMNS',
                'values':[debit_categories]
            }
        ]

    elif credit_range:
        # Update range to batch update (Column F = Categories column)
        credit_range = credit_range.replace("A","F")
        credit_range = credit_range.replace(":D",":F")
        
        # Update category list values of newly appended credit transactions
        for value in response['valueRanges'][0]['values'][-(int(credit_range_values[1])-int(credit_range_values[0]))-1:]:
            if(value[0] in categories):
                credit_categories.append(categories[value[0]])
            else:
                credit_categories.append('')

        data = [
            {
                'range':credit_range,
                'majorDimension':'COLUMNS',
                'values':[credit_categories]
            }
        ]

    # Request body
    batch_update_values_request_body = {
        # How the data should be interprested
        'value_input_option': 'USER_ENTERED',
        'data': data
    }

    # Calling spreadsheets.values.batchUpdate api
    request = service.spreadsheets().values().batchUpdate(spreadsheetId=spreadsheet_id, body=batch_update_values_request_body)
    response = request.execute()

def clean_old_csv_files():
    os.system('find ~/Downloads/ -type f -mtime +5 -iname "*csv" -delete')
    print('Cleaned csv files from ~/Downloads/ older than 5 days')

def open_google_sheet(spreadsheet_id):
    os.system('open -a /Applications/Safari.app https://docs.google.com/spreadsheets/d/'+spreadsheet_id+'/edit#gid=0')

# MAIN PROGRAM
def main():
    scope = [   "https://spreadsheets.google.com/feeds",
                'https://www.googleapis.com/auth/spreadsheets',
                "https://www.googleapis.com/auth/drive.file",
                "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_name("/Users/joshuanavarro/Dropbox/PythonProjects/Budget_App/creds.json", scope)
    clean_old_csv_files()
    spreadsheet_id = get_spreadsheet_id(creds)
    #sheet_id = get_sheet_id(creds,spreadsheet_id,'Debit')
    input_files = get_csv_files()
    updated_range = input_file_to_sheet(input_files,spreadsheet_id,creds)
    update_balance_column(updated_range,spreadsheet_id,creds)
    categorize(updated_range,spreadsheet_id,creds)
    open_google_sheet(spreadsheet_id)

if __name__ == '__main__':
    main()
