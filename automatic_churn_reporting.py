# access chartMogul data
import chartmogul

# access and write in the google sheet
import gspread
from gspread_formatting import *
from oauth2client.service_account import ServiceAccountCredentials

# to get the date
from datetime import date, timedelta

# to get the state of the last transaction
import stripe

# to get the new access for helpscout and store them in a json file
import json
import requests

# this part allows python to connect to the google sheet
# the mail address in the json file must be authorized as admin in the sheet option
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name("automatic_churn_reporting/test-python-954dc90706df.json", scope)
client = gspread.authorize(creds)
sheet = client.open('Churn reporting').sheet1

# those credentials are the couple (account token, secret key)
config = chartmogul.Config('account_token', 'secret_key')

# allow to access stripe API
stripe.api_key = "stripe_api_key"

# the following part update helpscout credentials because the expire every two hours
################################################################################################

# this code was used to create data.json the file that contains the credentials
'''
data = {}
data['credentials'] = []
data['credentials'].append({
    'refresh_token': 'refresh_token',
    'access_token' : 'access_token',
    'client_id': 'client_id',
    'client_secret': 'client_secret'
})

with open('helpscout_credentials.json', 'w') as outfile:
    json.dump(data, outfile)
'''

# get the data entered in parameter from the data.json file
def getData(token):
    jsonFile = open("automatic_churn_reporting/helpscout_credentials.json", "r") # Open the JSON file for reading
    data = json.load(jsonFile) # Read the JSON into the buffer
    jsonFile.close() # Close the JSON file

    response = data["credentials"][0][token]

    return response

# replace refresh and access tokens in data.json
def updateJsonFile(new_refresh_token, new_access_token):
    jsonFile = open("automatic_churn_reporting/helpscout_credentials.json", "r") # Open the JSON file for reading
    data = json.load(jsonFile) # Read the JSON into the buffer
    jsonFile.close() # Close the JSON file

    ## Working with buffered content
    data["credentials"][0]["refresh_token"] = new_refresh_token
    data["credentials"][0]["access_token"] = new_access_token

    ## Save our changes to JSON file
    jsonFile = open("automatic_churn_reporting/helpscout_credentials.json", "w+")
    jsonFile.write(json.dumps(data))
    jsonFile.close()

def requestNewToken(previousToken):
    data = {
      'refresh_token': previousToken,
      'client_id': 'client_id',
      'client_secret': 'client_secret',
      'grant_type': 'refresh_token'
    }

    response = requests.post('https://api.helpscout.net/v2/oauth2/token', data=data).json()
    '''print(response)'''
    return([response["refresh_token"], response["access_token"]])

def updateHelpscoutAccess():
    try:
        new_credentials = requestNewToken(getData("refresh_token"))
        new_access_token = new_credentials[1]
        new_refresh_token = new_credentials[0]
    except (Exception):
        new_access_token = getData("access_token")
        new_refresh_token = getData("refresh_token")
        
    updateJsonFile(new_refresh_token, new_access_token)



################################################################################################


# return the content of the cancel mail 
def cancel_reason(email, access_token):
    headers = {
    'Authorization': 'Bearer '+access_token,
    }
    params = (('query', '(body:'+email+' AND subject:"User cancel feedback")'),('status', 'all'))
    response = requests.get('https://api.helpscout.net/v2/conversations', headers=headers, params=params)
    try:
        message = str(response.json()["_embedded"]["conversations"][0]["preview"])
        reason = message.split("for the following reason :")[1]
        return(reason.lower().lstrip(' '))
    except (Exception):
        return("")


do_not_need = ["need", "besoin"]

def analyse_reason(reason):
    if len(reason)<4 and len(reason)>0:
        return "N/A"
    if "need" in reason or "besoin" in reason or "necesit" in reason:
        return "Do not need"
    if "expensive" in reason or "cher" in reason:
        return "Expensive"
    if "pause" in reason:
        return "Site paused"
    if "clos" in reason:
        return "Website closed"
    return ("TOCHANGE " + reason)




def next_available_row(worksheet, col):
    str_list = list(filter(None, worksheet.col_values(col)))
    return len(str_list)


def fill_next_row(worksheet, tab, date, row):
    worksheet.update_cell(row, 1, date)
    worksheet.update_cell(row, 2, tab[0])
    worksheet.update_cell(row, 3, tab[1])
    worksheet.update_cell(row, 4, tab[2])
    worksheet.update_cell(row, 5, tab[3])
    worksheet.update_cell(row, 6, tab[4])
    worksheet.update_cell(row, 7, tab[5])
    worksheet.update_cell(row, 13, tab[6])

    # if there is an error somewhere higlight the whole line in red
    for element in tab:
        if element == "error":
            format = cellFormat(backgroundColor=color(1, 0.9, 0.9))
            format_cell_range(worksheet, 'A'+str(row)+':'+'M'+str(row), format)
    
    # if there is a cancel somewhere higlight the whole line in grey
    for element in tab:
        if element == "Cancel":
            format = cellFormat(backgroundColor=color(0.8, 0.8, 0.8))
            format_cell_range(worksheet, 'A'+str(row)+':'+'M'+str(row), format)


def get_cms(uuid):
    try:
        code = int(chartmogul.Customer.retrieve(config, uuid=uuid).get().attributes.stripe["cms"])
    except Exception:
        return "error"
    cms = ["Unknown", "WordPress", "Shopify", "BigCommerce", "Jimdo", "Squarespace", "Wix",
           "Laravel", "Symfony", "Weebly", "Drupal", "October CMS", "Other Technology",
           "Webflow", "Prestashop", "Magento"]
    try:
        return cms[code]
    except Exception:
        return "error"


# I make a first call to the API in order to know haw many pages of 200 clients we have
customer_list = chartmogul.Customer.all(config, status="cancelled", per_page=200).get()
pages = customer_list.total_pages

# the following part find yesterday's date and stores it month day and year to be used later

yesterday = date.today() - timedelta(days=1)
yesterday.strftime('%m%d%y')
date_tab = str(yesterday).split("-")
date = str(yesterday).replace("-", "/")
month = int(date_tab[1])
day = int(date_tab[2])
year = int(date_tab[0])


# enter the date here and uncomment to select a precise day
'''
date = "10/19/2019"
date_tab = date.split("/")
month = int(date_tab[0])
day = int(date_tab[1])
year = int(date_tab[2])
'''

# there will be as many steps as pages
print(pages)
step = 0

# get the row in which we have to start writing
row = next_available_row(sheet, 1)


updateHelpscoutAccess()
# get the access token for HelpScout API
access = getData("access_token")


# we look at all the clients page by page
while pages - step > 0:
    last_page = chartmogul.Customer.all(config, status="cancelled", per_page=200, page=pages - step).get()
    step += 1
    print(step)
    for i in range(0, len(last_page.entries)):
        # Getting the client's subscription
        uuid = last_page.entries[i].uuid
        entry = chartmogul.Subscription.all(config, uuid=uuid, per_page=200).get().entries.pop()
        # We check that client's last subscription end_date is the day we are looking for
        if entry.end_date.day == day and entry.end_date.month == month and entry.end_date.year == year:
            plan = entry.plan.split("_")
            try:
                nature = plan[0].capitalize()
            except Exception:
                nature = "error"
            try:
                frequency = plan[1].capitalize()
            except Exception:
                frequency = "error"

            # here we determine if the customer failed or cancelled
            try:
                state = stripe.Charge.list(limit=1, customer=last_page.entries[i].external_id)
            except Exception:
                state = "error"

            # final step : printing all the results in the sheet
            if state == "error":
                tab = [last_page.entries[i].external_id, nature, frequency, "error", "", last_page.entries[i].email, get_cms(uuid)]
                row += 1
                fill_next_row(sheet, tab, date, row)

            else:
                if state.data[0].status == "failed":
                    tab = [last_page.entries[i].external_id, nature, frequency, "Fail", "", last_page.entries[i].email, get_cms(uuid)]
                    row += 1
                    fill_next_row(sheet, tab, date, row)
                else:
                    reason = cancel_reason(last_page.entries[i].email, access)
                    reason_filtered = analyse_reason(reason)
                    tab = [last_page.entries[i].external_id, nature, frequency, "Cancel", reason_filtered, last_page.entries[i].email,get_cms(uuid)]
                    row += 1
                    fill_next_row(sheet, tab, date, row)
