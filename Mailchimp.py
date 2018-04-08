import httplib2
import os
import json
import facebook
import requests
import webbrowser
import datetime

from apiclient import discovery
from oauth2client import client
from oauth2client import tools
from oauth2client.file import Storage

try:
    import argparse
    flags = argparse.ArgumentParser(parents=[tools.argparser]).parse_args()
except ImportError:
    flags = None

# If modifying these scopes, delete your previously saved credentials
# at ~/.credentials/sheets.googleapis.com-python-quickstart.json
# rm ~/.credentials/sheets.googleapis.com-python-quickstart.json
SCOPES = 'https://www.googleapis.com/auth/spreadsheets.readonly'
CLIENT_SECRET_FILE = 'client_secret.json'
APPLICATION_NAME = 'Mailchimp Email Generator'
FACEBOOK_TOKEN = '' #put your api key here
MAILCHIMP_API_KEY = '' #put your api key here
MAILCHIMP_ENDPOINT_URL = 'https://us11.api.mailchimp.com/3.0'
BIZTECH_FBPAGE_ID = '933982143299464'

def get_credentials():
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python-quickstart.json')

    store = Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(CLIENT_SECRET_FILE, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials

def readSheets(sheetURL):
    credentials = get_credentials()
    http = credentials.authorize(httplib2.Http())
    discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                    'version=v4')
    service = discovery.build('sheets', 'v4', http=http,
                              discoveryServiceUrl=discoveryUrl)

    spreadsheetId = sheetURL
    rangeName = 'Form Responses 1!A1:E'
    result = service.spreadsheets().values().get(
        spreadsheetId=spreadsheetId, range=rangeName).execute()
    values = result.get('values', [])
    if not values:
        print('No data found.')
    else:
        print(len(values))
        return values

def getFacebook(endpoint):
    graph = facebook.GraphAPI(access_token=FACEBOOK_TOKEN, version="2.7")
    return graph.request(endpoint)
    #page = graph.request('/185834998131645?fields=description')

def buildSubscriberJson(name, email):
    data = {}
    data['email_address'] = email
    data['status'] = 'subscribed'
    data['merge_fields'] = {'FNAME':name}
    return json.dumps(data)

def buildMembers(values):
    nameIndex = values[0].index("First Name")
    emailIndex = values[0].index("Email Address")
    membersString = ''
    for row in values[1:]:
        if ((len(row)>2) and (row[nameIndex] is not None) and (row[emailIndex] is not None)):
            membersString = membersString + ',' + buildSubscriberJson(row[nameIndex],row[emailIndex])
    return '{"members": [' + membersString[1:] + '], "update_existing": true}'

def mailChimpAddTemplate(eventId,eventInfo,campaignsId):
    eventCoverPhoto = eventInfo['cover']['source']
    eventName = eventInfo['name']
    url = MAILCHIMP_ENDPOINT_URL + '/campaigns/' + campaignsId + '/content'
    # place emailtemplate.txt in the same folder and have it contain an HTML version of your email template
    # you can dynamically replace fields using replace
    file = open('emailtemplate.txt','r').read().replace('+eventId+',eventId).replace('+eventCoverPhoto+',eventCoverPhoto).replace('+eventName+',eventName)
    requestParam ='{"html":"'+ file +'"}'
    
    r = requests.put(url,auth=('user',MAILCHIMP_API_KEY),data=requestParam)

def mailChimpCampaignCreateRequestParam(subject,listid):
    return '{"recipients":{"list_id":"'+listid+'"},"type":"regular","settings":{"subject_line":"'+subject+'","reply_to":"ubcbiztech@gmail.com","from_name":"BizBot"}}'

def getMailChimp(endpoint):
    url = MAILCHIMP_ENDPOINT_URL + endpoint
    r = requests.get(url,auth=('user',MAILCHIMP_API_KEY))
    return r.json()

def postMailChimp(endpoint,requestParam):
    url = MAILCHIMP_ENDPOINT_URL + endpoint
    r = requests.post(url,auth=('user',MAILCHIMP_API_KEY),data=requestParam)
    return r.json()

def deleteMailChimpList(listid):
    url = MAILCHIMP_ENDPOINT_URL + '/list/' + listid 
    r = requests.post(url,auth=('user',MAILCHIMP_API_KEY))
    return r.json()

def main():
    googleSheetsUrl = raw_input("paste google sheets URL: ")
    emailList = buildMembers(readSheets(googleSheetsUrl)) #get attendee emails and names
    upcomingEvents = getFacebook('/933982143299464/events?time_filter=upcoming')
    eventId=upcomingEvents['data'][0]['id']
    eventInfo = getFacebook('/'+eventId+'?fields=cover,name,start_time,end_time')
    eventCoverPhoto = eventInfo['cover']['source']
    eventName = eventInfo['name']
    listId = (postMailChimp('/lists','{"name":"'+eventName+'","contact":{"company":"UBC BizTech","address1":"2053 Main Mall","city":"Vancouver","state":"BC","zip":"V6T 1Z2","country":"CAN","phone":""},"permission_reminder":"You\'re receiving this email because you signed up for an event.","campaign_defaults":{"from_name":"BizBot","from_email":"ubcbiztech@gmail.com","subject":"","language":"en"},"email_type_option":false}')['id'])
    if ('status' in postMailChimp('/lists/'+listId,emailList)):
        print("error at subscribers to list") #add subscribers to list
        return
    campaignId = postMailChimp('/campaigns',mailChimpCampaignCreateRequestParam(eventName,listId))['id'] #create a campaign and attach it to the list
    mailChimpAddTemplate(eventId,eventInfo,campaignId) #add template to the campaign
    try:
        test = postMailChimp('/campaigns/'+campaignId+'/actions/test','{"test_emails":["raymond@ubcbiztech.com"],"send_type":"html"}')
    except ValueError:
        print("error at send test")


if __name__ == '__main__':
    main()


