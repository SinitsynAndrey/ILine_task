import json
import logging
import os
from datetime import datetime

import apiclient
import httplib2
import requests
from dotenv import load_dotenv
from oauth2client.service_account import ServiceAccountCredentials

load_dotenv('.env')
BASE_URL = os.getenv("BASE_URL")
logger = logging.getLogger(__name__)


def get_token(grant_type="authorization_code") -> None:
    url = BASE_URL + '/oauth2/access_token'
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("SECRET_KEY")

    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "grant_type": grant_type,
        "redirect_uri": os.getenv("REDIRECT_URL")
    }
    if grant_type == 'authorization_code':
        data.update({"code": os.getenv("AUTHORIZATION_CODE")})
    elif grant_type == 'refresh_token':
        data.update({"refresh_token": os.getenv("REFRESH_TOKEN")})
    res = requests.post(url, data=data)
    with open('tokens.json', 'w', encoding='utf-8') as j_f:
        json_format = json.loads(res.text)
        json.dump(json_format, j_f, ensure_ascii=False)


def get_events() -> str:
    url = BASE_URL + '/api/v4/events'
    access_token = os.getenv('ACCESS_TOKEN')
    headers = {
        'Authorization': f'Bearer {access_token}'
    }
    params = {'filter[type]': 'entity_linked'}
    res = requests.get(url, params=params, headers=headers)
    return res.text


def prepare_value(events: str) -> str:
    json_events = json.loads(events)
    try:
        data = {"date": datetime.timestamp(datetime.now()),
                "events": json_events['_embedded']}
        return json.dumps(data)
    except KeyError:
        logger.error("Need to refresh token")
        exit(1)


def get_cell(service) -> str:
    values = service.spreadsheets().values().get(
        spreadsheetId=os.getenv('SPREADSHEET_ID'),
        range='A:A',
        majorDimension='ROWS'
    ).execute()
    if 'values' in values:
        last_value = json.loads(values['values'][-1][0])
        date_value = datetime.fromtimestamp(last_value['date']).date()
        if date_value == datetime.now().date():
            return f'A{len(values["values"])}:A{len(values["values"])}'
        return f'A{len(values["values"]) + 1}:A{len(values["values"]) + 1}'
    else:
        return 'A1:A1'


def write_events() -> None:
    service = create_google_sheets_service()

    events = prepare_value(get_events())
    cell = get_cell(service)
    service.spreadsheets().values().update(
        range=cell,
        spreadsheetId=os.getenv('SPREADSHEET_ID'),
        valueInputOption='RAW',
        body={
            "range": cell,
            "majorDimension": 'COLUMNS',
            "values": [[events]]
        }
    ).execute()


def create_google_sheets_service():
    credentials_file = 'creds.json'

    credentials = ServiceAccountCredentials.from_json_keyfile_name(
        credentials_file,
        ['https://www.googleapis.com/auth/spreadsheets',
         'https://www.googleapis.com/auth/drive'])

    http_auth = credentials.authorize(httplib2.Http())
    service = apiclient.discovery.build('sheets', 'v4', http=http_auth)
    return service


if __name__ == "__main__":
    write_events()
