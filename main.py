import pygsheets
import requests
import pandas as pd
from dotenv import dotenv_values
from datetime import datetime
from datetime import timedelta

gc = pygsheets.authorize(client_secret='client_secret.json')
config = dotenv_values(".env")
currentMonth = datetime.now()


dateNow = datetime.now() + timedelta(days=1)
daysLess = dateNow - timedelta(days=30)
until = dateNow.strftime('%Y-%m-%d')
since = daysLess.strftime('%Y-%m-%d')

# abaixo, vamos conectar a API do pagerduty e retornar os dados necessários
headers = {
    'Content-Type': 'application/json',
    'Authorization': config['API_KEY'],
    'Accept': 'application/vnd.pagerduty+json;version=2'
}
url = f'https://stone.pagerduty.com/api/v1/schedules/PNMATVI?include%5B%5D=privileges&include%5B%5D=color&include%5B%5D=escalation_policies&overflow=true&include_oncall=true&rotation_target=2023-07-29T14%3A23%3A00.000-03%3A00&since={since}T00%3A00%3A00.000-03%3A00&until={until}T00%3A00%3A00.000-03%3A00&include_next_oncall_for_user=PT6UZH7'

response = requests.get(url, headers=headers)
data = response.json()

schedule = data['schedule']['final_schedule']['rendered_schedule_entries']

# abaixo, vamos tratar os dados e acrescentá-los a um dicionário 
scheduleObject = []
for i in schedule:
    dateStart = datetime.fromisoformat(i['start'])
    dateEnd = datetime.fromisoformat(i['end'])
    
    if dateStart.weekday() <= 4:
        startWorkHours  = dateEnd.replace(hour=8, minute=59, second=59)
        endWorkHours = dateStart.replace(hour=19, minute=0, second=0)
        results = ((startWorkHours  - endWorkHours))
        warningHours = str(results)

    if dateStart.weekday() > 5:
        finalSemana = dateEnd.replace(hour=8, minute=59, second=59)
        inicioSemana = dateStart.replace(hour=19, minute=0, second=0)
        results = ((finalSemana - inicioSemana))
        warningHours = str(results)

    if dateStart.weekday() == 4:
        results = ((dateEnd - endWorkHours) - timedelta(seconds=1))
        warningHours = str(results)

    if dateStart.weekday() == 5:
        result = ((dateEnd - dateStart) - timedelta(seconds=1))
        warningHours = str(results)

    objected = {
        "username": i['user']['summary'],
        "warningHours": warningHours,
        "dayWeekStart": dateStart.strftime('%A'),
        "startDate": datetime.fromisoformat(i['start']),
        "dayWeekEnd": dateEnd.strftime('%A'),
        "end": datetime.fromisoformat(i['end'])
    }
    scheduleObject.append(objected)

# funções para validar se existe uma planilha para o caso ou então criá-la 
gsheetName = f'Escala de Sobreaviso {currentMonth.strftime("%b")}-2023'

def spreadsheetExists(nome):
    files = gc.drive.list(q=f"name='{nome}' and mimeType='application/vnd.google-apps.spreadsheet'")
    return len(files) > 0

def spreadsheetCreated(nome):
    newSpreadsheet = gc.create(nome) 
    return newSpreadsheet

if spreadsheetExists(gsheetName):
    print('file already exists!')
else:
    spreadsheet = spreadsheetCreated(gsheetName)
    print('Spreadsheet created successfully!')
    
# função para preencher planilha    
def fillSpreadsheet(spreadsheet, data):
    guide = spreadsheet.sheet1
    df = pd.DataFrame(data)
    guide.set_dataframe(df, start='A1')


fillSpreadsheet(spreadsheet, scheduleObject)