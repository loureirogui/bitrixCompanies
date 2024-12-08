# Module for querying the Bitrix24 company database using CPF/CNPJ
# Developed by Guilherme Loureiro
# Consult companies by CNPJ and update or register as needed.

import requests
import openpyxl
import time
from art import text2art, tprint


# Bitrix field mappings
FIELD_COMPANY_TYPE = 'COMPANY_TYPE'
FIELD_TITLE = 'TITLE'
FIELD_CNPJ = 'UF_CRM_1701275490640'
FIELD_REVENUE = 'UF_CRM_1727441546022'
FIELD_EMAIL = 'EMAIL'
FIELD_PHONE = 'PHONE'
FIELD_RESPONSIBLE = 'UF_CRM_1727358267819'

tprint("Bitrix24", font="starwars")
print("Para começar, salve no diretório desse código sua planilha .xlsx com as empresas\n"
      "Que deseja cadastrar/atualizar. Em seguida, abra o arquivo bitrixCompanies em seu editor\n"
      "Para editar o código conforme orientações disponíveis nos comentários.")

# Configuration settings
spreadsheetName = input("Digite o nome da sua planilha de importação das empresas:\n")  # Enter the name of your import spreadsheet
bitrixToken = input("Digite seu token de API Bitrix24:\n")  # Insert your Bitrix token here

# Process phone numbers separated by ";"
def processPhones(phoneStr):
    if not phoneStr:
        return []
    phones = [phone.strip() for phone in phoneStr.split(";") if phone.strip()]
    return [{"VALUE": phone, "VALUE_TYPE": "WORK"} for phone in phones]

# Process emails separated by ";"
def processEmails(emailStr):
    if not emailStr:
        return []
    emails = [email.strip() for email in emailStr.split(";") if email.strip()]
    return [{"VALUE": email, "VALUE_TYPE": "WORK"} for email in emails]

# Query Bitrix24 for company by CNPJ
def queryCompany(cnpjField, cnpj, name, revenue, email, phone):
    bitrixUrl = f"https://setup.bitrix24.com.br/rest/301/{bitrixToken}/crm.company.list.json"

    params = {
        "order": {"DATE_CREATE": "ASC"},
        "filter": {cnpjField: cnpj},
        "select": ["ID", "TITLE", cnpjField],
        "start": 0
    }

    response = requests.post(bitrixUrl, json=params)

    if response.status_code == 200:
        result = response.json()
        if "error" in result:
            print(f"Error querying company; CNPJ: {cnpj}")
        else:
            companies = result.get("result", [])
            if companies:
                for company in companies:
                    bitrixId = company['ID']
                    print(f"Company found; CNPJ: {cnpj}. Updating data...")
                    updateCompany(bitrixId, cnpj, revenue, email, phone, name)
            else:
                print(f"Company not found; CNPJ: {cnpj}. Registering...")
                registerCompany(cnpj, revenue, email, phone, name)
    else:
        print(f"Request failed. Status code: {response.status_code}, Details: {response.text}")

# Update existing company
def updateCompany(bitrixId, cnpj, revenue, email, phone, name):
    bitrixUrlUpdate = f"https://setup.bitrix24.com.br/rest/301/{bitrixToken}/crm.company.update.json"

    fieldsToUpdate = {
        FIELD_TITLE: name,
        FIELD_CNPJ: cnpj,
        FIELD_REVENUE: revenue,
        FIELD_EMAIL: processEmails(email),
        FIELD_PHONE: processPhones(phone),
    }

    data = {
        "id": bitrixId,
        "fields": fieldsToUpdate,
        "params": {"REGISTER_SONET_EVENT": "Y"}
    }

    response = requests.post(bitrixUrlUpdate, json=data)

    if response.status_code == 200:
        result = response.json()
        if "error" in result:
            print(f"Error updating company: {result['error_description']}")
        else:
            print(f"Company updated successfully! ID: {result['result']}")
    else:
        print(f"Request failed. Status code: {response.status_code}, Details: {response.text}")

# Register a new company
def registerCompany(cnpj, revenue, email, phone, name):
    bitrixUrlAdd = f"https://setup.bitrix24.com.br/rest/301/{bitrixToken}/crm.company.add.json"

    fieldsToCreate = {
        FIELD_TITLE: name,
        FIELD_CNPJ: cnpj,
        FIELD_REVENUE: revenue,
        FIELD_EMAIL: processEmails(email),
        FIELD_PHONE: processPhones(phone),
    }

    data = {
        "fields": fieldsToCreate,
        "params": {"REGISTER_SONET_EVENT": "Y"}
    }

    response = requests.post(bitrixUrlAdd, json=data)

    if response.status_code == 200:
        result = response.json()
        if "error" in result:
            print(f"Error registering company: {result['error_description']}")
        else:
            print(f"Company registered successfully! ID: {result['result']}")
    else:
        print(f"Request failed. Status code: {response.status_code}, Details: {response.text}")

# Read spreadsheet and process rows
def processSpreadsheet():
    workbook = openpyxl.load_workbook(f'{spreadsheetName}.xlsx')
    sheet = workbook.active

    for row in sheet.iter_rows(min_row=2, min_col=1, max_col=23):
        cnpj = row[0].value
        revenue = row[1].value
        email = row[2].value
        phone = row[3].value
        name = row[4].value
        responsible = row[5].value
        clientStatus = row[6].value

        if clientStatus == 'Ativo':
            companyStatus = 'CUSTOMER'
        elif clientStatus == 'Inativo':
            companyStatus = 'UC_3HS5P9'
        else:
            companyStatus = None

        try:
            queryCompany(FIELD_CNPJ, cnpj, name, revenue, email, phone)
        except Exception as e:
            print(f"Error processing row. Details: {e}")
        time.sleep(0.5)

if __name__ == "__main__":
    processSpreadsheet()
