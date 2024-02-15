import os.path
import math
import logging
import sys
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

#Caso o escopo seja modificado, deletar o tokens.json do diretório
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

#O id e range da planilha usada
scores_sheet = "1ohImect67ak1vD2f6OmCHaeC1f4SHviZtcUAbj8M1UM"
range_name = "engenharia_de_software!A4:H27"
total_classes = 60

#Logger
logging.basicConfig(level=logging.INFO)

class Student:
    def __init__(self, registration, name, absences, p1, p2, p3):
        self.registration = registration
        self.name = name
        self.absences = absences
        self.scores = [p1, p2, p3]
        self.final_score_needed = None

    def calculate_mean_score(self):
        return sum(self.scores) / len(self.scores)

    def calculate_situation(self):
        max_absences = .25*total_classes
        if self.absences > max_absences:
            logging.info(f"{self.name} has failed due to absences.")
            return "Reprovado por Falta"

        mean_score = self.calculate_mean_score()
        mean_score /= 10

        if mean_score < 5:
            logging.info(f"{self.name} has failed due to low score.")
            return "Reprovado por Nota"
        elif mean_score < 7:
            #Calculando a média final e arredondando pra cima quando necessário
            self.final_score_needed = math.ceil((50 - (mean_score*6))/4)
            logging.info(f"{self.name} needs to take the final exam. Score needed: {self.final_score_needed}")
            return "Exame Final"
        else:
            logging.info(f"{self.name} has been approved.")
            return "Aprovado"

def update_student_situations(service, values):
    update_data = []
    sheet_range = range_name.split('!')[1]
    base_row = sheet_range.split(':')[0][1:]

    for i, row in enumerate(values):
        student = Student(row[0], row[1], int(row[2]), int(row[3]), int(row[4]), int(row[5]))
        situation = student.calculate_situation()

        #Retorna a nota final necessária mínima ou 0 caso não precise de prva final
        final_score_needed = student.final_score_needed if situation == "Exame Final" else 0

        #Preparando os valores que irão ser inseridos na planilha (situação and Nota para Aprovação Final)
        update_range = f"engenharia_de_software!G{base_row}:{sheet_range.split(':')[1][0]}{base_row}"
        #Incrementa o numero da linha para cada iteração
        base_row = str(int(base_row) +1)
        update_values = [[situation, final_score_needed]]

        #Usa o append para atualizar por lote
        update_data.append({
            'range': update_range,
            'values': update_values
        })
    
    body = {
        'valueInputOption': 'USER_ENTERED',
        'data': update_data
    }

    try:
        result = service.spreadsheets().values().batchUpdate(
            spreadsheetId=scores_sheet,
            body=body).execute()
        logging.info(f"{result.get('totalUpdatedCells')} células atualizadas.")
    except HttpError as e:
        logging.error(f"An error has occurred: {e}")

def main():
  creds = None

  if os.path.exists("token.json"):
    creds = Credentials.from_authorized_user_file("token.json", SCOPES)
  #Caso não existam credenciais válidas permite o user se logar.
  if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
      creds.refresh(Request())
    else:
        try:
            flow = InstalledAppFlow.from_client_secrets_file(
                "credentials.json", SCOPES
            )
            creds = flow.run_local_server(port=0)
        except FileNotFoundError as fnf_err:
            logging.error(f"Your credentials.json was not found in this application's directory. Visit the following website for instructions o how to download your credentials: https://developers.google.com/sheets/api/quickstart/python")
            sys.exit(1)

    #Salva as credenciais para rodar posteriormente sem necessidade de logar
    with open("token.json", "w") as token:
      token.write(creds.to_json())

  try:
        service = build("sheets", "V4", credentials=creds)

        #Chama a sheets API
        sheet = service.spreadsheets()
        result = (
            sheet.values()
            .get(spreadsheetId=scores_sheet, range=range_name)
            .execute()
        )
        values = result.get("values", [])

        if not values:
            print("No data found.")
        else:
            update_student_situations(service, values)

  except HttpError as err:
    logging.error(f"A HTTP error has occurred: {err}")


if __name__ == "__main__":
  main()
