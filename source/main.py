import traceback
import logging

import yaml
import pandas as pd
import gspread
from oauth2client.service_account import ServiceAccountCredentials

file_log = logging.FileHandler('log.log', encoding='utf-8')
logging.basicConfig(handlers=(file_log,),
                    format='[%(asctime)s | %(levelname)s | %(name)s]: %(message)s',
                    datefmt='%m.%d.%Y %H:%M:%S',
                    level=logging.INFO)

log = logging.getLogger("sheet-bot")

config_file = yaml.safe_load(open("config.yml", encoding='utf-8'))


def create_number_to_letter_dict():
    number_to_letter = {}

    # Числа от 1 до 26
    for num in range(1, 27):
        letter = chr(ord('A') + num - 1)
        number_to_letter[num] = letter

    # Числа больше 26
    for tens in range(1, 27):
        tens_letter = chr(ord('A') + tens - 1)
        for units in range(1, 27):
            units_letter = chr(ord('A') + units - 1)
            number = 26 + (tens - 1) * 26 + units
            letter = tens_letter + units_letter
            number_to_letter[number] = letter

    return number_to_letter


number_to_letter_dict = create_number_to_letter_dict()


def init_client():
    # устанавливаем учетные данные OAuth2 для доступа к API Google Sheets
    scope = ['https://www.googleapis.com/auth/spreadsheets',
             'https://spreadsheets.google.com/feeds',
             'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
    return gspread.authorize(creds)


def read_from_google_sheet(spreadsheet_name, worksheet_name):
    client = init_client()

    # получаем таблицу Google Sheets с указанным именем
    spreadsheet = client.open(spreadsheet_name)

    # получаем лист с указанным именем
    worksheet = spreadsheet.worksheet(worksheet_name)

    # получаем значения из листа и преобразуем их в dataframe
    data = worksheet.get_all_values()
    df = pd.DataFrame(data)

    return df


def write_to_google_sheet(dataframe, spreadsheet_name, worksheet_name):
    client = init_client()

    spreadsheet = client.open(spreadsheet_name)

    worksheet = spreadsheet.worksheet(worksheet_name)

    data = dataframe.values.tolist()

    for i in range(len(data)):
        for j in range(len(data[i])):
            if "|" in data[i][j]:
                worksheet.format(f"{number_to_letter_dict[j+1]}{i+1}", {
                    "backgroundColor": {
                        "red": 1.0,
                        "green": 1.0,
                        "blue": 0.0
                    }})
    data = list(map(lambda x: list(map(lambda y: str(y).split('|')[0], x)), data))

    # Записываем данные в Google Таблицу
    worksheet.update(f'A1', data)

    print(f"Данные записаны в лист '{worksheet_name}' в таблице '{spreadsheet_name}'.")


def parse_table(spreadsheet_name, worksheet_name):
    df = read_from_google_sheet(spreadsheet_name, worksheet_name)
    # Перебор столбцов
    for col in range(3, df.shape[1]):
        # Перебор строк
        all_col = df.iloc[1:len(df), col]
        date = df.iloc[0, col]
        if all_col.all() and date:
            continue
        if all_col.any() and date:
            for row in range(1, df.shape[0]):
                left_value = df.iloc[row, col - 1]
                right_value = df.iloc[row, col]
                if left_value and not right_value:
                    df.iloc[row, col] = left_value.replace("|", "")
                elif left_value != right_value:
                    df.iloc[row, col] = f"{df.iloc[row, col]}|"
        elif all_col.any() and not date:
            df.iloc[0, col] = f"забор|"
            for row in range(1, df.shape[0]):
                df.iloc[row, col] = f"{df.iloc[row, col]}|"
        else:
            return df


def main():
    try:
        print('Начало работы скрипта')
        table_name = config_file['table']['table_name']
        worksheet_names = config_file['table']['worksheet_names']
        for worksheet_name in worksheet_names:
            result = parse_table(table_name, worksheet_name)
            write_to_google_sheet(result, table_name, worksheet_name)
        print('Скрипт завершил работу')
    except Exception as ex:
        print("Произошла ошибка, отправьте файл log.log разработчику")
        log.error({"error": ex, "traceback": traceback.format_exc()})


if __name__ == '__main__':
    main()
    input('\n\nДля выхода из программы нажмите любую кнопку:')
