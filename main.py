import os
import requests
import json
import xml.etree.ElementTree as ET
import pandas as pd
import urllib3
import datetime

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# Dictionary of routers
router_dict = {
    'bb1.alm1': '217.196.30.129',
    'bb11.alm1': '95.141.143.133',
    'bb0.alm1': '95.141.143.132',
    'bb10.ast1': '217.196.24.10',
    'bb0.atr1': '217.196.16.4',
    'gw0.atr1': '217.196.16.14',
    'bb10.akb1': '95.141.142.17',
    'bb0.akt1': '217.196.20.2',
    'bb10.url2': '217.196.19.215',
    'gw0.url2': '217.196.19.214',
    'bb11.pvl1': '95.141.140.1',
    'bb11.krg1': '217.196.25.3',
    'bb10.shm1': '95.141.135.3',
    'bb10.akt1': '217.196.21.4'
}

URL = "https://billing.kaztranscom.kz/services_onyma_api/service.htms?name=OnymaApi"

HEADERS = {
    "Content-Type": "application/xml",
}

TAG_TO_COLUMN_MAP = {
    'f1': 'Группa',
    'f2': 'Лицевой счет',
    'f3': 'Тип договора',
    'f4': 'Название клиента',
    'f5': 'Учетное имя',
    'f6': 'Потребительский сегмент',
    'f7': 'Проект',
    'f8': 'Тарифный план',
    'f9': 'Ресурс',
    'f10': 'Подключение',
    'f11': 'Статус подключения',
    'f12': 'Дата подключения ресурса',
    'f13': 'Дата окончания ресурса',
    'f14': 'Маршрутизатор',
    'f15': 'Сеть',
    'f16': 'IP- адреса (диапазон IP-адресов)',
    'f17': 'Скорость подключения',
    'f18': 'Ед.изм.',
    'f19': 'Наименование технологии LM',
    'f20': 'Код классификатора',
    'f21': 'Услуга',
    'f22': 'ROUTER-NEW'
}

def extract_ip(x):
    return x.split('/')[0].strip() if isinstance(x, str) else x

def compare_excel_files(first_file, second_file):
    # Read the first Excel file
    df1 = pd.read_excel(first_file)
    df2 = pd.read_excel(second_file)

    # Extract IP addresses from the columns
    df1['IP'] = df1['IP- адреса (диапазон IP-адресов)'].apply(extract_ip)
    df1['Маршрутизатор'] = df1['Маршрутизатор'].apply(lambda x: x.split('-')[1].strip() if isinstance(x, str) else x)

    # Drop duplicates
    df1.drop_duplicates(subset=['IP', 'Маршрутизатор'], keep='first', inplace=True)

    # Result list
    results = []

    # Iterate over the first dataframe
    for index, row in df1.iterrows():
        ip = row['IP']
        router = row['Маршрутизатор']
        connection = row['Подключение']

        # Find the corresponding row in the second dataframe
        match = df2[df2['INTERFACE'] == ip]
        if not match.empty:
            router_name = match['ROUTER NAME'].values[0]
            router_ip = router_dict.get(router_name, 'Router not found in dictionary')

            # Check if the router IPs match
            status = 'correct' if router == router_ip else 'incorrect'

            # Append to the results
            results.append({
                'IP': ip,
                'Status': status,
                'Router IP': router_ip,
                'Connection': connection
            })

    # Convert results to a DataFrame and return
    result_df = pd.DataFrame(results)
    return result_df


class ExcelComparisonApp:
    
    def __init__(self):
        self.first_file_path = ""
        self.second_file_path = ""
        self.date = '2023.08.23'
        self.main_loop()
        
    def main_loop(self):
        """Main interaction loop where the user can select options."""
        while True:
            print("\nExcel Comparison Tool")
            print("1. Download from ONYMA")
            print("2. Download from SERVER")
            print("3. Take a result")
            print("4. Exit")
            print("5. Set Date")
            
            choice = input("Enter your choice (1/2/3/4/5): ")

            if choice == "1":
                self.browse_first_file()
            elif choice == "2":
                self.browse_second_file()
            elif choice == "3":
                self.compare_files()
            elif choice == "4":
                print("Exiting...")
                break
            elif choice == "5":
                self.set_date()
            else:
                print("Invalid choice. Please try again.")

    def get_token(self, email, password):
        data = {
            "email": email,
            "password": password
        }

        url = 'https://tc.kaztranscom.kz:9000/auth/sign-in'
        response = requests.post(url, json=data, verify=False)

        if response.status_code == 200:
            token = response.json().get('token', '')
            return token
        else:
            raise Exception("Failed to get the token.")


    def fetch_data_main(self, token, date):
        url = f'https://tc.jusanmobile.kz:9000/api/router_onyma/speeds'
        headers = {
            'accept': 'application/json, text/plain, */*',
            'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
            'authorization': f'Bearer {token}',
            'origin': 'https://tc.jusanmobile.kz',
            'referer': 'https://tc.jusanmobile.kz/',
            'sec-ch-ua': '"Not/A)Brand";v="99", "Google Chrome";v="115", "Chromium";v="115"',
            'sec-ch-ua-mobile': '?0',
            'sec-ch-ua-platform': 'Windows',
            'sec-fetch-dest': 'empty',
            'sec-fetch-mode': 'cors',
            'sec-fetch-site': 'same-site',
            'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36'
        }

        params = {
            'date': date,
            'ip_interface': '',
            'router_name': ''
        }

        response = requests.get(url, headers=headers, params=params, verify=False)

        if response.status_code == 200:
            data = json.loads(response.text)
            return data
        else:
            raise Exception("Failed to fetch the second file.")

    def process_data(self, data):
        processed_data = []
        
        for item in data['data']:
            processed_item = {
                'DATE AND TIME': item['insert_datetime'],
                'BRANCH': item['branch_contract'],
                'ROUTER NAME': item['router_name'],
                'INTERFACE NAME': item['interface_name'],
                'INTERFACE DESCRIPTION': item['interface_description'],
                'INPUT POLICY': item['in_policy_router'],
                'INPUT SPEED': item['in_speed_router'],
                'OUTPUT POLICY': item['out_policy_router'],
                'OUTPUT SPEED': item['out_speed_router'],
                'INTERFACE': item['ip_interface'],
                'GAP': '',
                'НОМЕР ЛС': item['clsrv'],
                'НАИМЕНОВАНИЕ': item['company_name'],
                'INPUT SPEED (ONYMA)': item['in_speed_onyma'],
                'OUTPUT SPEED (ONYMA)': item['out_speed_onyma'],
                'ДППП': item['dppp_name']
            }
            processed_data.append(processed_item)
        return processed_data
    
    
    def set_date(self):
        date_input = input("Enter the date (format YYYY.MM.DD): ")
        try:
            # Check if the date format is correct
            datetime.datetime.strptime(date_input, '%Y.%m.%d')
            self.date = date_input
            print(f"Date set to {self.date}")
        except ValueError:
            print("The date format is incorrect. Please use the format YYYY.MM.DD.")
     
    def fetch_data(self):
        email = 'r.khakimov@kaztranscom.kz'
        password = 'Jr78Ndg56Fsd#$'  
        date =   self.date          #self.date_entry.get()         

        try:
            token = self.get_token(email, password)
            data = self.fetch_data_main(token, date)
            processed_data = self.process_data(data)
            df = pd.DataFrame(processed_data)
            output_file = 'second_file.xlsx'
            df.to_excel(output_file, index=False)
            self.second_file_path = output_file
            print("Успешно", "Файл успешно загружен с сервера.")
        except Exception as e:
            print("Ошибка", f"Ошибка при загрузке файла с сервера: {e}")
    
    def fetch_from_server(self):
        batch_size = 1000
        skip = 0
        all_data = []

        SOAP_BODY_TEMPLATE = """<?xml version="1.0" encoding="UTF-8"?>
<soapenv:Envelope xmlns:soapenv="http://schemas.xmlsoap.org/soap/envelope/">
    <soapenv:Header>
        <heads:credentials xmlns:heads="http://www.onyma.ru/services/OnymaApi/heads/">
            <heads:username>mypo</heads:username>
            <heads:password>YAKHuENH</heads:password>
        </heads:credentials>
    </soapenv:Header>
    <soapenv:Body>
        <onyma_api_uapi_v_query_12257>
            <rows_limit>{limit}</rows_limit>
            <rows_skip>{skip}</rows_skip>
        </onyma_api_uapi_v_query_12257>
    </soapenv:Body>
</soapenv:Envelope>
"""
        
        while True:
            # Create the SOAP envelope with current skip and limit
            soap_body = SOAP_BODY_TEMPLATE.format(limit=batch_size, skip=skip)

            try:
                response = requests.get(URL, headers=HEADERS, data=soap_body, verify=False)
                root = ET.fromstring(response.content)

                current_data_batch = []
                for row in root.findall('.//row'):
                    record = {}
                    for child in row:
                        if child.tag in TAG_TO_COLUMN_MAP:
                            record[TAG_TO_COLUMN_MAP[child.tag]] = child.text
                    current_data_batch.append(record)

                # If no data is returned or less data than expected, stop fetching
                if not current_data_batch or len(current_data_batch) < batch_size:
                    break

                all_data.extend(current_data_batch)

                # Move to the next batch
                skip += batch_size

            except Exception as e:
                print("Ошибка", f"Ошибка при загрузке файла с сервера: {e}")
                return

        df = pd.DataFrame(all_data)
        output_file = "first_file.xlsx"
        df.to_excel(output_file, index=False)
        self.first_file_path = output_file
        print("Успешно", "Файл успешно загружен с сервера.")

     
    def browse_first_file(self):
        self.fetch_from_server()

    def browse_second_file(self):
        self.fetch_data()

    def compare_files(self):
        first_file = self.first_file_path
        second_file = self.second_file_path

        # Check if both files are selected
        if not first_file or not second_file:
            print("Error", "Пожалуйста выберите оба файла.")
            return
        
        # Read the Excel files and perform the comparison logic here
        try:
            df1 = pd.read_excel(first_file)
            df2 = pd.read_excel(second_file)

            # Your comparison logic goes here
            result_df = compare_excel_files(first_file, second_file)
            
            output_directory = os.path.dirname(first_file)
            output_filename = os.path.join(output_directory, "comparison_result.xlsx")
            
            ip_to_router_name = {v: k for k, v in router_dict.items()}
            result_df["Router Name"] = result_df["Router IP"].map(ip_to_router_name)
            result_df = result_df[['IP', 'Status', 'Router Name', 'Router IP', 'Connection']]
            
            df_second = pd.read_excel(second_file)
            ip_to_branch_mapping = df_second.set_index('INTERFACE')['BRANCH'].to_dict()
            if 'IP' in result_df.columns:
                result_df['Branch'] = result_df['IP'].map(ip_to_branch_mapping)
            else:
                print("Warning: 'IP' column not found in result_df.")
                        
            result_df.to_excel(output_filename, index=False)

            # Show the result to the user
            print("Результат сравнения", f"Сравнение завершено!\nРезультат сохранен в файл {output_filename}")
        except Exception as e:
            print("Error", f"An error occurred: {e}")


if __name__ == "__main__": 
    app = ExcelComparisonApp()
