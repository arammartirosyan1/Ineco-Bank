import pandas as pd
import random
import datetime
from datetime import date
from dateutil.relativedelta import relativedelta
import sys

ls = ["Հաճախորդ", "Հաճախորդի անուն/ազգանուն", "Ապահովադիր", "Ծննդյան ամսաթիվ",
          "ՀԾՀ", "Հաճախորդի անձնագրային տվյալներ", "Հեռախոսահամար", "Էլ. հասցե",
          "Գրանցման հասցե", "Բնակության հասցե", "Շահառու", "ՀՎՀՀ", "Ապահովագրվող գույքի տեսակ",
          "Գույքի ապահովագրական արժեք", "Ապահովագրական գումար", "Սակագին", "Ապահովագրվող գույքի հասցե",
          "Սեփականատեր անհատ", "Նոր պայմանագրի սկիզբ"]

ls_city = ['Երևան', 'Արմավիր', 'Արարատ', 'Արագածոտն', 'Կոտայք', 'Շիրակ', 'Լոռի', 'Տավուշ', 'Գեղարքունիք',
           'Վայոց Ձոր', 'Սյունիք']

ls_1 = {}

ani_to_uni = {"²": "Ա",
                  "³": "ա",
                  "´": "Բ",
                  "µ": "բ",
                  "¶": "Գ",
                  "·": "գ",
                  "¸": "Դ",
                  "¹": "դ",
                  "º": "Ե",
                  "»": "ե",
                  '¼': 'Զ',
                  '½': 'զ',
                  '¾': 'Է',
                  '¿': 'է',
                  'À': 'Ը',
                  'Á': 'ը',
                  'Â': 'Թ',
                  'Ã': 'թ',
                  'Ä': 'Ժ',
                  'Å': 'ժ',
                  'Æ': 'Ի',
                  'Ç': 'ի',
                  'È': 'Լ',
                  'É': 'լ',
                  'Ê': 'Խ',
                  'Ë': 'խ',
                  'Ì': 'Ծ',
                  'Í': 'ծ',
                  'Î': 'Կ',
                  'Ï': 'կ',
                  'Ð': 'Հ',
                  'Ñ': 'հ',
                  'Ò': 'Ձ',
                  'Ó': 'ձ',
                  'Ô': 'Ղ',
                  'Õ': 'ղ',
                  'Ö': 'Ճ',
                  '×': 'ճ',
                  'Ø': 'Մ',
                  'Ù': 'մ',
                  'Ú': 'Յ',
                  'Û': 'յ',
                  'Ü': 'Ն',
                  'Ý': 'ն',
                  'Þ': 'Շ',
                  'ß': 'շ',
                  'à': 'Ո',
                  'á': 'ո',
                  'â': 'Չ',
                  'ã': 'չ',
                  'ä': 'Պ',
                  'å': 'պ',
                  'æ': 'Ջ',
                  'ç': 'ջ',
                  'è': 'Ռ',
                  'é': 'ռ',
                  'ê': 'Ս',
                  'ë': 'ս',
                  'ì': 'Վ',
                  'í': 'վ',
                  'î': 'Տ',
                  'ï': 'տ',
                  'ð': 'Ր',
                  'ñ': 'ր',
                  'ò': 'Ց',
                  'ó': 'ց',
                  'ô': 'Ւ',
                  'õ': 'ւ',
                  'ö': 'Փ',
                  '÷': 'փ',
                  'ø': 'Ք',
                  'ù': 'ք',
                  'ú': 'Օ',
                  'û': 'օ',
                  'ü': 'Ֆ',
                  'ý': 'ֆ',
                  '¨': 'և',
                  "'": "՚",
                  '°': '՛',
                  '¯': '՜',
                  'ª': '՝',
                  '±': '՞',
                  '£': '։',
                  '­': '֊',
                  '§': '«',
                  '¦': '»',
                  '«': ',',
                  '©': '.',
                  '®': '…'

                  }

ls_2 = ["ք․", "ք.", "Ք․", "Ք.", "քաղաք", "Քաղաք", "մ․", "մ.", "Մ․", "Մ.", "մարզ", "Մարզ"]

df = pd.read_excel('accident.xlsx', header=None)
df.drop(df.index[:2], inplace=True)
len_row = df.columns[1::]

for i in len_row:
    try:
        df.replace({i: ani_to_uni}, regex=True, inplace=True)
    except:
        pass

for i in range(len(df)):
    data = df.values[i][1::]
    if len(str(data[2])) > 10:
        new = {"": data}
        exc_data = pd.DataFrame(new, index=ls)
        exc_data.to_excel('News/Excel_{}.xlsx'.format(i))

#########################################################
        if "Վարկառու/" in data[0]:
            for i in range(len(data)):
                try:
                    if i == 0:
                        x = str(data[i]).split()
                        ls_1["IS_INSURED_PHYSICAL"] = "1"
                except:
                    print("Error", data[i])

                try:
                    if i == 2:
                        x = str(data[i]).split()
                        ls_1["INSURED_NAME"] = x[0]
                        ls_1["INSURED_LAST_NAME"] = x[1]
                        ls_1["INSURED_SECOND_NAME"] = x[2]
                except:
                    print("Error", data[i])

                try:
                    if i == 3:
                        x = str(data[i]).split()[0]
                        x = datetime.datetime.strptime(x, "%Y-%m-%d").strftime("%d.%m.%Y")
                        ls_1["INSURED_BIRTHDAY"] = x
                except:
                    print("Error", data[i])

                try:
                    if i == 4:
                        x = str(data[i]).strip('.').strip('').split('.')
                        ls_1["INSURED_SOCIAL_CARD"] = x[0]
                        if int(x[0][:2]) > 50:
                            ls_1["INSURED_GENDER"] = "F"
                        else:
                            ls_1["INSURED_GENDER"] = "M"
                except:
                    print("Error", data[i])

                try:
                    if i == 5:
                        x = str(data[i]).split()
                        ls_1["INSURED_CITIZENSHIP"] = "ՀՀ"
                        ls_1["INSURED_PASSPORT_NUMBER"] = x[0]
                        ls_1["INSURED_PASSPORT_ISSUE_DATE"] = x[1].replace('/', '.')
                        ls_1["INSURED_PASSPORT_AUTHORITY"] = x[2]
                        start = datetime.datetime.strptime(x[1], "%d/%m/%Y")
                        end = start + relativedelta(years=+10)
                        ls_1["INSURED_PASSPORT_EXPIRY_DATE"] = end.strftime("%d.%m.%Y")
                except:
                    print("Error", data[i])

                try:
                    if i == 6:
                        x = str(data[i]).split('.')[0].strip()
                        if len(x) == 9:
                            ls_1["INSURED_MOBILE_PHONE"] = x
                        else:
                            x = "0" + x[::-1][:8][::-1]
                            ls_1["INSURED_MOBILE_PHONE"] = x
                except:
                    print("Error", data[i])

                try:
                    if i == 7:
                        x = str(data[i]).split()
                        ls_1["INSURED_MAIL"] = x[0]
                except:
                    print("Error", data[i])

                try:
                    if i == 8:
                        new_data = str(data[i]).replace(',', " ").strip()
                        new_data = new_data.split(" ")
                        for k in new_data[0:2]:
                            for j in ls_city:
                                if j in k:
                                    new_data_1 = str(data[i]).replace(k, '')
                                    for q in ls_2:
                                        if q in data[i]:
                                            qaxaq = q
                                    new_data_1 = new_data_1.replace(qaxaq, "").strip().strip(",").strip()
                                    ls_1["INSURED_REG_COUNTRY"] = "ARM"
                                    ls_1["INSURED_REG_REGION"] = j
                                    ls_1['INSURED_REG_FULL_ADDRESS'] = new_data_1
                                    if j == "Տավուշ":
                                        ls_1['INSURED_REG_CITY'] = "Այլ"
                                    else:
                                        ls_1['INSURED_REG_CITY'] = j
                except:
                    if i == 8:
                        new_data_1 = str(new_data_1).strip().strip(",").strip()
                        ls_1["INSURED_REG_COUNTRY"] = "ARM"
                        ls_1["INSURED_REG_REGION"] = j
                        ls_1['INSURED_REG_FULL_ADDRESS'] = new_data_1
                        if j == "Տավուշ":
                            ls_1['INSURED_REG_CITY'] = "Այլ"
                        else:
                            ls_1['INSURED_REG_CITY'] = j

                try:
                    if i == 9:
                        new_data = str(data[i]).replace(',', " ").strip()
                        new_data = new_data.split(" ")
                        for k in new_data[0:2]:
                            for j in ls_city:
                                if j in k:
                                    new_data_1 = str(data[i]).replace(k, '')
                                    for q in ls_2:
                                        if q in data[i]:
                                            qaxaq = q
                                    new_data_1 = new_data_1.replace(qaxaq, "").strip().strip(",").strip()
                                    ls_1["INSURED_LIVE_COUNTRY"] = "ARM"
                                    ls_1["INSURED_LIVE_REGION"] = j
                                    ls_1['INSURED_LIVE_FULL_ADDRESS'] = new_data_1
                                    if j == "Տավուշ":
                                        ls_1['INSURED_LIVE_CITY'] = "Այլ"
                                    else:
                                        ls_1['INSURED_LIVE_CITY'] = j
                                        i += 1
                except:
                    if i == 9:
                        try:
                            new_data_1 = str(new_data_1).strip().strip(",").strip()
                            ls_1["INSURED_LIVE_COUNTRY"] = "ARM"
                            ls_1["INSURED_LIVE_REGION"] = j
                            ls_1['INSURED_LIVE_FULL_ADDRESS'] = new_data_1
                            if j == "Տավուշ":
                                ls_1['INSURED_LIVE_CITY'] = "Այլ"
                            else:
                                ls_1['INSURED_LIVE_CITY'] = j
                        except:
                            ls_1["IS_INSURED_SAME_REG_LIVE_ADDRESS"] = '1'

                l = pd.Series(ls_1)
                l.to_json('News/Insured.json', indent=2, force_ascii=False)

for row in range(len(df)):
    data = df.values[row][1::]
    if len(str(data[2])) > 10:
        for i in range(len(data)):
            try:
                if i == 1:
                    x = str(data[i]).split()
                    ls_1["PA_PERSON_FIRST_NAME"] = x[0]
                    ls_1["PA_PERSON_LAST_NAME"] = x[1]
                    ls_1["PA_PERSON_SECOND_NAME"] = x[2]
            except:
                print("Error", data[i])

            try:
                if i == 3:
                    x = str(data[i]).split()[0]
                    x = datetime.datetime.strptime(x, "%Y-%m-%d").strftime("%d.%m.%Y")
                    ls_1["PA_PERSON_BIRTHDAY"] = x
            except:
                print("Error", data[i])

            try:
                if i == 4:
                    x = str(data[i]).strip('.').strip('').split('.')
                    ls_1["PA_PERSON_SOCIAL_CARD"] = x[0]
                    if int(x[0][:2]) > 50:
                        ls_1["PA_PERSON_GENDER"] = "F"
                    else:
                        ls_1["PA_PERSON_GENDER"] = "M"
            except:
                print("Error", data[i])

            try:
                if i == 5:
                    x = str(data[i]).split()
                    ls_1["PA_PERSON_CITIZENSHIP"] = "ՀՀ"
                    ls_1["PA_PERSON_PASSPORT_NUMBER"] = x[0]
                    ls_1["PA_PERSON_PASSPORT_ISSUE_DATE"] = x[1].replace('/', '.')
                    ls_1["PA_PERSON_PASSPORT_AUTHORITY"] = x[2]
                    start = datetime.datetime.strptime(x[1], "%d/%m/%Y")
                    end = start + relativedelta(years=+10)
                    ls_1["PA_PERSON_PASSPORT_EXPIRY_DATE"] = end.strftime("%d.%m.%Y")
            except:
                print("Error", data[i])

            try:
                if i == 6:
                    x = str(data[i]).split('.')[0].strip()
                    if len(x) == 9:
                        ls_1["PA_PERSON_MOBILE_PHONE"] = x
                    else:
                        x = "0" + x[::-1][:8][::-1]
                        ls_1["PA_PERSON_MOBILE_PHONE"] = x
            except:
                print("Error", data[i])

            try:
                if i == 7:
                    x = str(data[i]).split()
                    ls_1["PA_PERSON_MAIL"] = x[0]
            except:
                print("Error", data[i])

            try:
                if i == 8:
                    new_data = str(data[i]).replace(',', " ").strip()
                    new_data = new_data.split(" ")
                    for k in new_data[0:2]:
                        for j in ls_city:
                            if j in k:
                                new_data_1 = str(data[i]).replace(k, '')
                                for q in ls_2:
                                    if q in data[i]:
                                        qaxaq = q
                                new_data_1 = new_data_1.replace(qaxaq, "").strip().strip(",").strip()
                                ls_1["PA_PERSON_REG_COUNTRY"] = "ARM"
                                ls_1["PA_PERSON_REG_REGION"] = j
                                ls_1['PA_PERSON_REG_FULL_ADDRESS'] = new_data_1
                                if j == "Տավուշ":
                                    ls_1['PA_PERSON_REG_CITY'] = "Այլ"
                                else:
                                    ls_1['PA_PERSON_REG_CITY'] = j
            except:
                if i == 8:
                    new_data_1 = str(new_data_1).strip().strip(",").strip()
                    ls_1["PA_PERSON_REG_COUNTRY"] = "ARM"
                    ls_1["PA_PERSON_REG_REGION"] = j
                    ls_1['PA_PERSON_REG_FULL_ADDRESS'] = new_data_1
                    if j == "Տավուշ":
                        ls_1['PA_PERSON_REG_CITY'] = "Այլ"
                    else:
                        ls_1['PA_PERSON_REG_CITY'] = j

            try:
                if i == 14:
                    x = str(data[14]).strip().split('.')
                    ls_1['PA_RISK_AMOUNT'] = x[0]
                    risk = str(float(0.16) * int(x[0]) / 100).split('.')[0]
                    ls_1['PA_RISK_PREMIUM'] = risk
                    ls_1["POLICY_AMOUNT_CURRENCY"] = "AMD"
                    ls_1["PA_INSURANCE_RATE"] = "0.16"
            except:
                print("Error", data[i])

            try:
                if i == 18:
                    # Եթե նշված է օր վերցնում ենք այդ օրը որպես պայմանագրի սկիզբ
                    x = str(data[i]).split()[0]
                    start_send = datetime.datetime.strptime(x, "%Y-%m-%d")
                    end = start_send + datetime.timedelta(days=365)
                    ls_1["POLICY_FROM_DATE"] = start_send.strftime("%d.%m.%Y")
                    ls_1["POLICY_TO_DATE"] = end.strftime("%d.%m.%Y")
                    start_date = datetime.date.today().strftime("%d.%m.%Y")
                    ls_1["POLICY_CREATION_DATE"] = start_date
                    payment = str(risk), start_date
                    ls_1["POLICY_PAYMENT_SCHEDULE"] = ", ".join(payment)
            except:
                if i == 18:
                    # Եթե նշված չի օր վերցնում ենք այս օրը որպես պայմանագրի սկիզբ
                    start_date = datetime.date.today()
                    end_date = start_date + datetime.timedelta(days=365)
                    start = start_date.strftime("%d.%m.%Y")
                    end = end_date.strftime("%d.%m.%Y")
                    ls_1["POLICY_FROM_DATE"] = start
                    ls_1["POLICY_TO_DATE"] = end
                    ls_1["POLICY_CREATION_DATE"] = start
                    payment = str(risk), start
                    ls_1["POLICY_PAYMENT_SCHEDULE"] = ", ".join(payment)

        l = pd.Series(ls_1)
        l.to_json(f'News/PA{row}.json', indent=2, force_ascii=False)

        benef_data = pd.read_json('beneficiar.json', orient='index')[0]

        agent_data = pd.read_json('agent.json', orient='index')[0]

        result = pd.concat([benef_data, l, agent_data])

        result.to_json(f'News/PA_Format{row}.json', indent=2, force_ascii=False)
##########################################################################


        import requests
        import json
        import pandas as pd

        url = "https://testimex.efes.am/webservice/policy"
        data = pd.read_json('C:/Users/aramm/OneDrive - EFES ICJSC/Desktop/INECO/ACCIDENT/News/PA_Format0.json',
                            orient='index')[0]
        data = dict(data)

        payload = json.dumps(data)

        headers = {
            'Content-Type': 'application/json',
            'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJIUzI1NiJ9.eyJjb250ZXh0Ijp7ImNsaWVudCI6eyJpZCI6IjI3NyIsIm5hbWUiOiJJa2h0c3lhbmRyIn0sImVudiI6IlBST0QifSwiaXNzIjoid3d3LmltZXguYW0iLCJpYXQiOjE3MDc0ODE4NzEsImV4cCI6MTcwNzY1NDY3MX0.aWBYxeI0NnQZqfc63flgEbm2TAABnGPv_t6be2h_7wM'
        }

        response = requests.request("POST", url, headers=headers, data=payload)
        if response.status_code == 200:
            result = json.loads(response.text)
            print(result['result'])
        else:
            print(response.text)



