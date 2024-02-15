import pandas as pd
import datetime
from dateutil.relativedelta import relativedelta


ls_1 = {}

ls_city = ['Երևան', 'Արմավիր', 'Արարատ', 'Արագածոտն', 'Կոտայք', 'Շիրակ', 'Լոռի', 'Տավուշ', 'Գեղարքունիք',
           'Վայոց Ձոր', 'Սյունիք']

ls_2 = ["ք․", "ք.", "Ք․", "Ք.", "քաղաք", "Քաղաք", "մ․", "մ.", "Մ․", "Մ.", "մարզ", "Մարզ", "ՀՀ", "հհ"]

ls_person = ["Գույքի սեփականատեր", "", "Հասցե մարզ", "Հասցե ամբողջական", " Անձնագրի համար", "Երբ է տրվել", "Ում կողմից", " Ծննդյան ամսաթիվ", "Սոցիալական քարտի համար", "Հեռախոս"]


df = pd.read_excel("mortgage.xlsx", header=None)
data_insured = df.values[5]
for i in range(len(data_insured)):
    data = data_insured

    x = str(data[0]).strip().split()
    ls_1["INSURED_NAME"] = x[1]
    ls_1["INSURED_LAST_NAME"] = x[0]
    ls_1["INSURED_SECOND_NAME"] = x[2]

    x = str(data[2]).strip()
    if x == "Տավուշ":
        ls_1["INSURED_REG_REGION"] = x
        ls_1['INSURED_REG_CITY'] = "Այլ"
    else:
        ls_1["INSURED_REG_REGION"] = x
        ls_1['INSURED_REG_CITY'] = x

    try:
        if i == 3:
            new_data = str(data[i]).replace(',', "").strip()
            new_data = new_data.split(" ")
            for k in new_data[:3]:
                for j in ls_city:
                    if j in k:
                        new_data_1 = str(data[i]).replace(k, '')
                        for q in ls_2:
                            if q in data[i]:
                                qaxaq = q
                                new_data_1 = new_data_1.replace(qaxaq, "").strip(',').strip().strip(',').strip()
                                ls_1["INSURED_REG_COUNTRY"] = "ARM"
                                ls_1['INSURED_REG_FULL_ADDRESS'] = new_data_1


    except:
        new_data_1 = str(new_data_1).strip().strip(",").strip()
        ls_1["INSURED_REG_COUNTRY"] = "ARM"
        ls_1['INSURED_REG_FULL_ADDRESS'] = new_data_1

    x = str(data[4]).strip()
    ls_1['INSURED_CITIZENSHIP'] = "ՀՀ"
    ls_1["INSURED_PASSPORT_NUMBER"] = x.replace('՛', '')

    x = str(data[5]).strip()
    ls_1["INSURED_PASSPORT_ISSUE_DATE"] = x.replace('/', '.')
    start = datetime.datetime.strptime(x, "%d/%m/%Y")
    end = start + relativedelta(years=+10)
    ls_1["INSURED_PASSPORT_EXPIRY_DATE"] = end.strftime("%d.%m.%Y")

    ls_1["INSURED_PASSPORT_AUTHORITY"] = str(data[6]).strip()

    x = str(data[7]).split()[0]
    x = datetime.datetime.strptime(x, "%Y-%m-%d").strftime("%d.%m.%Y")
    ls_1["INSURED_BIRTHDAY"] = x

    x = str(data[8]).strip()
    ls_1["INSURED_SOCIAL_CARD"] = x
    if int(x[0]) >= 5:
        ls_1["INSURED_GENDER"] = "F"
    else:
        ls_1["INSURED_GENDER"] = "M"

    ls_1["INSURED_MOBILE_PHONE"] = "0" + str(data[9]).strip()[-8:]

    ls_1["INSURED_MAIL"] = str(df.values[6][7]).strip()


data = df.values[10]
for i in range(len(data)):
    new_data = str(data[0]).replace(',', " ").strip()
    new_data = new_data.split(" ")

    for k in new_data[0:3]:
        for j in ls_city:
            if j in k:
                new_data_1 = str(data[0]).replace(k, '')
                for q in ls_2:
                    if q in data[0]:
                        qaxaq = q
                        new_data_1 = new_data_1.replace(qaxaq, "").strip(",").strip().strip(",").strip()
                ls_1["PA_PERSON_LOCATION_TYPE"] = "REGISTRATION"
                ls_1["PA_PERSON_REG_COUNTRY"] = "ՀՀ"
                ls_1["PA_PERSON_REG_REGION"] = j
                ls_1['PA_PERSON_REG_FULL_ADDRESS'] = new_data_1
                if j == "Տավուշ":
                    ls_1['PA_PERSON_REG_CITY'] = "Այլ"
                else:
                    ls_1['PA_PERSON_REG_CITY'] = j

    ls_1["PROPERTY_DOCUMENT_NUMBER"] = str(data[2]).strip()

    x = str(data[4]).strip().split()
    ls_1["PROPERTY_NAME"] = x[0]
    ls_1["PROPERTY_TYPE"] = x[0]
    ls_1["PROPERTY_AREA_MEASURE"] = 'քմ'
    ls_1["PROPERTY_AREA"] = x[1]

    ls_1["PROPERTY_MARKET_PRICE"] = str(data[6]).split('.')[0].strip()

    ls_1["PROPERTY_SUM_INSURED"] = str(data[7]).split('.')[0].strip()
    ls_1["PROPERTY_RISK_AMOUNT"] = str(data[7]).split('.')[0].strip()
    ls_1["POLICY_AMOUNT_CURRENCY"] = "AMD"

    ls_1["PROPERTY_RISK_PREMIUM"] = str(data[9]).split('.')[0].strip()


l = pd.Series(ls_1)
l.to_json('News/Json.json', indent=2, force_ascii=False)


benef_data = pd.read_json('beneficiar.json', orient='index')[0]

agent_data = pd.read_json('agent.json', orient='index')[0]

result = pd.concat([benef_data, l, agent_data])

result.to_json('News/Format.json', indent=2, force_ascii=False)