import pandas as pd
import random
import datetime
from datetime import date
from dateutil.relativedelta import relativedelta
import sys


class Error(Exception):

    def __init__(self, message, value):
        self.message = message
        self.value = value

    def __repr__(self):
        return "Wrong for {} : {}".format(self.message, self.value)


# cd C:\Users\aramm\'OneDrive - EFES ICJSC'\Desktop\INECO BANK
# python translation.py Ineco I1


class Bank:
    def __init__(self, name, type):
        self.name = name
        self.type = type

        match name:
            case "Ineco":
                match type:
                    case "I1":
                        self.ins = False
                        self.ben = True
                        ben_true()
                    case "I2":
                        self.ins = False
                        self.ben = True


def ben_true():
    global qaxaq, new_data_1, j, data, risk
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

    df = pd.read_excel('INECO.xlsx', header=None)
    df.drop(df.index[:2], inplace=True)
    len_row = df.columns[1::]

    for i in len_row:
        try:
            df.replace({i: ani_to_uni}, regex=True, inplace=True)
        except:
            pass

    for i in range(len(df)):
        data = df.values[i][1:-1:]
        new = {"": data}
        exc_data = pd.DataFrame(new, index=ls, )
        exc_data.to_excel('NEWS/Excel_{}.xlsx'.format(i))

        ####################################################
        if i == 0:
            for i in range(len(data)):
                try:
                    if i == 0:
                        x = str(data[i]).split()
                        ls_1["IS_INSURED_PHYSICAL"] = "1"
                except:
                    print("Error", data[i])

                try:
                    if i == 1:
                        if str(data[1]).split() == str(data[2]).split():
                            pass
                        else:
                            x = str(data[i]).split()
                            ls_1["PROPERTY_OWNER_PERSON_FIRST_NAME"] = x[0]
                            ls_1["PROPERTY_OWNER_PERSON_LAST_NAME"] = x[1]
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
                        x = str(data[i]).split()
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
                        x = str(data[i]).split()[0]
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

                    if i == 10:
                        pass
                    if i == 11:
                        pass

                try:
                    if i == 12:
                        x = str(data[12]).split()
                        ls_1["PROPERTY_NAME"] = x[0]
                        ls_1["PROPERTY_TYPE"] = x[0]
                except:
                    print("Error", data[i])

                try:
                    if i == 13:
                        x = str(data[13]).split()
                        ls_1["PROPERTY_MARKET_PRICE"] = x[0]
                except:
                    print("Error", data[i])

                try:
                    if i == 14:
                        x = str(data[14]).split()
                        ls_1['PROPERTY_RISK_AMOUNT'] = x[0]
                        ls_1["PROPERTY_SUM_INSURED"] = x[0]
                        risk = str(float(0.16) * int(x[0]) / 100).split('.')[0]
                        ls_1['PROPERTY_RISK_PREMIUM'] = risk
                        ls_1["POLICY_AMOUNT_CURRENCY"] = "AMD"
                        ls_1["PROPERTY_INSURANCE_RATE"] = "0.16"
                except:
                    print("Error", data[i])

                    if i == 15:
                        pass

                try:
                    if i == 16:
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
                                    ls_1["PROPERTY_COUNTRY"] = "ARM"
                                    ls_1["PROPERTY_REGION"] = j
                                    ls_1['PROPERTY_FULL_ADDRESS'] = new_data_1
                                    if j == "Տավուշ":
                                        ls_1['PROPERTY_CITY'] = "Այլ"
                                    else:
                                        ls_1['PROPERTY_CITY'] = j

                except:
                    if i == 16:
                        new_data_1 = str(new_data_1).strip().strip(",").strip()
                        ls_1["PROPERTY_COUNTRY"] = "ARM"
                        ls_1["PROPERTY_REGION"] = j
                        ls_1['PROPERTY_FULL_ADDRESS'] = new_data_1
                        if j == "Տավուշ":
                            ls_1['PROPERTY_CITY'] = "Այլ"
                        else:
                            ls_1['PROPERTY_CITY'] = j

                try:
                    if i == 17:
                        # Եթե սեփականատեր դաշտը լրացված է(գրված է սոց քարտ) գրում ենք այդ տվերը և BPR դաշտում գրում 1
                        x = str(data[i]).split()[0]
                        if len(x) > 5:
                            ls_1["PROPERTY_OWNER_PERSON_SOCIAL_CARD"] = x
                            ls_1["PROPERTY_OWNER_PERSON_PERS_BPR_USE"] = "1"
                        else:
                            # Եթե սեփականատեր դաշտը բաց է թողվաց գրում ենք սեփականատեր հանդիսանում է ապահովադիրը
                            ls_1["IS_OWNER_PERSON_SAME_INSURED"] = "1"
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

                        # Եթե նշված է օր վերցնում ենք հաջորդ ամսվա այդ օրը
                        start_date = datetime.date.today().strftime("%d.%m.%Y")
                        start = datetime.datetime.strptime(start_date, "%d.%m.%Y")
                        end_date = start + relativedelta(months=+1)
                        end = end_date.strftime("%d.%m.%Y")
                        payment = str(risk), end
                        ls_1["POLICY_CREATION_DATE"] = start_date
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

                        # Եթե նշված չի օր վերցնում ենք գալիք ամսվա 15ը
                        start_date = datetime.date.today().strftime("%d.%m.%Y")
                        start = datetime.datetime.strptime(start_date, "%d.%m.%Y")
                        end_date = start + relativedelta(months=+1)
                        end = end_date.strftime("15.%m.%Y")
                        payment = str(risk), end
                        ls_1["POLICY_CREATION_DATE"] = start_date
                        ls_1["POLICY_PAYMENT_SCHEDULE"] = ", ".join(payment)


    ##########################
    l = pd.Series(ls_1)
    l.to_json('NEWS/Json.json', indent=2, force_ascii=False)

    benef_data = pd.read_json('beneficiar.json', orient='index')[0]

    excel_data = pd.read_json('NEWS/Json.json', orient='index')[0]

    agent_data = pd.read_json('agent.json', orient='index')[0]

    result = pd.concat([benef_data, excel_data, agent_data])

    result.to_json('NEWS/New_Format.json', indent=2, force_ascii=False)


bank = Bank('Ineco', 'I1')
