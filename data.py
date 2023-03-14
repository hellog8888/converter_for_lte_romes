import os
import pandas as pd
from transliterate import translit

dict_for_operator = \
    {
        'Открытое акционерное общество \"Мобильные ТелеСистемы\"': 1,
        'Публичное акционерное общество \"Мобильные ТелеСистемы\"': 1,
        'Открытое акционерное общество \"МегаФон\"': 2,
        'Публичное акционерное общество \"МегаФон\"': 2,
        'Общество с ограниченной ответственностью «Скартел»': 2,
        'Общество с ограниченной ответственностью \"Скартел\"': 2,
        'Публичное акционерное общество \"Ростелеком\"': 20,
        'Публичное акционерное общество «Ростелеком»': 20,
        'Общество с ограниченной ответственностью \"Т2 Мобайл\"': 20,
        'Открытое акционерное общество \"Вымпел-Коммуникации\"': 99,
        'Публичное акционерное общество \"Вымпел-Коммуникации\"': 99,
    }

def toFixed(numObj, digits=0):
    return f"{numObj:.{digits}f}"

count_ = 0
def count_for_uniq():
    global count_
    count_ += 1
    return count_

def dms2dd(dd):
    N = dd[:2] + ' ' + dd[3:5] + ' ' + dd[5:7]
    E = dd[8:10] + ' ' + dd[11:13] + ' ' + dd[13:]
    d1, m1, s1 = N.split(' ')
    d2, m2, s2 = E.split(' ')
    try:
        return f'{toFixed((int(d2) + int(m2) / 60 + int(s2) / 3600), 6)} {toFixed((int(d1) + int(m1) / 60 + int(s1) / 3600), 6)}'
    except ValueError:
        pass

def first_stage():

    file_LTE_all = pd.read_excel('LTE_ALL.xls')
    LTE_db = file_LTE_all.loc[:, ['Свидетельство', 'Организация', 'Широта/ Долгота', 'Идентификационный номер РЭС в сети связи', 'Место установки']]

    for i in range(len(LTE_db['Организация'])):
        if LTE_db['Организация'][i] in dict_for_operator:
            LTE_db['Организация'][i] = dict_for_operator[LTE_db['Организация'][i]]

    LTE_db['Широта/ Долгота'] = [dms2dd(str(x)) for x in LTE_db['Широта/ Долгота']]
    LTE_db['PosLatitude'] = LTE_db['Широта/ Долгота'].str[:9]
    LTE_db['PosLongitude'] = LTE_db['Широта/ Долгота'].str[10:]
    LTE_db = LTE_db.drop(['Широта/ Долгота'], axis=1)

    LTE_db['Идентификационный номер РЭС в сети связи'] = \
        [str(x).replace('MCC, MNC, eNB ID, Cell ID: ', '').replace('CI(ECI): ', '').replace('CI (ECI): ', '').replace('MAC:', '').replace(';', ',')
         for x in LTE_db['Идентификационный номер РЭС в сети связи']]
    LTE_db['Идентификационный номер РЭС в сети связи'] = LTE_db['Идентификационный номер РЭС в сети связи'].str.replace(' ', '')

    LTE_db['Свидетельство'] = LTE_db['Свидетельство'].str.replace('  № ', '-')
    LTE_db['Свидетельство'] = LTE_db['Свидетельство'].str.replace(' ', '-')
    LTE_db['Место установки'] = LTE_db['Место установки'].str.replace('  ', '')
    LTE_db['Место установки'] = LTE_db['Место установки'].str.replace('.,', ',', regex=False)
    LTE_db['Место установки'] = LTE_db['Место установки'].str.replace('\"\"', '')
    LTE_db['Место установки'] = LTE_db['Место установки'].str.replace('\"', '')
    LTE_db['Место установки'] = LTE_db['Место установки'].str.replace('Краснодарский край, ', '')
    LTE_db['Место установки'] = LTE_db['Место установки'].str.replace('Адыгея Респ, ', '')
    LTE_db['Место установки'] = LTE_db['Место установки'].str.replace('Республика Адыгея (Адыгея), ', '', regex=False)
    LTE_db['Место установки'] = [str(translit(x, 'ru', reversed=True)) for x in LTE_db['Место установки']]
    LTE_db['eNodeB_Name'] = LTE_db['Свидетельство'] + '--' + LTE_db['Место установки']
    LTE_db = LTE_db.drop(['Свидетельство'], axis=1)
    LTE_db = LTE_db.drop(['Место установки'], axis=1)

    LTE_db['Идентификационный номер РЭС в сети связи'] = LTE_db['Идентификационный номер РЭС в сети связи'].str.split(',')
    LTE_db = LTE_db.explode('Идентификационный номер РЭС в сети связи')

    LTE_db = LTE_db[['eNodeB_Name', 'PosLongitude', 'PosLatitude', 'Организация', 'Идентификационный номер РЭС в сети связи']]
    LTE_db.rename(columns={'Организация': 'MNC', 'Идентификационный номер РЭС в сети связи': 'CellID'}, inplace=True)
    LTE_db['UniqueID'] = '0'
    LTE_db['UniqueID'] = [count_for_uniq() for _ in LTE_db['UniqueID']]
    LTE_db['MCC'] = '250'
    LTE_db['PosAltitude'] = '0.00'
    LTE_db['Power'] = '0.0000'
    LTE_db['PosErrorLargeHalfAxis'] = '0.00'
    LTE_db['PosErrorSmallHalfAxis'] = '0.00'
    LTE_db['PosErrorDirection'] = '0.00'
    LTE_db['Direction'] = '0.00'
    LTE_db['IsDirected'] = '1'
    LTE_db['3GNC'] = ''
    LTE_db['2GNC'] = ''
    LTE_db['4GNC'] = ''
    LTE_db['EARFCN'] = '0'
    LTE_db['PhyCellID'] = '0'
    LTE_db = LTE_db[['UniqueID', 'eNodeB_Name', 'MNC', 'MCC', 'PosLatitude', 'PosLongitude', 'PosAltitude', 'Power', 'PosErrorLargeHalfAxis', 'PosErrorSmallHalfAxis', 'PosErrorDirection', 'Direction', 'IsDirected', '3GNC', '2GNC', '4GNC', 'EARFCN', 'CellID', 'PhyCellID']]

    LTE_db.to_csv('LTE_R&S LTE Scanner_[1]_trans_list_[].csv', sep=';', index=False)

def second_stage():
    with open('LTE_R&S LTE Scanner_[1]_trans_list_[].csv', mode="r", encoding='utf-8') as csv_file, open('LTE_R&S LTE Scanner_[1]_trans_list_[].txt', mode='w', encoding='utf-8') as txt_file:
        csv_file.seek(195)
        txt_file.write(csv_file.read())

def three_stage():
    with open('LTE_R&S LTE Scanner_[1]_trans_list_[].atd', mode="w", encoding='utf-8') as adt_file:
        header_rows_db = """[Main]
Type=ATD
[Table1]
Name=LTE_PNS_DATABASE
File=LTE_R&S LTE Scanner_[1]_trans_list_[].txt
Columns_Size=19
Columns0_Name=UniqueID
Columns0_Type=utULInt
Columns1_Name=eNodeB_Name
Columns1_Type=utDynChar
Columns2_Name=MNC
Columns2_Type=utULInt
Columns3_Name=MCC
Columns3_Type=utULInt
Columns4_Name=PosLongitude
Columns4_Type=utDouble
Columns5_Name=PosLatitude
Columns5_Type=utDouble
Columns6_Name=PosAltitude
Columns6_Type=utDouble
Columns7_Name=Power
Columns7_Type=utDouble
Columns8_Name=PosErrorLargeHalfAxis
Columns8_Type=utDouble
Columns9_Name=PosErrorSmallHalfAxis
Columns9_Type=utDouble
Columns10_Name=PosErrorDirection
Columns10_Type=utDouble
Columns11_Name=Direction
Columns11_Type=utUSInt
Columns12_Name=IsDirected
Columns12_Type=utUTInt
Columns13_Name=3GNC
Columns13_Type=utDynChar
Columns14_Name=2GNC
Columns14_Type=utDynChar
Columns15_Name=4GNC
Columns15_Type=utDynChar
Columns16_Name=EARFCN
Columns16_Type=utULInt
Columns17_Name=CellID
Columns17_Type=utULInt
Columns18_Name=PhyCellID
Columns18_Type=utUSInt
"""

        adt_file.write(header_rows_db)
        os.remove('LTE_R&S LTE Scanner_[1]_trans_list_[].csv')
        print('[Result file successfully]')

if __name__ == '__main__':
    first_stage()
    second_stage()
    three_stage()