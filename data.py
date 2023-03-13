import csv
import pandas as pd
from transliterate import translit, get_available_language_codes

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
    #open
    file_LTE_all = pd.read_excel('LTE_ALL.xls')
    LTE_db = file_LTE_all.loc[:, ['Свидетельство', 'Организация', 'Широта/ Долгота', 'Идентификационный номер РЭС в сети связи', 'Место установки']]

    #Организация
    for i in range(len(LTE_db['Организация'])):
        if LTE_db['Организация'][i] in dict_for_operator:
            LTE_db['Организация'][i] = dict_for_operator[LTE_db['Организация'][i]]

    #Широта/ Долгота
    LTE_db['Широта/ Долгота'] = [dms2dd(str(x)) for x in LTE_db['Широта/ Долгота']]
    LTE_db['Latitude'] = LTE_db['Широта/ Долгота'].str[:9]
    LTE_db['Longitude'] = LTE_db['Широта/ Долгота'].str[10:]
    LTE_db = LTE_db.drop(['Широта/ Долгота'], axis=1)

    #Идентификационный номер РЭС в сети связи
    LTE_db['Идентификационный номер РЭС в сети связи'] = \
        [str(x).replace('MCC, MNC, eNB ID, Cell ID: ', '').replace('CI(ECI): ', '').replace('CI (ECI): ', '').replace('MAC:', '').replace(';', ',')
         for x in LTE_db['Идентификационный номер РЭС в сети связи']]
    LTE_db['Идентификационный номер РЭС в сети связи'] = LTE_db['Идентификационный номер РЭС в сети связи'].str.replace(' ', '')

    #Свидетельство + Место установки
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
    LTE_db['Name'] = LTE_db['Свидетельство'] + '--' + LTE_db['Место установки']
    LTE_db = LTE_db.drop(['Свидетельство'], axis=1)
    LTE_db = LTE_db.drop(['Место установки'], axis=1)

    #split_CI_colums
    LTE_db['Идентификационный номер РЭС в сети связи'] = LTE_db['Идентификационный номер РЭС в сети связи'].str.split(',')
    LTE_db = LTE_db.explode('Идентификационный номер РЭС в сети связи')

    #colums_create
    LTE_db = LTE_db[['Name', 'Longitude', 'Latitude', 'Организация', 'Идентификационный номер РЭС в сети связи']]
    LTE_db.rename(columns={'Организация': 'MNC', 'Идентификационный номер РЭС в сети связи': 'CellIdentity'}, inplace=True)
    LTE_db['PosErrorDirection'] = '0'
    LTE_db['PosErrorLambda1'] = '0'
    LTE_db['PosErrorLambda2'] = '0'
    LTE_db['IsDirected'] = '0'
    LTE_db['Direction'] = '0'
    LTE_db['Power'] = '0'
    LTE_db['MaxPowerUsedForTowerEstimationbyPE'] = '0'
    LTE_db['TowerID'] = '0'
    LTE_db['MCC'] = '250'
    LTE_db['TAC'] = '0'
    LTE_db['EARFCN'] = '0'
    LTE_db['PhyCellID'] = '0'
    LTE_db = LTE_db[['Name', 'Longitude', 'Latitude', 'PosErrorDirection', 'PosErrorLambda1', 'PosErrorLambda2', 'IsDirected', 'Direction', 'Power', 'MaxPowerUsedForTowerEstimationbyPE', 'TowerID', 'MCC', 'MNC', 'TAC', 'CellIdentity', 'EARFCN', 'PhyCellID']]

    #write_to_csv_file
    LTE_db.to_csv('LTE_R&S LTE Scanner_[1]_trans_list_[].csv', sep=';', index=False)

def second_stage():
    with open('LTE_R&S LTE Scanner_[1]_trans_list_[].csv', mode="r+", encoding='utf-8') as trans_list_file:
        header_rows = """;exported by ROMES LTE BCCH View
;The name is either the name of the station found in the database or in the case no station was found in the data base "eNodeB <eNodeBID/CellID> <PhyCellID>
;Postion of 0.0/0.0 means unknown
;Only Stations with a valid Cell Id are exported.
;
"""
        data = trans_list_file.read()
        trans_list_file.seek(0)
        trans_list_file.write(header_rows + ";" + data)

def three_stage():
    with open('LTE_R&S LTE Scanner_[1]_trans_list_[].atd', mode="w", encoding='utf-8') as adt_file:
        header_rows_db = """; This ATD File defines the LTE BTS Table in the LTE BTS DataBase
;
;
[Main]
Type=ATD
[Table1]
Name=BTS_TABLE
File=LTE_R&S LTE Scanner_[1]_trans_list_[].txt
Columns_Size=17
;
;
; Cell Name
Columns0_Name=Name
Columns0_Type=utDynChar
;
; The Position Column Names (Pos...) must be given in this writing
; The Names are used by the data base software to make the special
; Column "Position with Error"
Columns1_Name=PosLongitude
Columns1_Type=utDouble
Columns2_Name=PosLatitude
Columns2_Type=utDouble
Columns3_Name=PosErrorDirection
Columns3_Type=utDouble
Columns4_Name=PosErrorLambda1
Columns4_Type=utDouble
Columns5_Name=PosErrorLambda2
Columns5_Type=utDouble
Columns6_Name=IsDirected
Columns6_Type=utUTInt
Columns7_Name=Direction
Columns7_Type=utUSInt
Columns8_Name=Power
Columns8_Type=utDouble
Columns9_Name=MaxPowerUsedForTowerEstimationbyPE
Columns9_Type=utDouble
Columns10_Name=TowerID
Columns10_Type=utUSInt
Columns11_Name=MCC
Columns11_Type=utUSInt
Columns12_Name=MNC
Columns12_Type=utUSInt
Columns13_Name=TAC
Columns13_Type=utUSInt
Columns14_Name=CellID
Columns14_Type=utULInt
Columns15_Name=EARFCN
Columns15_Type=utULInt
Columns16_Name=PhyCellID
Columns16_Type=utUSInt
"""
        adt_file.write(header_rows_db)
        print('[Result file successfully]')

if __name__ == '__main__':
    first_stage()
    second_stage()
    three_stage()
