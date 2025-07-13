import pandas as pd
import numpy as np
from pprint import pprint
from datetime import date

df = pd.read_excel('phB12E.xlsx')

# setup dátumu a názvu súboru
currentDate = date.today().strftime('%d.%m.%Y')
nazov_suboru = f'ABC_25_{currentDate}.xlsx'

# Transform
#----------------------------------------------------------------------------------
# vytiahnutie karty
df = df.loc[ df['Typ'] == 'Karta']

# vytiahnutie vydaných faktúr
df = df.loc[ df['Agenda'] == 'Vydané faktúry']

# tvorba stĺpca mesiac
df['Mesiac'] = df['Dátum'].dt.month_name()

# výpočet obratu
df['Obrat'] = round(df['Množstvo'] * df['Čiastka'], 2)

# produkty s kladným predajom ale záporným ziskom - nový df
zapornyPredaj = df.loc[ (df['Množstvo'] > 0) & (df['Zisk'] < 0) ]

# odfiltrovanie produktov so záporným ziskom
df = df.drop(df.loc[ (df['Množstvo'] > 0) & (df['Zisk'] < 0) ].index)

# odstránenie darčekov
df = df.drop(df.loc[ df['Kód'].str.contains('gift', case=False) ].index)

# odstránenie zápisov bez krajiny - testy a pod
df = df.drop(df.loc[ df['Krajina'].isna() ].index)

# extrakcia typu produktu
df[['Sklad', 'Typ_produktu']] = df['Členenie'].str.split('/', expand=True)

df[['Split_1','Split_2','Split_3','Split_4','Split_5']] = df['Kód'].str.split('-', expand=True)


# Transform - vytiahnutie mastrov z kodov produktu
#----------------------------------------------------------------------------------
# zvysok dodávateľov
df['Master'] = df['Split_1'] + '-' + df['Split_2']

# Dodavatel 1
df.loc[ df['Dodávateľ'] == 'Dodavatel 1', 'Master' ] = (df['Split_2'] + '-' + df['Split_3'])

# Dodavatel 2
df.loc[ df['Dodávateľ'] == 'Dodavatel 2', 'Master' ] = (df['Split_2'] + '-' + df['Split_3'])

# Dodavatel 3
df.loc[ df['Dodávateľ'] == 'Dodavatel 3', 'Master' ] = (df['Split_1'] + '-' + df['Split_3'])

# Dodavatel 4
df.loc[ df['Dodávateľ'] == 'Dodavatel 4', 'Master' ] = (df['Split_1'] + '-' + df['Split_2'] + '-' + df['Split_3'])

# Dodavatel 5 - ak Split_3 chýba -> kratší kód
df.loc[
    (df['Dodávateľ'] == 'Dodavatel 5') & (df['Split_3'].isna()),
    'Master'
] = df['Split_1']

# Dodavatel 5 - ak Split_3 existuje -> dlhší kód
df.loc[
    (df['Dodávateľ'] == 'Dodavatel 5') & (df['Split_3'].notna()),
    'Master'
] = df['Split_1'] + '-' + df['Split_2']

# Dodavatel 6 - ošetrenie neštandardneho kódu produktu
# Dodavatel 6 - ak Split_5 chýba -> kratší kód
df.loc[
    (df['Dodávateľ'] == 'Dodavatel 6') & (df['Split_5'].isna()),
    'Master'
] = df['Split_1'] + '-' + df['Split_2']

# Dodavatel 6 - ak Split_5 existuje -> nejedná sa o 02-S
df.loc[
    (df['Dodávateľ'] == 'Dodavatel 6') & (df['Split_3'].notna()),
    'Master'
] = df['Split_1'] + '-' + df['Split_2'] + '-' + df['Split_3']

# tvorba nového DF ktorý obsahuje iba potrebné stlpce pre ABC
finalData = df[['Krajina', 'Dodávateľ', 'Číslo', 'Typ_produktu', 'Kód', 'Master', 'Dátum', 'Mesiac', 'Množstvo', 'Obrat', 'Zisk']]

# export + celková ABC analýza po mesiacoch
# ---------------------------------------------------------------------------------------------
# získanie mesiaca
mesiace = finalData['Mesiac'].unique()

# definícia kvartálu
kvartaly = {
    '1Q': ['January', 'February', 'March'],
    '2Q': ['April', 'May', 'June'],
    '3Q': ['July', 'August', 'September'],
    '4Q': ['October', 'November', 'December']
}


with pd.ExcelWriter(nazov_suboru) as writer:
    for mesiac in mesiace:
        data_mesiac = finalData.loc[ finalData['Mesiac'] == mesiac ]

        tab = data_mesiac.groupby(['Master', 'Dodávateľ']).agg({
        'Množstvo': 'sum',
        'Obrat': 'sum',
        'Zisk': 'sum'
        })

        # výpočet podielu na obrate
        tab['Obrat_podiel'] = round((tab['Obrat'] / tab['Obrat'].sum()) * 100, 2)

        # výpočet podielu na zisku
        tab['Zisk_podiel'] = round((tab['Zisk'] / tab['Zisk'].sum()) * 100, 2)

        # pridanie kumulatívneho vypoctu pre OBRAT + ABC flagu
        tab = tab.sort_values(by='Obrat_podiel', ascending=False)
        tab['Obrat_cum'] = tab['Obrat_podiel'].cumsum()

        # pridanie kumulatívneho vypoctu pre ZISK
        tab = tab.sort_values(by='Zisk_podiel', ascending=False)
        tab['Zisk_cum'] = tab['Zisk_podiel'].cumsum()

        # pridanie flagov
        tab['ABC_obrat'] = tab['Obrat_cum'].apply(lambda x: 'A' if x <= 80 else( 'B' if x <=95 else 'C' ))
        tab['ABC_zisk'] = tab['Zisk_cum'].apply(lambda x: 'A' if x <= 80 else( 'B' if x <=95 else 'C' ))

        # výstupné tabuľky
        summary1 = tab.groupby('ABC_zisk').agg({
            'Obrat': 'sum',
            'Zisk': 'sum',
            'Obrat_cum': 'count'
        })

        summary2 = pd.crosstab(tab.reset_index()['Dodávateľ'], tab.reset_index()['ABC_zisk'])

        # názvy sheetov 
        nazov = mesiac[:3].capitalize()
        tab.to_excel(writer, sheet_name=f'{nazov}', index=True)
        summary1.to_excel(writer, sheet_name=f'{nazov}_summary', startrow=0)
        summary2.to_excel(writer, sheet_name=f'{nazov}_summary', startrow=6)

    # výpočet po Q
    for nazov_q, mesiace_q in kvartaly.items():
        # vyber len dáta pre daný kvartál
        data_q = finalData[finalData['Mesiac'].isin(mesiace_q)]

        if data_q.empty:
            continue

        tab = data_q.groupby(['Master', 'Dodávateľ']).agg({
            'Množstvo': 'sum',
            'Obrat': 'sum',
            'Zisk': 'sum'
        })

        tab['Obrat_podiel'] = round((tab['Obrat'] / tab['Obrat'].sum()) * 100, 2)
        tab['Zisk_podiel'] = round((tab['Zisk'] / tab['Zisk'].sum()) * 100, 2)

        tab = tab.sort_values(by='Obrat_podiel', ascending=False)
        tab['Obrat_cum'] = tab['Obrat_podiel'].cumsum()

        tab = tab.sort_values(by='Zisk_podiel', ascending=False)
        tab['Zisk_cum'] = tab['Zisk_podiel'].cumsum()

        tab['ABC_obrat'] = tab['Obrat_cum'].apply(lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C'))
        tab['ABC_zisk'] = tab['Zisk_cum'].apply(lambda x: 'A' if x <= 80 else ('B' if x <= 95 else 'C'))

        summary1 = tab.groupby('ABC_zisk').agg({
            'Obrat': 'sum',
            'Zisk': 'sum',
            'Obrat_cum': 'count'
        })

        summary2 = pd.crosstab(tab.reset_index()['Dodávateľ'], tab.reset_index()['ABC_zisk'])

        tab.to_excel(writer, sheet_name=f'{nazov_q}')
        summary1.to_excel(writer, sheet_name=f'{nazov_q}_summary', startrow=0)
        summary2.to_excel(writer, sheet_name=f'{nazov_q}_summary', startrow=6)
