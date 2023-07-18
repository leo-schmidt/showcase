import pandas as pd
pd.options.mode.chained_assignment = None
import numpy as np

def parse_excel(path):
    df = pd.read_excel(path, skiprows=15, usecols="A:F", skipfooter=2)

    # erste 2 Zeilen (leer) löschen
    df = df.drop([0,1])
    
    return df

def stack_columns(df):
    df_0_bis_3 = df[['Tiefe', 'Schläge']]
    df_3_bis_6 = df[['Tiefe.1', 'Schläge.1']].rename({'Tiefe.1' : 'Tiefe', 'Schläge.1' : 'Schläge'}, axis=1)
    df_6_bis_9 = df[['Tiefe.2', 'Schläge.2']].rename({'Tiefe.2' : 'Tiefe', 'Schläge.2' : 'Schläge'}, axis=1)

    df_stacked = pd.concat([df_0_bis_3, df_3_bis_6, df_6_bis_9], ignore_index=True).astype({'Schläge' : 'Int64', 'Tiefe' : float})

    return df_stacked

def write_csv(df, path):
    df.to_csv(path, index=False, sep=';')
    
def vorschachtung(df):
    """
    Erkennt Vorschachtung und fügt eine Null an letzter Stelle ein.
    """
    for index, row in df.iterrows():
        # Bereich 0.5 - 2.0 m
        # darüber und darunter nicht plausibel
        if index > 4 and index < 20:
            # Bedingung 1: vorherige Zeilen leer
            condition_1 = np.all(pd.isnull(df.loc[0:index-1, 'Schläge']))
            # Bedingung 2: aktuelle Zeile leer
            condition_2 = pd.isnull(df.loc[index, 'Schläge'])
            # Bedingung 3: nächste Zeile nicht leer
            condition_3 = not pd.isnull(df.loc[index+1, 'Schläge'])

            if condition_1 and condition_2 and condition_3:
                df.loc[index, 'Schläge'] = 0
                print('Vorschachtung bis ', df.loc[index, 'Tiefe'], 'm.')

    return df