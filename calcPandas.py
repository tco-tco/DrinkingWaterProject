import pandas as pd
import mysql.connector


def avarageCalc(tesisAdi,id,iladi,locgiris):
    connection=mysql.connector.connect(host="localhost", user= "root",password="password", database=f"{iladi}")
    df=pd.read_sql_query("SELECT * FROM {tesisadi} where ID='{ID}'".format(tesisadi=tesisAdi,ID=id), connection)


    #dfDrop=df.drop(["ID", "Parametre_Yonetmelik", "Parametre_Ölçüm", "CAS_NO", "Birim", "A1", "A2", "A3", "Giriş_Ortalama", "Giriş_Sonuç", "Çıkış_Ortalama", "Çıkış_Sonuç"], axis=1)
    dfDrop=df.drop(df.iloc[:,0:13], axis=1)

    dfDropt=dfDrop.drop(dfDrop.iloc[:,(int(f"{locgiris}")-1):], axis=1)

    dfDot=dfDropt.dropna(axis=1, how='all')

    #dfDot=dfDot.stack().str.replace(',','.').unstack()

    dfDot= dfDot.astype(float, errors = 'raise')

    # dataTypeSeries = dfDot.dtypes
    # print('Data type of each column of Dataframe :')
    # print(dataTypeSeries)

    dfDot['mean'] = (dfDot.mean(axis=1)).round(decimals = 10)
    
    if len(str(dfDot['mean'].values[0]))!=0:
        return dfDot['mean'].values[0]



def avarageCalcCıkıs(tesisAdi,id,iladi,locgiris):
    connection=mysql.connector.connect(host="localhost", user= "root",password="password", database=f"{iladi}")
    df=pd.read_sql_query("SELECT * FROM {tesisadi} where ID='{ID}'".format(tesisadi=tesisAdi,ID=id), connection)


    #dfDrop=df.drop(["ID", "Parametre_Yonetmelik", "Parametre_Ölçüm", "CAS_NO", "Birim", "A1", "A2", "A3", "Giriş_Ortalama", "Giriş_Sonuç", "Çıkış_Ortalama", "Çıkış_Sonuç"], axis=1)
    dfDrop=df.drop(df.iloc[:,:int(f"{13+locgiris}")],axis=1)

 
    dfDot=dfDrop.dropna(axis=1, how='all')
    #dfDot=dfDot.stack().str.replace(',','.').unstack()


    dfDot= dfDot.astype(float, errors = 'raise')




    # dataTypeSeries = dfDot.dtypes
    # print('Data type of each column of Dataframe :')
    # print(dataTypeSeries)

    dfDot['mean'] = (dfDot.mean(axis=1)).round(decimals = 10)
    

    if len(str(dfDot['mean'].values[0]))!=0:
        return dfDot['mean'].values[0]

