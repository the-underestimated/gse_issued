import pandas as pd
from xls2xlsx import XLS2XLSX
import numpy as np
import io

def readProcess_Order(dataOrder, dataRaw_1, dataRaw_2, dataRaw_3, dataRaw_4):
    try:
        orderDetailTable = pd.read_csv(dataOrder)
        orderDetailTableMerging = orderDetailTable[['GRB_HISTORY', 'UNIT COST', 'CURRENCY ']]
        orderDetailTableMerging.rename({
            'UNIT COST':'INDEXED_PRICE',
            'CURRENCY ':'INDEXED_CURRENCY',
            'GRB_HISTORY':'GRB',
            }, axis=1, inplace=True)
        orderDetailTableMerging = orderDetailTableMerging.set_index('GRB')
        orderDetailTable2 = orderDetailTable.rename({
            'GRB_HISTORY':'GRB',
            'ORDER PN ':'PN'
            },axis=1)

        orderDetailTable2.dropna(axis='columns', inplace=True)
        orderDetailTable2.sort_values(by=['CREATED DATE', 'PN'], ascending=[False,False], inplace=True)

        priceUSD_IDR_OD = orderDetailTable2['UNIT COST'] * 15000
        priceEUR_IDR_OD = orderDetailTable2['UNIT COST'] * 17000
        priceJPY_IDR_OD = orderDetailTable2['UNIT COST'] * 100
        priceIDR_IDR_OD = orderDetailTable2['UNIT COST'] * 1

        orderDetailTable2.loc[orderDetailTable2['CURRENCY '] == 'USD', 'INDEXED_PRICE'] = priceUSD_IDR_OD[orderDetailTable2['CURRENCY '] == 'USD']
        orderDetailTable2.loc[orderDetailTable2['CURRENCY '] == 'EUR', 'INDEXED_PRICE'] = priceEUR_IDR_OD[orderDetailTable2['CURRENCY '] == 'EUR']
        orderDetailTable2.loc[orderDetailTable2['CURRENCY '] == 'JPY', 'INDEXED_PRICE'] = priceJPY_IDR_OD[orderDetailTable2['CURRENCY '] == 'JPY']
        orderDetailTable2.loc[orderDetailTable2['CURRENCY '] == 'IDR', 'INDEXED_PRICE'] = priceIDR_IDR_OD[orderDetailTable2['CURRENCY '] == 'IDR']

        orderDetailTable2.rename({'CURRENCY ':'INDEXED_CURRENCY'}, axis=1, inplace=True)

        PNComparisonTable = orderDetailTable2.iloc[:,[6,27,2]].drop_duplicates(subset=['PN']).reset_index()
        PNComparisonTable = PNComparisonTable.drop(['CREATED DATE', 'index'], axis=1)
        PNComparisonTable.rename({'INDEXED_PRICE':'PRICE'}, axis=1, inplace=True)

        
        # Change XLS to XLSX
        '''
        file1Rename = XLS2XLSX(dataRaw_1)
        file1Rename = file1Rename.to_xlsx()
        file1Read = pd.read_excel(dataRaw_1, engine='xlrd')

        file2Rename = XLS2XLSX(dataRaw_2)
        file2Rename = file2Rename.to_xlsx()
        file2Read = pd.read_excel(dataRaw_2, engine='xlrd')

        file3Rename = XLS2XLSX(dataRaw_3)
        file3Rename = file3Rename.to_xlsx()
        file3Read = pd.read_excel(dataRaw_3, engine='xlrd')

        file4Rename = XLS2XLSX(dataRaw_4)
        file4Rename = file4Rename.to_xlsx()
        file4Read = pd.read_excel(dataRaw_4, engine='xlrd')
        '''
        file1Rename = XLS2XLSX(dataRaw_1)
        file1Rename.to_xlsx("INVENTORY_REPORT1.xlsx")

        file2Rename = XLS2XLSX(dataRaw_2)
        file2Rename.to_xlsx("INVENTORY_REPORT2.xlsx")

        file3Rename = XLS2XLSX(dataRaw_3)
        file3Rename.to_xlsx("INVENTORY_REPORT3.xlsx")

        file4Rename = XLS2XLSX(dataRaw_4)
        file4Rename.to_xlsx("INVENTORY_REPORT4.xlsx")

        file1Table = pd.read_excel("INVENTORY_REPORT1.xlsx")
        file1Table.rename({'GOODS_RCVD_BATCH':'BATCH', 'BATCH':'GRB', 'Unnamed: 26': 'CONDITION', 'AC': 'REGISTRASI_GSE'}, axis=1, inplace=True)

        file2Table = pd.read_excel("INVENTORY_REPORT2.xlsx")
        file2Table.rename({'GOODS_RCVD_BATCH':'BATCH', 'BATCH':'GRB', 'Unnamed: 26': 'CONDITION', 'AC': 'REGISTRASI_GSE'}, axis=1, inplace=True)

        file3Table = pd.read_excel("INVENTORY_REPORT3.xlsx")
        file3Table.rename({'GOODS_RCVD_BATCH':'BATCH', 'BATCH':'GRB', 'Unnamed: 26': 'CONDITION', 'AC': 'REGISTRASI_GSE'}, axis=1, inplace=True)

        file4Table = pd.read_excel("INVENTORY_REPORT4.xlsx")
        file4Table.rename({'GOODS_RCVD_BATCH':'BATCH', 'BATCH':'GRB', 'Unnamed: 26': 'CONDITION', 'AC': 'REGISTRASI_GSE'}, axis=1, inplace=True)

        # Merging 4 file jadi satu
        mergedTable = pd.concat([file1Table, file2Table, file3Table, file4Table])
        mergedTable = mergedTable[mergedTable.REGISTRASI_GSE.notnull()]
        mergedTable = mergedTable.set_index('GRB')
        mergedTable = pd.merge(mergedTable, orderDetailTableMerging, on='GRB', how='outer', suffixes=('_left', '_right')).reset_index()
        mergedTable.dropna(subset=['REGISTRASI_GSE'], inplace=True)

        mergedTable['ISSUED_ITEM_PRICE'] = mergedTable['INDEXED_PRICE']
        mergedTable['ISSUED_ITEM_PRICE'] = mergedTable['ISSUED_ITEM_PRICE'].fillna(-1)

        # Nerge dengan comparisonn table dari order detail
        mergedTable = mergedTable.merge(PNComparisonTable, on='PN', how='left').reset_index()

        mergedTable['PRICE'] = mergedTable['PRICE'].where(mergedTable['ISSUED_ITEM_PRICE'] == -1)
        mergedTable.loc[mergedTable['ISSUED_ITEM_PRICE'] == -1, 'ISSUED_ITEM_PRICE'] = mergedTable['PRICE']
        mergedTable = mergedTable.drop(['PRICE', 'index'], axis=1)

        mergedTable['ISSUED_ITEM_PRICE'] = mergedTable.ISSUED_ITEM_PRICE.replace('', np.nan, regex=True)
        mergedTable['ISSUED_ITEM_PRICE'].fillna(-1, inplace=True)

        priceUSD_IDR = mergedTable['INDEXED_PRICE'] * 15000
        priceEUR_IDR = mergedTable['INDEXED_PRICE'] * 17000
        priceJPY_IDR = mergedTable['INDEXED_PRICE'] * 100

        mergedTable.loc[mergedTable.INDEXED_CURRENCY == 'USD', 'ISSUED_ITEM_PRICE'] = priceUSD_IDR
        mergedTable.loc[mergedTable.INDEXED_CURRENCY == 'EUR', 'ISSUED_ITEM_PRICE'] = priceEUR_IDR
        mergedTable.loc[mergedTable.INDEXED_CURRENCY == 'JPY', 'ISSUED_ITEM_PRICE'] = priceJPY_IDR

        def removeDash(df):
            return df.replace('-', ' ')

        mergedTable['REGISTRASI_GSE'] = mergedTable['REGISTRASI_GSE'].apply(removeDash)

        mergedTable['REGISTRASI_GSE_PREFIX'] = mergedTable['REGISTRASI_GSE'].astype(str).str[:3]

        mergedTable['JENIS_GSE'] = ''

        listGSE = {
            'TPB':'Aircraft Towing Tractor',
            'TR ':'Baggage Towing Tractor',
            'CBL':'Conveyor Belt Loader',
            'BCL':'Conveyor Belt Loader',
            'BCC':'Baggage Cart',
            'HLL':'High Lift Loader',
            'HCT':'Highlift Catering Truck',
            'MT ':'Maintenance Truck',
            'TNG':'Tangga Teknik',
            'TG ':'Tangga Teknik',
            'TB ':'Towbar',
            'WST':'Water Service Truck',
            'WSC':'Water Service Truck',
            'LST':'Lavatory Service Truck',
            'ACC':'Air Conditioning Unit',
            'GTC':'Air Starter Unit / Gas Turbine Compressor',
            'PDL':'Pallet Dolly',
            'TBL':'Telescopic Boom Lift',
            'ECW':'Compressor',
            'CMS':'Compressor',
            'CMP':'Compressor',
            'GPU':'Ground Power Unit',
            'GRB':'Garbarata',
            'APB':'Apron Passenger Bus',
            'GEN':'Genset',
            'BUS':'Apron Passenger Bus',
            'ATW':'Aircraft Towing Tractor',
            'FLT':'Forklift',
            'SR ':'Tangga Teknik',
            'PBS':'Passenger Boarding Stair',
            'GPB':'Ground Power Battery',
            'A19':'Baggage Towing Tractor',
            'A24':'Baggage Towing Tractor',
            'CON':'-',
            'TLM':'!Needs Recheck',
            'PK ':'-',
            'HS ':'-',
            'BC ':'Baggage Cart',
            'BCT':'Baggage Cart'
            }

        for key, value in listGSE.items():
            mergedTable.loc[mergedTable['REGISTRASI_GSE_PREFIX'] == key, 'JENIS_GSE'] = value.upper()

        mergedTable_mask = mergedTable.mask(mergedTable['JENIS_GSE'] == '')
        mergedTable['JENIS_GSE'] = mergedTable_mask['JENIS_GSE'].fillna("Airside Operation Vehicle".upper())
        mergedTable['PRICE_X_QTY'] = mergedTable['ISSUED_ITEM_PRICE'] * (mergedTable['QTY'] - mergedTable['QTY_RETURN_STOCK'])

        if mergedTable.CREATED_DATE.dtype == '<M8[ns]':
            mergedTable['CREATED_DATE'] = mergedTable['CREATED_DATE'].dt.date

        dataRaw = mergedTable.iloc[:,[31,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,0,16,17,18,19,21,22,23,24,25,26,29,32]]
        dataRaw.rename({
            'CREATED_DATE':'ISSUED_DATE',
            'QTY':'QTY_ISSUED'
            }, axis=1, inplace=True)
        dataRaw.sort_values(by=['ISSUED_DATE', 'REGISTRASI_GSE', 'BATCH'], ascending=[True, True, True], inplace=True)
        dataRaw.reset_index(drop=True, inplace=True)

        dataProcessed = dataRaw.iloc[:,[0,14,18,15,6,27,21,20,28,19]]

        oldestDate = mergedTable['CREATED_DATE'].min()
        newestDate = mergedTable['CREATED_DATE'].max()

        return dataRaw, dataProcessed, oldestDate, newestDate

        #with pd.ExcelWriter('DATA_%s_%s.xlsx' %(oldestDate,newestDate), date_format='m/d/yyyy', datetime_format='m/d/yyyy HH:MM:SS', engine='xlsxwriter') as writer:
        #    dataRaw.to_excel(writer, sheet_name='DATA_RAW', index=False)
        #    dataProcessed.to_excel(writer, sheet_name='DATA_PROCESSED', index=False)
                
    except Exception as e:
        raise ValueError(e)