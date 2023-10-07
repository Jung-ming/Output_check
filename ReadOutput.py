import pandas as pd


def Read_Output(Output輸出檔):
    工作表名稱 = pd.ExcelFile(Output輸出檔).sheet_names
    if '匯出檔' in 工作表名稱:
        Output = pd.read_excel(Output輸出檔)
    else:
        Output_DIP = pd.read_excel(Output輸出檔)
        Output_SMT = pd.read_excel(Output輸出檔, sheet_name='SMT')
        Output = pd.concat(
            [Output_DIP[['母工單單號', '工號', '名稱規格', '工令量', 'SOURCE']], Output_SMT[['母工單單號', '工號', '名稱規格', '工令量', 'SOURCE']]],
            ignore_index=True)
        Output = Output.assign(尾數=None, 移轉小記=None, 總計=None, 餘數=None, 註記=None)

    return Output


def Read_DIP(DIP_file):
    DIP_Sheet1 = pd.read_excel(DIP_file)
    DIP_Sheet1.drop(['    '], axis=1)
    columns = ['工號', '名稱', '工單', 'MO', '工令量', 'SOURCE ', 'Unnamed: 6', '開始時間',
                          'DIP首件產出時間/數量', 'OUTPUT', '製程', '尾數 ', '移轉小記', '總計', '餘數',
                          'Unnamed: 15']
    DIP_Sheet1.columns = columns
    末位足標 = DIP_Sheet1.index[-1]

    DIP_四零四內帳 = pd.read_excel(DIP_file, header=1, sheet_name='四零四內帳')
    DIP_四零四TEST = pd.read_excel(DIP_file, header=1, sheet_name='四零四TEST')
    DIP_TEST測試部 = pd.read_excel(DIP_file, header=1, sheet_name='TEST測試部')
    DIP_ASSY組裝部 = pd.read_excel(DIP_file, header=1, sheet_name='ASSY組裝部')
    待檢查工作表 = [DIP_四零四內帳, DIP_四零四TEST, DIP_TEST測試部, DIP_ASSY組裝部]

    for 工作表 in 待檢查工作表:
        for 足標, 欄位 in 工作表.iterrows():
            if 欄位['工號'] in DIP_Sheet1['工號'].tolist():
                # 這裡將待檢查的工作表進行迭代，取得每一行的資料，如果有找到一樣的工單，就利用工單反查原本的對應足標
                # 然後將迭代的資料蓋過對應足標的資料
                對應足標 = DIP_Sheet1[DIP_Sheet1['工號'] == 欄位['工號']].index[0]
                DIP_Sheet1.at[對應足標, '尾數 '] = 欄位['尾數 ']
                DIP_Sheet1.at[對應足標, '移轉小記'] = 欄位['移轉小記']
                DIP_Sheet1.at[對應足標, '總計'] = 欄位['總計']
                DIP_Sheet1.at[對應足標, '餘數'] = 欄位['餘數']
                if 'Unnamed: 15' in list(工作表.columns):
                    DIP_Sheet1.at[對應足標, 'Unnamed: 15'] = 欄位['Unnamed: 15']
            else:
                末位足標 += 1
                DIP_Sheet1.at[末位足標, '工號'] = 欄位['工號']
                DIP_Sheet1.at[末位足標, '名稱'] = 欄位['名稱']
                DIP_Sheet1.at[末位足標, 'MO'] = 欄位['MO']
                DIP_Sheet1.at[末位足標, '工令量'] = 欄位['工令量']
                DIP_Sheet1.at[末位足標, 'SOURCE '] = 欄位['SOURCE ']
                DIP_Sheet1.at[末位足標, '尾數 '] = 欄位['尾數 ']
                DIP_Sheet1.at[末位足標, '移轉小記'] = 欄位['移轉小記']
                DIP_Sheet1.at[末位足標, '總計'] = 欄位['總計']
                DIP_Sheet1.at[末位足標, '餘數'] = 欄位['餘數']
                if 'Unnamed: 15' in list(工作表.columns):
                    DIP_Sheet1.at[末位足標, 'Unnamed: 15'] = 欄位['Unnamed: 15']

    # 直接將工號設定為索引，這樣就能直接用工號去找特定的資料，不用考慮足標的問題
    return DIP_Sheet1[['工號', '名稱', '工令量', 'SOURCE ', '尾數 ', '移轉小記', '總計', '餘數',
                       'Unnamed: 15']].set_index('工號')


if __name__ == "__main__":
    Output = Read_Output('Output輸出檔1005.xlsx')
    DIP = Read_DIP('112.10.5.xlsx')
    for 足標, 欄位 in Output.iterrows():
        if 欄位['母工單單號'] in DIP.index:
            Output.at[足標, '尾數'] = DIP.loc[欄位['母工單單號'], '尾數 ']
            Output.at[足標, '移轉小記'] = DIP.loc[欄位['母工單單號'], '移轉小記']
            Output.at[足標, '總計'] = DIP.loc[欄位['母工單單號'], '總計']
            Output.at[足標, '餘數'] = DIP.loc[欄位['母工單單號'], '餘數']
            Output.at[足標, '註記'] = DIP.loc[欄位['母工單單號'], 'Unnamed: 15']
        elif 欄位['工號'] in DIP.index:
            Output.at[足標, '尾數'] = DIP.loc[欄位['工號'], '尾數 ']
            Output.at[足標, '移轉小記'] = DIP.loc[欄位['工號'], '移轉小記']
            Output.at[足標, '總計'] = DIP.loc[欄位['工號'], '總計']
            Output.at[足標, '餘數'] = DIP.loc[欄位['工號'], '餘數']
            Output.at[足標, '註記'] = DIP.loc[欄位['工號'], 'Unnamed: 15']
        else:
            Output.at[足標, '尾數'] = '查無資料'
    Output.to_excel('比對結果.xlsx', index=False)
