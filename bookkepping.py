import pandas as pd
import datetime

class bookkepping:
    def __init__(self, path, col = None):
        print('读取文本数据...')
        try:
            self.book = pd.read_excel(path, sheet_name=col)
            for sheet_name, df in self.book.items():
                print(f'读取{sheet_name}')
        except:
            self.book = dict()
            print(f'没有找到文件{path}，会新建一个')
        print('*' * 50)
        
    def update_detail(self, sheet_name, sheet_date, detail_path, detail_date, start, end):
        print('更新明细表...')
        if sheet_name:
            book_detail = self.book[sheet_name]
            book_detail[sheet_date] = pd.to_datetime(book_detail[sheet_date]).dt.date
            book_check = book_detail[(book_detail[sheet_date] >= start) & (book_detail[sheet_date] <= end)]
        
        detail = pd.read_csv(detail_path, encoding = 'utf-16LE', on_bad_lines='skip', sep = '\t')
        detail['金额'] = pd.to_numeric(detail['金额'].apply(lambda x: x.replace(',', '') if isinstance(x, str) else x))
        detail.loc[:, '金额'] = pd.to_numeric(detail.loc[:, '金额'])
        if not set(detail['类型'].dropna().unique()).issubset({'收入', '支出', '转账'}): 
            raise Exception('请检查明细表类型')
        detail.loc[:,detail_date] = pd.to_datetime(detail[detail_date]).dt.date
        
        detail_io = detail[detail['类型'].isin(['收入', '支出'])]
        detail_transfer =  detail[detail['类型'].isin(['转账'])]
        detail_io = detail_io[['日期', '类型', '货币', '金额', '分类','账户', '备注']]
        detail_io = detail_io[(detail_io[detail_date] >= start) & (detail_io[detail_date] <= end)]
        
        new_detail = detail_io[~detail_io.apply(tuple, axis=1).isin(book_check.apply(tuple, axis=1))]
        print(f'新添明细数据\n{new_detail}')
        book_detail = pd.concat([book_detail, new_detail], ignore_index=True)
        book_detail.sort_values(sheet_date, ascending=False, inplace=True)
        self.book[sheet_name] = book_detail
        
        print('更新账户数据...')
        account = self.book['账户']
        detail_in = detail_transfer.copy()
        detail_in.loc[:, '账户'] = detail_in.loc[:, '付款']
        detail_in.loc[:, '金额'] = -detail_in.loc[:, '金额']
        detail_in = detail_in[['日期', '货币', '金额', '账户']]
        detail_out = detail_transfer.copy()
        detail_out.loc[:, '账户'] = detail_out.loc[:, '收款']
        detail_out = detail_out[['日期', '货币', '金额', '账户']]
        new_detail = pd.concat([new_detail[['日期', '货币', '金额', '账户']], detail_in, detail_out])
        
        account_update = new_detail.groupby(['账户', '货币'])['金额'].sum().reset_index()
        account = account.merge(account_update, how = 'outer', on = ['账户', '货币'], suffixes=('', '_更新'), sort=True).fillna(0)
        account['金额'] += account['金额_更新']
        account = account[['账户', '货币', '类型', '金额']]
        self.book['账户'] = account
        print('*' * 50)
        
    def save(self, output_path):
        print('保存记账...')
        with pd.ExcelWriter(output_path) as writer:
            for name, df in self.book.items():
                df.to_excel(writer, sheet_name=name, index=False, freeze_panes = (1,0))
        print('*' * 50)

if __name__ == '__main__':
    today = datetime.date.today()
    yesterday = (today - datetime.timedelta(days = 1))
    start = datetime.date(2025,4,14)
    start = today
    end = today
    
    input_path = 'D:/Personal/bookkepping/bookkepping.xlsx'
    book_path = 'D:/Personal/bookkepping/bookkepping.xlsx'
    detail_path = f"D:/Wechat Files/WeChat Files/wxid_g9cw6yoy90kn22/FileStorage/File/{start.strftime('%Y-%m')}/薄荷记账_{str(start.year)}年{str(start.month)}月{str(start.day)}日.csv"
    book = bookkepping(input_path)
    book.update_detail('明细', '日期', detail_path, '日期', start, end)
    book.save(book_path)
    