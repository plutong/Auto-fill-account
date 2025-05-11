#%% 包区域
import pandas as pd
import datetime
#%% 定义记账表读取与更新包
class bookkepping:
    def __init__(self, path, col = None):
        """
        加载原始记账数据
        arg:
            path: str -> 记账表地址
            col: str -> 读取的记账工作表名称
        output:
            self.book: pd.DataFrame -> 记账表
        """
        print(f'{'*'*50}\n读取原始记账表...')
        # 读取记账表
        try:
            self.book = pd.read_excel(path, sheet_name=col)
            for sheet_name, _ in self.book.items():
                print(f'读取了{sheet_name}')
        except:
            self.book = dict()
            print(f'没有找到文件{path}，会新建一个')
        print('*' * 50)
    
    def update_extract(self, detail_path):
        """
        读取更新数据并做初步处理
        arg:
            detail_path: str -> 更新记账表地址
        output:
            self.detail_io: pd.DataFrame -> 更新的收入/支出信息
            self.date_transfer: pd.DataFrame -> 更新的转账信息
        """
        print(f'{'*'*50}\n读取更新记账表...')
        # 读取数据并对列做数据转换
        detail = pd.read_csv(detail_path, encoding = 'utf-16LE', on_bad_lines='skip', sep = '\t')
        detail['金额'] = pd.to_numeric(detail['金额'].apply(lambda x: x.replace(',', '') if isinstance(x, str) else x))
        detail.loc[:, '金额'] = pd.to_numeric(detail.loc[:, '金额'])
        if not set(detail['类型'].dropna().unique()).issubset({'收入', '支出', '转账'}): 
            raise Exception('请检查明细表类型')
        detail.loc[:,'日期'] = pd.to_datetime(detail['日期']).dt.date
        # 根据类型将新增数据划分为 增减数据 与 转账数据
        detail_io, detail_transfer = detail[detail['类型'].isin(['收入', '支出'])], detail[detail['类型'].isin(['转账'])]
        detail_io = detail_io[['日期', '类型', '货币', '金额', '分类','账户', '备注']]
        # 将两类数据加载到类中
        self.detail_io = detail_io
        self.detail_transfer = detail_transfer

    def update_detail(self, name):
        """
        将支出/收入数据加载到明细记账表中
        arg:
            name: str -> 原始明细记账表的工作表名称
        update:
            self.book[name]: pd.DataFrame -> 更新后明细记账表
        """
        print(f'{'*'*50}\n更新明细记账表...')
        # 加载原始明细表
        book_detail = self.book[name]
        book_detail['日期'] = pd.to_datetime(book_detail['日期']).dt.date
        # 更新明细表
        new_detail = self.detail_io[~self.detail_io.apply(tuple, axis=1).isin(book_detail.apply(tuple, axis=1))]
        print(f'新添明细数据\n{new_detail}')
        book_detail = pd.concat([book_detail, new_detail], ignore_index=True)
        book_detail.sort_values('日期', ascending=False, inplace=True)
        self.book[name] = book_detail
        # 加载更新的明细表，后期更新账户数据表要用
        self.new_detail = new_detail

    def update_account(self, name):
        """
        根据更新账户记账表中
        arg:
            name: str -> 原始账户记账表的工作表名称
        update:
            self.book[name]: pd.DataFrame -> 更新账户记账表
        """
        print(f'{'*'*50}\n更新账户记账表...')
        # 加载原始账户表
        new_detail = self.new_detail
        account = self.book[name]
        # 将转账数据分解成流入数据
        detail_in = self.detail_transfer.copy()
        detail_in.loc[:, '账户'] = detail_in.loc[:, '付款']
        detail_in.loc[:, '金额'] = -detail_in.loc[:, '金额']
        detail_in = detail_in[['日期', '货币', '金额', '账户']]
        # 将转账数据分解成流出数据
        detail_out = self.detail_transfer.copy()
        detail_out.loc[:, '账户'] = detail_out.loc[:, '收款']
        detail_out = detail_out[['日期', '货币', '金额', '账户']]
        # 合装成账户的输入输出信息并更新
        new_detail = pd.concat([new_detail[['日期', '货币', '金额', '账户']], detail_in, detail_out])
        account_update = new_detail.groupby(['账户', '货币'])['金额'].sum().reset_index()
        print(f'时段内账户变动为\n{account_update}')
        account = account.merge(account_update, how = 'outer', on = ['账户', '货币'], suffixes=('', '_更新'), sort=True).fillna(0)
        account['金额'] += account['金额_更新']
        account = account[['账户', '货币', '类型', '金额']]
        self.book[name] = account
        
    def save(self, output_path):
        """
        将更新后数据加载到目标地址
        arg:
            output_path: str -> 保存更新后记账表的地址
        """
        print(f'{'*'*50}\n保存记账...')
        with pd.ExcelWriter(output_path) as writer:
            for name, df in self.book.items():
                df.to_excel(writer, sheet_name=name, index=False, freeze_panes = (1,0))
        print('*' * 50)

#%% 
def main(book_path, output_path, detail_path=None, date=None):
    """
    arg:
        date: datetime.datetime.date -> 更新的文件日期
        book_path: str -> 原始记账表地址
        detail_path: str -> 更新记账表地址
        output_path: str -> 输出记账表地址
    """
    # 初始化原始记账表
    book = bookkepping(book_path)
    # 更新明细与账户表
    if not detail_path:
        # 
        detail_path = f"D:/Wechat Files/WeChat Files/wxid_g9cw6yoy90kn22/FileStorage/File/{date.strftime('%Y-%m')}/薄荷记账_{str(date.year)}年{str(date.month)}月{str(date.day)}日.csv"
    book.update_extract(detail_path)
    book.update_detail(name='明细')
    book.update_account(name='账户')
    # 将更新后数据存入输出地址
    book.save(output_path=output_path)
        

if __name__ == '__main__':
    today = datetime.date.today()
    book_path = 'D:/Personal/bookkepping/bookkepping.xlsx'
    main(book_path, book_path, date=today)
    