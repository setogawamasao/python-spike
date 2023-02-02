import pandas as pd

# ヘッダとインデックス番号が、太字・枠線解除
df = pd.read_excel('data.xlsx', sheet_name='202004')
df.to_excel('pd_data.xlsx', sheet_name='new_sheet', header=False, index=False)