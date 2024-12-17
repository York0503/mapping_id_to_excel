# 依照excel指定欄位進行對比並更新資料

# config.ini 參數介紹
- 1.file_path: input excel file path
- 2.sheet_name: excel分頁名稱
- 3.mapping_column_name: 要比對是否相符的欄位(付以","隔開，ex: id,name)
- 4.update_column_name: 相符資料所需更新的欄位，請輸入該列表title

參考範例
```
[Files_1]
file_path = input_Excel/Product_CHT_ENG_Reference_new.xlsx
sheet_name = Snacks
mapping_column_name = item_no,item_name (中文)
update_column_name = item_name (EN)
```
