import urllib
import os
import json
import datetime

import requests
from openpyxl import load_workbook

def get_file_name(url):
    path=urllib.parse.urlparse(url).path
    return os.path.split(path)[-1]

def download_osaka_model_data_file():
    #大阪府の大阪モデルモニタリング指標のページにある　Excelファイルのリンク
    url="http://www.pref.osaka.lg.jp/attach/23711/00362734/sihyou.xlsx"
    excel_file_name=get_file_name(url)

    response=requests.get(url)
    if response.status_code!=requests.codes.ok:
        raise Exception("status_code!=200")

    print(excel_file_name)
    with open(excel_file_name,"wb") as f:
        f.write(response.content)
    return excel_file_name


#num_data_type データの種類の数　
def process_one_table(ws,min_row,num_data_type=3):
    max_col=ws.max_column
    
    data=[]
    for col in ws.iter_cols(min_row=min_row,max_row=min_row+num_data_type,min_col=2,max_col=max_col):
        date_value=col[0].value
        if date_value is None:
            break
        date=date_value.date()
        # print(date)
        d=list(map(lambda item:item.value,col[1:]))

        #１つでもNoneがあれば、終了
        if any([item is None for item in d]):
            break


        d_dict={"item"+str(i):item for (i,item) in enumerate(d)}
        data.append({
            "date":date_value.date(),
            **d_dict
        })

    return data

def _get_data(ws):
    num_data_type=3
    data=[]
    for i in range(100): #実際には4つぐらいしか処理しない。
        #日付+データ数の数+空白　の合計(num_data_type+2)だけずらして処理していく
        min_row=32+i*(num_data_type+2)
        # print(min_row)
        date_cell=ws.cell(min_row,2)
        if date_cell.value is None:
            break

        d=process_one_table(ws,min_row,num_data_type=num_data_type)
        data.extend(d)
            
    return data
def get_bed_data_for_mild_or_moderate_patients(wb):
    ws=wb["(参考③)"]
    return _get_data(ws)

def get_bed_data_for_severe_patients(wb):
    ws=wb["(3)"]
    return _get_data(ws)

def get_accommodation_facility_data(wb):
    ws=wb["(参考④)"]
    return _get_data(ws)

def rename_keys(array,names):
    for d in array:
        d["date"]=d["date"].isoformat()

        for i,name in enumerate(names):
            d[names[i]]=d.pop("item"+str(i))

def main():
    file_name=download_osaka_model_data_file()


    wb=load_workbook(file_name,data_only=True)

    bed_data_for_mild_moderate=get_bed_data_for_mild_or_moderate_patients(wb)
    #キーの名前変更
    rename_keys(bed_data_for_mild_moderate,["hospital_capacity","num_patients","percentage"])

    bed_data_for_severe=get_bed_data_for_severe_patients(wb)
    rename_keys(bed_data_for_severe,["hospital_capacity","num_patients","percentage"])

    data_accommodation_facility=get_accommodation_facility_data(wb)
    rename_keys(data_accommodation_facility,["num_rooms","num_patients","percentage"])

    data={
        "data":{
            "bed_data_for_mild_moderate":bed_data_for_mild_moderate,
            "bed_data_for_severe":bed_data_for_severe,
            "data_accommodation_facility":data_accommodation_facility
        }
    }

    with open("hospital_beds_data.json","w") as f:
        json.dump(data,f,indent=4)

if __name__ == "__main__":
    main()