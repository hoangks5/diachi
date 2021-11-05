import openpyxl
import  json
book = openpyxl.load_workbook('data.xlsx')

sheet = book.active
def update_json(s):  # Lưu dữ liệu xử lỹ vào một tệp tin để training. Sử dụng tệp .json thay vì .txt để dễ truy vấn
    with open("data.json",'r', encoding='utf-8') as fp:
        information1 = json.load(fp)
    information1["intents"].append({
        "title": s,
        "text": s
    })
    with open("data.json",'w',encoding='utf-8') as fp: # Thêm dữ liệu vào tệp JSON
        json.dump(information1, fp, indent=2,ensure_ascii=False)
    
for i in range(1,10606,1):
    diachi = sheet['E'+str(i)].value+' '+sheet['C'+str(i)].value+' '+sheet['A'+str(i)].value
    diachi = diachi.lower()
    update_json(diachi)