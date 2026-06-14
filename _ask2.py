import json, io
import server as s
def count(obj):
    r = json.loads(s.execute_query(f"ВЫБРАТЬ КОЛИЧЕСТВО(*) КАК К ИЗ {obj}", None, 1))
    return r["rows"][0][0]
res = {o: count("Справочник."+o) for o in ["Номенклатура","Контрагенты","ФизическиеЛица","Валюты","Склады"]}
res["Документ.РеализацияТоваровУслуг"] = count("Документ.РеализацияТоваровУслуг")
s.conn.close()
io.open('_ask2.json','w',encoding='utf-8').write(json.dumps(res, ensure_ascii=False, indent=2))
print('ok')
