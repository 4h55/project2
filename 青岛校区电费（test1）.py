import urllib.request, urllib.parse, urllib.error

from Tools.scripts.generate_opcode_h import header
from bs4 import BeautifulSoup
import re
import schedule
import time
import xlwt
import requests
import json



#创建excel表
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('山大building', cell_overwrite_ok=True)
worksheet1 = workbook.add_sheet('凤凰居6号楼', cell_overwrite_ok=True)
worksheet2 = workbook.add_sheet('B5号楼', cell_overwrite_ok=True)
worksheet3 = workbook.add_sheet('B2', cell_overwrite_ok=True)
worksheet4 = workbook.add_sheet('T1', cell_overwrite_ok=True)
worksheet5 = workbook.add_sheet('S1一多书院', cell_overwrite_ok=True)
worksheet6 = workbook.add_sheet('S11', cell_overwrite_ok=True)
worksheet7 = workbook.add_sheet('B9', cell_overwrite_ok=True)
worksheet8 = workbook.add_sheet('凤凰居9号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('凤凰居2号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('S5凤凰居5号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('凤凰居10号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('凤凰居2号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('凤凰居2号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('凤凰居2号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('凤凰居2号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('凤凰居2号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('凤凰居2号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('凤凰居2号楼', cell_overwrite_ok=True)
worksheet9 = workbook.add_sheet('凤凰居2号楼', cell_overwrite_ok=True)



col=('buildingig','building','宿舍号','剩余电费')
col1=('宿舍号','剩余电费')
for i in range(0,len(col)):
    worksheet.write(0,i,col[i])
for i in range(0,len(col1)):
    worksheet1.write(0,i,col1[i])
    worksheet2.write(0,i,col1[i])
    worksheet3.write(0,i,col1[i])
    worksheet4.write(0,i,col1[i])
    worksheet5.write(0,i,col1[i])
    worksheet6.write(0,i,col1[i])
    worksheet7.write(0,i,col1[i])
    worksheet8.write(0,i,col1[i])
    worksheet9.write(0,i,col1[i])

#得到buildingig和building
#以下数据均为抓包得到
    #响应头
headers = {

        "User-Agent": "Mozilla/5.0 (Linux; Android 9; RMX1931 Build/PQ3A.190605.09201023; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/91.0.4472.114 Safari/537.36",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",

     }
#请求体
json_data={"query_elec_building":{"retcode":"0", "errmsg":"请选择相应的楼栋：",  "aid":"0030000000002505", "account":"823872", "area":{"area":"青岛校区","areaname":"青岛校区"},  "buildingtab":[ { "buildingid":"1503975980", "building":"凤凰居6号楼" },{ "buildingid":"1661835273", "building":"B5号楼" },{ "buildingid":"1661835256", "building":"B2" },{ "buildingid":"1574231830", "building":"T1" },{ "buildingid":"1503975832", "building":"凤凰居1号楼" },{ "buildingid":"1503975832", "building":"S1一多书院" },{ "buildingid":"1599193777", "building":"S11" },{ "buildingid":"1693031698", "building":"B9" },{ "buildingid":"1503976004", "building":"凤凰居9号楼" },{ "buildingid":"1503975890", "building":"凤凰居2号楼" },{ "buildingid":"1503975967", "building":"S5凤凰居5号楼" },{ "buildingid":"1503976037", "building":"凤凰居10号楼" },{ "buildingid":"1503975890", "building":"S2从文书院" },{ "buildingid":"1693031710", "building":"阅海居B10楼" },{ "buildingid":"1693031698", "building":"阅海居B9楼" },{ "buildingid":"1574231835", "building":"T3" },{ "buildingid":"1503976004", "building":"S9凤凰居9号楼" },{ "buildingid":"1503975988", "building":"S7凤凰居7号楼" },{ "buildingid":"1503976037", "building":"S10凤凰居10号楼" },{ "buildingid":"1503975995", "building":"S8凤凰居8号楼" },{ "buildingid":"1599193777", "building":"凤凰居11/13号楼" },{ "buildingid":"1574231833", "building":"专家公寓2号楼" },{ "buildingid":"1503975902", "building":"凤凰居3号楼" },{ "buildingid":"1693031710", "building":"B10" },{ "buildingid":"1661835249", "building":"B1" },{ "buildingid":"1503975950", "building":"凤凰居4号楼" },{ "buildingid":"1503975980", "building":"S6凤凰居6号楼" }]}}
data=f"jsondata={urllib.parse.quote(json.dumps(json_data,ensure_ascii=False))}&funname=synjones.onecard.query.elec.building&json=true"


url="http://10.100.1.24:8988/web/Common/Tsm.html"
#得到数据
request =urllib.request.Request(url=url,headers=headers,data=data.encode("utf-8"),method='POST')
response = urllib.request.urlopen(request, timeout=6000)
html = response.read().decode("utf-8")
#利用正则提取得到关键信息
a=re.compile(r'"buildingid":"(.*?)",')
b=re.compile(r'"building":"(.*?)"')
buildingid=re.findall(a,html)
building=re.findall(b,html)

#得到buildingig
for i in range(0, len(buildingid)):
    worksheet.write(i + 1, 0, buildingid[i])
#得到building
for i in range(0, len(building)):
    worksheet.write(i + 1, 1, building[i])
#workbook.save('山大电费.xls')


#得到每个男生宿舍的电费
headers1 = {

        "User-Agent": "Mozilla/5.0 (Linux; Android 9; RMX1931 Build/PQ3A.190605.09201023; wv) AppleWebKit/537.36 (KHTML, like Gecko) Version/4.0 Chrome/91.0.4472.114 Safari/537.36",
        "Content-Type": "application/x-www-form-urlencoded; charset=UTF-8",

     }
#得到宿舍名称
c=101
b=100
list_name=[]
for i in range(0,149):
    c+=1
    a="B"+str(c)
    list_name.append(a)
    if c%100>24:
        b+=100
        c=b
for i in buildingid:
    for j in list_name:
        json_data={"query_elec_roominfo":{"retcode":"0", "errmsg":"房间当前剩余电量358.01",  "aid":"0030000000002505", "account":"823872",  "meterflag":"amt", "bal":"",  "price":"0", "pkgflag":"none", "area":{"area":"青岛校区","areaname":"青岛校区"},  "building":{"buildingid":i,"building":"a"},  "floor":{"floorid":"","floor":""},  "room":{"roomid":j,"room":j}, "pkgtab":[ ]}}
        data1 = f"jsondata={urllib.parse.quote(json.dumps(json_data, ensure_ascii=False))}&funname=synjones.onecard.query.elec.roominfo&json=true"
        request1 = urllib.request.Request(url=url, headers=headers, data=data1.encode("utf-8"), method='POST')
        response1 = urllib.request.urlopen(request1, timeout=6000)
        html1 = response1.read().decode("utf-8")
        a1=re.compile(r'"errmsg":"房间当前剩余电量(.*?)",')
        b1=re.findall(a1,html1)




