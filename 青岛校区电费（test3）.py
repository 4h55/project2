import urllib.request, urllib.parse, urllib.error
from win11toast import toast
import re
import schedule
import time
import xlwt
import requests
import json
import sys



#创建excel表
workbook = xlwt.Workbook(encoding='utf-8')
worksheet = workbook.add_sheet('山大building', cell_overwrite_ok=True)
col=('building','buildingid')
for i in range(0,len(col)):
    worksheet.write(0,i,col[i])
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
    worksheet.write(i + 1, 1, buildingid[i])
#得到building
for i in range(0, len(building)):
    worksheet.write(i + 1, 0, building[i])
workbook.save('山大电费.xls')


#查询宿舍的电费

#创建buildingid与building的字典
building_buildingid={}
for i in range(0, len(building)):
    building_buildingid[building[i]]=buildingid[i]
#得到用户需要查询的楼名
building_num=input('''
'1.凤凰居6号楼',
'2.B5号楼',
'3.B2',
'4.T1',
'5.凤凰居1号楼',
'6.S1一多书院',
'7.S11', 
'8.B9', 
'9.凤凰居9号楼',
'10.凤凰居2号楼',
'11.S5凤凰居5号楼',
'12.凤凰居10号楼',
'13.S2从文书院', 
'14.阅海居B10楼', 
'15.阅海居B9楼', 
'16.T3', 
'17.S9凤凰居9号楼', 
'18.S7凤凰居7号楼', 
'19.S10凤凰居10号楼', 
'20.S8凤凰居8号楼', 
'21.凤凰居11/13号楼',
'22.专家公寓2号楼', 
'23.凤凰居3号楼', 
'24.B10', 
'25.B1', 
'26.凤凰居4号楼', 
'27.S6凤凰居6号楼',
请选择你所居住的大楼的编号并输入
''')
#得到用户需要查询的buildingid
buildingid_name=building_buildingid.get(building[int(building_num)-1])
#得到房间编号
room_name=input("请输入你的房间编号")

#请求体的构建：将得到的buildingig放入请求体发送到服务器
json_data={"query_elec_roominfo":{"retcode":"0", "errmsg":"房间当前剩余电量358.01",  "aid":"0030000000002505", "account":"823872",  "meterflag":"amt", "bal":"",  "price":"0", "pkgflag":"none", "area":{"area":"青岛校区","areaname":"青岛校区"},  "building":{"buildingid":buildingid_name,"building":""},  "floor":{"floorid":"","floor":""},  "room":{"roomid":room_name,"room":room_name}, "pkgtab":[ ]}}
data1 = f"jsondata={urllib.parse.quote(json.dumps(json_data, ensure_ascii=False))}&funname=synjones.onecard.query.elec.roominfo&json=true"
request1 = urllib.request.Request(url=url, headers=headers, data=data1.encode("utf-8"), method='POST')
#得到返回的数据
response1 = urllib.request.urlopen(request1, timeout=6000)
html1 = response1.read().decode("utf-8")
#利用正则将剩余电费内容提取出来
a1=re.compile(r'"errmsg":"房间当前剩余电量(.*?)",')
elc=re.findall(a1,html1)


#在windows桌面上定时发送提醒
#用户自己设置电量还剩多少时自动提醒
a=int(input("请输入还剩多少度电时需要提醒"))
def job():

    if float(elc[0])<a:
    # 使用示例：
        toast('警告！电费不足，请尽快充值',f'电量仅剩{float(elc[0])}度',
              image=r'C:\Users\4h55\Downloads\029b01fc-ba29-4427-a906-4b9b07957ce6.jpg',
              duration='long',
              )
    else:
        toast('今晚不用担心断电问题啦', '请安心学习',
              image=r'"C:\Users\4h55\Downloads\59fc2ac28b2d7_610.jpg"',
              duration='long',
              )



schedule.every(int(input("请输入分钟数"))).minutes.do(job)
while True:
     schedule.run_pending()
     time.sleep(1)

    









