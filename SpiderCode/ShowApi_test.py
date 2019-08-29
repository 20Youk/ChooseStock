# -*- coding:utf-8 -*-
# Author:Youk.Lin
# 应用:
from ShowapiRequest import ShowapiRequest

r = Request("http://route.showapi.com/1653-1","my_appId","my_appSecret" )
r.addBodyPara("keyWords", "")
r.addBodyPara("page", "0")
r.addBodyPara("cityName", "成都")
r.addBodyPara("inDate", "")
r.addBodyPara("outDate", "")
r.addBodyPara("sortCode", "")
r.addBodyPara("returnFilter", "")
r.addBodyPara("star", "")
r.addBodyPara("feature", "")
r.addBodyPara("minPrice", "")
r.addBodyPara("maxPrice", "")
r.addBodyPara("facility", "")
r.addBodyPara("hotellabels", "")
# r.addFilePara("img", r"C:\Users\showa\Desktop\使用过的\4.png") #文件上传时设置
res = r.post()
print(res.text) # 返回信息