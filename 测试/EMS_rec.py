# -*- coding:utf-8 -*-
import requests
import urllib
from PIL import Image
import pytesseract


session = requests.Session()
image_url = 'http://www.11183.com.cn/ems/rand'
image_path = r'C:\MyProgram\file\image\rand.png'
url = 'http://www.11183.com.cn/ems/order/singleQuery_t'
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; rv:2.0.1) Gecko/20100101 Firefox/4.0.1'}
# get_website = session.get(url=image_url, headers=headers)
urllib.urlretrieve(image_url, image_path)
pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
im = Image.open(image_path)
text = pytesseract.image_to_string(im)

post_data = {'mailNum': '1062365610525', 'checkCode': text}
resp = session.post(url=url, headers=headers, data=post_data)
text = resp.text
print text
