# -*- coding:utf-8 -*-
from PIL import Image
import pytesseract
# from selenium import webdriver

# pytesseract.pytesseract.tesseract_cmd = 'C:\\Program Files (x86)\\Tesseract-OCR\\tesseract.exe'
# im = Image.open('../../file/2.png')
# text = pytesseract.image_to_string(im)
# print text

from selenium import webdriver
import time
from selenium.webdriver.chrome.options import Options

def take_screenshot(url, save_fn="../../file/capture.png"):
    chrome_option = Options()
    chrome_option.add_argument('--headless')
    chrome_option.add_argument('--disable-gpu')
    browser = webdriver.Chrome(chrome_options=chrome_option)
    browser.set_window_size(1200, 900)
    browser.get(url)  # Load page
    browser.execute_script("""
        (function () {
            var y = 0;
            var step = 100;
            window.scroll(0, 0);

            function f() {
                if (y < document.body.scrollHeight) {
                    y += step;
                    window.scroll(0, y);
                    setTimeout(f, 100);
                } else {
                    window.scroll(0, 0);
                    document.title += "scroll-done";
                }
            }

            setTimeout(f, 1000);
        })();
    """)

    for i in xrange(30):
        if "scroll-done" in browser.title:
            break
        time.sleep(10)

    browser.save_screenshot(save_fn)
    browser.close()

if __name__ == "__main__":

    take_screenshot("http://codingpy.com")