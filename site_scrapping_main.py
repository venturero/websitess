from selenium import webdriver
from PIL import Image
import time
import pandas as pd
import xlsxwriter
#from Screenshot import Screenshot_clipping
import os, os.path

file_count = sum(len(files) for _, _, files in os.walk(r'First create images files then write images file path here'))

df = pd.read_excel(r"excel path")#to get len of a column
driver = webdriver.Chrome(executable_path = "chromedriver(or other driver) path, to install driver please read read me")
#url = 'https://www.google.com/'
workbook = xlsxwriter.Workbook("excel file path to write")
unlocked_format = workbook.add_format({'locked': False})#to overwrite if needed
worksheet = workbook.add_worksheet()

images_script=[]

for i in range(len(df)):
    url = df["Company url"][i]
    try:
        driver.get(url)
        time.sleep(0.1)
        driver.save_screenshot("images\\{}_ss.png".format(i))#correct the file path according to your file path
        a="{}_ss.png".format(i)
        images_script.append(a)
        #worksheet.insert_image('E{}'.format(i + 2), 'images/{}_ss.png'.format(i), {'x_scale': 0.045, 'y_scale': 0.045})
        worksheet.write('E{}'.format(i + 2),"{}_ss.png".format(i))
    except:
        pass

driver.quit()
workbook.close()
"""
resource:
#https://github.com/batrlatom/aliexpress_feedback_images_scrapper/blob/master/FinalScrapper.ipynb
"""