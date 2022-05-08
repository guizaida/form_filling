from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import time
import sys
import configparser as cp
filename = './config.ini'
inifile = cp.ConfigParser()
inifile.read(filename,'UTF-8')
acc = inifile.get('Microsoft','acc')
password = inifile.get('Microsoft','pwd')
url = inifile.get('Microsoft','url')
options = Options()#瀏覽器設定
options.add_argument('--no-sandbox')
options.add_argument("--incognito")#無痕模式
options.add_experimental_option("detach", True)
driver = webdriver.Chrome(ChromeDriverManager().install(),chrome_options=options)
driver.get (url)
locator=(By.XPATH,'//*[@id="i0116"]')
WebDriverWait(driver,30,1).until(EC.presence_of_element_located(locator))#等待頁面載入完成
driver.find_element(By.XPATH,'//*[@id="i0116"]').send_keys(acc)
time.sleep(3)
driver.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
time.sleep(3)
driver.find_element(By.XPATH,'//*[@id="i0118"]').send_keys(password)
time.sleep(3)
driver.find_element(By.XPATH,'//*[@id="idSIButton9"]').click()
time.sleep(3)
driver.find_element(By.XPATH,'//*[@id="idBtn_Back"]').click()
time.sleep(3)
#今日身體狀況 ?(How is your health status today?)
driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[2]/div/div[2]/div/div[1]/div/label/input').click()#正常(Normal)
# driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[2]/div/div[2]/div/div[2]/div/label/input').click()#其他
# health_status ='請輸入身體狀況'
# driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[2]/div/div[2]/div/div[2]/div/label/div/div/input').sendkey(health_status)#其他
time.sleep(3)
#本次填寫時段
if sys.argv[1] == 'morning':#早上上班前填寫
    driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[3]/div/div[2]/div/div[1]/div/label/input').click()
    time.sleep(3)
    #今日的工作地點 
    if sys.argv[2] == 'office':#ECV辦公室(ECV Office)
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[4]/div/div[2]/div/div[1]/div/label/input').click()
        status = True
    elif sys.argv[2] == 'wfh':#在家工作(WFH)
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[4]/div/div[2]/div/div[2]/div/label/input').click()
        status = True
    elif sys.argv[2] == 'holiday':#排定休假，含事、病、防疫隔離假等
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[4]/div/div[2]/div/div[4]/div/label/input').click()
        status = True
    time.sleep(3)
    if status:
        #提醒一下~ 需要填Time sheet 的同仁，前一次的是否已經填寫了呢?
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[1]/div/label/input').click()#是
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[2]/div/label/input').click()#否
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[3]/div/label/input').click()#免填
        time.sleep(3)
        #在近兩周內是否有與發燒或是身體不適患者接觸?
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[6]/div/div[2]/div/div[1]/div/label/input').click()#是
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[6]/div/div[2]/div/div[2]/div/label/input').click()#否
        time.sleep(3)
        #請確認，台灣社交距離APP上檢視是否有與確診者的接觸距離?
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[7]/div/div[2]/div/div[1]/div/label/input').click()#是
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[7]/div/div[2]/div/div[2]/div/label/input').click()#否
        time.sleep(3)
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[4]/div[1]/button/div').click()#提交
        time.sleep(3)


#################################################################################################################################################################################
# driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[4]/div/div[2]/div/div[3]/div/label/input').click()#出差(Business Trip)
#################################################################################################################################################################################

elif sys.argv[1] == 'night':#晚上下班前
    driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[3]/div/div[2]/div/div[2]/div/label/input').click()#晚上下班前填寫
    time.sleep(3)
    status = False
    #今晚預計移動地點?
    driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[4]/div/div[2]/div/div[1]/div/label/input').click()#在家休息無外出
    # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[4]/div/div[2]/div/div[2]/div/label/input').click()#直接回家不繞路 (Go home directly)
    #如果選擇需要前往人潮聚集地 或者 前往醫療院所則需要打開此選項
    # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[4]/div/div[2]/div/div[3]/div/label/input').click()#需要前往人潮聚集地(如風景區、遊樂園、夜市、遶境...等)
    # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[4]/div/div[2]/div/div[4]/div/label/input').click()#前往醫療院所(如診所、醫院、衛生中心...等)
    #status = True
    if status != True:
        time.sleep(3)
        #提醒一下~ 需要填Time sheet 的同仁，前一次的是否已經填寫了呢?
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[1]/div/label/input').click()#是
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[2]/div/label/input').click()#否
        time.sleep(3)
        #在近兩周內是否有與確診者或是被匡列人士接觸?
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[6]/div/div[2]/div/div[1]/div/label/input').click()#是
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[6]/div/div[2]/div/div[2]/div/label/input').click()#否
        #在近兩周內是否有與發燒或是身體不適患者接觸?
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[7]/div/div[2]/div/div[1]/div/label/input').click()#是
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[7]/div/div[2]/div/div[2]/div/label/input').click()#否
        time.sleep(3)
        #請確認，台灣社交距離APP上檢視是否有與確診者的接觸距離?
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[8]/div/div[2]/div/div[1]/div/label/input').click()#是
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[8]/div/div[2]/div/div[2]/div/label/input').click()#否
    else:
        #我預計前往的地點
        place = '請填寫要前往之地點'
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div/input').send_keys(place)
        time.sleep(3)
        #提醒一下~ 需要填Time sheet 的同仁，今日是否填寫了呢?
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[6]/div/div[2]/div/div[1]/div/label/input').click()#是
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[6]/div/div[2]/div/div[2]/div/label/input').click()#否
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[6]/div/div[2]/div/div[3]/div/label/input').click()#免填
        time.sleep(3)
        #在近兩周內是否有與確診者或是被匡列人士接觸?
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[7]/div/div[2]/div/div[1]/div/label/input').click()#是
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[7]/div/div[2]/div/div[2]/div/label/input').click()#否
        time.sleep(3)
        #在近兩周內是否有與發燒或是身體不適患者接觸?
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[8]/div/div[2]/div/div[1]/div/label/input').click()#是
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[8]/div/div[2]/div/div[2]/div/label/input').click()#否
        time.sleep(3)
        #請確認，台灣社交距離APP上檢視是否有與確診者的接觸距離?
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[9]/div/div[2]/div/div[1]/div/label/input').click()#是
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[9]/div/div[2]/div/div[2]/div/label/input').click()#否
    time.sleep(3)
    driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[4]/div[1]/button/div').click()#提交
    time.sleep(3)

#假日後上班
elif sys.argv[1] == 'monday':#假日後上班(週一早上請選此項)
    driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[3]/div/div[2]/div/div[3]/div/label/input').click()
    time.sleep(3)
    #假日期間活動地點?
    driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[4]/div/div[2]/div/div[1]/div/label/input').click()#在家休息無外出
    time.sleep(3)
    #今日的工作地點?
    if sys.argv[2] == 'office':#ECV辦公室
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[1]/div/label/input').click()
        status = True
        time.sleep(3)
    elif sys.argv[2] == 'wfh':#WFH
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[2]/div/label/input').click()#今日的工作地點?在家工作(WFH)
        time.sleep(3)
        status = True
    elif sys.argv[2] == 'holiday':#排定休假，含事、病、防疫隔離假等
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[5]/div/div[2]/div/div[4]/div/label/input').click()
        status = True
    if status:
        #提醒一下~ 需要填Time sheet 的同仁，前一次的是否已經填寫了呢?
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[6]/div/div[2]/div/div[1]/div/label/input').click()#是
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[6]/div/div[2]/div/div[2]/div/label/input').click()#否
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[6]/div/div[2]/div/div[3]/div/label/input').click()#免填
        time.sleep(3)
        #在近兩周內是否有與發燒或是身體不適患者接觸?
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[7]/div/div[2]/div/div[1]/div/label/input').click()#是
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[7]/div/div[2]/div/div[2]/div/label/input').click()#否
        time.sleep(3)
        #請確認，台灣社交距離APP上檢視是否有與確診者的接觸距離?
        # driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[8]/div/div[2]/div/div[1]/div/label/input').click()#是
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[2]/div[8]/div/div[2]/div/div[2]/div/label/input').click()#否
        time.sleep(3)
        driver.find_element(By.XPATH,'//*[@id="form-container"]/div/div/div[1]/div/div[1]/div[3]/div[4]/div[1]/button/div').click()#提交
        time.sleep(3)
driver.quit()