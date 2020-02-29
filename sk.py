#!/usr/bin/python3
from selenium import webdriver
from Crypto.PublicKey import RSA
from Crypto.Cipher import AES, PKCS1_OAEP
from selenium.webdriver.support.ui import Select
from dateutil.relativedelta import relativedelta 
import glob
import os
import sys
import time
import pandas
import datetime
if sys.platform=='linux':
    wkdir="/home/eric"
    pth=wkdir+'/下載'
else:
    wkdir=os.environ['HOMEDRIVE']+os.environ['HOMEPATH']
    pth=wkdir+'/Downloads'
def Descrypt(filename):
    code = 'nooneknows'
    with open(filename, 'rb') as fobj:
        # 导入私钥
        private_key = RSA.import_key(open(pth+'/my_private_rsa_key.bin').read(), passphrase=code)
        # 会话密钥， 随机数，消息认证码，机密的数据
        enc_session_key, nonce, tag, ciphertext = [ fobj.read(x) 
                                                    for x in (private_key.size_in_bytes(), 
                                                    16, 16, -1) ]
        cipher_rsa = PKCS1_OAEP.new(private_key)
        session_key = cipher_rsa.decrypt(enc_session_key)
        cipher_aes = AES.new(session_key, AES.MODE_EAX, nonce)
        # 解密
        data = cipher_aes.decrypt_and_verify(ciphertext, tag)
        os.remove(pth+"/my_private_rsa_key.bin")
    data=data.decode('ascii')
    return data

def getCsv(qryTy):
    driver=webdriver.Chrome(wkdir+'/Documents/python/chromedriver')
    #driver.implicitly_wait(30)
    driver.get('https://etrade.yuanta.com.tw/tsweb/')
    #設定身份證
    driver.switch_to.frame("down")
    ele=driver.find_element_by_xpath('//*[@id="Login_frm"]/table[1]/tbody/tr[2]/td[2]/table/tbody/tr[2]/td[2]/input[2]')
    ele.click()
    #輸入帳號密碼
    data=Descrypt(wkdir+"/mydata.txt").rstrip().split(',')
    ele=driver.find_element_by_id('loginid')
    ele.send_keys(data[0])
    ele=driver.find_element_by_id('loginPWD')
    ele.send_keys(data[1])
    ele=driver.find_element_by_id('button1')
    ele.click()
    time.sleep(5)
    #個人帳戶
    driver.switch_to.parent_frame()
    driver.switch_to.frame('tb')
    ele=driver.find_element_by_id('m1')
    ele.click()
    #投資明細
    ele=driver.find_element_by_id('i8')
    ele.click()
    time.sleep(5)
    #設定期間
    driver.switch_to.parent_frame()
    driver.switch_to.frame('down')
    s1 = Select(driver.find_element_by_id('selType')) 
    s1.select_by_value(qryTy)
    #查詢
    ele=driver.find_element_by_id('search')
    ele.click()
    time.sleep(5)
    #儲存結果
    #driver.switch_to.parent_frame()
    #driver.switch_to.frame('down')
    #fd=open("html.html","w")
    #fd.write(driver.page_source)
    #fd.close()
    #匯出excel
    for f in glob.glob(pth+"/投資明細*.xls"):
        os.remove(f)
    ele=driver.find_element_by_name("btnExportExcel")
    ele.click()
    time.sleep(5)
    #轉CSV檔
    for f in glob.glob(wkdir+"/Dropbox/a*.csv"):
        os.remove(f)
    dfs=pandas.read_html(pth+'/投資明細.xls')
    df=dfs[0]
    df.columns=df.loc[0]
    df.drop([0],inplace=True)
    today = datetime.date.today()
    yr=str(today.year)
    #如果是上月，重新計算年度
    if qryTy=='4':
        yr=str((today - relativedelta(months=1)).year)
    df.to_csv(wkdir+'/Dropbox/a'+yr+'.csv', index=0,encoding='utf-8')
    #關閉瀏覽器
    driver.close()
    fd=open(wkdir+"/Dropbox/fn.txt","w")
    fd.write("/a"+yr+".csv,"+qryTy)
    fd.close()
    return wkdir+'/Dropbox/a'+yr+'.csv轉檔成功'

#MENU選擇
if os.path.exists(pth+'/my_private_rsa_key.bin'):
    sel=(input("\n(T)本年度 (3)本月 (4)上月：")).upper()
    if sel in('T','3','4'):
        print(getCsv(sel))
    else:
        print("錯誤: 只能輸入 3, 4, T!")
else:
    print(pth+'/my_private_rsa_key.bin'+"不存在！！！請下載此檔案，再執行本轉檔作業。")
sel=input("按Enter結束作業...")
exit()
