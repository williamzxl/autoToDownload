#!/usr/bin/env python3
# -*- coding:utf-8 -*-
from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from getUserInfo import XlUserInfo
import threading, time

class AutoDownload(object):
    def __init__(self,args, args2):
        self.args = args
        self.args2 = args2

    def openBrower(self):
        self.profile = webdriver.FirefoxProfile()
        self.profile.accept_untrusted_certs = True
        self.profile = webdriver.FirefoxProfile(r"ini3hkyu.myProfile")
        self.profile.set_preference("browser.download.useDownloadDir", True);
        if self.args2['downloadpath'] is None:
            self.profile.set_preference('browser.download.dir', 'c:\\')
        else:
            self.profile.set_preference('browser.download.dir', self.args2['downloadpath'])
            print(self.args2['downloadpath'])
        # self.profile.set_preference('browser.download.folderList', 2)
        # # self.profile.set_preference('browser.download.manager.showWhenStarting', True)
        # self.profile.set_preference('browser.download.manager.showWhenStarting', False)
        # self.profile.set_preference('browser.helperApps.neverAsk.saveToDisk',
        #                             'application/zip,application/xml,text/plain,text/xml')

        # application/octet-stream,text/html,text/webviewhtml,x-world/x-vrml,application/internet-property-stream,application/postscript,
        # x - world / x - vrml, text / x - vcard, text / x - setext, text / x - component, text / webviewhtml,application/x-compress

        self.driver = webdriver.Firefox(firefox_profile=self.profile)
        self.driver.implicitly_wait(30)
        return self.driver

    def openUrl(self):
        try:
            self.driver.get(self.args2['url'])
            self.driver.maximize_window()
        except:
            print("Failed to get {}".format(self.args2['url']))
        return self.driver

    def login(self):
        '''
        user_name
        pwd_name
        logIn_name
        '''
        self.driver.find_element_by_name(self.args['user_name']).send_keys(self.args2['uname'])
        if isinstance(self.args2['pwd'],float):
            self.driver.find_element_by_name(self.args['pwd_name']).send_keys(int(self.args2['pwd']))
        else:
            self.driver.find_element_by_name(self.args['pwd_name']).send_keys(self.args2['pwd'])
        self.driver.find_element_by_name(self.args['logIn_name']).click()
        self.driver.implicitly_wait(10)
        print(self.driver.window_handles)
        return self.driver

    def download(self):
        self.driver.implicitly_wait(15)
        self.driver.find_element_by_link_text(self.args['Search_Forms_text']).click()
        self.driver.implicitly_wait(30)
        self.driver.find_element_by_id(self.args['OSS_Num_type_id']).send_keys(int(self.args2['OSS_num']))
        self.driver.find_element_by_id(self.args['Search_button_id']).click()
        self.driver.implicitly_wait(10)
        self.driver.find_element_by_link_text(str(int(self.args2['OSS_num']))).click()
        self.driver.implicitly_wait(20)
        # Attachments_text
        self.driver.find_element_by_link_text(self.args['Attachments_text']).click()
        self.driver.implicitly_wait(10)
        try:
            for i in range(2,9):
                # xpath = r'//table[4]//tr[{}]/td[1]/a'.format(i)
                print('//table[4]//tr[{}]/td[1]/a'.format(i))
                time.sleep(5)
                self.driver.find_element_by_xpath('//table[4]//tr[{}]/td[1]/a'.format(i)).click()
                time.sleep(10)
        except:
            self.driver.find_element_by_xpath('//table[4]//tr[3]/td[1]/a').click()
            time.sleep(10)
            self.driver.find_element_by_xpath('//table[4]//tr[5]/td[1]/a').click()
            print("Finish to download")
        finally:
            time.sleep(300)
            print("Close the browser")
            self.driver.quit()


        # self.driver.find_element_by_xpath('//table[4]//tr[3]/td[1]/a').click()
        # self.driver.find_element_by_xpath('//table[4]//tr[4]/td[1]/a').click()
        # self.driver.find_element_by_xpath('//table[4]//tr[6]/td[1]/a').click()
        # self.driver.find_element_by_xpath('//table[4]//tr[5]/td[1]/a').click()

    def send_keys(self):
        self.driver.send_keys(Keys.ARROW_DOWN)
        print('william')
        self.driver.send_keys(Keys.ENTER)



    def quit(self):
        self.driver.quit()

    def Run(self):
        self.openBrower()
        self.openUrl()
        self.login()
        self.download()
        # self.quit()


if __name__ == '__main__':
    xl = XlUserInfo('userInfo.xlsx')
    userinfo = xl.get_sheetinfo_by_name('userInfo')
    webinfo = xl.get_sheetinfo_by_name('WebEle')
    print(userinfo)
    print(webinfo)
    # down_txt = AutoDownload('txt',webinfo,userinfo)
    down_load = AutoDownload(webinfo,userinfo)
    down_load.Run()


