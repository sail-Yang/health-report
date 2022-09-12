# This is a sample Python script.

# Press Shift+F10 to execute it or replace it with your code.
# Press Double Shift to search everywhere for classes, files, tool windows, actions, and settings.

import sys

from selenium.webdriver import Keys

import mainWindow
from PyQt5.QtWidgets import QApplication, QDialog, QMessageBox
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.common.exceptions import TimeoutException, NoSuchElementException
from selenium.webdriver.chrome.service import Service
from PyQt5.QtCore import QCoreApplication,Qt
import time
import json
from qt_material import apply_stylesheet


class MainDialog(QDialog):
    def __init__(self, parent=None):
        super(QDialog, self).__init__(parent)
        self.ui = mainWindow.Ui_Dialog()
        self.ui.setupUi(self)
        self.ui.startButton.clicked.connect(self.Login)
        self.ui.exitButton.clicked.connect(QCoreApplication.quit)
        self.initData()

    student = {
        'usrname': '',
        'passwd': '',
        'reason': "",
        'class': "计科2003",
        'province': "",
        'city': "",
        'qu': "",
        'living': ""
    }
    def initData(self):
        with open('infomation.json','r',encoding='utf-8') as jsonFile:
            data = json.load(jsonFile)
            self.ui.usrname.setText(data["usrname"])
            self.ui.passwd.setText(data["passwd"])
            self.ui.reasonText.setText(data["reason"])
            self.ui.Class.setText(data["class"])
            self.ui.province.setText(data["province"])
            self.ui.city.setText(data["city"])
            self.ui.qu.setText(data["qu"])
            self.ui.living.setText(data["living"])


    def saveData(self):
        with open("infomation.json",'w',encoding='utf-8') as jsonFile:
            json_str = json.dumps(self.student, indent=4)
            jsonFile.write(json_str)

    def Login(self):

        url = 'https://sso.yzu.edu.cn/login?service=https:%2F%2Fehall.yzu.edu.cn%2Finfoplus%2Flogin%3FretUrl%3Dhttps%253A%252F%252Fehall.yzu.edu.cn%252Finfoplus%252Fform%252FXNYQSB%252Fstart'
        s = Service('msedgedriver.exe')
        browser = webdriver.Edge(service=s)
        self.student['usrname'] = self.ui.usrname.text()
        self.student['passwd'] = self.ui.passwd.text()
        self.student['reason'] = self.ui.reasonText.toPlainText()
        self.student['class'] = self.ui.Class.text()
        self.student['province'] = self.ui.province.text()
        self.student['city'] = self.ui.city.text()
        self.student['qu'] = self.ui.qu.text()
        self.student['living'] = self.ui.living.text()

        try:
            browser.get(url)
            time.sleep(1)

            inputID = browser.find_element(By.XPATH, '//input[@name="username"]')
            inputID.send_keys(self.student['usrname'])

            inputPassword = browser.find_element(By.XPATH, '//input[@type="password"]')
            inputPassword.send_keys(self.student['passwd'])

            enterButton = browser.find_element(By.XPATH, '//button[@class="login-button ant-btn"]')
            enterButton.click()
            time.sleep(1)

            startButton = browser.find_element(By.XPATH, '//a[@id="preview_start_button"]')
            startButton.click()
            time.sleep(1)

            hasReport = True

            try:
                hasReportedButton = browser.find_element(By.XPATH, '//button[@class="dialog_button default fr"]')
                hasReportedButton.click()
                self.noPassMessageDialog()
            except NoSuchElementException:
                hasReport = False

            if(not hasReport):
                self.report(browser)

        except NoSuchElementException:
            print('No Element found')
            self.passwdWrongMessageDialog()
        except Exception as e:
            print(e)
        finally:
            self.saveData()
            browser.close()

    def report(self,browser):
        try:
            inputClass = browser.find_element(By.ID,'V1_CTRL8')
            inputClass.send_keys(self.student["class"])
            if (self.ui.IsSchool.isChecked()):
                isInSchoolButton = browser.find_element(By.XPATH, '//input[@id="fieldSFZX-0"]')
                isInSchoolButton.click()
            else:
                noIsInSchoolButton = browser.find_element(By.XPATH, '//input[@id="fieldSFZX-1"]')
                noIsInSchoolButton.click()
                time.sleep(1)
                inputProvince = browser.find_element(By.XPATH,'//div[@class="suggest_frame_outer"]//div[@class=V1_CTRL117345_activeDiv]//input')
                inputProvince.send_keys(self.student['province'])
                time.sleep(2)
                print("stop")
                inputProvince.send_keys(Keys.ENTER)
                inputCity = browser.find_element(By.XPATH,'//input[@placeholder="市.city" | @class="infoplus_control active_input validate[funcCall[checkRenderFormFields]]"]')
                inputCity.send_keys(self.student['city'])
                time.sleep(2)
                inputCity.send_keys(Keys.ENTER)
                inputQu = browser.find_element(By.XPATH, '//input[@placeholder="区.district" | @class="infoplus_control active_input validate[funcCall[checkRenderFormFields]]"]')
                inputQu.send_keys(self.student['qu'])
                inputQu.send_keys(Keys.ENTER)
                inputLiving= browser.find_element(By.ID,'V1_CTRL120')
                inputLiving.send_keys(self.student['living'])


            inputTravel = browser.find_element(By.XPATH, '//textarea[@class="xdRichTextBox infoplus_control '
                                                         'infoplus_textareaControl validate[funcCall['
                                                         'checkRenderFormFields]] infoplus_writable"]')
            inputTravel.send_keys(self.student["reason"])

            CheckBox1 = browser.find_element(By.ID,'V1_CTRL128')
            CheckBox1.click()
            CheckBox2 = browser.find_element(By.ID, 'V1_CTRL129')
            CheckBox2.click()
            CheckBox3 = browser.find_element(By.ID, 'V1_CTRL130')
            CheckBox3.click()
            CheckBox4 = browser.find_element(By.ID, 'V1_CTRL138')
            CheckBox4.click()

            inputUnderstand = browser.find_element(By.ID,'V1_CTRL115')
            inputUnderstand.click()

            tCheckBox1  = browser.find_element(By.ID, 'V1_CTRL109')
            tCheckBox1.click()
            tCheckBox2  = browser.find_element(By.ID, 'V1_CTRL111')
            tCheckBox2.click()
            tCheckBox3  = browser.find_element(By.ID, 'V1_CTRL114')
            tCheckBox3.click()
            tCheckBox4  = browser.find_element(By.ID, 'V1_CTRL133')
            tCheckBox4.click()
            tCheckBox5  = browser.find_element(By.ID, 'V1_CTRL135')
            tCheckBox5.click()
            tCheckBox6  = browser.find_element(By.ID, 'V1_CTRL136')
            tCheckBox6.click()

            promiseCheckBox  = browser.find_element(By.ID, 'V1_CTRL82')
            promiseCheckBox.click()

            submitBUtton = browser.find_element(By.CLASS_NAME,'command_button')
            submitBUtton.click()

            time.sleep(5)
        except NoSuchElementException:
            print("未找到元素")
        except Exception as e:
            print(e)




    def noPassMessageDialog(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Information)
        msg.setText("未到申报时间或者已经申报过了")
        msg.setWindowTitle("提示")
        msg.setWindowFlag(Qt.WindowStaysOnTopHint)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()

    def passwdWrongMessageDialog(self):
        msg = QMessageBox()
        msg.setIcon(QMessageBox.Warning)
        msg.setText("学号或密码可能输入错误，请点击OK重新确认(不要直接关闭浏览器)")
        msg.setWindowTitle("警告")
        msg.setWindowFlag(Qt.WindowStaysOnTopHint)
        msg.setStandardButtons(QMessageBox.Ok)
        msg.exec_()


# Press the green button in the gutter to run the script.
if __name__ == '__main__':
    myapp = QApplication(sys.argv)
    dialog = MainDialog()
    apply_stylesheet(myapp,theme='light_blue_500.xml')
    dialog.show()
    sys.exit(myapp.exec_())
