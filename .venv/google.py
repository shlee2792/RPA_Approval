import time

from selenium import webdriver
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from openpyxl.drawing.image import Image

#이번달 첫번째 날짜, 마지막 날짜 찾기 위함
from datetime import datetime, timedelta
from dateutil import relativedelta

#파일 업로드 창에서 제어하기 위함
import pyperclip, pyautogui

#openpyxl 모듈 import, 엑셀 사용위해서
import  openpyxl  as  op  

#엑셀시트 pdf 저장을 위해서 pywin32 install 
#win32com.client 오류시 pip insatll --upgrade pywin32==225 다운그레이드
import sys, os, win32com.client

#폴더생성 함수 설정
def makedirs(path): 
   try: 
        os.makedirs(path) 
   except OSError: 
       if not os.path.isdir(path): 
           raise


#이건 사용자가 입력할 수 있게? 아니면 고정 
upload_folder_path  = r"C:\RPA_Approval\Upload_file"
image_path = r"C:\RPA_Approval\Image"


#현재 월 받아오기
now = datetime.today()
input_year = str(now.year)
input_month = str(now.month) 
input_month_int = now.month #now.month int형 출력 확인


#영수증 PDF와 파견출장정보 PDF 생성 Part

#월별 주차 입력
week_calc_2022 = [4,4,5,4,4,5,4,4,5,4,4,5]

# 해당 월이 몇주차 인지 확인
week_calc_2022[input_month_int-1]


excel_path = r"C:\RPA_Approval"
wb = op.load_workbook(excel_path + "/sample.xlsx")


if week_calc_2022[input_month_int-1] == 4:
    ws_report = wb['4']
    temp_text = '4'
else:
    ws_report = wb['5']
    temp_text = '5'

# 파견 정보 해당 월 입력
ws_report["E1"].value = input_month + '월'

#식권영수증 추가
ws_receipt = wb['Receipt']

img = Image(image_path + "/영수증.jpg")
img.height = 650
img.width = 325

ws_receipt.add_image(img,"C6")


#변경내용 엑셀 저장
wb.save(excel_path + "/sample.xlsx")
wb.close()


#수정한 엑셀 파일 불러오기
excel = win32com.client.Dispatch("Excel.Application")
excel.Visible = False

#불러온 수정된 엑셀파일을 PDF로 바꿔준다.
wb_for_PDF = excel.Workbooks.Open(excel_path + "/sample.xlsx")

#해당 월의 시트 선택해서 수정
wb_for_PDF.WorkSheets(temp_text).Select()
wb_for_PDF.ActiveSheet.ExportAsFixedFormat(0, upload_folder_path + "/01_{}년_{}월_슈어소프트테크_파견_정보.pdf".format(input_year, input_month))

wb_for_PDF.WorkSheets('Receipt').Select()
wb_for_PDF.ActiveSheet.ExportAsFixedFormat(0, upload_folder_path + "/02_{}년_{}월_식권영수증.pdf".format(input_year, input_month))


wb_for_PDF.Close()
excel.Quit()


#해당폴더에 chromedriver 놓기
driver = webdriver.Chrome(r".\chromedriver.exe")
driver.get("https://suresofttech.hanbiro.net/ngw/app/#/sign")
driver.maximize_window()

driver.implicitly_wait(3)
driver.find_element(By.NAME, 'userid').send_keys("shlee")

driver.switch_to.frame("iframeLoginPassword")

driver.find_element(By.ID, 'p').send_keys('!dltkdgh86')
driver.find_element(By.ID, 'p').send_keys(Keys.RETURN)


driver.implicitly_wait(5)
driver.switch_to.default_content() #iframe 접속 종료
# select_authBtn = driver.find_element_by_xpath("/html/body/div[1]/div[2]/section[3]/div[1]/div/div[4]/div/nav-menu-react/nav/a[3]").click()
driver.find_element(By.XPATH,"//*[@id='main-navi']/nav-menu-react/nav/a[3]").click()



# 완결문서 선택
time.sleep(2)
driver.find_element(By.XPATH,"//*[@id='mCSB_4_container']/div/ul/li[4]/a").click()

documents = driver.find_elements(By.CSS_SELECTOR, 'span.text.approval-priority3')

for chk_document in documents:
    if "파견비신청서" in chk_document.text:
        chk_document.click()
        break

time.sleep(1)
# driver.find_element(By.CSS_SELECTOR, 'button.btn.btn-sm.btn-white.btn-primary').click()
driver.find_element(By.XPATH, '//*[@id="ngw.approval.container "]/split-screen-view/list-view/div/div[1]/div[3]/div/div[3]/div[1]/div/div[1]/button').click()

time.sleep(2)
try:
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR,'body > div.modal.fade.modal-type1.small.in > div > div > div.modal-footer.center > button').click()
    
    time.sleep(1)
    driver.find_element(By.CSS_SELECTOR,'body > div.modal.fade.modal-type1.in > div > div > div.modal-footer > button.btn.btn-sm.btn-info.no-border').click()
    
except:
    #창 안뜸
    print("no alert")


#기존 결재문서에 있는 Title 삭제 후 New Title 설정
driver.find_element(By.XPATH,'//*[@id="write-form"]/div/div[2]/div/div[2]/div/div[1]/div/input').clear()
driver.find_element(By.XPATH,'//*[@id="write-form"]/div/div[2]/div/div[2]/div/div[1]/div/input').send_keys("파견비신청서 - 22년 {}월".format(input_month))

now = datetime.today()
print("현재 날짜:", now.date())
 
this_month_first = datetime(now.year, now.month, 1)
print("이번달 첫째 날짜:", this_month_first.date()) 
 
next_month_first = this_month_first + relativedelta.relativedelta(months=1)
print("다음달 첫째 날짜:", next_month_first.date())
 
this_month_last = next_month_first - timedelta(seconds=1)
print("이번달 마지막 날짜:", this_month_last.date())

input_work_period = str(this_month_first.date()) + ' ~ ' + str(this_month_last.date())

#기안자 이름 가져오기
writter_name = driver.find_element(By.CSS_SELECTOR, "#write-form > div > div.field2.widget-container-col.visible > div > div.widget-body > div > div.approvalDraft.approvalPage.margin-bottom-10 > div > table > tbody > tr:nth-child(1) > td > table:nth-child(2) > tbody > tr > td > div.line-collapse > table > tbody > tr:nth-child(4) > td:nth-child(2)")

#iframe 선택
iframe = driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/section[3]/div[2]/div[2]/div[2]/div/div[3]/div[3]/div/split-screen-view/list-view/div/div[2]/div/div/content-write/div/div/form/div/div[2]/div/div[2]/div/div[2]/div/table/tbody/tr[2]/td/table[2]/tbody/tr/td/div[2]/han-editor/div[3]/div[1]/div[2]/div[1]/iframe")
driver.switch_to.frame(iframe)

#기한 부분 기존 텍스트 삭제 
driver.find_element(By.CSS_SELECTOR, "#tinymce > div > table:nth-child(1) > tbody > tr:nth-child(5) > td:nth-child(2) > p:nth-child(1) > span").clear()
# driver.find_element(By.XPATH, "//*[@id='tinymce']/div/table[1]/tbody/tr[5]/td[2]/p[1]/span").send_keys(str(this_month_first.date()) + ' ~ ')
# driver.find_element(By.CSS_SELECTOR, "#tinymce > div > table:nth-child(1) > tbody > tr:nth-child(5) > td:nth-child(2) > p:nth-child(1) > span").send_keys(str(this_month_first.date()) + ' ~ ')
                                      
driver.find_element(By.CSS_SELECTOR, "#tinymce > div > table:nth-child(1) > tbody > tr:nth-child(5) > td:nth-child(2) > p:nth-child(2) > span").clear()
# driver.find_element(By.XPATH, "//*[@id='tinymce']/div/table[1]/tbody/tr[5]/td[2]/p[2]/span").send_keys(str(this_month_last.date()))
# driver.find_element(By.CSS_SELECTOR, "#tinymce > div > table:nth-child(1) > tbody > tr:nth-child(5) > td:nth-child(2) > p:nth-child(2) > span").send_keys(str(this_month_last.date()))


#새로운 기간 입력
driver.find_element(By.CSS_SELECTOR, "#tinymce > div:nth-child(1) > table:nth-child(1) > tbody > tr:nth-child(5) > td:nth-child(2)").send_keys(input_work_period)

#중앙 정렬 - 다음에 하자...
# driver.find_element(By.XPATH, "//*[@id='tinymce']/div[1]/table[1]/tbody/tr[5]/td[2]").click()
# driver.find_element(By.CSS_SELECTOR,"#tinymce > div:nth-child(1) > table:nth-child(1) > tbody > tr:nth-child(5) > td:nth-child(2)").click() #입력칸 다시 선택
#가운데 정렬 완료
# driver.find_element(By.CSS_SELECTOR,"#write-form > div > div.field2.widget-container-col.visible > div > div.widget-body > div > div.approvalDraft.approvalPage.margin-bottom-10 > div > table > tbody > tr:nth-child(2) > td > table.width-100.bordered-td.no-border-top.approval-activex-area > tbody > tr > td > div.col-sm-12.tab_area.padding-10 > han-editor > div.tox.tox-tinymce > div.tox-editor-container > div.tox-editor-header > div.tox-toolbar-overlord > div.tox-toolbar__primary > div:nth-child(2) > button:nth-child(1)").click()
# driver.find_element(By.XPATH, "/html/body/div[7]/div/div/div/div/div[2]").click()

#iframe 접속 종료
driver.switch_to.default_content() 



#첨부파일 삭제 클릭
driver.find_element(By.XPATH, "//*[@id='attachment-list']/li[1]/span[2]/a/i").click()
driver.find_element(By.CSS_SELECTOR, "body > div.modal.fade.modal-type1.small.in > div > div > div.modal-footer.center > button.btn.btn-sm.btn-info.no-border").click()

time.sleep(1)

driver.find_element(By.XPATH, "//*[@id='attachment-list']/li/span[2]/a/i").click()
driver.find_element(By.CSS_SELECTOR, "body > div.modal.fade.modal-type1.small.in > div > div > div.modal-footer.center > button.btn.btn-sm.btn-info.no-border").click()


#파일 추가 클릭
driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/section[3]/div[2]/div[2]/div[2]/div/div[3]/div[3]/div/split-screen-view/list-view/div/div[2]/div/div/content-write/div/div/form/div/div[2]/div/div[2]/div/div[4]/div/div[1]/div/div[2]/div/div[1]/div/div[1]/div[1]/div[1]/div/span[1]").click()

#첫번째 업로드 파일 경로 입력
pyperclip.copy(upload_folder_path)

for i in range(5):
    # pyautogui.click()
    pyautogui.sleep(1)
    pyautogui.press('tab')

pyautogui.sleep(1)
pyautogui.press('space')
pyautogui.sleep(1)
pyautogui.hotkey("ctrl","v")
pyautogui.sleep(1)
pyautogui.press("enter")

for i in range(4):           
    # pyautogui.click()
    pyautogui.sleep(1)
    pyautogui.press("tab")    

pyautogui.sleep(1)
pyautogui.press("space")
pyautogui.sleep(1)
pyautogui.press("enter")

time.sleep(2)


#파일 추가 클릭
driver.find_element(By.XPATH, "/html/body/div[1]/div[2]/section[3]/div[2]/div[2]/div[2]/div/div[3]/div[3]/div/split-screen-view/list-view/div/div[2]/div/div/content-write/div/div/form/div/div[2]/div/div[2]/div/div[4]/div/div[1]/div/div[2]/div/div[1]/div/div[1]/div[1]/div[1]/div/span[1]").click()

#쉬프트 탭 두번 keydown
for i in range(2):
    # pyautogui.click()
    pyautogui.sleep(1)
    pyautogui.hotkey("shift","tab")

pyautogui.sleep(1)
pyautogui.press("down")
pyautogui.sleep(1)
pyautogui.press("enter")



driver.find_element(By.CSS_SELECTOR, "#ngw\.approval\.container\  > split-screen-view > list-view > div > div.content-header > div:nth-child(3) > div > div:nth-child(3) > div > div.messagebar-item-right > button").click()
time.sleep(1)
driver.find_element(By.CSS_SELECTOR, "body > div.modal.fade.modal-type1.middle.in > div > div > div.modal-body > form > div:nth-child(1) > div > label > span").click()

# driver.find_element(By.CSS_SELECTOR, "body > div.modal.fade.modal-type1.middle.in > div > div > div.modal-footer.center > button.btn.btn-sm.btn-info.no-border").click()

print("===="*40)









