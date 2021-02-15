from selenium import webdriver
from selenium.webdriver.chrome.options import Options
import pandas as pd
import openpyxl
import time
from bs4 import BeautifulSoup
import tkinter as tk
def open_webbrowser():
    global driv
    #Generate Fake User-Agent and open Browser instance
    options = Options()

    #options.add_argument(f'user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.88 Safari/537.36')
    options.add_argument('--ignore-certificate-errors')

    d=webdriver.Chrome(options=options, executable_path='chromedriver')
    #Remove navigator.webdriver Flag using JavaScript
    #d.execute_script("Object.defineProperty(navigator, 'webdriver', {get: () => undefined})")

    d.get('https://covid19.kdca.go.kr/')
    driv=d
def enter_excel(data_path,root,text_area,search_code):
    global driv
    d=driv
    error_names=[]
    df=pd.read_excel(data_path,engine='openpyxl',dtype=str)
    text_area['state']='normal'
    text_area.insert(tk.INSERT,'파일 경로: '+data_path+'\n\n')
    text_area.yview_pickplace("end")
    text_area['state']='disabled'
    #Navigate to webpage for process
    d.switch_to.frame(d.find_element_by_xpath('//*[@id="base"]'))
    d.implicitly_wait(5)
    d.find_element_by_xpath('//*[@id="mCSB_1_container"]/ul/li[1]/a').click()
    time.sleep(0.5)
    d.find_element_by_xpath('//*[@id="mCSB_1_container"]/ul/li[1]/ul/li/a').click()
    time.sleep(0.5)
    d.find_element_by_xpath('//*[@id="mCSB_1_container"]/ul/li[1]/ul/li/ul/li[1]').click()
    time.sleep(0.5)
    
    d.switch_to.frame(d.find_element_by_xpath('//*[@id="contents_body"]/iframe'))

    crashed_indicies=[]
    #Enter Information
    for idx in range(df.shape[0]):
        if (idx+1)%20==0:
            text_area['state']='normal'
            text_area.insert(tk.INSERT,f'{idx+1}/{df.shape[0]} 입력완료\n')
            text_area['state']='disabled'
            
        root.update()
        if not df['오류'].isnull()[idx]:
            crashed_indicies.append(idx)
            continue
        
        step=0
        #보고 버튼 클릭
        try:
            d.find_element_by_xpath('//*[@id="mbtnCreate"]').click()
            d.implicitly_wait(5)
        except:
            for x in range(1,len(d.window_handles)):
                d.switch_to.window(d.window_handles[x])
                d.close()
            d.switch_to.window(d.window_handles[0])
            d.switch_to.frame(d.find_element_by_xpath('//*[@id="base"]'))
            d.switch_to.frame(d.find_element_by_xpath('//*[@id="contents_body"]/iframe'))
            
            d.find_element_by_xpath('//*[@id="pbtnClose"]').click()
            d.implicitly_wait(5)
            d.find_element_by_xpath('//*[@id="mbtnCreate"]').click()
        step=1  
        try:
            #이름 입력
            name=df['성명'][idx]
            if df['성명'].isnull()[idx]:
                continue
            d.find_element_by_xpath('//*[@id="regist_pop"]/div[1]/div[1]/div[1]/div[1]/table/tbody/tr/td/table/tbody/tr[1]/td[1]/input').send_keys(name)
            step=2
            #주민등록번호 입력
            id_1,id_2=df['주민등록번호'][idx].split('-')
            foreign=df['외국인'][idx]
            d.find_element_by_xpath('//*[@id="ptxtPatntIhidnum1"]').send_keys(id_1)
            d.find_element_by_xpath('//*[@id="ptxtPatntIhidnum2"]').send_keys(id_2)
            step=3
            if df['외국인'].isnull()[idx]==False:
                #외국인일 시
                d.find_element_by_xpath('//*[@id="pchkFrgnrAt"]').click()
            step=4
            #성별, 나이 입력
            s=df['성별'][idx]
            if s=='여' or str(s)=='2':
                d.find_element_by_xpath('//*[@id="ptxtPatntSexdstnCd"]/option[3]').click()
            elif s=='남'or str(s)=='1':
                d.find_element_by_xpath('//*[@id="ptxtPatntSexdstnCd"]/option[2]').click()
            else:
                raise ValueError('성별이 잘못 기입되었습니다. ')
            step=5
            #직업 입력
            job_dict={'의회의원,고위임직원 및 관리자':1, '전문가':2,'기술공 및 준전문가':3,'사무종사자':4,'서비스종사자':5,'판매종사자':6,
                      '농업 및 어업숙련 종사자':7, '기능원 및 기능관련 종사자':8,'장치,기계조작 및 조립종사자':9,'단순노무 종사자':10,'군인':11,
                      '(전업)주부':12, '학생':13, '무직':14, '기타':15}
            job=df['직업'][idx]
            specific_job=df['상세직업'][idx]
            if type(job)==str and not job.isnumeric():
                #인덱스 번호로 변환 오류 직업은 모두 기타로 처리 
                if not job in job_dict:
                    job='기타'
                job=job_dict[job]
            d.find_element_by_xpath(f'//*[@id="pcmbPatntOccpCd"]/option[{str(int(job)+1)}]').click()
            if df['상세직업'].isnull()[idx]==False:
                d.find_element_by_xpath('//*[@id="ptxtOccpDtlInfo"]').send_keys(specific_job)
            step=6
            #휴대전화 입력
            no_1,no_2,no_3=df['휴대전화번호'][idx].split('-')
            d.find_element_by_xpath('//*[@id="ptxtPatntMbtlnum1"]').send_keys(no_1)
            d.find_element_by_xpath('//*[@id="ptxtPatntMbtlnum2"]').send_keys(no_2)
            d.find_element_by_xpath('//*[@id="ptxtPatntMbtlnum3"]').send_keys(no_3)
            step=7
            #주소 입력
            specific=df['상세주소'][idx]
            general=df['도로명주소'][idx]
            d.find_element_by_xpath('//*[@id="pbtnSearchRdnmadr"]').click()
            d.switch_to.window(d.window_handles[1])
            time.sleep(1)
            d.find_element_by_xpath('//*[@id="keyword"]').send_keys(general)
            d.find_element_by_xpath('//*[@id="serarchContentBox"]/div[1]/fieldset/span/input[2]').click()
            time.sleep(.5)
            d.find_element_by_xpath('//*[@id="roadAddrDiv1"]').click()
            time.sleep(.5)
            if df['상세주소'].isnull()[idx]==False:
                d.find_element_by_xpath('//*[@id="rtAddrDetail"]').send_keys(specific)
            d.find_element_by_xpath('//*[@id="resultData"]/div/a').click()
            
            d.switch_to.window(d.window_handles[0])
            d.switch_to.frame(d.find_element_by_xpath('//*[@id="base"]'))
            d.switch_to.frame(d.find_element_by_xpath('//*[@id="contents_body"]/iframe'))
            time.sleep(2)
            step=8
            #증상및증후
            symptoms=df['증상및징후'][idx]
            if df['증상및징후'].isnull()[idx]:
                symptoms='-'
            d.find_element_by_xpath('//*[@id="ptxtEidsSymptms"]').send_keys(symptoms)
            step=9
            #요양기관(상록수보건소 선택)
            search_code='31700543'
            d.implicitly_wait(5)
            d.find_element_by_xpath('//*[@id="pbtnSearchSttemntMdlcnst"]').click()
            d.implicitly_wait(5)
            d.find_element_by_xpath('//*[@id="txtSearchinsttNm"]').send_keys(search_code)
            d.implicitly_wait(5)
            d.find_element_by_xpath('//*[@id="ibtnPopSearch"]').click()
            time.sleep(5)
            d.find_element_by_xpath('//*[@id="hptlPopList"]/tbody').click()
            step=10
            #발병,진단,신고 일자
            yy,mm,dd=df['발병일'][idx].split()
            d.find_element_by_xpath('//*[@id="ptxtAtfssDe1"]').send_keys(yy)
            d.find_element_by_xpath('//*[@id="ptxtAtfssDe2"]').send_keys(mm)
            d.find_element_by_xpath('//*[@id="ptxtAtfssDe3"]').send_keys(dd)
            yy,mm,dd=df['진단일'][idx].split()
            d.find_element_by_xpath('//*[@id="ptxtDgnssDe1"]').send_keys(yy)
            d.find_element_by_xpath('//*[@id="ptxtDgnssDe2"]').send_keys(mm)
            d.find_element_by_xpath('//*[@id="ptxtDgnssDe3"]').send_keys(dd)
            yy,mm,dd=df['신고일'][idx].split()
            d.find_element_by_xpath('//*[@id="ptxtSttemntDe1"]').send_keys(yy)
            d.find_element_by_xpath('//*[@id="ptxtSttemntDe2"]').send_keys(mm)
            d.find_element_by_xpath('//*[@id="ptxtSttemntDe3"]').send_keys(dd)
            step=11
            #검사결과,입원여부,환자분류
            result_dict={'양성':1,'음성':2,'검사 진행중':3,'검사 미실시':4}
            admission_dict={'외래':2,'입원':1,'그 밖의 경우':3}
            patient_dict={'환자':1,'의사환자':2,'병원체보유자':3,'검사 거부자':5,'그 밖의 경우(환자아님)':4}

            result=df['검사결과'][idx]
            if type(result)==str and not result.isnumeric():
                result=str(result_dict[result])
            d.find_element_by_xpath(f'//*[@id="prdoDsndgnssInspctResultTyCd{result}"]').click()
            if result==1:
                d.switch_to.alert.accept()
                
            result=df['입원여부'][idx]
            if type(result)==str and not result.isnumeric():
                result=str(admission_dict[result])
            d.find_element_by_xpath(f'//*[@id="prdoHsptlzTyCd{result}"]').click()

            result=df['환자분류'][idx]
            if type(result)==str and not result.isnumeric():
                result=str(patient_dict[result])
            d.find_element_by_xpath(f'//*[@id="prdoPatntClCd{result}"]').click()
            
            step=12
            #특이사항 작성
            text=df['비고(특이사항)'][idx]
            d.find_element_by_xpath('//*[@id="ptxtRmInfo"]').clear()
            d.find_element_by_xpath('//*[@id="ptxtRmInfo"]').send_keys(text)
            step=13
            #진단의사 작성
            text=df['진단의사'][idx]
            d.find_element_by_xpath('//*[@id="ptxtSttemntDoctrNm"]').send_keys(text)
            step=14
            #국적 작성
            if df['외국인'].isnull()[idx]==False:
                d.find_element_by_xpath('//*[@id="ptxtPatntNlty"]').send_keys(df['외국인'][idx])
            step=15
            #소속 기관 작성(선택)
            if df['환자소속기관명'].isnull()[idx]==False:
                d.find_element_by_xpath('//*[@id="ptxtPatntPstinstNm"]').send_keys(df['환자소속기관명'][idx])
            if df['환자소속기관 시도 '].isnull()[idx]==False:
                d.find_element_by_xpath('//*[@id="ptxtPatntPstinstCtprvnNm"]').send_keys(df['환자소속기관주소 시도'][idx])
            if df['환자소속기관 시군구'].isnull()[idx]==False:
                d.find_element_by_xpath('//*[@id="ptxtPatntPstinstSignguNm"]').send_keys(df['환자소속기관주소 시군구'][idx])
            if df['환자소속기관 읍면동'].isnull()[idx]==False:
                d.find_element_by_xpath('//*[@id="ptxtPatntPstinstEmdNm"]').send_keys(df['환자소속기관주소 읍면동'][idx])
            if df['환자소속기관 상세주소'].isnull()[idx]==False:
                d.find_element_by_xpath('//*[@id="ptxtPatntPstinstAdres"]').send_keys(df['환자소속기관 상세주소'][idx])
            step=16
            #감염구역
            if df['추정감염지역'].isnull()[idx]==False:
                #국외 감염의 경우
                d.find_element_by_xpath('//*[@id="prdoPrsmpInfcAreaTyCd2"]').click()
                d.switch_to.alert.accept()
                d.find_element_by_xpath('//*[@id="ptxtPrsmpInfcNationNm"]').send_keys(df['추정감염지역'][idx])
                if df['체류기간(시작)'].isnull()[idx]==False:
                    yy,mm,dd=df['체류기간(시작)'][idx].split()
                    d.find_element_by_xpath('//*[@id="ptxtStayBeginDe1"]').send_keys(yy)
                    d.find_element_by_xpath('//*[@id="ptxtStayBeginDe2"]').send_keys(mm)
                    d.find_element_by_xpath('//*[@id="ptxtStayBeginDe3"]').send_keys(dd)
                if df['체류기간(종료)'].isnull()[idx]==False:
                    yy,mm,dd=df['체류기간(종료)'][idx].split()
                    d.find_element_by_xpath('//*[@id="ptxtStayEndDe1"]').send_keys(yy)
                    d.find_element_by_xpath('//*[@id="ptxtStayEndDe2"]').send_keys(mm)
                    d.find_element_by_xpath('//*[@id="ptxtStayEndDe3"]').send_keys(dd)
                if df['입국일'].isnull()[idx]==False:
                    yy,mm,dd=df['입국일'][idx].split()
                    d.find_element_by_xpath('//*[@id="ptxtPatntErccrDe1"]').send_keys(yy)
                    d.find_element_by_xpath('//*[@id="ptxtPatntErccrDe2"]').send_keys(mm)
                    d.find_element_by_xpath('//*[@id="ptxtPatntErccrDe3"]').send_keys(dd)
            step=17

            #최종 확인 및 보고
            d.find_element_by_xpath('//*[@id="pchkNA0012ErrCheck"]').click()
            if job!='14' and job!='12':
                d.find_element_by_xpath('//*[@id="pchkOccpCheck"]').click()
            if df['주민등록번호'][idx].split('-')[1][0]>'4':
                d.find_element_by_xpath('//*[@id="pchkErrCheck"]').click()

            d.find_element_by_xpath('//*[@id="pbtnCreateReport"]').click()
            d.switch_to.alert.accept()
            name=df['성명'][idx]
            text_area.insert(tk.INSERT,name+'의 정보를 성공적으로 기입했습니다.\n')
        except:
            crashed_indicies.append(idx)
            error_dict=['보고 버튼', '이름', '주민 번호','외국인','주민등록번호/성별','직업','휴대번호','주소','증상및징후',
                        '요양기관','일자','검사결과,입원여부,환자분류','비고(특이사항)','진단의사','국적','소속기관','해외감염','최종보고']
            name=df['성명'][idx]
            text_area['state']='normal'
            text_area.insert(tk.INSERT,name+'의 정보를 입력하는 중 '+error_dict[step]+'에서 오류가 발생했습니다.\n')
            text_area.yview_pickplace("end")
            text_area['state']='disabled'
            error_names.append(name)
            

    d.find_element_by_xpath('//*[@id="pbtnClose"]').click()
    text_area['state']='normal'
    text_area.insert(tk.INSERT,'\n기입에 실패한 사람 목록\n')
    text_area.insert(tk.INSERT,str(error_names))
    if len(crashed_indicies)==0:
        text_area.insert(tk.INSERT,'\n모두 성공적으로 입력했습니다.\n')
    else:
        file_name=data_path.split('/')[-1].split('.')[0]+'-error.xlsx'
        text_area.insert(tk.INSERT,'\n입력에 실패한 사람들을 '+ file_name+'에 저장합니다.\n')
        new_df=pd.DataFrame(data=df.loc[crashed_indicies,:],columns=df.columns)
        new_df.to_excel(file_name,index=False)
    text_area.insert(tk.INSERT,'sieunpark77@gmail.com\n')
    text_area['state']='disabled'
    d.quit()
