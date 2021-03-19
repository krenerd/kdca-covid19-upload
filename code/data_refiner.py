from selenium import webdriver
import pandas as pd
import openpyxl
import time
import tkinter as tk
from selenium.webdriver.chrome.options import Options

def refine_data(data_path,root,text_area):
    df=pd.read_excel(data_path,engine='openpyxl')

    chrome_options = Options()
    chrome_options.add_argument("--headless")
    d=webdriver.Chrome(executable_path='chromedriver',options=chrome_options)

    address_list=[]
    error_list=[]
    text_area['state']='normal'
    text_area.insert(tk.INSERT,'\n데이터 수정을 시작합니다. \n')
    text_area.yview_pickplace("end")
    text_area['state']='disabled'
    to_string_columns=['주민등록번호','휴대전화번호']
    for col in to_string_columns:
        df[col]=df[col].astype(str)
    for idx in range(df.shape[0]):
        if (idx+1)%10==0:
            text_area['state']='normal'
            text_area.insert(tk.INSERT,f'{idx+1}/{df.shape[0]} 수정완료\n')
            text_area.yview_pickplace("end")
            text_area['state']='disabled'
            
            
        root.update()
        error_list.append('')
        if df['성명'].isnull()[idx]:
            continue
        #도로명 주소 수정
        if df['도로명주소'].isnull()[idx]==False:
            address=df['도로명주소'][idx]
            d.get('https://www.juso.go.kr/openIndexPage.do')
            d.implicitly_wait(3)
            d.find_element_by_xpath('//*[@id="inputSearchAddr"]').send_keys(address)
            d.find_element_by_xpath('//*[@id="AKCFrm"]/fieldset/div/button').click()
            d.implicitly_wait(3)
            try:
                df.at[idx,'도로명주소']=d.find_element_by_xpath('//*[@id="list1"]/div[1]/span[2]').text
            except:
                error_list[-1]=error_list[-1]+'주소_검색실패  '
        else:
            pass
        #필수 항목 검사
        col_list=['성명','주민등록번호','성별','휴대전화번호','도로명주소','비고(특이사항)','진단의사','검사결과']
        for col in col_list:
            if df[col].isnull()[idx]:
                error_list[-1]=error_list[-1]+col+'미기입  '

        #주민등록번호, 전화번호 양식
        true_num=True
        n=str(df['주민등록번호'][idx]).split('.')[0].replace(' ','-').split('-')
        if len(n)==1:
            if len(n[0])==13:
                df.at[idx,'주민등록번호']=n[0][:6]+'-'+n[0][6:]
            elif len(n[0])<13:
                df.at[idx,'주민등록번호']='0'*(13-len(n[0]))+n[0][:-7]+'-'+n[0][-7:]
                #error_list[-1]=error_list[-1]+'(주의)주민번호를_확인해주세요  '
            else:
                true_num=False
                error_list[-1]=error_list[-1]+'주민등록번호_형식_오류  '
        elif len(n)==2:
            if len(n[0])!=6 or len(n[1])!=7:
                true_num=False
                error_list[-1]=error_list[-1]+'주민등록번호_형식_오류  '
        else:
            true_num=False
            error_list[-1]=error_list[-1]+'주민등록번호_형식_오류  '

        if true_num:
            mult=[2,3,4,5,6,7,8,9,2,3,4,5]
            n=df['주민등록번호'][idx].replace('-','')
            add=0
            for x in range(12):
                add+=mult[x]*int(n[x])
            if not int(n[-1])==(11-add%11)%10:
                error_list[-1]=error_list[-1]+'잘못된_주민등록번호  '
        
        tf=df['휴대전화번호'][idx].split('.')[0]
        n=str(tf).replace(' ','-').split('-')

        if len(n)==1:
            if len(n[0])==9:
                if n[0][0]=='1':
                    #10-xxx-xxxx
                    df.at[idx,'휴대전화번호']='0'+n[0][0:2]+'-'+n[0][2:5]+'-'+n[0][5:]
                else:
                    error_list[-1]=error_list[-1]+'전화번호_형식_오류  '
            elif len(n[0])==10:
                if n[0][0]=='0':
                    #010-xxx-xxxx
                    df.at[idx,'휴대전화번호']=n[0][0:3]+'-'+n[0][3:6]+'-'+n[0][6:]
                elif n[0][0]=='1':
                    #10-xxxx-xxxx
                    df.at[idx,'휴대전화번호']='0'+n[0][0:2]+'-'+n[0][2:6]+'-'+n[0][6:]
            elif len(n[0])==11:
                #010-xxxx-xxxx
                df.at[idx,'휴대전화번호']=n[0][0:3]+'-'+n[0][3:7]+'-'+n[0][7:]
            else:
                error_list[-1]=error_list[-1]+'전화번호_형식_오류  '
        elif len(n)!=3:
            error_list[-1]=error_list[-1]+'전화번호_형식_오류  '

        #Default 검사결과, 입원여부, 환자분류 입력
        #if df['검사결과'].isnull()[idx]:
        #    df['검사결과'][idx]='음성'
        if df['입원여부'].isnull()[idx]:
            df['입원여부'][idx]='2'
        if df['환자분류'].isnull()[idx]:
            df['환자분류'][idx]='4'
            
        #발병일, 진단일, 신고일 양식
        if df['발병일'].isnull()[idx]:
            df['발병일'][idx]='0000 00 00'
        n=df['발병일'][idx].split(' ')
        if len(n)!=3 and len(n[0])!=4 and len(n[1])!=2 and len(n[2])!=2:
            error_list[-1]=error_list[-1]+'발병일_형식_오류  '

        if df['진단일'].isnull()[idx]:
            df['진단일'][idx]='0000 00 00'  
        n=df['진단일'][idx].split(' ')
        if len(n)!=3 and len(n[0])!=4 and len(n[1])!=2 and len(n[2])!=2:
            error_list[-1]=error_list[-1]+'진단일_형식_오류  '

        if df['신고일'].isnull()[idx]:
            df['신고일'][idx]='0000 00 00'
        n=df['신고일'][idx].split(' ')
        if len(n)!=3 and len(n[0])!=4 and len(n[1])!=2 and len(n[2])!=2:
            error_list[-1]=error_list[-1]+'신고일_형식_오류  '
        #특수 사례 1: 체류기간 등등 미기입
        if df['추정감염지역'][idx]=='국내':
            df['추정감염지역'][idx]=np.nan
        col_list=['입국일','체류기간(시작)','체류기간(종료)','추정감염지역']
        s=0
        for col in col_list:
            s+=not df[col].isnull()[idx]
        if s!=0:
            for col in col_list:
                if df[col].isnull()[idx]:
                    df[col][idx]='0000 00 00'
        
    text_area['state']='normal'
    text_area.insert(tk.INSERT,"refined.xlsx로 파일 내보내는 중...")
    df['오류']=pd.Series(error_list)
    df.to_excel(data_path.split('/')[-1].split('.')[0]+'-refined.xlsx',index=False)
    d.quit()
    text_area.insert(tk.INSERT,'sieunpark77@gmail.com')
    text_area['state']='disabled'
