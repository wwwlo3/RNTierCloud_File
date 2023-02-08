# -*- coding: utf-8 -*-
import requests
import json
import datetime as DT
import re
from openpyxl import load_workbook
from openpyxl import Workbook
import pandas as pd
import numpy as np

INDEX = 0 # Search를 위한 INDEX 변수
Count = 0 # 정상 처리 고객 인원 수
Send_Count = 0 # SMS 예약별 전송 총 횟수 (삭제할 예정_확인용)
clunix_C = pd.read_excel('sample_list.xlsx', sheet_name='Worksheet') # Clunix 고객 명단 file 오픈
#print(clunix_C)

ToTal = clunix_C['사용자명'].count() # 현재 고객 총 인원 수
Rntier_person  = np.array(clunix_C['사용자명'].values) # RNTier Cloud 사용자 이름
Date_card = np.array(clunix_C['가입일'].values) # 카드 등록일
Card_reg = np.array(clunix_C['카드 등록 여부'].values) # 등록 : 'O' / 미등록 : 'X'
Phone_num = np.array(clunix_C['전화'].values) # 휴대폰 번호



# RNTier Cloud 10만클루 만기일에 따른 날짜 연산 함수
def DATE_CAL(conv):
    day30 = conv + DT.timedelta(days=30)# 30일 후
    day53 = conv + DT.timedelta(days=53)# 53일 후
    day60 = conv + DT.timedelta(days=60)# 60일 후

    day30 = str(DT.date(day30.year,day30.month,day30.day))
    day53 = str(DT.date(day53.year,day53.month,day53.day))
    day60 = str(DT.date(day60.year,day60.month,day60.day))
    Cal_DATE = [day30, day53, day60]
    return Cal_DATE # 리스트로 반환

 # rdate 옵션 함수
def RDATE_OPTION(DATE):
    R_list = []
    R_list.append(re.sub(r"[^0-9]", "", DATE[0]))
    R_list.append(re.sub(r"[^0-9]", "", DATE[1]))
    R_list.append(re.sub(r"[^0-9]", "", DATE[2]))  
    return R_list

while True:
    if INDEX >= ToTal:
        print("모든 리스트 작업이 끝났습니다.")
        break
    
     # sms_data에 최적화된 형식으로 변환작업 영역
    Phone_num[INDEX] = re.sub(r"[^0-9]", "", Phone_num[INDEX]) # 전화번호 추출 값을 항상 01012345678 형식으로 저장
    conv = DT.datetime.strptime(Date_card[INDEX], '%Y-%m-%d') # 카드 등록일 yyyymmdd 형식으로 변환
    DATE = DATE_CAL(conv)
    # sms_data_Option_Parameter
    rdate_list = RDATE_OPTION(DATE) # rdate_list / SMS 송신 예약 날짜 리스트 항상 3개의 목록으로 일정
    destination = Phone_num[INDEX] + '|' + Rntier_person[INDEX] # sms_data의 송신할 대상의 폰번호|성함

   
    # 임시 확인 결과 값
    if Card_reg[INDEX] =='O' and Date_card[INDEX] != None and Phone_num[INDEX] !="":
        Count+=1
        print(Rntier_person[INDEX], rdate_list, Card_reg[INDEX], Phone_num[INDEX])
        print("======================================================\n")
        # send_url = 'https://apis.aligo.in/send/' # 요청을 던지는 URL, 현재는 문자보내기

        # # ================================================================== 문자 보낼 때 필수 key값
        # # API key, userid, sender, receiver, msg
        # # API키, 알리고 사이트 아이디, 발신번호, 수신번호, 문자내용

        # print("=================== 카드 등록 완료에 대한 SMS를 요청 완료 하였습니다 ====================\n")
        # 카드 등록 했을 때 메세지 데이터
        # sms_data={'key': 'xx5g032ytpv4xrs4rzdnv74ax6fo4ezz', #api key
        #         'userid': 'clunix', # 알리고 사이트 아이디
        #         'sender': '0234865896', # 발신번호
        #         'receiver': Phone_num[INDEX], # 수신번호 (,활용하여 1000명까지 추가 가능)
        #         'msg': '%고객명%님 안녕하세요. 카드등록이 완료되어 10만 원 크레딧을 지급해드렸습니다.\n - RNTierCloud -', #문자 내용 
        #         'msg_type' : 'SMS', #메세지 타입 (SMS, LMS)
        #         'title' : 'RNTier Cloud 10만클루 이벤트', #메세지 제목 (장문에 적용)
        #         'destination' : destination, # %고객명% 치환용 입력
        #         'rdate' : rdate_list[0],
        #         #'rtime' : '1000',
        #         #'testmode_yn' : '' #테스트모드 적용 여부 Y/N
        # }

        # send_response = requests.post(send_url, data=sms_data)
        # print (send_response.json())
        print("=================== 30일 지났을 때 SMS를 요청 완료 하였습니다 ====================\n")
        # 10만클루 만기일 30일 지났을 때 데이터
        sms_data={'key': 'xx5g032ytpv4xrs4rzdnv74ax6fo4ezz', #api key
                'userid': 'clunix', # 알리고 사이트 아이디
                'sender': '0234865896', # 발신번호
                'receiver': Phone_num[INDEX], # 수신번호 (,활용하여 1000명까지 추가 가능)
                'msg': '%고객명%님 안녕하세요. 10만 이벤트 크레딧 소멸 기간까지 30일 남았습니다.', #문자 내용 
                'msg_type' : 'SMS', #메세지 타입 (SMS, LMS)
                'title' : 'RNTier Cloud 10만클루 이벤트', #메세지 제목 (장문에 적용)
                'destination' : destination, # %고객명% 치환용 입력
                'rdate' : rdate_list[0],
                'rtime' : '1000',
                #'testmode_yn' : '' #테스트모드 적용 여부 Y/N
        }
        print(sms_data)
        Send_Count +=1
        # send_response = requests.post(send_url, data=sms_data)
        # print (send_response.json())

        print("=================== 53일 지났을 때 SMS를 요청 완료 하였습니다 ====================\n")
        # 10만클루 만기일 53일 지났을 때 데이터
        sms_data={'key': 'xx5g032ytpv4xrs4rzdnv74ax6fo4ezz', #api key
                'userid': 'clunix', # 알리고 사이트 아이디
                'sender': '0234865896', # 발신번호
                'receiver': Phone_num[INDEX], # 수신번호 (,활용하여 1000명까지 추가 가능)
                'msg': '%고객명%님 안녕하세요. 10만 이벤트 크레딧 소멸 기간까지 7일 남았습니다.', #문자 내용 
                'msg_type' : 'SMS', #메세지 타입 (SMS, LMS)
                'title' : 'RNTier Cloud 10만클루 이벤트', #메세지 제목 (장문에 적용)
                'destination' : destination, # %고객명% 치환용 입력
                'rdate' : rdate_list[1],
                'rtime' : '1000',
                #'testmode_yn' : '' #테스트모드 적용 여부 Y/N
        }
        print(sms_data)
        Send_Count +=1
        # send_response = requests.post(send_url, data=sms_data)
        # print (send_response.json())

        print("=================== 60일 지났을 때 SMS를 요청 완료 하였습니다 ====================\n")
        # 10만클루 만기일 당일 데이터
        sms_data={'key': 'xx5g032ytpv4xrs4rzdnv74ax6fo4ezz', #api key
                'userid': 'clunix', # 알리고 사이트 아이디
                'sender': '0234865896', # 발신번호
                'receiver': Phone_num[INDEX], # 수신번호 (,활용하여 1000명까지 추가 가능)
                'msg': '%고객명%님 안녕하세요. 10만 이벤트 크레딧 소멸 예정입니다.', #문자 내용 
                'msg_type' : 'SMS', #메세지 타입 (SMS, LMS)
                'title' : 'RNTier Cloud 10만클루 이벤트', #메세지 제목 (장문에 적용)
                'destination' : destination, # %고객명% 치환용 입력
                'rdate' : rdate_list[2],
                'rtime' : '1000',
                #'testmode_yn' : '' #테스트모드 적용 여부 Y/N
        }
        print(sms_data)
        Send_Count +=1
        # send_response = requests.post(send_url, data=sms_data)
        # print (send_response.json())
    INDEX+=1
    

    
    

    
print('정상 처리 된 고객 인원 수 : {} 명'.format(Count))
print('총 예약 SMS 보낸 건 수 : {} 번'.format(Send_Count))