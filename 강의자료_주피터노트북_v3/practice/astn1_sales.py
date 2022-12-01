import pandas as pd
from pandas import DataFrame, Series
import cafle
from cafle import Index, Account
from cafle import Setattr

from practice.astn0_overview import idx


#분양매출 테이블 작성
분양_오피스텔 = DataFrame({
    "호실면적" : [  10,   11,   12,   13], #평
    "호실수"  : [ 120,   30,  120,   50], #실
    "평단가"  : [20.0, 20.0, 20.0, 20.0], #백만원/평
},  index   = ['A', 'B', 'C', 'D'])
B = 분양_오피스텔
B['호실면적m2'] = round(B['호실면적'] * cafle.PY, 2)
B['면적소계'] = B['호실면적'] * B['호실수']
B['면적소계m2'] = round(B['호실면적m2'] * B['호실수'], 2)
B['호실단가'] = B['호실면적'] * B['평단가']
B['매출소계'] = B['면적소계'] * B['평단가']

분양_근생 = DataFrame({
    "면적" : [140], #평
    "평단가" : [40.0], #백만원/평
},  index = ['F1'])
B = 분양_근생
B['면적m2'] = round(B['면적'] * cafle.PY, 2)
B['매출소계'] = B['면적'] * B['평단가']

분양테이블 = {
    "오피" : 분양_오피스텔,
    "근생" : 분양_근생,
}


#분양매출
분양매출 = {
    "오피" : 분양테이블['오피']['매출소계'].sum(),
    "근생" : 분양테이블['근생']['매출소계'].sum(),
    "합계" : 분양테이블['오피']['매출소계'].sum() + 분양테이블['근생']['매출소계'].sum()
}


#분양대금 납입일정
대금납입일정 = DataFrame({
    '구분': ['계약금', '중도금1', '중도금2', '중도금3', '중도금4', '잔금'],
    '오피': [   0.1,     0.1,     0.1,      0.1,     0.1,   0.5],
    '근생': [   0.1,     0.1,     0.1,      0.1,     0.0,   0.6],
}, index=  [idx.sales[0], idx.sales[5], idx.sales[10], idx.sales[15], idx.sales[20], idx.sales[-1]])

대금납입일정['납입오피'] = 대금납입일정['오피'] * 분양테이블['오피']['매출소계'].sum()
대금납입일정['납입근생'] = 대금납입일정['근생'] * 분양테이블['근생']['매출소계'].sum()
대금납입일정['납입소계'] = 대금납입일정['납입오피'] + 대금납입일정['납입근생']


#분양률 가정
분양률가정시나리오 = {}

분양률가정시나리오[1] = DataFrame({
    '오피': [   0.2,     0.2,     0.2,      0.2,     0.2],
    '근생': [   0.0,     0.0,     0.0,      0.0,     1.0],
}, index=  [idx.sales[0], idx.sales[6], idx.sales[12], idx.sales[18], idx.sales[-1]])

분양률가정시나리오[2] = DataFrame({
    '오피': [   1.0,     0.0,     0.0,      0.0,     0.0],
    '근생': [   1.0,     0.0,     0.0,      0.0,     0.0],
}, index=  [idx.sales[0], idx.sales[6], idx.sales[12], idx.sales[18], idx.sales[-1]])

분양률가정시나리오[3] = DataFrame({
    '오피': [   0.0,     0.0,     0.0,      0.0,     1.0],
    '근생': [   0.0,     0.0,     0.0,      0.0,     1.0],
}, index=  [idx.sales[0], idx.sales[6], idx.sales[12], idx.sales[18], idx.sales[-1]])

분양률가정 = 분양률가정시나리오[1]
분양률가정['계약오피'] = 분양률가정['오피'] * 분양테이블['오피']['매출소계'].sum()
분양률가정['계약근생'] = 분양률가정['근생'] * 분양테이블['근생']['매출소계'].sum()
분양률가정['계약소계'] = 분양률가정['계약오피'] + 분양률가정['계약근생']


#기본 함수 설정
def 현금스케줄(분양률가정, 대금납입일정):
    rslt = DataFrame(index = idx)
    rslt['계약율'] = 분양률가정
    rslt['납입율'] = 대금납입일정
    rslt = rslt.fillna(0.0)
    
    rslt[['계약율누적', '납입율누적']] = rslt.cumsum()
    rslt['현금율누적'] = rslt['계약율누적'] * rslt['납입율누적']
    rslt['현금율유입'] = rslt['현금율누적'].diff()
    rslt = rslt.fillna(0.0)
    
    return rslt


#Sales Account 설정
sales = Account(idx)
sales.오피 = sales.subacc('오피')
sales.근생 = sales.subacc('근생')

현금스케줄_오피 = 현금스케줄(분양률가정['오피'], 대금납입일정['오피'])
현금스케줄_근생 = 현금스케줄(분양률가정['근생'], 대금납입일정['근생'])

sales.오피.subscd(현금스케줄_오피.index, 현금스케줄_오피['현금율유입'] * 분양테이블['오피']['매출소계'].sum())
sales.근생.subscd(현금스케줄_근생.index, 현금스케줄_근생['현금율유입'] * 분양테이블['근생']['매출소계'].sum())

sales.분양테이블 = 분양테이블
sales.분양매출 = 분양매출
sales.대금납입일정 = 대금납입일정
sales.분양률가정 = 분양률가정


##Sales 함수 설정
@Setattr(sales.dct)
def get_salesamt(sls, acc, idxno):
    amt = sls.scd_out[idxno]
    sls.send(idxno, amt, acc, note=f"분양매출({sls.name})")
    return amt