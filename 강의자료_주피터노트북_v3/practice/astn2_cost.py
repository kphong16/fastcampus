import pandas as pd
from pandas import DataFrame, Series
import cafle
from cafle import Index, Account
from cafle import Setattr

from practice.astn0_overview import overview, idx
from practice.astn1_sales import sales


#Cost setting 함수
def costsetting(cst):
    cst.amtlst = Series(
        {key: val.amt for key, val in cst.dct.items()}
    )
    cst.amt = cst.amtlst.sum()
    cst.smry = {
        key: {
            subkey: getattr(subacc, subkey) for subkey in
                [var for var in subacc.vars if var not in ['index', 'name']]
        }
        for key, subacc in cst.dct.items()
    }
#공정률 설정
stdprc = pd.read_excel('data/standard_process_rate.xlsx')
lencstrn = len(idx.cstrn)
cstrnprc = stdprc[lencstrn][:lencstrn]
cstrnprc.index = idx.cstrn
cstrnprc._name = '공정률'


##Cost Account 설정
cost = Account(idx)

#토지비 가정
cost.토지비 = cost.subacc('토지비')
with cost.토지비 as 토지비:
    토지비.토지매입가액 = 토지비.subacc('토지매입가액')
    with 토지비.토지매입가액 as c:
        c.대지면적 = overview['area']['대지면적']['평']
        c.평단가 = 45#백만원/평
        c.amt = c.대지면적 * c.평단가
        c.addscd(idx.loan[0], c.amt)

    토지비.취득세 = 토지비.subacc('취득세')
    with 토지비.취득세 as c:
        c.취득세율 = 0.046
        c.과세표준 = 토지비.토지매입가액.amt
        c.amt = round(c.과세표준 * c.취득세율)
        c.addscd(idx.loan[0], c.amt)
        
    토지비.중개수수료 = 토지비.subacc('중개수수료')
    with 토지비.중개수수료 as c:
        c.수수료율 = 0.01
        c.취득가액 = 토지비.토지매입가액.amt
        c.amt = round(c.취득가액 * c.수수료율)
        c.addscd(idx.loan[0], c.amt)
        
    토지비.부수비용 = 토지비.subacc('부수비용')
    with 토지비.부수비용 as c:
        c.적용비율 = 0.003
        c.취득가액 = 토지비.토지매입가액.amt
        c.amt = round(c.취득가액 * c.적용비율)
        c.addscd(idx.loan[0], c.amt)
        
    costsetting(토지비)
    
#공사비 가정
cost.공사비 = cost.subacc('공사비')
with cost.공사비 as 공사비:
    공사비.철거비 = 공사비.subacc('철거비')
    with 공사비.철거비 as c:
        c.기존건물면적 = 700#평
        c.철거단가 = 0.3#백만원/평
        c.amt = c.기존건물면적 * c.철거단가
        c.addscd(idx.cstrn[0], c.amt)
        
    공사비.건축비 = 공사비.subacc('건축비')
    with 공사비.건축비 as c:
        c.공사단가 = 5.50#백만원
        c.연면적 = overview['area']['연면적']['평']
        c.amt = c.공사단가 * c.연면적
        c.공정률표 = DataFrame(cstrnprc)
        c.공정률표['공정금액'] = c.공정률표['공정률'] * c.amt
        c.addscd(idx.cstrn, c.공정률표['공정금액'])
        
    공사비.인입비 = 공사비.subacc('인입비')
    with 공사비.인입비 as c:
        c.적용비율 = 0.035#백만원
        c.연면적 = overview['area']['연면적']['평']
        c.amt = c.적용비율 * c.연면적
        c.addscd(idx.cstrn[0], c.amt)
        
    costsetting(공사비)
    
#간접공사비 가정
cost.간접공사비 = cost.subacc('간접공사비')
with cost.간접공사비 as 간접공사비:
    간접공사비.설계비 = 간접공사비.subacc('설계비')
    with 간접공사비.설계비 as c:
        c.적용단가 = 0.10#백만원/평
        c.연면적 = overview['area']['연면적']['평']
        c.amt = c.적용단가 * c.연면적
        c.addscd(idx.cstrn[0], c.amt)
        
    간접공사비.감리비 = 간접공사비.subacc('감리비')
    with 간접공사비.감리비 as c:
        c.월단가 = 15.0#백만원/월
        c.amt = c.월단가 * len(idx.cstrn)
        c.addscd(idx.cstrn, [c.월단가] * len(idx.cstrn))
        
    간접공사비.인허가등 = 간접공사비.subacc('인허가등')
    with 간접공사비.인허가등 as c:
        c.amt = 300.0#백만원
        c.addscd(idx.loan[0], c.amt)
        
    costsetting(간접공사비)
    
#판매비 가정
cost.판매비 = cost.subacc('판매비')
with cost.판매비 as 판매비:
    판매비.MH운영비 = 판매비.subacc('MH운영비')
    with 판매비.MH운영비 as c:
        c.월단가 = 60#백만원/월
        c.amt = c.월단가 * len(idx.sales)
        c.addscd(idx.sales, [c.월단가] * len(idx.sales))
        
    판매비.광고비 = 판매비.subacc('광고비')
    with 판매비.광고비 as c:
        c.분양매출 = sales.분양매출['합계']
        c.적용비율 = 0.015
        c.amt = c.분양매출 * c.적용비율
        c.월단가 = c.amt / len(idx.sales)
        c.addscd(idx.sales, [c.월단가] * len(idx.sales))
        
    판매비.분양수수료 = 판매비.subacc('분양수수료')
    with 판매비.분양수수료 as c:
        c.수수료율_오피 = 0.05
        c.수수료율_근생 = 0.08
        c.분양가정테이블 = sales.분양률가정.copy() #deep copy
        c.분양가정테이블['수수료오피'] = c.분양가정테이블['계약오피'] * c.수수료율_오피
        c.분양가정테이블['수수료근생'] = c.분양가정테이블['계약근생'] * c.수수료율_근생
        c.수수료_오피 = c.분양가정테이블['수수료오피'].sum()
        c.수수료_근생 = c.분양가정테이블['수수료근생'].sum()
        c.amt = c.수수료_오피 + c.수수료_근생
        c.addscd(c.분양가정테이블.index, c.분양가정테이블['수수료오피'])
        c.addscd(c.분양가정테이블.index, c.분양가정테이블['수수료근생'])
    
    costsetting(판매비)
    
#일반사업비 가정
cost.일반사업비 = cost.subacc('일반사업비')
with cost.일반사업비 as 일반사업비:
    일반사업비.신탁수수료 = 일반사업비.subacc('신탁수수료')
    with 일반사업비.신탁수수료 as c:
        c.분양매출 = sales.분양매출['합계']
        c.적용비율 = 0.02
        c.amt = c.분양매출 * c.적용비율
        c.addscd(idx.loan[0], c.amt)
    
    일반사업비.시행사운영비 = 일반사업비.subacc('시행사운영비')
    with 일반사업비.시행사운영비 as c:
        c.월단가 = 30.0#백만원
        c.amt = c.월단가 * len(idx.sales)
        c.addscd(idx.sales, c.월단가)
        
    일반사업비.금융조달비 = 일반사업비.subacc('금융조달비')
    with 일반사업비.금융조달비 as c:
        c.감정평가 = 50.0#백만원
        c.사업성평가 = 40.0#백만원
        c.법무법인 = 40.0#백만원
        c.법무사 = 10.0#백만원
        c.amt = c.감정평가 + c.사업성평가 + c.법무법인 + c.법무사
        c.addscd(idx.loan[0], c.amt)
        
    일반사업비.예비비 = 일반사업비.subacc('예비비')
    with 일반사업비.예비비 as c:
        c.분양매출 = sales.분양매출['합계']
        c.적용비율 = 0.01
        c.amt = c.분양매출 * c.적용비율
        c.월단가 = c.amt / len(idx.loan)
        c.addscd(idx.loan, c.월단가)
        
    costsetting(일반사업비)
    
#제세공과금 가정
cost.제세공과금 = cost.subacc('제세공과금')
with cost.제세공과금 as 제세공과금:
    제세공과금.상하수원인자부담금 = 제세공과금.subacc('상하수원인자부담금')
    with 제세공과금.상하수원인자부담금 as c:
        c.amt = 155.0#백만원
        c.addscd(idx.cstrn[0], c.amt)
    
    제세공과금.보존등기비 = 제세공과금.subacc('보존등기비')
    with 제세공과금.보존등기비 as c:
        c.과세표준 = cost.공사비.amt + cost.간접공사비.amt + cost.일반사업비.amt
        c.취득세 = c.과세표준 * 0.028      #2.80%
        c.지방교육세 = c.과세표준 * 0.0016  #0.16%
        c.농어촌특별세 = c.과세표준 * 0.0020 #0.20%
        c.amt = c.취득세 + c.지방교육세 + c.농어촌특별세
        c.addscd(idx.cstrn[-1], c.amt)

    costsetting(제세공과금)
    

##Cost 함수 설정
@Setattr(cost)
def estimate_cost_amt(cost, idxno):
    cst = cost.mrg2
    return cst.scd_in[idxno]

@Setattr(cost)
def pay_cost_amt(cost, acc, idxno):
    amtttl = 0
    for key, item in cost.dct.items():
        for key2, item2 in item.dct.items():
            amt = item2.scd_in[idxno]
            acc.send(idxno, amt, item2, note=item2.name)
            amtttl += amt
    return amtttl
