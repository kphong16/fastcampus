from cafle import Index, Account
from pandas import DataFrame

#Business Overview
overview = {}
overview["business"] = {
    "사업명": ["ㅇㅇ동 ㅁㅁ오피스텔 개발사업"],
    "위치": ["서울시 ㄷㄷ구 ㅇㅇ동 123번지 일원"],
    "지역지구": ["일반상업지구, 지구단위계획구역"],
    "주용도": ["업무시설(오피스텔), 판매시설"],
    "규모": ["지상15층 / 지하3층"],
    "주구조": ["철근콘크리트구조"],
    "주차대수": ["법정 161대", "계획 161대"]
}
overview["area"] = DataFrame({
    "대지면적":   [ 1_500,   454],
    "건축면적":   [   750,   227],
    "연면적" :  [ 13_000,  3_933],
    "용적률연면적": [9_500,  2_874],
},  index =       ['m2',   '평'])

#Index
idx = Index("2023.01", 30) #총사업기간 Index
idx.mtrt = 26 #대출 만기
idx.loan = Index("2023.03", idx.mtrt + 1) #자금조달 Index
idx.cstrn = Index("2023.04", 24) #공사기간 Index
idx.sales = Index("2023.04", 25) #분양기간 Index

#Account
acc = Account(idx) #Account 대분류
acc.oprtg = acc.subacc("oprtg") #운영계좌
acc.repay = acc.subacc("repay") #상환계좌
acc.tmp = acc.subacc("tmp") #임시계좌