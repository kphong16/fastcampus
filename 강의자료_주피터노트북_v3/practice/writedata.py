from datetime import date
import pandas as pd
from pandas import DataFrame, Series
from cafle import Write, extnddct

from practice.cashflow import overview, idx, acc, sales, cost, equity, loan

#################################
##1. Excel Workbook 생성
wb = Write("Result_Cashflow.xlsx")


#################################
##2. 사업개요 및 Assumption 데이터 출력
ws = wb.add_ws("assumption")
wb.ws["assumption"].set_column("A:X", 12)

ws("ASSUMPTION", wb.bold)                  #sheet 제목 출력
ws("Written at: " + wb.now)                #작성 일시 출력
ws.nextcell(2)                             #두칸 띄우기

#사업개요 출력
ws("[Business Overview]", wb.bold)
ws(overview["business"], fmtkey=wb.bold, fmt=wb.nml, valdrtn='col', drtn='col')
ws.nextcell(1)

#사업면적 출력
ws("[Business Area]", wb.bold)
ws(overview["area"], fmtkey=wb.bold, fmt=wb.nml, valdrtn='col', drtn='col')
ws.nextcell(1)

#Index 출력
ws("[Index]", wb.bold)
fmtd = [wb.bold, wb.month, wb.date, wb.date]
ws(["", "개월수", "시작일", "종료일"])
ws(["사업기간", len(idx) - 1, idx[0], idx[-1]], fmtd)
ws(["매출기간", len(idx.sales), idx.sales[0], idx.sales[-1]], fmtd)
ws(["대출기간", idx.mtrt, idx.loan[0], idx.loan[-1]], fmtd)
ws(["건축기간", len(idx.cstrn), idx.cstrn[0], idx.cstrn[-1]], fmtd)
ws.nextcell(1)

#Equity 출력
ws("[Equity]", wb.bold)
ws({"구분": [_.name for _ in equity.dct.values()]}, fmtkey=wb.bold, fmt=wb.nml, valdrtn='col')
ws({"출자금액": [_.amt for _ in equity.dct.values()]}, fmtkey=wb.bold, fmt=wb.num, valdrtn='col')
ws.nextcell(1)

#Loan 출력
ws("[Loan]", wb.bold)
ws({"구분": [_.name for _ in loan.dct.values()]}, fmtkey=wb.bold, fmt=wb.nml, valdrtn='col')
ws({"순위": [_.rank for _ in loan.dct.values()]}, fmtkey=wb.bold, fmt=wb.num, valdrtn='col')
ws({"대출금액": [_.ntnl.amt for _ in loan.dct.values()]}, fmtkey=wb.bold, fmt=wb.num, valdrtn='col')
ws({"최초인출": [_.ntnl.intlamt for _ in loan.dct.values()]}, fmtkey=wb.bold, fmt=wb.num, valdrtn='col')
ws({"한도인출": [_.ntnl.amt - _.ntnl.intlamt for _ in loan.dct.values()]}, fmtkey=wb.bold, fmt=wb.num, valdrtn='col')
ws({"수수료율": [_.fee.rate for _ in loan.dct.values()]}, fmtkey=wb.bold, fmt=wb.pct, valdrtn='col')
ws({"금리": [_.IR.rate for _ in loan.dct.values()]}, fmtkey=wb.bold, fmt=wb.pct, valdrtn='col')
ws.nextcell(1)

ws(["만기", idx.mtrt], [wb.bold, wb.month])
ws(["총 대출금액", sum([ln.ntnl.amt for ln in loan.dct.values()])], [wb.bold, wb.num])
ws.nextcell(1)

#분양매출 가정 출력
ws.setcell(4, 6)

ws("[분양매출가정_오피스텔]", wb.bold)
slstbl = sales.분양테이블['오피']
slssum = DataFrame(slstbl.sum(), columns=['합계']).T
slstbl = pd.concat([slstbl, slssum])
ws(slstbl, fmtkey=wb.bold, fmt=wb.num, valdrtn='row', drtn='col')
ws.nextcell(1)

ws("[분양매출가정_근생시설]", wb.bold)
slstbl = sales.분양테이블['근생']
slssum = DataFrame(slstbl.sum(), columns=['합계']).T
slstbl = pd.concat([slstbl, slssum])
ws(slstbl, fmtkey=wb.bold, fmt=wb.num, valdrtn='row', drtn='col')
ws.nextcell(1)

ws("[분양매출금액]", wb.bold)
ws(sales.분양매출, fmtkey=wb.bold, fmt=wb.num, valdrtn='row', drtn='col')
ws.nextcell(1)

ws("[분양대금 납입일정]", wb.bold)
tmpfmt = {'구분':wb.nml, '오피':wb.pct, '근생':wb.pct, '납입오피':wb.num, '납입근생':wb.num, '납입소계':wb.num}
slstbl = sales.대금납입일정
slssum = DataFrame(slstbl.sum(), columns=['합계']).T
slstbl = pd.concat([slstbl, slssum])
ws(slstbl, fmtidx=wb.date, fmtkey=wb.bold, fmt=tmpfmt, valdrtn='row', drtn='col')
ws.nextcell(1)

ws("[분양률 가정]", wb.bold)
tmpfmt = {'오피':wb.pct, '근생':wb.pct, '계약오피':wb.num, '계약근생':wb.num, '계약소계':wb.num}
slstbl = sales.분양률가정
slssum = DataFrame(slstbl.sum(), columns=['합계']).T
slstbl = pd.concat([slstbl, slssum])
ws(slstbl, fmtidx=wb.date, fmtkey=wb.bold, fmt=tmpfmt, valdrtn='row', drtn='col')


#################################
##3. Cashflow 출력
ws = wb.add_ws("cashflow")
wb.ws["cashflow"].set_column("A:A", 12)

ws("CASH FLOW", wb.bold)
ws("Written at: " + wb.now)
ws.nextcell(2)

#Cashflow data 출력
cell = ws("[Cashflow]", wb.bold)
ws.nextcell(2)
ws(idx, wb.date, valdrtn='col', drtn='col')
ws("합계", wb.nml)

ws.setcell(cell)
ws.nextcell(1, drtn='col')

사업비 = {key: {key2: list(item2.df.amt_in) for key2, item2 in item.dct.items()} for key, item in cost.dct.items()}
cfdct = {
    '운영_기초': extnddct(
        {'기초잔액': acc.oprtg.df.bal_strt},
    ),
    '자금조달': extnddct(
        {'Eqt_'+key: list(item.amt_out) for key, item in equity.dct.items()},
        {'Loan_'+key: list(item.ntnl.df.amt_out) for key, item in loan.dct.items()},
    ),
    '분양매출': extnddct(
        {'분양_'+key: list(item.amt_out) for key, item in sales.dct.items()},
    ),
    '운영_유입': extnddct(
        {'현금유입': list(acc.oprtg.df.amt_in)},
    ),
    '금융비용': extnddct(
        {'Fee_'+key: list(item.fee.df.amt_in) for key, item in loan.dct.items()},
        {'IR_'+key: list(item.IR.df.amt_in) for key, item in loan.dct.items()}
    ),
    **사업비,
    '상환_loan': extnddct(
        {'상환'+key: list(item.ntnl.df.amt_in) for key, item in loan.dct.items()},
    ),
    '상환_Eqt': extnddct(
        {'Eqt_'+key: list(item.df.amt_in) for key, item, in equity.dct.items()}
    ),
    '운영_유출': extnddct(
        {'현금유출': acc.oprtg.df.amt_out},
    ),
    '운영_기말': extnddct(
        {'기말잔액': acc.oprtg.df.bal_end}
    ),
}

dctsum = {}
for key, dct in cfdct.items():
    dctsum[key] = {}
    for key2, srs in dct.items():
        if key2 in ['기초잔액', '기말잔액']:
            tmpval = '-'
        else:
            tmpval = sum(srs)
        dctsum[key][key2] = pd.concat([Series(srs), Series([tmpval], index=['합계'])])    
ws(dctsum, fmtkey=wb.bold, fmt=wb.num, valdrtn='row', drtn='col')
cell = ws.nextcell(2)


#################################
##4. 요약 Cashflow 출력
ws.setcell(cell.row, 0)
cell = ws("[요약 Cashflow]", wb.bold)
ws.nextcell(2)
ws(idx, wb.date, valdrtn='col', drtn='col')
ws("합계", wb.nml)
ws.setcell(cell)
ws.nextcell(1, drtn='col')

cfdct = {
    '운영_기초': extnddct(
        {'기초잔액': acc.oprtg.df.bal_strt},
    ),
    '자금조달': extnddct(
        {'Eqt_'+key: list(item.amt_out) for key, item in equity.dct.items()},
        {'Loan_'+key: list(item.ntnl.df.amt_out) for key, item in loan.dct.items()},
    ),
    '분양매출': extnddct(
        {'분양_'+key: list(item.amt_out) for key, item in sales.dct.items()},
    ),
    '운영_유입': extnddct(
        {'현금유입': list(acc.oprtg.df.amt_in)},
    ),
    '금융비용': extnddct(
        {'Fee_'+key: list(item.fee.df.amt_in) for key, item in loan.dct.items()},
        {'IR_'+key: list(item.IR.df.amt_in) for key, item in loan.dct.items()}
    ),
    '사업비': extnddct(
        {key: list(item.mrg.df.amt_in) for key, item in cost.dct.items()},
    ),
    '상환_loan': extnddct(
        {'상환'+key: list(item.ntnl.df.amt_in) for key, item in loan.dct.items()},
    ),
    '상환_Eqt': extnddct(
        {'Eqt_'+key: list(item.df.amt_in) for key, item, in equity.dct.items()}
    ),
    '운영_유출': extnddct(
        {'현금유출': acc.oprtg.df.amt_out},
    ),
    '운영_기말': extnddct(
        {'기말잔액': acc.oprtg.df.bal_end}
    ),
}

dctsum = {}
for key, dct in cfdct.items():
    dctsum[key] = {}
    for key2, srs in dct.items():
        if key2 in ['기초잔액', '기말잔액']:
            tmpval = '-'
        else:
            tmpval = sum(srs)
        dctsum[key][key2] = pd.concat([Series(srs), Series([tmpval], index=['합계'])]) 
        
ws(dctsum, fmtkey=wb.bold, fmt=wb.num, valdrtn='row', drtn='col')


#################################
##4. Financing data 출력
ws = wb.add_ws("financing")
wb.ws["financing"].set_column("A:A", 12)

ws("FINANCING", wb.bold)
ws("Written at: " + wb.now)
cell = ws.nextcell(2)

ws.nextcell(3)
ws(idx, wb.date, valdrtn='col', drtn='col')
ws("합계", wb.nml)

ws.setcell(cell)
ws.nextcell(1, drtn='col')

fncdct = {}
for key, ln in loan.dct.items():
    fncdct['Loan_'+key] = wb.dctprt_loan(ln)

eqtdct = {}
for key, eqt in equity.dct.items():
    eqtdct['Eqt_'+key] = {
        '인출한도': eqt._df.scd_out,
        '상환예정': eqt._df.scd_in,
        '인출금액': eqt._df.amt_out,
        '상환금액': eqt._df.amt_in,
        '대출잔액': eqt._df.bal_end,
    }
fncdct['Equity'] = eqtdct

ttldct = {}
for key, dct in fncdct.items():
    ttldct[key] = {}
    for key2, dct2 in dct.items():
        ttldct[key][key2] = {}
        for key3, srs in dct2.items():
            if key3 in ['대출잔액', '누적이자', '누적수수료', '누적미인출']:
                tmpval = '-'
            else:
                tmpval = sum(srs)
            ttldct[key][key2][key3] = pd.concat([srs, Series([tmpval], index=['합계'])])

ws(ttldct, fmtkey=wb.bold, fmt=wb.num, valdrtn='row', drtn='col')


#################################
##5. Balance Sheet 출력
ws = wb.add_ws("balancesheet")
wb.ws["balancesheet"].set_column("A:X", 12)

ws("BALANCE SHEET", wb.bold)
ws("Written at: " + wb.now)
cell = ws.nextcell(2)


ws("[[분양매출]]", wb.bold)
ws("오피스텔", wb.bold, drtn='row')
slstbl = sales.분양테이블['오피'][['호실면적', '호실수', '평단가', '매출소계']]
slssum = DataFrame(slstbl.sum(), columns=['합계']).T
slstbl = pd.concat([slstbl, slssum])
ws(slstbl, fmtkey=wb.bold, fmt=wb.num, valdrtn='row', drtn='col')
ws.nextcell(-1, drtn='col')

ws("근생시설", wb.bold, drtn='row')
slstbl = sales.분양테이블['근생'][['면적', '평단가', '면적m2', '매출소계']]
slssum = DataFrame(slstbl.sum(), columns=['합계']).T
slstbl = pd.concat([slstbl, slssum])
ws(slstbl, fmtkey=wb.bold, fmt=wb.num, valdrtn='row', drtn='col')
ws.nextcell(-1, drtn='col')

ttl_sales = sales.분양매출['합계']
ws(["분양매출 합계", "", "", "", "", ttl_sales], [wb.bold, wb.nml, wb.nml, wb.nml, wb.nml, wb.numb])
ws.nextcell(2)


ws("[[사업비]]", wb.bold)
ws(["구분1", "구분2", "", "", "", "금액"], wb.bold)
ttl_costs = 0
#Write operating costs
for key1, item1 in cost.dct.items():
	ws(key1, wb.bold, drtn='row')
	subttl = 0
	for key2, item2 in item1.dct.items():
		amt = item2.bal_end[-1]
		ws([key2, "", "", "", amt], [wb.nml, wb.nml, wb.nml, wb.nml, wb.num])
		subttl += amt
	ws(["소계", "", "", "", subttl], [wb.bold, wb.nml, wb.nml, wb.nml, wb.numb])
	ws.nextcell(-1, drtn='col')
	ttl_costs += subttl
#Write financing costs
for key, item in loan.dct.items():
	ws("대출"+key, wb.bold, drtn='row')
	subttl = 0
	amt_fee = item.fee.bal_end[-1]
	ws(["Fee_"+key, "", "", "", amt_fee], [wb.nml, wb.nml, wb.nml, wb.nml, wb.num])
	amt_IR = item.IR.bal_end[-1]
	ws(["IR_"+key, "", "", "", amt_IR], [wb.nml, wb.nml, wb.nml, wb.nml, wb.num])
	subttl = amt_fee + amt_IR
	
	ws(["소계", "", "", "", subttl], [wb.bold, wb.nml, wb.nml, wb.nml, wb.numb])
	ws.nextcell(-1, drtn='col')
	ttl_costs += subttl
ws(["총사업비 합계", "", "", "", "", ttl_costs], [wb.bold, wb.nml, wb.nml, wb.nml, wb.nml, wb.numb])
ws.nextcell(2)
	

ws("[[사업이익]]", wb.bold)
ttl_profit = ttl_sales - ttl_costs
ws(["총 사업이익", "", "", "", "", ttl_profit], [wb.bold, wb.nml, wb.nml, wb.nml, wb.nml, wb.numb])
ws(["매출액 대비 이익률", "", "", "", "", ttl_profit / ttl_sales], [wb.nml, wb.nml, wb.nml, wb.nml, wb.nml, wb.pct])
ws(["총사업비 대비 이익률", "", "", "", "", ttl_profit / ttl_costs], [wb.nml, wb.nml, wb.nml, wb.nml, wb.nml, wb.pct])


#################################
##6. 엑셀 워크북 클로징
wb.close()

