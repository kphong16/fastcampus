from practice.astn0_overview import overview, idx, acc
from practice.astn1_sales import sales
from practice.astn2_cost import cost
from practice.astn3_equity import equity
from practice.astn4_loan import loan


for idxno in idx:
    #운영계좌 잔액 확인
    _ = acc.oprtg.bal_strt[idxno]
    acc.tmp.addamt(idxno, _, note="월초 잔액")

    #자기자본 인출
    _ = equity.withdraw_equity_amount(idxno, acc.oprtg)
    acc.tmp.addamt(idxno, _, note="자기자본 납입")

    #사업비 추정
    oprtg_estmtd = cost.estimate_cost_amt(idxno)

    lncst_estmtd = 0
    for ln in loan.getloan(reverse=True):
        lncst_estmtd += ln.estimate_fee_amt(idxno)
        lncst_estmtd += ln.estimate_IR_amt(idxno)

    cst_estmtd = oprtg_estmtd + lncst_estmtd
    acc.tmp.subscd(idxno, cst_estmtd, note="추정 사업비")
    
    #분양매출 인식
    for sls in sales.dct.values():
        _ = sls.get_salesamt(acc.oprtg, idxno)
        acc.tmp.addamt(idxno, _, note=f"분양매출({sls.name})")

    #대출원금 인출
    for ln in loan.getloan(reverse=True):
        ln.set_loan_withdrawable(idxno)
        _ = ln.withdraw_ntnl_fixed(acc.oprtg, idxno)
        acc.tmp.addamt(idxno, _, note=f"일시대출금({ln.name})")

    for ln in loan.getloan(reverse=True):
        _ = ln.withdraw_ntnl_flexible(acc.tmp, acc.oprtg, idxno)
        acc.tmp.addamt(idxno, _, note=f"한도대출금({ln.name})")

    #금융비용 지출
    for ln in loan.getloan(reverse=False):
        fee = ln.pay_fee_amt(acc.oprtg, idxno)
        IR = ln.pay_IR_amt(acc.oprtg, idxno)
        acc.tmp.subamt(idxno, fee, note=f"수수료({ln.name})")
        acc.tmp.subamt(idxno, IR, note=f"이자({ln.name})")

    #일반사업비 지출
    _ = cost.pay_cost_amt(acc.oprtg, idxno)
    acc.tmp.subamt(idxno, _, note="일반사업비")

    #대출원금 상환
    for ln in loan.getloan(reverse=False):
        amtrpy = ln.repay_ntnl_amt(acc.oprtg, idxno)
        acc.tmp.subamt(idxno, amtrpy, note=f"대출금상환({ln.name})")
        ln.setback_loan_unwithdrawable(idxno)
        
    #자기자본 상환
    if idxno == idx[-1]: #at business end
        _ = equity.repay_equity_amt(idxno, acc.oprtg)
        acc.tmp.subamt(idxno, _, note="자기자본 상환")

    #임시계좌 현금 조정
    acc.tmp.subamt(idxno, acc.tmp.bal_end[idxno], note="임시계좌 잔액 조정")