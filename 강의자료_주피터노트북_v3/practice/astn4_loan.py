import pandas as pd
from pandas import DataFrame, Series
import cafle
from cafle import Index, Account
from cafle import Setattr, round_up

from practice.astn0_overview import overview, idx


#Loan account 설정
loan = Account(idx)
with loan as l:
    l.mtrt = idx.mtrt
    l.is_repaid_all = False

#Loan tra. 설정
loan.tra = loan.subacc('tra')
with loan.tra as tra:
    tra.rank = 0
    tra.is_wtdrbl = False
    tra.is_repaid = False
    
    tra.ntnl = tra.subacc('ntnl')
    with tra.ntnl as n:
        n.amt = 40_000#백만원
        n.intlamt = 5_000#백만원
        n.subscd(idx.loan[0], n.amt)
        n.addscd(idx.loan[-1], n.amt)
        
    tra.IR = tra.subacc('IR')
    with tra.IR as i:
        i.rate = 0.06
        i.cycle = 1#개월
        i.rate_cycle = i.rate / 12 * i.cycle
        
    tra.fee = tra.subacc('fee')
    with tra.fee as f:
        f.rate = 0.02
        
#Loan trb. 설정
loan.trb = loan.subacc('trb')
with loan.trb as trb:
    trb.rank = 1
    trb.is_wtdrbl = False
    trb.is_repaid = False
    
    trb.ntnl = trb.subacc('ntnl')
    with trb.ntnl as n:
        n.amt = 5_000#백만원
        n.intlamt = 5_000#백만원
        n.subscd(idx.loan[0], n.amt)
        n.addscd(idx.loan[-1], n.amt)
        
    trb.IR = trb.subacc('IR')
    with trb.IR as i:
        i.rate = 0.09
        i.cycle = 1#개월
        i.rate_cycle = i.rate / 12 * i.cycle
        
    trb.fee = trb.subacc('fee')
    with trb.fee as f:
        f.rate = 0.06
        

##Loan 조달 함수
#1)하위 loan 추출 함수
@Setattr(loan)
def getloan(loan, reverse=False):
    lst = list(loan.dct.values())
    fn = lambda x: x.rank
    lst.sort(key = fn, reverse = reverse)
    for ln in lst:
        yield ln

#2)금융비용 추정 함수
@Setattr(loan.dct)
def estimate_fee_amt(ln, idxno):
    if idxno == idx.loan[0]:
        feeamt = ln.ntnl.amt * ln.fee.rate
        ln.fee.addscd(idxno, feeamt, note=f"수수료({ln.name})")
        return feeamt
    return 0

@Setattr(loan.dct)
def estimate_IR_amt(ln, idxno):
    if ln.is_wtdrbl is False:
        return 0
    if ln.is_repaid is True:
        return 0
    ntnlbal = -ln.ntnl.bal_strt[idxno]
    IRamt = ntnlbal * ln.IR.rate_cycle
    if IRamt > 0.0:
        ln.IR.addscd(idxno, IRamt, note=f"이자({ln.name})")
        return IRamt
    return 0

#3)대출원금 인출
@Setattr(loan.dct)
def set_loan_withdrawable(ln, idxno):
    if idxno == idx.loan[0]:
        ln.is_wtdrbl = True
        
@Setattr(loan.dct)
def withdraw_ntnl_fixed(ln, acc, idxno):
    if idxno != idx.loan[0]:
        return 0
    amt_wtdrw = ln.ntnl.intlamt
    ln.ntnl.send(idxno, amt_wtdrw, acc, note=f"일시대출금({ln.name})")
    return amt_wtdrw

@Setattr(loan.dct)
def withdraw_ntnl_flexible(ln, acctmp, acc, idxno):
    if idxno < idx.loan[0]:
        return 0
    if ln.is_wtdrbl is False:
        return 0
    if ln.is_repaid is True:
        return 0
    amttopay = acctmp.scd_out[idxno] - acctmp.bal_end[idxno]
    amttopay = max(round_up(amttopay, -2), 0)
    amtscdout = ln.ntnl.rsdl_out_cum[idxno]
    
    amt_wtdrw = min(amttopay, amtscdout)
    ln.ntnl.send(idxno, amt_wtdrw, acc, note=f"한도대출금({ln.name})")
    return amt_wtdrw
    
#4)금융비용 지출
@Setattr(loan.dct)
def pay_fee_amt(ln, acc, idxno):
    feeamt = ln.fee.scd_in[idxno]
    acc.send(idxno, feeamt, ln.fee, note=f"수수료({ln.name})")
    return feeamt

@Setattr(loan.dct)
def pay_IR_amt(ln, acc, idxno):
    IRamt = ln.IR.scd_in[idxno]
    acc.send(idxno, IRamt, ln.IR, note=f"이자({ln.name})")
    return IRamt

#5)대출원금 상환
@Setattr(loan.dct)
def repay_ntnl_amt(ln, acc, idxno):
    #at maturity
    if idxno >= idx.loan[-1]:
        amt_scd_in = ln.ntnl.rsdl_in_cum[idxno] - ln.ntnl.rsdl_out_cum[idxno]
        acc.send(idxno, amt_scd_in, ln.ntnl, note=f"대출금상환({ln.name})")
        return amt_scd_in
    return 0

@Setattr(loan.dct)
def setback_loan_unwithdrawable(ln, idxno):
    #at maturity
    if idxno >= idx.loan[-1]:
        ln.is_wtdrbl = False