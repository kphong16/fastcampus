import pandas as pd
from pandas import DataFrame, Series
import cafle
from cafle import Index, Account
from cafle import Setattr
from functools import partial

from practice.astn0_overview import overview, idx


##Equity Account 설정
equity = Account(idx)

equity.보통주 = equity.subacc('보통주')
with equity.보통주 as e:
    e.amt = 5_000#백만원
    e.rate = 1.0#지분율 100%
    e.subscd(idx[0], e.amt, note="보통주 출자금")
    
    
##Equity Account 함수 설정
@Setattr(equity)
def withdraw_equity_amount(equity, idxno, oprtg):
    amt_wtdrw = 0
    for key, item in equity.dct.items():
        _df = item.jnlscd.loc[item.jnlscd.index == idxno]
        for index, row in _df.iterrows():
            item.send(idxno, row.amt_out, oprtg, "자기자본 납입")
            amt_wtdrw += row.amt_out
    return amt_wtdrw
#equity.withdraw_equity_amount = partial(withdraw_equity_amount, equity)

@Setattr(equity)
def repay_equity_amt(equity, idxno, oprtg):
    if idxno < idx[-1]:
        return 0
    amt_bal = oprtg.bal_end[idxno]
    for key, item in equity.dct.items():
        amt_rpy = item.rate * amt_bal
        oprtg.send(idxno, amt_rpy, item, note=f"자기자본상환({item.name})")
    return amt_bal