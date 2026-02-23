import pyreadstat
import pandas as pd
import math

df, meta = pyreadstat.read_sav(r'D:\SPSS_PS\Employee data.sav')
df['gender'] = df['gender'].map({'m': 'Male', 'f': 'Female'})

pivot = df.pivot_table(
    values='salbegin',
    index='educ',
    columns='gender',
    aggfunc=['mean', 'count'],
    margins=True,
    margins_name='Total'
)

print('Beginning Salary by Gender and Education')
print('=' * 85)
header = '{:<12} {:>14} {:>10} {:>14} {:>10} {:>14} {:>10}'.format(
    'Education', 'Mean Female', 'N Female', 'Mean Male', 'N Male', 'Mean Total', 'N Total')
print(header)
print('-' * 85)

for educ_val in pivot.index:
    label = str(int(educ_val)) + ' yrs' if educ_val != 'Total' else 'Total'

    def get(col_tuple):
        return pivot.loc[educ_val, col_tuple] if col_tuple in pivot.columns else float('nan')

    mf = get(('mean', 'Female'))
    nf = get(('count', 'Female'))
    mm = get(('mean', 'Male'))
    nm = get(('count', 'Male'))
    mt = get(('mean', 'Total'))
    nt = get(('count', 'Total'))

    def fmt_money(v):
        return ('${:,.0f}'.format(v)) if not math.isnan(v) else '-'

    def fmt_n(v):
        return str(int(v)) if not math.isnan(v) else '-'

    print('{:<12} {:>14} {:>10} {:>14} {:>10} {:>14} {:>10}'.format(
        label, fmt_money(mf), fmt_n(nf), fmt_money(mm), fmt_n(nm), fmt_money(mt), fmt_n(nt)))
