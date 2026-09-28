"""
Workbook builder — DDM course project, Group 02 Section A.

STANDALONE. Recomputes the entire analysis from the three raw GST registers and
writes the 16-sheet self-documenting workbook. Run either directly:

    python build_workbook.py

or via analysis.py, which runs the analysis with full console verification first.

Requires: pandas, numpy, openpyxl. Raw files must sit in the same folder.
"""
import pandas as pd, numpy as np, os
from itertools import combinations
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

import os
_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
def _find_data_dir(base):
    """Locate the folder holding the three raw registers: ./ then ../Data."""
    import os
    probe = "SALE_DATA_2324.xlsx"
    for cand in (base, os.path.join(base, "Data"),
                 os.path.join(os.path.dirname(base), "Data")):
        if os.path.exists(os.path.join(cand, probe)):
            return cand
    return base
SRC = _find_data_dir(_SCRIPT_DIR)
OUT = os.path.join(_SCRIPT_DIR, 'LG_Kent_Marketing_Analytics_FY2324_FY2526.xlsx')

# ---------------- LOAD & PREPROCESS ----------------
f2 = pd.read_excel(f'{SRC}/SALE_DATA_2324.xlsx', sheet_name='Sheet1'); f2['FISCAL_YEAR']='FY2023-24'
f3 = pd.read_excel(f'{SRC}/SALE_DATA_2425_NEW.xlsx', sheet_name='Sheet1'); f3['FISCAL_YEAR']='FY2024-25'
f1 = pd.read_excel(f'{SRC}/SALE_DATA.xlsx', sheet_name='Sheet1'); f1['FISCAL_YEAR']='FY2025-26'
raw_rows = len(f2)+len(f3)+len(f1)

common = sorted(set(f2.columns) & set(f3.columns) & set(f1.columns))
d = pd.concat([f2,f3,f1], ignore_index=True)
dups = int(d.duplicated().sum())
d = d.drop_duplicates().reset_index(drop=True)

d['INVOICE_DATE'] = pd.to_datetime(d['INVOICEDATE'], errors='coerce', dayfirst=True)
d['MONTH'] = d['INVOICE_DATE'].dt.month
d['FISCAL_MONTH'] = d['MONTH'].apply(lambda m: (m-4)%12+1)
FM_NAMES = ['Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec','Jan','Feb','Mar']
d['FISCAL_MONTH_NAME'] = d['FISCAL_MONTH'].apply(lambda x: FM_NAMES[int(x)-1])

HSN = {841510:'Air Conditioner',852872:'Television',841810:'Refrigerator',
       84501100:'Washing Machine Automatic',84501200:'Washing Machine Semi-Automatic',
       842121:'Water Purifier RO'}
d['CATEGORY'] = d['HSNCODE'].map(HSN).fillna('Other')
d['LINE_REVENUE'] = (d['TOTAL'].fillna(0)+d['SGSTAMT'].fillna(0)+
                     d['CGSTAMT'].fillna(0)+d['IGSTAMT'].fillna(0))
d['ZERO_VALUE_FLAG'] = (d['GSTRATE'].fillna(0)==0)
d['BAD_SKU_FLAG'] = d['PRODUCT'].astype(str).str.fullmatch(r'\d+').fillna(False)
d['CASH_CUSTOMER_FLAG'] = d['CUSTOMERNAME'].astype(str).str.contains('CASH CUSTOMER',case=False,na=False)
d['BRANCH'] = d['INVOICENO'].astype(str).str.extract(r'^([A-Z]+)/')[0]

clean_rows = len(d)
inv = d.drop_duplicates('INVOICENO')
n_inv = inv['INVOICENO'].nunique()
tot_rev = inv['GRANDAMOUNT'].sum()
named = d[~d['CASH_CUSTOMER_FLAG']]
n_named = named['CUSTOMERNAME'].nunique()

FYS = ['FY2023-24','FY2024-25','FY2025-26']
CATS = ['Air Conditioner','Television','Refrigerator','Other',
        'Washing Machine Automatic','Washing Machine Semi-Automatic','Water Purifier RO']

# ---------------- ANALYSIS A ----------------
catfy = d.pivot_table(index='CATEGORY',columns='FISCAL_YEAR',values='LINE_REVENUE',aggfunc='sum').reindex(CATS).fillna(0)
ll = d[d['FISCAL_MONTH']<=10]
catll = ll.pivot_table(index='CATEGORY',columns='FISCAL_YEAR',values='LINE_REVENUE',aggfunc='sum').reindex(CATS).fillna(0)
month_rev = inv.pivot_table(index='FISCAL_MONTH',columns='FISCAL_YEAR',values='GRANDAMOUNT',aggfunc='sum').fillna(0)
month_cnt = inv.pivot_table(index='FISCAL_MONTH',columns='FISCAL_YEAR',values='INVOICENO',aggfunc='count').fillna(0)
cat_month = d.pivot_table(index='CATEGORY',columns='FISCAL_MONTH',values='LINE_REVENUE',aggfunc='sum').reindex(CATS).fillna(0)

# ---------------- ANALYSIS B ----------------
def basket_pairs(series_of_sets, label):
    N=len(series_of_sets); rows=[]
    for a,b in combinations(CATS,2):
        na=sum(1 for s in series_of_sets if a in s); nb=sum(1 for s in series_of_sets if b in s)
        both=sum(1 for s in series_of_sets if a in s and b in s)
        if both==0 or na==0 or nb==0: continue
        sup=both/N
        rows.append([label,a,b,N,both,sup,both/na,both/nb,sup/((na/N)*(nb/N))])
    return rows

inv_sets = d.groupby('INVOICENO')['CATEGORY'].apply(set)
cust_sets = named.groupby('CUSTOMERNAME')['CATEGORY'].apply(set)
pairs = basket_pairs(inv_sets,'Invoice basket') + basket_pairs(cust_sets,'Customer basket')
pairs_df = pd.DataFrame(pairs,columns=['BASKET_TYPE','CATEGORY_A','CATEGORY_B','BASKETS','PAIR_BASKETS',
                                       'SUPPORT','CONFIDENCE_A_TO_B','CONFIDENCE_B_TO_A','LIFT'])
pairs_df = pairs_df.sort_values(['BASKET_TYPE','LIFT'],ascending=[True,False])

sm=[]
for fy in FYS:
    s=named[named['FISCAL_YEAR']==fy]
    nc=s.groupby('CUSTOMERNAME')['CATEGORY'].nunique()
    rv=s.drop_duplicates('INVOICENO').groupby('CUSTOMERNAME')['GRANDAMOUNT'].sum()
    for lab,idx in [('Multi-category',nc[nc>=2].index),('Single-category',nc[nc==1].index)]:
        sm.append([fy,lab,len(idx),rv[idx].mean(),rv[idx].sum()])
sm_df=pd.DataFrame(sm,columns=['FISCAL_YEAR','CUSTOMER_TYPE','CUSTOMERS','AVG_TOTAL_SPEND','TOTAL_REVENUE'])

# ---------------- ANALYSIS C ----------------
pc = d[d['BASICRATE']>0].copy()
pc['PROD_MEAN'] = pc.groupby('PRODUCT')['BASICRATE'].transform('mean')
pc['PRICE_INDEX'] = pc['BASICRATE']/pc['PROD_MEAN']
pidx = pc.pivot_table(index='CATEGORY',columns='FISCAL_MONTH',values='PRICE_INDEX',aggfunc='mean').reindex(CATS)
cov = (pidx.std(axis=1)/pidx.mean(axis=1)).rename('COEFFICIENT_OF_VARIATION').reset_index()
cov = cov.sort_values('COEFFICIENT_OF_VARIATION',ascending=False)

# ---------------- ANALYSIS D ----------------
ninv = named.drop_duplicates('INVOICENO')[['INVOICENO','CUSTOMERNAME','INVOICE_DATE','GRANDAMOUNT']]
snap = ninv['INVOICE_DATE'].max()
rfm = ninv.groupby('CUSTOMERNAME').agg(RECENCY_DAYS=('INVOICE_DATE',lambda x:(snap-x.max()).days),
                                       FREQUENCY=('INVOICENO','nunique'),
                                       MONETARY=('GRANDAMOUNT','sum'))
rfm['R_SCORE']=pd.qcut(rfm['RECENCY_DAYS'].rank(method='first'),5,labels=[5,4,3,2,1]).astype('float').fillna(3).astype(int)
rfm['F_SCORE']=pd.cut(rfm['FREQUENCY'],bins=[0,1,2,3,5,10**9],labels=[1,2,3,4,5]).astype('float').fillna(1).astype(int)
rfm['M_SCORE']=pd.qcut(rfm['MONETARY'].rank(method='first'),5,labels=[1,2,3,4,5]).astype('float').fillna(3).astype(int)
rfm['RFM_SCORE']=rfm['R_SCORE'].astype(str)+rfm['F_SCORE'].astype(str)+rfm['M_SCORE'].astype(str)
def seg(r):
    R,F,M=r['R_SCORE'],r['F_SCORE'],r['M_SCORE']
    if F>=4 and M>=4: return 'Champions'
    if F>=3 and M>=3: return 'Loyal'
    if R>=4 and F<=2: return 'Potential Loyalist'
    if R<=2 and M>=4: return 'At Risk (high value)'
    if R<=2: return 'Lapsed / Not Yet Due'
    return 'Needs Attention'
rfm['SEGMENT']=rfm.apply(seg,axis=1)
segsum = rfm.groupby('SEGMENT').agg(CUSTOMERS=('MONETARY','size'),REVENUE=('MONETARY','sum'),
        AVG_MONETARY=('MONETARY','mean'),AVG_FREQUENCY=('FREQUENCY','mean'),
        AVG_RECENCY_DAYS=('RECENCY_DAYS','mean')).reset_index()
segsum['REVENUE_SHARE']=segsum['REVENUE']/segsum['REVENUE'].sum()
segsum=segsum.sort_values('REVENUE',ascending=False)

srt=rfm['MONETARY'].sort_values(ascending=False); cum=srt.cumsum()/srt.sum()
top10=cum.iloc[int(len(srt)*.1)-1]; top20=cum.iloc[int(len(srt)*.2)-1]
once=int((rfm['FREQUENCY']==1).sum())

print('prep done', clean_rows, n_inv, n_named)

# ================= WRITE WORKBOOK =================
wb = openpyxl.Workbook(); wb.remove(wb.active)
F='Arial'
H_FILL=PatternFill('solid',fgColor='1F3864'); H_FONT=Font(name=F,bold=True,color='FFFFFF',size=10)
T_FONT=Font(name=F,bold=True,size=14,color='1F3864'); S_FONT=Font(name=F,bold=True,size=11,color='1F3864')
BODY=Font(name=F,size=10); BOLD=Font(name=F,size=10,bold=True)
TOT_FILL=PatternFill('solid',fgColor='DDEBF7')
THIN=Side(style='thin',color='BFBFBF'); BORD=Border(left=THIN,right=THIN,top=THIN,bottom=THIN)
WRAP=Alignment(wrap_text=True,vertical='top')

def sheet(name): 
    ws=wb.create_sheet(name); ws.sheet_view.showGridLines=False; return ws
def title(ws,r,txt,font=T_FONT):
    ws.cell(r,1,txt).font=font; return r+1
def note(ws,r,txt,width=9):
    c=ws.cell(r,1,txt); c.font=Font(name=F,size=10,italic=True,color='444444'); c.alignment=WRAP
    ws.merge_cells(start_row=r,start_column=1,end_row=r,end_column=width)
    ws.row_dimensions[r].height=max(15,15*(len(txt)//110+1)); return r+1
def header(ws,r,cols):
    for i,h in enumerate(cols,1):
        c=ws.cell(r,i,h); c.font=H_FONT; c.fill=H_FILL; c.border=BORD
        c.alignment=Alignment(wrap_text=True,vertical='center',horizontal='center')
    ws.row_dimensions[r].height=30; return r+1
def widths(ws,ws_widths):
    for i,w in enumerate(ws_widths,1): ws.column_dimensions[get_column_letter(i)].width=w

NUM='#,##0'; PCT='0.0%'; DEC4='0.0000'

# ---- README ----
ws=sheet('README'); widths(ws,[110])
r=title(ws,1,'LG / KENT CONSUMER DURABLES DEALERSHIP — MARKETING ANALYTICS WORKBOOK')
r=note(ws,r+1,'Three fiscal years of GST sales-register data (FY2023-24, FY2024-25, FY2025-26) from a multi-branch consumer durables dealership in Varanasi, Uttar Pradesh, trading in LG white goods and Kent water purification products.')
r=title(ws,r+1,'Business problem addressed',S_FONT)
for t in ['On a like-for-like basis the business has declined 7.9% over two years while serving MORE customers.',
          'Six of seven categories carry 10-15 year replacement cycles and are structurally depleting.',
          'Air Conditioner is the only growing category (+16.8%) but is negatively associated with every other category, so it generates single-category customers worth ~1/3 of multi-category ones.']:
    r=note(ws,r,'• '+t)
r=title(ws,r+1,'Contents of this workbook',S_FONT)
contents=[('README','This sheet — business context, contents guide.'),
 ('DATA_DICTIONARY','Every column defined; original vs derived.'),
 ('PREPROCESSING_LOG','Every cleaning step in order, rows affected, reason.'),
 ('DATA_GAP_CHECK','CRITICAL: FY2024-25 is incomplete in Feb and Mar. Read before any YoY comparison.'),
 ('ANALYSIS_A_CATEGORY','Category revenue by fiscal year, full-year reported basis, with reconciliation.'),
 ('ANALYSIS_A2_LIKE_FOR_LIKE','Category revenue on the comparable Apr-Jan window. THIS IS THE RELIABLE VIEW.'),
 ('ANALYSIS_A_SEASONALITY','Monthly revenue and invoice counts; category revenue by month.'),
 ('ANALYSIS_B_BASKET','Market basket: support, confidence, lift for every category pair.'),
 ('ANALYSIS_B_CROSSSELL','Single vs multi-category customers and their value gap.'),
 ('ANALYSIS_C_PRICING','Product-normalised monthly price index and coefficient of variation.'),
 ('ANALYSIS_D_RFM','RFM segmentation, segment sizes and revenue concentration.'),
 ('KEY_FINDINGS','Numbered findings with supporting numbers.'),
 ('RECOMMENDATIONS','Seven prioritised recommendations with rationale and risk.'),
 ('FURTHER_SCOPE','What else this data supports, and what needs data the firm lacks.'),
 ('LIMITATIONS','What this dataset cannot support. Read before drawing conclusions.'),
 ('CLEANED_MASTER_DATA','Full cleaned line-item dataset with all derived columns.'),
 ('INVOICE_LEVEL','One row per invoice.')]
r=header(ws,r,['Sheet','What it contains'])
ws.column_dimensions['A'].width=32; ws.column_dimensions['B'].width=95
for a,b in contents:
    ws.cell(r,1,a).font=BOLD; ws.cell(r,1).border=BORD
    c=ws.cell(r,2,b); c.font=BODY; c.border=BORD; c.alignment=WRAP; r+=1
r=title(ws,r+1,'Key verified numbers',S_FONT)
for k,v,f in [('Source line-item rows',raw_rows,NUM),('Exact duplicate rows removed',dups,NUM),
              ('Cleaned master rows',clean_rows,NUM),('Distinct invoices',n_inv,NUM),
              ('Total billed revenue (Rs)',tot_rev,NUM),('Named customers (cash excluded)',n_named,NUM)]:
    ws.cell(r,1,k).font=BODY; c=ws.cell(r,2,v); c.font=BOLD; c.number_format=f; r+=1

# ---- DATA DICTIONARY ----
ws=sheet('DATA_DICTIONARY'); widths(ws,[30,16,80])
r=title(ws,1,'DATA DICTIONARY'); r=header(ws,r+1,['Column','Origin','Definition'])
dd=[('INVOICENO','Original','Invoice number. Repeats across the lines of one invoice.'),
('INVOICEDATE','Original','Invoice date as supplied (day-first text/date).'),
('CUSTOMERNAME','Original','Buyer name. The only customer identifier available; no unique key exists.'),
('PRODUCT','Original','SKU description. A few rows contain purely numeric codes (see BAD_SKU_FLAG).'),
('HSNCODE','Original','GST HSN classification code. Basis of the CATEGORY mapping.'),
('QUANTITY','Original','Units on the line.'),
('BASICRATE','Original','Per-unit rate before tax. Negotiated price, not list price.'),
('TOTAL','Original','Line taxable value, excluding GST.'),
('CGSTAMT / SGSTAMT / IGSTAMT','Original','GST amounts on the line.'),
('GRANDAMOUNT','Original','INVOICE-level total including tax. Constant across all lines of an invoice. Must NOT be used to split revenue by category.'),
('GSTRATE','Original','GST rate on the line. Zero indicates bundled/warranty items.'),
('FISCAL_YEAR','Derived','Source file tag: FY2023-24, FY2024-25 or FY2025-26.'),
('INVOICE_DATE','Derived','INVOICEDATE parsed day-first to a true date.'),
('MONTH','Derived','Calendar month number.'),
('FISCAL_MONTH','Derived','Fiscal month index, April = 1 through March = 12.'),
('FISCAL_MONTH_NAME','Derived','Month name for readability.'),
('CATEGORY','Derived','Product category mapped from HSNCODE. Top codes cover 96.6% of revenue; all else = Other.'),
('LINE_REVENUE','Derived','TOTAL + CGSTAMT + SGSTAMT + IGSTAMT. The CORRECT basis for category-level revenue. Reconciles to ~100.1% of distinct-invoice GRANDAMOUNT.'),
('ZERO_VALUE_FLAG','Derived','TRUE where GSTRATE = 0 (bundled or warranty items).'),
('BAD_SKU_FLAG','Derived','TRUE where PRODUCT is a purely numeric string (mis-entered master data).'),
('CASH_CUSTOMER_FLAG','Derived','TRUE where the buyer is an anonymous walk-in. EXCLUDED from all customer-level analysis.'),
('BRANCH','Derived','Alphabetic prefix of INVOICENO, identifying the billing branch/series.')]
for a,b,c_ in dd:
    ws.cell(r,1,a).font=BOLD; ws.cell(r,2,b).font=BODY
    cc=ws.cell(r,3,c_); cc.font=BODY; cc.alignment=WRAP
    for i in range(1,4): ws.cell(r,i).border=BORD
    r+=1

# ---- PREPROCESSING LOG ----
ws=sheet('PREPROCESSING_LOG'); widths(ws,[8,46,16,80])
r=title(ws,1,'PREPROCESSING LOG'); r=header(ws,r+1,['Step','Action','Rows affected','Reason'])
log=[(1,'Unioned three fiscal-year files',f'{raw_rows:,} in',f'Joined on the {len(common)} common columns. FY2025-26 uniquely has GSTIN/SALESLEDGER; the earlier two uniquely have STATE/COUNTRY/BUYERMAILINGNAME/ADDRESS2/CASHBACK — retained as optional, not used as join keys.'),
(2,'Tagged FISCAL_YEAR','All','Identify source year for every row.'),
(3,'Parsed INVOICEDATE day-first','All','Source dates are day-first; derived MONTH, FISCAL_MONTH (Apr=1..Mar=12), FISCAL_MONTH_NAME.'),
(4,'Removed exact duplicate rows',f'{dups} removed','Byte-identical repeated lines. Cleaned master = {:,} rows.'.format(clean_rows)),
(5,'Mapped HSNCODE to CATEGORY','All','Top HSN codes cover 96.6% of revenue; all remaining codes grouped as Other.'),
(6,'Computed LINE_REVENUE','All','CORRECTION APPLIED: category is a LINE-level attribute, so category revenue must use line-level value (TOTAL + all GST), never invoice-level GRANDAMOUNT. Attributing GRANDAMOUNT to each category on an invoice double-counts every multi-line invoice (40-43% of invoices) and roughly doubles category totals.'),
(7,'Flagged zero-value lines',f'{int(d.ZERO_VALUE_FLAG.sum()):,} flagged','GSTRATE = 0. Bundled/warranty items. Flagged, not deleted.'),
(8,'Flagged bad SKU codes',f'{int(d.BAD_SKU_FLAG.sum()):,} flagged','PRODUCT is purely numeric — mis-entered master data. Flagged, not deleted.'),
(9,'Flagged cash customers',f'{int(d.CASH_CUSTOMER_FLAG.sum()):,} flagged','Anonymous walk-ins with no recoverable identity. EXCLUDED from all customer-level analysis.'),
(10,'Derived BRANCH','All','Alphabetic prefix of INVOICENO.'),
(11,'Built invoice-level table',f'{n_inv:,} invoices','GRANDAMOUNT verified constant within each invoice, so it is safe as an invoice-level total.'),
(12,'Identified FY2024-25 data gap','See DATA_GAP_CHECK','Feb and Mar FY2024-25 are materially incomplete. Like-for-like Apr-Jan analysis added as a result.')]
for a,b,c_,e in log:
    ws.cell(r,1,a).font=BOLD; ws.cell(r,2,b).font=BODY; ws.cell(r,3,c_).font=BODY
    cc=ws.cell(r,4,e); cc.font=BODY; cc.alignment=WRAP
    for i in range(1,5): ws.cell(r,i).border=BORD; ws.cell(r,i).alignment=WRAP
    ws.row_dimensions[r].height=max(15,14*(len(e)//95+1)); r+=1

# ---- DATA GAP CHECK ----
ws=sheet('DATA_GAP_CHECK'); widths(ws,[16,18,18,18,14])
r=title(ws,1,'DATA COMPLETENESS CHECK — CRITICAL, READ FIRST')
r=note(ws,r+1,'February FY2024-25 contains only 298 invoices and March FY2024-25 only 35, against roughly 800 and 1,100 in the same months of the two adjacent years. A functioning dealership does not fall from 1,054 invoices in November to 35 in March and then recover to 1,249 in the same month the following year. This is almost certainly a TRUNCATED EXPORT, not a demand collapse. Every full-year comparison involving FY2024-25 is therefore understated, and the headline V-shape (-13.7% then +9.3%) is largely an artefact. Use ANALYSIS_A2_LIKE_FOR_LIKE for reliable trend readings.')
r=title(ws,r+1,'Table — Invoice count by fiscal month',S_FONT)
r=header(ws,r,['Fiscal month']+FYS)
first=r
for i in range(1,13):
    ws.cell(r,1,FM_NAMES[i-1]).font=BODY
    for j,fy in enumerate(FYS,2):
        v=int(month_cnt.loc[i,fy]) if i in month_cnt.index and fy in month_cnt.columns else 0
        c=ws.cell(r,j,v); c.font=BODY; c.number_format=NUM
    for k in range(1,5): ws.cell(r,k).border=BORD
    r+=1
ws.cell(r,1,'TOTAL').font=BOLD
for j in range(2,5):
    c=ws.cell(r,j,f'=SUM({get_column_letter(j)}{first}:{get_column_letter(j)}{r-1})')
    c.font=BOLD; c.number_format=NUM; c.fill=TOT_FILL
r+=2
r=title(ws,r,'Table — February + March revenue comparison (Rs)',S_FONT)
r=header(ws,r,['Fiscal year','Feb + Mar revenue','Feb + Mar invoices','Assessment'])
fm_rev={fy: float(month_rev.loc[[11,12],fy].sum()) for fy in FYS}
fm_cnt={fy: int(month_cnt.loc[[11,12],fy].sum()) for fy in FYS}
for fy in FYS:
    ws.cell(r,1,fy).font=BODY
    c=ws.cell(r,2,fm_rev[fy]); c.font=BODY; c.number_format=NUM
    c=ws.cell(r,3,fm_cnt[fy]); c.font=BODY; c.number_format=NUM
    ws.cell(r,4,'INCOMPLETE' if fy=='FY2024-25' else 'Normal').font=BOLD if fy=='FY2024-25' else BODY
    for k in range(1,5): ws.cell(r,k).border=BORD
    r+=1
short=(fm_rev['FY2023-24']+fm_rev['FY2025-26'])/2-fm_rev['FY2024-25']
ws.cell(r,1,'Estimated shortfall (Rs)').font=BOLD
c=ws.cell(r,2,short); c.font=BOLD; c.number_format=NUM; c.fill=TOT_FILL

# ---- ANALYSIS A ----
def cat_table(ws,r,tab,label):
    r=title(ws,r,label,S_FONT)
    r=header(ws,r,['Category']+FYS+['YoY FY24-25','YoY FY25-26','2-yr change'])
    first=r
    for cat in CATS:
        ws.cell(r,1,cat).font=BODY
        for j,fy in enumerate(FYS,2):
            c=ws.cell(r,j,float(tab.loc[cat,fy])); c.font=BODY; c.number_format=NUM
        ws.cell(r,5,f'=IF(B{r}=0,"",C{r}/B{r}-1)').number_format=PCT
        ws.cell(r,6,f'=IF(C{r}=0,"",D{r}/C{r}-1)').number_format=PCT
        ws.cell(r,7,f'=IF(B{r}=0,"",D{r}/B{r}-1)').number_format=PCT
        for k in range(1,8): ws.cell(r,k).border=BORD
        r+=1
    ws.cell(r,1,'TOTAL').font=BOLD
    for j in range(2,5):
        c=ws.cell(r,j,f'=SUM({get_column_letter(j)}{first}:{get_column_letter(j)}{r-1})')
        c.font=BOLD; c.number_format=NUM; c.fill=TOT_FILL
    ws.cell(r,5,f'=C{r}/B{r}-1').number_format=PCT; ws.cell(r,5).font=BOLD; ws.cell(r,5).fill=TOT_FILL
    ws.cell(r,6,f'=D{r}/C{r}-1').number_format=PCT; ws.cell(r,6).font=BOLD; ws.cell(r,6).fill=TOT_FILL
    ws.cell(r,7,f'=D{r}/B{r}-1').number_format=PCT; ws.cell(r,7).font=BOLD; ws.cell(r,7).fill=TOT_FILL
    return r+1, r

ws=sheet('ANALYSIS_A_CATEGORY'); widths(ws,[34,18,18,18,14,14,14])
r=title(ws,1,'ANALYSIS A — CATEGORY REVENUE BY FISCAL YEAR (FULL-YEAR REPORTED BASIS)')
r=note(ws,r+1,'METHOD: revenue summed at LINE level using LINE_REVENUE = TOTAL + CGST + SGST + IGST. Category is a line-level attribute, so invoice-level GRANDAMOUNT cannot be used here — doing so double-counts every multi-line invoice. WARNING: the FY2024-25 column is understated because that file is incomplete in Feb and Mar (see DATA_GAP_CHECK). Use ANALYSIS_A2_LIKE_FOR_LIKE for reliable trends.')
r,totrow = cat_table(ws,r+1,catfy,'Table A1 — Category revenue by fiscal year (Rs, LINE_REVENUE)')
r=title(ws,r+1,'Reconciliation check — line revenue vs distinct-invoice totals',S_FONT)
r=header(ws,r,['Fiscal year','Sum of LINE_REVENUE','Distinct-invoice GRANDAMOUNT','Ratio'])
for fy in FYS:
    ws.cell(r,1,fy).font=BODY
    c=ws.cell(r,2,float(catfy[fy].sum())); c.font=BODY; c.number_format=NUM
    c=ws.cell(r,3,float(inv[inv['FISCAL_YEAR']==fy]['GRANDAMOUNT'].sum())); c.font=BODY; c.number_format=NUM
    c=ws.cell(r,4,f'=B{r}/C{r}'); c.font=BOLD; c.number_format='0.0000'
    for k in range(1,5): ws.cell(r,k).border=BORD
    r+=1
r=note(ws,r+1,'A ratio of ~1.001 confirms the decomposition is sound: category revenues sum to the year total. Under the earlier erroneous method the parts exceeded the whole by roughly 2x.')

ws=sheet('ANALYSIS_A2_LIKE_FOR_LIKE'); widths(ws,[34,18,18,18,14,14,14])
r=title(ws,1,'ANALYSIS A2 — LIKE-FOR-LIKE CATEGORY REVENUE (APRIL-JANUARY WINDOW)')
r=note(ws,r+1,'METHOD: identical to Analysis A but restricted to fiscal months 1-10 (April to January), the window in which all three years are comparably complete. This removes the FY2024-25 Feb-Mar data gap from the comparison. THIS IS THE RELIABLE VIEW OF THE TREND.')
r,_ = cat_table(ws,r+1,catll,'Table A2 — Category revenue, April-January window (Rs, LINE_REVENUE)')
r=title(ws,r+1,'Interpretation',S_FONT)
for t in ['There was NO collapse in FY2024-25 — on a comparable basis it was essentially flat at -1.3%. The apparent -13.7% was a reporting artefact.',
          'There was NO recovery in FY2025-26 either — the business declined 6.6%, and is down 7.9% across two years. The V-shape is an illusion; the trend is steady erosion.',
          'Air Conditioner is the ONLY category above its FY2023-24 level (+16.8%). Every other category is materially below: Television -24.2%, Washing Machine Automatic -32.0%, Water Purifier RO -20.4%, Refrigerator -16.0%.',
          'This is the pattern the replacement-cycle hypothesis predicts: the short-cycle, climate-driven category grows while 10-15 year cycle categories deplete.']:
    r=note(ws,r,'• '+t)

# ---- SEASONALITY ----
ws=sheet('ANALYSIS_A_SEASONALITY'); widths(ws,[16,18,18,18,14,14,14,14])
r=title(ws,1,'ANALYSIS A — SEASONALITY')
r=note(ws,r+1,'METHOD: monthly totals use distinct-invoice GRANDAMOUNT, which is correct because it is a whole-invoice measure. Category-by-month uses LINE_REVENUE.')
r=title(ws,r+1,'Table A3 — Total invoice revenue by fiscal month (Rs)',S_FONT)
r=header(ws,r,['Fiscal month']+FYS+['Avg share of year'])
first=r
for i in range(1,13):
    ws.cell(r,1,FM_NAMES[i-1]).font=BODY
    for j,fy in enumerate(FYS,2):
        c=ws.cell(r,j,float(month_rev.loc[i,fy]) if i in month_rev.index else 0); c.font=BODY; c.number_format=NUM
    ws.cell(r,5,f'=AVERAGE(B{r}/B${first+12},C{r}/C${first+12},D{r}/D${first+12})').number_format=PCT
    for k in range(1,6): ws.cell(r,k).border=BORD
    r+=1
ws.cell(r,1,'TOTAL').font=BOLD
for j in range(2,5):
    c=ws.cell(r,j,f'=SUM({get_column_letter(j)}{first}:{get_column_letter(j)}{r-1})'); c.font=BOLD; c.number_format=NUM; c.fill=TOT_FILL
r+=2
r=note(ws,r,'April-June contributes 39-49% of annual revenue; the July-September monsoon trough is a remarkably stable 12.0-12.8% in every year. The stability of the trough indicates the seasonal shape is structural, not a feature of any one year. February and March FY2024-25 reflect the data gap, not demand.')
r=title(ws,r+1,'Table A4 — Category revenue by fiscal month, pooled across all years (Rs)',S_FONT)
r=header(ws,r,['Category']+FM_NAMES)
widths(ws,[34]+[13]*12)
for cat in CATS:
    ws.cell(r,1,cat).font=BODY
    for i in range(1,13):
        v=float(cat_month.loc[cat,i]) if i in cat_month.columns else 0
        c=ws.cell(r,i+1,v); c.font=BODY; c.number_format=NUM
    for k in range(1,14): ws.cell(r,k).border=BORD
    r+=1

# ---- ANALYSIS B ----
ws=sheet('ANALYSIS_B_BASKET'); widths(ws,[18,32,32,12,13,12,15,15,10])
r=title(ws,1,'ANALYSIS B — MARKET BASKET ANALYSIS (SUPPORT, CONFIDENCE, LIFT)')
r=note(ws,r+1,'METHOD: association analysis run at two levels. INVOICE BASKET treats each invoice as a transaction (what is bought together in one visit). CUSTOMER BASKET treats each named customer as a basket of every category they bought across three years (portfolio ownership). For durables the customer view matters because 10-15 year replacement cycles mean the same household rarely buys two categories on the same day. LIFT > 1 = positive association; LIFT < 1 = negative association.')
r=header(ws,r+1,['Basket type','Category A','Category B','Baskets','Pair baskets','Support','Confidence A→B','Confidence B→A','Lift'])
for _,x in pairs_df.iterrows():
    ws.cell(r,1,x['BASKET_TYPE']).font=BODY; ws.cell(r,2,x['CATEGORY_A']).font=BODY; ws.cell(r,3,x['CATEGORY_B']).font=BODY
    for j,(v,fmt) in enumerate([(int(x['BASKETS']),NUM),(int(x['PAIR_BASKETS']),NUM),(x['SUPPORT'],DEC4),
                                (x['CONFIDENCE_A_TO_B'],DEC4),(x['CONFIDENCE_B_TO_A'],DEC4),(x['LIFT'],'0.000')],4):
        c=ws.cell(r,j,v); c.font=BOLD if j==9 else BODY; c.number_format=fmt
    if x['LIFT']>=1.2: ws.cell(r,9).fill=PatternFill('solid',fgColor='C6EFCE')
    elif x['LIFT']<1.0: ws.cell(r,9).fill=PatternFill('solid',fgColor='FFC7CE')
    for k in range(1,10): ws.cell(r,k).border=BORD
    r+=1
r=title(ws,r+1,'WHY LIFT AND NOT CONFIDENCE ALONE',S_FONT)
r=note(ws,r,'Support, confidence and lift each answer a different question and each fails on its own. SUPPORT = P(A and B): how often the pair occurs at all, i.e. whether a rule is frequent enough to be worth acting on commercially. CONFIDENCE = P(B|A): the conditional probability of the consequent given the antecedent. LIFT = confidence / P(B): confidence corrected for the consequent\'s own base rate. Confidence alone systematically favours large categories, because a popular consequent appears in many baskets by chance, so a rule can be simultaneously HIGH-confidence and NEGATIVELY associated.')
r=header(ws,r,['Rule (antecedent -> consequent)','Support','Confidence','Lift','Reading'])
for a,b,c_,e,g in [('Refrigerator -> Other',0.0902,0.328,0.856,'Highest support AND confidence in the set, yet lift < 1. Other is a large residual category appearing by chance; the pair is NEGATIVELY associated. Acting on confidence alone promotes exactly the wrong rule.'),
                   ('Refrigerator -> Television',0.0790,0.287,1.354,'Strong on BOTH dimensions - second-highest support and a solid lift. The most reliable pairing, and the anchor for next-best-category targeting.'),
                   ('Television -> WM Semi-Automatic',0.0353,0.167,1.611,'Under half the support and roughly half the confidence of the first rule, but the STRONGEST lift - a genuine complementary relationship.'),
                   ('Air Conditioner -> Water Purifier',0.0040,0.014,0.475,'Weakest on every measure. The AC occasion closes no other sale.')]:
    ws.cell(r,1,a).font=BODY
    for j,(v,fmt) in enumerate([(b,DEC4),(c_,'0.000'),(e,'0.000')],2):
        cc=ws.cell(r,j,v); cc.font=BOLD if j==4 else BODY; cc.number_format=fmt
    cc=ws.cell(r,5,g); cc.font=BODY; cc.alignment=WRAP
    for k in range(1,6): ws.cell(r,k).border=BORD
    ws.row_dimensions[r].height=max(30,13*(len(g)//60+1)); r+=1
ws.column_dimensions['E'].width=70
r=title(ws,r+1,'Key reading',S_FONT)
for t in ['Television, Refrigerator and both Washing Machine variants form a tightly associated cluster (lift 1.28-1.61) — these are household-formation purchases evaluated together.',
          'AIR CONDITIONER HAS LIFT BELOW 1.0 AGAINST EVERY OTHER CATEGORY, at both invoice and customer level. An AC buyer is materially LESS likely than average to buy anything else.',
          'This is the central commercial problem: the only growing category is the least cross-sold. Every incremental AC customer tends to become another single-category buyer.']:
    r=note(ws,r,'• '+t)

ws=sheet('ANALYSIS_B_CROSSSELL'); widths(ws,[16,22,14,20,22])
r=title(ws,1,'ANALYSIS B — SINGLE VS MULTI-CATEGORY CUSTOMER VALUE')
r=note(ws,r+1,'METHOD: named customers only (cash customers excluded). A multi-category customer bought from 2+ categories within that fiscal year. Spend uses distinct-invoice GRANDAMOUNT.')
r=header(ws,r+1,['Fiscal year','Customer type','Customers','Avg spend (Rs)','Total revenue (Rs)'])
for _,x in sm_df.iterrows():
    ws.cell(r,1,x['FISCAL_YEAR']).font=BODY; ws.cell(r,2,x['CUSTOMER_TYPE']).font=BODY
    c=ws.cell(r,3,int(x['CUSTOMERS'])); c.font=BODY; c.number_format=NUM
    c=ws.cell(r,4,float(x['AVG_TOTAL_SPEND'])); c.font=BOLD; c.number_format=NUM
    c=ws.cell(r,5,float(x['TOTAL_REVENUE'])); c.font=BODY; c.number_format=NUM
    for k in range(1,6): ws.cell(r,k).border=BORD
    r+=1
r=title(ws,r+1,'Value gap and trend',S_FONT)
r=header(ws,r,['Fiscal year','Multi/single value ratio','Single-category share of customers'])
for fy in FYS:
    m=sm_df[(sm_df.FISCAL_YEAR==fy)&(sm_df.CUSTOMER_TYPE=='Multi-category')].iloc[0]
    s=sm_df[(sm_df.FISCAL_YEAR==fy)&(sm_df.CUSTOMER_TYPE=='Single-category')].iloc[0]
    ws.cell(r,1,fy).font=BODY
    c=ws.cell(r,2,m['AVG_TOTAL_SPEND']/s['AVG_TOTAL_SPEND']); c.font=BOLD; c.number_format='0.00"x"'
    c=ws.cell(r,3,s['CUSTOMERS']/(s['CUSTOMERS']+m['CUSTOMERS'])); c.font=BOLD; c.number_format=PCT
    for k in range(1,4): ws.cell(r,k).border=BORD
    r+=1
r=note(ws,r+1,'Multi-category customers are worth 3.4-3.7x single-category ones in every year, yet the single-category share is RISING (64.1% to 69.4%). The firm is acquiring more customers at declining value per customer.')

# ---- ANALYSIS C ----
ws=sheet('ANALYSIS_C_PRICING'); widths(ws,[34]+[11]*12+[16])
r=title(ws,1,'ANALYSIS C — SEASONAL PRICING')
r=note(ws,r+1,'METHOD: each transaction price (BASICRATE) is indexed against the mean BASICRATE of that same PRODUCT across all three years. This normalisation is essential — a raw average price by month would reflect product mix, not price movement. The index is then averaged by category and fiscal month. Rows with BASICRATE <= 0 excluded. An index of 1.000 means the product sold at its own three-year average price.')
r=title(ws,r+1,'Table C1 — Product-normalised price index by category and fiscal month',S_FONT)
r=header(ws,r,['Category']+FM_NAMES+['Coeff. of variation'])
for cat in CATS:
    if cat not in pidx.index: continue
    ws.cell(r,1,cat).font=BODY
    for i in range(1,13):
        v=pidx.loc[cat,i] if i in pidx.columns else None
        c=ws.cell(r,i+1,float(v) if pd.notna(v) else None); c.font=BODY; c.number_format='0.000'
    cvv=cov[cov.CATEGORY==cat]['COEFFICIENT_OF_VARIATION']
    c=ws.cell(r,14,float(cvv.iloc[0]) if len(cvv) else None); c.font=BOLD; c.number_format=DEC4
    for k in range(1,15): ws.cell(r,k).border=BORD
    r+=1
r=title(ws,r+1,'Finding — prices do NOT move with the season',S_FONT)
for t in ['Air Conditioner price varies by less than 1% across the year (coefficient of variation 0.0072) despite a threefold swing in demand between the April-June peak and the monsoon trough.',
          'Television — a category with NO weather seasonality — is nearly twice as price-volatile (0.0128) as Air Conditioner.',
          'The firm is therefore not practising seasonal pricing on air conditioners at all, in one of the most seasonally extreme AC markets in India. The revenue-management lever is entirely unused.',
          'CAVEAT: these are negotiated transaction prices, not experimentally controlled prices. The finding that price does not vary is robust; the inference that varying it would raise revenue is untested. Test before rolling out.']:
    r=note(ws,r,'• '+t)

# ---- ANALYSIS D ----
ws=sheet('ANALYSIS_D_RFM'); widths(ws,[28,14,20,18,16,18,14])
r=title(ws,1,'ANALYSIS D — RFM SEGMENTATION')
r=note(ws,r+1,f'METHOD: named customers only ({n_named:,}; cash customers excluded). Recency = days from the customer\'s last invoice to the dataset end ({snap.date()}). Frequency = distinct invoices. Monetary = sum of distinct-invoice GRANDAMOUNT. R and M scored by quintile; F banded because the distribution is highly skewed.')
r=note(ws,r,'ESSENTIAL CAVEAT: RFM is an FMCG framework. On goods with 10-15 year replacement cycles observed over only three years, a LOW FREQUENCY SCORE IS THE EXPECTED CONDITION OF A SATISFIED CUSTOMER, not evidence of defection. Read this as a value-concentration and contactability map, NOT as a churn diagnosis.')
r=title(ws,r+1,'Table D1 — Segment summary',S_FONT)
r=header(ws,r,['Segment','Customers','Revenue (Rs)','Avg monetary (Rs)','Avg frequency','Avg recency (days)','Revenue share'])
first=r
for _,x in segsum.iterrows():
    ws.cell(r,1,x['SEGMENT']).font=BOLD
    for j,(v,fmt) in enumerate([(int(x['CUSTOMERS']),NUM),(float(x['REVENUE']),NUM),(float(x['AVG_MONETARY']),NUM),
                                (float(x['AVG_FREQUENCY']),'0.00'),(float(x['AVG_RECENCY_DAYS']),'0')],2):
        c=ws.cell(r,j,v); c.font=BODY; c.number_format=fmt
    ws.cell(r,7,f'=C{r}/$C${first+len(segsum)}').number_format=PCT
    for k in range(1,8): ws.cell(r,k).border=BORD
    r+=1
ws.cell(r,1,'TOTAL').font=BOLD
for j,col in [(2,'B'),(3,'C')]:
    c=ws.cell(r,j,f'=SUM({col}{first}:{col}{r-1})'); c.font=BOLD; c.number_format=NUM; c.fill=TOT_FILL
r+=2
r=title(ws,r,'Table D2 — Revenue concentration and frequency structure',S_FONT)
r=header(ws,r,['Metric','Value'])
for k,v,fmt in [('Named customers',n_named,NUM),
                ('Customers transacting ONCE in three years',once,NUM),
                ('Share of customers transacting only once',once/len(rfm),PCT),
                ('Top 10% of customers share of revenue',float(top10),PCT),
                ('Top 20% of customers share of revenue',float(top20),PCT),
                ('Median recency (days)',float(rfm.RECENCY_DAYS.median()),'0'),
                ('Mean frequency (invoices per customer)',float(rfm.FREQUENCY.mean()),'0.00')]:
    ws.cell(r,1,k).font=BODY; c=ws.cell(r,2,v); c.font=BOLD; c.number_format=fmt
    for i in range(1,3): ws.cell(r,i).border=BORD
    r+=1
r=note(ws,r+1,'Revenue is concentrated: the top decile of customers produces ~46% of revenue. The highest-value segment shows a frequency far outside a household durables replacement pattern, which — together with a parallel invoice series billed to trading names — indicates it is substantially COMMERCIAL and SUB-DEALER accounts (hotels, guesthouses, contractors, resellers) rather than households. It should be managed with account-management methods, not consumer marketing.')

# ---- KEY FINDINGS ----
ws=sheet('KEY_FINDINGS'); widths(ws,[6,46,95])
r=title(ws,1,'KEY FINDINGS'); r=header(ws,r+1,['#','Finding','Supporting evidence'])
kf=[(1,'The headline trend is an artefact; the real trend is steady erosion','Reported full-year revenue falls 13.7% then rises 9.3%, suggesting a dip and recovery. On the comparable April-January window the figures are 364,458,534 then 359,611,784 then 335,718,314 — i.e. -1.3% then -6.6%, a 7.9% decline over two years. There was no collapse and no recovery.'),
(2,'Six of seven categories are structurally depleting on 10-15 year replacement cycles','Like-for-like two-year change: Washing Machine Automatic -32.0%, Television -24.2%, Water Purifier RO -20.4%, Refrigerator -16.0%, Washing Machine Semi-Automatic -11.9%, Other -10.2%. Indian manufacturers now offer 9-10 year compressor and 5-20 year motor warranties, so these categories are not due for replacement.'),
(3,'Air Conditioner is the only growing category','+16.8% over two years like-for-like (116,949,119 to 136,574,334). Consistent with AC total-cost-of-ownership inflecting from year five — an iFOREST survey found ~80% of ACs over five years old need annual refrigerant refilling versus ~33% of newer units — plus extreme local climate (June maxima 44-47C) and tourism-driven hospitality demand.'),
(4,'The only growing category is the LEAST cross-sold — the central problem','Air Conditioner has LIFT BELOW 1.0 against every other category at invoice level (0.475 to 0.849) and still below 1.0 at customer level. Meanwhile Television, Refrigerator and Washing Machines cluster together at lifts of 1.28-1.61. AC is a standalone, urgency-driven purchase occasion that does not lead into the portfolio.'),
(5,'Customer value is falling even as customer count rises','Multi-category customers are worth 3.4-3.7x single-category ones, but the single-category share rose from 64.1% to 65.7% to 69.4%. The firm served 7,999 named customers in FY2025-26 versus 6,886 two years earlier, yet earned less like-for-like revenue.'),
(6,'Pricing does not respond to seasonality at all','Air Conditioner monthly price index coefficient of variation is 0.0072 — under 1% movement — despite April-June contributing 39-49% of annual revenue and July-September only ~12%. Television, with no weather seasonality, is more volatile at 0.0128. The revenue-management lever is unused.'),
(7,'Purchase frequency is structurally very low','16,541 of 19,736 named customers (83.8%) transacted exactly once in three years; only 897 transacted three or more times. This is the arithmetically expected outcome for decade-cycle goods observed over three years, NOT evidence of a retention failure.'),
(8,'Revenue is concentrated in a commercial sub-segment','The top decile of customers generates 46.1% of revenue and the top quintile 59.9%. The highest-value segment shows purchase frequency far outside a household pattern and is billed substantially through a parallel invoice series to trading names, indicating B2B and sub-dealer accounts.'),
(9,'Two analytical errors were caught that would have inverted the conclusions','(a) Category revenue was initially computed by attributing invoice-level GRANDAMOUNT to every category on the invoice, roughly doubling totals and wrongly showing AC as third-largest and declining 9.5%. (b) The FY2024-25 file is missing ~56.7m of Feb-Mar billing, manufacturing a false recovery narrative.')]
for a,b,c_ in kf:
    ws.cell(r,1,a).font=BOLD
    cb=ws.cell(r,2,b); cb.font=BOLD; cb.alignment=WRAP
    cc=ws.cell(r,3,c_); cc.font=BODY; cc.alignment=WRAP
    for i in range(1,4): ws.cell(r,i).border=BORD
    ws.row_dimensions[r].height=max(30,13*(len(c_)//92+1)); r+=1

# ---- RECOMMENDATIONS ----
ws=sheet('RECOMMENDATIONS'); widths(ws,[6,12,42,72,46])
r=title(ws,1,'RECOMMENDATIONS')
r=note(ws,r+1,'Ordered by expected impact relative to implementation difficulty.',5)
r=header(ws,r,['#','Priority','Recommendation','Rationale (with supporting numbers)','Main risk'])
recs=[(1,'Critical','Capture customer identity (mobile number) at point of sale on every invoice, including cash sales','There is no customer key in the data today — customers are matched on name strings alone, and 1,200-1,800 line items per year are billed anonymously to CASH CUSTOMER. Every target list this analysis produces is unactionable without this. It is the precondition for recommendations 2, 3 and 5.','Requires a change to billing discipline at the counter; staff compliance must be enforced.'),
(2,'High','Launch a structured Air Conditioner to second-category conversion programme','AC is the only growing category (+16.8%) yet has lift below 1.0 against every other category. Converting even 10% of the 5,554 single-category customers to a second category is worth more than any plausible AC pricing gain, given the 3.4-3.7x value differential. Use the AC installation visit — a home touchpoint no other category reliably produces — and a deferred incentive redeemable during the monsoon trough.','AC buyers may be structurally different (tenants, hospitality, single urgent need) and not cross-sellable. Test cheaply before scaling.'),
(3,'High','Deploy next-best-category targeting against the single-category base','Use the empirical lift matrix rather than assumption: Refrigerator owners toward Washing Machine Semi-Automatic (customer-level lift 1.600) and Television (1.347); Television owners toward Washing Machine Semi-Automatic (1.589) and Washing Machine Automatic (1.477). Time the approach to the October festive secondary peak.','Contactability — the firm may hold names without usable contact details. Depends on recommendation 1.'),
(4,'Medium','Run a controlled monsoon-trough stimulation test','Revenue runs at ~12% of the annual total in July-September against 39-49% in April-June, yet the AC price index moves less than 1% (CoV 0.0072). The firm has no evidence on trough elasticity because it has never varied the lever. Test on a limited set of high-volume SKUs, measured against prior-year SKU baselines.','Reputational. A trust-and-referral dealership may be damaged by visible discounting — prefer bundles, extended warranty or free installation, which vary effective price while preserving headline price.'),
(5,'Medium','Separate and formally account-manage the commercial and sub-dealer book','The highest-value segment produces a disproportionate share of revenue from a small minority of customers, with purchase frequency far outside a household pattern and billing through a parallel trading-name invoice series. Give it named ownership, negotiated terms, pre-season contact and stock reservation ahead of the April peak.','Concentration is also exposure — losing a handful of sub-dealers would be materially damaging. Manage to protect as well as to grow.'),
(6,'Medium','Phase working capital, inventory and staffing to the demand curve','April-June carries 39-49% of annual revenue while July-September carries a stable 12.0-12.8%. Resourcing flat against a demand curve that varies threefold is demonstrably inefficient. This is a margin gain requiring no change in customer behaviour, so it is among the most reliably achievable.','Requires supplier and cash-flow coordination; over-correcting risks stockouts in the peak.'),
(7,'Investigate separately','Assess a fast-moving electronics adjacency (mobiles, audio, wearables, accessories) as a SEPARATE feasibility study','The structural finding that 83.8% of customers transact once in three years means the business fundamentally lacks purchase frequency. A sub-5-year replacement category would generate repeat visits, and each visit is an opportunity to cross-sell the high-value durables that carry the margin. It would also raise local visibility and footfall.','NOT EVALUABLE FROM THIS DATASET — it contains zero fast-moving electronics transactions. This is a hypothesis generated by the analysis, not a conclusion supported by it. Requires its own demand study, competitive scan against mobile retail and e-commerce, and channel-fit testing. Cost structure, SKU turnover and staff skills all differ from the current business.')]
for a,b,c_,e,g in recs:
    ws.cell(r,1,a).font=BOLD; ws.cell(r,2,b).font=BOLD
    for col,val in [(3,c_),(4,e),(5,g)]:
        cc=ws.cell(r,col,val); cc.font=BODY; cc.alignment=WRAP
    for i in range(1,6): ws.cell(r,i).border=BORD; ws.cell(r,i).alignment=WRAP
    ws.row_dimensions[r].height=max(45,13*(len(e)//70+1)); r+=1

# ---- LIMITATIONS ----
ws=sheet('LIMITATIONS'); widths(ws,[6,44,100])
r=title(ws,1,'LIMITATIONS OF THE DATASET AND THE ANALYSIS')
r=note(ws,r+1,'These bound what can legitimately be concluded. Acknowledging them is not a weakness of the study; concealing them would be.',3)
r=header(ws,r,['#','Limitation','Detail and consequence'])
lims=[(1,'The observation window is far too short for the product category','The firm sells goods with 10-15 year replacement cycles and the study observes three years — between one fifth and one third of a single cycle. The data structurally CANNOT distinguish a lost customer from one not yet due. The 83.8% single-purchase rate is the expected arithmetic outcome, not a retention failure. A genuine CLV model would need 10+ years of history; no analytical sophistication substitutes for that observation period.'),
(2,'FY2024-25 data completeness gap','February FY2024-25 has 298 invoices and March only 35, against ~800 and ~1,100 in adjacent years — an estimated 56.7m shortfall. Mitigated by the April-January like-for-like window, but the mitigation is imperfect: it discards two real months from the other years and assumes the gap is confined to Feb-Mar, which cannot be verified from inside the data. The firm should confirm whether this is a truncated export or genuine disruption.'),
(3,'No customer identifier','Customers are identified only by name strings — no unique key, no phone, no address normalisation. Name matching both splits one customer across spelling variants and merges distinct customers sharing a common name. All customer counts, repeat rates and segment sizes carry unquantified error. Separately, 1,200-1,800 anonymous cash line items per year are excluded entirely and may differ systematically from the named population.'),
(4,'No cost or margin data','This is a sales register, not a P&L. There is no cost of goods, landed cost or expense allocation. Every value statement — CLV, the multi-category premium, segment concentration — is REVENUE only. A high-revenue category may carry thin margin, and a cross-sell that raises revenue could reduce profit. Re-evaluate against margin data before committing investment.'),
(5,'Effectively no promotional or discount data','Discount, cashback and freight ledgers are populated in well under 1% of ~67,000 rows. This made promotional effectiveness, discount elasticity and display planning impossible to assess — three taught techniques excluded for this reason. It also leaves an open question: does the business genuinely not discount, or is discounting happening informally inside negotiated BASICRATE and never recorded? If the latter, Analysis C measures recorded rather than realised price.'),
(6,'Pricing evidence is correlational, not experimental','Observed prices are negotiated transaction prices. Variation confounds pricing decisions with customer negotiation, salesperson discretion, product age, mix and channel. The finding that AC prices do not vary seasonally is robust as a description; the inference that varying them would raise revenue is untested. Hence the recommendation is an experiment, not a repricing.'),
(7,'The replacement-cycle explanation is directional, not proven','Supported by published lifespan and warranty data and consistent with the observed pattern, but the dataset contains no product age, no installed-base data and no first-purchase-versus-replacement indicator. It explains the pattern plausibly; it does not establish causation. Competitive entry, credit availability and e-commerce substitution cannot be excluded from within this data.'),
(8,'No external or market context','The data contains only this firm\'s own sales — no market size, no competitor data, no share. It is therefore impossible to tell whether the 7.9% like-for-like decline reflects a shrinking market in which the firm holds share, or a growing market in which it is losing share. These have opposite strategic implications.'),
(9,'Branch network changes excluded by scope','Two branches ceased trading between FY2023-24 and FY2025-26 and one new branch opened in FY2025-26. These structural changes affect aggregate figures but were excluded from the narrative because network expansion is a growth-strategy rather than a marketing question. Aggregate trend lines contain this unmodelled effect.'),
(10,'Category mapping is approximate','Categories are assigned by HSN code. The top codes cover 96.6% of revenue, but the residual Other category aggregates 70+ further codes spanning accessories, small appliances, spares and installation materials and is not a coherent commercial category. Between 33 and 93 rows per year carry purely numeric product codes indicating mis-entered master data; these were flagged, not removed.')]
for a,b,c_ in lims:
    ws.cell(r,1,a).font=BOLD
    cb=ws.cell(r,2,b); cb.font=BOLD; cb.alignment=WRAP
    cc=ws.cell(r,3,c_); cc.font=BODY; cc.alignment=WRAP
    for i in range(1,4): ws.cell(r,i).border=BORD
    ws.row_dimensions[r].height=max(30,13*(len(c_)//97+1)); r+=1

# ---- FURTHER SCOPE ----
ws=sheet('FURTHER_SCOPE'); widths(ws,[6,46,22,74])
r=title(ws,1,'FURTHER SCOPE — WHAT ELSE THIS DATA SUPPORTS')
r=note(ws,r+1,'Answers the question of whether there is scope for additional analysis. Split by what is feasible on the data as it stands versus what needs data the firm does not currently hold.',4)
r=header(ws,r,['#','Analysis','Feasibility','Requirement / note'])
fs=[(1,'SKU-level assortment and dead-stock rationalisation','Feasible now','1,900-2,100 SKUs sold per year. Identify true revenue drivers per branch and long-tail SKUs specific branches should stop stocking. Screened in; not carried forward for scope only.'),
(2,'Branch-level basket comparison','Feasible now','Do cross-category baskets differ by branch? Would show whether the AC isolation problem is uniform or concentrated in particular showrooms.'),
(3,'Margin-based CLV','Needs cost data','Current CLV is revenue-only. A high-revenue category may carry thin margin, and a cross-sell that lifts revenue could reduce profit. Needs cost of goods per SKU.'),
(4,'Price elasticity','Needs an experiment','Observed prices are negotiated, so variation confounds pricing with negotiation, mix and channel. Requires the controlled monsoon-trough test in Recommendation 4.'),
(5,'Replacement-timing survival analysis','Needs 10+ years','The single most valuable future analysis for this business. Requires installation-date capture starting now; three years cannot separate a lost customer from one not yet due.'),
(6,'Conjoint, market-share simulation, advertising effectiveness and planning, promotion and display planning, network/viral','Not feasible','Require primary survey data, competitor volumes, ad spend and exposure, a promotion calendar, or referral/social-graph data. None exist in a GST sales register.')]
for a,b,c_,e in fs:
    ws.cell(r,1,a).font=BOLD; ws.cell(r,2,b).font=BOLD; ws.cell(r,3,c_).font=BODY
    cc=ws.cell(r,4,e); cc.font=BODY; cc.alignment=WRAP
    for i in range(1,5): ws.cell(r,i).border=BORD; ws.cell(r,i).alignment=WRAP
    ws.row_dimensions[r].height=max(30,13*(len(e)//72+1)); r+=1

# ---- DATA SHEETS ----
keep=['FISCAL_YEAR','INVOICENO','INVOICE_DATE','FISCAL_MONTH','FISCAL_MONTH_NAME','BRANCH','CUSTOMERNAME',
      'PRODUCT','HSNCODE','CATEGORY','QUANTITY','BASICRATE','TOTAL','CGSTAMT','SGSTAMT','IGSTAMT',
      'LINE_REVENUE','GRANDAMOUNT','GSTRATE','ZERO_VALUE_FLAG','BAD_SKU_FLAG','CASH_CUSTOMER_FLAG']
master=d[keep].copy(); master['INVOICE_DATE']=master['INVOICE_DATE'].dt.strftime('%Y-%m-%d')
ws=wb.create_sheet('CLEANED_MASTER_DATA')
ws.append(keep)
for c in ws[1]: c.font=H_FONT; c.fill=H_FILL
for row in master.itertuples(index=False): ws.append(list(row))
ws.freeze_panes='A2'

invcols=['FISCAL_YEAR','INVOICENO','INVOICE_DATE','FISCAL_MONTH_NAME','BRANCH','CUSTOMERNAME','GRANDAMOUNT','CASH_CUSTOMER_FLAG']
invt=inv[invcols].copy(); invt['INVOICE_DATE']=invt['INVOICE_DATE'].dt.strftime('%Y-%m-%d')
cats_per_inv=d.groupby('INVOICENO')['CATEGORY'].apply(lambda s:', '.join(sorted(set(s))))
invt['CATEGORIES_ON_INVOICE']=invt['INVOICENO'].map(cats_per_inv)
ws=wb.create_sheet('INVOICE_LEVEL')
ws.append(list(invt.columns))
for c in ws[1]: c.font=H_FONT; c.fill=H_FILL
for row in invt.itertuples(index=False): ws.append(list(row))
ws.freeze_panes='A2'

wb.save(OUT)
print('SAVED', OUT, os.path.getsize(OUT))
