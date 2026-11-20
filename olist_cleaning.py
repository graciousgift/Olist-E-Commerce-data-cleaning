"""Olist Brazilian E-Commerce - data cleaning, matching & verification pipeline.
Input : E_Commerse_Brazil_Data.xlsx (8 sheets, the Kaggle Olist CSVs combined)
Output: Olist_Data_Cleaning_Portfolio.xlsx
Usage : python olist_cleaning.py <input.xlsx> <output.xlsx>
"""
import sys, unicodedata, pandas as pd, numpy as np
src, dst = sys.argv[1], sys.argv[2]
X = pd.read_excel(src, sheet_name=None, dtype=str)
orders, trans, cust, pay, items, geo, prod, rev = [X[k] for k in X]

# ---------- 1. LIST MATCHING: product categories (Portuguese) vs translation list ----------
trans.columns = ['product_category_name', 'product_category_name_english']
if trans.iloc[0, 0] == 'product_category_name':          # header was loaded as a data row
    trans = trans.iloc[1:].reset_index(drop=True)
manual = {  # categories present in products but missing from the translation list
    'pc_gamer': 'pc_gamer',
    'portateis_cozinha_e_preparadores_de_alimentos': 'portable_kitchen_food_preparers'}
cats = prod.product_category_name.value_counts(dropna=True).rename_axis('product_category_name').reset_index(name='products')
cm = cats.merge(trans, how='left', on='product_category_name')
cm['match_status'] = np.where(cm.product_category_name_english.notna(), 'Matched', 'No match in translation list')
cm['english_final'] = cm.product_category_name_english.fillna(cm.product_category_name.map(manual))
cm['action'] = np.where(cm.match_status == 'Matched', 'None',
                        np.where(cm.english_final.notna(), 'Translation added manually (not in source list)', 'Needs review'))
unused = trans[~trans.product_category_name.isin(prod.product_category_name)]
cat_map = dict(zip(cm.product_category_name, cm.english_final))

# ---------- 2. PRODUCTS ----------
num = ['product_name_lenght','product_description_lenght','product_photos_qty','product_weight_g','product_length_cm','product_height_cm','product_width_cm']
for c in num: prod[c] = pd.to_numeric(prod[c])
prod = prod.rename(columns={'product_name_lenght':'product_name_length','product_description_lenght':'product_description_length'})
p = prod.copy()
p['category_english'] = p.product_category_name.map(cat_map)
p['flag_missing_category'] = p.product_category_name.isna()
p['flag_missing_dimensions'] = p[['product_weight_g','product_length_cm','product_height_cm','product_width_cm']].isna().any(axis=1)
p['flag_zero_weight'] = p.product_weight_g.eq(0)
key = ['product_category_name','product_name_length','product_description_length','product_photos_qty','product_weight_g','product_length_cm','product_height_cm','product_width_cm']
has = p.product_category_name.notna() & ~p.flag_missing_dimensions
p['possible_duplicate_group'] = np.nan
grp = p[has].groupby(key, sort=False).ngroup()
size = p[has].groupby(key, sort=False)['product_id'].transform('size')
p.loc[has & (size > 1).reindex(p.index, fill_value=False), 'possible_duplicate_group'] = grp[size > 1]
p['possible_duplicate_group'] = p.possible_duplicate_group.map(lambda v: f'DG{int(v):04d}' if pd.notna(v) else '')
items['price'] = items.price.astype(float); items['freight_value'] = items.freight_value.astype(float)
sales = items.groupby('product_id').agg(units_sold=('order_id', 'size'), revenue=('price', 'sum')).reset_index()
p = p.merge(sales, how='left', on='product_id'); p[['units_sold','revenue']] = p[['units_sold','revenue']].fillna(0)
p['revenue'] = p.revenue.round(2)
def status(r):
    s = []
    if r.flag_missing_category: s.append('Missing category/name/description/photo data')
    if r.flag_missing_dimensions: s.append('Missing weight/dimensions')
    if r.flag_zero_weight: s.append('Weight = 0')
    if r.possible_duplicate_group: s.append('Possible duplicate listing')
    return '; '.join(s) if s else 'OK'
p['verification_status'] = p.apply(status, axis=1)
research = p[p.flag_missing_category | p.flag_missing_dimensions | p.flag_zero_weight].copy()
research = research.sort_values('revenue', ascending=False)[['product_id','verification_status','units_sold','revenue','product_weight_g','product_length_cm','product_height_cm','product_width_cm']]
for c in ['researched_category','source_url','verified_by','date_verified']: research[c] = ''
research.insert(0, 'priority_rank', range(1, len(research) + 1))

# ---------- 3. REVIEWS: duplicate review_id ----------
rd = rev[rev.review_id.duplicated(keep=False)].sort_values(['review_id','order_id'])[['review_id','order_id','review_score','review_creation_date']].copy()
rd['note'] = 'Same review_id attached to several orders (one review covering a multi-order purchase). Not a true duplicate - dedupe by review_id only for review-level stats.'

# ---------- 4. GEOLOCATION ----------
def norm(s): return unicodedata.normalize('NFKD', str(s)).encode('ascii','ignore').decode().lower().strip()
g = geo.copy(); g['lat'] = g.geolocation_lat.astype(float); g['lng'] = g.geolocation_lng.astype(float)
g_rows = len(g); g_exact = int(g.duplicated().sum())
g['city_clean'] = g.geolocation_city.map(norm)
g['out_of_brazil'] = (g.lat > 5.3) | (g.lat < -34) | (g.lng > -34) | (g.lng < -74)
cities_before = g.geolocation_city.nunique(); cities_after = g.city_clean.nunique(); oob = int(g.out_of_brazil.sum())
good = g[~g.out_of_brazil]
zipdim = good.groupby('geolocation_zip_code_prefix').agg(
    city=('city_clean', lambda s: s.mode().iat[0]), state=('geolocation_state', lambda s: s.mode().iat[0]),
    lat=('lat','median'), lng=('lng','median'), source_rows=('lat','size')).reset_index()
zipdim['lat'] = zipdim.lat.round(6); zipdim['lng'] = zipdim.lng.round(6)
cust_zip_missing = int((~cust.customer_zip_code_prefix.isin(g.geolocation_zip_code_prefix)).sum())

# ---------- 5. ORDER VERIFICATION ----------
o = orders.copy()
for c in o.columns[3:]: o[c] = pd.to_datetime(o[c])
pay['payment_value'] = pay.payment_value.astype(float)
exc = []
def add(df, issue, detail):
    d = df[['order_id']].copy(); d['issue'] = issue; d['detail'] = detail(df) if callable(detail) else detail; exc.append(d)
add(o[(o.order_status=='delivered') & o.order_delivered_customer_date.isna()], 'Delivered but no delivery date', 'Status=delivered, delivery date blank')
add(o[(o.order_status!='delivered') & o.order_delivered_customer_date.notna()], 'Delivery date but status not delivered', lambda d: 'Status=' + d.order_status)
add(o[o.order_delivered_carrier_date < o.order_purchase_timestamp], 'Carrier pickup before purchase', 'Impossible sequence')
add(o[o.order_delivered_customer_date < o.order_delivered_carrier_date], 'Delivered before carrier pickup', 'Impossible sequence')
tot = pd.concat([items.groupby('order_id').apply(lambda d: (d.price + d.freight_value).sum()).rename('items_total'),
                 pay.groupby('order_id').payment_value.sum().rename('paid_total')], axis=1)
tot['diff'] = (tot.items_total - tot.paid_total).round(2)
mm = tot[tot.items_total.notna() & tot.paid_total.notna() & (tot['diff'].abs() > 1)].reset_index()
add(mm.assign(), 'Payment total differs from items+freight by > R$1', lambda d: 'items+freight=' + d.items_total.round(2).astype(str) + ' paid=' + d.paid_total.round(2).astype(str))
add(o[~o.order_id.isin(items.order_id)], 'Order has no line items', lambda d: 'Status=' + d.order_status)
add(o[~o.order_id.isin(pay.order_id)], 'Order has no payment record', lambda d: 'Status=' + d.order_status)
exc = pd.concat(exc, ignore_index=True)
ec = exc.issue.value_counts()

# ---------- 6. ISSUES LOG ----------
nrow = lambda d: len(d)
log = [
 ('Products','Category names with no match in translation list', len(cm[cm.match_status!='Matched']), f"{int(cm[cm.match_status!='Matched'].products.sum())} products affected", 'Added English names manually and flagged them','Fixed'),
 ('Products','Translation sheet header loaded as data row (Column1/Column2)', 1, 'Row 1 held the real headers', 'Promoted to header, removed from data','Fixed'),
 ('Products','Misspelled column names (lenght)', 2, 'product_name_lenght, product_description_lenght', 'Renamed to _length','Fixed'),
 ('Products','Missing category, name length, description length, photos', int(p.flag_missing_category.sum()), f"{p.flag_missing_category.mean():.1%} of products", 'Kept; flagged; listed in Research_Queue ranked by revenue','Flagged - needs source data'),
 ('Products','Missing weight and dimensions', int(p.flag_missing_dimensions.sum()), '', 'Flagged; in Research_Queue','Flagged'),
 ('Products','Weight recorded as 0 g', int(p.flag_zero_weight.sum()), '', 'Flagged; in Research_Queue','Flagged'),
 ('Products','Possible duplicate listings (identical category, text lengths, photos, weight, dimensions)', int((p.possible_duplicate_group!='').sum()), f"{p[p.possible_duplicate_group!=''].possible_duplicate_group.nunique()} groups", 'Grouped with DG id; NOT deleted (different product_ids can be genuinely different SKUs) - needs manual check','Flagged'),
 ('Products','Duplicate product_id', int(prod.product_id.duplicated().sum()), '', 'None found','Passed'),
 ('Reviews','review_id appears on more than one order', int(rev.review_id.duplicated().sum()), f"{rev.review_id.duplicated(keep=False).sum()} rows involved", 'Kept (not true duplicates); listed in Duplicate_Reviews','Documented'),
 ('Geolocation','Exact duplicate rows', g_exact, f'{g_exact/g_rows:.1%} of {g_rows:,} rows', 'Removed by collapsing to one row per zip prefix (Geo_Zip_Lookup)','Fixed'),
 ('Geolocation','City spelling variants (accents/case/encoding)', cities_before - cities_after, f'{cities_before:,} distinct names -> {cities_after:,} after normalising', 'Lower-cased, accents removed, trimmed','Fixed'),
 ('Geolocation','Coordinates outside Brazil', oob, '', 'Excluded from zip lookup','Fixed'),
 ('Customers','Customer zip prefixes not found in geolocation', cust_zip_missing, '', 'Left blank in lookup; flagged','Flagged'),
 ('Customers','customer_id is per order; customer_unique_id is per person', len(cust) - cust.customer_unique_id.nunique(), f'{len(cust):,} ids vs {cust.customer_unique_id.nunique():,} people', 'Use customer_unique_id for customer counts','Documented'),
] + [('Orders', k, int(v), '', 'Listed in Order_Exceptions', 'Flagged') for k, v in ec.items()]
log = pd.DataFrame(log, columns=['table','issue','records','detail','action_taken','status'])
log.insert(0, 'id', [f'I{n:02d}' for n in range(1, len(log)+1)])

# ---------- 7. WRITE ----------
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
pcols = ['product_id','product_category_name','category_english','product_name_length','product_description_length','product_photos_qty','product_weight_g','product_length_cm','product_height_cm','product_width_cm','units_sold','revenue','possible_duplicate_group','verification_status']
with pd.ExcelWriter(dst, engine='openpyxl') as w:
    pd.DataFrame({'x':[]}).to_excel(w, 'Summary', index=False)
    log.to_excel(w, 'Issues_Log', index=False)
    cm[['product_category_name','product_category_name_english','english_final','products','match_status','action']].to_excel(w, 'Category_Matching', index=False)
    p[pcols].to_excel(w, 'Products_Clean', index=False)
    research.to_excel(w, 'Research_Queue', index=False)
    exc.to_excel(w, 'Order_Exceptions', index=False)
    rd.to_excel(w, 'Duplicate_Reviews', index=False)
    zipdim.to_excel(w, 'Geo_Zip_Lookup', index=False)
wb = load_workbook(dst)
hf = PatternFill('solid', fgColor='1F3864')
for ws in wb:
    if ws.title == 'Summary': continue
    for c in ws[1]: c.font = Font(bold=True, color='FFFFFF'); c.fill = hf; c.alignment = Alignment(wrap_text=True, vertical='center')
    ws.freeze_panes = 'A2'; ws.auto_filter.ref = ws.dimensions
    for i, col in enumerate(ws.columns, 1):
        m = max((len(str(c.value)) for c in list(col)[:300] if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(i)].width = min(max(m + 2, 10), 60)
ws = wb['Issues_Log']
for r in ws.iter_rows(min_row=2):
    for c in r[3:]: c.alignment = Alignment(wrap_text=True, vertical='top')
ws.column_dimensions['D'].width = 55; ws.column_dimensions['F'].width = 55
s = wb['Summary']; s.sheet_view.showGridLines = False
s['A1'] = 'Olist Brazilian E-Commerce - Data Cleaning, Matching & Verification'; s['A1'].font = Font(bold=True, size=16, color='1F3864')
s['A2'] = 'Source: Kaggle "Brazilian E-Commerce Public Dataset by Olist" (8 tables). Sheet checks below are live formulas on the cleaned sheets.'
rows = [('Metric','Value'),
 ('Products reviewed', '=COUNTA(Products_Clean!A2:A%d)' % (len(p)+1)),
 ('Products with all details OK', '=COUNTIF(Products_Clean!N2:N%d,"OK")' % (len(p)+1)),
 ('Products missing category', '=COUNTIF(Products_Clean!N2:N%d,"Missing category*")' % (len(p)+1)),
 ('Products in possible-duplicate groups', '=COUNTIF(Products_Clean!M2:M%d,"DG*")' % (len(p)+1)),
 ('Category names matched to translation list', '=COUNTIF(Category_Matching!E2:E%d,"Matched")' % (len(cm)+1)),
 ('Category names fixed manually', '=COUNTIF(Category_Matching!F2:F%d,"Translation added*")' % (len(cm)+1)),
 ('Products still in research queue', '=COUNTA(Research_Queue!B2:B%d)' % (len(research)+1)),
 ('Revenue (R$) tied to research-queue products', '=SUM(Research_Queue!E2:E%d)' % (len(research)+1)),
 ('Order exceptions logged', '=COUNTA(Order_Exceptions!A2:A%d)' % (len(exc)+1)),
 ('Geolocation rows in source', g_rows), ('Geolocation exact duplicates removed', g_exact),
 ('Geolocation zip prefixes after clean-up', '=COUNTA(Geo_Zip_Lookup!A2:A%d)' % (len(zipdim)+1)),
 ('Issues logged', '=COUNTA(Issues_Log!A2:A%d)' % (len(log)+1))]
for n, (a, b) in enumerate(rows, 4):
    s.cell(n, 1, a); s.cell(n, 2, b)
    if n == 4: s.cell(n,1).font = s.cell(n,2).font = Font(bold=True, color='FFFFFF'); s.cell(n,1).fill = s.cell(n,2).fill = hf
    else: s.cell(n, 2).number_format = '#,##0'
s.column_dimensions['A'].width = 48; s.column_dimensions['B'].width = 18
s.cell(19,1,'Sheets').font = Font(bold=True)
for n, t in enumerate(['Issues_Log - every problem found, rule applied and status','Category_Matching - product categories matched against the translation list','Products_Clean - 32,951 products with English category, flags and verification status','Research_Queue - products needing missing details, ranked by revenue, with blank columns for research','Order_Exceptions - orders failing date / payment / completeness checks','Duplicate_Reviews - review_ids repeated across orders','Geo_Zip_Lookup - geolocation table cleaned to one row per zip prefix'], 20): s.cell(n,1,t)
wb.save(dst)
print(log.to_string()); print(len(research), len(exc), len(zipdim), cm[cm.match_status!='Matched'])
print('unused translations', len(unused))
