"""Amazon seller research assistant. Research only on pages you are permitted to access.
Run: python agent.py --input seeds.csv --out reports
Or: python agent.py --asin B0D9VS4PV7 --marketplace US --out reports
No CAPTCHAs, authentication, geolocation or bot defenses are bypassed.
"""
from __future__ import annotations
import argparse, csv, json, re, time, logging, os, sys, tempfile
import math
from dataclasses import dataclass, asdict, field
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse, parse_qs, quote
from bs4 import BeautifulSoup
from playwright.sync_api import sync_playwright
from openpyxl import Workbook, load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

MARKETS = {
'US':('www.amazon.com','ATVPDKIKX0DER'), 'CA':('www.amazon.ca','A2EUQ1WTGCTBG2'),
'MX':('www.amazon.com.mx','A1AM78C64UM0Y8'), 'UK':('www.amazon.co.uk','A1F83G8C2ARO7P'),
'DE':('www.amazon.de','A1PA6795UKMFR9'), 'FR':('www.amazon.fr','A13V1IB3VIYZZH'),
'ES':('www.amazon.es','A1RKKUPIHCS9HS'), 'IT':('www.amazon.it','APJ6JRA9NG5V4'),
'PL':('www.amazon.pl','A1C3SOZRARQ6R3'), 'NL':('www.amazon.nl','A1805IZSGTT6HS'),
'BE':('www.amazon.com.be','AMEN7PMS3EDWL'), 'SE':('www.amazon.se','A2NODRKZP88ZB9'),
'AE':('www.amazon.ae','A2VIGQ35RCS4UG'), 'SA':('www.amazon.sa','A17E79C6D8DWNP'),
'IN':('www.amazon.in','A21TJRUUN4KGV'), 'AU':('www.amazon.com.au','A39IBJ37TRP1C6'),
'JP':('www.amazon.co.jp','A1VC38T7YXB528')
}
CURRENCY = {'US':'USD','CA':'CAD','MX':'MXN','UK':'GBP','DE':'EUR','FR':'EUR','ES':'EUR','IT':'EUR','PL':'PLN',
'NL':'EUR','BE':'EUR','SE':'SEK','AE':'AED','SA':'SAR','IN':'INR','AU':'AUD','JP':'JPY'}
# Default delivery postal codes (central city of each country). AE/SA choose a city instead of a code.
ZIPS = {'US':'10001','CA':'M5V 2T6','MX':'06600','UK':'SW1A 1AA','DE':'10115','FR':'75001','ES':'28001',
'IT':'00100','PL':'00-001','NL':'1012 AB','BE':'1000','SE':'111 20','IN':'110001','AU':'2000','JP':'100-0001'}
CITY_MARKETS = {'AE':'Dubai','SA':'Riyadh'}
ASIN_RE = re.compile(r'(?<![A-Z0-9])([A-Z0-9]{10})(?![A-Z0-9])', re.I)
DEVICE_RE = re.compile(r'(?i)\b(?:iPhone\s*(?:1[1-9]|[2-9][0-9]?)(?:\s*(?:Pro\s*Max|Pro|Plus|Mini))?|AirPods\s*(?:Pro\s*[123]?|[1234](?:st|nd|rd|th)?\s*Gen(?:eration)?|Max))\b')
PRICE_RE = re.compile(r'(?<!\d)(\d{1,4}(?:[.,]\d{2})?)(?!\d)')
ISO_RE = re.compile(r'\b(?:US|CA|MX|GB|UK|DE|FR|ES|IT|PL|NL|BE|SE|AE|SA|IN|AU|JP|CN|HK|SG|TR|VN|KR|TW|BR)\b')
EMAIL_RE = re.compile(r'\b[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}\b', re.I)
PHONE_RE = re.compile(r'(?<!\w)(?:\+?\d[\d\s().\-]{6,}\d)(?!\w)')
REGISTRATION_RE = re.compile(r'(?i)\b(?:registration(?:\s+number|\s+no\.?)?|company\s+number|handelsregisternummer|registre\s+du\s+commerce|registro\s+mercantil|ust-?id|vat\s*(?:id|number)|eori)\s*[:#-]?\s*([A-Z0-9][A-Z0-9 ./-]{3,40})')
LOG = logging.getLogger('amazon-agent')
BASE_DIR = Path(__file__).resolve().parent

def clean(s): return re.sub(r'\s+', ' ', s or '').strip()
def txt(node): return clean(node.get_text(' ', strip=True)) if node else ''
def pick(soup, *selectors):
    for selector in selectors:
        elem = soup.select_one(selector)
        if elem and txt(elem): return txt(elem)
    return ''
def num(s):
    s = clean(s).replace('\u202f','').replace(' ','')
    m = re.search(r'\d[\d.,]*', s)
    if not m:return None
    v=m.group(0)
    if ',' in v and '.' in v: v=v.replace(',','') if v.rfind('.')>v.rfind(',') else v.replace('.','').replace(',','.')
    elif v.count(',')==1: v=v.replace(',','.') if len(v.split(',')[1])<=2 else v.replace(',','')
    elif v.count('.')>1: v=v.replace('.','')
    try: return float(v)
    except ValueError:return None

def extract_contact_data(text):
    emails=list(dict.fromkeys(EMAIL_RE.findall(text or '')))
    phones=[]
    for value in PHONE_RE.findall(text or ''):
        value=clean(value)
        digits=re.sub(r'\D','',value)
        if 7 <= len(digits) <= 16: phones.append(value)
    registrations=[]
    for match in REGISTRATION_RE.finditer(text or ''):
        value=clean(match.group(1)).rstrip('.,;')
        if value and value not in registrations:registrations.append(value)
    return emails,phones,registrations

def classify_shipping(text):
    value=clean(text)
    low=value.lower()
    if not value:return 'UNKNOWN'
    if any(token in low for token in ('ships from amazon','versand durch amazon','expédié par amazon','spedito da amazon','enviado por amazon')):
        return 'AMAZON'
    if any(token in low for token in ('ships from seller','versand durch den verkäufer','versand durch verkäufer','expédié par le vendeur','spedito dal venditore')):
        return 'SELLER'
    return 'SHOWN_TEXT_ONLY'

def asin_from(text):
    p=urlparse(text)
    match=re.search(r'/(?:dp|gp/product|product)/([A-Z0-9]{10})',p.path,re.I)
    if match:return match.group(1).upper()
    found=ASIN_RE.search(text.upper())
    return found.group(1).upper() if found and not text.startswith('http') else None

def market_from_url(url,default='US'):
    host=(urlparse(url).hostname or '').lower()
    for market,(domain,_) in MARKETS.items():
        if host in (domain,domain.removeprefix('www.')):return market
    return default

def url_asin(asin,market): return f'https://{MARKETS[market][0]}/dp/{asin}'
def url_seller(seller,market):return f'https://{MARKETS[market][0]}/sp?ie=UTF8&seller={quote(seller)}'
def url_store(seller,market):return f'https://{MARKETS[market][0]}/s?me={quote(seller)}&marketplaceID={MARKETS[market][1]}'

def get_specs(soup):
    values={}
    for tr in soup.select('#productDetails_detailBullets_sections1 tr, #productDetails_techSpec_section_1 tr, .prodDetTable tr'):
        k=txt(tr.select_one('th'));v=txt(tr.select_one('td'))
        if k and v:values[k]=v
    for li in soup.select('#detailBullets_feature_div li, #detailBulletsWrapper_feature_div li'):
        k=li.select_one('.a-text-bold')
        if k:
            key=txt(k).strip(' :\u200e');value=txt(li)
            if value.lower().startswith(key.lower()):value=value[len(key):].strip(' :\u200e')
            if value:values.setdefault(key,value)
    return values

def parse_product(html, asin, market):
    s=BeautifulSoup(html,'lxml')
    title=pick(s,'#productTitle','h1#title','#title')
    bullets=[txt(e) for e in s.select('#feature-bullets .a-list-item') if txt(e)]
    specs=get_specs(s)
    image=s.select_one('#landingImage, #imgBlkFront, #main-image')
    image_url=''
    if image:
        image_url=image.get('data-old-hires') or image.get('src') or ''
        if image_url.startswith('data:'): image_url=''
    if not image_url:
        meta=s.select_one('meta[property="og:image"]')
        image_url=meta.get('content','') if meta else ''
    price_text=pick(s,'#corePrice_feature_div .a-offscreen','#corePriceDisplay_desktop_feature_div .a-offscreen','#priceblock_ourprice','#priceblock_dealprice','#apex_desktop .a-price .a-offscreen')
    seller_href=''; seller_name=pick(s,'#sellerProfileTriggerId','#merchant-info a[href*="seller="]')
    for a in s.select('a[href*="seller="]'):
        href=a.get('href','')
        if '/sp?' in href or 'seller=' in href:
            seller_href=href
            if not seller_name: seller_name=txt(a)
            break
    sm=re.search(r'[?&]seller=([A-Z0-9]+)',seller_href,re.I)
    seller_id=sm.group(1) if sm else ''
    shipping=pick(s,'#fulfillerInfoFeature_feature_div','#fulfillerInfoFeature_feature_div .offer-display-feature-text-message','#tabular-buybox-text .tabular-buybox-text')
    if not shipping:
        buybox=s.select('#tabular-buybox .tabular-buybox-text-container')
        shipping=txt(buybox[0]) if buybox else ''
    avail=pick(s,'#availability span','#availability')
    avail_l=avail.lower()
    if not title: status='UNVERIFIED'
    elif any(x in avail_l for x in ('currently unavailable','out of stock','nicht verfügbar','indisponible','agotado','non disponibile')):status='NO'
    elif any(x in avail_l for x in ('in stock','available','auf lager','en stock','disponibile','disponível')):status='YES'
    else:status='UNKNOWN'
    rating_text=pick(s,'#acrPopover .a-icon-alt','#averageCustomerReviews .a-icon-alt','#acrPopover')
    rating=num(rating_text)
    review_text=pick(s,'#acrCustomerReviewText','#acrCustomerReviewLink')
    bsr=' | '.join(v for k,v in specs.items() if 'best sellers rank' in k.lower() or 'bestseller' in k.lower() or 'rang' in k.lower())
    first_available=' | '.join(v for k,v in specs.items() if 'date first available' in k.lower() or 'first available' in k.lower() or 'mise en ligne' in k.lower())
    var_nodes=s.select('#twister .swatchAvailable, #twister .swatchSelect, #variation_color_name li, #variation_size_name li')
    variation='YES' if len(var_nodes)>1 else 'UNKNOWN' if not s.select_one('#twister') else 'NO'
    offers=pick(s,'#olp-upd-new-used a','#olp_feature_div a','#moreBuyingChoices_feature_div')
    offer_count=None
    m=re.search(r'(\d[\d.,]*)\s*(?:offers?|options?|new)',offers,re.I)
    if m:offer_count=int(num(m.group(1)) or 0)
    device= ', '.join(dict.fromkeys(m.group(0) for m in DEVICE_RE.finditer(title+' '+' '.join(bullets))))
    return {'marketplace':market,'asin':asin,'product_url':url_asin(asin,market),'image_url':image_url,
        'title':title,'highlights':' | '.join(bullets),'device_model':device,'price_text':price_text,
        'price_numeric':num(price_text),'currency':CURRENCY.get(market,''),'delivery_zip':'','buybox':'YES' if price_text and (seller_name or shipping) else 'UNKNOWN',
        'in_stock':status,'ships_from':shipping,'ships_from_entity':classify_shipping(shipping),
        'sold_by':seller_name,'seller_id':seller_id,
        'variations':variation,'offer_count':offer_count,'rating':rating,'rating_count':int(num(review_text) or 0) if review_text else None,
        'bsr':bsr,'date_first_available':first_available,'item_details':specs,'source':'amazon_page',
        'observed_at':datetime.now(timezone.utc).isoformat(timespec='seconds'),'status':'OK' if title else 'UNVERIFIED'}

def parse_seller(html, seller_id,market):
    s=BeautifulSoup(html,'lxml');all_text=txt(s)
    title=pick(s,'#sellerName','#seller-name','h1','h2')
    legal_name=''; address=''; returns=''
    emails,phones,registrations=extract_contact_data(all_text)
    vat_match=re.search(r'(?i)\bVAT(?:\s*(?:ID|number|no\.?)?)?\s*[:#-]?\s*([A-Z]{2}[A-Z0-9 .\-/]{4,30})',all_text)
    eori_match=re.search(r'(?i)\bEORI\s*(?:number|no\.?)?\s*[:#-]?\s*([A-Z]{2}[A-Z0-9 .\-/]{4,30})',all_text)
    company_match=re.search(r'(?i)\b(?:company\s+(?:registration\s+)?(?:number|no\.?)|registration\s+number|handelsregisternummer)\s*[:#-]?\s*([A-Z0-9 .\-/]{4,40})',all_text)
    documents=[]
    for link in s.select('a[href]'):
        label=clean(link.get_text(' ',strip=True));href=link.get('href','')
        if re.search(r'(?i)(impressum|legal|business|registration|register|vat|eori|terms)',label+' '+href):
            documents.append(f'{label or href} -> {href}')
    # Block-specific parsing; preserve full content rather than fabricating precise labels.
    for node in s.select('div,section'):
        t=txt(node)
        if len(t)>12000:continue
        if ('Detailed Seller Information' in t or 'Business Name:' in t) and not legal_name:
            m=re.search(r'Business Name:\s*(.{3,180}?)(?=Business Address:|$)',t,re.I)
            if m:legal_name=clean(m.group(1))
            a=re.search(r'Business Address:\s*(.{5,450}?)(?=Shipping Policies|Products|Help|$)',t,re.I)
            if a:address=clean(a.group(1))
        if 'Return & Refund Policies' in t and not returns and len(t)<650:
            returns=t
    positive=None;positive_text=''
    m=re.search(r'(\d{1,3}(?:[.,]\d+)?)\s*%\s*positive',all_text,re.I)
    if m:positive=num(m.group(1));positive_text=m.group(0)
    rating=None
    m=re.search(r'(\d(?:[.,]\d)?)\s*(?:out of|von|sur|de)\s*5',all_text,re.I)
    if m:rating=num(m.group(1))
    rating_count=None
    m=re.search(r'(\d[\d,. ]*)\s*(?:ratings|bewertungen|évaluations|reviews|feedback)',all_text,re.I)
    if m and num(m.group(1)) is not None:rating_count=int(num(m.group(1)))
    feedback={}
    for period in ('1 month','3 months','12 months','lifetime'):
        feedback[period]={'rating': None, 'ratings':None, 'positive_percent':None, 'status':'NOT_COLLECTED'}
    addr_iso=''
    if address:
        m=ISO_RE.findall(address.upper());addr_iso=m[-1] if m else ''
        if addr_iso=='UK':addr_iso='GB'
    return {'seller_id':seller_id,'marketplace':market,'name':title,'seller_url':url_seller(seller_id,market),
            'storefront_url':url_store(seller_id,market),'positive_12m_percent':positive,
            'positive_12m_text':positive_text,'visible_rating':rating,'visible_rating_count':rating_count,'feedback':feedback,
            'legal_name':legal_name,'business_address':address,'country_code':addr_iso,
            'contact_email':' | '.join(emails),'contact_phone':' | '.join(phones),
            'company_registration_number':(company_match.group(1).strip() if company_match else (registrations[0] if registrations else '')),
            'vat_number':vat_match.group(1).strip() if vat_match else '',
            'eori_number':eori_match.group(1).strip() if eori_match else '',
            'registration_documents':' | '.join(dict.fromkeys(documents)),
            'contact_status':'OBSERVED' if emails or phones else 'NOT_SHOWN',
            'registration_status':'OBSERVED' if company_match or vat_match or eori_match or registrations or documents else 'NOT_SHOWN',
            'return_policy':returns,'observed_at':datetime.now(timezone.utc).isoformat(timespec='seconds'),
            'status':'OK' if legal_name or positive is not None else 'PARTIAL'}

def collect_feedback_windows(page, seller):
    """Try to navigate ordinary 1/3/12/lifetime dropdown. Mark windows unverified if not identified."""
    choices={'1 month':['1 month','30 days','30 tage','1 mois'],
             '3 months':['3 months','90 days','3 monate','3 mois'],
             '12 months':['12 months','365 days','12 monate','12 mois'],
             'lifetime':['lifetime','all time','gesamte laufzeit','depuis le début']}
    for period, synonyms in choices.items():
        found=False
        for selector in ['select#feedback-time-period','select[name*="time"]','select[name*="period"]','select']:
            try:
                loc=page.locator(selector).first
                if not loc.count():continue
                opts=loc.locator('option')
                for i in range(min(opts.count(),20)):
                    label=clean(opts.nth(i).inner_text()).lower()
                    if any(k in label for k in synonyms):
                        value=opts.nth(i).get_attribute('value')
                        if value is None:continue
                        loc.select_option(value=value,timeout=5000)
                        page.wait_for_timeout(700)
                        found=True;break
                if found:break
            except Exception:continue
        # Pages commonly require further clicks; do not label general seller rating as time-specific.
        if found:
            text_page=txt(BeautifulSoup(page.content(),'lxml'))
            m=re.search(r'(\d(?:[.,]\d)?)\s*(?:out of|von|sur|de)\s*5\s*(?:stars?|stern|étoiles?)?',text_page,re.I)
            n=re.search(r'(\d[\d,\. ]*)\s*(?:ratings|bewertungen|évaluations)',text_page,re.I)
            p=re.search(r'(\d{1,3}(?:[.,]\d+)?)\s*%\s*positive',text_page,re.I)
            seller['feedback'][period]={'rating':num(m.group(1)) if m else None,
                'ratings':int(num(n.group(1))) if n and num(n.group(1)) is not None else None,
                'positive_percent':num(p.group(1)) if p else None,
                'status':'OBSERVED' if m or n or p else 'UNVERIFIED'}
    return seller

def detect_block(page):
    t=page.content().lower()
    return any(x in t for x in ['enter the characters you see below','robot check','captcha'])

def navigate(page,url,wait_ms=1200):
    try:
        response=page.goto(url,wait_until='domcontentloaded',timeout=35000)
        page.wait_for_timeout(wait_ms)
        if detect_block(page):return '', 'BLOCKED_CHALLENGE'
        return page.content(), f'HTTP_{response.status if response else "UNKNOWN"}'
    except Exception as exc:return '',f'ERROR:{type(exc).__name__}:{str(exc)[:130]}'

def storefront_asins(html):
    s=BeautifulSoup(html,'lxml');result=[]
    for el in s.select('[data-asin]'):
        asin=(el.get('data-asin') or '').upper()
        if re.fullmatch(r'[A-Z0-9]{10}',asin) and asin not in result:result.append(asin)
    return result

def collect_storefront(page,seller_id,market,max_pages=3):
    collected=[]; seen=set();log=[]
    for index in range(1,max_pages+1):
        url=url_store(seller_id,market)+f'&page={index}'
        html,state=navigate(page,url)
        log.append({'url':url,'status':state,'observed_at':datetime.now(timezone.utc).isoformat(timespec='seconds')})
        if not html:break
        page_asins=storefront_asins(html)
        novel=[a for a in page_asins if a not in seen]
        if not novel:break
        seen.update(novel);collected.extend(novel)
    return collected,log

PRODUCT_COLS=[('marketplace','Marketplace'),('asin','ASIN'),('product_url','Product URL'),('image_url','Image URL'),
('title','Product Title'),('highlights','Item Highlights'),('device_model','Device Model'),('price_text','Price Raw'),
('price_numeric','Price Number'),('currency','Currency'),('buybox','Buy Box'),('in_stock','In Stock'),
('ships_from','Ships From Raw'),('ships_from_entity','Ships From Entity'),('sold_by','Sold By'),('seller_id','Seller ID'),('variations','Variations'),
('offer_count','Offer Count'),('rating','Product Rating'),('rating_count','Review Count'),('bsr','Best Sellers Rank'),
('date_first_available','Date First Available'),('observed_at','Observed UTC'),('status','Data Status'),
('delivery_zip','Delivery ZIP / City')]
SELLER_COLS=['seller_id','marketplace','name','seller_url','storefront_url','legal_name','business_address',
'country_code','positive_12m_percent','positive_12m_text','visible_rating','visible_rating_count',
'contact_email','contact_phone','company_registration_number','vat_number','eori_number',
'registration_documents','contact_status','registration_status','return_policy','page_status','observed_at','status']

def safe_filename(s):return re.sub(r'[^A-Za-z0-9_.-]','_',s)[:85]

def style_sheet(sheet, ncols, default_width=20):
    for c in range(1,ncols+1):
        sheet.column_dimensions[get_column_letter(c)].width=default_width
        cell=sheet.cell(row=1,column=c)
        cell.fill=PatternFill('solid',fgColor='183153');cell.font=Font(color='FFFFFF',bold=True)
        cell.alignment=Alignment(wrap_text=True,vertical='center')
    sheet.row_dimensions[1].height=35
    sheet.freeze_panes='A2'

ILLEGAL_RE=re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f]')
def cell_value(v):
    if isinstance(v,(dict,list)):v=json.dumps(v,ensure_ascii=False)
    if isinstance(v,str):
        v=ILLEGAL_RE.sub('',v)
        # Values from exports/pages are data, never executable Excel formulas.
        return "'"+v if v.startswith('=') else v
    return v

def export_report(out_path,seller,products,events):
    wb=Workbook()
    dashboard=wb.active;dashboard.title='Dashboard'
    for row in [['SELLER INTELLIGENCE','VALUE'],['Seller ID',seller['seller_id']],
       ['Marketplace',seller['marketplace']],['Seller',seller.get('name','')],['Legal Country',seller.get('country_code','')],
       ['Observed products',len(products)],['Product rating < 4.1',sum(isinstance(p.get('rating'),(float,int)) and p['rating']<4.1 for p in products)],
       ['Stock YES',sum(p.get('in_stock')=='YES' for p in products)],['Catalog completeness','OBSERVED ONLY, NOT VERIFIED FULL'],
       ['Collected UTC',datetime.now(timezone.utc).isoformat(timespec='seconds')]]:
        dashboard.append([cell_value(v) for v in row])
    style_sheet(dashboard,2);dashboard.column_dimensions['A'].width=35;dashboard.column_dimensions['B'].width=48
    sht=wb.create_sheet('Products');heads=[c[1] for c in PRODUCT_COLS]
    sht.append(heads)
    for p in products:sht.append([cell_value(p.get(k)) for k,_ in PRODUCT_COLS])
    style_sheet(sht,len(heads))
    for col,width in {'C':44,'D':44,'E':54,'F':65,'G':27,'M':35,'T':38}.items():sht.column_dimensions[col].width=width
    if products:
        # Highlight whole product row whenever a numeric rating is known to be below 4.1.
        rating_col=get_column_letter([k for k,_ in PRODUCT_COLS].index('rating')+1)
        sht.conditional_formatting.add(f'A2:{get_column_letter(len(heads))}{len(products)+1}',
            FormulaRule(formula=[f'AND(ISNUMBER(${rating_col}2),${rating_col}2<4.1)'],
                        fill=PatternFill('solid',bgColor='FEE2E2',fgColor='FEE2E2'),font=Font(color='991B1B')))
    profile=wb.create_sheet('Seller Profile')
    profile.append(['Seller property','Value'])
    for k in SELLER_COLS:profile.append([k,cell_value(str(seller.get(k,'')) if seller.get(k) is not None else '')])
    for period in ('1 month','3 months','12 months','lifetime'):
        d=seller.get('feedback',{}).get(period,{})
        for fieldn in ('rating','ratings','positive_percent','status'):
            profile.append([f'Feedback {period} / {fieldn}',str(d.get(fieldn,'')) if d.get(fieldn) is not None else ''])
    style_sheet(profile,2);profile.column_dimensions['A'].width=38;profile.column_dimensions['B'].width=98
    details=wb.create_sheet('Item Details');details.append(['ASIN','Field','Value'])
    for p in products:
        for k,v in p.get('item_details',{}).items():details.append([p['asin'],cell_value(k),cell_value(v)])
    style_sheet(details,3);details.column_dimensions['B'].width=40;details.column_dimensions['C'].width=90
    log=wb.create_sheet('Run Log');log.append(['UTC','URL','Status'])
    for e in events:log.append([e.get('observed_at',''),e.get('url',''),cell_value(e.get('status',''))])
    style_sheet(log,3);log.column_dimensions['B'].width=80;log.column_dimensions['C'].width=52
    wb.save(str(out_path))

SEED_COLS=('asin','url','product url','product_url','link','amazon url')

def load_seeds(file):
    path=Path(file)
    if path.suffix.lower() in ('.csv','.tsv'):
        with path.open(encoding='utf-8-sig',newline='') as f:
            reader=csv.DictReader(f,delimiter='\t' if path.suffix.lower()=='.tsv' else ',')
            return list(reader)
    if path.suffix.lower()=='.json':
        val=json.loads(path.read_text(encoding='utf-8'))
        return val if isinstance(val,list) else val.get('items',[])
    if path.suffix.lower()=='.xlsx':
        wb=load_workbook(str(path),data_only=True)
        result=[]; found_header=False
        try:
            for ws in wb.worksheets:
                rows=ws.iter_rows(); headers=None
                for n,r in enumerate(rows):
                    if n >= 10:break
                    candidate=[clean(str(c.value or '')).lower() for c in r]
                    if any(h in SEED_COLS for h in candidate):
                        headers=candidate;found_header=True;break
                if headers is None:continue
                for r in rows:
                    row={}
                    for h,c in zip(headers,r):
                        if not h:continue
                        v=str(c.value).strip() if c.value is not None else ''
                        target=c.hyperlink.target if c.hyperlink else None
                        if h in SEED_COLS and target and market_from_url(target,None):v=target
                        row[h]=v
                    if any(row.values()):result.append(row)
        finally:
            wb.close()
        if found_header:return result
        raise ValueError('No ASIN or URL column found in XLSX.')
    raise ValueError('Supported input: CSV, TSV, JSON, XLSX')

def market_from_name(name):
    low=name.lower()
    for market,(domain,_) in sorted(MARKETS.items(),key=lambda item:len(item[1][0]),reverse=True):
        if domain.removeprefix('www.') in low:return market
    for tok in re.split(r'[^a-z0-9]+',low):
        t={'gb':'uk'}.get(tok,tok).upper()
        if len(tok)==2 and t in MARKETS:return t
    return None

def load_folder(folder,default_market=None):
    """Read Xray files deterministically, without interactive prompts (also safe in CI)."""
    folder=Path(folder)
    files=sorted(f for f in folder.iterdir() if f.is_file() and f.suffix.lower()=='.xlsx' and not f.name.startswith('~$'))
    if not files:raise ValueError(f'No .xlsx files in {folder}. Add a Helium 10 Xray export first.')
    seeds=[]
    for f in files:
        market=market_from_name(f.stem) or default_market
        try:
            found=resolve_seeds(load_seeds(f),market)
            if not found:raise ValueError('No valid ASIN seeds found')
        except Exception as exc:raise ValueError(f'{f.name}: {exc}') from exc
        LOG.info('%s: %d ASIN(s), default marketplace %s',f.name,len(found),market)
        seeds.extend(found)
    return list(dict.fromkeys(seeds))

def resolve_seeds(rows,default_market):
    result=[]
    for row in rows:
        if not isinstance(row,dict):raise ValueError('Every input row must be an object with ASIN or URL')
        row={clean(str(k)).lower():v for k,v in row.items() if k}
        values=[str(row[k]).strip() for k in SEED_COLS if row.get(k)]
        asin=next((asin_from(v) for v in values if asin_from(v)),None)
        if not asin:
            LOG.warning('Could not resolve ASIN from %r',values);continue
        url_market=next((market_from_url(v,None) for v in values if asin_from(v)==asin and market_from_url(v,None)),None)
        market=clean(str(row.get('marketplace') or row.get('market') or url_market or default_market or '')).upper()
        market={'GB':'UK'}.get(market,market)
        if market not in MARKETS:
            raise ValueError(f'Marketplace missing or invalid for {asin}: {market!r}. Use --marketplace US (or another code), a Marketplace column, an Amazon URL, or a file name such as Xray_US.xlsx.')
        result.append((asin,market))
    return list(dict.fromkeys(result))

def _set_delivery_once(page,market,zipcode):
    """Set the ordinary 'Deliver to' postal code via the site's own UI. Returns (status, shown_location)."""
    if not zipcode:return 'NO_ZIP_FOR_MARKET',''
    html,state=navigate(page,f'https://{MARKETS[market][0]}/')
    if not html:return state,''
    try:
        try:page.wait_for_load_state('load',timeout=15000)
        except Exception:pass
        probe={'SA':'#GLUXCityList','AE':'#GLUXCityWithDistrictCityInput'}.get(market,'#GLUXZipUpdateInput')
        for _ in range(3):  # the header link is sometimes clicked before its handler is attached
            page.locator('#nav-global-location-popover-link, #glow-ingress-block').first.click(timeout=8000)
            try:
                page.locator(probe).first.wait_for(state='attached' if market=='SA' else 'visible',timeout=5000);break
            except Exception:page.wait_for_timeout(1000)
        if market=='SA':
            page.locator('#GLUXCityList').first.wait_for(state='attached',timeout=8000)
            page.locator('#GLUXCityList').first.select_option(label=zipcode)
        elif market=='AE':
            city=page.locator('#GLUXCityWithDistrictCityInput').first
            city.wait_for(timeout=8000);city.click();city.type(zipcode,delay=80);page.wait_for_timeout(1200)
            page.locator('#GLUXCityWithDistrictCityList li',has_text=zipcode).first.click(timeout=5000)
            page.locator('#GLUXCityWithDistrictApplyButton').first.click(timeout=5000)
        else:
            inp=page.locator('#GLUXZipUpdateInput').first
            inp.wait_for(timeout=8000)
            parts=zipcode.split(' ')
            if page.locator('#GLUXZipUpdateInput_1').count() and len(parts)==2:
                inp.fill(parts[0]);page.locator('#GLUXZipUpdateInput_1').first.fill(parts[1]);inp.press('Enter')
            else:
                inp.click();inp.type(zipcode,delay=60);inp.press('Enter')
            page.wait_for_timeout(2000)
            if inp.is_visible():  # Enter did not submit: use the Apply button
                page.locator('#GLUXZipUpdate input, #GLUXZipUpdate').first.click(timeout=5000)
        page.wait_for_timeout(1500)
        for sel in ('#GLUXConfirmClose','button[name="glowDoneButton"]','.a-popover-footer .a-button-primary input'):
            try:
                b=page.locator(sel).first
                if b.count() and b.is_visible():b.click(timeout=3000);break
            except Exception:continue
        page.wait_for_timeout(1200)
        page.reload(wait_until='domcontentloaded');page.wait_for_timeout(1000)
        shown=clean(page.locator('#glow-ingress-line2').first.inner_text(timeout=3000))
        ok=zipcode.replace(' ','').lower() in shown.replace(' ','').lower()
        return ('OK' if ok else 'UNVERIFIED'),shown
    except Exception as exc:return f'ERROR:{type(exc).__name__}:{str(exc)[:100]}',''

def set_delivery(page,market,zipcode,attempts=3):
    status,shown='ERROR',''
    for _ in range(attempts):
        status,shown=_set_delivery_once(page,market,zipcode)
        if status in ('OK','NO_ZIP_FOR_MARKET','BLOCKED_CHALLENGE'):break
        page.wait_for_timeout(1500)
    return status,shown

def export_run_report(out_path,seeds,products,groups,events):
    """Export an overall result, including unavailable seed products."""
    wb=Workbook(); summary=wb.active;summary.title='Summary'
    for row in [
        ['Run summary','Value'],['Seed ASINs',len(seeds)],
        ['Seller reports',sum(bool(g.get('seller')) for g in groups.values())],
        ['Observed / unavailable products',len(products)],
        ['Products with status OK',sum(p.get('status')=='OK' for p in products)],
        ['Products requiring attention',sum(p.get('status')!='OK' for p in products)],
        ['Collected UTC',datetime.now(timezone.utc).isoformat(timespec='seconds')],
        ['Coverage','Observed products only; unavailable data remains blank'],
    ]:summary.append(row)
    style_sheet(summary,2);summary.column_dimensions['A'].width=40;summary.column_dimensions['B'].width=70
    sheet=wb.create_sheet('Products');sheet.append([label for _,label in PRODUCT_COLS])
    for p in products:sheet.append([cell_value(p.get(k)) for k,_ in PRODUCT_COLS])
    style_sheet(sheet,len(PRODUCT_COLS));sheet.column_dimensions['E'].width=54
    sellers=wb.create_sheet('Sellers');sellers.append(SELLER_COLS)
    for group in groups.values():
        seller=group.get('seller')
        if seller:sellers.append([cell_value(seller.get(k)) for k in SELLER_COLS])
    style_sheet(sellers,len(SELLER_COLS))
    inputs=wb.create_sheet('Input Seeds');inputs.append(['ASIN','Marketplace','Product URL'])
    for asin,market in seeds:inputs.append([asin,market,url_asin(asin,market)])
    style_sheet(inputs,3);inputs.column_dimensions['C'].width=65
    log=wb.create_sheet('Run Log');log.append(['UTC','URL','Status'])
    for event in events:log.append([cell_value(event.get(k,'')) for k in ('observed_at','url','status')])
    style_sheet(log,3);log.column_dimensions['B'].width=80;log.column_dimensions['C'].width=70
    for ws in (sheet,sellers,inputs,log):ws.auto_filter.ref=ws.dimensions
    out_path=Path(out_path)
    with tempfile.NamedTemporaryFile(dir=out_path.parent,suffix='.xlsx',delete=False) as tmp:
        tmp_path=Path(tmp.name)
    try:
        wb.save(tmp_path);tmp_path.replace(out_path)
    finally:
        wb.close();tmp_path.unlink(missing_ok=True)

def missing_product(asin,market,status):
    return {'asin':asin,'marketplace':market,'product_url':url_asin(asin,market),
            'status':status,'observed_at':datetime.now(timezone.utc).isoformat(timespec='seconds')}

def run(seeds,out_dir,headless=False,max_pages=3,max_products=100,delay=2.0,zips=None):
    out_dir=Path(out_dir);out_dir.mkdir(parents=True,exist_ok=True)
    groups={};events=[];products={}
    zips={**ZIPS,**CITY_MARKETS,**(zips or {})};located={}
    def ensure_location(page,market,sink):
        if market in located:return
        status,shown=set_delivery(page,market,zips.get(market))
        located[market]=zips.get(market,'') if status=='OK' else ''
        sink.append({'url':f'delivery:{market}:{zips.get(market,"")}','status':f'{status} {shown}'.strip(),
            'observed_at':datetime.now(timezone.utc).isoformat(timespec='seconds')})
        if status!='OK':LOG.warning('Delivery location for %s not confirmed: %s',market,status)
    try:
      with sync_playwright() as pw:
        browser=pw.chromium.launch(headless=headless)
        context=browser.new_context(locale='en-US',viewport={'width':1450,'height':1050})
        page=context.new_page()
        for asin,market in seeds:
            ensure_location(page,market,events)
            url=url_asin(asin,market);html,state=navigate(page,url)
            events.append({'url':url,'status':state,'observed_at':datetime.now(timezone.utc).isoformat(timespec='seconds')})
            if not html:
                products[(market,asin)]=missing_product(asin,market,state)
                continue
            p=parse_product(html,asin,market);p['delivery_zip']=located.get(market,'')
            products[(market,asin)]=p
            if not p.get('seller_id'):
                p['status']='NO_SELLER_ID' if p.get('title') else 'UNVERIFIED'
                LOG.warning('No seller ID for %s %s; cannot build seller-specific report',asin,market)
                continue
            key=(market,p['seller_id'])
            group=groups.setdefault(key,{'products':{},'seller':None,'events':[]})
            group['products'][asin]=p
            time.sleep(delay)
        for (market,seller_id),g in groups.items():
            ensure_location(page,market,g['events'])
            url=url_seller(seller_id,market);html,state=navigate(page,url)
            g['events'].append({'url':url,'status':state,'observed_at':datetime.now(timezone.utc).isoformat(timespec='seconds')})
            seller=parse_seller(html,seller_id,market) if html else parse_seller('',seller_id,market)
            seller['page_status']=state
            if not html:
                seller['status']=state
            if html: seller=collect_feedback_windows(page,seller)
            g['seller']=seller
            time.sleep(delay)
            catalog,logs=collect_storefront(page,seller_id,market,max_pages)
            g['events'].extend(logs)
            for asin in catalog[:max_products]:
                if asin in g['products']:continue
                url=url_asin(asin,market);html,state=navigate(page,url)
                g['events'].append({'url':url,'status':state,'observed_at':datetime.now(timezone.utc).isoformat(timespec='seconds')})
                if not html:continue
                p=parse_product(html,asin,market);p['delivery_zip']=located.get(market,'')
                if p['seller_id'] and p['seller_id']!=seller_id:
                    p['status']='DIFFERENT_BUYBOX_SELLER'
                    # Not proof seller does not offer this ASIN; retain with explicit ambiguity.
                g['products'][asin]=p
                time.sleep(delay)
            filename=out_dir/f'{market}_{safe_filename(seller_id)}.xlsx'
            export_report(filename,seller,list(g['products'].values()),g['events'])
            (out_dir/f'{market}_{safe_filename(seller_id)}.json').write_text(json.dumps({'seller':seller,'products':list(g['products'].values()),'events':g['events']},ensure_ascii=False,indent=2),encoding='utf-8')
            LOG.info('Wrote %s (%d observed ASINs)',filename,len(g['products']))
        browser.close()
    except Exception as exc:
        events.append({'url':'run','status':f'ERROR:{type(exc).__name__}:{exc}',
                       'observed_at':datetime.now(timezone.utc).isoformat(timespec='seconds')})
        raise
    finally:
        for group in groups.values():
            events.extend(group['events'])
            for p in group['products'].values():products.setdefault((p['marketplace'],p['asin']),p)
        for asin,market in seeds:products.setdefault((market,asin),missing_product(asin,market,'NOT_PROCESSED'))
        export_run_report(out_dir/'result.xlsx',seeds,list(products.values()),groups,events)
        LOG.info('Overall result: %s',out_dir/'result.xlsx')
    return len(groups)

def load_config(path):
    defaults={'input_dir':'xray helium10','output_dir':'reports','marketplace':None,
              'headless':False,'max_pages':3,'max_products':100,'delay':2.0}
    raw=json.loads(Path(path).read_text(encoding='utf-8-sig'))
    if not isinstance(raw,dict):raise ValueError('Config must contain a JSON object')
    unknown=set(raw)-set(defaults)
    if unknown:raise ValueError(f'Unknown config keys: {sorted(unknown)}')
    defaults.update(raw)
    for key in ('input_dir','output_dir'):
        if not isinstance(defaults[key],str) or not defaults[key].strip():raise ValueError(f'{key} must be a nonempty path')
    if defaults['marketplace'] is not None:
        defaults['marketplace']=str(defaults['marketplace']).strip().upper()
        if defaults['marketplace'] not in MARKETS:raise ValueError('Invalid config marketplace')
    if not isinstance(defaults['headless'],bool):raise ValueError('headless must be true or false')
    for key in ('max_pages','max_products'):
        if type(defaults[key]) is not int or defaults[key]<1:raise ValueError(f'{key} must be a positive integer')
    if type(defaults['delay']) not in (float,int) or not math.isfinite(defaults['delay']) or defaults['delay']<0:
        raise ValueError('delay must be a finite nonnegative number')
    return defaults

def project_path(value):
    path=Path(value).expanduser()
    return path if path.is_absolute() else BASE_DIR/path

def main(argv=None):
    ap=argparse.ArgumentParser(description=__doc__)
    source=ap.add_mutually_exclusive_group()
    source.add_argument('--input',help='Xray folder or one CSV/TSV/JSON/XLSX file')
    source.add_argument('--asin',help='Seed ASIN; requires marketplace')
    source.add_argument('--url',help='Amazon product URL')
    ap.add_argument('--config',default=str(BASE_DIR/'config.json'),help='JSON config file')
    ap.add_argument('--marketplace',type=str.upper,choices=sorted(MARKETS),help='Fallback market when row, URL and file name do not identify one')
    ap.add_argument('--out',help='Output root; each run creates its own timestamped folder')
    ap.add_argument('--headless',action=argparse.BooleanOptionalAction,default=None,help='No browser window; automatically enabled in CI')
    ap.add_argument('--max-pages',type=int)
    ap.add_argument('--max-products',type=int)
    ap.add_argument('--delay',type=float)
    ap.add_argument('--validate-only',action='store_true',help='Validate input and list ASINs without visiting Amazon')
    ap.add_argument('--zip',action='append',default=[],metavar='MARKET=CODE',help='Override delivery postal code, e.g. --zip US=90210 (repeatable)')
    args=ap.parse_args(argv);logging.basicConfig(level=logging.INFO,format='%(levelname)s %(message)s')
    try:
        config=load_config(project_path(args.config))
        for key in ('marketplace','headless','max_pages','max_products','delay'):
            if getattr(args,key) is None:setattr(args,key,config[key])
        if os.environ.get('CI','').lower() in ('1','true','yes'):args.headless=True
        if args.max_pages<1 or args.max_products<1 or not math.isfinite(args.delay) or args.delay<0:
            raise ValueError('max-pages and max-products must be positive; delay must be finite and nonnegative')
        if args.asin or args.url:
            seeds=resolve_seeds([{'asin':args.asin} if args.asin else {'url':args.url}],args.marketplace)
        else:
            src=project_path(args.input or config['input_dir'])
            if not src.exists():raise ValueError(f'Input not found: {src}')
            seeds=load_folder(src,args.marketplace) if src.is_dir() else resolve_seeds(load_seeds(src),market_from_name(src.stem) or args.marketplace)
        if not seeds:raise ValueError('No valid ASIN seeds found')
        zips={}
        for override in args.zip:
            key,sep,value=override.partition('=');key=key.strip().upper();value=value.strip()
            if not sep or key not in MARKETS or not value:raise ValueError('--zip must be MARKET=CODE, e.g. US=90210')
            zips[key]=value
    except (ValueError,OSError) as exc:ap.error(str(exc))
    if args.validate_only:
        print(json.dumps({'seed_count':len(seeds),'seeds':[{'asin':a,'marketplace':m} for a,m in seeds]},ensure_ascii=False,indent=2))
        return 0
    out=project_path(args.out or config['output_dir'])/datetime.now(timezone.utc).strftime('%Y%m%d_%H%M%S_%f_UTC')
    out.mkdir(parents=True,exist_ok=True)
    file_log=logging.FileHandler(out/'run.log',encoding='utf-8')
    file_log.setFormatter(logging.Formatter('%(asctime)s %(levelname)s %(message)s'))
    logging.getLogger().addHandler(file_log)
    try:
        n=run(seeds,out,args.headless,args.max_pages,args.max_products,args.delay,zips)
        print(f'Finished: {n} seller report(s). Result: {out / "result.xlsx"}')
        if not n:LOG.error('No seller reports collected. See result.xlsx / Run Log for unavailable products.')
        return 0 if n else 1
    except Exception:
        LOG.exception('Run failed. Diagnostic result, if saved: %s',out/'result.xlsx')
        return 1
    finally:
        logging.getLogger().removeHandler(file_log);file_log.close()
if __name__=='__main__':sys.exit(main())
