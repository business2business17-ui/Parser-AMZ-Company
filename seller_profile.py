"""Parse seller-owned information and period-specific feedback from public HTML."""
import re
from datetime import datetime, timezone
from urllib.parse import urljoin

from bs4 import BeautifulSoup

PERIOD_IDS = {'1 month': 'thirty', '3 months': 'ninety', '12 months': 'year', 'lifetime': 'lifetime'}
LABELS = {
    'legal_name': r'Business Name|Firmenname',
    'business_address': r'Business Address|Geschäftsadresse',
    'contact_email': r'E-?mail(?: Address)?|Customer Service E-?mail',
    'contact_phone': r'(?:Customer Service |Business )?(?:Phone(?: Number)?|Telephone|Telefonnummer)',
    'company_registration_number': r'(?:Company )?Registration (?:Number|No[.]?)|Company (?:Number|No[.]?)|Handelsregisternummer',
    'vat_number': r'VAT(?: Registration)?(?: (?:Number|ID|No[.]?))?|USt-?Id(?:Nr[.]?)?',
    'eori_number': r'EORI(?: (?:Number|No[.]?))?',
    '_other': r'Business Type|Trade Register|Registered Office|Contact|Detailed Seller Information',
}


def text(node):
    return ' '.join(node.stripped_strings) if node else ''


def count(value):
    match = re.search(r'[0-9][0-9.,\s\u00a0\u202f]*', value)
    return int(re.sub(r'\D', '', match.group())) if match else None


def rating(value):
    match = re.search(r'(?<!\d)([0-5](?:[.,]\d+)?)(?!\d)', value)
    return float(match.group(1).replace(',', '.')) if match else None


def parse_seller_profile(html, seller_id, market, seller_url, storefront_url):
    soup = BeautifulSoup(html, 'lxml')
    title = soup.select_one('#seller-name, #sellerName')
    legal = soup.select_one('#page-section-detail-seller-info, #detailed-seller-info')
    if legal is None:
        candidates = [n for n in soup.select('section,div') if re.search(r'Business Name\s*:', text(n))]
        legal = min(candidates, key=lambda n: len(text(n)), default=None)
    legal_text = text(legal)
    boundary = '|'.join(LABELS.values())
    values = {}
    for key, label in LABELS.items():
        match = re.search(r'(?:' + label + r')\s*:\s*(.*?)(?=\s*(?:' + boundary + r')\s*:|$)', legal_text, re.I)
        values[key] = match.group(1).strip() if match else ''
    # Contacts must be seller-labelled or seller-owned links, never the site footer.
    scopes = [n for n in (legal, soup.select_one('#page-section-about-seller')) if n is not None]
    emails = re.findall(r'[A-Z0-9._%+\-]+@[A-Z0-9.\-]+\.[A-Z]{2,}', values['contact_email'], re.I)
    phones = [values['contact_phone']] if values['contact_phone'] else []
    documents = []
    for scope in scopes:
        for link in scope.select('a[href]'):
            href = link['href']; label = text(link)
            if href.lower().startswith('mailto:'):emails.append(href[7:].split('?')[0])
            if href.lower().startswith('tel:'):phones.append(href[4:])
            if re.search(r'registration (?:certificate|document)|certificate of incorporation|impressum|handelsregister|company register', label, re.I):
                target = urljoin(seller_url, href)
                if target.startswith(('http://', 'https://')) and '/ap/' not in target:documents.append(target)
    values['contact_email'] = ' | '.join(dict.fromkeys(emails))
    values['contact_phone'] = ' | '.join(dict.fromkeys(phones))
    values.pop('_other', None)
    address_parts = values['business_address'].split()
    country = address_parts[-1] if address_parts and re.fullmatch(r'[A-Z]{2}', address_parts[-1]) else ''
    summary = text(soup.select_one('#seller-info-feedback-summary'))
    positive = re.search(r'(\d{1,3}(?:[.,]\d+)?)\s*%\s*positive', summary, re.I)
    feedback = {}
    for period, suffix in PERIOD_IDS.items():
        block = soup.select_one('#rating-' + suffix)
        value = rating(text(block.select_one('.ratings-reviews'))) if block else None
        number = count(text(block.select_one('.ratings-reviews-count'))) if block else None
        feedback[period] = {'rating': value, 'ratings': number, 'positive_percent': None,
                            'status': 'OBSERVED' if value is not None or number is not None else 'NOT_COLLECTED'}
    positive_12m = float(positive.group(1).replace(',', '.')) if positive and '12 months' in summary else None
    feedback['12 months']['positive_percent'] = positive_12m
    primary = feedback['12 months']
    # Use the exact decimal in the period block; the header star icon is rounded.
    visible_rating = primary['rating']
    visible_count = primary['ratings']
    if visible_count is None:
        match = re.search(r'\(([0-9][0-9.,\s]*) ratings\)', summary, re.I)
        visible_count = count(match.group(1)) if match else None
    has_profile = bool(title or legal)
    missing = 'NOT_SHOWN' if has_profile else 'NOT_COLLECTED'
    return {
        'seller_id': seller_id, 'marketplace': market, 'name': text(title),
        'seller_url': seller_url, 'storefront_url': storefront_url, **values,
        'country_code': 'GB' if country == 'UK' else country,
        'positive_12m_percent': positive_12m, 'positive_12m_text': positive.group() if positive_12m is not None else '',
        'visible_rating': visible_rating, 'visible_rating_count': visible_count,
        'feedback_period': '12 months' if visible_rating is not None or visible_count is not None else '',
        'feedback': feedback, 'registration_documents': ' | '.join(dict.fromkeys(documents)),
        'contact_status': 'OBSERVED' if emails or phones else missing,
        'registration_status': 'OBSERVED' if any(values[k] for k in ('company_registration_number','vat_number','eori_number')) or documents else missing,
        'return_policy': text(soup.select_one('#page-section-return-refunds, #page-section-return-policy, #page-section-returns')),
        'observed_at': datetime.now(timezone.utc).isoformat(timespec='seconds'),
        'status': 'OK' if has_profile else 'PROFILE_UNVERIFIED',
    }
