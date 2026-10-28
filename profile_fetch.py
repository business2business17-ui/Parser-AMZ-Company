"""Load a public seller profile, wait for its content and retain failed responses."""
import json
import logging
import time
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import urlparse

from bs4 import BeautifulSoup
from playwright.sync_api import TimeoutError as PlaywrightTimeoutError

PROFILE_SELECTOR = '#seller-name, #sellerName, #page-section-detail-seller-info, #detailed-seller-info'
MANUAL_STATES = ('CHECKPOINT_REQUIRED', 'BLOCKED_CHALLENGE', 'LOGIN_REQUIRED')


class ProfileAccessRequired(RuntimeError):
    """The run needs human verification or investigation before it can continue."""


def display_url(url):
    parsed = urlparse(url)
    if parsed.path.startswith(('/checkpoint/', '/ap/')):
        return parsed._replace(query='', fragment='').geturl()
    return url


def response_problem(html, status, url):
    soup = BeautifulSoup(html, 'lxml')
    # Inspect visible text/forms, not arbitrary JS containing the word "captcha".
    text = ' '.join(soup.stripped_strings).lower()
    path = urlparse(url).path.lower()
    if path.startswith('/checkpoint/') or status == 'CHECKPOINT_REQUIRED':
        return 'CHECKPOINT_REQUIRED'
    if status == 'BLOCKED_CHALLENGE' or soup.select_one('input#captchacharacters, form[action*="validateCaptcha"]') or any(
        token in text for token in ('enter the characters you see below', 'robot check', 'verify you are a human')
    ):
        return 'BLOCKED_CHALLENGE'
    if path.startswith('/ap/') or soup.select_one('form[name="signIn"], input#ap_password'):
        return 'LOGIN_REQUIRED'
    if status in ('HTTP_401', 'HTTP_403', 'HTTP_429'):
        return status
    if 'access denied' in text:
        return 'ACCESS_DENIED'
    if status.startswith('HTTP_') and status != 'HTTP_200':
        return status
    if any(token in text for token in ("sorry! something went wrong", 'sorry, something went wrong', 'service unavailable')):
        return 'TEMPORARY_ERROR_PAGE'
    if status.startswith('ERROR:'):
        return status
    return ''


def manual_navigation(page, url, navigate, status, timeout=180):
    """Let the user complete the site's UI. Never click or solve its verification."""
    logging.getLogger('amazon-agent').warning(
        '%s: complete the verification in the open browser within %s seconds. '
        'Do not close the browser. The parser will continue after you finish.', status, timeout)
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        page.wait_for_timeout(1000)
        html = page.content()
        if response_problem(html, 'HTTP_200', page.url) not in MANUAL_STATES:
            return navigate(page, url)
    raise ProfileAccessRequired('Verification was not completed. Resume with --no-headless --resume after completing the site verification.')


def save_diagnostic(folder, market, seller_id, attempt, html, metadata):
    folder = Path(folder);folder.mkdir(parents=True, exist_ok=True)
    stem = folder / f'{market}_{seller_id}_attempt{attempt}'
    stem.with_suffix('.html').write_text(html, encoding='utf-8')
    stem.with_suffix('.json').write_text(json.dumps(metadata, ensure_ascii=False, indent=2), encoding='utf-8')
    return str(stem.with_suffix('.html'))


def fetch_profile(page, seller_id, market, url, navigate, parse, diagnostics_dir, manual=False):
    events = []
    seller = None
    terminal = MANUAL_STATES + ('HTTP_401', 'HTTP_403', 'HTTP_429', 'ACCESS_DENIED')
    for attempt in (1, 2):
        html, http_status = navigate(page, url)
        actual_url = page.url
        problem = response_problem(html, http_status, actual_url)
        if manual and problem in MANUAL_STATES:
            try:
                html, http_status = manual_navigation(page, url, navigate, problem)
                actual_url = page.url
                problem = response_problem(html, http_status, actual_url)
            except ProfileAccessRequired:
                # Preserve the challenge and diagnostics before the caller stops.
                html = page.content()
                actual_url = page.url
        seller = parse(html if not problem else '', seller_id, market)
        # A successful HTTP request is not proof that the profile has rendered.
        if not problem and seller['status'] != 'OK':
            try:
                page.locator(PROFILE_SELECTOR).first.wait_for(state='attached', timeout=8000)
            except PlaywrightTimeoutError:
                pass
            html = page.content()
            actual_url = page.url
            problem = response_problem(html, http_status, actual_url)
            seller = parse(html if not problem else '', seller_id, market)
        status = problem or seller['status']
        title = page.title()
        diagnostic = ''
        if status != 'OK':
            if not html:
                html = page.content()
            diagnostic = save_diagnostic(diagnostics_dir, market, seller_id, attempt, html, {
                'requested_url': url, 'final_url': display_url(actual_url), 'page_title': title,
                'http_status': http_status, 'profile_status': status, 'html_length': len(html),
                'observed_at': datetime.now(timezone.utc).isoformat(timespec='seconds'),
            })
            seller['contact_status'] = seller['registration_status'] = 'NOT_COLLECTED'
        seller.update(status=status, page_status=http_status, profile_source=display_url(actual_url),
                      page_title=title, diagnostic_html=diagnostic)
        events.append({'url': url, 'observed_at': datetime.now(timezone.utc).isoformat(timespec='seconds'),
                       'status': f'{http_status}; {status}; attempt={attempt}; title={title}; final_url={display_url(actual_url)}'
                                 + (f'; diagnostic={diagnostic}' if diagnostic else '')})
        if status == 'OK' or status in terminal:
            return seller, events
        if attempt == 1:
            # One ordinary retry only; do not rotate sessions or bypass challenges.
            page.wait_for_timeout(3000)
    return seller, events
