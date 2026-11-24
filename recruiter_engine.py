"""
Recruiter AI Agent - Mode A (X-Ray Sourcing & Vietnam Talent Mapping)
Automates Google X-Ray search company-by-company in Vietnam without touching LinkedIn login.
"""

import os
import re
import time
import random
import urllib.parse
from typing import List, Dict, Any, Callable, Optional, Tuple
import requests
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

def build_xray_queries(
    companies: List[str],
    titles: List[str],
    backgrounds: List[str],
    strict_headline_bg: bool = True,
) -> Dict[str, str]:
    """
    Builds clean, fast Google X-Ray queries for Vietnam LinkedIn profiles:
    - 1 clean query per title (avoids Google's broken title OR relaxation).
    - Multiple Background & Skills lines are treated as OR (e.g. "Treasury" OR "ALM" OR "ALCM" OR "Asset").
    - Smart job-ad fallback (-jobs -tuyển -"việc làm") is triggered ONLY on companies where Page 1 is polluted.
    """
    queries = {}
    clean_titles = [t.strip() for t in titles if t.strip()] or [""]
    clean_bgs = [b.strip() for b in backgrounds if b.strip()]
    if len(clean_bgs) == 1:
        bg_clause = f'"{clean_bgs[0]}"'
    elif len(clean_bgs) > 1:
        bg_clause = "(" + " OR ".join([f'"{b}"' for b in clean_bgs]) + ")"
    else:
        bg_clause = ""

    for comp in companies:
        c = comp.strip()
        if not c:
            continue
        comp_queries = []
        for t in clean_titles:
            parts = [f'site:vn.linkedin.com/in "{c}"']
            if t:
                parts.append(f'"{t}"')
            if bg_clause:
                parts.append(bg_clause)
            comp_queries.append(" ".join(parts))
        queries[c] = "\n".join(comp_queries)

    return queries


def _extract_card_metadata(card_text: str, raw_title: str, target_company: str) -> Tuple[bool, str, str]:
    """
    Extracts structured metadata from a Google search card for a LinkedIn profile:
    - Verifies the candidate CURRENTLY works at target_company (filtering out ex-employees).
    - Extracts the Current Job Title from Google's 'Location · Current Job Title · Current Company' line.
    Returns (is_current_company, detected_company, meta_current_title).
    """
    comp_lower = target_company.strip().lower()
    lines = [ln.strip() for ln in card_text.splitlines() if ln.strip()]

    # 1. Check Google's structured metadata line: "<Location> · <Current Title> · <Current Company>"
    ignore_markers = ("kinh nghiệm:", "experience:", "giáo dục:", "education:", "followers", "người theo dõi", "kết nối", "connections")
    for ln in lines:
        if "·" in ln and not any(m in ln.lower() for m in ignore_markers):
            parts = [p.strip() for p in ln.split("·") if p.strip()]
            if len(parts) >= 3:
                meta_title = parts[1]
                meta_company = parts[2]
                if comp_lower in meta_company.lower():
                    return True, meta_company, meta_title
                else:
                    # Structured current company is explicitly a DIFFERENT company -> reject
                    return False, meta_company, meta_title

    # 2. Check LinkedIn's standard meta description ("Experience: <Current Company>" / "Kinh nghiệm: <Current Company>")
    exp_match = re.search(r'(?:Kinh nghiệm|Experience)\s*:\s*([^·\n]+)', card_text, flags=re.IGNORECASE)
    if exp_match:
        exp_company = exp_match.group(1).strip()
        if comp_lower in exp_company.lower():
            return True, exp_company, ""
        else:
            return False, exp_company, ""

    # 3. Fallback: Check if target_company is explicitly in the headline (and not "Ex-<company>")
    title_lower = raw_title.lower()
    if comp_lower in title_lower:
        ex_pattern = rf'\b(?:ex|former)[\s\-@]*{re.escape(comp_lower)}\b'
        if not re.search(ex_pattern, title_lower):
            return True, target_company, ""

    # 4. Fallback for truncated headlines (e.g. "Senior Talent Acquisition Specialist at ..."):
    # Check if snippet explicitly states "at <company>" / "@ <company>" / "tại <company>" without a past date range
    snippet_lower = card_text.lower()
    at_pattern = rf'(?:\bat\b|\btại\b|@)\s*{re.escape(comp_lower)}'
    if re.search(at_pattern, snippet_lower):
        # Ensure it's not an old past role (e.g. "2015 - 2019 ... Home Credit")
        past_date_pattern = rf'20[012]\d\s*[-–]\s*(?:thg\s*\d+\s*)?20(?:1\d|2[0-4]).{{0,60}}{re.escape(comp_lower)}'
        if not re.search(past_date_pattern, snippet_lower):
            return True, target_company, ""

    return False, "", ""


def _resolve_google_link(context, page, a_el, href: str) -> str:
    """
    Reliably resolves a Google search result link (including encrypted /goto?url=... redirects)
    to the actual LinkedIn profile URL.
    """
    # 1. Direct LinkedIn URL
    if "linkedin.com/in/" in href and "google." not in href:
        return href.split("?")[0]

    # 2. Check if URL is embedded in query params (e.g. /url?q=https://vn.linkedin.com/in/...)
    if "linkedin.com" in href:
        parsed = urllib.parse.urlparse(href)
        qs = urllib.parse.parse_qs(parsed.query)
        for key in ("q", "url"):
            if key in qs and qs[key]:
                val = qs[key][0]
                if "linkedin.com/in/" in val:
                    return val.split("?")[0]

    # 3. Try fast Referer-authenticated request on /goto or /url
    if href.startswith("/goto?") or href.startswith("/url?"):
        try:
            resp = page.request.get(
                f"https://www.google.com{href}",
                headers={"Referer": page.url},
                max_redirects=0,
                timeout=8000,
            )
            loc = resp.headers.get("location", "")
            if "linkedin.com/in/" in loc:
                return loc.split("?")[0]
        except Exception:
            pass

        # 4. Fallback: Ctrl+Click in live browser context (scroll into view first, fast 0.1s polling, 1 retry)
        for _attempt in range(2):
            try:
                a_el.scroll_into_view_if_needed(timeout=3000)
                with context.expect_page(timeout=8000) as new_page_info:
                    a_el.click(modifiers=["Control"])
                new_page = new_page_info.value
                try:
                    for _ in range(50):
                        cur_url = new_page.url
                        if "linkedin.com/in/" in cur_url:
                            return cur_url.split("?")[0]
                        time.sleep(0.1)
                finally:
                    new_page.close()
            except Exception:
                time.sleep(0.3)

    return ""

def _fetch_serper_page(query: str, api_key: str, page_num: int = 0) -> List[Dict[str, str]]:
    """Fetches one page of Google results via Serper.dev API (2,500 free credits)."""
    url = "https://google.serper.dev/search"
    headers = {
        "X-API-KEY": api_key.strip(),
        "Content-Type": "application/json",
    }
    payload = {
        "q": query,
        "gl": "vn",
        "hl": "en",
        "num": 10,
        "page": page_num + 1,
    }
    resp = requests.post(url, headers=headers, json=payload, timeout=15)
    if resp.status_code in (401, 403):
        raise ValueError("Invalid Serper.dev API key. Please check your key in the sidebar.")
    if resp.status_code in (402, 429):
        raise ValueError("Serper.dev quota exhausted or rate-limited.")
    resp.raise_for_status()
    data = resp.json()

    items = []
    for org in data.get("organic", []):
        title = org.get("title", "").strip()
        subtitle = org.get("subtitle", "").strip()
        link = org.get("link", "").strip()
        snippet = org.get("snippet", "").strip()
        # Combine subtitle (Location · Current Title · Current Company), attributes, and snippet
        attr_parts = []
        attrs = org.get("attributes")
        if isinstance(attrs, dict):
            attr_parts = [f"{k}: {v}" for k, v in attrs.items() if v]
        card_text = "\n".join([p for p in [title, subtitle, " · ".join(attr_parts), snippet] if p])
        items.append({"raw_title": title, "url": link, "card_text": card_text})
    return items


def _fetch_google_cse_page(query: str, api_key: str, cx_id: str, page_num: int = 0) -> List[Dict[str, str]]:
    """Fetches one page of Google results via Official Google Custom Search JSON API (100 free/day)."""
    start_idx = page_num * 10 + 1
    if start_idx > 91:
        return []
    url = "https://www.googleapis.com/customsearch/v1"
    params = {
        "key": api_key.strip(),
        "cx": cx_id.strip(),
        "q": query,
        "start": start_idx,
        "num": 10,
        "gl": "vn",
        "hl": "en",
    }
    resp = requests.get(url, params=params, timeout=15)
    if resp.status_code in (400, 403):
        err_msg = ""
        try:
            err_msg = resp.json().get("error", {}).get("message", "")
        except Exception:
            pass
        raise ValueError(f"Google Custom Search API error ({resp.status_code}): {err_msg or 'Check your API Key & Search Engine ID (CX).'}")
    if resp.status_code == 429:
        raise ValueError("Google Custom Search daily quota (100/day) reached. Resets tomorrow.")
    resp.raise_for_status()
    data = resp.json()

    items = []
    for it in data.get("items", []):
        title = it.get("title", "").strip()
        link = it.get("link", "").strip()
        snippet = it.get("snippet", "").strip()
        og_desc = ""
        try:
            metatags = it.get("pagemap", {}).get("metatags", [])
            if metatags and isinstance(metatags, list):
                og_desc = metatags[0].get("og:description", "") or metatags[0].get("description", "")
        except Exception:
            pass
        card_text = "\n".join([p for p in [title, og_desc, snippet] if p])
        items.append({"raw_title": title, "url": link, "card_text": card_text})
    return items


def _execute_api_sourcing(
    companies: List[str],
    titles: List[str],
    backgrounds: List[str],
    max_per_company: int,
    max_pages: int,
    strict_headline_bg: bool,
    search_provider: str,
    api_key: str,
    cse_cx: str,
    status_callback: Optional[Callable[[str], None]],
) -> Tuple[List[Dict[str, str]], List[Dict[str, Any]]]:
    """Fast, zero-rate-limit X-Ray sourcing via Serper.dev or Google Custom Search JSON API."""
    queries = build_xray_queries(companies, titles, backgrounds, strict_headline_bg=strict_headline_bg)
    all_candidates = []
    company_status = []
    seen_urls = set()
    seen_identities = set()
    job_markers = ("tuyển", "việc làm", "joboko", "topcv", "jobs › view", "contract)", "rocketreach", "/jobs/view")

    def log(msg: str):
        print(msg)
        if status_callback:
            status_callback(msg)

    bg_summary = ", ".join(backgrounds[:3]) if backgrounds else "Technology"
    clean_title_filters = [t.strip().lower() for t in titles if t.strip()]
    clean_bg_filters = [b.strip().lower() for b in backgrounds if b.strip()]

    def _scan_api_page(q_text: str, comp_name: str, current_count: int, page_num: int) -> Tuple[int, bool, int]:
        if search_provider == "serper":
            raw_items = _fetch_serper_page(q_text, api_key, page_num)
        else:
            raw_items = _fetch_google_cse_page(q_text, api_key, cse_cx, page_num)

        added = 0
        job_polluted = False
        for item in raw_items:
            if current_count + added >= max_per_company:
                break
            raw_title = item["raw_title"]
            url = item["url"]
            card_text = item["card_text"]

            if any(jm in card_text.lower() or jm in url.lower() for jm in job_markers):
                job_polluted = True
                continue

            # Accept direct LinkedIn URLs AND Google /goto?url=... redirects (Serper returns /goto for ~80% of VN LinkedIn profiles)
            is_linkedin_direct = "linkedin.com/in/" in url.lower()
            is_google_goto = "google.com/goto?" in url.lower()
            if not is_linkedin_direct and not is_google_goto:
                continue

            clean_url = url.split("?")[0] if is_linkedin_direct else url
            clean_header = re.sub(r'\s*\|\s*LinkedIn.*$', '', raw_title, flags=re.IGNORECASE).strip()
            parts = clean_header.split(" - ")
            if len(parts) >= 2:
                cand_name = parts[0].strip()
                cand_title = " - ".join(parts[1:]).strip()
            else:
                cand_name = clean_header
                cand_title = ""

            is_current_company, detected_company, meta_current_title = _extract_card_metadata(card_text, raw_title, comp_name)
            if not is_current_company:
                continue

            combined_role = f"{cand_title} {meta_current_title}".lower()
            if clean_title_filters and not any(tf in combined_role for tf in clean_title_filters):
                continue
            if strict_headline_bg and clean_bg_filters and not any(bf in combined_role for bf in clean_bg_filters):
                continue

            clean_comp = detected_company if (detected_company and len(detected_company) <= 35 and "..." not in detected_company) else comp_name
            identity_key = (cand_name.lower(), comp_name.lower())
            if (is_linkedin_direct and clean_url in seen_urls) or identity_key in seen_identities:
                continue

            if is_linkedin_direct:
                seen_urls.add(clean_url)
            seen_identities.add(identity_key)

            meta_is_garbled = bool(re.search(r'\bthg\b|\b(19|20)\d{2}\b|\bnăm\b|\bmonths?\b|\byears?\b', meta_current_title.lower())) if meta_current_title else True
            display_title = meta_current_title if not meta_is_garbled else (cand_title if cand_title else (titles[0] if titles else "Professional"))

            all_candidates.append({
                "name": cand_name,
                "title": display_title,
                "industry": bg_summary,
                "company": clean_comp,
                "linkedin_link": clean_url,
            })
            added += 1

        return added, job_polluted, len(raw_items)

    provider_label = "Serper.dev API" if search_provider == "serper" else "Google Custom Search API"
    for c_idx, company in enumerate(companies, 1):
        comp_name = company.strip()
        if not comp_name:
            continue

        query_str = queries.get(comp_name, "")
        sub_queries = [q for q in query_str.split("\n") if q.strip()]
        log(f"[{c_idx}/{len(companies)}] ({provider_label}) Mapping: {comp_name}...")

        comp_found_count = 0
        for sq in sub_queries:
            if comp_found_count >= max_per_company:
                break

            active_q = sq
            for page_num in range(max_pages):
                if comp_found_count >= max_per_company:
                    break
                found_here, was_job_polluted, total_items = _scan_api_page(active_q, comp_name, comp_found_count, page_num)
                comp_found_count += found_here

                if page_num == 0 and found_here == 0 and was_job_polluted and comp_found_count < max_per_company:
                    active_q = f'{sq} -jobs -tuyển -"việc làm"'
                    found_fb, _, total_items = _scan_api_page(active_q, comp_name, comp_found_count, 0)
                    comp_found_count += found_fb
                    found_here += found_fb

                if found_here == 0 or total_items < 8:
                    break

        log(f"-> {comp_name}: Found {comp_found_count} profiles.")
        company_status.append({
            "company": comp_name,
            "status": "Found" if comp_found_count > 0 else "Not Found",
            "count": comp_found_count,
            "query": query_str,
        })

    return all_candidates, company_status


def execute_xray_sourcing(
    companies: List[str],
    titles: List[str],
    backgrounds: List[str],
    max_per_company: int = 5,
    max_pages: int = 3,
    strict_headline_bg: bool = True,
    search_provider: str = "browser",
    api_key: str = "",
    cse_cx: str = "",
    status_callback: Optional[Callable[[str], None]] = None
) -> Tuple[List[Dict[str, str]], List[Dict[str, Any]]]:
    """
    Executes automated Google X-Ray company-by-company mapping in Vietnam.
    Supports:
      - 'browser'    : Free local Playwright Chrome X-Ray (default)
      - 'serper'     : Serper.dev Search API (fast, 2,500 free credits)
      - 'google_cse' : Official Google Custom Search JSON API (100 free/day)
    """
    if search_provider in ("serper", "google_cse"):
        return _execute_api_sourcing(
            companies=companies,
            titles=titles,
            backgrounds=backgrounds,
            max_per_company=max_per_company,
            max_pages=max_pages,
            strict_headline_bg=strict_headline_bg,
            search_provider=search_provider,
            api_key=api_key,
            cse_cx=cse_cx,
            status_callback=status_callback,
        )

    from playwright.sync_api import sync_playwright

    queries = build_xray_queries(companies, titles, backgrounds, strict_headline_bg=strict_headline_bg)
    all_candidates = []
    company_status = []
    seen_urls = set()

    def log(msg: str):
        print(msg)
        if status_callback:
            status_callback(msg)

    bg_summary = ", ".join(backgrounds[:3]) if backgrounds else "Technology"

    with sync_playwright() as p:
        browser = p.chromium.launch(
            channel="chrome",
            headless=True,
            args=[
                "--disable-blink-features=AutomationControlled",
                "--no-sandbox",
            ],
        )
        context = browser.new_context(
            user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
            locale="en-US",
        )
        page = context.new_page()

        def _scan_search_page(q_text: str, comp_name: str, clean_title_filters, clean_bg_filters, current_count: int, start_offset: int = 0) -> Tuple[int, bool, int]:
            """Loads one headless Google result page (gl=vn), extracts valid candidates, and returns (added, job_polluted, total_h3s)."""
            search_url = f"https://www.google.com/search?q={urllib.parse.quote(q_text)}&hl=en&gl=vn" + (f"&start={start_offset}" if start_offset > 0 else "")
            added = 0
            job_polluted = False
            total_h3s = 0
            try:
                page.goto(search_url, wait_until="domcontentloaded", timeout=20000)
                if "sorry/index" in page.url or "consent.google" in page.url:
                    log(f"-> Google rate-limit/interstitial hit on {comp_name}. Cooling down 5s...")
                    time.sleep(5)
                    page.goto(search_url, wait_until="domcontentloaded", timeout=20000)
                try:
                    page.wait_for_selector("h3", timeout=6000)
                except Exception:
                    pass
                time.sleep(random.uniform(1.0, 1.5))

                consent_btn = page.query_selector("button:has-text('Accept all'), button:has-text('I agree'), #L2AGLb")
                if consent_btn:
                    consent_btn.click()
                    time.sleep(0.8)

                h3_elements = page.query_selector_all("h3")
                total_h3s = len(h3_elements)
                if not h3_elements:
                    return 0, False, 0

                job_markers = ("tuyển", "việc làm", "joboko", "topcv", "jobs › view", "contract)", "rocketreach")
                for h3 in h3_elements:
                    if current_count + added >= max_per_company:
                        break
                    raw_title = h3.inner_text().strip()
                    if not raw_title or "Google" in raw_title:
                        continue

                    container = h3.evaluate_handle("el => el.closest('div.MjjYud') || el.closest('div.g') || el.closest('[data-hveid]')")
                    card_text = container.evaluate("el => el ? el.innerText : ''") if container else ""
                    card_lower = card_text.lower()
                    if any(jm in card_lower for jm in job_markers):
                        job_polluted = True
                        continue

                    clean_header = re.sub(r'\s*\|\s*LinkedIn.*$', '', raw_title, flags=re.IGNORECASE).strip()
                    parts = clean_header.split(" - ")
                    if len(parts) >= 2:
                        cand_name = parts[0].strip()
                        cand_title = " - ".join(parts[1:]).strip()
                    else:
                        cand_name = clean_header
                        cand_title = ""

                    is_current_company, detected_company, meta_current_title = _extract_card_metadata(card_text, raw_title, comp_name)
                    if not is_current_company:
                        continue

                    combined_role = f"{cand_title} {meta_current_title}".lower()
                    if clean_title_filters and not any(tf in combined_role for tf in clean_title_filters):
                        continue
                    if strict_headline_bg and clean_bg_filters and not any(bf in combined_role for bf in clean_bg_filters):
                        continue

                    a_el = h3.evaluate_handle("el => el.closest('a')")
                    if not a_el:
                        continue
                    href = a_el.get_attribute("href") or ""
                    real_linkedin_url = _resolve_google_link(context, page, a_el, href)

                    if real_linkedin_url and real_linkedin_url not in seen_urls:
                        seen_urls.add(real_linkedin_url)
                        clean_comp = detected_company if (detected_company and len(detected_company) <= 35 and "..." not in detected_company) else comp_name
                        meta_is_garbled = bool(re.search(r'\bthg\b|\b(19|20)\d{2}\b|\bnăm\b|\bmonths?\b|\byears?\b', meta_current_title.lower())) if meta_current_title else True
                        display_title = meta_current_title if not meta_is_garbled else (cand_title if cand_title else (titles[0] if titles else "Professional"))

                        all_candidates.append({
                            "name": cand_name,
                            "title": display_title,
                            "industry": bg_summary,
                            "company": clean_comp,
                            "linkedin_link": real_linkedin_url
                        })
                        added += 1
            except Exception as e:
                log(f"-> Error searching {comp_name}: {str(e)}")
            return added, job_polluted, total_h3s

        for c_idx, company in enumerate(companies, 1):
            comp_name = company.strip()
            if not comp_name:
                continue

            query_str = queries.get(comp_name, "")
            sub_queries = [q for q in query_str.split("\n") if q.strip()]
            log(f"[{c_idx}/{len(companies)}] X-Ray mapping: {comp_name} (up to {max_pages} pages/query)...")

            comp_found_count = 0
            clean_title_filters = [t.strip().lower() for t in titles if t.strip()]
            clean_bg_filters = [b.strip().lower() for b in backgrounds if b.strip()]

            for sq in sub_queries:
                if comp_found_count >= max_per_company:
                    break

                # Fresh context per title query (100% free, zero API credits used)
                try:
                    context.close()
                except Exception:
                    pass
                context = browser.new_context(
                    user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
                    locale="en-US",
                )
                page = context.new_page()

                active_q = sq
                for page_num in range(max_pages):
                    if comp_found_count >= max_per_company:
                        break
                    found_here, was_job_polluted, total_h3s = _scan_search_page(
                        active_q, comp_name, clean_title_filters, clean_bg_filters,
                        comp_found_count, start_offset=page_num * 10
                    )
                    comp_found_count += found_here

                    # Smart fallback on Page 1: ONLY retry with -jobs if Page 1 found 0 AND had job-ad pollution
                    if page_num == 0 and found_here == 0 and was_job_polluted and comp_found_count < max_per_company:
                        active_q = f'{sq} -jobs -tuyển -"việc làm"'
                        found_fb, _, total_h3s = _scan_search_page(
                            active_q, comp_name, clean_title_filters, clean_bg_filters,
                            comp_found_count, start_offset=0
                        )
                        comp_found_count += found_fb
                        found_here += found_fb

                    # Only stop paging early if Google ran out of results on this page (<8 search results)
                    if total_h3s < 8:
                        break

            log(f"-> {comp_name}: Found {comp_found_count} profiles.")

            status_label = "Found" if comp_found_count > 0 else "Not Found"
            company_status.append({
                "company": comp_name,
                "status": status_label,
                "count": comp_found_count,
                "query": query_str
            })

            time.sleep(random.uniform(1.0, 1.8))

        browser.close()
        
    return all_candidates, company_status

def export_mapping_recap_to_excel(
    candidates: List[Dict[str, str]], 
    company_status: List[Dict[str, Any]], 
    file_path: str = "Recruiter_Vietnam_Sourcing_Recap.xlsx"
) -> str:
    """
    Exports a 2-sheet formatted Excel report:
    Sheet 1: Sourced Candidates
    Sheet 2: Search Coverage Recap
    """
    wb = openpyxl.Workbook()
    
    # Styles
    header_fill = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    data_font = Font(name="Calibri", size=10)
    link_font = Font(name="Calibri", size=10, color="0000FF", underline="single")
    found_fill = PatternFill(start_color="D1E7DD", end_color="D1E7DD", fill_type="solid")
    not_found_fill = PatternFill(start_color="F8D7DA", end_color="F8D7DA", fill_type="solid")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    # SHEET 1: Sourced Candidates
    ws1 = wb.active
    ws1.title = "Sourced Candidates"
    headers1 = ["Name", "Title", "Industry / Background", "Company", "LinkedIn Link"]
    ws1.append(headers1)
    for col_idx in range(1, len(headers1) + 1):
        cell = ws1.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws1.row_dimensions[1].height = 26
    
    row_num = 2
    for cand in candidates:
        ws1.cell(row=row_num, column=1, value=cand.get("name", "")).font = data_font
        ws1.cell(row=row_num, column=2, value=cand.get("title", "")).font = data_font
        ws1.cell(row=row_num, column=3, value=cand.get("industry", "")).font = data_font
        ws1.cell(row=row_num, column=4, value=cand.get("company", "")).font = data_font
        
        link_val = cand.get("linkedin_link", "")
        link_cell = ws1.cell(row=row_num, column=5, value=link_val)
        if link_val and link_val.startswith("http"):
            link_cell.hyperlink = link_val
            link_cell.font = link_font
        else:
            link_cell.font = data_font
            
        for c_idx in range(1, 6):
            c = ws1.cell(row=row_num, column=c_idx)
            c.border = thin_border
            c.alignment = Alignment(vertical="center")
        ws1.row_dimensions[row_num].height = 20
        row_num += 1
        
    widths1 = {1: 22, 2: 38, 3: 24, 4: 22, 5: 45}
    for col_idx, width in widths1.items():
        ws1.column_dimensions[get_column_letter(col_idx)].width = width
        
    # SHEET 2: Search Coverage Recap
    ws2 = wb.create_sheet(title="Search Coverage Recap")
    headers2 = ["Company", "Status", "Candidates Sourced", "X-Ray Search Query"]
    ws2.append(headers2)
    for col_idx in range(1, len(headers2) + 1):
        cell = ws2.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    ws2.row_dimensions[1].height = 26
    
    r2 = 2
    for stat in company_status:
        comp_name = stat.get("company", "")
        status_val = stat.get("status", "Not Found")
        count_val = stat.get("count", 0)
        query_val = stat.get("query", "")
        
        ws2.cell(row=r2, column=1, value=comp_name).font = data_font
        status_cell = ws2.cell(row=r2, column=2, value=status_val)
        status_cell.alignment = Alignment(horizontal="center", vertical="center")
        status_cell.font = Font(name="Calibri", size=10, bold=True)
        if status_val == "Found":
            status_cell.fill = found_fill
        else:
            status_cell.fill = not_found_fill
            
        ws2.cell(row=r2, column=3, value=count_val).alignment = Alignment(horizontal="center", vertical="center")
        ws2.cell(row=r2, column=4, value=query_val).font = data_font
        
        for c_idx in range(1, 5):
            ws2.cell(row=r2, column=c_idx).border = thin_border
        ws2.row_dimensions[r2].height = 20
        r2 += 1
        
    widths2 = {1: 25, 2: 15, 3: 20, 4: 65}
    for col_idx, width in widths2.items():
        ws2.column_dimensions[get_column_letter(col_idx)].width = width
        
    wb.save(file_path)
    return file_path
