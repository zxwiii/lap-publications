"""
Web of Science Starter API Publication Synchronizer for LAP Lab
================================================================
Official integration with Clarivate WoS Starter API for ResearcherID: J-8774-2017
Author: LAP Lab (Xingwang Zhu)
Domain: https://zxw.ac.cn/

Workflow Architecture:
1. Image Extraction:
   - Reads 'DOI' and '图片链接' from '朱兴旺文章_完整整合版.xlsx' ONLY
   - Automatically populates and updates 'doi_image_map.json'
   - Uses 'doi_image_map.json' to attach TOC images to WoS publications
2. Strict Representative Papers Control:
   - Reads DOIs from 'selected_doi.txt' ONLY
   - No automatic top journal promotion: represents exactly the user's curated list
3. Academic Badges:
   - ESI TOP 1% loaded from 'esi_top_1_percent.txt'
   - Corr./First Author loaded from 'corresponding_titles.txt' + author string analysis
4. AuthorKeywords-Driven Deep Topic Categorization:
   - Directly analyzes Title + Clarivate WoS official Author Keywords
   - Standard 9 Topics strictly ordered:
     Photocatalysis, CO2 Reduction, 2D Materials, H2 Evolution, N2 Conversion,
     Metal Oxides, Environmental Remediation, Defect Engineering, Other Research.
5. Official Clarivate WoS Starter API Querying:
   - Pulls 146+ publications directly with official citation counts
   - Crossref caching for volume and page enrichment
"""

import os
import sys
import re
import json
import time
import ssl
import urllib.request
import urllib.parse
from datetime import datetime

if sys.platform.startswith('win'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
if BASE_DIR:
    os.chdir(BASE_DIR)

# ==================== Configuration ====================
# API Key resolution: 1. Environment variable (GitHub Actions) -> 2. Local wos_api_key.txt (never pushed to git)
KEY_FILE = os.path.join(BASE_DIR, "wos_api_key.txt")
API_KEY = os.environ.get("WOS_API_KEY", "").strip()
if not API_KEY and os.path.exists(KEY_FILE):
    try:
        with open(KEY_FILE, "r", encoding="utf-8") as f:
            API_KEY = f.read().strip()
    except Exception:
        pass

if not API_KEY:
    print("[ERROR] WOS_API_KEY is missing! Please set the environment variable WOS_API_KEY or save it in wos_api_key.txt")
    sys.exit(1)
RESEARCHER_ID = "J-8774-2017"
BASE_URL = "https://api.clarivate.com/apis/wos-starter/v1/documents"
IMG_PREFIX = "https://32031610.s21i.faiusr.com/"

# Core Files
MASTER_EXCEL_FILE = os.path.join(BASE_DIR, "朱兴旺文章_完整整合版.xlsx")
IMAGE_MAP_FILE = os.path.join(BASE_DIR, "doi_image_map.json")
SELECTED_FILE = os.path.join(BASE_DIR, "selected_doi.txt")
ESI_FILE = os.path.join(BASE_DIR, "esi_top_1_percent.txt")
CORR_FILE = os.path.join(BASE_DIR, "corresponding_titles.txt")
CROSSREF_CACHE_FILE = os.path.join(BASE_DIR, "crossref_cache.json")
OUTPUT_JS_FILE = os.path.join(BASE_DIR, "publications_data.js")

# Topic Priority Order
TOPIC_ORDER = [
    'Photocatalysis',
    'CO2 Reduction',
    '2D Materials',
    'H2 Evolution',
    'N2 Conversion',
    'Metal Oxides',
    'Environmental Remediation',
    'Defect Engineering',
    'Other Research'
]

# Journal Top Tiers & Priority
HIGH_IMPACT_JOURNALS = [
    ("Nature Energy", "Nat. Energy", True),
    ("Nature Communications", "NC", True),
    ("Nat. Commun.", "NC", True),
    ("Advanced Materials", "Adv. Mater.", True),
    ("Adv. Mater.", "Adv. Mater.", True),
    ("Advanced Energy Materials", "AEM", True),
    ("Adv. Energy Mater.", "AEM", True),
    ("Angewandte Chemie", "Angew", True),
    ("Angew. Chem.", "Angew", True),
    ("ACS Nano", "ACS Nano", True),
    ("Chemical Engineering Science", "CES", True),
    ("AIChE Journal", "AIChE J.", True),
    ("AIChE J.", "AIChE J.", True),
    ("ACS Catalysis", "ACS Catal.", True),
    ("ACS Catal.", "ACS Catal.", True),
    ("Advanced Functional Materials", "AFM", True),
    ("Advanced Science", "Adv.Sci", False),
    ("Adv. Sci.", "Adv.Sci", False),
    ("Applied Catalysis B", "ACB", False),
    ("Appl. Catal. B", "ACB", False),
    ("Chemical Engineering Journal", "CEJ", False),
    ("Chem. Eng. J.", "CEJ", False),
    ("Chinese Journal of Catalysis", "Chin. J. Catal.", False),
    ("Chin. J. Catal.", "Chin. J. Catal.", False),
    ("Small", "Small", False),
    ("Nano Energy", "Nano Energy", False),
    ("Journal of Materials Chemistry A", "JMCA", False),
    ("J. Mater. Chem. A", "JMCA", False),
    ("Journal of Energy Chemistry", "JEC", False),
    ("J. Energy Chem.", "JEC", False),
    ("ACS Materials Letters", "ACS Mater. Lett.", False),
    ("Environmental Science & Technology", "ES&T", False),
    ("Water Research", "Water Res.", False),
    ("Nano Letters", "Nano Lett.", False),
    ("Aggregate", "Aggregate", False),
]

def clean_doi(raw):
    if not raw:
        return ""
    d = str(raw).strip().lower()
    d = re.sub(r'^(https?://)?(dx\.)?doi\.org/', '', d)
    return d.strip()

def normalize_title(t):
    if not t:
        return ""
    s = re.sub(r"[^a-z0-9]", "", str(t).lower())
    return s.replace("towards", "toward")

def clean_html_tags(text):
    if not text:
        return ""
    return re.sub(r'<[^>]+>', '', text).strip()

def get_journal_info(journal_str):
    if not journal_str:
        return "", False
    j_low = journal_str.lower()
    for full_j, short_j, is_top in HIGH_IMPACT_JOURNALS:
        pattern = r'\b' + re.escape(full_j.lower()) + r'\b'
        if re.search(pattern, j_low) or full_j.lower() in j_low:
            return short_j, is_top
    return "", False

def refined_topics(title, keywords_list=None):
    """
    Categorizes topics directly from Title + Clarivate WoS official Author Keywords.
    """
    clean_t = clean_html_tags(title)
    kw_str = " ".join(keywords_list) if keywords_list else ""
    combined = (clean_t + " " + kw_str).lower()
    topics = []

    # 1. Photocatalysis (光催化)
    if ('photocataly' in combined or 'photo-cataly' in combined or 'photoreduction' in combined or 
        'photo-oxidation' in combined or 'photooxidation' in combined or 'photoinduced' in combined):
        topics.append("Photocatalysis")

    # 2. CO2 Reduction
    if (re.search(r'\bco2\b', combined) or 'carbon dioxide' in combined or 'photoreduction' in combined or 
        'co2rr' in combined or 'reduction of co2' in combined or 'co2 conversion' in combined or 
        'co2 photoreduction' in combined or 'co2 capture' in combined or 'bicarbonate' in combined or
        'syngas' in combined or 'reverse water gas shift' in combined or 'rwgs' in combined or
        'co 2' in combined):
        topics.append("CO2 Reduction")

    # 3. 2D Materials (formerly Black Phosphorus & 2D Materials)
    if ('black phosphorus' in combined or re.search(r'\bbp\b', combined) or 'phosphorene' in combined or 
        'g-c3n4' in combined or 'carbon nitride' in combined or 'c3n4' in combined or 'mxene' in combined or 
        'mxenes' in combined or '2d' in combined or 'nanosheet' in combined or 'nanosheets' in combined or 
        'ultrathin' in combined or 'two-dimensional' in combined or 'mos2' in combined or 'graphene' in combined or 
        'atomic layers' in combined or 'c4n3' in combined):
        topics.append("2D Materials")

    # 4. H2 Evolution
    if ('hydrogen' in combined or re.search(r'\bher\b', combined) or 'water splitting' in combined or 
        'water photoreduction' in combined or 'h2 evolution' in combined or 'h2 generation' in combined or 
        'h2 production' in combined or re.search(r'\bh2\b', combined) or 'water oxidation' in combined or
        'overall water splitting' in combined or re.search(r'\boer\b', combined) or 'h-2 evolution' in combined or
        'oxygen evolution reaction' in combined):
        topics.append("H2 Evolution")

    # 5. N2 Conversion
    t_no_cn = combined.replace('carbon nitride', '').replace('carbon nitrides', '').replace('g-c3n4', '').replace('c3n4', '')
    if (re.search(r'\bn2\b', t_no_cn) or 'nitrogen reduction' in t_no_cn or 'nitrogen fixation' in t_no_cn or 
        'n2 reduction' in t_no_cn or 'n2-to-nh3' in t_no_cn or 'nitrate' in t_no_cn or 'nitrite' in t_no_cn or 
        'ammonia' in t_no_cn or re.search(r'\bnh3\b', t_no_cn) or 'urea synthesis' in t_no_cn):
        topics.append("N2 Conversion")

    # 6. Metal Oxides
    if ('oxide' in combined or 'oxides' in combined or 'bismuth' in combined or re.search(r'\bbi\d', combined) or 
        re.search(r'\bbio[a-z0-9]', combined) or 'tio2' in combined or 'titanium' in combined or 
        'mil-125' in combined or 'in2o3' in combined or 'w18o49' in combined or 'wo3' in combined or 
        'cds' in combined or 'zno' in combined or 'fe2o3' in combined or 'cofe2o4' in combined or 
        'mof' in combined or 'heterojunction' in combined or 'heterostructure' in combined or 
        'schottky' in combined or 'sb2wo6' in combined or 'znin2s4' in combined or 'layered double hydroxide' in combined or
        'ldh' in combined or 'co3o4' in combined or 'metal-organic frameworks' in combined):
        topics.append("Metal Oxides")

    # 7. Environmental Remediation
    if ('degradation' in combined or 'removal' in combined or 'fenton' in combined or 'peroxymonosulfate' in combined or 
        re.search(r'\bpms\b', combined) or 'pollutant' in combined or 'tetracycline' in combined or 'antibiotic' in combined or 
        'bpa' in combined or 'bisphenol' in combined or 'rhodamine' in combined or 'methylene blue' in combined or 
        'microplastics' in combined or 'purification' in combined or 'seawater purification' in combined or 
        'sensor' in combined or 'sensing' in combined or 'advanced oxidation' in combined):
        topics.append("Environmental Remediation")

    # 8. Defect Engineering (formerly Phase & Defect Engineering)
    if ('defect engineering' in combined or 'defect-engineering' in combined or 'phase engineering' in combined or 
        'oxygen vacancy' in combined or 'oxygen vacancies' in combined or 'crystal phase' in combined or 
        'phase transition' in combined or 'carbon vacancy' in combined or 'heterophase' in combined or 'defective' in combined):
        topics.append("Defect Engineering")

    if not topics:
        topics.append("Other Research")

    # Sort topics by specified order
    topics.sort(key=lambda x: TOPIC_ORDER.index(x) if x in TOPIC_ORDER else 99)
    return topics

def is_first_author(authors_str):
    if not authors_str:
        return False
    first_auth = authors_str.split(';')[0].strip()
    return bool(re.search(r'^(Zhu,\s*Xingwang|Zhu,\s*X\.|朱兴旺)', first_auth, re.I))

def sync_images_from_excel_to_json():
    """
    Reads DOI + 图片链接 from 朱兴旺文章_完整整合版.xlsx and updates doi_image_map.json.
    """
    image_map = {}
    if os.path.exists(IMAGE_MAP_FILE):
        try:
            with open(IMAGE_MAP_FILE, 'r', encoding='utf-8') as f:
                image_map = json.load(f)
        except Exception:
            pass

    if os.path.exists(MASTER_EXCEL_FILE):
        try:
            import openpyxl
            wb = openpyxl.load_workbook(MASTER_EXCEL_FILE, data_only=True)
            sheet = wb.active
            header_row = [cell.value for cell in sheet[1]]
            col_doi = 6
            col_img = 10

            for idx, h in enumerate(header_row):
                if not h:
                    continue
                h_str = str(h).strip().lower()
                if 'doi' in h_str:
                    col_doi = idx
                elif '图片链接' in h_str or 'img' in h_str:
                    col_img = idx

            for row in sheet.iter_rows(min_row=2, values_only=True):
                if not row or len(row) <= col_doi:
                    continue
                doi = clean_doi(row[col_doi]) if col_doi < len(row) else ""
                img = str(row[col_img]).strip() if col_img < len(row) and row[col_img] else ""
                if doi and img:
                    image_map[doi] = img

            with open(IMAGE_MAP_FILE, 'w', encoding='utf-8') as f:
                json.dump(image_map, f, ensure_ascii=False, indent=2)
            print(f"[Image Sync] Synced {len(image_map)} image links into {os.path.basename(IMAGE_MAP_FILE)} (from {os.path.basename(MASTER_EXCEL_FILE)})")
        except Exception as e:
            print(f"[Warning] Failed to sync images from Excel: {e}")

    return image_map

def load_selected_dois():
    """
    Reads representative papers strictly from selected_doi.txt.
    """
    selected = set()
    if os.path.exists(SELECTED_FILE):
        with open(SELECTED_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line_str = line.strip()
                if not line_str or line_str.startswith('#'):
                    continue
                d = clean_doi(line_str)
                if d:
                    selected.add(d)
        print(f"[Config] Loaded {len(selected)} representative papers strictly from {os.path.basename(SELECTED_FILE)}")
    return selected

def load_esi_dois():
    """
    Reads ESI Top 1% DOIs from esi_top_1_percent.txt.
    """
    esi = set()
    if os.path.exists(ESI_FILE):
        with open(ESI_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line_str = line.strip()
                if not line_str or line_str.startswith('#'):
                    continue
                d = clean_doi(line_str)
                if d:
                    esi.add(d)
        print(f"[Config] Loaded {len(esi)} ESI Top 1% DOIs from {os.path.basename(ESI_FILE)}")
    return esi

def load_corr_titles():
    """
    Reads corresponding author titles from corresponding_titles.txt.
    """
    corr_titles = []
    if os.path.exists(CORR_FILE):
        with open(CORR_FILE, 'r', encoding='utf-8') as f:
            for line in f:
                line_str = line.strip()
                if not line_str or line_str.startswith('#'):
                    continue
                corr_titles.append(normalize_title(line_str))
        print(f"[Config] Loaded {len(corr_titles)} corresponding author titles from {os.path.basename(CORR_FILE)}")
    return corr_titles

MONTH_MAP = {
    'JAN': 1, 'FEB': 2, 'MAR': 3, 'APR': 4, 'MAY': 5, 'JUN': 6,
    'JUL': 7, 'AUG': 8, 'SEP': 9, 'OCT': 10, 'NOV': 11, 'DEC': 12
}

def parse_month_int(m_str):
    if not m_str:
        return 0
    m_upper = str(m_str).upper()
    for k, v in MONTH_MAP.items():
        if k in m_upper:
            return v
    return 0

def fetch_wos_documents():
    print(f"[WoS API] Connecting to Clarivate Web of Science for ResearcherID: {RESEARCHER_ID}...")
    ctx = ssl.create_default_context()
    page = 1
    limit = 50
    all_hits = []

    while True:
        params = {
            'db': 'WOS',
            'q': f'AI={RESEARCHER_ID}',
            'limit': limit,
            'page': page
        }
        url = BASE_URL + '?' + urllib.parse.urlencode(params)
        req = urllib.request.Request(url, headers={
            'X-ApiKey': API_KEY,
            'Accept': 'application/json'
        })
        try:
            with urllib.request.urlopen(req, context=ctx, timeout=20) as resp:
                data = json.loads(resp.read().decode('utf-8'))
                hits = data.get('hits', [])
                if not hits:
                    break
                all_hits.extend(hits)
                total_avail = data.get('metadata', {}).get('total', 0)
                print(f"  -> Successfully fetched Page {page} ({len(hits)} papers, total {len(all_hits)}/{total_avail})")
                if len(all_hits) >= total_avail:
                    break
                page += 1
                time.sleep(0.2)
        except Exception as e:
            print(f"[Error] WoS API error (Page {page}): {e}")
            break

    print(f"[WoS API] Fetch complete! Successfully retrieved {len(all_hits)} official publications from Clarivate.")
    return all_hits

def enrich_with_crossref(doi, crossref_cache):
    if not doi:
        return {}
    doi_clean = clean_doi(doi)
    if doi_clean in crossref_cache:
        return crossref_cache[doi_clean]

    url = f"https://api.crossref.org/works/{urllib.parse.quote(doi_clean)}"
    req = urllib.request.Request(url, headers={"User-Agent": "LAPLabWebsite/1.0 (mailto:zxw@ujs.edu.cn)"})
    try:
        ctx = ssl.create_default_context()
        with urllib.request.urlopen(req, context=ctx, timeout=5) as resp:
            data = json.loads(resp.read().decode('utf-8'))
            msg = data.get('message', {})
            info = {
                'volume': msg.get('volume', ''),
                'page': msg.get('page', '') or msg.get('article-number', '')
            }
            crossref_cache[doi_clean] = info
            return info
    except Exception:
        pass
    return {}

def run_sync():
    print("=" * 70)
    print("🚀 LAP Lab (zxw.ac.cn) Web of Science Official Auto-Sync Engine")
    print("=" * 70)

    # 1. Extract DOI + 图片链接 from Excel -> doi_image_map.json ONLY
    image_map = sync_images_from_excel_to_json()

    # 2. Load Selected DOIs strictly from selected_doi.txt
    selected_dois = load_selected_dois()

    # 3. Load ESI DOIs from esi_top_1_percent.txt
    esi_dois = load_esi_dois()

    # 4. Load Corr Author Titles from corresponding_titles.txt
    corr_norm_titles = load_corr_titles()

    # 5. Load Crossref Cache
    crossref_cache = {}
    if os.path.exists(CROSSREF_CACHE_FILE):
        try:
            with open(CROSSREF_CACHE_FILE, 'r', encoding='utf-8') as f:
                crossref_cache = json.load(f)
        except Exception:
            pass

    # 6. Fetch WoS Hits
    hits = fetch_wos_documents()
    if not hits:
        print("[Error] No publications retrieved from WoS API. Retaining existing publications_data.js to avoid disruption.")
        return

    articles = []
    new_crossref_fetched = 0

    for h in hits:
        # Title
        raw_title = clean_html_tags(h.get('title', ''))
        title = re.sub(r'\s+', ' ', raw_title).strip()
        norm_t = normalize_title(title)

        # Identifiers
        doi = clean_doi(h.get('identifiers', {}).get('doi', ''))

        # Authors
        authors_list = []
        for author in h.get('names', {}).get('authors', []):
            name = author.get('displayName') or author.get('wosStandard') or ''
            if name:
                authors_list.append(name.strip())
        authors_str = '; '.join(authors_list)

        # Journal / Source
        src = h.get('source', {})
        raw_journal = src.get('sourceTitle', '')
        journal = ' '.join([w.capitalize() if w.lower() not in ['and', 'of', 'for', 'in', 'on', 'the', '&'] else w.lower() for w in raw_journal.split()])
        journal = re.sub(r'\bAcs\b', 'ACS', journal)
        journal = re.sub(r'\bRsc\b', 'RSC', journal)
        journal = re.sub(r'\bAiche\b', 'AIChE', journal)
        journal = re.sub(r'\bJcis\b', 'JCIS', journal)

        year = str(src.get('publishYear', ''))
        vol = str(src.get('volume', '') or '')
        pages = ''
        if isinstance(src.get('pages'), dict):
            pages = src.get('pages', {}).get('range', '') or src.get('pages', {}).get('begin', '')
        elif src.get('pages'):
            pages = str(src.get('pages'))

        # Crossref enrichment if vol/pages missing
        if doi and (not vol or not pages):
            c_info = enrich_with_crossref(doi, crossref_cache)
            if c_info:
                if not vol and c_info.get('volume'):
                    vol = c_info.get('volume')
                if not pages and c_info.get('page'):
                    pages = c_info.get('page')
                new_crossref_fetched += 1

        # Image link lookup directly from doi_image_map.json
        img = image_map.get(doi, '')
        if img.startswith(IMG_PREFIX):
            img = img[len(IMG_PREFIX):]
        elif img.startswith("//32031610.s21i.faiusr.com/"):
            img = img[len("//32031610.s21i.faiusr.com/"):]

        # High Impact Journal detection
        short_journal, _ = get_journal_info(journal)

        # Representative / Selected Paper: STRICTLY controlled by selected_doi.txt
        is_sel = (doi in selected_dois)

        # Highly Cited ESI TOP 1%
        is_esi = (doi in esi_dois)

        # Author identity (First or Corr)
        is_corr = False
        for ct in corr_norm_titles:
            if ct and (ct == norm_t or (len(ct) > 20 and (ct in norm_t or norm_t in ct))):
                is_corr = True
                break

        is_first = is_first_author(authors_str)
        is_foc = is_corr or is_first

        # Citations count from WoS
        wos_citations = 0
        for cite_entry in h.get('citations', []):
            if cite_entry.get('db') == 'WOS':
                wos_citations = cite_entry.get('count', 0)
                break

        # Author Keywords directly from WoS
        author_kws = h.get('keywords', {}).get('authorKeywords', [])

        # Match topics directly using Title + WoS official Author Keywords
        topics = refined_topics(title, author_kws)

        articles.append({
            'year': year,
            'volume': vol,
            'pages': pages,
            'journal': journal,
            'top_journal': short_journal,
            'doi': doi,
            'title': title,
            'authors': authors_str,
            'topics': topics,
            'image': img,
            'is_selected': is_sel,
            'is_highly_cited': is_esi,
            'is_first_or_corr': is_foc,
            'citations': wos_citations,
            'publish_month': src.get('publishMonth', '')
        })

    # Sort all publications by Year descending, then month descending, then title
    def parse_year_int(y_str):
        try:
            return int(y_str)
        except Exception:
            return 0

    articles.sort(key=lambda x: (
        parse_year_int(x['year']),
        parse_month_int(x.get('publish_month', '')),
        x['title']
    ), reverse=True)

    # Assign sequential IDs: Total down to 1
    total_count = len(articles)
    for idx, a in enumerate(articles):
        a['id'] = total_count - idx

    # Format compact JSON for front-end
    compact_list = []
    for a in articles:
        compact_list.append({
            'id': a['id'],
            'y': a['year'],
            'v': a['volume'],
            'p': a['pages'],
            't': a['title'],
            'a': a['authors'],
            'j': a['journal'],
            'd': a['doi'],
            'img': a['image'],
            'tp': a['topics'],
            'sel': 1 if a['is_selected'] else 0,
            'esi': 1 if a['is_highly_cited'] else 0,
            'foc': 1 if a['is_first_or_corr'] else 0,
            'top': a['top_journal'],
            'c': a.get('citations', 0)
        })

    # Save Crossref Cache
    if new_crossref_fetched > 0:
        try:
            with open(CROSSREF_CACHE_FILE, 'w', encoding='utf-8') as f:
                json.dump(crossref_cache, f, ensure_ascii=False, indent=2)
            print(f"[Crossref] Cached {new_crossref_fetched} new volume/pages entries.")
        except Exception:
            pass

    # Save publications_data.js
    data_json_min = json.dumps(compact_list, ensure_ascii=False, separators=(',', ':'))
    with open(OUTPUT_JS_FILE, 'w', encoding='utf-8') as f:
        f.write(f"window.PUBLICATIONS_IMG_PREFIX = '{IMG_PREFIX}';\nwindow.PUBLICATIONS_DATA = {data_json_min};\n")

    print("=" * 70)
    print(f"🎉 [Success] Synchronized {len(compact_list)} official WoS publications into publications_data.js!")
    print(f"   - Selected Papers (strictly from selected_doi.txt): {sum(1 for a in compact_list if a['sel'])}")
    print(f"   - ESI TOP 1% Highly Cited (from esi_top_1_percent.txt): {sum(1 for a in compact_list if a['esi'])}")
    print(f"   - Corr. / First Author papers (from corresponding_titles.txt): {sum(1 for a in compact_list if a['foc'])}")
    print(f"   - TOC Images matched (from doi_image_map.json): {sum(1 for a in compact_list if a['img'])}")
    print("=" * 70)

if __name__ == '__main__':
    run_sync()
