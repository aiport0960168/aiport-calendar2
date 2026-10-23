import csv
import datetime
import hashlib
import io
import time
import os
import re
import json
import random
import unicodedata
from html import escape as _html_escape
from difflib import SequenceMatcher
from urllib.parse import urljoin, urlparse, parse_qs, quote
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm, Pt, RGBColor
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

try:
    import segno
except ImportError:
    segno = None

BASE_URL = "https://www.kumamoto-aiport.com/cal_copy/pub/default.aspx"
AVAIL_URL = "https://www.kumamoto-aiport.com/reservation_sts/pub/default.aspx"
CACHE_PATH = "cache.json"
HTML_PATH = "index.html"

# ---- Participants (Google Forms) ----
# Fill these three after creating the input form (see the setup guide).
# While empty, the participants feature stays off and the page is unchanged.
PARTICIPANT_FORM_BASE = os.environ.get(
    "AIPORT_PARTICIPANT_FORM_BASE",
    "https://docs.google.com/forms/d/e/1FAIpQLSf2JkhmXlmeohGnR_iF85LoZyI66Xeh086NkwxY1fXbSb4M-A/viewform",
)
PARTICIPANT_ENTRY = {
    "date": os.environ.get("AIPORT_PARTICIPANT_ENTRY_DATE", "entry.385141083"),
    "title": os.environ.get("AIPORT_PARTICIPANT_ENTRY_TITLE", "entry.1526388592"),
    "org": os.environ.get("AIPORT_PARTICIPANT_ENTRY_ORG", "entry.18611752"),
    "key": os.environ.get("AIPORT_PARTICIPANT_ENTRY_KEY", "entry.1974106417"),
}
PARTICIPANT_CSV_URL = os.environ.get(
    "AIPORT_PARTICIPANT_CSV_URL",
    "https://docs.google.com/spreadsheets/d/13jG9rfVADRq0T77W7n2h5g2ORid_mi_1lV6_IBXRlCo/export?format=csv&gid=1403488226",
)

# sumitsuki-kakko (U+3010 / U+3011) built via chr() to keep this file ASCII-safe
LB = chr(0x3010)
RB = chr(0x3011)

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
}


def get_months_to_fetch(now):
    if (now.year == 2027 and now.month >= 4) or (now.year > 2027):
        return 3
    return 6


def classify_room(text):
    has_event = any(k in text for k in ["イベントコーナー", "イベント広場", "イベント"])
    has_seminar = any(k in text for k in ["会議室セミナー室", "会議・セミナー室", "会議セミナー室", "会議室", "セミナー室"])

    if has_event and has_seminar:
        return "イベントコーナー・会議セミナー室"
    elif has_event:
        return "イベントコーナー"
    elif has_seminar:
        return "会議セミナー室"
    else:
        return "その他・未記載"


def extract_bracket_field(lines, keywords):
    pattern = r"^" + LB + r"\s*(?:" + "|".join(keywords) + r")\s*" + RB + r"\s*(.*)"
    for i, line in enumerate(lines):
        m = re.match(pattern, line)
        if m:
            val = m.group(1).strip()
            if not val and i + 1 < len(lines):
                val = lines[i + 1].strip()
            return val
    return ""


def fetch_detail_info(session, detail_url):
    time_mark = LB + "時間" + RB
    try:
        res = session.get(detail_url, headers=HEADERS, timeout=15)
        res.encoding = "utf-8"
        soup = BeautifulSoup(res.text, "html.parser")
        main_div = soup.find("div", id="main") or soup.find("div", id="cont_body") or soup

        raw_time = ""
        for tr in main_div.find_all("tr"):
            tr_text = tr.get_text(" ", strip=True)
            if time_mark in tr_text:
                m = re.search(re.escape(time_mark) + r"\s*(.+)", tr_text)
                if m:
                    raw_time = m.group(1).strip()
                break

        full_text_raw = main_div.get_text(" ", strip=True)
        lines = [line.strip() for line in main_div.get_text(separator="\n").splitlines() if line.strip()]

        org_name = extract_bracket_field(lines, ["団体名", "主催", "主催者"])
        raw_place = extract_bracket_field(lines, ["場 所", "場所", "会 場", "会場"])
        fee = extract_bracket_field(lines, ["参加費", "料金", "費用"])

        if not raw_time:
            raw_time = extract_bracket_field(lines, ["時 間", "時間"])

        room_cat = classify_room(raw_place)
        if room_cat == "その他・未記載":
            room_cat = classify_room(full_text_raw)

        return {
            "time": raw_time,
            "room": room_cat,
            "place": raw_place,
            "org": org_name,
            "fee": fee
        }
    except Exception as e:
        print(f"detail fetch error ({detail_url}): {e}")
        return {"time": "", "room": "その他・未記載", "place": "", "org": "", "fee": ""}


def fetch_availability_day(session, target_date):
    """空き状況ページから1日分の予約済み時間帯を取得する。取れなければNone"""
    params = {"c_id": 4, "sely": target_date.year, "selm": target_date.month, "seld": target_date.day}
    for attempt in range(3):
        try:
            res = session.get(AVAIL_URL, params=params, headers=HEADERS, timeout=30)
            res.encoding = "utf-8"
            soup = BeautifulSoup(res.text, "html.parser")
            title_el = soup.find(id="mttl")
            if title_el:
                m = re.search(r"(\d{4})年\s*(\d{1,2})月\s*(\d{1,2})日", title_el.get_text())
                if m:
                    shown = datetime.date(int(m.group(1)), int(m.group(2)), int(m.group(3)))
                    if shown != target_date:
                        return None
            status_div = soup.find("div", class_="reservationStatus")
            if not status_div:
                return None
            result = {}
            for title_div in status_div.find_all("div", class_="title"):
                room_name = title_div.get_text(strip=True)
                key = "イベントコーナー" if "イベント" in room_name else "会議セミナー室"
                slots = []
                body_text = title_div.find_next_sibling("div", class_="body_text")
                if body_text:
                    for li in body_text.find_all("li"):
                        mm = re.search(r"(\d{1,2}:\d{2})\s*～\s*(\d{1,2}:\d{2})", li.get_text())
                        if mm:
                            slots.append(mm.group(1) + "～" + mm.group(2))
                result[key] = slots
            return result if result else None
        except Exception as e:
            if attempt == 2:
                print(f"空き状況取得失敗 ({target_date}): {e}")
                return None
            time.sleep(2)
    return None


def _parse_time_range(time_str):
    """time文字列（例: 10時00分～12時30分）を (開始分, 終了分) に変換する。読めなければNone"""
    t = (time_str or "").replace("〜", "～").replace("~", "～").replace(" ", "").replace("　", "")
    m = re.search(r"(\d{1,2})時(?:(\d{1,2})分|(半))?\s*～\s*(\d{1,2})時(?:(\d{1,2})分|(半))?", t)
    if not m:
        return None

    def _min(h, mi, half):
        return int(h) * 60 + (int(mi) if mi else (30 if half else 0))

    return (_min(m.group(1), m.group(2), m.group(3)), _min(m.group(4), m.group(5), m.group(6)))


def _venues_of(room):
    """会場カテゴリ文字列が含む会場の集合を返す（両方利用は両会場を返す）"""
    venues = set()
    if "イベント" in (room or ""):
        venues.add("イベントコーナー")
    if "会議" in (room or "") or "セミナー" in (room or ""):
        venues.add("会議・セミナー室")
    return venues


def find_double_use(events):
    """同じ日・同じ会場で時間帯が重なる行事ペアを検出する（会場の二重使用チェック）。
    中止行事と時間が読み取れない行事は対象外。団体名は判定に使わない。"""
    cancel_mark = LB + "中止" + RB
    by_date = {}
    for ev in events:
        if cancel_mark in (ev.get("title") or ""):
            continue
        rng = _parse_time_range(ev.get("time") or "")
        venues = _venues_of(ev.get("room") or "")
        if rng is None or not venues:
            continue
        by_date.setdefault(ev["date"], []).append((rng, venues, ev))
    pairs = []
    for date_str in sorted(by_date):
        lst = by_date[date_str]
        for i in range(len(lst)):
            for j in range(i + 1, len(lst)):
                (s1, e1), v1, a = lst[i]
                (s2, e2), v2, b = lst[j]
                shared = [v for v in ["イベントコーナー", "会議・セミナー室"] if v in v1 and v in v2]
                if shared and s1 < e2 and s2 < e1:
                    pairs.append({
                        "date": date_str,
                        "weekday": a.get("weekday", ""),
                        "venues": shared,
                        "overlap": (max(s1, s2), min(e1, e2)),
                        "a": a,
                        "b": b,
                    })
    return pairs


def _org_norm(text, strong=False):
    """団体名の正規化。strong=True では中黒・読点も除去する（表記ゆれ吸収用）"""
    t = unicodedata.normalize("NFKC", text or "")
    t = re.sub(r"\s+", "", t)
    if strong:
        t = t.replace("・", "").replace("、", "")
    return t


def find_org_name_issues(events):
    """団体名の表記ゆれ・類似名を検出する（要確認リスト用。間違いの断定はしない）。
    - 表記ゆれの疑い: 空白・全角半角・中黒・読点を除くと同一になるのに原文が異なる名前同士
    - 類似名の候補: 類似度が高い、または一方が他方の接頭辞（数字だけの違いは期・班の区別として除外）"""
    counts = {}
    last = {}
    for ev in events:
        org = (ev.get("org") or "").strip()
        if not org:
            continue
        counts[org] = counts.get(org, 0) + 1
        if ev.get("date", "") > last.get(org, ""):
            last[org] = ev.get("date", "")
    names = sorted(counts)
    strong = {n: _org_norm(n, strong=True) for n in names}
    issues = []
    groups = {}
    for n in names:
        groups.setdefault(strong[n], []).append(n)
    for key, g in groups.items():
        if len(g) > 1:
            issues.append({"kind": "表記ゆれの疑い", "names": sorted(g)})
    for i in range(len(names)):
        for j in range(i + 1, len(names)):
            a, b = names[i], names[j]
            sa, sb = strong[a], strong[b]
            if sa == sb:
                continue
            if re.sub(r"[0-9]", "", sa) == re.sub(r"[0-9]", "", sb):
                continue
            ratio = SequenceMatcher(None, sa, sb).ratio()
            shorter, longer = (sa, sb) if len(sa) <= len(sb) else (sb, sa)
            is_prefix = len(shorter) >= 5 and longer.startswith(shorter)
            if ratio >= 0.88 or is_prefix:
                issues.append({"kind": "類似名の候補", "names": [a, b]})
    for iss in issues:
        iss["counts"] = {n: counts[n] for n in iss["names"]}
        iss["last"] = {n: last.get(n, "") for n in iss["names"]}
    return issues


def participant_key(e):
    # Stable key that links a form response to one event.
    base = (e.get("url") or "") or (e["date"] + "|" + (e.get("title") or ""))
    return hashlib.sha1(base.encode("utf-8")).hexdigest()[:10]


def participant_form_link(e):
    # Google Forms URL with the event's date / title / org / key prefilled.
    if not PARTICIPANT_FORM_BASE or not all(PARTICIPANT_ENTRY.values()):
        return ""
    params = [
        (PARTICIPANT_ENTRY["date"], e["date"]),
        (PARTICIPANT_ENTRY["title"], e.get("title", "")),
        (PARTICIPANT_ENTRY["org"], e.get("org", "")),
        (PARTICIPANT_ENTRY["key"], participant_key(e)),
    ]
    qs = "&".join(f"{k}={quote(str(v), safe='')}" for k, v in params)
    sep = "&" if "?" in PARTICIPANT_FORM_BASE else "?"
    return PARTICIPANT_FORM_BASE + sep + "usp=pp_url&" + qs


def fetch_participants():
    # Read the responses CSV; return {event_key: (total, non_members)}.
    # Rows come oldest first, so a later row overwrites an earlier one.
    counts = {}
    if not PARTICIPANT_CSV_URL:
        return counts
    try:
        if PARTICIPANT_CSV_URL.startswith("http"):
            resp = requests.get(PARTICIPANT_CSV_URL, headers=HEADERS, timeout=20)
            resp.raise_for_status()
            text = resp.content.decode("utf-8-sig")
        else:
            with open(PARTICIPANT_CSV_URL, encoding="utf-8-sig") as f:
                text = f.read()
    except Exception as ex:
        print(f"participants: failed to load responses ({ex})")
        return counts
    try:
        rows = list(csv.DictReader(io.StringIO(text)))
    except Exception as ex:
        print(f"participants: failed to parse responses ({ex})")
        return counts
    if not rows:
        return counts
    headers = list(rows[0].keys())

    def _find_col(*needles):
        for h in headers:
            hn = (h or "").strip()
            if any(n in hn for n in needles):
                return h
        return None

    key_col = _find_col("行事キー", "キー")
    total_col = _find_col("総参加")
    if total_col is None:
        total_col = _find_col("参加人数", "参加数")
    nonmem_col = _find_col("会員外")
    if key_col is None or total_col is None:
        print("participants: key/total column not found in responses")
        return counts

    def _num(row, col):
        if col is None:
            return 0
        raw = (row.get(col) or "").strip().replace(",", "").replace("人", "")
        try:
            return int(float(raw))
        except ValueError:
            return None

    for row in rows:
        k = (row.get(key_col) or "").strip()
        total = _num(row, total_col)
        if not k or total is None:
            continue
        nonmem = _num(row, nonmem_col)
        counts[k] = (total, nonmem if nonmem is not None else 0)
    print(f"participants: {len(counts)} events have counts")
    return counts


def generate_html(events, updated_str, week_start, week_end, period_start, period_end, availability=None, radio_monday=None, radio_filename=None):
    events_json = json.dumps(events, ensure_ascii=False)
    availability_json = json.dumps(availability or {}, ensure_ascii=False)
    double_use_pairs = find_double_use(events)
    double_use_card = ""
    if double_use_pairs:
        print(f"会場の二重使用チェック: {len(double_use_pairs)}組を検出しました。")

        def _fmt_min(mins):
            return f"{mins // 60}:{mins % 60:02d}"

        dup_items = []
        for p in double_use_pairs:
            a, b = p["a"], p["b"]
            d_obj = datetime.date.fromisoformat(p["date"])
            venue_label = "・".join(p["venues"])
            ov_label = _fmt_min(p["overlap"][0]) + "～" + _fmt_min(p["overlap"][1])
            dup_items.append(
                '<div class="dup-item">'
                f'<div class="dup-date">{d_obj.month}月{d_obj.day}日（{p["weekday"]}） '
                f'<span class="badge badge-dup">{_html_escape(venue_label)}</span> '
                f'<span class="dup-ov">重なり {ov_label}</span></div>'
                f'<div class="dup-ev">{_html_escape(a.get("time", ""))}　{_html_escape(a.get("title", ""))}（{_html_escape(a.get("org", ""))}） <a class="btn-link" href="{_html_escape(a.get("url", ""))}" target="_blank" rel="noopener">公式で確認</a></div>'
                f'<div class="dup-ev">{_html_escape(b.get("time", ""))}　{_html_escape(b.get("title", ""))}（{_html_escape(b.get("org", ""))}） <a class="btn-link" href="{_html_escape(b.get("url", ""))}" target="_blank" rel="noopener">公式で確認</a></div>'
                '</div>'
            )
        double_use_card = (
            '<button type="button" class="alert-chip dup" onclick="document.getElementById(\'dupModal\').classList.add(\'open\')">'
            f'<span>⚠ 会場の二重使用チェック</span><span class="chip-open">{len(double_use_pairs)}組あります（クリックで詳細）</span></button>'
            '<div class="modal-overlay" id="dupModal" onclick="if(event.target===this)this.classList.remove(\'open\')">'
            '<div class="modal-box">'
            '<div class="modal-head">'
            f'<h2 class="h2-dup">⚠ 会場の二重使用チェック（{len(double_use_pairs)}組）</h2>'
            '<button type="button" class="modal-close" onclick="document.getElementById(\'dupModal\').classList.remove(\'open\')">✕ 閉じる</button></div>'
            '<p class="dup-note">同じ日・同じ会場で時間帯が重なっている予定です。登録ミスの可能性があるため、元の申込内容を確認してください。'
            '同じ団体がイベントコーナーと会議・セミナー室を同時に使う場合は会場が異なるため、ここには含まれません。</p>'
            + "".join(dup_items)
            + '</div></div>'
        )
    org_latest = {}
    for e in events:
        _oname = e.get("org", "")
        if _oname:
            _cur = org_latest.get(_oname)
            if _cur is None or e["date"] >= _cur[0]:
                org_latest[_oname] = (e["date"], e.get("url", ""))
    org_issues = find_org_name_issues(events)
    org_check_card = ""
    if org_issues:
        print(f"団体名チェック: {len(org_issues)}件の要確認候補を検出しました。")
        org_items = []
        for iss in org_issues:
            lines = []
            for n in iss["names"]:
                ld = iss["last"].get(n, "")
                ld_label = ""
                if ld:
                    d2 = datetime.date.fromisoformat(ld)
                    ld_label = f"・最終 {d2.month}/{d2.day}"
                _ourl = org_latest.get(n, ("", ""))[1]
                _olink = f' <a class="btn-link" href="{_html_escape(_ourl)}" target="_blank" rel="noopener">公式で確認</a>' if _ourl else ""
                lines.append(f'<div class="dup-ev">{_html_escape(n)}（{iss["counts"][n]}件{ld_label}）{_olink}</div>')
            org_items.append(
                '<div class="dup-item org-item">'
                f'<div class="dup-date"><span class="badge badge-org">{iss["kind"]}</span></div>'
                + "".join(lines)
                + '</div>'
            )
        org_check_card = (
            '<button type="button" class="alert-chip org" onclick="document.getElementById(\'orgModal\').classList.add(\'open\')">'
            f'<span>団体名チェック（要確認）</span><span class="chip-open">{len(org_issues)}件あります（クリックで詳細）</span></button>'
            '<div class="modal-overlay" id="orgModal" onclick="if(event.target===this)this.classList.remove(\'open\')">'
            '<div class="modal-box">'
            '<div class="modal-head">'
            f'<h2 class="h2-org">団体名チェック（要確認 {len(org_issues)}件）</h2>'
            '<button type="button" class="modal-close" onclick="document.getElementById(\'orgModal\').classList.remove(\'open\')">✕ 閉じる</button></div>'
            '<p class="org-note">似ている団体名の候補です。間違いと決まったものではありません（別の団体のこともあります）。'
            '申込書や団体登録の正式名称と照合してください。表記ゆれは、件数の多い表記への統一が目安です。「公式で確認」は、その名前で直近の行事の公式ページを開きます。</p>'
            + "".join(org_items)
            + '</div></div>'
        )
    miss_fields = [("fee", "料金"), ("org", "団体名"), ("time", "時間"), ("title", "タイトル")]
    miss_counts = {label: 0 for _, label in miss_fields}
    miss_events = []
    for e in events:
        miss = [label for key, label in miss_fields if not (e.get(key) or "").strip()]
        if miss:
            for label in miss:
                miss_counts[label] += 1
            miss_events.append((e, miss))
    missing_card = ""
    if miss_events:
        print(f"missing-field check: {len(miss_events)} events have empty fields")
        miss_summary = "・".join(f"{label}{miss_counts[label]}件" for _, label in miss_fields if miss_counts[label])
        miss_items = []
        for e, miss in sorted(miss_events, key=lambda x: (x[0]["date"], x[0]["time"], x[0]["id"])):
            dd = datetime.date.fromisoformat(e["date"])
            badges = "".join(f'<span class="badge badge-miss">{_html_escape(label)}</span> ' for label in miss)
            miss_items.append(
                '<div class="dup-item miss-item">'
                f'<div class="dup-date">{dd.month}月{dd.day}日（{e.get("weekday", "")}） {badges}</div>'
                f'<div class="dup-ev">{_html_escape(e.get("time", ""))}　{_html_escape(e.get("title", ""))}（{_html_escape(e.get("org", ""))}） <a class="btn-link" href="{_html_escape(e.get("url", ""))}" target="_blank" rel="noopener">公式で確認</a></div>'
                '</div>'
            )
        missing_card = (
            '<button type="button" class="alert-chip miss" onclick="document.getElementById(\'missModal\').classList.add(\'open\')">'
            f'<span>未入力チェック</span><span class="chip-open">{len(miss_events)}件の行事に空欄があります（クリックで詳細）</span></button>'
            '<div class="modal-overlay" id="missModal" onclick="if(event.target===this)this.classList.remove(\'open\')">'
            '<div class="modal-box">'
            '<div class="modal-head">'
            f'<h2 class="h2-miss">未入力チェック（{len(miss_events)}件）</h2>'
            '<button type="button" class="modal-close" onclick="document.getElementById(\'missModal\').classList.remove(\'open\')">✕ 閉じる</button></div>'
            f'<p class="miss-note">項目別の空欄件数：{miss_summary}。バッジは空欄の項目を示します。元の申込内容を確認して、管理システム側で補ってください。</p>'
            + "".join(miss_items)
            + '</div></div>'
        )
    participants = fetch_participants()
    csv_ready = bool(PARTICIPANT_CSV_URL)
    form_ready = bool(PARTICIPANT_FORM_BASE and all(PARTICIPANT_ENTRY.values()))
    monthly_stats = {}
    for e in events:
        ym = e["date"][:7]
        st = monthly_stats.setdefault(ym, {"total": 0, "event": 0, "seminar": 0, "both": 0, "other": 0, "cancelled": 0, "participants": 0, "nonmembers": 0, "entered": 0, "orgs": set()})
        st["total"] += 1
        room = e.get("room", "")
        if room == "イベントコーナー・会議セミナー室":
            st["both"] += 1
        elif room == "イベントコーナー":
            st["event"] += 1
        elif room == "会議セミナー室":
            st["seminar"] += 1
        else:
            st["other"] += 1
        if "中止" in e.get("title", ""):
            st["cancelled"] += 1
        if csv_ready and "中止" not in e.get("title", ""):
            _pc = participants.get(participant_key(e))
            if _pc is not None:
                st["participants"] += _pc[0]
                st["nonmembers"] += _pc[1]
                st["entered"] += 1
        org_name = (e.get("org") or "").strip()
        if org_name:
            st["orgs"].add(org_name)
    stats_rows = []
    for ym in sorted(monthly_stats.keys()):
        st = monthly_stats[ym]
        y_str, m_str = ym.split("-")
        _people_cells = f'<td>{st["participants"]}</td><td>{st["nonmembers"]}</td><td>{st["entered"]}</td>' if csv_ready else ""
        stats_rows.append(
            f'<tr><td class="date-cell">{int(y_str)}年{int(m_str)}月</td>'
            f'<td>{st["total"]}</td><td>{st["event"]}</td><td>{st["seminar"]}</td>'
            f'<td>{st["both"]}</td><td>{st["other"]}</td><td>{len(st["orgs"])}</td>{_people_cells}<td>{st["cancelled"]}</td></tr>'
        )
    stats_people_th = '<th>参加人数</th><th>うち会員外</th><th>人数入力済</th>' if csv_ready else ""
    stats_people_note = '参加人数は入力フォームの回答（行事ごとに最新の回答を採用）の集計です。人数入力済は、参加人数が入力されている行事の数です。' if csv_ready else ""
    stats_card = (
        '<button type="button" class="action-card stats" onclick="document.getElementById(\'statsModal\').classList.add(\'open\')">'
        '📊 月次利用サマリ<span>クリックで月別の利用状況を開きます</span></button>'
        '<div class="modal-overlay" id="statsModal" onclick="if(event.target===this)this.classList.remove(\'open\')">'
        '<div class="modal-box wide">'
        '<div class="modal-head">'
        '<h2 class="h2-stats">月次利用サマリ</h2>'
        '<button type="button" class="modal-close" onclick="document.getElementById(\'statsModal\').classList.remove(\'open\')">✕ 閉じる</button></div>'
        '<p class="stats-note">現在の掲載データからの自動集計です。行事数は中止分も含みます。利用団体数は団体名の種類数（空欄を除く）です。'
        + stats_people_note
        + '</p>'
        '<div class="table-wrap"><table><thead><tr>'
        '<th>月</th><th>行事数</th><th>イベントコーナー</th><th>会議セミナー室</th><th>両方</th><th>その他</th><th>利用団体数</th>'
        + stats_people_th
        + '<th>中止</th>'
        '</tr></thead><tbody>'
        + "".join(stats_rows)
        + '</tbody></table></div></div></div>'
    )
    people_card = ""
    if form_ready:
        try:
            run_date = datetime.date.fromisoformat(updated_str[:10])
        except Exception:
            run_date = week_start
        run_ym = run_date.strftime("%Y-%m")

        def _qr_svg(link):
            if segno is None:
                return ""
            try:
                return segno.make(link, error="m").svg_inline(scale=3, border=1, dark="#1e293b")
            except Exception:
                return ""

        qr_end = run_date + datetime.timedelta(days=7)
        qr_events = [e for e in events
                     if run_date.isoformat() <= e["date"] <= qr_end.isoformat()
                     and "中止" not in e.get("title", "")]
        qr_events.sort(key=lambda e: (e["date"], e.get("time", ""), e["id"]))
        qr_blocks = []
        _cur_date = None
        for e in qr_events:
            if e["date"] != _cur_date:
                _cur_date = e["date"]
                d_obj = datetime.date.fromisoformat(e["date"])
                qr_blocks.append(f'<div class="qr-date">{d_obj.month}月{d_obj.day}日（{e["weekday"]}）</div>')
            link = participant_form_link(e)
            _pc = participants.get(participant_key(e))
            if _pc is not None:
                status = f'<span class="qr-done">入力済み：{_pc[0]}人（うち会員外{_pc[1]}人）</span>'
            else:
                status = f'<a class="btn-link" href="{_html_escape(link)}" target="_blank" rel="noopener">入力フォームを開く</a>'
            qr_blocks.append(
                '<div class="qr-item">'
                f'<div class="qr-fig">{_qr_svg(link)}</div>'
                '<div class="qr-meta">'
                f'<div class="qr-ev">{_html_escape(e.get("time", ""))}　{_html_escape(e.get("title", ""))}</div>'
                f'<div class="qr-org">{_html_escape(e.get("org", ""))}／{_html_escape(e.get("room", ""))}　{status}</div>'
                '</div></div>'
            )
        month_done = [e for e in events
                      if e["date"][:7] == run_ym and e["date"] <= run_date.isoformat()
                      and "中止" not in e.get("title", "")]
        month_entered = [e for e in month_done if participant_key(e) in participants]
        month_people = sum(participants[participant_key(e)][0] for e in month_entered)
        month_nonmem = sum(participants[participant_key(e)][1] for e in month_entered)
        past_missing = [e for e in month_done if participant_key(e) not in participants]
        past_missing.sort(key=lambda e: (e["date"], e.get("time", ""), e["id"]))
        pm_items = []
        for e in past_missing[:60]:
            d_obj = datetime.date.fromisoformat(e["date"])
            link = participant_form_link(e)
            pm_items.append(
                '<div class="dup-ev">'
                f'{d_obj.month}/{d_obj.day}（{e["weekday"]}） {_html_escape(e.get("time", ""))}　{_html_escape(e.get("title", ""))}（{_html_escape(e.get("org", ""))}） '
                f'<a class="btn-link" href="{_html_escape(link)}" target="_blank" rel="noopener">人数を入力</a></div>'
            )
        pm_html = ""
        if pm_items:
            pm_html = (
                '<h3 class="people-h3">今月分で、参加人数がまだ入っていない行事（終了分）</h3>'
                + "".join(pm_items)
            )
        people_card = (
            '<button type="button" class="action-card people" onclick="document.getElementById(\'peopleModal\').classList.add(\'open\')">'
            '👥 参加人数の入力・状況<span>QRコードやリンクから参加人数を入力／今月の入力状況</span></button>'
            '<div class="modal-overlay" id="peopleModal" onclick="if(event.target===this)this.classList.remove(\'open\')">'
            '<div class="modal-box wide">'
            '<div class="modal-head">'
            '<h2 class="h2-people">参加人数の入力・状況</h2>'
            '<button type="button" class="modal-close" onclick="document.getElementById(\'peopleModal\').classList.remove(\'open\')">✕ 閉じる</button></div>'
            f'<p class="people-note">今月（{run_date.month}月）の参加人数の合計は <b>{month_people}人</b>（うち会員外 {month_nonmem}人）です。入力済みは、終了した行事 {len(month_done)} 件中 {len(month_entered)} 件です。スタッフが紙の活動報告書から転記する想定です（団体からの直接入力も可能です）。同じ行事に複数の回答があった場合は最新の回答を採用します。</p>'
            '<h3 class="people-h3">直近8日間の行事（QRコード）</h3>'
            '<p class="people-note">QRコードを読み取るか「入力フォームを開く」を押すと、その行事の情報（日付・行事名・団体名）が入ったフォームが開きます。総参加数と会員外人数だけ入力して送信してください。</p>'
            '<div id="peoplePrint">'
            + "".join(qr_blocks)
            + '</div>'
            '<p><button type="button" class="btn-print" onclick="printQrSheet()">🖨 QRシートを印刷用に開く</button></p>'
            + pm_html
            + '</div></div>'
            '<script>function printQrSheet(){var s=document.getElementById("peoplePrint").innerHTML;var w=window.open("","_blank");if(!w){alert("ポップアップがブロックされました。ブラウザの設定で許可してください。");return;}'
            'w.document.write("<!DOCTYPE html><html lang=\\"ja\\"><head><meta charset=\\"utf-8\\"><title>参加人数入力QRシート</title><style>body{font-family:sans-serif;margin:16px}.qr-item{display:flex;gap:12px;align-items:center;border-bottom:1px solid #ccc;padding:8px 0}.qr-fig svg{width:96px;height:96px}.qr-date{font-weight:bold;margin-top:14px}.qr-ev{font-weight:bold}.qr-org{font-size:12px;color:#555}.qr-done{color:#15803d;font-weight:bold}a{display:none}</style></head><body><h1>参加人数入力QRシート</h1>"+s+"</body></html>");w.document.close();}</script>'
        )
    new_groups = {}
    for e in events:
        fs = e.get("firstSeen")
        if fs:
            new_groups.setdefault(fs, []).append(e)
    new_entries_card = ""
    if new_groups:
        total_new = sum(len(v) for v in new_groups.values())
        print(f"new entries: {total_new} events in {len(new_groups)} firstSeen groups")
        group_html = []
        for fs_date in sorted(new_groups.keys(), reverse=True):
            g = sorted(new_groups[fs_date], key=lambda x: (x["date"], x["time"], x["id"]))
            d_fs = datetime.date.fromisoformat(fs_date)
            by_date = {}
            for e in g:
                by_date[e["date"]] = by_date.get(e["date"], 0) + 1
            summary_parts = []
            for dstr in sorted(by_date.keys()):
                dd = datetime.date.fromisoformat(dstr)
                summary_parts.append(f"{dd.month}/{dd.day}\u306b{by_date[dstr]}\u4ef6")
            ev_lines = []
            for e in g:
                dd = datetime.date.fromisoformat(e["date"])
                cancelled_mark = ""
                if "\u4e2d\u6b62" in e.get("title", ""):
                    cancelled_mark = ' <span class="badge badge-dup">\u4e2d\u6b62</span>'
                ev_lines.append(
                    f'<div class="dup-ev">{dd.month}/{dd.day}\uff08{e.get("weekday", "")}\uff09 '
                    f'{_html_escape(e.get("time", ""))}\u3000{_html_escape(e.get("title", ""))}'
                    f'\uff08{_html_escape(e.get("org", ""))}\uff09 [{_html_escape(e.get("room", ""))}]{cancelled_mark} <a class="btn-link" href="{_html_escape(e.get("url", ""))}" target="_blank" rel="noopener">公式で確認</a></div>'
                )
            summary_text = "\u3001".join(summary_parts)
            group_html.append(
                '<div class="dup-item new-item">'
                f'<div class="dup-date">{d_fs.month}\u6708{d_fs.day}\u65e5\u306b\u521d\u78ba\u8a8d\uff08{len(g)}\u4ef6\uff09</div>'
                f'<div class="dup-ev">\u5bfe\u8c61\u65e5: {summary_text}</div>'
                + "".join(ev_lines)
                + '</div>'
            )
        new_entries_card = (
            '<button type="button" class="alert-chip new" onclick="document.getElementById(\'newModal\').classList.add(\'open\')">'
            f'<span>\u65b0\u898f\u5165\u529b\u30c1\u30a7\u30c3\u30af</span><span class="chip-open">{total_new}\u4ef6\u3042\u308a\u307e\u3059\uff08\u30af\u30ea\u30c3\u30af\u3067\u8a73\u7d30\uff09</span></button>'
            '<div class="modal-overlay" id="newModal" onclick="if(event.target===this)this.classList.remove(\'open\')">'
            '<div class="modal-box">'
            '<div class="modal-head">'
            f'<h2 class="h2-new">\u65b0\u898f\u5165\u529b\u30c1\u30a7\u30c3\u30af\uff08{total_new}\u4ef6\uff09</h2>'
            '<button type="button" class="modal-close" onclick="document.getElementById(\'newModal\').classList.remove(\'open\')">\u2715 \u9589\u3058\u308b</button></div>'
            '<p class="new-note">\u65b0\u3057\u304f\u30ab\u30ec\u30f3\u30c0\u30fc\u306b\u73fe\u308c\u305f\u4e88\u5b9a\u3092\u3001\u521d\u3081\u3066\u78ba\u8a8d\u3067\u304d\u305f\u65e5\uff08\u521d\u78ba\u8a8d\u65e5\uff09\u3054\u3068\u306b\u307e\u3068\u3081\u3066\u3044\u307e\u3059\u3002\u624b\u5165\u529b\u76f4\u5f8c\u306e\u898b\u76f4\u3057\u306b\u4f7f\u3063\u3066\u304f\u3060\u3055\u3044\u3002\u8a18\u9332\u306f\u3053\u306e\u6a5f\u80fd\u306e\u5c0e\u5165\u65e5\u304b\u3089\u59cb\u307e\u3063\u3066\u3044\u307e\u3059\u3002</p>'
            + "".join(group_html)
            + '</div></div>'
        )
    radio_label = ""
    radio_card = ""
    radio_note = ""
    if radio_monday:
        wd_jp_r = ["月", "火", "水", "木", "金", "土", "日"]
        radio_label = f"{radio_monday.month}月{radio_monday.day}日（{wd_jp_r[radio_monday.weekday()]}）"
    if radio_monday and radio_filename:
        radio_card = (f'<a class="action-card radio" href="{radio_filename}" download>'
                      f'📻 ラジオ原稿ダウンロード（Word）<span>対象: {radio_label}放送分</span></a>')
        radio_note = (f'<p>※ラジオ原稿は月曜放送分です。毎週水曜17時に翌週分へ切り替わります。'
                      f'今ダウンロードできるのは{radio_label}放送分です。</p>')
    cancel_label = "「" + LB + "中止" + RB + "」の予定を隠す"
    only_label = "「" + LB + "中止" + RB + "」だけを表示"
    wd_jp = ["月", "火", "水", "木", "金", "土", "日"]
    week_label = (f"{week_start.month}月{week_start.day}日（{wd_jp[week_start.weekday()]}）"
                  f"〜{week_end.month}月{week_end.day}日（{wd_jp[week_end.weekday()]}）")
    period_label = (f"{period_start.month}月{period_start.day}日（{wd_jp[period_start.weekday()]}）"
                    f"〜{period_end.month}月{period_end.day}日（{wd_jp[period_end.weekday()]}）")
    html = f"""<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>あいぽーとセンター業務ポータル</title>
<style>
  * {{ box-sizing: border-box; }}
  body {{ font-family: "Helvetica Neue", Arial, "Hiragino Kaku Gothic ProN", "Hiragino Sans", Meiryo, sans-serif; background: #eef2f7; color: #1e293b; margin: 0; padding: 20px 12px 40px; }}
  .container {{ max-width: 1200px; margin: 0 auto; }}
  .app-header {{ background: linear-gradient(135deg, #1e3a8a, #2563eb); color: #fff; border-radius: 14px; padding: 22px 26px; box-shadow: 0 4px 14px rgba(30,58,138,0.25); display: flex; justify-content: space-between; align-items: flex-start; gap: 12px; flex-wrap: wrap; }}
  .app-header h1 {{ margin: 0 0 6px; font-size: 1.45rem; }}
  .updated {{ font-size: 0.85rem; opacity: 0.92; }}
  .subtitle {{ font-size: 0.85rem; opacity: 0.85; margin-bottom: 4px; }}
  .btn-update-sm {{ background: rgba(255,255,255,0.15); border: 1px solid rgba(255,255,255,0.65); color: #fff; padding: 6px 14px; border-radius: 999px; font-size: 0.82rem; font-weight: bold; text-decoration: none; white-space: nowrap; }}
  .btn-update-sm:hover {{ background: rgba(255,255,255,0.3); }}
  .filters-head {{ grid-column: 1 / -1; display: flex; justify-content: space-between; align-items: center; }}
  .filters-head h2 {{ margin: 0; font-size: 1.05rem; color: #1e293b; }}
  .btn-clear {{ padding: 7px 14px; border: 1px solid #cbd5e1; background: #fff; border-radius: 999px; cursor: pointer; font-weight: bold; font-size: 0.85rem; color: #475569; }}
  .btn-clear:hover {{ border-color: #dc2626; color: #dc2626; }}
  .actions {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 12px; margin: 16px 0 12px; }}
  .action-card {{ display: block; background: #fff; border-radius: 12px; padding: 14px 16px; text-decoration: none; color: #1e293b; box-shadow: 0 2px 8px rgba(15,23,42,0.06); border-left: 6px solid #2563eb; font-weight: bold; }}
  .action-card span {{ display: block; margin-top: 4px; font-size: 0.8rem; font-weight: normal; color: #64748b; }}
  .action-card:hover {{ box-shadow: 0 4px 14px rgba(15,23,42,0.14); }}
  .action-card.docx {{ border-left-color: #16a34a; }}
  .action-card.excel {{ border-left-color: #217346; }}
  .action-card.radio {{ border-left-color: #d97706; }}
  .action-card.stats {{ border-left-color: #7c3aed; }}
  button.action-card {{ font: inherit; text-align: left; cursor: pointer; border-top: none; border-right: none; border-bottom: none; width: 100%; }}
  .info-box {{ background: #eff6ff; border: 1px solid #bfdbfe; border-radius: 10px; padding: 10px 14px; margin-bottom: 16px; }}
  .info-box p {{ margin: 3px 0; font-size: 0.8rem; color: #475569; }}
  .alert-chip {{ display: flex; justify-content: space-between; align-items: center; gap: 10px; width: 100%; text-align: left; padding: 11px 16px; border-radius: 10px; margin-bottom: 10px; cursor: pointer; font-weight: bold; font-size: 0.95rem; box-shadow: 0 2px 8px rgba(15,23,42,0.06); }}
  .alert-chip.dup {{ background: #fef2f2; border: 1px solid #fecaca; border-left: 6px solid #dc2626; color: #b91c1c; }}
  .alert-chip.org {{ background: #fffbeb; border: 1px solid #fde68a; border-left: 6px solid #d97706; color: #b45309; }}
  .alert-chip.new {{ background: #eff6ff; border: 1px solid #bfdbfe; border-left: 6px solid #2563eb; color: #1d4ed8; }}
  .h2-new {{ color: #1d4ed8; }}
  .new-note {{ margin: 0 0 10px; font-size: 0.82rem; color: #1e3a8a; }}
  .dup-item.new-item {{ border-color: #bfdbfe; }}
  .stats-note {{ margin: 0 0 10px; font-size: 0.82rem; color: #64748b; }}
  .h2-stats {{ color: #6d28d9; }}
  .action-card.people {{ border-left-color: #059669; }}
  .h2-people {{ color: #059669; }}
  .people-note {{ margin: 0 0 10px; font-size: 0.82rem; color: #334155; }}
  .people-h3 {{ margin: 14px 0 6px; font-size: 0.95rem; color: #065f46; }}
  .qr-date {{ font-weight: bold; margin: 12px 0 4px; color: #1e293b; }}
  .qr-item {{ display: flex; gap: 12px; align-items: center; border: 1px solid #e2e8f0; border-radius: 10px; padding: 8px 10px; margin-bottom: 8px; background: #fff; }}
  .qr-fig svg {{ width: 84px; height: 84px; display: block; }}
  .qr-ev {{ font-weight: bold; }}
  .qr-org {{ font-size: 0.82rem; color: #64748b; margin-top: 2px; }}
  .qr-done {{ color: #15803d; font-weight: bold; }}
  .btn-print {{ font: inherit; cursor: pointer; border: 1px solid #059669; color: #059669; background: #ecfdf5; border-radius: 8px; padding: 6px 12px; }}
  .modal-box.wide {{ max-width: 920px; }}
  .alert-chip.miss {{ background: #f0fdfa; border: 1px solid #99f6e4; border-left: 6px solid #0d9488; color: #0f766e; }}
  .h2-miss {{ color: #0f766e; }}
  .miss-note {{ margin: 0 0 10px; font-size: 0.82rem; color: #134e4a; }}
  .dup-item.miss-item {{ border-color: #99f6e4; }}
  .badge-miss {{ background: #ccfbf1; color: #0f766e; }}
  .alert-chip:hover {{ box-shadow: 0 4px 14px rgba(15,23,42,0.12); }}
  .chip-open {{ font-size: 0.82rem; white-space: nowrap; }}
  .modal-overlay {{ display: none; position: fixed; inset: 0; background: rgba(15,23,42,0.55); z-index: 50; padding: 20px 12px; }}
  .modal-overlay.open {{ display: block; }}
  .modal-box {{ background: #fff; border-radius: 12px; max-width: 760px; margin: 0 auto; max-height: calc(100vh - 40px); overflow-y: auto; padding: 18px 20px; }}
  .modal-head {{ display: flex; justify-content: space-between; align-items: center; gap: 10px; margin-bottom: 6px; }}
  .modal-head h2 {{ margin: 0; font-size: 1.1rem; }}
  .h2-dup {{ color: #b91c1c; }}
  .h2-org {{ color: #b45309; }}
  .modal-close {{ padding: 7px 14px; border: 1px solid #cbd5e1; background: #fff; border-radius: 999px; cursor: pointer; font-weight: bold; font-size: 0.85rem; color: #475569; white-space: nowrap; }}
  .modal-close:hover {{ border-color: #dc2626; color: #dc2626; }}
  .dup-note {{ margin: 0 0 10px; font-size: 0.82rem; color: #7f1d1d; }}
  .dup-item {{ background: #fff; border: 1px solid #fecaca; border-radius: 8px; padding: 10px 12px; margin-bottom: 8px; }}
  .dup-item:last-child {{ margin-bottom: 0; }}
  .dup-date {{ font-weight: bold; color: #1e293b; margin-bottom: 4px; }}
  .badge-dup {{ background: #fee2e2; color: #b91c1c; }}
  .dup-ov {{ color: #b91c1c; font-weight: bold; font-size: 0.85rem; }}
  .dup-ev {{ font-size: 0.9rem; color: #334155; padding: 2px 0; }}
  .org-note {{ margin: 0 0 10px; font-size: 0.82rem; color: #78350f; }}
  .dup-item.org-item {{ border-color: #fde68a; }}
  .badge-org {{ background: #fef3c7; color: #b45309; }}
  .card {{ background: #fff; border-radius: 12px; box-shadow: 0 2px 8px rgba(15,23,42,0.06); }}
  .filters {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(210px, 1fr)); gap: 12px; padding: 16px; margin-bottom: 16px; }}
  .filter-group label {{ display: block; font-size: 0.85rem; font-weight: bold; margin-bottom: 6px; color: #475569; }}
  .filter-group input[type="text"], .filter-group input[type="date"] {{ width: 100%; padding: 9px; border: 1px solid #cbd5e1; border-radius: 8px; font-size: 0.95rem; }}
  .check-group {{ display: flex; align-items: flex-end; gap: 14px; padding-bottom: 8px; flex-wrap: wrap; }}
  .check-group label {{ display: inline-flex; align-items: center; gap: 6px; font-size: 0.9rem; margin: 0; }}
  .check-group input {{ width: 17px; height: 17px; }}
  .tabs {{ display: flex; gap: 8px; margin-bottom: 12px; flex-wrap: wrap; }}
  .tab-btn {{ padding: 9px 16px; border: 1px solid #cbd5e1; background: #fff; color: #475569; border-radius: 999px; cursor: pointer; font-weight: bold; font-size: 0.9rem; transition: 0.2s; }}
  .tab-btn:hover {{ border-color: #2563eb; color: #2563eb; }}
  .tab-btn.active {{ background: #2563eb; border-color: #2563eb; color: #fff; }}
  .count-bar {{ font-size: 1.05rem; font-weight: bold; margin: 0 2px 10px; color: #1e293b; }}
  .table-wrap {{ overflow-x: auto; border-radius: 12px; }}
  table {{ width: 100%; border-collapse: separate; border-spacing: 0; font-size: 0.95rem; background: #fff; }}
  th, td {{ padding: 11px 12px; border-bottom: 1px solid #e2e8f0; text-align: left; vertical-align: top; }}
  thead th {{ background: #f1f5f9; color: #334155; white-space: nowrap; position: sticky; top: 0; z-index: 2; border-bottom: 2px solid #cbd5e1; }}
  tbody tr:nth-child(even) {{ background: #f8fafc; }}
  tbody tr:hover {{ background: #eff6ff; }}
  tr.date-start td {{ border-top: 2px solid #cbd5e1; }}
  tr.cancelled {{ background: #fef2f2 !important; }}
  .date-cell {{ white-space: nowrap; font-weight: bold; }}
  .wd-sat {{ color: #2563eb; }}
  .wd-sun {{ color: #dc2626; }}
  .time-cell {{ white-space: nowrap; color: #334155; }}
  .title-cell {{ font-weight: bold; color: #0f172a; min-width: 180px; }}
  .org-cell {{ color: #475569; }}
  .badge {{ display: inline-block; padding: 4px 9px; border-radius: 999px; font-size: 0.78rem; font-weight: bold; white-space: nowrap; }}
  .badge-event {{ background: #fef3c7; color: #b45309; }}
  .badge-seminar {{ background: #dbeafe; color: #1d4ed8; }}
  .badge-both {{ background: #ede9fe; color: #6d28d9; }}
  .badge-other {{ background: #f1f5f9; color: #64748b; }}
  a.btn-link {{ display: inline-block; padding: 5px 12px; background: #fff; border: 1px solid #2563eb; color: #2563eb; text-decoration: none; border-radius: 999px; font-size: 0.82rem; font-weight: bold; white-space: nowrap; }}
  a.btn-link:hover {{ background: #2563eb; color: #fff; }}
  .empty {{ text-align: center; color: #94a3b8; padding: 36px 12px; font-size: 1rem; }}
  .avail-card {{ padding: 18px; margin-bottom: 16px; }}
  .avail-head {{ display: flex; justify-content: space-between; align-items: center; flex-wrap: wrap; gap: 10px; }}
  .avail-head h2 {{ margin: 0; font-size: 1.15rem; color: #1e293b; }}
  .avail-controls {{ display: flex; align-items: center; gap: 8px; }}
  .avail-controls input[type="date"] {{ padding: 8px; border: 1px solid #cbd5e1; border-radius: 8px; font-size: 0.95rem; }}
  .avail-nav {{ padding: 8px 12px; border: 1px solid #cbd5e1; background: #fff; border-radius: 8px; cursor: pointer; font-weight: bold; color: #475569; }}
  .avail-nav:hover {{ border-color: #2563eb; color: #2563eb; }}
  .avail-date-label {{ font-size: 1.05rem; font-weight: bold; margin: 12px 0 8px; color: #1e293b; }}
  .avail-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(300px, 1fr)); gap: 16px; }}
  .room-panel {{ border: 1px solid #e2e8f0; border-radius: 10px; padding: 14px; }}
  .room-panel h3 {{ margin: 0 0 10px; }}
  .room-panel h4 {{ margin: 14px 0 4px; font-size: 0.9rem; color: #475569; }}
  .timeline {{ display: flex; height: 20px; border-radius: 6px; overflow: hidden; border: 1px solid #cbd5e1; background: #dcfce7; }}
  .seg-booked {{ background: #f87171; height: 100%; }}
  .tl-scale {{ position: relative; height: 16px; font-size: 0.7rem; color: #94a3b8; margin-top: 2px; }}
  .tl-scale span {{ position: absolute; transform: translateX(-50%); }}
  .tl-scale span:first-child {{ transform: none; }}
  .tl-scale span:last-child {{ transform: none; }}
  .slot-list {{ margin-top: 8px; display: flex; flex-wrap: wrap; gap: 6px; }}
  .slot {{ display: inline-block; padding: 4px 10px; border-radius: 999px; font-size: 0.82rem; font-weight: bold; }}
  .slot.booked {{ background: #fee2e2; color: #b91c1c; }}
  .slot.free {{ background: #dcfce7; color: #15803d; }}
  .slot.nodata {{ background: #f1f5f9; color: #64748b; }}
  .day-ev {{ padding: 6px 0; border-bottom: 1px dashed #e2e8f0; font-size: 0.9rem; }}
  .day-ev:last-child {{ border-bottom: none; }}
  .day-ev-time {{ font-weight: bold; color: #334155; margin-right: 8px; white-space: nowrap; }}
  .day-ev-org {{ color: #64748b; font-size: 0.82rem; }}
  .day-ev-none {{ color: #94a3b8; font-size: 0.9rem; }}
  @media (max-width: 640px) {{ .app-header {{ padding: 18px; }} .app-header h1 {{ font-size: 1.2rem; }} th, td {{ padding: 9px 8px; }} }}
</style>
</head>
<body>
<div class="container">
  <header class="app-header">
    <div>
      <h1>あいぽーとセンター業務ポータル</h1>
      <div class="subtitle">予定検索・空き状況確認・各種出力（週報／あいぽーと通信／ラジオ原稿）</div>
      <div class="updated">最終自動更新: {updated_str} （毎日1回自動更新）</div>
    </div>
    <a class="btn-update-sm" href="https://github.com/aiport0960168/aiport-calendar2/actions/workflows/update_data.yml" target="_blank" rel="noopener">🔄 最新情報に更新</a>
  </header>

  <div class="actions">
    <a class="action-card docx" href="weekly.docx" download>📄 週報ダウンロード（Word）<span>対象: {week_label}分</span></a>
    <a class="action-card excel" href="newsletter.xlsx" download>📊 あいぽーと通信用カレンダー（Excel）<span>対象: {period_label}分</span></a>
    {radio_card}
    {stats_card}
    {people_card}
  </div>

  <div class="info-box">
    <p>※データは毎朝6時00分に自動更新されます。手動で更新する場合は右上の「最新情報に更新」から、GitHubのページで「Run workflow」を押してください。</p>
    <p>※週報は今週の月曜〜日曜の7日分（1日1ページ）です。木曜17時以降は翌週分に切り替わります。</p>
    <p>※あいぽーと通信用カレンダーは16日〜翌月15日分です。毎月15日17時に、翌月16日〜翌々月15日分へ切り替わります。</p>
    {radio_note}
  </div>

  {double_use_card}

  {org_check_card}

  {new_entries_card}

  {missing_card}

  <div class="card avail-card">
    <div class="avail-head">
      <h2>🗓 空き状況（日付で確認）</h2>
      <div class="avail-controls">
        <button type="button" class="avail-nav" id="availPrev">◀ 前日</button>
        <button type="button" class="avail-nav" id="availToday">今日</button>
        <input type="date" id="availDate">
        <button type="button" class="avail-nav" id="availNext">翌日 ▶</button>
      </div>
    </div>
    <div class="avail-date-label" id="availDateLabel"></div>
    <div class="avail-grid">
      <div class="room-panel">
        <h3><span class="badge badge-event">イベントコーナー</span></h3>
        <div class="timeline" id="tlEvent"></div>
        <div class="tl-scale"><span style="left:0%">8:30</span><span style="left:28%">12:00</span><span style="left:52%">15:00</span><span style="left:76%">18:00</span><span style="right:0">21:00</span></div>
        <div class="slot-list" id="slotsEvent"></div>
        <h4>この日の行事</h4>
        <div id="dayEventsEvent"></div>
      </div>
      <div class="room-panel">
        <h3><span class="badge badge-seminar">会議・セミナー室</span></h3>
        <div class="timeline" id="tlSeminar"></div>
        <div class="tl-scale"><span style="left:0%">8:30</span><span style="left:28%">12:00</span><span style="left:52%">15:00</span><span style="left:76%">18:00</span><span style="right:0">21:00</span></div>
        <div class="slot-list" id="slotsSeminar"></div>
        <h4>この日の行事</h4>
        <div id="dayEventsSeminar"></div>
      </div>
    </div>
  </div>

  <div class="filters card">
    <div class="filters-head">
      <h2>🔍 行事を検索</h2>
      <button type="button" class="btn-clear" id="clearFilters">条件をクリア</button>
    </div>
    <div class="filter-group">
      <label>開始日</label>
      <input type="date" id="startDate">
    </div>
    <div class="filter-group">
      <label>終了日</label>
      <input type="date" id="endDate">
    </div>
    <div class="filter-group">
      <label>キーワード検索（タイトル・団体名・時間）</label>
      <input type="text" id="keyword" placeholder="スペース区切りで複数検索（例: 囲碁 相談）">
    </div>
    <div class="filter-group check-group">
      <label><input type="checkbox" id="hideCancel" checked> {cancel_label}</label>
      <label><input type="checkbox" id="onlyCancel"> {only_label}</label>
    </div>
  </div>

  <div class="tabs">
    <button class="tab-btn active" onclick="setTab('all', this)">すべて</button>
    <button class="tab-btn" onclick="setTab('イベントコーナー', this)">イベントコーナー</button>
    <button class="tab-btn" onclick="setTab('会議セミナー室', this)">会議セミナー室</button>
    <button class="tab-btn" onclick="setTab('イベントコーナー・会議セミナー室', this)">イベントコーナー・会議セミナー室</button>
    <button class="tab-btn" onclick="setTab('その他・未記載', this)">その他</button>
  </div>

  <div class="count-bar" id="countBar"></div>

  <div class="table-wrap card">
    <table>
      <thead>
        <tr>
          <th>日付</th>
          <th>時間</th>
          <th>場所区分</th>
          <th>タイトル</th>
          <th>団体名</th>
          <th>詳細</th>
        </tr>
      </thead>
      <tbody id="tableBody"></tbody>
    </table>
  </div>
</div>

<script>
const events = {events_json};
let currentTab = 'all';

// startDate defaults to today in the viewer's local timezone (JST-safe)
const _d = new Date();
const todayStr = _d.getFullYear() + '-' + String(_d.getMonth() + 1).padStart(2, '0') + '-' + String(_d.getDate()).padStart(2, '0');
document.getElementById('startDate').value = todayStr;

function setTab(tab, btn) {{
  currentTab = tab;
  document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
  btn.classList.add('active');
  render();
}}

function getBadgeClass(room) {{
  if (room.includes('イベント') && room.includes('会議')) return 'badge-both';
  if (room.includes('イベント')) return 'badge-event';
  if (room.includes('会議')) return 'badge-seminar';
  return 'badge-other';
}}

function render() {{
  const start = document.getElementById('startDate').value;
  const end = document.getElementById('endDate').value;
  const kwRaw = document.getElementById('keyword').value.trim().toLowerCase();
  const keywords = kwRaw ? kwRaw.split(/\\s+/) : [];
  const hideCancel = document.getElementById('hideCancel').checked;
  const onlyCancel = document.getElementById('onlyCancel').checked;

  const filtered = events.filter(ev => {{
    if (start && ev.date < start) return false;
    if (end && ev.date > end) return false;
    if (currentTab !== 'all' && !ev.room.includes(currentTab)) return false;
    if (onlyCancel && !ev.title.includes('中止')) return false;
    if (!onlyCancel && hideCancel && ev.title.includes('中止')) return false;
    if (keywords.length > 0) {{
      const target = (ev.title + ' ' + ev.org + ' ' + ev.time + ' ' + ev.place).toLowerCase();
      if (!keywords.every(k => target.includes(k))) return false;
    }}
    return true;
  }});

  document.getElementById('countBar').textContent = '検索結果: ' + filtered.length + ' 件';

  const tbody = document.getElementById('tableBody');
  let rows = '';
  let prevDate = '';
  for (const ev of filtered) {{
    const parts = ev.date.split('-');
    const dateLabel = parseInt(parts[1]) + '月' + parseInt(parts[2]) + '日';
    const wdClass = ev.weekday === '土' ? 'wd-sat' : (ev.weekday === '日' ? 'wd-sun' : '');
    let cls = '';
    if (ev.date !== prevDate) cls += ' date-start';
    if (ev.title.includes('中止')) cls += ' cancelled';
    prevDate = ev.date;
    rows += '<tr' + (cls ? ' class="' + cls.trim() + '"' : '') + '>'
      + '<td class="date-cell">' + dateLabel + ' (<span class="' + wdClass + '">' + ev.weekday + '</span>)</td>'
      + '<td class="time-cell">' + ev.time + '</td>'
      + '<td><span class="badge ' + getBadgeClass(ev.room) + '">' + ev.room + '</span></td>'
      + '<td class="title-cell">' + ev.title + '</td>'
      + '<td class="org-cell">' + ev.org + '</td>'
      + '<td><a class="btn-link" href="' + ev.url + '" target="_blank" rel="noopener">公式詳細</a></td>'
      + '</tr>';
  }}
  if (!rows) rows = '<tr><td colspan="6" class="empty">該当する予定がありません</td></tr>';
  tbody.innerHTML = rows;
}}

document.querySelectorAll('#startDate, #endDate, #keyword, #hideCancel, #onlyCancel').forEach(el => {{
  el.addEventListener('input', render);
  el.addEventListener('change', render);
}});

render();

const availability = {availability_json};
const AVAIL_START = 510;
const AVAIL_END = 1260;

function toMin(hhmm) {{
  const p = hhmm.split(':');
  return parseInt(p[0]) * 60 + parseInt(p[1]);
}}

function shiftAvailDate(delta) {{
  const inp = document.getElementById('availDate');
  if (!inp.value) return;
  const d = new Date(inp.value + 'T00:00:00');
  d.setDate(d.getDate() + delta);
  inp.value = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  renderAvail();
}}

function renderRoom(suffix, roomKey, dateStr) {{
  const tl = document.getElementById('tl' + suffix);
  const slotsEl = document.getElementById('slots' + suffix);
  const evEl = document.getElementById('dayEvents' + suffix);
  const day = availability[dateStr];
  if (!day) {{
    tl.innerHTML = '';
    slotsEl.innerHTML = '<span class="slot nodata">空き状況データがありません</span>';
  }} else {{
    const slots = day[roomKey] || [];
    let pos = AVAIL_START;
    let segs = '';
    for (const s of slots) {{
      const ab = s.split('～');
      const st = Math.max(toMin(ab[0]), AVAIL_START);
      const en = Math.min(toMin(ab[1]), AVAIL_END);
      if (en <= st) continue;
      if (st > pos) segs += '<div style="width:' + ((st - pos) / (AVAIL_END - AVAIL_START) * 100) + '%"></div>';
      segs += '<div class="seg-booked" title="予約済 ' + s + '" style="width:' + ((en - st) / (AVAIL_END - AVAIL_START) * 100) + '%"></div>';
      pos = en;
    }}
    tl.innerHTML = segs;
    slotsEl.innerHTML = slots.length
      ? slots.map(s => '<span class="slot booked">予約済 ' + s + '</span>').join('')
      : '<span class="slot free">予約はありません（全日空き）</span>';
  }}
  const matchKey = roomKey === 'イベントコーナー' ? 'イベント' : '会議';
  const evs = events.filter(ev => ev.date === dateStr && ev.room.includes(matchKey));
  evEl.innerHTML = evs.length
    ? evs.map(ev => '<div class="day-ev"><span class="day-ev-time">' + ev.time + '</span>' + ev.title + ' <span class="day-ev-org">' + ev.org + '</span></div>').join('')
    : '<div class="day-ev-none">行事の登録はありません</div>';
}}

function renderAvail() {{
  const dateStr = document.getElementById('availDate').value;
  if (!dateStr) return;
  const d = new Date(dateStr + 'T00:00:00');
  const wd = ['日', '月', '火', '水', '木', '金', '土'][d.getDay()];
  document.getElementById('availDateLabel').textContent = (d.getMonth() + 1) + '月' + d.getDate() + '日（' + wd + '）の空き状況と行事';
  renderRoom('Event', 'イベントコーナー', dateStr);
  renderRoom('Seminar', '会議セミナー室', dateStr);
}}

document.getElementById('availDate').value = todayStr;
document.getElementById('availDate').addEventListener('change', renderAvail);
document.getElementById('availPrev').addEventListener('click', () => shiftAvailDate(-1));
document.getElementById('availNext').addEventListener('click', () => shiftAvailDate(1));
document.getElementById('availToday').addEventListener('click', () => {{ document.getElementById('availDate').value = todayStr; renderAvail(); }});
renderAvail();

function clearFilters() {{
  document.getElementById('startDate').value = todayStr;
  document.getElementById('endDate').value = '';
  document.getElementById('keyword').value = '';
  document.getElementById('hideCancel').checked = true;
  document.getElementById('onlyCancel').checked = false;
  currentTab = 'all';
  document.querySelectorAll('.tab-btn').forEach((b, i) => b.classList.toggle('active', i === 0));
  render();
}}
document.getElementById('clearFilters').addEventListener('click', clearFilters);
</script>
</body>
</html>"""
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)


def _normalize_time(time_str):
    """「13時～15時」→「13時00分～15時00分」のように、分なしの時刻に00分を補う"""
    return re.sub(r"(\d+)時(?![\d半])", r"\1時00分", time_str or "")


def _disp_width(text):
    return sum(1.0 if unicodedata.east_asian_width(c) in ("W", "F") else 0.5 for c in text)


def _balanced_lines(text, body_size, label_units=6):
    """長い文字列を行末に孤立文字が出ないようバランスよく分ける"""
    cap = 139 / (body_size * 0.3528) - label_units
    if _disp_width(text) <= cap:
        return [text]
    no_head = set("」』）)］｝、。．，．・：；？！…―〉》々ゃゅょっぁぃぅぇぉゎー")
    no_tail = set("「『（(［｛〈《")
    natural_after = set("・、。，．／/ 」』）)】")
    natural_before = set("「『（(【")
    chars = list(text)
    widths = []
    w = 0.0
    for c in chars:
        w += 1.0 if unicodedata.east_asian_width(c) in ("W", "F") else 0.5
        widths.append(w)
    total = widths[-1]
    min_tail = 6.0
    target = total / 2

    def find_split(use_kinsoku):
        best_i, best_score = None, None
        for i in range(1, len(chars)):
            w1 = widths[i - 1]
            w2 = total - w1
            if w1 > cap or w2 < min_tail:
                continue
            if use_kinsoku and (chars[i] in no_head or chars[i - 1] in no_tail):
                continue
            score = abs(w1 - target)
            if chars[i - 1] not in natural_after and chars[i] not in natural_before:
                score += 3.0
            if best_score is None or score < best_score:
                best_score, best_i = score, i
        return best_i

    i = find_split(True)
    if i is None:
        i = find_split(False)
    if i is None:
        return [text]
    line1 = text[:i]
    rest = text[i:]
    if _disp_width(rest) > cap:
        return [line1] + _balanced_lines(rest, body_size, 0)
    return [line1, rest]


def _start_hour(time_str):
    m = re.search(r"(\d+)\s*時", time_str or "")
    return int(m.group(1)) if m else 99


def main():
    jst = datetime.timezone(datetime.timedelta(hours=9))
    now = datetime.datetime.now(jst)
    months_count = get_months_to_fetch(now)

    cache = {}
    if os.path.exists(CACHE_PATH):
        try:
            with open(CACHE_PATH, "r", encoding="utf-8") as f:
                cache = json.load(f)
        except Exception:
            cache = {}

    session = requests.Session()
    all_events = []
    seen_keys = set()

    curr_year = now.year
    curr_month = now.month
    weekdays = ["月", "火", "水", "木", "金", "土", "日"]

    for _ in range(months_count):
        params = {"c_id": 15, "sely": curr_year, "selm": curr_month}
        print(f"{curr_year}年{curr_month}月のカレンダーを取得中...")

        try:
            res = session.get(BASE_URL, params=params, headers=HEADERS, timeout=20)
            res.encoding = "utf-8"
            soup = BeautifulSoup(res.text, "html.parser")
        except Exception as e:
            print(f"カレンダー取得失敗 ({curr_year}/{curr_month}): {e}")
            break

        cal_table = soup.find("table", class_="table1")
        if cal_table:
            links = cal_table.find_all("a", href=re.compile(r"detail\.aspx"))
            for a in links:
                href = a.get("href", "")
                detail_url = urljoin(BASE_URL, href)

                qs = parse_qs(urlparse(detail_url).query)
                event_id = qs.get("id", [""])[0]
                y = int(qs.get("sely", [curr_year])[0])
                m = int(qs.get("selm", [curr_month])[0])
                d = int(qs.get("seld", [1])[0])

                date_obj = datetime.date(y, m, d)
                date_str = date_obj.strftime("%Y-%m-%d")
                wd_str = weekdays[date_obj.weekday()]
                full_title = a.get("title") or a.get_text(strip=True)

                key = f"{event_id}_{date_str}"
                if key in seen_keys:
                    continue
                seen_keys.add(key)

                if event_id in cache:
                    info = cache[event_id]
                else:
                    info = fetch_detail_info(session, detail_url)
                    if info["room"] == "その他・未記載":
                        info["room"] = classify_room(full_title)
                    cache[event_id] = info
                    time.sleep(0.25)

                all_events.append({
                    "id": event_id,
                    "date": date_str,
                    "weekday": wd_str,
                    "time": _normalize_time(info["time"]),
                    "room": info["room"],
                    "place": info["place"],
                    "title": full_title,
                    "org": info["org"],
                    "fee": info["fee"],
                    "url": detail_url
                })

        if curr_month == 12:
            curr_year += 1
            curr_month = 1
        else:
            curr_month += 1

    all_events.sort(key=lambda x: (x["date"], _start_hour(x["time"]), x["time"], x["id"]))

    with open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)

    apply_first_seen(all_events, now.strftime("%Y-%m-%d"))

    availability = {}
    if all_events:
        last_date = max(datetime.date.fromisoformat(e["date"]) for e in all_events)
        day = now.date()
        while day <= last_date:
            day_avail = fetch_availability_day(session, day)
            if day_avail is not None:
                availability[day.strftime("%Y-%m-%d")] = day_avail
            time.sleep(0.25)
            day += datetime.timedelta(days=1)
    print(f"空き状況: {len(availability)}日分を取得しました。")

    radio_monday = _radio_monday(now)
    radio_history = dict(_RADIO_SEED)
    if os.path.exists(RADIO_HISTORY_PATH):
        try:
            with open(RADIO_HISTORY_PATH, encoding="utf-8") as f:
                radio_history.update(json.load(f))
        except Exception as e:
            print(f"ラジオ履歴の読み込み失敗: {e}")
    radio_picks = prepare_radio_picks(all_events, radio_monday, radio_history)
    with open(RADIO_HISTORY_PATH, "w", encoding="utf-8") as f:
        json.dump(radio_history, f, ensure_ascii=False, indent=2)
    radio_filename = generate_radio_docx(radio_picks, radio_monday)
    print(f"ラジオ原稿: {radio_monday}放送分を {radio_filename} に出力しました。")

    updated_str = now.strftime("%Y-%m-%d %H:%M")
    week_start = _report_monday(now)
    week_end = week_start + datetime.timedelta(days=6)
    period_start = _newsletter_period_start(now)
    period_end = _period_end(period_start)
    generate_html(all_events, updated_str, week_start, week_end, period_start, period_end, availability, radio_monday, radio_filename)
    print(f"完了: {len(all_events)}件の予定を含む index.html を生成しました。")
    generate_weekly_docx(all_events, week_start)
    print(f"週報: {week_start}からの7日分を weekly.docx に出力しました。")
    n_news = generate_newsletter_excel(all_events, period_start)
    print(f"あいぽーと通信用: {period_start}〜{period_end}の{n_news}件を newsletter.xlsx に出力しました。")


WEEKLY_DOCX_PATH = "weekly.docx"
NEWSLETTER_XLSX_PATH = "newsletter.xlsx"


def _fit_sizes(event_col, seminar_col):
    """日付・見出しは固定サイズ。本文を1ページに収まる最大にする"""
    factor = 1.40  # 行間の実効係数（Wordの行間倍率0.78適用後の実測ベース）
    usable = 182.0  # 1ページに使える高さ(mm)
    date_pt = 24
    head_pt = 18

    def col_lines(events, body):
        lines = 0
        for e in events:
            lines += 1
            lines += len(_balanced_lines(e.get("org", ""), body, 6))
            lines += len(_balanced_lines(e.get("title", ""), body, 7))
            lines += 1
        return lines

    def total_height(body):
        h = date_pt * factor * 0.3528
        h += head_pt * factor * 0.3528
        h += max(col_lines(event_col, body),
                 col_lines(seminar_col, body)) * body * factor * 0.3528
        return h

    body = 30.0
    while body > 9.5 and total_height(body) > usable:
        body -= 0.5
    # 1行に収まる行は割らない：全行が1行に収まる最大サイズを優先する
    limit = 30.0
    for e in event_col + seminar_col:
        for lu, value in ((6, e.get("org", "")), (7, e.get("title", "")),
                          (6, e.get("fee", ""))):
            w = lu + _disp_width(value)
            if w > 0:
                limit = min(limit, 139 / (w * 0.3528))
        w = 1.0 + _disp_width(e.get("time", ""))
        limit = min(limit, 139 / (w * 0.3528))
    no_split = int(limit * 2) / 2
    if no_split >= 15:
        body = min(body, no_split)
    return (date_pt, head_pt, body)


def _newsletter_period_start(now):
    """あいぽーと通信の対象期間の開始日。基本は当月16日、毎月15日17時以降は翌月16日に切り替える。"""
    y, m = now.year, now.month
    if now.day > 15 or (now.day == 15 and now.hour >= 17):
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return datetime.date(y, m, 16)


def _period_end(start):
    y, m = (start.year + 1, 1) if start.month == 12 else (start.year, start.month + 1)
    return datetime.date(y, m, 16) - datetime.timedelta(days=1)


def _to_hhmm(time_str):
    s = (time_str or "").replace(" ", "").replace("　", "")
    s = re.sub(r"(\d+)時(\d+)分", lambda m: f"{int(m.group(1))}:{m.group(2)}", s)
    s = re.sub(r"(\d+)時", lambda m: f"{int(m.group(1))}:00", s)
    return s


def _place_marker(room):
    if room == "イベントコーナー":
        return "イ"
    if room == "会議セミナー室":
        return "会"
    if room == "イベントコーナー・会議セミナー室":
        return "イ・会"
    return ""


def _report_monday(now):
    """週報対象週の月曜日。基本は今週の月曜、木曜17時以降〜日曜は翌週の月曜にする。"""
    monday = (now - datetime.timedelta(days=now.weekday())).date()
    if (now.weekday() == 3 and now.hour >= 17) or now.weekday() in (4, 5, 6):
        monday += datetime.timedelta(days=7)
    return monday


def _radio_monday(now):
    """ラジオ原稿の対象放送日（月曜）。水曜17時以降は翌週の月曜にする。"""
    monday = (now - datetime.timedelta(days=now.weekday())).date()
    if (now.weekday() == 2 and now.hour >= 17) or now.weekday() in (3, 4, 5, 6):
        monday += datetime.timedelta(days=7)
    return monday


RADIO_HISTORY_PATH = "radio_history.json"
FIRST_SEEN_PATH = "first_seen.json"


def apply_first_seen(events, today_str):
    """Attach firstSeen (the JST date an event URL first appeared) to each event.

    State is kept in first_seen.json (committed back to the repo by Actions).
    On the very first run (no state file), every current event is seeded with
    None ("before tracking started") so nothing is reported as new; from the
    second run on, URLs that appear for the first time get today's date.
    """
    first_seen = {}
    exists = os.path.exists(FIRST_SEEN_PATH)
    if exists:
        try:
            with open(FIRST_SEEN_PATH, "r", encoding="utf-8") as f:
                first_seen = json.load(f)
        except Exception as e:
            print(f"first_seen load failed: {e}")
            exists = False
    for e in events:
        if e["url"] not in first_seen:
            first_seen[e["url"]] = today_str if exists else None
        e["firstSeen"] = first_seen[e["url"]]
    with open(FIRST_SEEN_PATH, "w", encoding="utf-8") as f:
        json.dump(first_seen, f, ensure_ascii=False, indent=2)
    n_new = sum(1 for e in events if e.get("firstSeen") == today_str)
    print(f"firstSeen: {n_new} events first confirmed today")
    return first_seen
RADIO_GREEN = RGBColor(0x00, 0xB0, 0x50)
RADIO_RED = RGBColor(0xFF, 0x00, 0x00)
RADIO_FONT = "Meiryo UI"

# 過去に手作業で作った原稿の履歴（履歴ファイルが無い期間の種データ）
_RADIO_SEED = {
    "2026-09-14": [
        {"org": "ヨガ健康クラブ", "title": "初心者向けヨガレッスン"},
        {"org": "世界の民族舞踊ペガサス", "title": "世界各地の民族の踊りを楽しもう"},
        {"org": "NPO法人でんでん虫の会", "title": "コミュニケーション交流"},
        {"org": "スマイリー健＆アロハ・スターハワイアンズ", "title": "ウクレレの弾き方等のレクチャー"},
        {"org": "シニアクラブうらら", "title": "シニアフィットネス・出前講座"},
        {"org": "NPO法人熊本わくわく体操ボランティア協会", "title": "歌謡曲・ナツメロ・流行歌・童謡に合わせてリズミカルに体を動かすイス腰掛け福祉体操"},
        {"org": "さわやか南京玉すだれを楽しむ会", "title": "南京玉すだれの公演と体験会"},
    ],
    "2026-10-12": [
        {"org": "エンジョイズ", "title": "童謡・唱歌・懐かしい流行歌を楽しむ会"},
        {"org": "楽しいオリガミ", "title": "折紙製作勉強会"},
        {"org": "KKA高齢者健康友の会", "title": "みんな笑顔で楽しく健康体操"},
        {"org": "NPO法人傾聴ネットキーステーション", "title": "傾聴カフェ、花ことば"},
        {"org": "令和9年度助成事業説明会", "title": "令和9年度助成事業説明会"},
        {"org": "くまもと学生ボランティアFESTA2026", "title": "くまもと学生ボランティアFESTA2026"},
        {"org": "あいぽーとFESTIVAL2026", "title": "つなぐ・つなげるあいぽーと"},
    ],
    "2026-10-19": [
        {"org": "フラ・オ・レアレア", "title": "フラダンスレッスン"},
        {"org": "囲碁普及会", "title": "囲碁の対局・初心者の指導と研究"},
        {"org": "楽画喜の会", "title": "絵画制作活動"},
        {"org": "NPO法人成年後見安心サポートネット熊本", "title": "成年後見無料相談会"},
        {"org": "熊本健康マージャン協会", "title": "マージャン協会の座談会の打合せ"},
        {"org": "ほっとステージBINGO", "title": "ベビーマッサージ教室"},
        {"org": "fit(0-0)フィットれいわ", "title": "写真整理で交流会"},
    ],
}


def _radio_candidates(events, date_str, excluded_orgs, used_orgs):
    cands = [e for e in events
             if e.get("date") == date_str
             and "中止" not in e.get("title", "")
             and e.get("org")
             and e["org"] not in excluded_orgs
             and e["org"] not in used_orgs]
    free = [e for e in cands if (e.get("fee") or "").strip() == "無料"]
    return free or cands


def _pick_radio_events(events, monday, excluded_orgs, rng):
    picks = []
    used = set()
    for i in range(7):
        day_str = (monday + datetime.timedelta(days=i)).isoformat()
        pool = _radio_candidates(events, day_str, excluded_orgs, used)
        if pool:
            chosen = rng.choice(pool)
            used.add(chosen["org"])
            picks.append({"org": chosen["org"], "title": chosen["title"]})
        else:
            picks.append({"org": "", "title": ""})
    return picks


def _radio_pick_still_valid(events, monday, idx, pick):
    if not pick.get("org"):
        return False
    day_str = (monday + datetime.timedelta(days=idx)).isoformat()
    for e in events:
        if (e.get("date") == day_str and e.get("org") == pick["org"]
                and e.get("title") == pick["title"] and "中止" not in e.get("title", "")):
            return True
    return False


def prepare_radio_picks(events, monday, history):
    """対象週の選定を返す。履歴にあれば再利用（無効になった日だけ再選定）、無ければ新規にランダム選定して履歴へ記録。"""
    key = monday.isoformat()
    past_keys = sorted(k for k in history if k < key)
    excluded = set()
    for k in past_keys[-2:]:
        for p in history[k]:
            if p.get("org"):
                excluded.add(p["org"])
    rng = random.Random(key)
    if key in history:
        picks = [dict(p) for p in history[key]]
        used = {p["org"] for p in picks if p.get("org")}
        for i, p in enumerate(picks):
            if not _radio_pick_still_valid(events, monday, i, p):
                old_org = p.get("org")
                pool = _radio_candidates(events, (monday + datetime.timedelta(days=i)).isoformat(),
                                         excluded, used - {old_org})
                if pool:
                    chosen = rng.choice(pool)
                    picks[i] = {"org": chosen["org"], "title": chosen["title"]}
                    used.discard(old_org)
                    used.add(chosen["org"])
                else:
                    picks[i] = {"org": "", "title": ""}
        history[key] = picks
        return picks
    picks = _pick_radio_events(events, monday, excluded, rng)
    history[key] = picks
    return picks


def generate_radio_docx(picks, monday):
    """ラジオ原稿（朝リズム Let's go あいぽーと）をWordで生成し、ファイル名を返す。"""
    filename = f"letsgo_{monday:%m%d}.docx"
    date_label = f"{monday.month}/{monday.day}"
    black = RGBColor(0, 0, 0)
    doc = Document()
    for section in doc.sections:
        section.top_margin = Pt(36)
        section.bottom_margin = Pt(36)
        section.left_margin = Pt(50)
        section.right_margin = Pt(50)
    style = doc.styles["Normal"]
    style.font.name = RADIO_FONT
    style.font.size = Pt(11)
    style.element.rPr.rFonts.set(qn("w:eastAsia"), RADIO_FONT)
    style.paragraph_format.space_after = Pt(0)
    style.paragraph_format.space_before = Pt(0)

    def add_run(p, text, color, size=11):
        r = p.add_run(text)
        r.font.name = RADIO_FONT
        r.element.rPr.rFonts.set(qn("w:eastAsia"), RADIO_FONT)
        r.font.size = Pt(size)
        r.font.color.rgb = color
        return r

    tbl = doc.add_table(rows=1, cols=1)
    tblPr = tbl._tbl.tblPr
    borders = OxmlElement("w:tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement(f"w:{edge}")
        el.set(qn("w:val"), "single")
        el.set(qn("w:sz"), "6")
        el.set(qn("w:space"), "0")
        el.set(qn("w:color"), "000000")
        borders.append(el)
    tblPr.append(borders)
    cell = tbl.cell(0, 0)
    p = cell.paragraphs[0]
    add_run(p, "國津様", black)
    p = cell.add_paragraph()
    add_run(p, "お世話になっております。", black)
    add_run(p, date_label, RADIO_GREEN)
    add_run(p, "（月）朝リズム（Let’s go あいぽーと）原稿です。", black)
    p = cell.add_paragraph()
    add_run(p, "メールもお送りしております。よろしくお願いします。浦田", black)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    p.paragraph_format.space_before = Pt(10)
    p.paragraph_format.space_after = Pt(8)
    add_run(p, date_label, RADIO_GREEN, 18)
    add_run(p, "（月）放送分", black, 18)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(4)
    add_run(p, LB + "あいぽーとの紹介" + RB, black).bold = True
    intro = [
        "毎週月曜日のこの時間は、「熊本市市民活動支援センター・あいぽーと」に関しての情報をお送りします。",
        "あいぽーとは、熊本市内で活動されているボランティアやNPOなど、様々な分野の市民活動を支援する総合拠点です。",
        "活動や交流の場や情報の提供、相談などを行い、これから活動しようと考えている人や、すでに活動に取り組んでいる人を幅広くサポートしています。",
        "場所は、中央区大江にあるウェルパルくまもと１階です。開館時間は朝８時３０分から夜９時までです。",
        "尚、NPO法人に関する相談については午前１０時から午後７時までです。",
        "相談の予約も可能です。",
        "休館日は、毎月第２木曜日と年末年始です。公共交通機関をご利用の場合は、市電・バスともに「交通局前」下車です。",
    ]
    for t in intro:
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(2)
        add_run(p, t, black)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    add_run(p, "－" * 43, black)
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(8)
    add_run(p, "今週の「あいぽーと」も、様々な分野の団体が活動されますよ。", black)
    verbs = ["を開催します", "を開催します", "を行います", "を開催します", "を開催します", "などを行います", "を開催します"]
    wd_names = ["月曜日", "火曜日", "水曜日", "木曜日", "金曜日", "土曜日", "日曜日"]
    for i in range(7):
        pick = picks[i] if i < len(picks) else {"org": "", "title": ""}
        org = pick.get("org") or "（未定）"
        title = pick.get("title") or "（未定）"
        p = doc.add_paragraph()
        p.paragraph_format.space_after = Pt(9)
        add_run(p, wd_names[i] + "は、", black)
        add_run(p, "「" + org + "」が「" + title + "」" + verbs[i] + "。参加費は無料", RADIO_RED)
        add_run(p, "です。", black)
    p = doc.add_paragraph()
    p.paragraph_format.space_before = Pt(6)
    add_run(p, "今回ご紹介した団体の開催時間などはあいぽーとへお問い合せ下さい。今回ご紹介した活動以外にも、あいぽーとの場所で活動される団体さんの情報は、あいぽーとのホームページ「あいぽーとカレンダー」で紹介しておりますので、「あいぽーと」とひらがなで検索しご覧ください。", black)
    doc.save(filename)
    return filename


# 見本docxと同じ和文フォント（名称に半角カナを含むためエスケープ表記）
JP_FONT = "HGP創英角\uff7a\uff9e\uff7c\uff6f\uff78UB"


def _set_jp_font(doc):
    """文書全体の既定フォントを見本と同じにする"""
    style_el = doc.styles["Normal"].element
    rpr = style_el.find(qn("w:rPr"))
    if rpr is None:
        rpr = OxmlElement("w:rPr")
        style_el.append(rpr)
    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:ascii"), "Meiryo UI")
    rfonts.set(qn("w:hAnsi"), "Meiryo UI")
    rfonts.set(qn("w:eastAsia"), JP_FONT)
    rfonts.set(qn("w:cs"), "Meiryo UI")
    doc.styles["Normal"].font.size = Pt(12)
    pf = doc.styles["Normal"].paragraph_format
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    pf.line_spacing = 0.78


def _shade_cell(cell, fill):
    """表セルの背景色を指定する"""
    tc = cell._tc
    tcpr = tc.find(qn("w:tcPr"))
    if tcpr is None:
        tcpr = OxmlElement("w:tcPr")
        tc.insert(0, tcpr)
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    tcpr.append(shd)


def generate_weekly_docx(events, week_start):
    """次の月曜〜日曜の7日分を1日1ページのWord文書（weekly.docx）に出力する"""
    doc = Document()
    _set_jp_font(doc)
    section = doc.sections[0]
    section.page_width = Mm(297)
    section.page_height = Mm(210)
    section.orientation = WD_ORIENT.LANDSCAPE
    section.top_margin = Mm(13)
    section.bottom_margin = Mm(13)
    section.left_margin = Mm(5)
    section.right_margin = Mm(5)
    weekdays = ["月", "火", "水", "木", "金", "土", "日"]
    fsp = "\u3000"

    for offset in range(7):
        day = week_start + datetime.timedelta(days=offset)
        date_str = day.strftime("%Y-%m-%d")
        wd = weekdays[day.weekday()]
        day_events = [e for e in events
                      if e.get("date") == date_str and "中止" not in e.get("title", "")]
        day_events.sort(key=lambda x: (_start_hour(x.get("time", "")), x.get("time", ""), x.get("id", "")))

        event_col, seminar_col = [], []
        for e in day_events:
            room = e.get("room", "")
            has_event = "イベント" in room
            has_seminar = "会議" in room
            if has_event and has_seminar:
                event_col.append(e)
                seminar_col.append(e)
            elif has_event:
                event_col.append(e)
            elif has_seminar:
                seminar_col.append(e)
            else:
                seminar_col.append(e)

        date_size, head_size, body_size = _fit_sizes(event_col, seminar_col)

        p = doc.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_run = p.add_run(f"{day.month}月{day.day}日（{wd}）{fsp}◎どなたでも参加できます")
        date_run.bold = True
        date_run.font.size = Pt(date_size)

        table = doc.add_table(rows=2, cols=2)
        table.style = "Table Grid"
        for hdr_text, hdr_cell in (("イベントコーナー", table.rows[0].cells[0]),
                                   ("会議・セミナー室", table.rows[0].cells[1])):
            hdr_cell.text = hdr_text
            hdr_run = hdr_cell.paragraphs[0].runs[0]
            hdr_run.bold = True
            hdr_run.font.size = Pt(head_size)
            _shade_cell(hdr_cell, "D9D9D9")

        for col_events, cell in ((event_col, table.rows[1].cells[0]),
                                 (seminar_col, table.rows[1].cells[1])):
            use_first = True
            for e in col_events:
                if use_first:
                    time_p = cell.paragraphs[0]
                    use_first = False
                else:
                    time_p = cell.add_paragraph()
                time_run = time_p.add_run("◎" + e.get("time", ""))
                time_run.bold = True
                time_run.font.size = Pt(body_size)
                for label, value, lu in (("・使用団体：", e.get("org", ""), 6),
                                         ("・内" + fsp * 3 + "容：", e.get("title", ""), 7)):
                    parts = _balanced_lines(value, body_size, lu)
                    for j, part in enumerate(parts):
                        body_p = cell.add_paragraph()
                        body_run = body_p.add_run((label if j == 0 else "") + part)
                        body_run.font.size = Pt(body_size)
                fee_p = cell.add_paragraph()
                fee_run = fee_p.add_run("・参 加 費：" + e.get("fee", ""))
                fee_run.font.size = Pt(body_size)

        if offset < 6:
            doc.add_page_break()

    doc.save(WEEKLY_DOCX_PATH)


def generate_newsletter_excel(events, period_start):
    """あいぽーと通信用カレンダー（16日〜翌月15日）をExcelで出力する"""
    period_end = _period_end(period_start)
    wd_jp = ["月", "火", "水", "木", "金", "土", "日"]
    group_order = {"イ": 0, "イ・会": 1, "会": 2, "": 3}

    picked = []
    for e in events:
        if "中止" in e.get("title", ""):
            continue
        d = datetime.date.fromisoformat(e["date"])
        if not (period_start <= d <= period_end):
            continue
        marker = _place_marker(e.get("room", ""))
        picked.append((d, group_order[marker], _start_hour(e.get("time", "")),
                       e.get("time", ""), e, marker))
    picked.sort(key=lambda r: (r[0], r[1], r[2], r[3]))

    wb = Workbook()
    ws = wb.active
    ws.title = "あいぽーとカレンダー"
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.orientation = "portrait"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    thin = Side(style="thin", color="666666")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws.merge_cells("A1:E1")
    c = ws["A1"]
    c.value = f"あいぽーとカレンダー {period_start.month}月（後半）～{period_end.month}月（前半）"
    c.font = Font(name="Meiryo UI", size=14, bold=True)

    ws.merge_cells("A2:E2")
    c = ws["A2"]
    c.value = ("※受付分までの予定を記載しています。団体の判断で中止となる場合もありますので、"
               "事前に主催団体へ確認のうえ参加ください。「あいぽーとカレンダー」は、"
               "あいぽーとホームページにも掲載しています。掲載している内容は、一般の方も参加できます。"
               "主催団体のルールに則って参加ください。")
    c.font = Font(name="Meiryo UI", size=9)
    c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.row_dimensions[2].height = 30

    ws.merge_cells("A3:E3")
    c = ws["A3"]
    c.value = "場所：イ：イベントコーナー　会：会議・セミナー室"
    c.font = Font(name="Meiryo UI", size=10)

    fill = PatternFill("solid", fgColor="D9D9D9")
    for col, h in enumerate(["期日", "時間", "主催", "内容", "場所"], start=1):
        c = ws.cell(row=4, column=col, value=h)
        c.font = Font(name="Meiryo UI", size=11, bold=True)
        c.fill = fill
        c.border = border
        c.alignment = Alignment(horizontal="center", vertical="center")

    for i, w in enumerate([14, 15, 34, 52, 8], start=1):
        ws.column_dimensions[chr(64 + i)].width = w

    r = 5
    spans = []
    cur_date = None
    span_start = None
    for d, _g, _h, _t, e, marker in picked:
        if d != cur_date:
            if cur_date is not None:
                spans.append((span_start, r - 1))
            cur_date = d
            span_start = r
            c = ws.cell(row=r, column=1, value=f"{d.month}月{d.day}日（{wd_jp[d.weekday()]}）")
            c.alignment = Alignment(vertical="top")
        for col, v in ((2, _to_hhmm(e.get("time", ""))), (3, e.get("org", "")),
                       (4, e.get("title", "")), (5, marker)):
            c = ws.cell(row=r, column=col, value=v)
            c.alignment = Alignment(wrap_text=(col in (3, 4)), vertical="top",
                                    horizontal="center" if col == 5 else "left")
        for col in range(1, 6):
            cell = ws.cell(row=r, column=col)
            cell.border = border
            cell.font = Font(name="Meiryo UI", size=11)
        r += 1
    if cur_date is not None:
        spans.append((span_start, r - 1))
    for s, e_row in spans:
        if e_row > s:
            ws.merge_cells(start_row=s, start_column=1, end_row=e_row, end_column=1)

    wb.save(NEWSLETTER_XLSX_PATH)
    return len(picked)


if __name__ == "__main__":
    main()
