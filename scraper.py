import datetime
import time
import os
import re
import json
import unicodedata
from urllib.parse import urljoin, urlparse, parse_qs
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.section import WD_ORIENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Mm, Pt
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

BASE_URL = "https://www.kumamoto-aiport.com/cal_copy/pub/default.aspx"
CACHE_PATH = "cache.json"
HTML_PATH = "index.html"

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


def generate_html(events, updated_str, week_start, week_end, period_start, period_end):
    events_json = json.dumps(events, ensure_ascii=False)
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
<title>あいぽーと 予定検索システム</title>
<style>
  body {{ font-family: "Helvetica Neue", Arial, "Hiragino Kaku Gothic ProN", Meiryo, sans-serif; background: #f5f7fa; color: #333; margin: 0; padding: 20px; }}
  .container {{ max-width: 1100px; margin: 0 auto; background: #fff; padding: 24px; border-radius: 12px; box-shadow: 0 2px 10px rgba(0,0,0,0.06); }}
  h1 {{ margin-top: 0; font-size: 1.5rem; color: #1e293b; border-bottom: 2px solid #e2e8f0; padding-bottom: 12px; }}
  .updated {{ font-size: 0.85rem; color: #64748b; margin-bottom: 16px; }}
  .filters {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(220px, 1fr)); gap: 12px; background: #f8fafc; padding: 16px; border-radius: 8px; border: 1px solid #e2e8f0; margin-bottom: 20px; }}
  .filter-group label {{ display: block; font-size: 0.85rem; font-weight: bold; margin-bottom: 6px; color: #475569; }}
  .filter-group input[type="text"], .filter-group input[type="date"] {{ width: 100%; padding: 8px; border: 1px solid #cbd5e1; border-radius: 6px; box-sizing: border-box; font-size: 0.95rem; }}
  .tabs {{ display: flex; gap: 8px; margin-bottom: 16px; flex-wrap: wrap; }}
  .tab-btn {{ padding: 10px 18px; border: none; background: #e2e8f0; color: #475569; border-radius: 8px; cursor: pointer; font-weight: bold; font-size: 0.95rem; transition: 0.2s; }}
  .tab-btn.active {{ background: #2563eb; color: #fff; }}
  .count-bar {{ font-weight: bold; margin-bottom: 10px; color: #334155; }}
  .table-wrap {{ overflow-x: auto; }}
  table {{ width: 100%; border-collapse: collapse; font-size: 0.92rem; }}
  th, td {{ padding: 10px 12px; border-bottom: 1px solid #e2e8f0; text-align: left; }}
  th {{ background: #f1f5f9; color: #334155; white-space: nowrap; }}
  tr:hover {{ background: #f8fafc; }}
  .badge {{ display: inline-block; padding: 4px 8px; border-radius: 6px; font-size: 0.8rem; font-weight: bold; white-space: nowrap; }}
  .badge-event {{ background: #fef3c7; color: #b45309; }}
  .badge-seminar {{ background: #dbeafe; color: #1d4ed8; }}
  .badge-both {{ background: #ede9fe; color: #6d28d9; }}
  .badge-other {{ background: #f1f5f9; color: #64748b; }}
  a.btn-link {{ display: inline-block; padding: 5px 10px; background: #f1f5f9; color: #2563eb; text-decoration: none; border-radius: 6px; font-size: 0.85rem; white-space: nowrap; }}
  a.btn-link:hover {{ background: #e2e8f0; }}
  .btn-update {{ display: inline-block; margin-left: 12px; padding: 8px 16px; background: #2563eb; color: #fff; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 0.9rem; }}
  .btn-update:hover {{ background: #1d4ed8; }}
  .update-note {{ font-size: 0.8rem; color: #64748b; margin: -8px 0 16px; }}
  .btn-docx {{ display: inline-block; margin-left: 12px; padding: 8px 16px; background: #16a34a; color: #fff; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 0.9rem; }}
  .btn-docx:hover {{ background: #15803d; }}
  .btn-excel {{ display: inline-block; margin-left: 12px; padding: 8px 16px; background: #217346; color: #fff; text-decoration: none; border-radius: 8px; font-weight: bold; font-size: 0.9rem; }}
  .btn-excel:hover {{ background: #1a5c38; }}
</style>
</head>
<body>
<div class="container">
  <h1>あいぽーと イベントコーナー・会議セミナー室 予定検索</h1>
  <div class="updated">最終自動更新: {updated_str} （毎日1回自動更新）
    <a class="btn-update" href="https://github.com/aiport0960168/aiport-calendar2/actions/workflows/update_data.yml" target="_blank" rel="noopener">🔄 最新情報に更新</a>
    <a class="btn-docx" href="weekly.docx" download>📄 週報ダウンロード（Word）</a>
    <a class="btn-excel" href="newsletter.xlsx" download>📊 あいぽーと通信用カレンダー（Excel）</a>
  </div>
  <p class="update-note">※データは毎朝6時00分に自動更新されます。「最新情報に更新」を押すとGitHubのページが開きます。「Run workflow」を押すと数分で最新のデータに更新されます。</p>
  <p class="update-note">※週報は今週の月曜〜日曜の7日分（1日1ページ）です。木曜17時以降は翌週分に切り替わります。今ダウンロードできるのは{week_label}分です。</p>
  <p class="update-note">※あいぽーと通信用カレンダーは16日〜翌月15日分です。毎月15日17時に、翌月16日〜翌々月15日分へ切り替わります。今ダウンロードできるのは{period_label}分です。</p>

  <div class="filters">
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
    <div class="filter-group" style="display:flex; align-items:flex-end; padding-bottom:6px; gap:16px;">
      <label style="cursor:pointer; margin:0;"><input type="checkbox" id="hideCancel" checked> {cancel_label}</label>
      <label style="cursor:pointer; margin:0;"><input type="checkbox" id="onlyCancel"> {only_label}</label>
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

  <div class="table-wrap">
    <table>
      <thead>
        <tr>
          <th>日付</th>
          <th>時間</th>
          <th>場所区分</th>
          <th>タイトル</th>
          <th>団体名</th>
          <th>場所詳細</th>
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
  for (const ev of filtered) {{
    rows += '<tr>'
      + '<td style="white-space:nowrap; font-weight:bold;">' + ev.date + ' (' + ev.weekday + ')</td>'
      + '<td style="white-space:nowrap;">' + ev.time + '</td>'
      + '<td><span class="badge ' + getBadgeClass(ev.room) + '">' + ev.room + '</span></td>'
      + '<td style="font-weight:bold;">' + ev.title + '</td>'
      + '<td>' + ev.org + '</td>'
      + '<td>' + ev.place + '</td>'
      + '<td><a class="btn-link" href="' + ev.url + '" target="_blank" rel="noopener">公式詳細</a></td>'
      + '</tr>';
  }}
  tbody.innerHTML = rows;
}}

document.querySelectorAll('#startDate, #endDate, #keyword, #hideCancel, #onlyCancel').forEach(el => {{
  el.addEventListener('input', render);
  el.addEventListener('change', render);
}});

render();
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

    updated_str = now.strftime("%Y-%m-%d %H:%M")
    week_start = _report_monday(now)
    week_end = week_start + datetime.timedelta(days=6)
    period_start = _newsletter_period_start(now)
    period_end = _period_end(period_start)
    generate_html(all_events, updated_str, week_start, week_end, period_start, period_end)
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
