import streamlit as st
import pandas as pd
import datetime
import calendar
import io
import math
import random
import re
import pathlib
import hashlib
import json
from ortools.sat.python import cp_model
import jpholiday

# ==========================================
# 共通の定数
# ==========================================
WEEKDAYS_JA = ["月", "火", "水", "木", "金", "土", "日"]
NIGHT_SHIFTS = ['A宿直', 'B宿直', '外来宿直']
DAY_SHIFTS = ['A日直', 'B日直', '外来日直']
ALL_SHIFT_TYPES = ['A宿直', 'B宿直', '外来宿直', 'A日直', 'B日直', '外来日直']
RESULT_COLUMNS = ["日付", "平日/休日", "A日直", "A宿直", "B日直", "B宿直", "外来日直", "外来宿直"]
NUMERIC_DEFAULTS = {'希望優先度(数字が大きいほど優先)': 1, '最低空ける日数': 5, '月間最小回数': 0, '月間最大回数': 5, '休日最大回数': 4, **{s + '上限': 2 for s in ALL_SHIFT_TYPES}}
NG_COLUMN = 'NG日(半角カンマ区切り)'
# 宿直・日直それぞれの月間上限（A・B・外来を合わせた回数）。空欄なら制限なし
GROUP_CAP_COLUMNS = {'宿直上限': ['A宿直', 'B宿直', '外来宿直'], '日直上限': ['A日直', 'B日直', '外来日直']}
REQUEST_COLUMN = '希望日(半角カンマ区切り)'


def clean_text(value):
    return '' if pd.isna(value) else str(value).strip()


# ==========================================
# 重いCSV読み込みを一瞬で終わらせる魔法（キャッシュ機能）
# ==========================================
def _read_csv_any_encoding(file_bytes):
    """
    UTF-8（BOMあり・なし）→ Shift-JIS(cp932) の順に試してDataFrameを返す。
    UTF-8は形式の合わないファイルでは確実に失敗するため、先に試すと誤判定が起きにくい。
    どれでも読めない場合は、文字化けしたまま進まないようにエラーにする。
    io.BytesIO はread後にポインタが末尾へ移動するため、毎回新しく作る。
    """
    for encoding in ('utf-8-sig', 'cp932'):
        try:
            return pd.read_csv(io.BytesIO(file_bytes), encoding=encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError("CSVの文字コードを判定できませんでした。Excelの「名前を付けて保存」で「CSV UTF-8」形式を選んで保存し直してください。")


def normalize_shift_names(df):
    """以前のCSVの枠名・上限列・希望枠を現在の名称へ変換する。"""
    aliases = {"宿直A": "A宿直", "宿直B": "B宿直", "A当直": "A宿直", "B当直": "B宿直", "日直A": "A日直", "日直B": "B日直"}
    for old, new in aliases.items():
        for suffix in ("", "上限"):
            old_col, new_col = old + suffix, new + suffix
            if old_col not in df.columns:
                continue
            if new_col not in df.columns:
                df = df.rename(columns={old_col: new_col})
            else:
                blank = df[new_col].fillna("").astype(str).str.strip().eq("")
                df.loc[blank, new_col] = df.loc[blank, old_col]
                df = df.drop(columns=[old_col])
    if REQUEST_COLUMN in df.columns:
        def normalize_request(value):
            if not isinstance(value, str):
                return value
            for old, new in aliases.items():
                value = value.replace(old, new)
            return value
        df[REQUEST_COLUMN] = df[REQUEST_COLUMN].map(normalize_request)
    return df


@st.cache_data
def parse_staff_csv(file_bytes):
    df = normalize_shift_names(_read_csv_any_encoding(file_bytes))
    # 旧CSVの列名も受け付け、画面・出力CSVでは新名称に統一する。
    weekday_column = "翌日PM duty"
    for old_column in ("原則、宿直を外す曜日", "入れない曜日(半角カンマ区切り)", "入れない曜日"):
        if old_column not in df.columns:
            continue
        if weekday_column not in df.columns:
            df = df.rename(columns={old_column: weekday_column})
        else:
            # 両方の列がある場合は新名称の値を優先し、空欄だけ旧列で補う。
            blank = df[weekday_column].fillna("").astype(str).str.strip().eq("")
            df.loc[blank, weekday_column] = df.loc[blank, old_column]
            df = df.drop(columns=[old_column])
    return df


@st.cache_data
def parse_fixed_csv(file_bytes):
    df = normalize_shift_names(_read_csv_any_encoding(file_bytes))
    if '区分' in df.columns:
        df = df.rename(columns={'区分': '平日/休日'})
    return df


def parse_shift_date(date_value, target_year, target_month):
    if pd.isna(date_value): return None
    if isinstance(date_value, datetime.datetime): return date_value.date()
    if isinstance(date_value, datetime.date): return date_value
    text = str(date_value).strip().translate(str.maketrans('０１２３４５６７８９／－', '0123456789/-'))
    text = re.sub(r'\([月火水木金土日]\)$', '', text).strip()
    full = re.fullmatch(r'(\d{4})[年/.-](\d{1,2})[月/.-](\d{1,2})日?(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?', text)
    short = re.fullmatch(r'(\d{1,2})[月/.-](\d{1,2})日?', text)
    if full:
        y, m, d = map(int, full.groups())
    elif short:
        m, d = map(int, short.groups()); y = target_year
        if m - target_month >= 6: y -= 1
        elif target_month - m >= 6: y += 1
    else:
        return None
    try:
        return datetime.date(y, m, d)
    except ValueError:
        return None


# ==========================================
# 休日・曜日・NG・希望日の読み取り（アプリ全体でこの関数だけを使う）
# ==========================================
def is_holiday_date(date_obj, target_year, target_month, custom_holidays):
    """土日・祝日、または対象月で「休日にする」にチェックした日なら True。"""
    return (
        date_obj.weekday() >= 5
        or jpholiday.is_holiday(date_obj)
        or ((date_obj.year, date_obj.month) == (target_year, target_month) and date_obj.day in custom_holidays)
    )


def pm_duty_weekdays(text):
    """「水,木」のような入力を曜日番号（月=0〜日=6）のリストにする。"""
    text = clean_text(text)
    return [i for i, w in enumerate(WEEKDAYS_JA) if w in text]


def checked_day_items(value, year, month, ng=False):
    text = clean_text(value).translate(str.maketrans('０１２３４５６７８９，：', '0123456789,:'))
    items = []
    for token in text.split(','):
        if not token.strip():
            continue
        parts = token.strip().split(':')
        if len(parts) > 2 or not re.fullmatch(r'\d+', parts[0].strip()):
            raise ValueError(f'「{token}」：日付は整数、複数は半角カンマで入力してください。')
        d = int(parts[0].strip())
        if not 1 <= d <= calendar.monthrange(year, month)[1]:
            raise ValueError(f'「{token}」：{year}年{month}月にない日付です。')
        kind = parts[1].strip() if len(parts) == 2 else ('全NG' if ng else None)
        if kind not in (['全NG', '日NG', '宿NG', 'OK'] if ng else ALL_SHIFT_TYPES + [None]):
            raise ValueError(f'「{token}」：枠名・NGの種類を確認してください。')
        items.append((d, kind))
    if ng:
        by_day = {}
        for d, kind in items:
            if d in by_day and by_day[d] != kind:
                raise ValueError(f'{d}日のNG指定が重複して異なっています。1種類にしてください。')
            by_day[d] = kind
    return list(dict.fromkeys(items))


def parse_ng_dict(value, year, month):
    """NG日の文字列を {日: 'OK'以外の種類} にする。入力チェック済みの値を想定。"""
    try:
        return {d: kind for d, kind in checked_day_items(value, year, month, True) if kind != 'OK'}
    except ValueError:
        return {}


def parse_requests(value, year, month):
    """希望日の文字列を「日だけの希望」と「(日, 枠)の希望」に分ける。"""
    try:
        items = checked_day_items(value, year, month, False)
    except ValueError:
        return [], []
    return [d for d, kind in items if kind is None], [(d, kind) for d, kind in items if kind is not None]


def effective_ng(value, holiday):
    """
    保存されているNG指定を、その日に実際に効く表示へ読み替える。
    保存値そのものは書き換えないので、休日設定を戻せば元の指定が復活する。
    （平日は日直がないため、全NGは宿NGと同じ意味になり、日NGは意味を持たない）
    """
    if holiday:
        return value if value in ("OK", "全NG", "日NG", "宿NG") else "OK"
    return "宿NG" if value in ("全NG", "宿NG") else "OK"


# ==========================================
# カレンダー一括操作用の裏側ロジック
# ==========================================
def pm_duty_restricted(date_obj, weekdays, target_year, target_month, custom_holidays, next_month_special_holiday=False):
    """翌日が勤務日の場合だけ、指定曜日の宿直を制限する。"""
    next_date = date_obj + datetime.timedelta(days=1)
    next_month_first = datetime.date(target_year, target_month, calendar.monthrange(target_year, target_month)[1]) + datetime.timedelta(days=1)
    next_is_holiday = (
        is_holiday_date(next_date, target_year, target_month, custom_holidays)
        or (next_month_special_holiday and next_date == next_month_first)
    )
    return date_obj.weekday() in weekdays and not next_is_holiday


def set_all_ng(doc_name, y, m, ndays, val, custom_hols=()):
    for d in range(1, ndays + 1):
        if val == "OK":
            st.session_state[f"ng_{doc_name}_{y}_{m}_{d}"] = "OK"
        else:
            # 「全日NGにする」が押された場合、休日は「全NG」、平日は「宿NG」にする
            is_hol = is_holiday_date(datetime.date(y, m, d), y, m, custom_hols)
            st.session_state[f"ng_{doc_name}_{y}_{m}_{d}"] = "全NG" if is_hol else "宿NG"


# ==========================================
# 勤務間隔制約：自動割当は確定勤務・月外勤務との間隔も守る
# ==========================================
def add_interval_constraints(
    model,
    shifts,
    doctors,
    daily_active_shifts,
    num_days,
    target_year,
    target_month,
    min_intervals,
    past_worked_dates,
    future_worked_dates,
    absolute_req_days,
    absolute_req_specific,
):
    """確定同士は維持し、自動割当と全勤務の間隔を制限する。"""
    for doc in doctors:
        gap = min_intervals[doc]
        fixed_days = set(absolute_req_days[doc]) | {d for d, s in absolute_req_specific[doc]}
        external = set(past_worked_dates.get(doc, [])) | set(future_worked_dates.get(doc, []))
        worked = {}
        for d in range(1, num_days + 1):
            variables = [shifts[d, doc, s] for s in daily_active_shifts.get(d, []) if (d, doc, s) in shifts]
            if not variables: continue
            worked[d] = model.NewBoolVar(f'working_{doc}_{d}')
            model.AddMaxEquality(worked[d], variables)
            dt = datetime.date(target_year, target_month, d)
            if d not in fixed_days and any(abs((dt - ext).days) <= gap for ext in external):
                model.Add(worked[d] == 0)
        for d in worked:
            for other in range(d + 1, min(num_days, d + gap) + 1):
                if other in worked and not (d in fixed_days and other in fixed_days):
                    model.Add(worked[d] + worked[other] <= 1)


# ==========================================
# ブラウザ内で動く表示部品（NGカレンダー・結果の表）
# HTMLは固定の「入れ物」にし、中身はデータとして渡す。
# こうすることで、何度操作しても部品（一時ファイル）は増えない。
# ==========================================
@st.cache_resource
def horizontal_ng_component(html_source):
    import tempfile
    from pathlib import Path
    import streamlit.components.v1 as components
    directory = Path(tempfile.mkdtemp(prefix="shift_ng_calendar_"))
    (directory / "index.html").write_text(html_source, encoding="utf-8")
    asset_id = hashlib.sha256(html_source.encode("utf-8")).hexdigest()[:16]
    return components.declare_component(f"shift_ng_{asset_id}", path=str(directory))


HORIZONTAL_NG_HTML = r"""<!doctype html><html lang="ja"><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><style>
/* NGカレンダー。各日の上段＝日直、下段＝宿直。押すとその勤務だけ 可⇔NG が切り替わる。
   色：日直NG・宿直NGとも同じテラコッタ。休日はマスを淡いローズにする。 */
:root{--ink:#243041;--sub:#7a8494;--line:#e3e7ed;--seg:#d8dee6;--day:#cf6a35;--night:#cf6a35;--ok:#2f8a5b;
      --hol:#fff6f5;--holline:#f0dcdc;--sun:#c8414a;--sat:#2f68b3}
*{box-sizing:border-box}
body{margin:0;font-family:system-ui,-apple-system,"Hiragino Sans","Yu Gothic UI","Meiryo",sans-serif;color:var(--ink);background:#fff;font-size:14px;-webkit-font-smoothing:antialiased}
.legend{display:flex;flex-wrap:wrap;gap:6px 16px;align-items:center;font-size:12px;color:var(--sub);margin:0 2px 12px}
.legend i{display:inline-block;width:12px;height:12px;border-radius:3px;vertical-align:-1px;margin-right:5px}
.strip{display:flex;gap:8px;overflow-x:auto;padding:2px 2px 14px}
.strip.month{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:6px;overflow:visible;padding-bottom:6px}
.wk{text-align:center;font-size:15px;font-weight:800;letter-spacing:.06em;color:#4a5566;padding:0 0 6px}
.blank{border-radius:12px;background:#fafbfc}
.cell{flex:0 0 112px;background:#fff;border:1px solid var(--line);border-radius:12px;padding:8px 7px 7px;
      box-shadow:0 1px 2px rgba(16,24,40,.04);display:flex;flex-direction:column;align-items:stretch;min-width:0}
.month .cell{flex:none}
.cell[data-holiday="true"]{background:var(--hol);border-color:var(--holline)}
.date{text-align:center;font-size:21px;font-weight:800;line-height:26px;white-space:nowrap;font-variant-numeric:tabular-nums}
.date small{font-size:14px;font-weight:700;margin-left:1px}
.date.holiday{color:var(--sun)}.date.saturday{color:var(--sat)}
.flag{height:26px;display:flex;align-items:center;justify-content:center}
.pm{display:inline-flex;align-items:center;gap:3px;color:#bf5700;background:#fff0c2;border:1px solid #ef9b20;border-radius:4px;padding:0 5px;font-size:13px;font-weight:900;line-height:18px;white-space:nowrap}
.pm small{font-size:9.5px;font-weight:700;line-height:1.15}
.seg{border:1px solid var(--seg);border-radius:8px;overflow:hidden;background:#fff}
.half{height:34px;display:flex;align-items:center;justify-content:center;gap:6px;font-size:12.5px;font-weight:700;color:var(--ink);background:#fff;
      cursor:pointer;user-select:none;-webkit-tap-highlight-color:transparent;transition:background .15s,color .15s}
.half + .half{border-top:1px solid var(--line)}
.half .st{font-size:11.5px;font-weight:800;color:var(--ok)}
.half:hover{background:#f5f7fa}
.half:focus-visible{outline:2px solid #4c8dff;outline-offset:-2px}
.half.ng{color:#fff}.half.ng .st{color:rgba(255,255,255,.92)}
.half.day.ng{background:var(--day)}.half.night.ng{background:var(--night)}
.half.none{cursor:default;color:#c4cad3;font-weight:500;background:repeating-linear-gradient(135deg,#fafbfc 0 6px,#f3f5f8 6px 7px)}
.half.none:hover{background:repeating-linear-gradient(135deg,#fafbfc 0 6px,#f3f5f8 6px 7px)}
.st .m{display:none}
.slot{height:34px;flex:none}
.cell.locked{background:#f1f3f6 !important;border:1px dashed #c3cad5 !important}
.locked-note{height:69px;display:flex;align-items:center;justify-content:center;text-align:center;font-size:12px;font-weight:700;color:#8a94a6;line-height:1.4}
#note{margin:8px 2px;font-size:12px;color:var(--sub);min-height:16px}
/* 先生用画面：まとめて全日NG／すべてOKにするボタン */
#bulk{display:none;gap:8px;justify-content:flex-end;margin:0 2px 10px}
#bulk button{background:#fff;color:var(--ink);border:1px solid #c9d2de;border-radius:8px;padding:8px 14px;font:700 13px system-ui,sans-serif;cursor:pointer}
#bulk button:hover{background:#f5f7fa}
#apply{background:#ff4b4b;color:#fff;border:0;border-radius:8px;padding:11px 18px;font:700 14px system-ui,sans-serif;cursor:pointer;box-shadow:0 1px 2px rgba(0,0,0,.08)}
/* 反映していない変更があるときは、ボタンを目立たせる */
#apply.dirty{background:#e8590c;box-shadow:0 0 0 3px #ffd8a8;animation:pulse 1.6s ease-in-out infinite}
@keyframes pulse{0%,100%{box-shadow:0 0 0 3px #ffd8a8}50%{box-shadow:0 0 0 7px #ffe8cc}}
/* 管理者画面で変更がないときは、押す必要がないことが分かる見た目にする */
#apply.clean{background:#eef1f5;color:#7a8494;box-shadow:none;cursor:default}
#note.dirty{color:#c2410c;font-weight:800;font-size:13px}
/* 先生用画面：保存ボタンのすぐ上の、管理者へのコメント欄 */
#commentBox{display:none;margin:4px 2px 12px}
#commentBox label{display:block;font-size:14px;font-weight:700;margin-bottom:5px;color:var(--ink)}
#commentBox label small{font-weight:500;color:var(--sub)}
#comment{width:100%;min-height:76px;resize:vertical;padding:8px 10px;border:1px solid #c9d2de;border-radius:8px;font:14px/1.5 system-ui,sans-serif;color:var(--ink)}
#comment:focus{outline:2px solid #4c8dff;outline-offset:0;border-color:transparent}
@media(max-width:600px){
 .strip.month{gap:3px}
 .month .cell{padding:5px 2px 3px;border-radius:8px}
 .month .date{font-size:17px;line-height:21px}
 .wk{font-size:13px}
 .month .flag{height:28px}
 .month .pm{font-size:11px;padding:0 2px;gap:1px;line-height:14px}
 .month .pm small{font-size:7.5px;white-space:normal;max-width:30px;overflow-wrap:anywhere}
 .month .seg{border-radius:6px}
 .month .slot{height:30px}
 .month .locked-note{height:61px;font-size:10px}
 .month .half{height:30px;font-size:10px;gap:0;flex-direction:column;line-height:1.1}
 .month .half .st{font-size:9.5px}
 .month .st .w,.month .none .w{display:none}.month .st .m,.month .none .m{display:inline}
 .month .blank{border-radius:8px}
}
</style><body>

<div id="bulk"><button type="button" id="allNg">全日NGにする</button><button type="button" id="allOk">すべてOKに戻す</button></div><div id="strip" class="strip" aria-label="NG日カレンダー"></div><div id="note"></div><div id="commentBox"><label for="comment">管理者へのコメント・要望 <small>（任意）</small></label><textarea id="comment" maxlength="500"></textarea></div><button id="apply" type="button">NG日を反映する</button><script>
let version=null,days=[],draft=[],cells=[];let mode='row';let frameHeight=0;
function send(type,data={}){window.parent.postMessage({isStreamlitMessage:true,type,...data},'*')}
function resize(){let h=document.body.scrollHeight+10;if(h!==frameHeight){frameHeight=h;send('streamlit:setFrameHeight',{height:h})}}
function combine(dayNg,nightNg){return dayNg&&nightNg?'全NG':dayNg?'日NG':nightNg?'宿NG':'OK'}
function paintHalf(el,ng){
  el.classList.toggle('ng',ng);el.setAttribute('aria-pressed',String(ng));
  el.querySelector('.st').innerHTML=ng?'<span class="w">NG</span><span class="m">✕</span>':'<span class="w">可</span><span class="m">○</span>';
}
function paint(cell,value){
  cell.dataset.state=value;
  let db=cell.querySelector('.half.day:not(.none)'),nb=cell.querySelector('.half.night');
  if(db)paintHalf(db,value==='全NG'||value==='日NG');
  paintHalf(nb,value==='全NG'||value==='宿NG');
}
window.addEventListener('message',(event)=>{
  if(!event.data||event.data.type!=='streamlit:render')return;
  let a=event.data.args;
  reportDirty=!!a.report_dirty;baseLabel=a.button_label||(a.doctor+'先生のNG日を反映する');pendingWord=a.pending_word||'反映';
  // 先生用画面だけ、保存ボタンの上にコメント欄を出す（保存済みのコメントが変わったときだけ書き換える）
  if(a.comment_enabled){document.getElementById('commentBox').style.display='block';
    if(a.comment!==commentBase){commentBase=a.comment||'';document.getElementById('comment').value=commentBase}}
  if(a.bulk_buttons)document.getElementById('bulk').style.display='flex';
  if(a.version!==version){
    version=a.version;mode=a.mode||'row';days=a.days;draft=days.map(d=>d.value);sentDirty=false;cells=[];
    let strip=document.getElementById('strip');
    strip.replaceChildren();
    strip.className='strip '+(mode==='month'?'month':'');
    if(mode==='month'){
      ['月','火','水','木','金','土','日'].forEach((w,i)=>{let el=document.createElement('div');el.className='wk';el.textContent=w;if(i===6)el.style.color='var(--sun)';if(i===5)el.style.color='var(--sat)';strip.append(el)});
      for(let i=0;i<a.offset;i++){let el=document.createElement('div');el.className='blank';strip.append(el)}
    }
    days.forEach((d,i)=>{
      const holiday=d.options.includes('日NG');
      let cell=document.createElement('div');cell.className='cell';cell.dataset.holiday=String(holiday);
      let date=document.createElement('div');date.className='date '+d.kind;
      date.innerHTML=d.day+(mode==='month'?'':'<small>（'+d.weekday+'）</small>');
      let flag=document.createElement('div');flag.className='flag';
      if(d.warning){let pm=document.createElement('span');pm.className='pm';pm.innerHTML='⚠︎<small>翌日PM duty</small>';
        pm.title='翌日PM duty（翌日は平日）。原則としてこの日の宿直は外れます。';flag.append(pm)}
      // 日直・宿直を押すと、その勤務だけ 可⇔NG が切り替わる（休日は両方を別々に指定できる）
      const toggle=(which)=>{
        let v=draft[i];let dayNg=v==='全NG'||v==='日NG',nightNg=v==='全NG'||v==='宿NG';
        if(which==='day')dayNg=!dayNg;else nightNg=!nightNg;
        let nv=combine(dayNg,nightNg);if(!d.options.includes(nv))return;
        draft[i]=nv;paint(cell,nv);updateDirty();
      };
      const half=(which,label)=>{let h=document.createElement('div');h.className='half '+which;h.tabIndex=0;h.setAttribute('role','button');
        h.setAttribute('aria-label',d.day+'日の'+label);h.innerHTML='<span class="lb">'+label+'</span><span class="st"></span>';
        h.onclick=()=>toggle(which);h.onkeydown=(e)=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();toggle(which)}};return h};
      if(d.locked){   // 年末年始など、先生がNGを入力できない日
        cell.classList.add('locked');cell.append(date,flag);
        let ln=document.createElement('div');ln.className='locked-note';ln.innerHTML='別途<br>決定';cell.append(ln);strip.append(cell);return;
      }
      let seg=document.createElement('div');seg.className='seg';
      if(holiday)seg.append(half('day','日直'));
      seg.append(half('night','宿直'));
      cell.append(date,flag);
      // 平日は日直がないので上段を空け、宿直ボタンの位置を休日とそろえる
      if(!holiday){let sp=document.createElement('div');sp.className='slot';cell.append(sp)}
      cell.append(seg);paint(cell,d.value);strip.append(cell);cells[i]=cell;
    });
    if(mode==='month'){let blanks=(7-(a.offset+days.length)%7)%7;for(let i=0;i<blanks;i++){let el=document.createElement('div');el.className='blank';strip.append(el)}}
    document.getElementById('note').textContent=mode==='month'?'':'左右へスクロールして選べます。';
    updateDirty(true);
  }
  resize();
});
// 反映していない変更（日数）を数えて、ボタンと案内の見た目を切り替える。
// 管理者画面（report_dirty）では、変更の有無をアプリにも知らせる（先生を切り替えるときの注意に使う）
let sentDirty=false,reportDirty=false,baseLabel='',pendingWord='反映',commentBase=null;
document.getElementById('comment').addEventListener('input',()=>{updateDirty();resize()});
function updateDirty(initial){
  const n=draft.filter((v,i)=>v!==days[i].value).length;
  const ta=document.getElementById('comment');
  const cm=commentBase!==null&&ta.value!==commentBase;   // コメントを書き換えたか（先生用画面）
  const btn=document.getElementById('apply'),note=document.getElementById('note');
  btn.classList.toggle('dirty',n>0||cm);
  btn.classList.toggle('clean',n===0&&!cm&&reportDirty);
  btn.textContent=n>0?baseLabel+'（未'+pendingWord+' '+n+'日）':(cm?baseLabel+'（コメント未'+pendingWord+'）':(reportDirty?'✓ 反映済み（変更はありません）':baseLabel));
  note.classList.toggle('dirty',n>0);
  if(n>0)note.textContent='⚠ まだ'+pendingWord+'していない変更が '+n+'日分あります。選び終わったら下のボタンを押してください。';
  else if(!initial&&note.classList)note.textContent='';
  if(reportDirty&&!initial&&(n>0)!==sentDirty){
    sentDirty=n>0;
    send('streamlit:setComponentValue',{value:{kind:'dirty',dirty:sentDirty,version,token:Date.now().toString()+'-'+Math.random()},dataType:'json'});
  }
}
// まとめて変更（「別途決定」の日は変えない）。保存ボタンを押すまでは保存されない
function setAll(ng){days.forEach((d,i)=>{if(d.locked||!cells[i])return;
  const v=ng?(d.options.includes('全NG')?'全NG':'宿NG'):'OK';draft[i]=v;paint(cells[i],v)});updateDirty()}
document.getElementById('allNg').onclick=()=>setAll(true);document.getElementById('allOk').onclick=()=>setAll(false);
document.getElementById('apply').onclick=()=>{
  if(reportDirty&&!draft.some((v,i)=>v!==days[i].value))return;   // 変更がなければ何もしない
  send('streamlit:setComponentValue',{value:{kind:'apply',version,values:draft,comment:document.getElementById('comment').value,token:Date.now().toString()+'-'+Math.random()},dataType:'json'});
  document.getElementById('note').textContent='処理しています…';};
new ResizeObserver(resize).observe(document.body);send('streamlit:componentReady',{apiVersion:1});resize();
</script></body></html>"""


HOVER_TABLE_HTML = r"""<!doctype html><html lang="ja"><meta charset="utf-8"><style>
*{box-sizing:border-box}body{margin:0;font:14px system-ui,sans-serif;color:#243247;background:white}
.scroll{overflow-x:auto;border:1px solid #d6dce5;border-radius:6px}
table{border-collapse:separate;border-spacing:0;width:100%;min-width:720px;table-layout:fixed}
th,td{overflow-wrap:anywhere;padding:7px 9px;border-bottom:1px solid #e4e7ed;border-right:1px solid #e4e7ed;text-align:left}
th{white-space:nowrap;position:sticky;top:0;background:#f3f5f8;z-index:1}
.doctor{display:inline-block;max-width:100%;white-space:normal;border-radius:4px;padding:2px 4px;outline-offset:1px;cursor:default}
.doctor.match{background:#ffeb70!important;color:#111!important;outline:2px solid #bf7100;font-weight:800}
.doctor:focus-visible{outline:2px solid #bf7100}
.note{font-size:16px;line-height:1.7;font-weight:500;color:#243247;margin:0 0 12px;overflow-wrap:anywhere}
</style><body>
<p class="note">名前にカーソルを合わせると同じ医師を強調します。クリックするとカーソルを外しても強調表示が続き、同じ名前をもう一度クリックすると解除します。スマートフォンではタップで同じ操作ができます。</p>
<div class="scroll"><table><colgroup><col style="width:12%"><col style="width:10%"><col span="6" style="width:13%"></colgroup>
<thead><tr id="head"></tr></thead><tbody id="body"></tbody></table></div>
<script>
function notify(type,data={}){window.parent.postMessage({isStreamlitMessage:true,type,...data},'*');}
let lastHeight=0;
function resizeFrame(){const h=Math.ceil(document.body.getBoundingClientRect().height)+8;if(h!==lastHeight){lastHeight=h;notify('streamlit:setFrameHeight',{height:h});}}
let pinned=null;
function highlight(name){document.querySelectorAll('.doctor').forEach(el=>el.classList.toggle('match',name!==null&&el.dataset.doctor===name));}
function render(table){
  const head=document.getElementById('head');head.replaceChildren();
  table.columns.forEach(c=>{const th=document.createElement('th');th.textContent=c;head.append(th);});
  const body=document.getElementById('body');body.replaceChildren();
  table.rows.forEach(cells=>{
    const tr=document.createElement('tr');
    cells.forEach(cell=>{
      const td=document.createElement('td');td.style.cssText=cell.style;
      cell.parts.forEach((p,i)=>{
        if(i>0)td.append('、');
        if(p.doctor===null){td.append(p.text);return;}
        const s=document.createElement('span');
        s.className='doctor';s.tabIndex=0;s.dataset.doctor=p.doctor;s.textContent=p.text;
        s.addEventListener('mouseenter',()=>highlight(p.doctor));
        s.addEventListener('mouseleave',()=>highlight(pinned));
        s.addEventListener('focus',()=>highlight(p.doctor));
        s.addEventListener('blur',()=>highlight(pinned));
        s.addEventListener('click',()=>{pinned=pinned===p.doctor?null:p.doctor;highlight(pinned);});
        s.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();s.click();}});
        td.append(s);
      });
      tr.append(td);
    });
    body.append(tr);
  });
  highlight(pinned);requestAnimationFrame(resizeFrame);
}
window.addEventListener('message',e=>{if(e.data&&e.data.type==='streamlit:render')render(e.data.args.table);});
document.addEventListener('keydown',e=>{if(e.key==='Escape'){pinned=null;highlight(null);}});
new ResizeObserver(resizeFrame).observe(document.body);
if(document.fonts)document.fonts.ready.then(resizeFrame);
notify('streamlit:componentReady',{apiVersion:1});
</script></body></html>"""


def hover_schedule_data(df, shift_columns, doctors, color_style):
    """結果の表の中身をデータとして組み立てる（HTMLは固定なので部品は増えない）。"""
    doctor_set = set(doctors)
    rows = []
    for _, row in df.iterrows():
        cells = []
        for col in df.columns:
            value = str(row[col])
            if col in shift_columns:
                style = color_style(value)
                tokens = [x.strip() for x in re.split(r'[、,]', value)]
                parts = [{"text": t, "doctor": t if t in doctor_set else None} for t in tokens]
            else:
                style = 'color:#ff4b4b;font-weight:bold;' if row.get('平日/休日') == '休日' else ''
                parts = [{"text": value, "doctor": None}]
            cells.append({"style": style, "parts": parts})
        rows.append(cells)
    return {"columns": [str(c) for c in df.columns], "rows": rows}


# ==========================================
# 医師×日付の一覧表（決まった当直・当直NG・翌日PM duty）
# ==========================================
# 枠ごとの色（背景色, 文字色）
MATRIX_SHIFT_COLORS = {
    'A宿直': ('#1f3c74', '#ffffff'),
    'B宿直': ('#6f9a37', '#ffffff'),
    '外来宿直': ('#e3a1ab', '#3a1218'),
    'A日直': ('#c8d4e6', '#1f2d44'),
    'B日直': ('#d4e6bf', '#24361a'),
    '外来日直': ('#f6d2d7', '#4a1c22'),
}
MATRIX_NG_COLOR = ('#ffffff', '#c0392b')          # ×（当直NG）
MATRIX_PM_COLOR = ('#ffe9b3', '#7a4b00')          # PM（翌日PM dutyで宿直を外す日）
MATRIX_HOLIDAY_HEAD = '#c0392b'
MATRIX_WISH_FRAME = '#8B4A1C'   # 希望が通った当直の囲み線（茶色の太線）
# 右端の回数の列（病棟・外来・宿直・日直・平日・休日・合計）
MATRIX_WARD_SHIFTS = ('A日直', 'B日直', 'A宿直', 'B宿直')
MATRIX_OUTPATIENT_SHIFTS = ('外来宿直', '外来日直')
MATRIX_COUNT_COLUMNS = [('ward', '病棟'), ('outpatient', '外来'), ('night', '宿直'), ('day', '日直'), ('weekday', '平日'), ('holiday', '休日'), ('total', '合計')]
# 2列ずつのセットがひと目でわかるように色分けする（見出しの色, 数字欄の色）
MATRIX_COUNT_GROUP_COLORS = {
    'night': ('#cfe0f5', '#eef4fc'), 'day': ('#cfe0f5', '#eef4fc'),            # 宿直・日直：青
    'ward': ('#d5ead0', '#eef7ec'), 'outpatient': ('#d5ead0', '#eef7ec'),      # 病棟・外来：緑
    'weekday': ('#f8dcc4', '#fdf1e7'), 'holiday': ('#f8dcc4', '#fdf1e7'),      # 平日・休日：橙
    'total': ('#d9dde3', '#f3f4f6'),                                           # 合計：灰
}
MATRIX_COUNT_GROUP_STARTS = ('ward', 'night', 'weekday', 'total')  # セットの左端（区切り線を太くする）
MATRIX_SATURDAY_HEAD = '#1f6fbf'


def build_duty_matrix(df_result, staff_df, year, month, custom_holidays, next_month_special_holiday=False, comments=None):
    """
    医師ごと・日ごとのマスの内容を作る。
    優先順：決まった当直 ＞ ×（当直NG）＞ PM（翌日PM duty）。
    """
    num_days = len(df_result)
    days = []
    for d in range(1, num_days + 1):
        dt = datetime.date(year, month, d)
        hol = str(df_result.iloc[d - 1]['平日/休日']) == '休日'
        kind = 'holiday' if (hol and (dt.weekday() == 6 or jpholiday.is_holiday(dt) or d in custom_holidays)) else ('saturday' if dt.weekday() == 5 else 'weekday')
        # 当直が決まっていない枠（⚠️不足）
        shortage = []
        for s in RESULT_COLUMNS[2:]:
            m = re.search(r'⚠️不足\((\d+)名\)', str(df_result.iloc[d - 1][s]))
            if m:
                shortage.append({'slot': s, 'count': int(m.group(1))})
        days.append({'day': d, 'weekday': WEEKDAYS_JA[dt.weekday()], 'holiday': hol, 'kind': kind, 'shortage': shortage,
                     'ng_count': 0})   # その日にNG（翌日PM dutyで宿直を外す先生を含む）の先生の数

    # 備考は「備考」で始まる列から読む（列名の書き方の違いに対応）
    note_columns = [c for c in staff_df.columns if str(c).strip().startswith("備考")]
    rows = []
    for _, staff in staff_df.iterrows():
        name = str(staff['先生の名前'])
        ng = parse_ng_dict(staff.get(NG_COLUMN, ''), year, month)
        pm_days = pm_duty_weekdays(staff.get('翌日PM duty', ''))
        wish_days, wish_specific = parse_requests(staff.get(REQUEST_COLUMN, ''), year, month)
        cells = []
        weekday_count = holiday_count = ward_count = outpatient_count = night_count = day_count = 0
        for info in days:
            d = info['day']
            r = df_result.iloc[d - 1]
            # 日ごとのNG人数：その日にNG（平日は宿NG・全NG、休日は日NGも）か、翌日PM dutyで宿直を外す先生を数える
            if effective_ng(ng.get(d, 'OK'), info['holiday']) != 'OK' or pm_duty_restricted(
                    datetime.date(year, month, d), pm_days, year, month, custom_holidays, next_month_special_holiday):
                info['ng_count'] += 1
            assigned = [s for s in RESULT_COLUMNS[2:] if name in [x.strip() for x in re.split(r'[、,]', str(r[s]))]]
            if assigned:
                if info['holiday']:
                    holiday_count += len(assigned)
                else:
                    weekday_count += len(assigned)
                night_count += sum(1 for s in assigned if s in NIGHT_SHIFTS)
                day_count += sum(1 for s in assigned if s in DAY_SHIFTS)
                ward_count += sum(1 for s in assigned if s in MATRIX_WARD_SHIFTS)
                outpatient_count += sum(1 for s in assigned if s in MATRIX_OUTPATIENT_SHIFTS)
                # 希望が通った日（日付だけの希望はどの枠でも、枠指定の希望はその枠に入ったとき）
                wish = d in wish_days or any((d, s) in wish_specific for s in assigned)
                text = '/'.join(assigned)
                cells.append({'type': 'shift', 'text': text, 'shift': assigned[0], 'wish': wish})
                continue
            kind = ng.get(d)
            pm = pm_duty_restricted(datetime.date(year, month, d), pm_days, year, month, custom_holidays, next_month_special_holiday)
            # 翌日PM dutyの日にNGも付いている場合は、斜線の上に×を出す（斜線だけ＝NGは付いていない）
            if info['holiday'] and (pm or kind in ('全NG', '宿NG', '日NG')):
                # 休日は、上半分＝日直、下半分＝宿直に分けて、NGの側に×を出す（全NGは上下とも×）。
                # 翌日PM dutyで宿直を外す日は、下半分を斜線にする
                top = '×' if kind in ('全NG', '日NG') else ''
                bottom = '×' if kind in ('全NG', '宿NG') else ''
                cells.append({'type': 'pmhalf' if pm else 'half', 'top': top, 'bottom': bottom, 'text': '×' if (top or bottom) else ''})
            elif kind in ('全NG', '宿NG'):
                cells.append({'type': 'ng', 'text': '×', 'pm': pm})
            elif pm:
                cells.append({'type': 'pm', 'text': ''})
            else:
                cells.append({'type': 'empty', 'text': ''})
        rows.append({
            'name': name,
            'pm_duty': '・'.join(WEEKDAYS_JA[i] for i in pm_days),
            'note': "\n".join(v for v in (clean_text(staff.get(c, "")) for c in note_columns) if v and v.lower() not in ("nan", "none")),
            'comment': (comments or {}).get(name, ("", ""))[0],
            'comment_at': (comments or {}).get(name, ("", ""))[1],
            'cells': cells,
            'night': night_count, 'day': day_count, 'ward': ward_count, 'outpatient': outpatient_count,
            'weekday': weekday_count, 'holiday': holiday_count, 'total': weekday_count + holiday_count,
        })
    return days, rows


def slot_label(text, newline):
    """
    セルに入れる枠名の改行位置を決める。「外来宿直」は「外来／宿直」で区切り、
    同じ日に2つの当直がある場合（A日直/A宿直）は1つずつ改行する。
    """
    parts = []
    for name in str(text).split("/"):
        if name.startswith("外来") and len(name) > 2:
            name = "外来" + newline + name[2:]
        parts.append(name)
    return newline.join(parts)


def shortage_slots(days):
    """不足がある枠を、表の列と同じ順で返す。"""
    used = {x['slot'] for info in days for x in info['shortage']}
    return [s for s in RESULT_COLUMNS[2:] if s in used]


# 当直が多すぎる先生を目立たせる基準（この回数以上でアラート。0はアラートなし）。画面で変えられる
MATRIX_COUNT_ALERTS = {'night': 4, 'day': 0, 'ward': 0, 'outpatient': 0, 'weekday': 0, 'holiday': 3}
MATRIX_ALERT_LABELS = {'night': '宿直', 'day': '日直', 'ward': '病棟', 'outpatient': '外来', 'weekday': '平日', 'holiday': '休日', 'total': '合計'}


def count_is_alert(key, value, alerts=None):
    limit = (MATRIX_COUNT_ALERTS if alerts is None else alerts).get(key) or 0
    return limit > 0 and isinstance(value, (int, float)) and value >= limit


def count_alert_messages(rows, alerts=None):
    """アラートになる先生の一覧（表の上に出す文）。"""
    alerts = MATRIX_COUNT_ALERTS if alerts is None else alerts
    messages = []
    for key, _ in MATRIX_COUNT_COLUMNS:
        if not alerts.get(key):
            continue
        hits = [f"{r['name']}（{r[key]}回）" for r in rows if count_is_alert(key, r.get(key), alerts)]
        if hits:
            messages.append(f"{MATRIX_ALERT_LABELS[key]}{alerts[key]}回以上：" + "、".join(hits))
    return messages


def count_alert_text(alerts=None):
    """凡例用：「宿直4回以上・休日3回以上」"""
    alerts = MATRIX_COUNT_ALERTS if alerts is None else alerts
    parts = [f"{MATRIX_ALERT_LABELS[k]}{alerts[k]}回以上" for k, _ in MATRIX_COUNT_COLUMNS if alerts.get(k)]
    return "・".join(parts)


def count_cell_css(key, is_head):
    head_bg, body_bg = MATRIX_COUNT_GROUP_COLORS[key]
    css = f"background:{head_bg if is_head else body_bg};"
    if key in MATRIX_COUNT_GROUP_STARTS:
        css += "border-left:2px solid #7b8698;"
    return css


def memo_preview(row, length=7):
    """メモ欄の見出し：アイコン（💬コメント・📝備考）と、最初の数文字だけ。"""
    icons = ("💬" if row.get("comment") else "") + ("📝" if row.get("note") else "")
    text = (row.get("comment") or row.get("note") or "").replace("\n", " ")
    short = text if len(text) <= length else text[:length] + "…"
    return f"{icons} {short}".strip()


def memo_cell_html(row, tip_up=False):
    """メモ欄：ふだんは短く表示し、押すと全文が開く（スマートフォンでも使える）。"""
    if not row.get("comment") and not row.get("note"):
        return '<td class="memo"></td>'
    parts = []
    if row.get("comment"):
        parts.append(f'<b>💬 コメント</b>（{html_escape(row.get("comment_at", ""))}）<br>' + html_escape(row["comment"]).replace("\n", "<br>"))
    if row.get("note"):
        parts.append('<b>📝 備考</b><br>' + html_escape(row["note"]).replace("\n", "<br>"))
    full = "<br><br>".join(parts)
    # パソコン：カーソルを合わせるとすぐ横に全文（.tip）。スマートフォン（カーソルがない）：タップでその場に開く（details）
    # 表の下のほうの行は、全文を上向きに開いて表からはみ出さないようにする
    up = " up" if tip_up else ""
    return (f'<td class="memo"><details><summary>{html_escape(memo_preview(row))}</summary>'
            f'<div class="full">{full}</div></details><div class="tip{up}">{full}</div></td>')


def duty_matrix_html(days, rows, year, month, alerts=None):
    def head_color(info):
        return MATRIX_HOLIDAY_HEAD if info['kind'] == 'holiday' else (MATRIX_SATURDAY_HEAD if info['kind'] == 'saturday' else '#243247')

    css = """
<style>
.dm-wrap{overflow-x:auto;border:1px solid #cfd6e0;border-radius:6px;margin:4px 0 8px}
.dm{border-collapse:collapse;font:12px system-ui,sans-serif;white-space:nowrap;background:#fff;color:#243247}
.dm th,.dm td{border:1px solid #d8dde5;padding:0;text-align:center;height:26px}
.dm thead th{background:#f3f5f8;font-weight:700;padding:2px 3px;min-width:30px}
.dm .nm{position:sticky;left:0;background:#fff;text-align:left;padding:0 5px;font-weight:600;z-index:1;min-width:52px;white-space:nowrap}
.dm thead .nm{background:#f3f5f8;z-index:2}
.dm .pmcol{padding:0 3px;color:#566174;font-size:11px;min-width:34px}
.dm thead .pmcol{line-height:1.2}
.dm td.hol{background:#fdf1f1}
.dm .cell{display:flex;align-items:center;justify-content:center;height:100%;min-height:26px;font-weight:700;font-size:10px;padding:1px 2px;width:38px;margin:0 auto;white-space:normal;line-height:1.15;text-align:center;word-break:keep-all}
.dm .ng{color:#c0392b;font-size:14px}
.dm .pm{background:repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px)}
.dm td.pmhalf{position:relative;min-width:38px;background:repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px) left bottom/100% 50% no-repeat,#fdf1f1}
.dm td.pmhalf span{position:absolute;left:0;right:0;height:50%;display:flex;align-items:center;justify-content:center;color:#c0392b;font-weight:700;font-size:13px;line-height:1}
.dm td.pmhalf span.t{top:0} .dm td.pmhalf span.b{bottom:0}
.dm td.half{position:relative;min-width:38px}
.dm td.half span{position:absolute;left:0;right:0;height:50%;display:flex;align-items:center;justify-content:center;color:#c0392b;font-weight:700;font-size:13px;line-height:1}
.dm td.half span.t{top:0} .dm td.half span.b{bottom:0}
.dm td.half::after{content:"";position:absolute;left:5px;right:5px;top:50%;border-top:1px dashed #dcb8b8}
/* 行の境目を見やすくする：先生の行を1行おきに薄いグレーにする（休日の色・斜線などの上に薄く重ねる） */
/* 右端の回数の列（色をマスごとに指定している）まで同じ縞にするため !important を付ける */
.dm tbody tr:not(.ngcount):nth-child(odd) td{background-image:linear-gradient(rgba(36,64,95,.075),rgba(36,64,95,.075)) !important}
/* 休日の翌日PM dutyのマス（下半分が斜線）は、縞の行でも斜線を消さない */
.dm tbody tr:not(.ngcount):nth-child(odd) td.pmhalf{background-image:linear-gradient(rgba(36,64,95,.075),rgba(36,64,95,.075)),
  repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px) !important;
  background-size:100% 50%,100% 50% !important;background-position:left top,left bottom !important;background-repeat:no-repeat !important}
.dm tr.short.first td{border-top:2px solid #c0392b}
.dm tr.short td.nm{color:#c0392b}
.dm th.miss{background:#fde2e0}
.dm .sum{padding:0 6px;font-weight:700;min-width:34px}
.dm .memo{min-width:40px;width:60px;text-align:left;padding:0 4px;font-size:11px;vertical-align:middle;position:relative}
.dm thead th.memo{text-align:center;font-size:12px}
.dm .memo details summary{cursor:pointer;list-style:none;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:56px;color:#425268}
.dm .memo details summary::-webkit-details-marker{display:none}
.dm .memo details[open] summary{color:#1f3c74;font-weight:700}
.dm .memo .full{white-space:normal;width:220px;padding:4px 6px;margin:2px 0 4px;background:#f6f8fb;border:1px solid #d8dde5;border-radius:4px;line-height:1.45}
/* パソコン（カーソルがある端末）：カーソルを合わせた瞬間に全文を表示する */
.dm .memo .tip{display:none;position:absolute;left:calc(100% - 2px);top:0;z-index:20;width:260px;white-space:normal;text-align:left;
  padding:7px 9px;background:#fff;border:1px solid #c9d2de;border-radius:6px;box-shadow:0 6px 18px rgba(16,24,40,.18);font-size:12px;line-height:1.5;color:#243247}
.dm .memo .tip.up{top:auto;bottom:0}
@media (hover:hover){.dm .memo details .full{display:none}.dm .memo details summary{cursor:default}
  .dm td.memo:hover{z-index:20}.dm td.memo:hover .tip{display:block}}
@media (hover:none){.dm .memo .tip{display:none !important}}
.dm-legend{display:flex;flex-wrap:wrap;gap:10px;font-size:12px;margin:2px 0 6px;color:#243247}
.dm-legend span{display:inline-flex;align-items:center;gap:4px}
.dm-legend i{display:inline-block;min-width:34px;padding:1px 4px;border:1px solid #d8dde5;border-radius:3px;font-style:normal;font-weight:700;text-align:center;font-size:11px}
</style>"""
    legend = ['<div class="dm-legend">']
    for s, (bg, fg) in MATRIX_SHIFT_COLORS.items():
        legend.append(f'<span><i style="background:{bg};color:{fg}">{s}</i></span>')
    legend.append(f'<span><i style="box-shadow:inset 0 0 0 3px {MATRIX_WISH_FRAME}">&nbsp;&nbsp;&nbsp;</i>茶色の枠：希望日が通った当直</span>')
    legend.append('<span><i style="background:#fdf1f1">&nbsp;</i>薄い赤の列：土日祝・休日にした日</span>')
    legend.append('<span><i style="color:#c0392b">×</i>当直NG（平日は宿直NG）</span>')
    legend.append('<span><i style="color:#c0392b;background:#fdf1f1;line-height:1;padding:0 4px">×<br>&nbsp;</i>'
                  '休日のマス：上半分＝日直、下半分＝宿直。上の×＝日直NG、下の×＝宿直NG（上下とも×＝終日NG）</span>')
    legend.append('<span><i style="background:repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px)">&nbsp;</i>斜線：翌日PM dutyで宿直を外す日（×があればNGも付いている）</span>')
    legend.append('<span><i style="background:repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px) left bottom/100% 50% no-repeat,#fdf1f1">&nbsp;</i>'
                  '下半分だけ斜線：休日の翌日PM dutyの日。宿直は外し、日直には入れる（×の意味は休日のマスと同じ）</span>')
    legend.append('<span><i>💬📝</i>メモ欄：先生からのコメント・備考（カーソルを合わせると全文）</span>')
    legend.append('<span><i style="color:#c0392b;border-color:#c0392b">⚠️ 未決定</i>赤線の下：当直が決まっていない枠</span>')
    legend.append('<span><i style="background:#ffd2a1;color:#7a3a00">3</i>NG人数の行：その日にNGの先生の数（翌日PM dutyで宿直を外す先生を含む）</span>')
    if count_alert_text(alerts):
        legend.append('<span><i style="color:#c0392b;background:#ffd8d8;border-color:#e03131">⚠</i>右端の赤いマス：' + count_alert_text(alerts) + '</span>')
    legend.append('</div>')

    head = ['<tr><th class="nm">名前</th><th class="pmcol">翌日<br>PM duty</th><th class="memo">メモ</th>']
    for info in days:
        miss_class = ' class="miss"' if info['shortage'] else ''
        miss_class = f' class="c{info["day"]}{" miss" if info["shortage"] else ""}"'
        head.append(f'<th{miss_class} style="color:{head_color(info)}">{info.get("label", info["day"])}<br>{info["weekday"]}</th>')
    head.append(''.join(f'<th class="sum" style="{count_cell_css(key, True)}">{label}</th>' for key, label in MATRIX_COUNT_COLUMNS) + '</tr>')

    body = []
    # 日ごとのNG人数（翌日PM dutyを含む）。多い日ほど濃い色にする
    n_doctors = max(len(rows), 1)
    ng_tr = ['<tr class="ngcount"><td class="nm" colspan="3" style="background:#ffe2c2;color:#7a3a00;font-weight:800">NG人数（翌日PM duty含む）</td>']
    for info in days:
        ratio = info.get('ng_count', 0) / n_doctors
        bg = '#ffffff' if ratio == 0 else ('#ffefdc' if ratio < 0.25 else ('#ffd2a1' if ratio < 0.5 else '#ffad5c'))
        ng_tr.append(f'<td class="c{info["day"]}" style="background:{bg};text-align:center;font-weight:700;color:#7a3a00">{info.get("ng_count", 0)}</td>')
    ng_tr.append(f'<td class="sum" colspan="{len(MATRIX_COUNT_COLUMNS)}" style="font-size:11px;background:#ffe2c2;color:#7a3a00;font-weight:800">／{len(rows)}名中</td></tr>')
    body.append(''.join(ng_tr))
    for idx, row in enumerate(rows):
        tip_up = idx >= max(len(rows) - 6, 6)
        tr = [f'<tr><td class="nm">{html_escape(row["name"])}</td><td class="pmcol">{html_escape(row["pm_duty"])}</td>{memo_cell_html(row, tip_up)}']
        for info, cell in zip(days, row['cells']):
            col = f'c{info["day"]}'   # 列（日付）ごとの印。カーソルを合わせた列を色付けするのに使う
            td_class = f' class="{col} hol"' if info['holiday'] else f' class="{col}"'
            if cell['type'] == 'shift':
                bg, fg = MATRIX_SHIFT_COLORS.get(cell['shift'], ('#eee', '#000'))
                wish_css = f'box-shadow:inset 0 0 0 3px {MATRIX_WISH_FRAME},inset 0 0 0 4px #ffffff;' if cell.get('wish') else ''
                tr.append(f'<td class="{col}"><div class="cell" style="background:{bg};color:{fg};{wish_css}">{slot_label(cell["text"], "<br>")}</div></td>')
            elif cell['type'] == 'ng':
                pm_class = ' pm' if cell.get('pm') else ''
                tr.append(f'<td{td_class}><div class="cell ng{pm_class}">{cell["text"]}</div></td>')
            elif cell['type'] == 'pm':
                tr.append(f'<td class="{col}"><div class="cell pm"></div></td>')
            elif cell['type'] in ('pmhalf', 'half'):
                hol_class = ' hol' if cell['type'] == 'half' else ''
                tr.append(f'<td class="{col}{hol_class} {cell["type"]}"><span class="t">{cell["top"]}</span><span class="b">{cell["bottom"]}</span></td>')
            else:
                tr.append(f'<td{td_class}></td>')
        tr.append(''.join(
            (f'<td class="sum" style="{count_cell_css(key, False)}background:#ffd8d8;color:#c0392b;font-weight:800;box-shadow:inset 0 0 0 2px #e03131" '
             f'title="{MATRIX_ALERT_LABELS[key]}が{(alerts or MATRIX_COUNT_ALERTS).get(key)}回以上です">⚠{row[key]}</td>')
            if count_is_alert(key, row[key], alerts) else f'<td class="sum" style="{count_cell_css(key, False)}">{row[key]}</td>'
            for key, _ in MATRIX_COUNT_COLUMNS) + '</tr>')
        body.append(''.join(tr))
    # 当直が決まっていない枠：枠ごとに1行、その枠の色で表示
    for k, slot in enumerate(shortage_slots(days)):
        first = ' first' if k == 0 else ''
        tr = [f'<tr class="short{first}"><td class="nm">⚠️ 未決定</td><td></td><td></td>']
        bg, fg = MATRIX_SHIFT_COLORS[slot]
        slot_total = 0
        for info in days:
            count = sum(x['count'] for x in info['shortage'] if x['slot'] == slot)
            if count:
                slot_total += 1
                label = slot_label(slot, "<br>") + ("" if count == 1 else f"×{count}")
                tr.append(f'<td class="c{info["day"]}"><div class="cell" style="background:{bg};color:{fg};white-space:normal;line-height:1.15;font-size:10px">{label}</div></td>')
            else:
                tr.append(f'<td class="c{info["day"]}"></td>')
        tr.append(f'<td class="sum" colspan="{len(MATRIX_COUNT_COLUMNS)}" style="color:#c0392b">{slot_total}枠</td></tr>')
        body.append(''.join(tr))
    # カーソルを合わせたセルの「行」と「列」を薄い黄色で示す（縦横に長い表での見間違いを防ぐ）
    hl = "#fff1b8"
    hover_css = ("<style>"
                 f".dm tbody tr:hover td{{background-color:{hl} !important}}"
                 f".dm tbody tr:hover td.nm{{background-color:#ffe08a !important}}"
                 + "".join(f".dm:has(.c{info['day']}:hover) .c{info['day']}{{background-color:{hl} !important}}"
                           f".dm:has(.c{info['day']}:hover) th.c{info['day']}{{background-color:#ffe08a !important}}" for info in days)
                 + ".dm tbody td:hover{outline:2px solid #e0a800;outline-offset:-2px}"
                 "</style>")
    return css + hover_css + ''.join(legend) + f'<div class="dm-wrap"><table class="dm"><thead>{"".join(head)}</thead><tbody>{"".join(body)}</tbody></table></div>'


def html_escape(text):
    import html
    return html.escape(str(text))


def xlsx_label_size(label):
    """Excelのマスに入れる枠名の文字の大きさ。「外来宿直」など4文字以上は小さめにして、マスの幅に必ず収める
    （Excelの「縮小して全体を表示」だけに頼ると、表示によっては左どなりのマスに文字がはみ出すことがあるため）。"""
    longest = max((len(x) for x in str(label).split("\n")), default=0)
    return 6.5 if longest >= 4 else 8


def duty_matrix_xlsx(days, rows, year, month, alerts=None, title=None):
    """一覧表をExcelファイル（印刷向けの書式付き）にする。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side, GradientFill
    from openpyxl.styles.fills import Stop
    from openpyxl.formatting.rule import CellIsRule
    from openpyxl.comments import Comment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = f"{year}年{month}月" if title is None else "当直表"
    thin = Side(style='thin', color='B8C0CC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    def fill(color):
        return PatternFill('solid', fgColor=color.lstrip('#'))

    def stripe_color(rgb):
        """色に薄いグレー（画面の縞と同じ紺色を7.5%）を重ねた色。"""
        rgb = str(rgb)[-6:]
        base = [int(rgb[k:k + 2], 16) for k in (0, 2, 4)]
        return ''.join(f"{round(v * 0.925 + t * 0.075):02X}" for v, t in zip(base, (36, 64, 95)))
    # 翌日PM dutyの日：Excelでは斜線を使わず、薄い黄色の塗りで表す
    PM_COLOR = 'F7DC8A'
    PM_FILL = fill(PM_COLOR)
    # 休日の翌日PM dutyの日（上＝日直、下＝宿直）：上半分は休日の薄い色のまま、下半分を黄色にする
    # （Excelはマスの半分だけを塗れないため、途中から色が変わる塗りで表す）
    PM_HALF_FILL = GradientFill(type='linear', degree=90, stop=(Stop('FDF1F1', 0), Stop('FDF1F1', 0.45),
                                                                Stop(PM_COLOR, 0.6), Stop(PM_COLOR, 1)))

    ws.cell(row=1, column=1, value=f"{year}年{month}月分 当直表" if title is None else f"{title} 当直表").font = Font(bold=True, size=14)
    header_row, first_day_col = 3, 4
    ws.cell(row=header_row, column=1, value='名前')
    ws.cell(row=header_row, column=2, value='翌日PM duty')
    ws.cell(row=header_row, column=3, value='メモ')
    for i, info in enumerate(days):
        c = ws.cell(row=header_row, column=first_day_col + i, value=f"{info.get('label', info['day'])}\n{info['weekday']}")
        color = MATRIX_HOLIDAY_HEAD if info['kind'] == 'holiday' else (MATRIX_SATURDAY_HEAD if info['kind'] == 'saturday' else '#243247')
        c.font = Font(bold=True, color=color.lstrip('#'))
    sum_col = first_day_col + len(days)
    n_counts = len(MATRIX_COUNT_COLUMNS)
    # 休日の列の目印（回数の式で使う。印刷には出ないよう、この行は隠す）
    flag_row = header_row - 1
    for i, info in enumerate(days):
        ws.cell(row=flag_row, column=first_day_col + i, value='休' if info['holiday'] else '平')
    ws.row_dimensions[flag_row].hidden = True
    first_letter = get_column_letter(first_day_col)
    last_letter = get_column_letter(first_day_col + len(days) - 1)
    count_letter = {key: get_column_letter(sum_col + j) for j, (key, _) in enumerate(MATRIX_COUNT_COLUMNS)}

    def count_formula(key, r):
        """その先生の行の回数を数える式。枠名（A宿直など）は1つに「直」が1つ、「宿直」「日直」「外来」も1回ずつ入る。"""
        rng = f"{first_letter}{r}:{last_letter}{r}"
        def occurrences(word):
            return f'SUMPRODUCT((LEN({rng})-LEN(SUBSTITUTE({rng},"{word}","")))/{len(word)})'
        c = {k: f"{v}{r}" for k, v in count_letter.items()}
        if key == 'night':
            return "=" + occurrences("宿直")
        if key == 'day':
            return "=" + occurrences("日直")
        if key == 'outpatient':
            return "=" + occurrences("外来")
        if key == 'ward':
            return f"={c['night']}+{c['day']}-{c['outpatient']}"
        if key == 'holiday':
            flags = f"{first_letter}${flag_row}:{last_letter}${flag_row}"
            return f'=SUMPRODUCT((LEN({rng})-LEN(SUBSTITUTE({rng},"直","")))*({flags}="休"))'
        if key == 'weekday':
            return f"={c['night']}+{c['day']}-{c['holiday']}"
        return f"={c['night']}+{c['day']}"   # total
    for j, (_, label) in enumerate(MATRIX_COUNT_COLUMNS):
        ws.cell(row=header_row, column=sum_col + j, value=label)
    for col in range(1, sum_col + n_counts):
        c = ws.cell(row=header_row, column=col)
        c.alignment = center; c.border = border; c.fill = fill('#E6EBF2')
        if c.font is None or not c.font.bold:
            c.font = Font(bold=True)
    group_side = Side(style='medium', color='7B8698')
    red_side = Side(style='medium', color='E03131')
    wish_side = Side(style='thick', color=MATRIX_WISH_FRAME.lstrip('#').upper())
    wish_border = Border(left=wish_side, right=wish_side, top=wish_side, bottom=wish_side)

    def style_count_cell(c, key, is_head):
        head_bg, body_bg = MATRIX_COUNT_GROUP_COLORS[key]
        c.fill = fill(head_bg if is_head else body_bg)
        if key in MATRIX_COUNT_GROUP_STARTS:
            c.border = Border(left=group_side, right=thin, top=c.border.top, bottom=c.border.bottom)

    for j, (key, _) in enumerate(MATRIX_COUNT_COLUMNS):
        style_count_cell(ws.cell(row=header_row, column=sum_col + j), key, True)

    # 見出しのすぐ下の行：日ごとのNG人数（翌日PM dutyを含む）
    ng_row = header_row + 1
    ng_title = ws.cell(row=ng_row, column=1, value='NG人数（翌日PM duty含む）')
    ng_title.font = Font(bold=True, size=9, color='7A3A00')
    for col in range(1, first_day_col):
        ws.cell(row=ng_row, column=col).fill = fill('#FFE2C2')
    n_doctors = max(len(rows), 1)
    for i, info in enumerate(days):
        cnt = info.get('ng_count', 0)
        c = ws.cell(row=ng_row, column=first_day_col + i, value=cnt)
        ratio = cnt / n_doctors
        c.fill = fill('#FFFFFF' if ratio == 0 else ('#FFEFDC' if ratio < 0.25 else ('#FFD2A1' if ratio < 0.5 else '#FFAD5C')))
        c.font = Font(bold=True, color='7A3A00'); c.alignment = center; c.border = border
    for col in range(1, first_day_col):
        ws.cell(row=ng_row, column=col).border = border
    ws.cell(row=ng_row, column=sum_col, value=f"／{len(rows)}名中").font = Font(bold=True, size=9, color='7A3A00')
    for j in range(n_counts):
        ws.cell(row=ng_row, column=sum_col + j).fill = fill('#FFE2C2')
    ws.row_dimensions[ng_row].height = 20

    for r_idx, row in enumerate(rows, header_row + 2):
        ws.cell(row=r_idx, column=1, value=row['name']).font = Font(bold=True)
        ws.cell(row=r_idx, column=2, value=row['pm_duty'])
        if row.get('comment') or row.get('note'):
            # セルには短く表示し、全文はセルのコメント（マウスを乗せると表示）と「メモ全文」シートに入れる
            # セルにはメモの全文を1行で入れる（表示はセルの幅で切れるが、文字はすべて残る。コピーや検索もできる）
            icons = ("💬" if row.get('comment') else "") + ("📝" if row.get('note') else "")
            one_line = " ／ ".join(t.replace("\r", "").replace("\n", " ") for t in (row.get('comment'), row.get('note')) if t)
            memo = ws.cell(row=r_idx, column=3, value=icons + one_line)
            memo.font = Font(size=8)
            full = []
            if row.get('comment'):
                full.append(f"【コメント】{row['comment']}")
            if row.get('note'):
                full.append(f"【備考】{row['note']}")
            memo.comment = Comment("\n".join(full), "当直作成アプリ", width=260, height=120)
        for i, (info, cell) in enumerate(zip(days, row['cells'])):
            # 当直が1つの日は1行で入れる（「外来宿直」も改行しない。幅に入らない分は文字を少し小さくして収める）
            # 同じ日に2つ当直がある日は、枠名を1つずつ改行する（「外来宿直」の途中では改行しない）
            label = str(cell['text']).replace('/', '\n') if cell['type'] == 'shift' else slot_label(cell['text'], '\n')
            c = ws.cell(row=r_idx, column=first_day_col + i, value=label or None)
            if cell['type'] == 'shift':
                bg, fg = MATRIX_SHIFT_COLORS.get(cell['shift'], ('#EEEEEE', '#000000'))
                c.fill = fill(bg); c.font = Font(bold=True, size=xlsx_label_size(label), color=fg.lstrip('#'))
            elif cell['type'] == 'ng':
                c.font = Font(bold=True, color='C0392B')
                if cell.get('pm'):
                    c.fill = PM_FILL
                elif info['holiday']:
                    c.fill = fill('#FDF1F1')
            elif cell['type'] == 'pm':
                c.fill = PM_FILL
            elif cell['type'] == 'half':
                # 休日の日直だけNG・宿直だけNG：×を小さめにして、上（日直）・下（宿直）に寄せる
                c.fill = fill('#FDF1F1')
                c.font = Font(bold=True, size=8, color='C0392B')
                c.value = f"{cell.get('top') or ' '}\n{cell.get('bottom') or ' '}"   # 2行にして、×を上の行（日直）・下の行（宿直）に置く
                c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            elif cell['type'] == 'pmhalf':
                c.fill = PM_HALF_FILL
                c.font = Font(bold=True, size=8, color='C0392B')
                if cell.get('top') or cell.get('bottom'):
                    c.value = f"{cell.get('top') or ' '}\n{cell.get('bottom') or ' '}"
                c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
            elif info['holiday']:
                c.fill = fill('#FDF1F1')
        if row.get('comment') or row.get('note'):
            # 右どなりの日付のマスが空だと、メモの長い文字が当直表の上にはみ出して見えるので、空白を1つ入れて止める
            nxt = ws.cell(row=r_idx, column=first_day_col)
            if nxt.value in (None, ""):
                nxt.value = " "
        for j, (key, _) in enumerate(MATRIX_COUNT_COLUMNS):
            # 回数は式にする（Excelでマスに枠名を書き足したり消したりすると、右の回数も自動で変わる）
            ws.cell(row=r_idx, column=sum_col + j, value=count_formula(key, r_idx)).font = Font(bold=True)
        for col in range(1, sum_col + n_counts):
            c = ws.cell(row=r_idx, column=col)
            c.border = border
            c.alignment = center if col > 1 else Alignment(vertical='center')
        # メモ欄は1行で表示（折り返すと行が高くなるため折り返さない。はみ出す分はセルの幅で切れて見える）
        ws.cell(row=r_idx, column=3).alignment = Alignment(horizontal='left', vertical='center', wrap_text=False)
        # 希望日が通った当直をオレンジの太線で囲む
        for i, cell in enumerate(row['cells']):
            if cell.get('wish'):
                ws.cell(row=r_idx, column=first_day_col + i).border = wish_border
        for j, (key, _) in enumerate(MATRIX_COUNT_COLUMNS):
            style_count_cell(ws.cell(row=r_idx, column=sum_col + j), key, False)
            # アラート（回数が多い先生を赤く）は条件付き書式にする（Excelで直した回数にも合わせて色が変わる）
            limit = (MATRIX_COUNT_ALERTS if alerts is None else alerts).get(key) or 0
            if limit > 0:
                ref = f"{get_column_letter(sum_col + j)}{r_idx}"
                ws.conditional_formatting.add(ref, CellIsRule(
                    operator='greaterThanOrEqual', formula=[str(limit)],
                    fill=PatternFill(start_color='FFD8D8', end_color='FFD8D8', fill_type='solid'),
                    font=Font(bold=True, color='C0392B'),
                    border=Border(left=red_side, right=red_side, top=red_side, bottom=red_side)))
        # 1行で入れた当直のマスは、折り返さずに縮小して収める（行が縦長にならないように）
        for i, cell in enumerate(row['cells']):
            if cell['type'] == 'shift' and '/' not in str(cell['text']):
                ws.cell(row=r_idx, column=first_day_col + i).alignment = Alignment(
                    horizontal='center', vertical='center', wrap_text=False, shrink_to_fit=True)
        # 同じ日に2つ当直がある医師の行だけ、2段で表示できる高さにする
        max_lines = max([str(cell['text']).count('/') + 1 for cell in row['cells']
                         if cell['type'] == 'shift' and '/' in str(cell['text'])] or [1])
        ws.row_dimensions[r_idx].height = max(20, 11 * max_lines + 6)
        # 画面と同じく、先生の行を1行おきに薄いグレーにする（当直の枠の色はそのまま。休日・PM dutyなどの色は少し暗くする）
        if (r_idx - header_row - 2) % 2 == 1:
            shift_cols = {first_day_col + i for i, cell in enumerate(row['cells']) if cell['type'] == 'shift'}
            for col in range(1, sum_col + n_counts):
                if col in shift_cols:
                    continue
                c = ws.cell(row=r_idx, column=col)
                if isinstance(c.fill, GradientFill):
                    c.fill = GradientFill(type='linear', degree=90, stop=[
                        Stop(stripe_color(st.color.rgb), st.position) for st in c.fill.stop])
                elif c.fill is not None and c.fill.fill_type == 'solid':
                    c.fill = fill(stripe_color(c.fill.fgColor.rgb))
                else:
                    c.fill = fill(stripe_color('FFFFFF'))

    # 当直が決まっていない枠：枠ごとに1行、その枠の色で表示
    last_row = header_row + 1 + len(rows)
    red = Side(style='medium', color='C0392B')
    # 表と「未決定」の行の境目は、マスの枠線ではなく、細い赤い行で区切る
    # （マスに赤線を付けると、Excelで未決定の枠をドラッグして移したときに赤線までついて行くため）
    if shortage_slots(days):
        last_row += 1
        for col in range(1, sum_col + n_counts):
            ws.cell(row=last_row, column=col).fill = fill('#C0392B')
        ws.row_dimensions[last_row].height = 4
    for k, slot in enumerate(shortage_slots(days)):
        last_row += 1
        bg, fg = MATRIX_SHIFT_COLORS[slot]
        ws.cell(row=last_row, column=1, value='⚠️ 未決定').font = Font(bold=True, color='C0392B')
        slot_total = 0
        for i, info in enumerate(days):
            count = sum(x['count'] for x in info['shortage'] if x['slot'] == slot)
            c = ws.cell(row=last_row, column=first_day_col + i)
            c.border = border
            if count:
                slot_total += 1
                c.value = slot + ('' if count == 1 else f'×{count}')
                c.fill = fill(bg); c.font = Font(bold=True, size=xlsx_label_size(c.value), color=fg.lstrip('#'))
                ws.cell(row=header_row, column=first_day_col + i).fill = fill('#FDE2E0')
            c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=False, shrink_to_fit=True)
        # 残りの枠の数も式にする（Excelで枠を表の中へ移すと、自動で減る）
        rng = f"{get_column_letter(first_day_col)}{last_row}:{get_column_letter(first_day_col + len(days) - 1)}{last_row}"
        ws.cell(row=last_row, column=sum_col + n_counts - 1,
                value=f'=COUNTIF({rng},"*直*")&"枠"').font = Font(bold=True, color='C0392B')
        for col in list(range(1, first_day_col)) + list(range(sum_col, sum_col + n_counts)):
            c = ws.cell(row=last_row, column=col)
            c.border = border
            c.alignment = center if col > 1 else Alignment(vertical='center')
        ws.row_dimensions[last_row].height = 24

    # 凡例：見本のマスは表と同じ幅の日付の列に、表と同じ見た目で置き、説明はその右に書く
    legend_row = last_row + 2
    one_line = Alignment(horizontal='center', vertical='center', wrap_text=False, shrink_to_fit=True)
    middle = Alignment(vertical='center')   # 説明の文字は、見本のマスと高さをそろえて上下中央にする
    lg = ws.cell(row=legend_row, column=1, value='凡例'); lg.font = Font(bold=True); lg.alignment = middle
    ws.cell(row=legend_row, column=2, value='当直の枠（色）').alignment = middle
    for k, (label, (bg, fg)) in enumerate(MATRIX_SHIFT_COLORS.items()):
        c = ws.cell(row=legend_row, column=first_day_col + k * 2, value=label)
        c.fill = fill(bg); c.font = Font(bold=True, size=xlsx_label_size(label), color=fg.lstrip('#')); c.alignment = one_line; c.border = border
    ws.row_dimensions[legend_row].height = 20

    def legend_line(offset, text):
        r = legend_row + offset
        ws.cell(row=r, column=first_day_col + 1, value=text).alignment = middle
        ws.row_dimensions[r].height = 20
        sample = ws.cell(row=r, column=first_day_col)
        sample.border = border; sample.alignment = one_line
        return sample

    s1 = legend_line(1, '×：当直NG（平日は宿直NG）。休日のマスは上＝日直、下＝宿直で、上の×＝日直NG、下の×＝宿直NG（上下とも×＝終日NG）')
    s1.value = '×'; s1.font = Font(bold=True, color='C0392B')
    legend_line(2, '黄色：翌日PM dutyで宿直を外す日（×があればNGも付いている）').fill = PM_FILL
    legend_line(4, '薄い赤の列：土日祝・休日にした日').fill = fill('#FDF1F1')
    s4 = legend_line(5, '茶色の太い枠：希望日が通った当直')
    # 枠が見えやすいよう、淡い色の枠名を見本にする
    sample_slot = 'B日直' if 'B日直' in MATRIX_SHIFT_COLORS else next(iter(MATRIX_SHIFT_COLORS))
    first_bg, first_fg = MATRIX_SHIFT_COLORS[sample_slot]
    s4.value = sample_slot; s4.fill = fill(first_bg)
    s4.font = Font(bold=True, size=8, color=first_fg.lstrip('#')); s4.border = wish_border
    s5 = legend_line(6, '赤線より下の「⚠️ 未決定」の行：当直が決まっていない枠（その枠の色で表示）')
    s5.border = Border(left=thin, right=thin, bottom=thin, top=red)
    s5.value = '未決定'; s5.font = Font(bold=True, size=8, color='C0392B')
    s6 = legend_line(7, 'メモ：💬＝先生からのコメント、📝＝備考（セルにカーソルを乗せると全文。「メモ全文」シートにも一覧）')
    s6.value = '💬📝'; s6.border = Border()
    s8 = legend_line(8, 'NG人数の行（見出しのすぐ下）：その日にNGの先生の数（翌日PM dutyで宿直を外す先生を含む）')
    s8.value = 3; s8.fill = fill('#FFD2A1'); s8.font = Font(bold=True, color='7A3A00')
    legend_line(3, '下半分だけ黄色：休日の翌日PM dutyの日。宿直は外し、日直には入れる（×の意味は休日のマスと同じ）').fill = PM_HALF_FILL
    if count_alert_text(alerts):
        s7 = legend_line(9, '右端の赤い枠の回数：' + count_alert_text(alerts))
        s7.value = next(v for k, v in alerts.items() if v) if alerts else 4
        s7.fill = fill('#FFD8D8'); s7.font = Font(bold=True, color='C0392B')
        s7.border = Border(left=red_side, right=red_side, top=red_side, bottom=red_side)

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 11
    for i in range(len(days)):
        ws.column_dimensions[get_column_letter(first_day_col + i)].width = 5.2
    for j in range(n_counts):
        ws.column_dimensions[get_column_letter(sum_col + j)].width = 5
    ws.row_dimensions[header_row].height = 30
    ws.freeze_panes = ws.cell(row=header_row + 2, column=first_day_col)
    # A3横向き・1ページに収めて印刷
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    # 2枚目のシート：コメント・備考の全文（印刷して確認する用）
    memo_rows = [r for r in rows if r.get('comment') or r.get('note')]
    if memo_rows:
        ws2 = wb.create_sheet("メモ全文")
        ws2.cell(row=1, column=1, value=(f"{year}年{month}月" if title is None else title) + " 先生からのコメント・備考").font = Font(bold=True, size=13)
        for j, label in enumerate(["名前", "先生からのコメント", "コメントの保存日時", "医師条件の備考"], 1):
            c = ws2.cell(row=3, column=j, value=label)
            c.font = Font(bold=True); c.fill = fill('#E6EBF2'); c.border = border
        for i, r in enumerate(memo_rows, 4):
            for j, v in enumerate([r['name'], r.get('comment', ''), r.get('comment_at', ''), r.get('note', '')], 1):
                c = ws2.cell(row=i, column=j, value=v or None)
                c.border = border; c.alignment = Alignment(vertical='top', wrap_text=True)
        for col, w in zip("ABCD", [14, 60, 20, 40]):
            ws2.column_dimensions[col].width = w
        ws2.page_setup.orientation = 'landscape'
        ws2.page_setup.paperSize = ws2.PAPERSIZE_A4
        ws2.sheet_properties.pageSetUpPr.fitToPage = True
        ws2.page_setup.fitToWidth = 1
        ws2.page_setup.fitToHeight = 0

    buffer = io.BytesIO()
    wb.save(buffer)
    return buffer.getvalue()


# ==========================================
# 入力チェック
# ==========================================
def validate_staff_inputs(df, year, month):
    df = df.copy()
    errors = []
    if '先生の名前' not in df:
        return df, ['医師条件CSVに「先生の名前」の列がありません。']
    df['先生の名前'] = df['先生の名前'].map(clean_text)
    blank = df['先生の名前'].eq('')
    for i, row in df[blank].iterrows():
        if any(clean_text(v) for k, v in row.items() if k != '先生の名前'):
            errors.append(f'医師条件 {i+1}行目：条件が入力されていますが医師名が空欄です。')
    df = df[~blank].reset_index(drop=True)
    for name in df.loc[df['先生の名前'].duplicated(), '先生の名前'].unique():
        errors.append(f'医師名「{name}」が重複しています。')
    for col, default in NUMERIC_DEFAULTS.items():
        if col not in df:
            df[col] = default
    for col in [NG_COLUMN, REQUEST_COLUMN, '翌日PM duty'] + list(GROUP_CAP_COLUMNS):
        if col not in df:
            df[col] = ''
    for col in GROUP_CAP_COLUMNS:
        df[col] = df[col].astype(object)   # 空欄（制限なし）と数字が混ざるため
    for i, row in df.iterrows():
        name = row['先生の名前']
        if any(c in name for c in ',、') or name == '-' or '⚠️不足' in name:
            errors.append(f'医師名「{name}」：カンマ・読点・不足表示・単独の「-」は使用できません。')
        for col, default in NUMERIC_DEFAULTS.items():
            raw = row[col]
            try:
                n = default if not clean_text(raw) else float(raw)
                if not math.isfinite(n) or n < 0 or n != int(n) or n > 1000000:
                    raise ValueError()
                df.at[i, col] = int(n)
            except (TypeError, ValueError, OverflowError):
                errors.append(f'{name}：{col}は0〜1000000の整数にしてください。')
        try:
            if float(df.at[i, '月間最小回数']) > float(df.at[i, '月間最大回数']):
                errors.append(f'{name}：月間最小回数は月間最大回数以下にしてください。')
        except (ValueError, TypeError):
            pass
        # 宿直上限・日直上限は空欄なら制限なし
        for col in GROUP_CAP_COLUMNS:
            raw = clean_text(row.get(col, ''))
            if raw.lower() in ('', 'nan', 'none', '<na>'):
                df.at[i, col] = ''
                continue
            try:
                n = float(raw)
                if not math.isfinite(n) or n < 0 or n != int(n) or n > 1000:
                    raise ValueError()
                df.at[i, col] = int(n)
            except (TypeError, ValueError, OverflowError):
                errors.append(f'{name}：{col}は0以上の整数にしてください（空欄なら制限なし）。')
        for col, ng in [(NG_COLUMN, True), (REQUEST_COLUMN, False)]:
            try:
                items = checked_day_items(row[col], year, month, ng)
                df.at[i, col] = ','.join(str(d) if kind is None or (ng and kind == '全NG') else f'{d}:{kind}' for d, kind in items)
            except ValueError as e:
                errors.append(f'{name}：{col.split("(")[0]} {e}')
        weekdays = clean_text(row['翌日PM duty']).replace('，', ',')
        if any(w.strip() not in WEEKDAYS_JA for w in weekdays.split(',') if w.strip()):
            errors.append(f'{name}：翌日PM dutyは「水,木」のように曜日を入力してください。')
    return df, errors


def validate_fixed_inputs(df, doctors, year, month, holidays=None):
    """確定当直の入力チェック。holidays を渡すと、対象月の平日に日直が入っていないかも確認する。"""
    errors = []
    if df is None:
        return None, errors
    df = df.copy()
    if '日付' not in df:
        return df, ['確定当直CSVに「日付」の列がありません。']
    unknown = [c for c in df if c not in ['日付', '平日/休日'] + ALL_SHIFT_TYPES]
    for c in unknown:
        if df[c].map(clean_text).ne('').any():
            errors.append(f'確定当直：未対応の列「{c}」に入力があります。枠名を確認してください。')
    for i, row in df.iterrows():
        raw = clean_text(row['日付'])
        populated = any(clean_text(row.get(s, '')) not in ('', '-') for s in ALL_SHIFT_TYPES)
        if not raw and not populated:
            continue
        dt = parse_shift_date(raw, year, month)
        if dt is None:
            errors.append(f'確定当直 {i+1}行目：日付「{raw}」を確認してください。「11/1」または「2026/11/1」の形で入力します。')
            continue
        df.at[i, '日付'] = dt.isoformat()
        for s in ALL_SHIFT_TYPES:
            val = clean_text(row.get(s, ''))
            if val in ('', '-'):
                continue
            if (holidays is not None and s in DAY_SHIFTS and (dt.year, dt.month) == (year, month)
                    and not is_holiday_date(dt, year, month, holidays)):
                errors.append(f'確定当直 {dt.month}/{dt.day} {s}：平日には日直がありません。この日に日直を設ける場合は、上のカレンダーで「休日にする」にチェックを入れてください。')
            names = list(dict.fromkeys(n.strip() for n in re.split('[、,]', val)))
            for name in names:
                if name not in doctors:
                    errors.append(f'確定当直 {dt.month}/{dt.day} {s}：医師名「{name}」が名簿と一致しません。')
            df.at[i, s] = '、'.join(names)
    return df, errors


def input_signature(year, month, staff, holidays, multi, fixed, next_month_special_holiday=False, shortage_priority=None):
    payload = [year, month, bool(next_month_special_holiday), staff.to_csv(index=False), sorted(holidays), sorted((d, s, c) for (d, s), c in multi.items()), fixed.to_csv(index=False),
               sorted(normalize_shortage_priority(shortage_priority).items())]
    return hashlib.sha256(json.dumps(payload, ensure_ascii=False).encode()).hexdigest()


# ==========================================
# 人数が足りないときに、どの枠を優先して埋めるか
# 数字が小さいほど優先。同じ数字の枠は同順位として扱う。
# ==========================================
DEFAULT_SHORTAGE_PRIORITY = {'A宿直': 1, 'B宿直': 1, 'A日直': 1, 'B日直': 1, '外来宿直': 2, '外来日直': 3}


def normalize_shortage_priority(priority):
    """未指定や不正な値は初期設定に戻す。"""
    result = dict(DEFAULT_SHORTAGE_PRIORITY)
    for s, rank in (priority or {}).items():
        if s in result:
            try:
                rank = int(rank)
            except (TypeError, ValueError):
                continue
            if 1 <= rank <= len(ALL_SHIFT_TYPES):
                result[s] = rank
    return result


def shortage_priority_groups(priority):
    """優先度の高い順に、同順位の枠をまとめたリストを返す。例：[['A宿直','B宿直'], ['外来宿直']]"""
    priority = normalize_shortage_priority(priority)
    ranks = sorted(set(priority.values()))
    return [[s for s in ALL_SHIFT_TYPES if priority[s] == r] for r in ranks]


def shortage_priority_text(priority):
    groups = shortage_priority_groups(priority)
    if len(groups) == 1:
        return "すべての枠を同順位として、不足の合計が少なくなる方針で作成しました。"
    order = "、".join(f"第{i}優先：{'・'.join(g)}" for i, g in enumerate(groups, 1))
    return f"不足を減らす優先順位（{order}）に沿って作成しました。条件によっては優先度の高い枠にも不足が残ります。"


# ==========================================
# 作成結果の点検（確定指定で条件を超えた箇所などを知らせる）
# ==========================================
def audit_schedule(result, staff, year, month, holidays, multi, past, future, next_month_special_holiday=False):
    warnings = []
    for _, row in staff.iterrows():
        name = row['先生の名前']
        worked = []; counts = dict.fromkeys(ALL_SHIFT_TYPES, 0); hol_count = 0
        ng = parse_ng_dict(row.get(NG_COLUMN, ''), year, month)
        pm_weekdays = pm_duty_weekdays(row.get('翌日PM duty', ''))
        for _, r in result.iterrows():
            dt = parse_shift_date(r['日付'], year, month)
            for slot in ALL_SHIFT_TYPES:
                if name not in [n.strip() for n in re.split('[、,]', str(r[slot]))]:
                    continue
                worked.append(dt); counts[slot] += 1
                hol_count += int(r['平日/休日'] == '休日')
                kind = ng.get(dt.day, 'OK')
                night = slot in NIGHT_SHIFTS
                if kind == '全NG' or (kind == '宿NG' and night) or (kind == '日NG' and not night):
                    warnings.append(f'{name}：{dt.month}/{dt.day} {slot}はNG指定より確定指定を優先しました。')
                if night and pm_duty_restricted(dt, pm_weekdays, year, month, holidays, next_month_special_holiday):
                    warnings.append(f'{name}：{dt.month}/{dt.day} {slot}は翌日PM dutyの制限より確定指定を優先しました。')
        total = sum(counts.values())
        if total < int(row['月間最小回数']):
            warnings.append(f'{name}：月間最小回数の目標{int(row["月間最小回数"])}回に対し、{total}回です。')
        group_items = [(col, sum(counts[s] for s in slots), int(row[col])) for col, slots in GROUP_CAP_COLUMNS.items()
                       if clean_text(row.get(col, '')) not in ('', 'nan')]
        for label, value, cap in [('月間最大回数', total, int(row['月間最大回数'])), ('休日最大回数', hol_count, int(row['休日最大回数']))] + group_items + [(s + '上限', counts[s], int(row[s + '上限'])) for s in ALL_SHIFT_TYPES]:
            if value > cap:
                warnings.append(f'{name}：確定指定に伴い{label}を超えています（設定{cap}回／結果{value}回）。')
        dates = sorted(set(worked) | set((past or {}).get(name, [])) | set((future or {}).get(name, [])))
        for a, b in zip(dates, dates[1:]):
            if (a.year, a.month) != (year, month) and (b.year, b.month) != (year, month): continue
            gap = (b - a).days - 1
            if gap < int(row['最低空ける日数']):
                warnings.append(f'{name}：確定勤務同士の間隔が設定未満です（{a.month}/{a.day}→{b.month}/{b.day}、空き{gap}日／設定{int(row["最低空ける日数"])}日）。確定勤務を維持しました。')
    for _, r in result.iterrows():
        d = parse_shift_date(r['日付'], year, month).day
        for s in ALL_SHIFT_TYPES:
            names = [n.strip() for n in re.split('[、,]', str(r[s])) if n.strip() not in ('', '-') and '⚠️不足' not in n]
            need = multi.get((d, s), 1)
            if len(names) > need:
                warnings.append(f'{month}/{d} {s}：設定{need}名に対して{len(names)}名を配置しています。')
    return warnings


# ==========================================
# 当直計算ロジック
# ==========================================
def add_type_cap(model, worked, forced_vars, cap, bound):
    # 上限を緩めるのは、その枠で実際に確定した回数が上限を超える分だけ。
    extra = model.NewIntVar(0, bound, 'fixed_type_extra')
    model.AddMaxEquality(extra, [0, sum(forced_vars) - cap])
    model.Add(sum(worked) <= cap + extra)


def _generate_shift_core(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None, locked_days=()):
    locked_days = set(locked_days)   # 「別途決定」の日：別の方法で当番を決めるので、ここでは割り当てない（確定当直だけ残す）
    _, num_days = calendar.monthrange(target_year, target_month)
    days = range(1, num_days + 1)

    def is_holiday(d):
        return is_holiday_date(datetime.date(target_year, target_month, d), target_year, target_month, custom_holidays)

    def safe_int(val, default_val):
        if pd.isna(val): return default_val
        try:
            return int(float(val))
        except (ValueError, TypeError):
            return default_val

    doctors = staff_df['先生の名前'].astype(str).tolist()
    ng_days_dict = {}
    req_days = {}
    req_specific = {}
    req_priority = {}
    hard_weekdays = {}
    min_intervals = {}
    min_shifts_total = {}
    max_shifts_total = {}
    max_hol_shifts_per_doc = {}
    max_shifts_per_type = {}
    group_caps = {}

    absolute_req_days = {doc: [] for doc in doctors}
    absolute_req_specific = {doc: [] for doc in doctors}
    past_worked_dates = {doc: [] for doc in doctors}
    future_worked_dates = {doc: [] for doc in doctors}

    # --- 確定当直（先月・今月・来月）の読み取り ---
    fixed_specific = {doc: set() for doc in doctors}
    if fixed_df is not None:
        for _, row in fixed_df.iterrows():
            date_obj = parse_shift_date(row.get('日付', ''), target_year, target_month)
            if date_obj is None:
                continue
            for s_type in ALL_SHIFT_TYPES:
                if s_type in row and pd.notna(row[s_type]):
                    for doc_val in re.split(r'[、,]+', str(row[s_type])):
                        doc_val = doc_val.strip()
                        if doc_val not in doctors:
                            continue
                        if (date_obj.year, date_obj.month) == (target_year, target_month):
                            absolute_req_specific[doc_val].append((date_obj.day, s_type))
                            fixed_specific[doc_val].add((date_obj.day, s_type))
                        elif date_obj < datetime.date(target_year, target_month, 1):
                            past_worked_dates[doc_val].append(date_obj)
                        else:
                            future_worked_dates[doc_val].append(date_obj)

    # --- 医師条件の読み取り ---
    for _, row in staff_df.iterrows():
        doc = str(row['先生の名前'])
        hard_weekdays[doc] = pm_duty_weekdays(row.get('翌日PM duty', ''))
        ng_days_dict[doc] = parse_ng_dict(row.get(NG_COLUMN, ''), target_year, target_month)
        req_days[doc], req_specific[doc] = parse_requests(row.get(REQUEST_COLUMN, ''), target_year, target_month)
        req_days[doc] = [d for d in req_days[doc] if d not in locked_days]
        req_specific[doc] = [(d, s) for d, s in req_specific[doc] if d not in locked_days]
        req_priority[doc] = safe_int(row.get('希望優先度(数字が大きいほど優先)'), 1)
        min_intervals[doc] = safe_int(row.get('最低空ける日数'), 5)
        min_shifts_total[doc] = safe_int(row.get('月間最小回数'), 0)
        max_shifts_total[doc] = safe_int(row.get('月間最大回数'), 5)
        max_hol_shifts_per_doc[doc] = safe_int(row.get('休日最大回数'), 4)
        max_shifts_per_type[doc] = {s: safe_int(row.get(s + '上限'), 2) for s in ALL_SHIFT_TYPES}
        group_caps[doc] = {col: safe_int(row.get(col), None) for col in GROUP_CAP_COLUMNS}

    # --- その日に存在する枠（平日は宿直のみ、休日は日直・宿直） ---
    daily_active_shifts = {d: (NIGHT_SHIFTS + DAY_SHIFTS) if is_holiday(d) else list(NIGHT_SHIFTS) for d in days}

    # 存在しない枠の指定はエラーにする（優先度に関係なく同じ扱い）
    invalid_requests = []
    for doc in doctors:
        for d, slot in req_specific[doc]:
            if slot not in daily_active_shifts[d]:
                invalid_requests.append(f'{doc}：{d}日の{slot}は設定されていません。平日に日直を希望する場合は、上のカレンダーでその日を「休日にする」にしてください。')
        for d, slot in fixed_specific[doc]:
            if slot not in daily_active_shifts[d]:
                invalid_requests.append(f'{doc}：確定当直の{d}日の{slot}は、この日には設定されていない枠です。')
    if invalid_requests:
        return None, False, list(dict.fromkeys(invalid_requests)), None, None

    # --- 優先度100以上の希望を確定扱いにする ---
    for doc in doctors:
        if req_priority[doc] >= 100:
            absolute_req_days[doc].extend(req_days[doc])
            absolute_req_specific[doc].extend(req_specific[doc])
        absolute_req_specific[doc] = sorted(set(absolute_req_specific[doc]))
        specified_days = {d for d, s in absolute_req_specific[doc]}
        absolute_req_days[doc] = sorted(set(absolute_req_days[doc]) - specified_days)
        past_worked_dates[doc] = sorted(set(past_worked_dates[doc]))
        future_worked_dates[doc] = sorted(set(future_worked_dates[doc]))
        # 確定した日はNG指定より確定を優先する
        all_abs_dates = set(absolute_req_days[doc]) | specified_days
        ng_days_dict[doc] = {d: v for d, v in ng_days_dict[doc].items() if d not in all_abs_dates}

    def required_count(d, s):
        fixed_docs_count = sum(1 for doc in doctors if (d, s) in absolute_req_specific[doc])
        if d in locked_days:
            return fixed_docs_count
        return max(multi_slots_dict.get((d, s), 1), fixed_docs_count)

    def total_cap(doc):
        return max(max_shifts_total[doc], len(absolute_req_days[doc]) + len(absolute_req_specific[doc]))

    # ------------------------------------------------------------------
    # 通常計算と不足許容計算で共通の条件。
    # ここを直せば両方の計算に反映される（片方だけ直し忘れることがない）。
    # ------------------------------------------------------------------
    def add_common_constraints(model, x):
        worked_all = {}
        holiday_worked = {}
        blocked_by_ng = {"全NG": ALL_SHIFT_TYPES, "日NG": DAY_SHIFTS, "宿NG": NIGHT_SHIFTS}
        for doc in doctors:
            specific_days = {d for d, _ in absolute_req_specific[doc]}
            fixed_dates = set(absolute_req_days[doc]) | specific_days

            # 1日に入れる勤務は原則1つ（確定で複数ある日だけ例外）
            for d in days:
                fixed_count = sum(1 for sd, _ in absolute_req_specific[doc] if sd == d)
                model.Add(sum(x[(d, doc, s)] for s in daily_active_shifts[d]) <= max(1, fixed_count))

            # 「別途決定」の日は、確定当直以外は入れない
            for d in locked_days:
                if d in days:
                    for s in daily_active_shifts[d]:
                        if (d, s) not in absolute_req_specific[doc] and d not in absolute_req_days[doc]:
                            model.Add(x[(d, doc, s)] == 0)

            # NG日（全NG / 日NG / 宿NG）
            for d, ng_type in ng_days_dict[doc].items():
                for s in daily_active_shifts[d]:
                    if s in blocked_by_ng.get(ng_type, []):
                        model.Add(x[(d, doc, s)] == 0)

            # 翌日PM duty：宿直のみ制限。翌日が休日なら制限しない。確定日は対象外。
            for d in days:
                if d in fixed_dates:
                    continue
                if pm_duty_restricted(datetime.date(target_year, target_month, d), hard_weekdays[doc], target_year, target_month, custom_holidays, next_month_special_holiday):
                    for s in NIGHT_SHIFTS:
                        model.Add(x[(d, doc, s)] == 0)

            # 確定（枠指定なしの日は、その日のいずれか1枠）
            for d in absolute_req_days[doc]:
                if d not in specific_days:
                    model.AddExactlyOne(x[(d, doc, s)] for s in daily_active_shifts[d])
            for d, s in absolute_req_specific[doc]:
                model.Add(x[(d, doc, s)] == 1)

            # 枠ごとの上限
            for s_type in ALL_SHIFT_TYPES:
                active_days = [d for d in days if s_type in daily_active_shifts[d]]
                if not active_days:
                    continue
                worked = [x[(d, doc, s_type)] for d in active_days]
                forced_vars = [x[(d, doc, s_type)] for d in active_days if (d, s_type) in absolute_req_specific[doc] or d in absolute_req_days[doc]]
                add_type_cap(model, worked, forced_vars, max_shifts_per_type[doc][s_type], num_days)

            # 宿直上限・日直上限（A・B・外来を合わせた回数。空欄なら制限なし）
            for col, slots in GROUP_CAP_COLUMNS.items():
                cap = group_caps[doc].get(col)
                if cap is None:
                    continue
                worked = [x[(d, doc, s)] for d in days for s in daily_active_shifts[d] if s in slots]
                if not worked:
                    continue
                forced_vars = [x[(d, doc, s)] for d in days for s in daily_active_shifts[d] if s in slots
                               and ((d, s) in absolute_req_specific[doc] or d in absolute_req_days[doc])]
                add_type_cap(model, worked, forced_vars, cap, num_days)

            # 月間上限
            all_vars = [x[(d, doc, s)] for d in days for s in daily_active_shifts[d]]
            model.Add(sum(all_vars) <= total_cap(doc))
            worked_all[doc] = all_vars

            # 休日上限
            hol_vars = [x[(d, doc, s)] for d in days if is_holiday(d) for s in daily_active_shifts[d]]
            abs_hol_count = sum(1 for d in absolute_req_days[doc] if is_holiday(d)) + sum(1 for d, _ in absolute_req_specific[doc] if is_holiday(d))
            model.Add(sum(hol_vars) <= max(max_hol_shifts_per_doc[doc], abs_hol_count))
            holiday_worked[doc] = hol_vars

        # 勤務間隔
        add_interval_constraints(
            model=model, shifts=x, doctors=doctors, daily_active_shifts=daily_active_shifts,
            num_days=num_days, target_year=target_year, target_month=target_month,
            min_intervals=min_intervals, past_worked_dates=past_worked_dates,
            future_worked_dates=future_worked_dates, absolute_req_days=absolute_req_days,
            absolute_req_specific=absolute_req_specific,
        )
        return worked_all, holiday_worked

    def build_schedule(value_of, missing_of=None):
        """計算結果を表（DataFrame）にする。不足があれば「⚠️不足(n名)」を書き込む。"""
        rows = []
        for d in days:
            date_obj = datetime.date(target_year, target_month, d)
            row = {"日付": f"{target_month}/{d}({WEEKDAYS_JA[date_obj.weekday()]})", "平日/休日": "休日" if is_holiday(d) else "平日"}
            for s in ALL_SHIFT_TYPES:
                row[s] = "-"
            for s in daily_active_shifts[d]:
                names = [doc for doc in doctors if value_of((d, doc, s)) == 1]
                missing = missing_of((d, s)) if missing_of else 0
                if missing > 0:
                    names.append(f"⚠️不足({missing}名)")
                if names:
                    row[s] = "、".join(names)
            rows.append(row)
        return pd.DataFrame(rows, columns=RESULT_COLUMNS)

    # ==================================================================
    # 通常計算
    # ==================================================================
    model = cp_model.CpModel()
    shifts = {(d, doc, s): model.NewBoolVar(f'shift_d{d}_{doc}_{s}') for d in days for doc in doctors for s in daily_active_shifts[d]}
    objective_terms = []

    over_caps = {}
    for d in days:
        for s in daily_active_shifts[d]:
            over_caps[(d, s)] = model.NewIntVar(0, len(doctors), f'over_cap_d{d}_{s}')
            model.Add(sum(shifts[(d, doc, s)] for doc in doctors) == required_count(d, s) + over_caps[(d, s)])
            objective_terms.append(over_caps[(d, s)] * -50000)

    worked_all, holiday_worked = add_common_constraints(model, shifts)

    # 月間最小回数（目標。届かない分は減点）
    for doc in doctors:
        shortfall = model.NewIntVar(0, min_shifts_total[doc], f'min_shortfall_{doc}')
        model.Add(sum(worked_all[doc]) + shortfall >= min(min_shifts_total[doc], total_cap(doc)))
        objective_terms.append(shortfall * -10000)

    # 休日回数の偏りを減らす
    max_hol_shifts = model.NewIntVar(0, num_days * 6, 'max_hol_shifts')
    for doc in doctors:
        model.Add(sum(holiday_worked[doc]) <= max_hol_shifts)

    # 希望日（優先度100未満）は加点
    for doc in doctors:
        if req_priority[doc] < 100:
            weight = req_priority[doc] * 100
            for d in req_days[doc]:
                for s in daily_active_shifts[d]:
                    objective_terms.append(shifts[(d, doc, s)] * weight)
            for d, s_name in req_specific[doc]:
                objective_terms.append(shifts[(d, doc, s_name)] * weight)

    model.Maximize(sum(objective_terms) - max_hol_shifts * 1000)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0
    solver.parameters.random_seed = random.randint(1, 10000)
    status = solver.Solve(model)

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        result_df = build_schedule(lambda key: solver.Value(shifts[key]))
        over_cap_warnings = []
        for d in days:
            date_obj = datetime.date(target_year, target_month, d)
            for s in daily_active_shifts[d]:
                if solver.Value(over_caps[(d, s)]) > 0:
                    over_cap_warnings.append(f"{target_month}/{d}({WEEKDAYS_JA[date_obj.weekday()]}) の「{s}」枠")
        warnings = []
        if status == cp_model.FEASIBLE:
            warnings.append("必要人数を満たす案です。時間内に最適化を完了したことまでは確認できていません。")
        if over_cap_warnings:
            warnings.append("⚠️ **【重要】次の枠は、確定当直や「希望優先度100以上」の指定が重なったため、設定した人数を超えて配置しています（そのまま当直案にしています）:**")
            warnings.extend([f"・{w}" for w in over_cap_warnings])
        return result_df, True, warnings, past_worked_dates, future_worked_dates

    # ==================================================================
    # バックアップ（不足を許容する計算）
    # ==================================================================
    reasons = []
    if status == cp_model.MODEL_INVALID:
        return None, False, ["計算モデルが無効です。入力条件と数値の範囲を確認してください。"], None, None
    if status == cp_model.UNKNOWN:
        reasons.append("通常計算では時間内に案を見つけられませんでした。条件が不可能と確定したわけではありません。")
    try:
        relax_model = cp_model.CpModel()
        r_shifts = {(d, doc, s): relax_model.NewBoolVar(f'r_shift_d{d}_{doc}_{s}') for d in days for doc in doctors for s in daily_active_shifts[d]}
        dummies = {}
        r_excess = []
        for d in days:
            for s in daily_active_shifts[d]:
                # 不足人数はその枠の必要人数を超えない
                dummies[(d, s)] = relax_model.NewIntVar(0, required_count(d, s), f'dummy_d{d}_{s}')
                overflow = relax_model.NewIntVar(0, len(doctors), f'r_over_{d}_{s}')
                r_excess.append(overflow)
                relax_model.Add(sum(r_shifts[(d, doc, s)] for doc in doctors) + dummies[(d, s)] == required_count(d, s) + overflow)

        add_common_constraints(relax_model, r_shifts)

        # 不足人数を、画面で指定した優先順位の群ごとに最小化する。
        # 優先度の低い群から順に重みを決め、上位群の重みは
        # 「下位群すべてで起こりうる最大の損失」より大きくする。
        # こうすると、上位群の不足を1名減らすことが、下位群の不足をいくら減らすことよりも優先される。
        missing_score = 0
        accumulated_bound = 0
        for group in reversed(shortage_priority_groups(shortage_priority)):
            weight = accumulated_bound + 1
            group_keys = [key for key in dummies if key[1] in group]
            missing_score += weight * sum(dummies[key] for key in group_keys)
            accumulated_bound += weight * sum(required_count(d, s) for d, s in group_keys)
        excess_bound = len(r_excess) * len(doctors)
        relax_model.Minimize(missing_score * (excess_bound + 1) + sum(r_excess))

        relax_solver = cp_model.CpSolver()
        relax_solver.parameters.max_time_in_seconds = 15.0
        relax_status = relax_solver.Solve(relax_model)

        if relax_status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            partial_df = build_schedule(lambda key: relax_solver.Value(r_shifts[key]),
                                        lambda key: relax_solver.Value(dummies[key]))
            bottlenecks = []
            missing_by_shift = {s: 0 for s in ALL_SHIFT_TYPES}
            for d in days:
                for s in daily_active_shifts[d]:
                    val = relax_solver.Value(dummies[(d, s)])
                    if val > 0:
                        bottlenecks.append(f"・{target_month}/{d} の「{s}」")
                        missing_by_shift[s] += val
            if bottlenecks:
                reasons.append(shortage_priority_text(shortage_priority))
                if relax_status == cp_model.FEASIBLE:
                    reasons.append("計算時間内に得られた案です。優先順位に沿った不足の最小化が完了したことまでは確認できていません。")
                reasons.append("🚨 **以下の枠で必要人数が不足しています:**")
                reasons.extend(bottlenecks)
                reasons.append("")
                reasons.append("📊 **【不足している枠の合計】**")
                for s, count in sorted(missing_by_shift.items(), key=lambda x: x[1], reverse=True):
                    if count > 0:
                        reasons.append(f"・{s}： 計 {count} 枠不足")
            return partial_df, False, reasons, past_worked_dates, future_worked_dates
        if relax_status == cp_model.INFEASIBLE:
            reasons.append("不足枠を許容しても条件が矛盾し、当直案を作成できませんでした。確定指定と上限を確認してください。")
        elif relax_status == cp_model.MODEL_INVALID:
            reasons.append("不足枠計算のモデルが無効です。入力値を確認してください。")
        else:
            reasons.append("不足枠を含む案も時間内に見つかりませんでした。条件が不可能と確定したわけではありません。")
    except Exception as e:
        reasons.append(f"⚠️ 部分的な当直表の作成中にもエラーが発生しました。詳細: {e}")

    return None, False, reasons, None, None


def generate_shift(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None, locked_days=()):
    staff_df, errors = validate_staff_inputs(staff_df, target_year, target_month)
    fixed_df, fixed_errors = validate_fixed_inputs(fixed_df, staff_df.get('先生の名前', pd.Series(dtype=str)).tolist(), target_year, target_month, custom_holidays)
    errors += fixed_errors
    if staff_df.empty: errors.append('医師を1名以上入力してください。')
    for (d, s), count in multi_slots_dict.items():
        if not isinstance(d, int) or not 1 <= d <= calendar.monthrange(target_year, target_month)[1] or s not in ALL_SHIFT_TYPES or not isinstance(count, int) or not 2 <= count <= 10:
            errors.append(f'増員設定「{d}日 {s} {count}名」を確認してください。')
        elif s in DAY_SHIFTS:
            dt = datetime.date(target_year, target_month, d)
            if not is_holiday_date(dt, target_year, target_month, custom_holidays):
                errors.append(f'{d}日の日直増員：先にカレンダーでその日を「休日にする」にしてください。')
    if errors: return None, False, errors, None, None
    result, success, warnings, past, future = _generate_shift_core(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority, locked_days)
    if result is not None:
        warnings += audit_schedule(result, staff_df, target_year, target_month, custom_holidays, multi_slots_dict, past, future, next_month_special_holiday)
        success = not any('⚠️不足' in str(v) for s in ALL_SHIFT_TYPES for v in result[s])
    return result, success, list(dict.fromkeys(warnings)), past, future


def show_input_errors(messages):
    messages = list(dict.fromkeys(str(m).strip() for m in messages if str(m).strip()))
    st.error(f"入力内容を確認してください（{len(messages)}件）")
    st.dataframe(pd.DataFrame({"修正が必要な項目": messages}), hide_index=True, use_container_width=True)


def show_schedule_notice(result, messages, alerts=None):
    messages = list(dict.fromkeys(str(m).strip() for m in messages if str(m).strip()))
    shortage_rows = []
    if result is not None:
        for _, row in result.iterrows():
            for slot in ALL_SHIFT_TYPES:
                match = re.search(r'⚠️不足\((\d+)名\)', str(row.get(slot, '')))
                if match:
                    shortage_rows.append(int(match.group(1)))
    # 最重要の結果を、説明より先に表示する。
    if result is None:
        st.error("当直案を作成できませんでした。")
    elif shortage_rows:
        st.warning(f"不足が残っています：{len(shortage_rows)}枠、合計{sum(shortage_rows)}名分（医師別の当直一覧のいちばん下、赤線の下で確認できます）")
    else:
        st.success("必要人数を満たす当直案ができました。")

    details = []
    for message in messages:
        if message.startswith(('🚨 **以下の枠', '📊 **【不足している枠')):
            continue
        if re.fullmatch(r'・\d+/\d+ の「.+」', message) or re.fullmatch(r'・.+： 計 \d+ 枠不足', message):
            continue
        details.append(message)
    alerts = alerts or []
    if details or result is None or shortage_rows or alerts:
        with st.container(border=True):
            st.markdown("**計算状況・条件の注意事項と対処方法**")
            if alerts:
                # 当直回数のアラートも、計算の注意事項と同じ枠にまとめて出す
                st.markdown("🔴 **当直が多い先生がいます**（医師別の当直一覧の右端の赤い枠。基準は表のすぐ上の「⚠ アラートの基準」で変えられます）")
                st.markdown("\n".join("- " + a for a in alerts))
            for message in details:
                st.markdown(message)
            if result is None or shortage_rows:
                st.markdown("**条件を見直す場合**")
                st.write("入力内容を確認し、実際に調整できる範囲で、月間・休日・枠別の上限や勤務間隔を見直して再作成してください。")
            if shortage_rows:
                st.write("不足している枠は、医師別の当直一覧のいちばん下（赤線の下の「⚠️ 未決定」）で確認できます。"
                         "手で埋める場合は、緑のボタン「📥 この一覧をExcelでダウンロード」で開き、未決定の枠を×のないマスへドラッグして移してください。"
                         "右の回数と残りの枠の数は自動で変わります。")


# ==========================================
# 先生ごとの専用URLからNG日を入力してもらう仕組み
# 入力内容は管理者のGoogleスプレッドシートに保存する。
#   ・tokens      … 先生ごとの固定キー（専用URLの「?key=」の部分）
#   ・settings    … 受付中の年月・休日設定・締め切り状態
#   ・submissions … 先生が保存するたびに1行ずつ追記（最新の行が有効。履歴として残る）
# ==========================================
DEFAULT_APP_URL = "https://touchoku-app.streamlit.app"
REMOTE_SHEETS = {
    "tokens": ["先生", "キー", "翌日PM duty", "作成日時"],
    "settings": ["項目", "値"],
    "submissions": ["年月", "先生", "NG日", "保存日時", "コメント", "種別"],   # 種別：管理者がリセットした行は「リセット」
    "saved_state": ["項目", "値"],   # 作業の途中保存（年月・休日・医師条件・確定当直など）
}


def remote_now():
    """日本時間の現在時刻（記録用の文字列）。"""
    from zoneinfo import ZoneInfo
    return datetime.datetime.now(ZoneInfo("Asia/Tokyo")).strftime("%Y/%m/%d %H:%M:%S")


def remote_secret(name, default=None):
    try:
        return st.secrets[name]
    except Exception:
        return default


def remote_service_account_info():
    """
    Googleの鍵を読み込む。次のどちらの書き方でもよい。
      ・gcp_service_account_json = ''' （ダウンロードしたJSONファイルの中身をそのまま） '''
      ・[gcp_service_account] の表に1項目ずつ書く
    """
    raw = remote_secret("gcp_service_account_json")
    if raw:
        return json.loads(str(raw))
    table = remote_secret("gcp_service_account")
    return dict(table) if table is not None else None


def remote_is_configured():
    has_key = remote_secret("gcp_service_account_json") is not None or remote_secret("gcp_service_account") is not None
    return has_key and remote_secret("ng_sheet_url") is not None


@st.cache_resource
def remote_spreadsheet():
    """保存用スプレッドシートに接続する（接続は使い回す）。"""
    import gspread
    client = gspread.service_account_from_dict(remote_service_account_info())
    return client.open_by_url(st.secrets["ng_sheet_url"])


def remote_worksheet(name):
    """シートがなければ見出し付きで作る。"""
    import gspread
    sh = remote_spreadsheet()
    try:
        return sh.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=200, cols=len(REMOTE_SHEETS[name]))
        ws.update(values=[REMOTE_SHEETS[name]], range_name="A1", value_input_option="RAW")
        return ws


@st.cache_data(ttl=20, show_spinner=False)
def remote_read(name):
    """シートの中身を辞書のリストで返す。数字に変換されないよう、すべて文字列のまま読む。"""
    values = remote_worksheet(name).get_all_values()
    if not values:
        return []
    header = values[0]
    expected = REMOTE_SHEETS[name]
    # あとから列を増やした場合（例：コメント列）、古い見出しのままでも新しい列名で読む
    if len(header) < len(expected) and header == expected[:len(header)]:
        header = expected
    return [dict(zip(header, row + [""] * (len(header) - len(row)))) for row in values[1:] if any(c.strip() for c in row)]


def remote_clear_cache():
    remote_read.clear()


def remote_append(name, row):
    ws = remote_worksheet(name)
    expected = REMOTE_SHEETS[name]
    # 列が足りなければ増やし、見出しも新しい形にそろえる（以前作ったシートでもそのまま使える）
    if ws.col_count < len(expected):
        ws.add_cols(len(expected) - ws.col_count)
    if ws.row_values(1) != expected:
        ws.update(values=[expected], range_name="A1", value_input_option="RAW")
    ws.append_row([str(v) for v in row], value_input_option="RAW")
    remote_clear_cache()


def remote_rewrite(name, rows):
    """シートを丸ごと書き直す（tokens・settings用。行数が少ないもの）。"""
    ws = remote_worksheet(name)
    ws.clear()
    ws.update(values=[REMOTE_SHEETS[name]] + [[str(v) for v in r] for r in rows], range_name="A1", value_input_option="RAW")
    remote_clear_cache()


def remote_settings():
    return {r.get("項目", ""): r.get("値", "") for r in remote_read("settings")}


def remote_save_settings(**changes):
    current = remote_settings()
    current.update({k: str(v) for k, v in changes.items()})
    remote_rewrite("settings", [[k, v] for k, v in current.items()])


def remote_tokens():
    """{先生: {"キー":…, "翌日PM duty":…}}"""
    return {r["先生"]: r for r in remote_read("tokens") if r.get("先生")}


def remote_new_key():
    import secrets
    return secrets.token_urlsafe(9)


def remote_sync_doctors(doctor_rows):
    """
    名簿の先生全員に固定キーを用意する。すでにある先生のキーは変えない（URLは固定）。
    doctor_rows: [(先生, 翌日PM duty)]
    """
    existing = remote_tokens()
    rows = []
    for name, pm in doctor_rows:
        old = existing.get(name)
        rows.append([name, old["キー"] if old and old.get("キー") else remote_new_key(), pm, old.get("作成日時", remote_now()) if old else remote_now()])
    # 名簿から外れた先生の行も残す（キーを再利用されないよう、記録として保持）
    listed = {name for name, _ in doctor_rows}
    for name, r in existing.items():
        if name not in listed:
            rows.append([name, r.get("キー", ""), r.get("翌日PM duty", ""), r.get("作成日時", "")])
    remote_rewrite("tokens", rows)


def remote_regenerate_key(name):
    rows = []
    for doc, r in remote_tokens().items():
        key = remote_new_key() if doc == name else r.get("キー", "")
        rows.append([doc, key, r.get("翌日PM duty", ""), remote_now() if doc == name else r.get("作成日時", "")])
    remote_rewrite("tokens", rows)


def qr_svg(text, size_mm=38):
    """文字列をQRコード（SVG）にする。追加の部品は不要（reportlabの機能を使う）。"""
    from reportlab.graphics.barcode import qrencoder
    qr = qrencoder.QRCode(None, qrencoder.QRErrorCorrectLevel.M)
    qr.addData(text)
    qr.make()
    n = qr.getModuleCount()
    quiet = 4   # 周りの余白（読み取りに必要）
    cells = "".join(f"M{c + quiet},{r + quiet}h1v1h-1z" for r in range(n) for c in range(n) if qr.isDark(r, c))
    total = n + quiet * 2
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {total} {total}" width="{size_mm}mm" height="{size_mm}mm" '
            f'shape-rendering="crispEdges"><rect width="{total}" height="{total}" fill="#fff"/><path d="{cells}" fill="#000"/></svg>')


def remote_qr_cards_html(cards):
    """
    先生ごとの専用URLを、A4・1人1枚の案内書類にした印刷用HTMLを作る。
    ふつうの書類の体裁のまま、点線で三つ折りにできる（折り目に文字やQRコードがかからない配置）。
    cards: [(先生, 専用URL)]
    """
    import html as _html
    pages = []
    for name, url in cards:
        n = _html.escape(name)
        pages.append(f"""
<section class="sheet">
  <div class="panel p1">
    <div class="to">{n} 先生</div>
    <h1>当直NG日入力用URLのご案内</h1>
    <p class="lead">毎月の当直表を作成するため、当直に入れない日（NG日）を、スマートフォンから入力していただくようお願いいたします。中のQRコード（またはURL）が、<b>{n} 先生専用</b>の入力ページです。</p>
  </div>
  <div class="fold"></div>
  <div class="panel p2">
    <div class="qrbox">
      {qr_svg(url, size_mm=48)}
      <div class="qr-side">
        <div class="big">スマートフォンのカメラで<br>QRコードを読み取ってください</div>
        <div class="small">QRコードにカメラを向けると、NG日の入力ページが開きます。</div>
        <div class="url-label">URL</div>
        <div class="url">{_html.escape(url)}</div>
      </div>
    </div>
  </div>
  <div class="fold"></div>
  <div class="panel p3">
    <h2>入力のしかた</h2>
    <ol>
      <li>QRコードを読み取り、入力ページを開きます。</li>
      <li>入れない日の「日直」「宿直」のボタンを押して、NGにします。<br>
          <span class="sub">休日は「日直」と「宿直」を別々に指定します。終日入れない日は両方押してください。</span></li>
      <li>「NG日とコメントを保存する」を押し、「✅ 保存しました」と出れば完了です。</li>
    </ol>
    <h2>ご注意</h2>
    <ul>
      <li><b>このURLは毎月同じです。</b>締め切りまでは何度でも修正でき、<b>最後に保存した内容</b>が採用されます。</li>
      <li>英語の画面（Zzzz…）が出たら、青いボタン「Yes, get this app back up!」を押して1分ほどお待ちください。</li>
      <li>このURLは<b>先生専用</b>です。他の方には渡さないでください。紛失した場合は担当者までご連絡ください。</li>
    </ul>
  </div>
</section>""")
    return f"""<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>当直NG日入力用URLのご案内</title>
<style>
@page {{ size: A4 portrait; margin: 0; }}
* {{ box-sizing: border-box; }}
body {{ margin: 0; font-family: "Hiragino Mincho ProN", "Yu Mincho", "MS Mincho", serif; color: #111; background: #e9edf2; }}
.bar {{ padding: 12px 16px; background: #eef3f8; border-bottom: 1px solid #cfd8e3; font-size: 14px; line-height: 1.6;
        font-family: "Hiragino Kaku Gothic ProN", "Yu Gothic", "Meiryo", sans-serif; }}
.bar button {{ font-size: 15px; padding: 8px 18px; margin-right: 12px; cursor: pointer; }}
.sheet {{ width: 210mm; height: 297mm; margin: 8mm auto; background: #fff; box-shadow: 0 1px 4px rgba(0,0,0,.15);
         break-after: page; page-break-after: always; overflow: hidden; font-size: 10.5pt; line-height: 1.7; }}
/* A4の縦297mmを3等分（99mm）。折り目の近くには何も置かない */
.panel {{ height: 99mm; padding: 0 22mm; display: flex; flex-direction: column; justify-content: center; }}
.fold {{ height: 0; border-top: 0.8px dashed #8a94a6; margin: 0 6mm; }}
.p1 {{ justify-content: flex-start; padding-top: 16mm; }}
.to {{ font-size: 15pt; border-bottom: 1px solid #111; align-self: flex-start; padding: 0 4mm 0.5mm 0; }}
h1 {{ margin: 9mm 0 6mm; text-align: center; font-size: 17pt; font-weight: 700; letter-spacing: .05em; }}
h2 {{ margin: 0 0 1.5mm; font-size: 12pt; font-weight: 700; border-left: 3px solid #111; padding-left: 2.5mm; line-height: 1.4; }}
h2 + ol {{ margin-bottom: 4mm; }}
.lead {{ margin: 0; text-indent: 1em; font-size: 11pt; line-height: 1.8; }}
.qrbox {{ padding: 6mm 8mm; border: 1px solid #555; display: flex; align-items: center; gap: 9mm; }}
.qr-side .big {{ font-size: 13pt; font-weight: 700; line-height: 1.6; }}
.qr-side .small {{ margin-top: 2mm; font-size: 10pt; }}
.url-label {{ margin-top: 3mm; font-size: 9pt; color: #444; }}
.url {{ font-size: 8.5pt; word-break: break-all; font-family: monospace; color: #333; line-height: 1.4; }}
ol, ul {{ margin: 0; padding-left: 7mm; }}
li {{ margin: .4mm 0; }}
.sub {{ font-size: 9pt; color: #444; }}
@media print {{ body {{ background: #fff; }} .bar {{ display: none; }} .sheet {{ margin: 0; box-shadow: none; }} }}
</style></head><body>
<div class="bar"><button onclick="window.print()">🖨️ 印刷する</button>
A4・縦・片面で印刷してください（{len(cards)}名分・1人1枚）。印刷の設定で「倍率：100%（実際のサイズ）」「余白：なし」を選ぶと、折り線の位置が正しくなります。上の点線は外側へ、下の点線は内側へ折ると、宛名が表に出て、QRコードは内側に隠れます。</div>
{''.join(pages)}
</body></html>"""


def remote_latest_submissions(month_key):
    """その年月の、先生ごとの最新の提出 {先生: {"NG日":…, "保存日時":…}}"""
    latest = {}
    for r in remote_read("submissions"):
        if r.get("年月") == month_key and r.get("先生"):
            latest[r["先生"]] = r   # 後ろの行ほど新しい
    # 最後が管理者のリセットなら、その先生は「未提出」に戻す（履歴は残る）
    return {name: r for name, r in latest.items() if r.get("種別", "") != "リセット"}


def remote_reset_submission(month_key, name):
    """先生の提出をリセットする（未提出に戻す）。行は消さずに「リセット」の行を足すので、履歴に残る。"""
    remote_append("submissions", [month_key, name, "", remote_now(), "", "リセット"])


def remote_doctor_url(key):
    base = str(remote_secret("app_url", DEFAULT_APP_URL)).rstrip("/")
    return f"{base}/?key={key}"


def ng_dict_to_text(values_by_day):
    """{日: NG種類} をアプリのNG日の書き方（例：3,10:宿NG）にする。"""
    items = []
    for d in sorted(values_by_day):
        v = values_by_day[d]
        if v == "全NG":
            items.append(str(d))
        elif v in ("日NG", "宿NG"):
            items.append(f"{d}:{v}")
    return ",".join(items)


def ng_text_summary(text, year, month, holidays):
    """NG日の文字列を「3日、10日(宿直NG)」のような読みやすい形にする。"""
    parts = []
    for d, v in sorted(parse_ng_dict(text, year, month).items()):
        shown = effective_ng(v, is_holiday_date(datetime.date(year, month, d), year, month, holidays))
        if shown == "全NG":
            parts.append(f"{d}日")
        elif shown == "日NG":
            parts.append(f"{d}日(日直NG)")
        elif shown == "宿NG":
            parts.append(f"{d}日(宿直NG)")
    return "、".join(parts)


def ng_calendar_days(year, month, holidays, next_month_special_holiday, pm_weekdays, values_by_day, locked_days=()):
    """NGカレンダー部品に渡す日ごとのデータを作る。locked_days の日は、先生がNGを入力できない（年末年始など）。"""
    days = []
    for d in range(1, calendar.monthrange(year, month)[1] + 1):
        dt = datetime.date(year, month, d)
        hol = is_holiday_date(dt, year, month, holidays)
        days.append({
            "day": d, "weekday": WEEKDAYS_JA[dt.weekday()],
            "options": ["OK", "全NG", "日NG", "宿NG"] if hol else ["OK", "宿NG"],
            "value": effective_ng(values_by_day.get(d, "OK"), hol),
            "warning": pm_duty_restricted(dt, pm_weekdays, year, month, holidays, next_month_special_holiday),
            "kind": "saturday" if dt.weekday() == 5 and not jpholiday.is_holiday(dt) and d not in holidays else ("holiday" if hol else "weekday"),
            "locked": d in locked_days,
        })
        if d in locked_days:
            days[-1]["value"] = "OK"
    return days


def parse_month_key(month_key):
    m = re.fullmatch(r"(\d{4})-(\d{1,2})", str(month_key).strip())
    if not m:
        return None
    y, mo = int(m.group(1)), int(m.group(2))
    return (y, mo) if 1 <= mo <= 12 else None


def parse_holiday_list(text):
    return sorted({int(x) for x in re.findall(r"\d+", str(text))})


def ng_receipt_image(doctor, year, month, saved_at, values_by_day, holidays):
    """
    提出内容（保存済みのNG日）をカレンダー画像にして表示する。
    画像はブラウザの中で描くので、スマートフォンでも日本語が正しく出る。
    スマートフォンは画像を長押しして保存、パソコンはボタンでダウンロードできる。
    """
    import streamlit.components.v1 as components
    days = []
    for d in range(1, calendar.monthrange(year, month)[1] + 1):
        dt = datetime.date(year, month, d)
        hol = is_holiday_date(dt, year, month, holidays)
        days.append({
            "d": d, "hol": hol,
            "sat": dt.weekday() == 5 and not jpholiday.is_holiday(dt) and d not in holidays,
            "ng": effective_ng(values_by_day.get(d, "OK"), hol),
        })
    data = {"doctor": doctor, "year": year, "month": month, "saved": saved_at,
            "offset": datetime.date(year, month, 1).weekday(), "days": days}
    weeks = (data["offset"] + len(days) + 6) // 7
    html_src = """
<div style="font-family:system-ui,sans-serif">
<img id="img" alt="提出内容の画像" style="width:100%;max-width:540px;border:1px solid #d6dce5;border-radius:8px;display:block">
<a id="dl" style="display:inline-block;margin-top:8px;padding:8px 14px;background:#ff4b4b;color:#fff;border-radius:6px;text-decoration:none;font-size:14px">画像をダウンロード</a>
<div style="font-size:12px;color:#566174;margin-top:6px">スマートフォンの場合は、画像を長押しして「写真に保存」（または「画像を保存」）を選んでください。</div>
</div>
<script>
const D = __DATA__;
(function(){
const W = 1080, topY = 230, cellH = 150, rows = __WEEKS__, H = topY + 60 + rows * cellH + 110;
const c = document.createElement('canvas'); c.width = W; c.height = H;
const g = c.getContext('2d');
const F = (s, w) => `${w||400} ${s}px "Hiragino Kaku Gothic ProN","Hiragino Sans","Yu Gothic","Meiryo","Noto Sans JP",sans-serif`;
g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
g.fillStyle = '#243247'; g.font = F(46, 700); g.fillText(`${D.year}年${D.month}月　当直NG日（提出控え）`, 50, 80);
g.font = F(40, 700); g.fillText(`${D.doctor} 先生`, 50, 145);
g.fillStyle = '#566174'; g.font = F(26); g.fillText(`保存日時：${D.saved}`, 50, 195);
const cw = (W - 100) / 7, x0 = 50, y0 = topY;
['月','火','水','木','金','土','日'].forEach((w, i) => {
  g.fillStyle = i === 6 ? '#c92336' : i === 5 ? '#1670c5' : '#243247';
  g.font = F(28, 700); g.textAlign = 'center'; g.fillText(w, x0 + cw * i + cw / 2, y0 + 40);
});
g.textAlign = 'left';
const DAY_NG = '#cf6a35', NIGHT_NG = '#cf6a35';   // 日直NG・宿直NGとも同じ色
let count = 0;
D.days.forEach((day, i) => {
  const p = D.offset + i, col = p % 7, row = Math.floor(p / 7);
  const x = x0 + col * cw, y = y0 + 60 + row * cellH;
  g.fillStyle = day.hol ? '#fff6f5' : '#ffffff'; g.fillRect(x, y, cw, cellH);
  g.strokeStyle = '#cbd2dc'; g.lineWidth = 2; g.strokeRect(x, y, cw, cellH);
  g.fillStyle = day.sat ? '#1670c5' : day.hol ? '#c92336' : '#243247';
  g.font = F(30, 700); g.fillText(String(day.d), x + 12, y + 40);
  const dayNg = day.ng === '全NG' || day.ng === '日NG', nightNg = day.ng === '全NG' || day.ng === '宿NG';
  if (dayNg || nightNg) count++;
  g.font = F(22, 700); g.textAlign = 'center';
  if (dayNg) { g.fillStyle = DAY_NG; g.fillRect(x + 8, y + 58, cw - 16, 30); g.fillStyle = '#fff'; g.fillText('日直NG', x + cw / 2, y + 81); }
  if (nightNg) { g.fillStyle = NIGHT_NG; g.fillRect(x + 8, y + 92, cw - 16, 30); g.fillStyle = '#fff'; g.fillText('宿直NG', x + cw / 2, y + 115); }
  g.textAlign = 'left';
});
g.fillStyle = '#243247'; g.font = F(26);
const fy = y0 + 60 + rows * cellH + 50;
g.fillText(count ? `NG日：${count}日` : 'NG日なし（どの日も入れます）', 50, fy);
g.fillStyle = '#566174'; g.font = F(22);
g.fillText('締め切りまでは保存し直せます。最後に保存した内容が採用されます。', 50, fy + 40);
const url = c.toDataURL('image/png');
document.getElementById('img').src = url;
const a = document.getElementById('dl'); a.href = url; a.download = `NG日_${D.year}年${D.month}月_${D.doctor}.png`;
})();
</script>
""".replace("__DATA__", json.dumps(data, ensure_ascii=False)).replace("__WEEKS__", str(weeks))
    components.html(html_src, height=int(540 * (230 + 60 + weeks * 150 + 110) / 1080) + 100)



# ==========================================
# 作業の途中保存と復元（スプレッドシートの saved_state シートに保存）
# ==========================================
WORK_STATE_CHUNK = 40000   # スプレッドシートの1つのマスに入る文字数（上限5万字）より少なめに区切る


def work_state_save(snapshot):
    text = json.dumps(snapshot, ensure_ascii=False)
    if len(text) > WORK_STATE_CHUNK * 20:
        raise ValueError("保存する内容が大きすぎます（医師の人数や確定当直の行数を確認してください）。")
    # 長いときは「内容」「内容2」「内容3」…と複数のマスに分けて保存する
    chunks = [text[i:i + WORK_STATE_CHUNK] for i in range(0, len(text), WORK_STATE_CHUNK)] or [""]
    rows = [["保存日時", remote_now()]] + [["内容" if i == 0 else f"内容{i + 1}", c] for i, c in enumerate(chunks)]
    remote_rewrite("saved_state", rows)


def work_state_info():
    """保存済みの作業 (保存日時, 内容) を返す。なければ (None, None)。"""
    try:
        rows = {r.get("項目", ""): r.get("値", "") for r in remote_read("saved_state")}
    except Exception:
        return None, None
    if not rows.get("内容"):
        return None, None
    text = rows["内容"]
    n = 2
    while rows.get(f"内容{n}"):   # 複数のマスに分けて保存した分をつなげる
        text += rows[f"内容{n}"]
        n += 1
    try:
        return rows.get("保存日時", ""), json.loads(text)
    except ValueError:
        return None, None


def work_state_restore():
    """保存した作業を画面に戻す（ボタンのコールバック。画面を作る前に実行される）。"""
    saved_at, data = work_state_info()
    if not data:
        st.session_state["work_state_message"] = ("error", "保存された作業が見つかりませんでした。")
        return
    y, m = int(data["year"]), int(data["month"])
    for key in list(st.session_state.keys()):
        if key.startswith(("ng_", "hol_", "lock_", "next_month_special_holiday_", "shortage_priority_")):
            del st.session_state[key]
    st.session_state["sel_year"] = y
    st.session_state["sel_month"] = m
    for d in data.get("holidays", []):
        st.session_state[f"hol_{y}_{m}_{int(d)}"] = True
    for d in data.get("locked", []):
        st.session_state[f"lock_{y}_{m}_{int(d)}"] = True
    st.session_state[f"next_month_special_holiday_{y}_{m}"] = bool(data.get("next_special"))
    for name, cv in (data.get("comments") or {}).items():
        if isinstance(cv, (list, tuple)) and cv and str(cv[0]).strip():
            st.session_state[f"ng_import_comment_{name}_{y}_{m}"] = (str(cv[0]), str(cv[1]) if len(cv) > 1 else "")
    for k, v in (data.get("alerts") or {}).items():
        if k in MATRIX_COUNT_ALERTS:
            st.session_state[f"alert_{k}"] = int(v)
    for slot, rank in (data.get("priority") or {}).items():
        st.session_state[f"shortage_priority_{slot}"] = int(rank)
    st.session_state["restored_state"] = data
    st.session_state["multi_editor_gen"] = st.session_state.get("multi_editor_gen", 0) + 1   # 増員設定の表を読み込んだ内容で作り直す
    # ファイルのアップロード欄を空にして、復元した内容を使う
    st.session_state["uploader_gen"] = st.session_state.get("uploader_gen", 0) + 1
    st.session_state.pop("last_uploaded_file_id", None)
    st.session_state["work_state_message"] = ("success", f"{saved_at} に保存した作業（{y}年{m}月）を読み込みました。")
    # 読み込んだことは、上の帯のボタンの右に1か所だけ出す（次に途中保存するまで表示しておく）
    st.session_state["work_restore_note"] = f"{str(saved_at)[5:16]} 保存分（{y}年{m}月）を読み込みました"
    st.session_state.pop("work_save_note", None)
    st.session_state["work_note_fresh"] = True   # スマートフォンでは、読み込んだ直後に画面上部へ数秒だけ表示する
    # 読み込んだあとは「続きから」の小さな枠を閉じる（見出しを目に見えない形で変えると、閉じた状態で作り直される）
    st.session_state["work_restore_gen"] = st.session_state.get("work_restore_gen", 0) + 1


def carry_over_staff(year, month):
    """
    途中保存してある医師条件（回数・間隔・翌日PM duty・備考など）を、今の年月に引き継ぐ（ボタンのコールバック）。
    月ごとに変わる NG日・希望日・先生のコメントは空にする。
    """
    saved_at, data = work_state_info()
    if not data or not data.get("staff_csv"):
        st.session_state["carry_over_message"] = ("error", "引き継げる医師条件が見つかりませんでした。先に「💾 作業を途中保存する」で保存してください。")
        return
    df = parse_staff_csv(str(data["staff_csv"]).encode("utf-8"))
    for col in (NG_COLUMN, REQUEST_COLUMN):
        if col in df.columns:
            df[col] = ""
    # 画面に残っている今月のNG日・取り込みの記録・先生のコメントを消す（新しい医師条件から読み直す）
    for key in list(st.session_state.keys()):
        if key.startswith("ng_"):
            del st.session_state[key]
    state = dict(st.session_state.get("restored_state") or {"year": year, "month": month})
    state["staff_csv"] = df.to_csv(index=False)
    st.session_state["restored_state"] = state
    st.session_state["uploader_gen"] = st.session_state.get("uploader_gen", 0) + 1
    st.session_state.pop("last_uploaded_file_id", None)
    st.session_state["carry_over_gen"] = st.session_state.get("carry_over_gen", 0) + 1   # 小さな枠を閉じる
    st.session_state["carry_over_message"] = (
        "success", f"{saved_at} に途中保存した{data.get('year')}年{data.get('month')}月の医師条件（{len(df)}名）を引き継ぎました。"
                   "NG日・希望日・先生のコメントは空にしてあります。")


def admin_logged_in(widget_key):
    """管理用パスワードの確認。成功すると、同じ画面のほかの管理機能でも入力不要になる。"""
    if st.session_state.get("remote_admin_ok"):
        return True
    admin_password = remote_secret("admin_password")
    if not admin_password:
        st.caption("管理用パスワード（admin_password）が設定されていません。")
        return False
    pw = st.text_input("管理用パスワード", type="password", key=widget_key)
    if pw:
        if pw == str(admin_password):
            st.session_state["remote_admin_ok"] = True
            st.rerun()
        else:
            st.error("パスワードが違います。")
    return False


# 途中保存ボタンは、画面の右上に固定して表示する（スクロールしても付いてくる）。
# 押されたことを覚えておき、全部の入力を読み終えた位置（作成ボタンの上）で保存し、結果をボタンの下に表示する。
WORK_SAVE_SLOT = None
work_save_request = False

WORK_SAVE_FLOAT_CSS = """
<style>
/* 画面いちばん上の帯（右上に Fork・⋮ がある白い帯）の空いている所に置く。表や入力欄には重ならない */
.st-key-work_save_float {
    position: fixed; top: 0.45rem; left: 1rem; right: auto; z-index: 1000002; margin: 0 !important;
    width: auto !important; max-width: 60vw;
    display: flex !important; flex-direction: row !important; align-items: center; gap: 10px;
    background: transparent;
}
.st-key-work_save_float > div { width: auto !important; flex: 0 0 auto !important; }
.st-key-work_save_float button { min-height: 2.2rem; padding: 0.2rem 0.9rem; font-weight: 700; white-space: nowrap; }
.st-key-work_save_float p { margin: 0; white-space: nowrap; }
.work-save-note { font-size: 0.85rem; color: #8ff0b6; font-weight: 700; }
/* 保存済みの表示はボタンの左に置く（右側には、Streamlit の「実行中」の表示が出るため） */
.st-key-work_save_float > div:has(.work-save-note) { order: 3; }
.work-save-note.busy { color: #e3e9f1; }
.work-save-note.bad { color: #ffb3b3; }
@keyframes work-note-fade0 { to { opacity: 0; visibility: hidden; } }
@keyframes work-note-fade1 { to { opacity: 0; visibility: hidden; } }
/* スマートフォン：右上のStreamlitのアイコンと重ならないよう左に寄せ、ボタンの文字を短くする */
@media (max-width: 640px) {
    .st-key-work_save_float { left: 0.5rem; right: auto; top: 0.55rem; max-width: calc(100vw - 6rem); gap: 6px; }
    .st-key-work_save_float button { padding: 0.1rem 0.55rem; min-height: 2rem; }
    .st-key-work_save_float > div:nth-child(1) button p,
    .st-key-work_save_float > div:nth-child(2) button p { font-size: 0 !important; }
    .st-key-work_save_float > div:nth-child(1) button p::after { content: "💾 途中保存"; font-size: 0.8rem; }
    .st-key-work_save_float > div:nth-child(2) button p::after { content: "📂 続きから"; font-size: 0.8rem; }
    /* 帯の中には入りきらないので、ふだんは隠す。保存・読み込みの直後だけ、帯のすぐ下の中央に数秒表示する */
    .work-save-note { display: none; }
    .work-save-note.fresh {
        display: block; position: fixed; top: 3.9rem; left: 50%; transform: translateX(-50%);
        z-index: 1000003; width: max-content; max-width: calc(100vw - 2rem); white-space: normal; text-align: center;
        background: #1f3550; color: #ffffff; border: 2px solid #8ff0b6; border-radius: 999px;
        padding: 0.5rem 1rem; font-size: 0.9rem; box-shadow: 0 4px 14px rgba(0,0,0,.25);
        animation: work-note-fade0 0.6s ease 6s forwards;
    }
    /* 続けて保存したときも表示し直されるよう、交互に別の名前の動きを使う */
    .work-save-note.fresh.n1 { animation-name: work-note-fade1; }
    .work-save-note.fresh.busy { border-color: #e3e9f1; animation: none; }
    .work-save-note.fresh.bad { background: #7a1f1f; border-color: #ffb3b3; animation: none; }
}
</style>
"""


def work_save_floating():
    """画面右上の帯に「前回の作業の続きから始める」と「作業を途中保存する」を並べて置く。"""
    global work_save_request, WORK_SAVE_SLOT
    if not remote_is_configured():
        return
    st.markdown(WORK_SAVE_FLOAT_CSS, unsafe_allow_html=True)
    msg = st.session_state.pop("work_state_message", None)
    if msg and msg[0] != "success":
        st.error(msg[1])
    has_popover = hasattr(st, "popover")
    with st.container(key="work_save_float"):
        # 並び：いちばん左に「作業を途中保存する」、その右に「前回の作業の続きから始める」
        # --- 作業を途中保存する ---
        logged_in = bool(st.session_state.get("remote_admin_ok"))
        if logged_in:
            if st.button("💾 作業を途中保存する", key="work_save_button", type="primary"):
                work_save_request = True
        else:
            # パスワード未入力のときは、押すと入力欄が開くボタンにする
            box = st.popover("💾 作業を途中保存する") if has_popover \
                else st.expander("💾 作業を途中保存する", expanded=False)
            with box:
                st.caption("途中保存には管理用パスワードが必要です。")
                admin_logged_in("work_save_pw")

        # --- 前回の作業の続きから始める（押すと小さな枠が開く） ---
        _restore_label = "📂 前回の作業の続きから始める" + "\u200b" * st.session_state.get("work_restore_gen", 0)
        box = st.popover(_restore_label) if has_popover \
            else st.expander(_restore_label, expanded=False)
        with box:
            if admin_logged_in("work_restore_pw"):
                saved_at, saved = work_state_info()
                if saved:
                    st.write(f"保存日時：**{saved_at}**")
                    st.write(f"{saved['year']}年{saved['month']}月の作業")
                    st.button("この作業を読み込む", key="work_restore", type="primary", on_click=work_state_restore)
                    st.caption("年月・休日と別途決定の日・増員設定・確定当直・医師条件（NG日を含む）・取り込んだ先生のコメント・不足時の優先順位・アラートの基準が、"
                               "途中保存したときの状態に戻ります。今の画面の入力内容は置き換わります。")
                else:
                    st.caption("まだ途中保存された作業はありません。左の「💾 作業を途中保存する」で保存できます。")
            else:
                st.caption("読み込みには管理用パスワードが必要です。")

        # --- 保存・読み込みの結果（ボタンの右に1行） ---
        if logged_in:
            # 「続きから読み込んだ」「途中保存した」のどちらも、同じ場所に1行だけ出す（次の保存・読み込みまで表示しておく）
            WORK_SAVE_SLOT = st.empty()
            _fresh = ""
            if st.session_state.pop("work_note_fresh", False):
                st.session_state["work_note_count"] = st.session_state.get("work_note_count", 0) + 1
                _fresh = f" fresh n{st.session_state['work_note_count'] % 2}"
            if st.session_state.get("work_restore_note"):
                WORK_SAVE_SLOT.markdown(f"<span class='work-save-note{_fresh}'>📂 {st.session_state['work_restore_note']}</span>", unsafe_allow_html=True)
            elif st.session_state.get("work_save_note"):
                WORK_SAVE_SLOT.markdown(f"<span class='work-save-note{_fresh}'>✅ {st.session_state['work_save_note']} 途中保存しました</span>", unsafe_allow_html=True)
            _dirty = st.session_state.pop("work_save_dirty", None)
            if _dirty:
                st.toast(f"{'・'.join(_dirty)}先生のカレンダーに、反映していない変更があります。"
                         "この変更は途中保存に含まれていません。「反映する」を押してから、もう一度途中保存してください。", icon="⚠️")

def doctor_entry_page(key):
    """専用URLで開いた先生用の画面。その先生のNG日だけを入力・保存できる。"""
    st.markdown("<style>[data-testid='stSidebar']{display:none}</style>", unsafe_allow_html=True)
    if not remote_is_configured():
        st.error("現在、NG日の入力を受け付けていません（準備中です）。担当者にお問い合わせください。")
        return
    try:
        tokens = remote_tokens()
        settings = remote_settings()
    except Exception:
        st.error("データを読み込めませんでした。しばらくしてから、もう一度開いてください。")
        return
    doctor = next((name for name, r in tokens.items() if r.get("キー") and r.get("キー") == key), None)
    if doctor is None:
        st.error("このURLは無効です。担当者から届いた最新のURLを開いてください。")
        return
    # 期間指定版（年末年始など）で受付中のときは、期間のカレンダーを出す
    if str(settings.get("受付中の年月", "")).startswith(PERIOD_KEY_PREFIX):
        doctor_entry_page_period(key)
        return

    ym = parse_month_key(settings.get("受付中の年月", ""))
    if ym is None:
        st.info(f"{doctor} 先生、現在受け付けているNG日の入力はありません。")
        return
    year, month = ym
    holidays = parse_holiday_list(settings.get("特別休日", ""))
    next_special = settings.get("翌月1日特別休日", "") == "True"
    closed = settings.get("締め切り", "") == "True"
    deadline = settings.get("締め切り日", "").strip()

    st.title(f"{year}年{month}月　当直NG日の入力")
    st.subheader(f"🩺 {doctor} 先生")
    latest = remote_latest_submissions(f"{year}-{month}").get(doctor)
    saved_text = latest.get("NG日", "") if latest else ""

    if closed:
        st.warning("受付は締め切りました。変更がある場合は、担当者に直接ご連絡ください。")
        if latest:
            summary = ng_text_summary(saved_text, year, month, holidays) or "NG日なし"
            st.write(f"提出内容（{latest.get('保存日時', '')} 保存）：{summary}")
            if latest.get("コメント", "").strip():
                st.write(f"管理者へのコメント：{latest.get('コメント', '')}")
        else:
            st.write("この月のNG日は提出されていません。")
        return

    st.info(
        (f"締め切り：{deadline}\n\n" if deadline else "")
        + "締め切りまでは、何度でも保存し直せます。**最後に保存した内容が採用されます**"
        "（前に保存した内容は、新しい内容に置き換わります）。"
    )
    just_saved = st.session_state.pop("remote_just_saved", False)
    if just_saved:
        try:
            st.toast("保存しました", icon="✅")   # 画面の隅に短く表示（古いStreamlitでは出ないだけ）
        except Exception:
            pass

    def saved_status_box(show_comment=True):
        """保存済みの内容を、はっきり分かる形で表示する（カレンダーの上と下の両方に出す）。"""
        if latest:
            summary = ng_text_summary(saved_text, year, month, holidays) or "NG日なし（どの日も入れます）"
            title = "✅ 保存しました" if just_saved else "✅ 提出済みです"
            st.success(f"**{title}**（{latest.get('保存日時', '')}）\n\n"
                       f"**現在の提出内容：** {summary}\n\n"
                       "変更したい場合は、カレンダーのボタンを押し直して、もう一度保存してください。")
            # コメントは、アプリの案内文と混ざらないように別の枠で表示する
            comment = latest.get("コメント", "").strip()
            if comment and show_comment:
                with st.container(border=True):
                    st.markdown("💬 **保存されている管理者へのコメント**")
                    st.text(f"「{comment}」")
        else:
            st.warning("**まだ提出されていません。** 当直に入れない日の「日直」「宿直」のボタンを押してNGにし、いちばん下の「NG日とコメントを保存する」を押してください。"
                       "NG日がない場合も、そのまま保存ボタンを押してください（「NG日なし」として提出されます）。")

    saved_status_box()

    # 管理者へのコメント（任意）は、カレンダーの下・保存ボタンのすぐ上に出す（カレンダー部品の中にある）
    comment_key = f"remote_comment_{year}_{month}"
    saved_comment = latest.get("コメント", "") if latest else ""
    st.caption("下のカレンダーで「日直」「宿直」のボタンを押してNGにし、いちばん下の「NG日とコメントを保存する」を押してください。"
               "コメントだけを直したいときも、同じボタンを押してください（NG日はそのまま）。")

    # 以前の紙の用紙（休日は日直・宿直まとめて×）から変わった点を、カレンダーのすぐ上で強調する
    st.markdown("""
<div style="border:2px solid #8a94a6;background:#f4f6f9;border-radius:10px;padding:12px 14px;margin:6px 0 10px">
 <div style="font-weight:800;font-size:1.05rem;color:#22345a">⚠ 休日は「日直」と「宿直」を別々に指定します</div>
 <div style="margin-top:4px;font-size:.95rem;color:#243247">終日入れない日は、<b>両方</b>押してください。</div>
 <div style="display:flex;flex-wrap:wrap;gap:10px;margin-top:10px;font-size:.82rem;font-weight:700">
  <div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">日直<span>NG</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">宿直<span>NG</span></div></div></div><div style="margin-top:4px;color:#243247">終日NG</div></div><div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#fff;color:#243041;">日直<span style="color:#2f8a5b">可</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">宿直<span>NG</span></div></div></div><div style="margin-top:4px;color:#243247">宿直だけNG<br><span style="color:#566174">日直は入れる</span></div></div><div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">日直<span>NG</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#fff;color:#243041;border-top:1px solid #e3e7ed;">宿直<span style="color:#2f8a5b">可</span></div></div></div><div style="margin-top:4px;color:#243247">日直だけNG<br><span style="color:#566174">宿直は入れる</span></div></div>
 </div>
</div>""", unsafe_allow_html=True)

    layout = st.radio("表示", ["月間カレンダー", "1日〜月末を横一列"], horizontal=True, key="remote_layout")
    pm_weekdays = pm_duty_weekdays(tokens[doctor].get("翌日PM duty", ""))
    values = parse_ng_dict(saved_text, year, month)
    locked_days = set(parse_holiday_list(settings.get("入力不可日", "")))
    days = ng_calendar_days(year, month, holidays, next_special, pm_weekdays, values, locked_days)
    if locked_days:
        st.caption("灰色の日（" + "・".join(f"{d}日" for d in sorted(locked_days)) + "）は、別の方法で当番を決めるため、NGは入力できません。")
    revision = hashlib.sha256(json.dumps([year, month, doctor, layout, days], ensure_ascii=False).encode()).hexdigest()
    component_key = f"remote_ng_{'row' if layout != '月間カレンダー' else 'month'}_{year}_{month}"
    response = horizontal_ng_component(HORIZONTAL_NG_HTML)(
        days=days, mode="row" if layout != "月間カレンダー" else "month",
        offset=datetime.date(year, month, 1).weekday(), doctor=doctor,
        button_label="NG日とコメントを保存する", pending_word="保存",
        comment_enabled=True, comment=saved_comment, bulk_buttons=True,
        version=revision, key=component_key, default=None,
    )
    seen_key = component_key + "_last_token"
    if isinstance(response, dict) and response.get("token") != st.session_state.get(seen_key):
        st.session_state[seen_key] = response.get("token")
        new_values = response.get("values")
        if response.get("version") == revision and isinstance(new_values, list) and len(new_values) == len(days) \
                and all(v in item["options"] for v, item in zip(new_values, days)):
            # 締め切り直前に保存された場合に備えて、保存の直前にもう一度確認する
            remote_clear_cache()
            if remote_settings().get("締め切り", "") == "True":
                st.error("受付が締め切られたため、保存できませんでした。担当者に直接ご連絡ください。")
                return
            # 画面で選べない日（平日の日NGなど）の過去の値は、そのまま残す
            merged = {d: v for d, v in values.items() if d not in locked_days}   # 入力できない日のNGは残さない
            for d, (v, item) in enumerate(zip(new_values, days), 1):
                if d in locked_days:
                    continue
                if v != item["value"]:
                    if v == "OK":
                        merged.pop(d, None)
                    else:
                        merged[d] = v
            try:
                now = remote_now()
                comment_text = str(response.get("comment", saved_comment) or "").strip()[:500]
                st.session_state[comment_key] = comment_text
                remote_append("submissions", [f"{year}-{month}", doctor, ng_dict_to_text(merged), now, comment_text])
            except Exception:
                st.error("保存できませんでした。通信状態を確認して、もう一度保存ボタンを押してください。")
                return
            st.session_state["remote_just_saved"] = True
            st.rerun()
    # 保存した直後だけ、保存ボタンのすぐ下にも表示する（スマートフォンで画面を上に戻さなくても確認できるように）
    if just_saved:
        saved_status_box(show_comment=False)
    if latest:
        with st.expander("📷 提出内容を画像で残す（控えとして）", expanded=just_saved):
            st.caption("保存済みの内容をカレンダーの画像にします。保存し直した場合は、画像も新しい内容になります。")
            ng_receipt_image(doctor, year, month, latest.get("保存日時", ""), parse_ng_dict(saved_text, year, month), holidays)
    st.caption("このページのURLは、先生専用です。ほかの方には共有しないでください。")


def doctor_entry_page_period(key):
    """期間指定版（年末年始など）で受付中のときの、先生用の入力画面。"""
    st.markdown("<style>[data-testid='stSidebar']{display:none}</style>", unsafe_allow_html=True)
    if not remote_is_configured():
        st.error("現在、NG日の入力を受け付けていません（準備中です）。担当者にお問い合わせください。")
        return
    try:
        tokens = remote_tokens()
        settings = remote_settings()
    except Exception:
        st.error("データを読み込めませんでした。しばらくしてから、もう一度開いてください。")
        return
    doctor = next((name for name, r in tokens.items() if r.get("キー") and r.get("キー") == key), None)
    if doctor is None:
        st.error("このURLは無効です。担当者から届いた最新のURLを開いてください。")
        return

    P = TouchokuPeriod.from_key(settings.get("受付中の年月", ""))
    if P is None:
        st.info(f"{doctor} 先生、現在受け付けているNG日の入力はありません。")
        return
    year, month = P.year, P.month
    holidays = parse_holiday_list(settings.get("特別休日", ""))
    next_special = settings.get("翌月1日特別休日", "") == "True"
    closed = settings.get("締め切り", "") == "True"
    deadline = settings.get("締め切り日", "").strip()

    st.title(f"{P.title()}　当直NG日の入力")
    st.subheader(f"🩺 {doctor} 先生")
    latest = remote_latest_submissions(P.key).get(doctor)
    saved_text = latest.get("NG日", "") if latest else ""

    if closed:
        st.warning("受付は締め切りました。変更がある場合は、担当者に直接ご連絡ください。")
        if latest:
            summary = period_ng_summary(saved_text, P, holidays) or "NG日なし"
            st.write(f"提出内容（{latest.get('保存日時', '')} 保存）：{summary}")
            if latest.get("コメント", "").strip():
                st.write(f"管理者へのコメント：{latest.get('コメント', '')}")
        else:
            st.write("この期間のNG日は提出されていません。")
        return

    st.info(
        (f"締め切り：{deadline}\n\n" if deadline else "")
        + "締め切りまでは、何度でも保存し直せます。**最後に保存した内容が採用されます**"
        "（前に保存した内容は、新しい内容に置き換わります）。"
    )
    just_saved = st.session_state.pop("remote_just_saved", False)
    if just_saved:
        try:
            st.toast("保存しました", icon="✅")   # 画面の隅に短く表示（古いStreamlitでは出ないだけ）
        except Exception:
            pass

    def saved_status_box(show_comment=True):
        """保存済みの内容を、はっきり分かる形で表示する（カレンダーの上と下の両方に出す）。"""
        if latest:
            summary = period_ng_summary(saved_text, P, holidays) or "NG日なし（どの日も入れます）"
            title = "✅ 保存しました" if just_saved else "✅ 提出済みです"
            st.success(f"**{title}**（{latest.get('保存日時', '')}）\n\n"
                       f"**現在の提出内容：** {summary}\n\n"
                       "変更したい場合は、カレンダーのボタンを押し直して、もう一度保存してください。")
            # コメントは、アプリの案内文と混ざらないように別の枠で表示する
            comment = latest.get("コメント", "").strip()
            if comment and show_comment:
                with st.container(border=True):
                    st.markdown("💬 **保存されている管理者へのコメント**")
                    st.text(f"「{comment}」")
        else:
            st.warning("**まだ提出されていません。** 当直に入れない日の「日直」「宿直」のボタンを押してNGにし、いちばん下の「NG日とコメントを保存する」を押してください。"
                       "NG日がない場合も、そのまま保存ボタンを押してください（「NG日なし」として提出されます）。")

    saved_status_box()

    # 管理者へのコメント（任意）は、カレンダーの下・保存ボタンのすぐ上に出す（カレンダー部品の中にある）
    comment_key = f"remote_comment_{P.key}"
    saved_comment = latest.get("コメント", "") if latest else ""
    st.caption("下のカレンダーで「日直」「宿直」のボタンを押してNGにし、いちばん下の「NG日とコメントを保存する」を押してください。"
               "コメントだけを直したいときも、同じボタンを押してください（NG日はそのまま）。")

    # 以前の紙の用紙（休日は日直・宿直まとめて×）から変わった点を、カレンダーのすぐ上で強調する
    st.markdown("""
<div style="border:2px solid #8a94a6;background:#f4f6f9;border-radius:10px;padding:12px 14px;margin:6px 0 10px">
 <div style="font-weight:800;font-size:1.05rem;color:#22345a">⚠ 休日は「日直」と「宿直」を別々に指定します</div>
 <div style="margin-top:4px;font-size:.95rem;color:#243247">終日入れない日は、<b>両方</b>押してください。</div>
 <div style="display:flex;flex-wrap:wrap;gap:10px;margin-top:10px;font-size:.82rem;font-weight:700">
  <div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">日直<span>NG</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">宿直<span>NG</span></div></div></div><div style="margin-top:4px;color:#243247">終日NG</div></div><div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#fff;color:#243041;">日直<span style="color:#2f8a5b">可</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">宿直<span>NG</span></div></div></div><div style="margin-top:4px;color:#243247">宿直だけNG<br><span style="color:#566174">日直は入れる</span></div></div><div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">日直<span>NG</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#fff;color:#243041;border-top:1px solid #e3e7ed;">宿直<span style="color:#2f8a5b">可</span></div></div></div><div style="margin-top:4px;color:#243247">日直だけNG<br><span style="color:#566174">宿直は入れる</span></div></div>
 </div>
</div>""", unsafe_allow_html=True)

    layout = st.radio("表示", ["月間カレンダー", "横一列"], horizontal=True, key="remote_layout_period")
    pm_weekdays = pm_duty_weekdays(tokens[doctor].get("翌日PM duty", ""))
    values = period_parse_ng(saved_text, P)
    locked_days = set(parse_holiday_list(settings.get("入力不可日", "")))
    days = period_ng_calendar_days(P, holidays, next_special, pm_weekdays, values, locked_days)
    if locked_days:
        st.caption("灰色の日（" + "・".join(P.label(d) for d in sorted(locked_days) if 1 <= d <= P.n) + "）は、別の方法で当番を決めるため、NGは入力できません。")
    revision = hashlib.sha256(json.dumps([P.key, doctor, layout, days], ensure_ascii=False).encode()).hexdigest()
    component_key = f"remote_ng_{'row' if layout != '月間カレンダー' else 'month'}_{P.key}"
    response = horizontal_ng_component(HORIZONTAL_NG_HTML)(
        days=days, mode="row" if layout != "月間カレンダー" else "month",
        offset=P.start.weekday(), doctor=doctor,
        button_label="NG日とコメントを保存する", pending_word="保存",
        comment_enabled=True, comment=saved_comment, bulk_buttons=True,
        version=revision, key=component_key, default=None,
    )
    seen_key = component_key + "_last_token"
    if isinstance(response, dict) and response.get("token") != st.session_state.get(seen_key):
        st.session_state[seen_key] = response.get("token")
        new_values = response.get("values")
        if response.get("version") == revision and isinstance(new_values, list) and len(new_values) == len(days) \
                and all(v in item["options"] for v, item in zip(new_values, days)):
            # 締め切り直前に保存された場合に備えて、保存の直前にもう一度確認する
            remote_clear_cache()
            if remote_settings().get("締め切り", "") == "True":
                st.error("受付が締め切られたため、保存できませんでした。担当者に直接ご連絡ください。")
                return
            # 画面で選べない日（平日の日NGなど）の過去の値は、そのまま残す
            merged = {d: v for d, v in values.items() if d not in locked_days}   # 入力できない日のNGは残さない
            for d, (v, item) in enumerate(zip(new_values, days), 1):
                if d in locked_days:
                    continue
                if v != item["value"]:
                    if v == "OK":
                        merged.pop(d, None)
                    else:
                        merged[d] = v
            try:
                now = remote_now()
                comment_text = str(response.get("comment", saved_comment) or "").strip()[:500]
                st.session_state[comment_key] = comment_text
                remote_append("submissions", [P.key, doctor, period_ng_text(merged, P), now, comment_text])
            except Exception:
                st.error("保存できませんでした。通信状態を確認して、もう一度保存ボタンを押してください。")
                return
            st.session_state["remote_just_saved"] = True
            st.rerun()
    # 保存した直後だけ、保存ボタンのすぐ下にも表示する（スマートフォンで画面を上に戻さなくても確認できるように）
    if just_saved:
        saved_status_box(show_comment=False)
    st.caption("このページのURLは、先生専用です。ほかの方には共有しないでください。")


def imported_comments(year, month):
    """
    アプリに取り込んだ先生のコメント {先生: (コメント, 保存日時)}。
    提出されただけでまだ取り込んでいないコメントは含めない（アプリをリセットしたら消え、途中保存から戻せる）。
    """
    prefix, suffix = "ng_import_comment_", f"_{year}_{month}"
    return {k[len(prefix):-len(suffix)]: tuple(v) for k, v in st.session_state.items()
            if k.startswith(prefix) and k.endswith(suffix) and v and str(v[0]).strip()}


def remote_comments_safe(year, month):
    """
    その年月に先生方が専用URLから書いたコメント {先生: (コメント, 保存日時)}。
    接続設定がない・読み込めない場合は空にする（NG入力の画面を止めないため）。
    """
    if not remote_is_configured():
        return {}
    try:
        latest = remote_latest_submissions(f"{year}-{month}")
    except Exception:
        return {}
    return {name: (r.get("コメント", "").strip(), r.get("保存日時", ""))
            for name, r in latest.items() if r.get("コメント", "").strip()}

NG_BANDS = {"全NG": {"day", "night"}, "日NG": {"day"}, "宿NG": {"night"}}


def merge_ng(a, b):
    """2つのNG指定 {日: 種類} を合わせる。例：日NG＋宿NG → 全NG。"""
    merged = {}
    for d in set(a) | set(b):
        bands = NG_BANDS.get(a.get(d), set()) | NG_BANDS.get(b.get(d), set())
        if bands == {"day", "night"}:
            merged[d] = "全NG"
        elif bands == {"day"}:
            merged[d] = "日NG"
        elif bands == {"night"}:
            merged[d] = "宿NG"
    return merged


def planned_import(name, year, month, submitted, mode):
    """
    取り込み後のNG日を計算する。
    合わせる：管理者が入れたNG（CSV・カレンダー）と先生の提出の両方を残す。
              前回取り込んだ分は差し替えるので、先生が提出し直してNGを減らした場合も正しく反映される。
    置き換える：先生の提出内容だけにする。
    戻り値：(取り込み後, 管理者側の元のNG)
    """
    num_days = calendar.monthrange(year, month)[1]
    current = {d: st.session_state.get(f"ng_{name}_{year}_{month}_{d}", "OK") for d in range(1, num_days + 1)}
    current = {d: v for d, v in current.items() if v != "OK"}
    last = st.session_state.get(f"ng_import_last_{name}_{year}_{month}")
    base = dict(current)
    if last:
        # 前回の取り込み結果から変わっていない日は、取り込み前の管理者の値に戻してから合わせ直す
        last_result, last_base = last
        base = {}
        for d in set(current) | set(last_result):
            if current.get(d) == last_result.get(d):
                if last_base.get(d):
                    base[d] = last_base[d]
            elif current.get(d):
                base[d] = current[d]
    result = submitted if mode == "replace" else merge_ng(base, submitted)
    return result, base


def import_remote_ng(year, month, submissions, mode="merge"):
    """先生方の提出内容を、アプリのNGカレンダー（画面の保存値）に取り込む。"""
    num_days = calendar.monthrange(year, month)[1]
    for doctor, r in submissions.items():
        submitted = parse_ng_dict(r.get("NG日", ""), year, month)
        result, base = planned_import(doctor, year, month, submitted, mode)
        # 取り消せるように、最初に取り込む前のNG日を覚えておく（取り込み直しても、いちばん最初の状態に戻せる）
        undo_key = f"ng_import_undo_{doctor}_{year}_{month}"
        if undo_key not in st.session_state:
            before = {d: st.session_state.get(f"ng_{doctor}_{year}_{month}_{d}", "OK") for d in range(1, num_days + 1)}
            st.session_state[undo_key] = {d: v for d, v in before.items() if v != "OK"}
        for d in range(1, num_days + 1):
            st.session_state[f"ng_{doctor}_{year}_{month}_{d}"] = result.get(d, "OK")
        st.session_state[f"ng_import_last_{doctor}_{year}_{month}"] = (result, base)
        # 先生のコメントも、取り込んだときだけアプリ（医師別の当直一覧など）に表示する
        comment = str(r.get("コメント", "")).strip()
        if comment:
            st.session_state[f"ng_import_comment_{doctor}_{year}_{month}"] = (comment, r.get("保存日時", ""))
        else:
            st.session_state.pop(f"ng_import_comment_{doctor}_{year}_{month}", None)
    how = "管理者の入力と合わせて" if mode == "merge" else "先生の提出内容に置き換えて"
    st.session_state["remote_import_message"] = f"{len(submissions)}名分のNG日を、{how}取り込みました。"


def import_undo_doctors(year, month, names):
    """URLからの取り込みを取り消せる先生（取り込む前の状態を覚えている先生）。"""
    return [n for n in names if f"ng_import_undo_{n}_{year}_{month}" in st.session_state]


def undo_remote_import(year, month, names):
    """URLから取り込んだNG日を取り消し、取り込む前のNG日に戻す。"""
    num_days = calendar.monthrange(year, month)[1]
    done = []
    for name in names:
        before = st.session_state.pop(f"ng_import_undo_{name}_{year}_{month}", None)
        if before is None:
            continue
        for d in range(1, num_days + 1):
            st.session_state[f"ng_{name}_{year}_{month}_{d}"] = before.get(d, "OK")
        st.session_state.pop(f"ng_import_last_{name}_{year}_{month}", None)
        st.session_state.pop(f"ng_import_comment_{name}_{year}_{month}", None)
        st.session_state.pop(f"ng_dirty::{name}", None)
        done.append(name)
    if done:
        st.toast(f"↩️ {'、'.join(done)} 先生のNG日を、URLから取り込む前の状態に戻しました。" if len(done) <= 3
                 else f"↩️ {len(done)}名分のNG日を、URLから取り込む前の状態に戻しました。")


def remote_admin_panel(year, month, doctor_rows, custom_holidays, next_month_special_holiday, day_is_holiday, locked_days=()):
    """管理者用：受付の開始・締め切り、専用URL、提出状況、取り込み。"""
    # 管理用パスワードが必要な欄であることが一目で分かるように、紺色の枠で囲み、上に印を付ける
    st.markdown("""
<style>
.st-key-remote_admin_box{border:2px solid #3b5b92;background:#f3f6fb;border-radius:12px;padding:10px 12px 6px}
.st-key-remote_admin_box [data-testid="stExpander"] details{background:#fff}
.admin-badge{display:inline-flex;align-items:center;gap:6px;font-size:.85rem;font-weight:800;color:#fff;background:#3b5b92;border-radius:999px;padding:2px 12px}
.admin-badge.on{background:#2f8a5b}
.admin-note{font-size:.85rem;color:#3b4a63;margin-left:8px}
</style>""", unsafe_allow_html=True)
    # 上のカレンダー（休日にする・別途決定）を、この画面で今まさに変えたかどうか。
    # 開いた直後（まだ何も設定していない状態）で先生側の設定を消してしまわないよう、
    # 変えたときだけ自動で先生側に反映する。
    _cal_now = (year, month, tuple(sorted(custom_holidays)), tuple(sorted(locked_days)), bool(next_month_special_holiday))
    _cal_prev = st.session_state.get("remote_cal_prev")
    # 年月を切り替えただけのときは「変えた」とみなさない
    cal_changed = _cal_prev is not None and _cal_prev[:2] == _cal_now[:2] and _cal_prev != _cal_now
    st.session_state["remote_cal_prev"] = _cal_now
    logged_in = bool(st.session_state.get("remote_admin_ok"))
    box = st.container(key="remote_admin_box")
    with box:
        if logged_in:
            st.markdown("<span class='admin-badge on'>🔓 管理者専用（ログイン中）</span>", unsafe_allow_html=True)
        else:
            st.markdown("<span class='admin-badge'>🔒 管理者専用</span><span class='admin-note'>管理用パスワードを入力すると使えます</span>",
                        unsafe_allow_html=True)
    with box, st.expander("📨 先生方からのNG日入力（専用URL）", expanded=False):
        if not remote_is_configured():
            st.info("この機能を使うには、最初に保存用のGoogleスプレッドシートとの接続設定が必要です。設定方法は担当者向けの手順書を参照してください。")
            return
        admin_password = remote_secret("admin_password")
        if not admin_password:
            st.error("管理用パスワード（admin_password）が設定されていません。Secretsに追加してください。")
            return
        if not st.session_state.get("remote_admin_ok"):
            # ログインしていないと、上のカレンダーの「休日にする」「別途決定」は先生側に届かないので知らせる
            try:
                _st = remote_settings()
                if _st.get("受付中の年月", "") == f"{year}-{month}" and (
                        parse_holiday_list(_st.get("入力不可日", "")) != sorted(locked_days)
                        or parse_holiday_list(_st.get("特別休日", "")) != sorted(custom_holidays)):
                    st.warning("上のカレンダーの「休日にする」「別途決定」が、先生側の画面とまだ違います。"
                               "管理用パスワードを入れて「締め切り日・休日設定を更新」を押すと反映されます。")
            except Exception:
                pass
            st.info("🔒 この欄は**管理者専用**です。受付の開始・締め切り、専用URLや配布用の案内書類の作成、提出されたNG日の取り込みができます。"
                    "使うには、管理用パスワードを入力してください。")
            pw = st.text_input("管理用パスワード", type="password", key="remote_admin_pw")
            if pw:
                if pw == str(admin_password):
                    st.session_state["remote_admin_ok"] = True
                    st.rerun()
                else:
                    st.error("パスワードが違います。")
            return

        try:
            settings = remote_settings()
            tokens = remote_tokens()
        except Exception as e:
            st.error(f"スプレッドシートを読み込めませんでした。接続設定を確認してください。詳細：{e}")
            return

        st.caption("毎月の流れ：① 受付を開始する → ② URLの案内書類を配る（初回のみ）→ ③ 提出状況を確認 → ④ 締め切る → ⑤ 提出されたNG日を取り込む")
        month_key = f"{year}-{month}"
        active = settings.get("受付中の年月", "")
        closed = settings.get("締め切り", "") == "True"
        if active == month_key:
            status = "🔒 締め切り済み" if closed else "🟢 受付中"
            st.markdown(f"**{year}年{month}月：{status}**（締め切り日の表示：{settings.get('締め切り日', '') or '未設定'}）")
        elif str(active).startswith(PERIOD_KEY_PREFIX):
            _ap = TouchokuPeriod.from_key(active)
            st.warning(f"現在は、期間指定版のページで「{_ap.title() if _ap else active}」の受付をしています。"
                       f"ここで{year}年{month}月の受付を開始すると、期間指定版の受付は終わります。")
        elif active:
            st.warning(f"現在の受付は「{active.replace('-', '年')}月」です。上で選んでいる{year}年{month}月とは異なります。")
        else:
            st.markdown("**現在、受付していません。**")

        # --- 受付の開始・締め切り ---
        deadline = st.text_input("先生方に表示する締め切り日（例：10月25日）", value=settings.get("締め切り日", "") if active == month_key else "", key="remote_deadline")
        _n_days = calendar.monthrange(year, month)[1]
        _saved_locked = [d for d in parse_holiday_list(settings.get("入力不可日", "")) if d <= _n_days] if active == month_key else []
        locked_sel = sorted(locked_days)
        locked_text = ",".join(map(str, locked_sel))
        if locked_sel:
            st.caption("先生がNGを入力できない日（上のカレンダーの「別途決定」）：" + "・".join(f"{month}/{d}" for d in locked_sel))
        c1, c2, c3 = st.columns(3)
        if c1.button(f"{year}年{month}月の受付を開始する", key="remote_open", use_container_width=True,
                     help="上で設定した年月・休日・別途決定の日で受付を始めます。先生の名簿と専用URLも更新します。"):
            remote_sync_doctors(doctor_rows)
            remote_save_settings(**{"受付中の年月": month_key, "特別休日": ",".join(map(str, sorted(custom_holidays))),
                                    "翌月1日特別休日": bool(next_month_special_holiday), "締め切り": False, "締め切り日": deadline,
                                    "入力不可日": locked_text})
            st.rerun()
        if active == month_key:
            if not closed and c2.button("🔒 締め切る（先生側の編集を止める）", key="remote_close", use_container_width=True):
                remote_save_settings(**{"締め切り": True})
                st.rerun()
            if closed and c2.button("受付を再開する", key="remote_reopen", use_container_width=True):
                remote_save_settings(**{"締め切り": False})
                st.rerun()
            if c3.button("締め切り日・休日設定を更新", key="remote_update", use_container_width=True,
                         help="受付中に締め切り日を変えた場合に押します（上のカレンダーの「休日にする」「別途決定」は自動で反映されます）。"):
                remote_save_settings(**{"特別休日": ",".join(map(str, sorted(custom_holidays))),
                                        "翌月1日特別休日": bool(next_month_special_holiday), "締め切り日": deadline,
                                        "入力不可日": locked_text})
                st.rerun()
            # 上のカレンダーの「休日にする」「別途決定」を変えたら、先生側の画面にも自動で反映する
            saved_holidays = parse_holiday_list(settings.get("特別休日", ""))
            saved_next = settings.get("翌月1日特別休日", "") == "True"
            cal_differs = saved_holidays != sorted(custom_holidays) or sorted(locked_sel) != sorted(_saved_locked) \
                    or saved_next != bool(next_month_special_holiday)
            if cal_differs and not cal_changed:
                st.warning("上のカレンダーの「休日にする」「別途決定」が、先生側の画面と違います。"
                           "いまの内容を先生側に反映するには「締め切り日・休日設定を更新」を押してください。")
            elif cal_differs:
                try:
                    remote_save_settings(**{"特別休日": ",".join(map(str, sorted(custom_holidays))),
                                            "翌月1日特別休日": bool(next_month_special_holiday),
                                            "入力不可日": locked_text})
                    st.toast("上のカレンダーの「休日にする」「別途決定」を、先生側の画面に反映しました。")
                except Exception as e:
                    st.warning(f"先生側の画面に休日・別途決定を反映できませんでした。「締め切り日・休日設定を更新」を押してください。詳細：{e}")

        # --- 専用URL ---
        st.markdown("##### 🔗 先生ごとの専用URL（固定）")
        missing = [name for name, _ in doctor_rows if name not in tokens]
        if missing:
            st.caption(f"まだURLがない先生：{'、'.join(missing)}")
            if st.button("名簿の先生全員のURLを用意する", key="remote_sync"):
                remote_sync_doctors(doctor_rows)
                st.rerun()
        url_rows = [{"先生": name, "専用URL": remote_doctor_url(tokens[name]["キー"])} for name, _ in doctor_rows if name in tokens]
        if url_rows:
            st.dataframe(pd.DataFrame(url_rows), hide_index=True, use_container_width=True,
                         column_config={"専用URL": st.column_config.LinkColumn("専用URL（右上のアイコンでコピーできます）")})
            st.download_button("📥 URL一覧をCSVでダウンロード", pd.DataFrame(url_rows).to_csv(index=False).encode("utf-8-sig"),
                               file_name="NG日入力_専用URL一覧.csv", mime="text/csv", key="remote_url_csv")
            # 紙で配る用：QRコード付きの案内書類（A4・1人1枚）
            card_targets = st.multiselect("案内書類を作る先生（URLを作り直した先生だけ印刷する場合などは絞り込めます）",
                                          [r["先生"] for r in url_rows], default=[r["先生"] for r in url_rows], key="remote_card_targets")
            if card_targets:
                url_by_name = {r["先生"]: r["専用URL"] for r in url_rows}
                try:
                    st.download_button(f"🖨️ 配布用の案内書類（QRコード付き・{len(card_targets)}名分）をダウンロード",
                                       remote_qr_cards_html([(n, url_by_name[n]) for n in card_targets]).encode("utf-8"),
                                       file_name="当直NG日入力用URLのご案内.html", mime="text/html", key="remote_qr_cards", type="primary")
                    st.caption("ダウンロードしたファイルを開くとブラウザに表示されるので、「印刷する」を押してA4・片面で印刷します（1人1枚）。"
                               "点線で三つ折りにすると（上の線は外側へ、下の線は内側へ折る）、宛名と題名が表に出て、QRコードは内側に隠れます。封筒に入れて各先生に渡してください。")
                except ImportError:
                    st.error("案内書類を作るには、requirements.txt に「reportlab」の1行を追加してください。")
            if st.checkbox("URLを作り直す（URLが漏れてしまった場合）", key="remote_regen_open"):
                target = st.selectbox("作り直す先生", [r["先生"] for r in url_rows], key="remote_regen_target")
                st.caption("作り直すと、古いURLは使えなくなります。新しいURLをその先生に送り直してください。")
                if st.button("この先生のURLを作り直す", key="remote_regen"):
                    remote_regenerate_key(target)
                    st.rerun()

        # --- 提出状況と取り込み ---
        st.markdown(f"##### 📋 {year}年{month}月の提出状況")
        if st.button("🔄 最新の状態に更新", key="remote_refresh"):
            remote_clear_cache()
            st.rerun()
        submissions = remote_latest_submissions(month_key)
        status_rows = []
        for name, _ in doctor_rows:
            r = submissions.get(name)
            text = r.get("NG日", "") if r else ""
            status_rows.append({
                "先生": name,
                "状況": "✅ 提出済み" if r else "⏳ 未提出",
                "最終保存": r.get("保存日時", "") if r else "",
                "NG日": (ng_text_summary(text, year, month, custom_holidays) or "NG日なし") if r else "",
                "コメント": r.get("コメント", "") if r else "",
            })
        done = sum(1 for r in status_rows if r["状況"].startswith("✅"))
        with_comment = sum(1 for r in status_rows if r["コメント"].strip())
        st.caption(f"提出済み {done} / {len(status_rows)} 名" + (f"　💬 コメントあり {with_comment} 名（表の「コメント」列）" if with_comment else ""))
        st.dataframe(pd.DataFrame(status_rows), hide_index=True, use_container_width=True)

        flash = st.session_state.pop("remote_import_message", None)
        if flash:
            st.success(flash)
        _locked_now = set(parse_holiday_list(settings.get("入力不可日", ""))) if active == month_key else set()

        def _drop_locked(r):
            # 先生がNGを入力できない日（年末年始など）のNGは取り込まない
            if not _locked_now:
                return r
            kept = {d: v for d, v in parse_ng_dict(r.get("NG日", ""), year, month).items() if d not in _locked_now}
            return {**r, "NG日": ng_dict_to_text(kept)}
        importable = {name: _drop_locked(r) for name, r in submissions.items() if name in {n for n, _ in doctor_rows}}
        if importable:
            mode = "replace"   # 取り込むと、その先生のNG日は提出内容に置き換わる
            changes = []
            for name, r in importable.items():
                submitted = parse_ng_dict(r.get("NG日", ""), year, month)
                result, _ = planned_import(name, year, month, submitted, mode)
                now = {d: st.session_state.get(f"ng_{name}_{year}_{month}_{d}", "OK") for d in day_is_holiday}
                now = {d: v for d, v in now.items() if v != "OK"}
                if result != now:
                    changes.append({"先生": name,
                                    "今のNG日（アプリ）": ng_text_summary(ng_dict_to_text(now), year, month, custom_holidays) or "なし",
                                    "取り込むNG日（提出）": ng_text_summary(r.get("NG日", ""), year, month, custom_holidays) or "なし",
                                    "コメント": str(r.get("コメント", "")).strip()})
            if changes:
                st.markdown("**取り込むと変わる先生**")
                st.dataframe(pd.DataFrame(changes), hide_index=True, use_container_width=True)
                changed_names = [c["先生"] for c in changes]
                # 取り込む先生を選べる（はじめは全員。対象の先生が変わったら選び直しになる）
                sel_key = "remote_import_targets_" + hashlib.sha256("|".join(changed_names).encode()).hexdigest()[:12]
                targets = st.multiselect("取り込む先生（×で外せます）", changed_names, default=changed_names, key=sel_key)
                targets = [n for n in (targets or []) if n in changed_names]
                st.button(f"⬇️ 選んだ先生のNG日を取り込む（{len(targets)}名分）", key="remote_import", type="primary",
                          disabled=not targets,
                          on_click=import_remote_ng, args=(year, month, {n: importable[n] for n in targets}, mode))
                st.caption("取り込むと、その先生のNG日は提出内容に置き換わります（CSVなどで入れていたNGは消えます）。"
                           "選ばなかった先生や、提出していない先生のNG日は変わりません。取り込んだあとも、下のカレンダーで個別に直せます。")
            else:
                st.caption("提出内容は、すべてアプリに取り込み済みです。")
        # --- 取り込みの取り消し ---
        undoable = import_undo_doctors(year, month, [n for n, _ in doctor_rows])
        if undoable:
            with st.container(border=True):
                st.markdown("**↩️ 取り込みを取り消す**")
                st.caption("URLから取り込んだNG日を取り消して、取り込む前のNG日（CSVや手入力の分）に戻します。"
                           "取り込んだあとにカレンダーで手直しした分も元に戻ります。先生の提出内容そのものは消えないので、もう一度取り込めます。")
                cu1, cu2 = st.columns([3, 2])
                undo_target = cu1.selectbox("取り消す先生", undoable, key="remote_undo_target", label_visibility="collapsed")
                cu2.button("この先生の取り込みを取り消す", key="remote_undo_one", use_container_width=True,
                           on_click=undo_remote_import, args=(year, month, [undo_target]))
                if len(undoable) > 1:
                    st.button(f"取り込んだ{len(undoable)}名全員の取り込みを取り消す", key="remote_undo_all",
                              on_click=undo_remote_import, args=(year, month, undoable))
        unknown = [name for name in submissions if name not in {n for n, _ in doctor_rows}]
        if unknown:
            st.caption(f"名簿にない先生の提出（取り込み対象外）：{'、'.join(unknown)}")

        # --- 先生の提出のリセット ---
        if submissions and st.checkbox("先生の提出をリセットする（入力をやり直してもらう場合など）", key="remote_reset_open"):
            reset_names = list(submissions)
            # 選択肢（提出済みの先生）が変わったら選び直しになるよう、名前の組み合わせで欄を分ける
            reset_key = "remote_reset_targets_" + hashlib.sha256("|".join(reset_names).encode()).hexdigest()[:12]
            reset_targets = [n for n in (st.multiselect("リセットする先生（複数選べます）", reset_names, key=reset_key) or [])
                             if n in reset_names]
            st.caption("リセットすると、その先生は「未提出」に戻り、先生のURLの画面もNG日・コメントが空の状態になります。"
                       "これまでの提出は「変更の履歴」に残ります。すでにアプリに取り込んだNG日は変わらないので、"
                       "必要なら「↩️ 取り込みを取り消す」も使ってください。")

            def _do_reset(names):
                for name in names:
                    remote_reset_submission(month_key, name)
                st.session_state["remote_reset_confirm"] = False   # 確認のチェックを外しておく（続けて押してしまわないように）
                st.session_state["remote_import_message"] = f"{'、'.join(names)} 先生の提出をリセットしました（未提出に戻しました）。"
            if reset_targets:
                if st.checkbox(f"選んだ{len(reset_targets)}名（{'、'.join(reset_targets)}）の提出をリセットしてよい", key="remote_reset_confirm"):
                    st.button(f"🗑️ {len(reset_targets)}名の提出をリセットする", key="remote_reset_do",
                              on_click=_do_reset, args=(reset_targets,))

        # （この欄自体が折りたたみの中にあるため、履歴はチェックボックスで開く）
        if st.checkbox("変更の履歴を表示（だれが・いつ・どう変えたか）", key="remote_history_open"):
            history = [r for r in remote_read("submissions") if r.get("年月") == month_key]
            if history:
                st.dataframe(pd.DataFrame([{"保存日時": r.get("保存日時", ""), "先生": r.get("先生", ""),
                                            "NG日": "（管理者がリセット）" if r.get("種別", "") == "リセット"
                                            else (ng_text_summary(r.get("NG日", ""), year, month, custom_holidays) or "NG日なし"),
                                            "コメント": r.get("コメント", "")}
                                           for r in reversed(history)]), hide_index=True, use_container_width=True)
            else:
                st.caption("まだ履歴はありません。")


# ==========================================
# 当直の枠は足りそう？（目安）：全体・枠の種類別・日ごとの見える化
# ==========================================
def _cap_num(value, default=None):
    """医師条件の数値を読む。空欄なら default（上限なし＝None）。"""
    v = pd.to_numeric(pd.Series([value]), errors="coerce").iloc[0]
    return default if pd.isna(v) else max(0, int(v))


def doctor_capacity(row):
    """
    1人の先生が入れる回数の上限を、条件から求める（目安）。
    月間最大回数・休日最大回数・宿直上限・日直上限・各枠の上限のうち、いちばん厳しいものが効く。
    """
    total = _cap_num(row.get("月間最大回数"), 0)
    hol = _cap_num(row.get("休日最大回数"), 0)
    per = {s: _cap_num(row.get(s + "上限"), 0) for s in ALL_SHIFT_TYPES}
    g_night = _cap_num(row.get("宿直上限"))
    g_day = _cap_num(row.get("日直上限"))
    night = min([total, sum(per[s] for s in NIGHT_SHIFTS)] + ([g_night] if g_night is not None else []))
    day = min([total, hol, sum(per[s] for s in DAY_SHIFTS)] + ([g_day] if g_day is not None else []))
    cap = {s: min(per[s], night if s in NIGHT_SHIFTS else day) for s in ALL_SHIFT_TYPES}
    return {"total": min(total, night + day), "night": night, "day": day,
            "holiday": min(hol, night + day), "per": cap}


def slot_needs(year, month, num_days, custom_holidays, locked_days, multi):
    """埋める必要がある枠の数（枠の種類別・宿直合計・日直合計・休日の枠）。"""
    per = dict.fromkeys(ALL_SHIFT_TYPES, 0)
    holiday = 0
    for d in range(1, num_days + 1):
        if d in locked_days:
            continue
        is_hol = is_holiday_date(datetime.date(year, month, d), year, month, custom_holidays)
        for s in ALL_SHIFT_TYPES:
            if s in DAY_SHIFTS and not is_hol:
                continue
            n = multi.get((d, s), 1)
            per[s] += n
            holiday += n if is_hol else 0
    return {"per": per, "night": sum(per[s] for s in NIGHT_SHIFTS), "day": sum(per[s] for s in DAY_SHIFTS),
            "holiday": holiday, "total": sum(per.values())}


def capacity_rows(staff_df, needs):
    """横棒に出す行：(名前, 必要な枠数, 入れる回数の合計, 上限の列名)。必要数0の行は出さない。"""
    caps = [doctor_capacity(r) for _, r in staff_df.iterrows()]
    rows = [("全体（日直＋宿直）", needs["total"], sum(c["total"] for c in caps), "月間最大回数"),
            ("宿直（A・B・外来の合計）", needs["night"], sum(c["night"] for c in caps), "宿直上限"),
            ("日直（A・B・外来の合計）", needs["day"], sum(c["day"] for c in caps), "日直上限"),
            ("休日の枠（日直＋宿直）", needs["holiday"], sum(c["holiday"] for c in caps), "休日最大回数")]
    rows += [(s, needs["per"][s], sum(c["per"][s] for c in caps), s + "上限") for s in ALL_SHIFT_TYPES]
    return [r for r in rows if r[1] > 0], sum(c["total"] for c in caps)


def capacity_status(need, cap):
    """'short'（足りない）・'tight'（ぎりぎり）・'ok'（足りている）。"""
    if cap < need:
        return "short"
    if cap - need <= max(1, -(-need // 10)):   # 余裕が1割以下（最低1回）ならぎりぎり
        return "tight"
    return "ok"


CAPACITY_CSS = """
<style>
.capbars{margin:.2rem 0 .4rem}
.capbars .r{display:flex;align-items:center;gap:10px;margin:7px 0}
.capbars .n{width:12.5rem;flex:0 0 auto;font-size:.9rem;text-align:right;line-height:1.2}
.capbars .b{position:relative;flex:1;height:20px;background:#eef0f3;border-radius:4px;min-width:6rem}
.capbars .f{height:100%;border-radius:4px}
.capbars .need{position:absolute;top:-5px;bottom:-5px;width:3px;background:#1d2733;border-radius:2px}
.capbars .m{width:10.5rem;flex:0 0 auto;font-weight:700;font-size:.9rem;line-height:1.2}
.capbars .ok{color:#2e7d43} .capbars .tight{color:#c77700} .capbars .short{color:#c62828}
.caplg{font-size:.8rem;color:#555;margin-top:2px}
@media (max-width:640px){.capbars .r{flex-wrap:wrap;gap:2px 8px;margin:9px 0}.capbars .n{width:100%;text-align:left;font-size:.85rem}.capbars .n br{display:none}.capbars .n span{margin-left:6px}.capbars .m{width:7.5rem;font-size:.8rem}}
</style>
"""

_CAP_COLORS = {"ok": "#5cb85c", "tight": "#f0ad4e", "short": "#d9534f"}


def capacity_bars_html(rows):
    """枠の種類ごとの横棒。棒＝入れる回数の合計、黒い線＝必要な枠数。"""
    top = max([max(need, cap) for _, need, cap, _ in rows] + [1]) * 1.08
    out = ["<div class='capbars'>"]
    for name, need, cap, _ in rows:
        stt = capacity_status(need, cap)
        msg = (f"{need - cap}枠 足りません" if stt == "short"
               else f"余裕 {cap - need}回（ぎりぎり）" if stt == "tight" else f"余裕 {cap - need}回")
        out.append(f"<div class='r'><div class='n'>{html_escape(name)}<br><span style='font-size:.75rem;color:#667'>必要 {need}枠／入れる {cap}回</span></div>"
                   f"<div class='b'><div class='f' style='width:{cap / top * 100:.1f}%;background:{_CAP_COLORS[stt]}'></div>"
                   f"<div class='need' style='left:{need / top * 100:.1f}%'></div></div>"
                   f"<div class='m {stt}'>{msg}</div></div>")
    out.append("</div><div class='caplg'>🟩 足りている　🟧 ぎりぎり（余裕が1割以下）　🟥 足りない　｜　棒＝その枠に入れる回数の合計　黒い線＝必要な枠数</div>")
    return "".join(out)


# ==========================================
# 期間指定版（年末年始など）：今月の1日から、翌月の指定日までを1つの当直表として作る
# 別ページ（pages フォルダ）と、先生用の入力画面で使う。通常の画面では使わない。
# 期間の中の日は「通し番号」（1＝今月1日、…）で扱い、画面には「12/31」のように表示する。
# ==========================================
PERIOD_KEY_PREFIX = "P:"
# 期間指定版では datetime.date / calendar を期間用に差し替えることがあるので、本物を別名で持っておく
_REAL_DATE = datetime.date
_REAL_MONTHRANGE = calendar.monthrange


class TouchokuPeriod:
    """当直を作る期間（今月1日〜翌月の指定日）。"""

    def __init__(self, year, month, extra_days=0):
        self.year, self.month = int(year), int(month)
        self.start = _REAL_DATE(self.year, self.month, 1)
        month_last = _REAL_DATE(self.year, self.month, _REAL_MONTHRANGE(self.year, self.month)[1])
        nxt = month_last + datetime.timedelta(days=1)
        self.extra = max(0, min(int(extra_days), _REAL_MONTHRANGE(nxt.year, nxt.month)[1]))
        self.end = month_last + datetime.timedelta(days=self.extra)
        self.n = (self.end - self.start).days + 1
        self.dates = [self.start + datetime.timedelta(days=i) for i in range(self.n)]
        self.key = f"{PERIOD_KEY_PREFIX}{self.start.isoformat()}~{self.end.isoformat()}"

    @classmethod
    def from_key(cls, key):
        m = re.fullmatch(r"P:(\d{4})-(\d{2})-01~(\d{4})-(\d{2})-(\d{2})", str(key).strip())
        if not m:
            return None
        y, mo = int(m.group(1)), int(m.group(2))
        end = _REAL_DATE(int(m.group(3)), int(m.group(4)), int(m.group(5)))
        month_last = _REAL_DATE(y, mo, _REAL_MONTHRANGE(y, mo)[1])
        return cls(y, mo, (end - month_last).days)

    def date(self, i):
        return self.dates[i - 1]

    def index(self, dt):
        if isinstance(dt, datetime.datetime):
            dt = dt.date()
        return (dt - self.start).days + 1 if self.start <= dt <= self.end else None

    def label(self, i):
        d = self.date(i)
        return f"{d.month}/{d.day}"

    def example(self, next_month=False):
        """エラー文の入力例（今月の日なら「12/30」、翌月の日なら「1/2」など）。"""
        if next_month and self.extra:
            d = self.date(self.n - self.extra + min(2, self.extra))
        else:
            d = self.date(self.n - self.extra - 1)
        return f"{d.month}/{d.day}"

    def title(self):
        s, e = self.start, self.end
        if (s.year, s.month) == (e.year, e.month):
            return f"{s.year}年{s.month}/{s.day}〜{e.month}/{e.day}"
        if s.year == e.year:
            return f"{s.year}年{s.month}/{s.day}〜{e.month}/{e.day}"
        return f"{s.year}年{s.month}/{s.day}〜{e.year}年{e.month}/{e.day}"


def period_is_holiday(P, i, custom_holidays):
    dt = P.date(i)
    return dt.weekday() >= 5 or jpholiday.is_holiday(dt) or i in custom_holidays


def period_is_holiday_date(P, dt, custom_holidays):
    i = P.index(dt)
    return dt.weekday() >= 5 or jpholiday.is_holiday(dt) or (i is not None and i in custom_holidays)


def period_pm_restricted(P, i, weekdays, custom_holidays, next_special=False):
    """翌日が勤務日の場合だけ、指定曜日の宿直を制限する（期間の翌日が病院独自の休みなら next_special）。"""
    dt = P.date(i)
    nd = dt + datetime.timedelta(days=1)
    next_is_holiday = period_is_holiday_date(P, nd, custom_holidays) or (next_special and nd == P.end + datetime.timedelta(days=1))
    return dt.weekday() in weekdays and not next_is_holiday


def period_checked_items(value, P, ng=False):
    """「12/30:A宿直, 1/2」や通し番号「3」を、(通し番号, 種類) のリストにする。"""
    text = clean_text(value).translate(str.maketrans('０１２３４５６７８９，：／', '0123456789,:/'))
    items = []
    for token in text.split(','):
        if not token.strip():
            continue
        parts = token.strip().split(':')
        head = parts[0].strip()
        if len(parts) > 2:
            raise ValueError(f'「{token}」：書き方を確認してください（例：{P.example()} または {P.example(True)}:A宿直）。')
        m = re.fullmatch(r'(\d{1,2})/(\d{1,2})', head)
        if m:
            mo, dd = int(m.group(1)), int(m.group(2))
            i = next((k for k, dt in enumerate(P.dates, 1) if (dt.month, dt.day) == (mo, dd)), None)
            if i is None:
                raise ValueError(f'「{token}」：{P.title()}の期間にない日付です。')
        elif re.fullmatch(r'\d+', head):
            i = int(head)
            if not 1 <= i <= P.n:
                raise ValueError(f'「{token}」：翌月の日は「{P.example(True)}」のように月も書いてください。')
        else:
            raise ValueError(f'「{token}」：日付は「{P.example()}」のように書き、複数は半角カンマで区切ってください。')
        kind = parts[1].strip() if len(parts) == 2 else ('全NG' if ng else None)
        if kind not in (['全NG', '日NG', '宿NG', 'OK'] if ng else ALL_SHIFT_TYPES + [None]):
            raise ValueError(f'「{token}」：枠名・NGの種類を確認してください。')
        items.append((i, kind))
    if ng:
        by_day = {}
        for i, kind in items:
            if i in by_day and by_day[i] != kind:
                raise ValueError(f'{P.label(i)}のNG指定が重複して異なっています。1種類にしてください。')
            by_day[i] = kind
    return list(dict.fromkeys(items))


def period_items_text(items, P, ng=False):
    """(通し番号, 種類) を「12/30,1/2:宿NG」のような文字にする。"""
    out = []
    for i, kind in items:
        if kind is None or (ng and kind == '全NG'):
            out.append(P.label(i))
        elif not (ng and kind == 'OK'):
            out.append(f"{P.label(i)}:{kind}")
    return ",".join(out)


def period_parse_ng(value, P):
    try:
        return {i: kind for i, kind in period_checked_items(value, P, True) if kind != 'OK'}
    except ValueError:
        return {}


def period_parse_requests(value, P):
    try:
        items = period_checked_items(value, P, False)
    except ValueError:
        return [], []
    return [i for i, kind in items if kind is None], [(i, kind) for i, kind in items if kind is not None]


def period_ng_text(values_by_day, P):
    """{通し番号: NG種類} を「12/30,1/2:宿NG」にする。"""
    return period_items_text([(i, values_by_day[i]) for i in sorted(values_by_day)], P, ng=True)


def period_ng_summary(text, P, holidays):
    parts = []
    for i, v in sorted(period_parse_ng(text, P).items()):
        shown = effective_ng(v, period_is_holiday(P, i, holidays))
        if shown == "全NG":
            parts.append(P.label(i))
        elif shown == "日NG":
            parts.append(f"{P.label(i)}(日直NG)")
        elif shown == "宿NG":
            parts.append(f"{P.label(i)}(宿直NG)")
    return "、".join(parts)


def period_ng_calendar_days(P, holidays, next_special, pm_weekdays, values_by_day, locked_days=()):
    """NGカレンダー部品に渡す日ごとのデータ（期間版）。翌月に入った日は「1/1」のように月も表示する。"""
    days = []
    for i in range(1, P.n + 1):
        dt = P.date(i)
        hol = period_is_holiday(P, i, holidays)
        days.append({
            "day": (P.label(i) if (dt.month != P.month or i == 1) else dt.day),
            "weekday": WEEKDAYS_JA[dt.weekday()],
            "options": ["OK", "全NG", "日NG", "宿NG"] if hol else ["OK", "宿NG"],
            "value": "OK" if i in locked_days else effective_ng(values_by_day.get(i, "OK"), hol),
            "warning": period_pm_restricted(P, i, pm_weekdays, holidays, next_special),
            "kind": "saturday" if dt.weekday() == 5 and not jpholiday.is_holiday(dt) and i not in holidays else ("holiday" if hol else "weekday"),
            "locked": i in locked_days,
        })
    return days


def period_add_interval_constraints(
    model,
    shifts,
    doctors,
    daily_active_shifts,
    num_days,
    period,
    min_intervals,
    past_worked_dates,
    future_worked_dates,
    absolute_req_days,
    absolute_req_specific,
):
    """確定同士は維持し、自動割当と全勤務の間隔を制限する。"""
    for doc in doctors:
        gap = min_intervals[doc]
        fixed_days = set(absolute_req_days[doc]) | {d for d, s in absolute_req_specific[doc]}
        external = set(past_worked_dates.get(doc, [])) | set(future_worked_dates.get(doc, []))
        worked = {}
        for d in range(1, num_days + 1):
            variables = [shifts[d, doc, s] for s in daily_active_shifts.get(d, []) if (d, doc, s) in shifts]
            if not variables: continue
            worked[d] = model.NewBoolVar(f'working_{doc}_{d}')
            model.AddMaxEquality(worked[d], variables)
            dt = period.date(d)
            if d not in fixed_days and any(abs((dt - ext).days) <= gap for ext in external):
                model.Add(worked[d] == 0)
        for d in worked:
            for other in range(d + 1, min(num_days, d + gap) + 1):
                if other in worked and not (d in fixed_days and other in fixed_days):
                    model.Add(worked[d] + worked[other] <= 1)


def _period_generate_core(P, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None, locked_days=()):
    target_year, target_month = P.year, P.month
    locked_days = set(locked_days)   # 「別途決定」の日：別の方法で当番を決めるので、ここでは割り当てない（確定当直だけ残す）
    num_days = P.n
    days = range(1, num_days + 1)

    def is_holiday(d):
        return period_is_holiday(P, d, custom_holidays)

    def safe_int(val, default_val):
        if pd.isna(val): return default_val
        try:
            return int(float(val))
        except (ValueError, TypeError):
            return default_val

    doctors = staff_df['先生の名前'].astype(str).tolist()
    ng_days_dict = {}
    req_days = {}
    req_specific = {}
    req_priority = {}
    hard_weekdays = {}
    min_intervals = {}
    min_shifts_total = {}
    max_shifts_total = {}
    max_hol_shifts_per_doc = {}
    max_shifts_per_type = {}
    group_caps = {}

    absolute_req_days = {doc: [] for doc in doctors}
    absolute_req_specific = {doc: [] for doc in doctors}
    past_worked_dates = {doc: [] for doc in doctors}
    future_worked_dates = {doc: [] for doc in doctors}

    # --- 確定当直（先月・今月・来月）の読み取り ---
    fixed_specific = {doc: set() for doc in doctors}
    if fixed_df is not None:
        for _, row in fixed_df.iterrows():
            date_obj = parse_shift_date(row.get('日付', ''), target_year, target_month)
            if date_obj is None:
                continue
            for s_type in ALL_SHIFT_TYPES:
                if s_type in row and pd.notna(row[s_type]):
                    for doc_val in re.split(r'[、,]+', str(row[s_type])):
                        doc_val = doc_val.strip()
                        if doc_val not in doctors:
                            continue
                        if P.index(date_obj) is not None:
                            absolute_req_specific[doc_val].append((P.index(date_obj), s_type))
                            fixed_specific[doc_val].add((P.index(date_obj), s_type))
                        elif date_obj < P.start:
                            past_worked_dates[doc_val].append(date_obj)
                        else:
                            future_worked_dates[doc_val].append(date_obj)

    # --- 医師条件の読み取り ---
    for _, row in staff_df.iterrows():
        doc = str(row['先生の名前'])
        hard_weekdays[doc] = pm_duty_weekdays(row.get('翌日PM duty', ''))
        ng_days_dict[doc] = period_parse_ng(row.get(NG_COLUMN, ''), P)
        req_days[doc], req_specific[doc] = period_parse_requests(row.get(REQUEST_COLUMN, ''), P)
        req_days[doc] = [d for d in req_days[doc] if d not in locked_days]
        req_specific[doc] = [(d, s) for d, s in req_specific[doc] if d not in locked_days]
        req_priority[doc] = safe_int(row.get('希望優先度(数字が大きいほど優先)'), 1)
        min_intervals[doc] = safe_int(row.get('最低空ける日数'), 5)
        min_shifts_total[doc] = safe_int(row.get('月間最小回数'), 0)
        max_shifts_total[doc] = safe_int(row.get('月間最大回数'), 5)
        max_hol_shifts_per_doc[doc] = safe_int(row.get('休日最大回数'), 4)
        max_shifts_per_type[doc] = {s: safe_int(row.get(s + '上限'), 2) for s in ALL_SHIFT_TYPES}
        group_caps[doc] = {col: safe_int(row.get(col), None) for col in GROUP_CAP_COLUMNS}

    # --- その日に存在する枠（平日は宿直のみ、休日は日直・宿直） ---
    daily_active_shifts = {d: (NIGHT_SHIFTS + DAY_SHIFTS) if is_holiday(d) else list(NIGHT_SHIFTS) for d in days}

    # 存在しない枠の指定はエラーにする（優先度に関係なく同じ扱い）
    invalid_requests = []
    for doc in doctors:
        for d, slot in req_specific[doc]:
            if slot not in daily_active_shifts[d]:
                invalid_requests.append(f'{doc}：{P.label(d)}の{slot}は設定されていません。平日に日直を希望する場合は、上のカレンダーでその日を「休日にする」にしてください。')
        for d, slot in fixed_specific[doc]:
            if slot not in daily_active_shifts[d]:
                invalid_requests.append(f'{doc}：確定当直の{P.label(d)}の{slot}は、この日には設定されていない枠です。')
    if invalid_requests:
        return None, False, list(dict.fromkeys(invalid_requests)), None, None

    # --- 優先度100以上の希望を確定扱いにする ---
    for doc in doctors:
        if req_priority[doc] >= 100:
            absolute_req_days[doc].extend(req_days[doc])
            absolute_req_specific[doc].extend(req_specific[doc])
        absolute_req_specific[doc] = sorted(set(absolute_req_specific[doc]))
        specified_days = {d for d, s in absolute_req_specific[doc]}
        absolute_req_days[doc] = sorted(set(absolute_req_days[doc]) - specified_days)
        past_worked_dates[doc] = sorted(set(past_worked_dates[doc]))
        future_worked_dates[doc] = sorted(set(future_worked_dates[doc]))
        # 確定した日はNG指定より確定を優先する
        all_abs_dates = set(absolute_req_days[doc]) | specified_days
        ng_days_dict[doc] = {d: v for d, v in ng_days_dict[doc].items() if d not in all_abs_dates}

    def required_count(d, s):
        fixed_docs_count = sum(1 for doc in doctors if (d, s) in absolute_req_specific[doc])
        if d in locked_days:
            return fixed_docs_count
        return max(multi_slots_dict.get((d, s), 1), fixed_docs_count)

    def total_cap(doc):
        return max(max_shifts_total[doc], len(absolute_req_days[doc]) + len(absolute_req_specific[doc]))

    # ------------------------------------------------------------------
    # 通常計算と不足許容計算で共通の条件。
    # ここを直せば両方の計算に反映される（片方だけ直し忘れることがない）。
    # ------------------------------------------------------------------
    def add_common_constraints(model, x):
        worked_all = {}
        holiday_worked = {}
        blocked_by_ng = {"全NG": ALL_SHIFT_TYPES, "日NG": DAY_SHIFTS, "宿NG": NIGHT_SHIFTS}
        for doc in doctors:
            specific_days = {d for d, _ in absolute_req_specific[doc]}
            fixed_dates = set(absolute_req_days[doc]) | specific_days

            # 1日に入れる勤務は原則1つ（確定で複数ある日だけ例外）
            for d in days:
                fixed_count = sum(1 for sd, _ in absolute_req_specific[doc] if sd == d)
                model.Add(sum(x[(d, doc, s)] for s in daily_active_shifts[d]) <= max(1, fixed_count))

            # 「別途決定」の日は、確定当直以外は入れない
            for d in locked_days:
                if d in days:
                    for s in daily_active_shifts[d]:
                        if (d, s) not in absolute_req_specific[doc] and d not in absolute_req_days[doc]:
                            model.Add(x[(d, doc, s)] == 0)

            # NG日（全NG / 日NG / 宿NG）
            for d, ng_type in ng_days_dict[doc].items():
                for s in daily_active_shifts[d]:
                    if s in blocked_by_ng.get(ng_type, []):
                        model.Add(x[(d, doc, s)] == 0)

            # 翌日PM duty：宿直のみ制限。翌日が休日なら制限しない。確定日は対象外。
            for d in days:
                if d in fixed_dates:
                    continue
                if period_pm_restricted(P, d, hard_weekdays[doc], custom_holidays, next_month_special_holiday):
                    for s in NIGHT_SHIFTS:
                        model.Add(x[(d, doc, s)] == 0)

            # 確定（枠指定なしの日は、その日のいずれか1枠）
            for d in absolute_req_days[doc]:
                if d not in specific_days:
                    model.AddExactlyOne(x[(d, doc, s)] for s in daily_active_shifts[d])
            for d, s in absolute_req_specific[doc]:
                model.Add(x[(d, doc, s)] == 1)

            # 枠ごとの上限
            for s_type in ALL_SHIFT_TYPES:
                active_days = [d for d in days if s_type in daily_active_shifts[d]]
                if not active_days:
                    continue
                worked = [x[(d, doc, s_type)] for d in active_days]
                forced_vars = [x[(d, doc, s_type)] for d in active_days if (d, s_type) in absolute_req_specific[doc] or d in absolute_req_days[doc]]
                add_type_cap(model, worked, forced_vars, max_shifts_per_type[doc][s_type], num_days)

            # 宿直上限・日直上限（A・B・外来を合わせた回数。空欄なら制限なし）
            for col, slots in GROUP_CAP_COLUMNS.items():
                cap = group_caps[doc].get(col)
                if cap is None:
                    continue
                worked = [x[(d, doc, s)] for d in days for s in daily_active_shifts[d] if s in slots]
                if not worked:
                    continue
                forced_vars = [x[(d, doc, s)] for d in days for s in daily_active_shifts[d] if s in slots
                               and ((d, s) in absolute_req_specific[doc] or d in absolute_req_days[doc])]
                add_type_cap(model, worked, forced_vars, cap, num_days)

            # 月間上限
            all_vars = [x[(d, doc, s)] for d in days for s in daily_active_shifts[d]]
            model.Add(sum(all_vars) <= total_cap(doc))
            worked_all[doc] = all_vars

            # 休日上限
            hol_vars = [x[(d, doc, s)] for d in days if is_holiday(d) for s in daily_active_shifts[d]]
            abs_hol_count = sum(1 for d in absolute_req_days[doc] if is_holiday(d)) + sum(1 for d, _ in absolute_req_specific[doc] if is_holiday(d))
            model.Add(sum(hol_vars) <= max(max_hol_shifts_per_doc[doc], abs_hol_count))
            holiday_worked[doc] = hol_vars

        # 勤務間隔
        period_add_interval_constraints(
            model=model, shifts=x, doctors=doctors, daily_active_shifts=daily_active_shifts,
            num_days=num_days, period=P,
            min_intervals=min_intervals, past_worked_dates=past_worked_dates,
            future_worked_dates=future_worked_dates, absolute_req_days=absolute_req_days,
            absolute_req_specific=absolute_req_specific,
        )
        return worked_all, holiday_worked

    def build_schedule(value_of, missing_of=None):
        """計算結果を表（DataFrame）にする。不足があれば「⚠️不足(n名)」を書き込む。"""
        rows = []
        for d in days:
            date_obj = P.date(d)
            row = {"日付": f"{P.label(d)}({WEEKDAYS_JA[date_obj.weekday()]})", "平日/休日": "休日" if is_holiday(d) else "平日"}
            for s in ALL_SHIFT_TYPES:
                row[s] = "-"
            for s in daily_active_shifts[d]:
                names = [doc for doc in doctors if value_of((d, doc, s)) == 1]
                missing = missing_of((d, s)) if missing_of else 0
                if missing > 0:
                    names.append(f"⚠️不足({missing}名)")
                if names:
                    row[s] = "、".join(names)
            rows.append(row)
        return pd.DataFrame(rows, columns=RESULT_COLUMNS)

    # ==================================================================
    # 通常計算
    # ==================================================================
    model = cp_model.CpModel()
    shifts = {(d, doc, s): model.NewBoolVar(f'shift_d{d}_{doc}_{s}') for d in days for doc in doctors for s in daily_active_shifts[d]}
    objective_terms = []

    over_caps = {}
    for d in days:
        for s in daily_active_shifts[d]:
            over_caps[(d, s)] = model.NewIntVar(0, len(doctors), f'over_cap_d{d}_{s}')
            model.Add(sum(shifts[(d, doc, s)] for doc in doctors) == required_count(d, s) + over_caps[(d, s)])
            objective_terms.append(over_caps[(d, s)] * -50000)

    worked_all, holiday_worked = add_common_constraints(model, shifts)

    # 月間最小回数（目標。届かない分は減点）
    for doc in doctors:
        shortfall = model.NewIntVar(0, min_shifts_total[doc], f'min_shortfall_{doc}')
        model.Add(sum(worked_all[doc]) + shortfall >= min(min_shifts_total[doc], total_cap(doc)))
        objective_terms.append(shortfall * -10000)

    # 休日回数の偏りを減らす
    max_hol_shifts = model.NewIntVar(0, num_days * 6, 'max_hol_shifts')
    for doc in doctors:
        model.Add(sum(holiday_worked[doc]) <= max_hol_shifts)

    # 希望日（優先度100未満）は加点
    for doc in doctors:
        if req_priority[doc] < 100:
            weight = req_priority[doc] * 100
            for d in req_days[doc]:
                for s in daily_active_shifts[d]:
                    objective_terms.append(shifts[(d, doc, s)] * weight)
            for d, s_name in req_specific[doc]:
                objective_terms.append(shifts[(d, doc, s_name)] * weight)

    model.Maximize(sum(objective_terms) - max_hol_shifts * 1000)

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = 60.0
    solver.parameters.random_seed = random.randint(1, 10000)
    status = solver.Solve(model)

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        result_df = build_schedule(lambda key: solver.Value(shifts[key]))
        over_cap_warnings = []
        for d in days:
            date_obj = P.date(d)
            for s in daily_active_shifts[d]:
                if solver.Value(over_caps[(d, s)]) > 0:
                    over_cap_warnings.append(f"{P.label(d)}({WEEKDAYS_JA[date_obj.weekday()]}) の「{s}」枠")
        warnings = []
        if status == cp_model.FEASIBLE:
            warnings.append("必要人数を満たす案です。時間内に最適化を完了したことまでは確認できていません。")
        if over_cap_warnings:
            warnings.append("⚠️ **【重要】次の枠は、確定当直や「希望優先度100以上」の指定が重なったため、設定した人数を超えて配置しています（そのまま当直案にしています）:**")
            warnings.extend([f"・{w}" for w in over_cap_warnings])
        return result_df, True, warnings, past_worked_dates, future_worked_dates

    # ==================================================================
    # バックアップ（不足を許容する計算）
    # ==================================================================
    reasons = []
    if status == cp_model.MODEL_INVALID:
        return None, False, ["計算モデルが無効です。入力条件と数値の範囲を確認してください。"], None, None
    if status == cp_model.UNKNOWN:
        reasons.append("通常計算では時間内に案を見つけられませんでした。条件が不可能と確定したわけではありません。")
    try:
        relax_model = cp_model.CpModel()
        r_shifts = {(d, doc, s): relax_model.NewBoolVar(f'r_shift_d{d}_{doc}_{s}') for d in days for doc in doctors for s in daily_active_shifts[d]}
        dummies = {}
        r_excess = []
        for d in days:
            for s in daily_active_shifts[d]:
                # 不足人数はその枠の必要人数を超えない
                dummies[(d, s)] = relax_model.NewIntVar(0, required_count(d, s), f'dummy_d{d}_{s}')
                overflow = relax_model.NewIntVar(0, len(doctors), f'r_over_{d}_{s}')
                r_excess.append(overflow)
                relax_model.Add(sum(r_shifts[(d, doc, s)] for doc in doctors) + dummies[(d, s)] == required_count(d, s) + overflow)

        add_common_constraints(relax_model, r_shifts)

        # 不足人数を、画面で指定した優先順位の群ごとに最小化する。
        # 優先度の低い群から順に重みを決め、上位群の重みは
        # 「下位群すべてで起こりうる最大の損失」より大きくする。
        # こうすると、上位群の不足を1名減らすことが、下位群の不足をいくら減らすことよりも優先される。
        missing_score = 0
        accumulated_bound = 0
        for group in reversed(shortage_priority_groups(shortage_priority)):
            weight = accumulated_bound + 1
            group_keys = [key for key in dummies if key[1] in group]
            missing_score += weight * sum(dummies[key] for key in group_keys)
            accumulated_bound += weight * sum(required_count(d, s) for d, s in group_keys)
        excess_bound = len(r_excess) * len(doctors)
        relax_model.Minimize(missing_score * (excess_bound + 1) + sum(r_excess))

        relax_solver = cp_model.CpSolver()
        relax_solver.parameters.max_time_in_seconds = 15.0
        relax_status = relax_solver.Solve(relax_model)

        if relax_status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
            partial_df = build_schedule(lambda key: relax_solver.Value(r_shifts[key]),
                                        lambda key: relax_solver.Value(dummies[key]))
            bottlenecks = []
            missing_by_shift = {s: 0 for s in ALL_SHIFT_TYPES}
            for d in days:
                for s in daily_active_shifts[d]:
                    val = relax_solver.Value(dummies[(d, s)])
                    if val > 0:
                        bottlenecks.append(f"・{P.label(d)} の「{s}」")
                        missing_by_shift[s] += val
            if bottlenecks:
                reasons.append(shortage_priority_text(shortage_priority))
                if relax_status == cp_model.FEASIBLE:
                    reasons.append("計算時間内に得られた案です。優先順位に沿った不足の最小化が完了したことまでは確認できていません。")
                reasons.append("🚨 **以下の枠で必要人数が不足しています:**")
                reasons.extend(bottlenecks)
                reasons.append("")
                reasons.append("📊 **【不足している枠の合計】**")
                for s, count in sorted(missing_by_shift.items(), key=lambda x: x[1], reverse=True):
                    if count > 0:
                        reasons.append(f"・{s}： 計 {count} 枠不足")
            return partial_df, False, reasons, past_worked_dates, future_worked_dates
        if relax_status == cp_model.INFEASIBLE:
            reasons.append("不足枠を許容しても条件が矛盾し、当直案を作成できませんでした。確定指定と上限を確認してください。")
        elif relax_status == cp_model.MODEL_INVALID:
            reasons.append("不足枠計算のモデルが無効です。入力値を確認してください。")
        else:
            reasons.append("不足枠を含む案も時間内に見つかりませんでした。条件が不可能と確定したわけではありません。")
    except Exception as e:
        reasons.append(f"⚠️ 部分的な当直表の作成中にもエラーが発生しました。詳細: {e}")

    return None, False, reasons, None, None


def period_validate_staff_inputs(df, P):
    df = df.copy()
    errors = []
    if '先生の名前' not in df:
        return df, ['医師条件CSVに「先生の名前」の列がありません。']
    df['先生の名前'] = df['先生の名前'].map(clean_text)
    blank = df['先生の名前'].eq('')
    for i, row in df[blank].iterrows():
        if any(clean_text(v) for k, v in row.items() if k != '先生の名前'):
            errors.append(f'医師条件 {i+1}行目：条件が入力されていますが医師名が空欄です。')
    df = df[~blank].reset_index(drop=True)
    for name in df.loc[df['先生の名前'].duplicated(), '先生の名前'].unique():
        errors.append(f'医師名「{name}」が重複しています。')
    for col, default in NUMERIC_DEFAULTS.items():
        if col not in df:
            df[col] = default
    for col in [NG_COLUMN, REQUEST_COLUMN, '翌日PM duty'] + list(GROUP_CAP_COLUMNS):
        if col not in df:
            df[col] = ''
    for col in GROUP_CAP_COLUMNS:
        df[col] = df[col].astype(object)   # 空欄（制限なし）と数字が混ざるため
    for i, row in df.iterrows():
        name = row['先生の名前']
        if any(c in name for c in ',、') or name == '-' or '⚠️不足' in name:
            errors.append(f'医師名「{name}」：カンマ・読点・不足表示・単独の「-」は使用できません。')
        for col, default in NUMERIC_DEFAULTS.items():
            raw = row[col]
            try:
                n = default if not clean_text(raw) else float(raw)
                if not math.isfinite(n) or n < 0 or n != int(n) or n > 1000000:
                    raise ValueError()
                df.at[i, col] = int(n)
            except (TypeError, ValueError, OverflowError):
                errors.append(f'{name}：{col}は0〜1000000の整数にしてください。')
        try:
            if float(df.at[i, '月間最小回数']) > float(df.at[i, '月間最大回数']):
                errors.append(f'{name}：月間最小回数は月間最大回数以下にしてください。')
        except (ValueError, TypeError):
            pass
        # 宿直上限・日直上限は空欄なら制限なし
        for col in GROUP_CAP_COLUMNS:
            raw = clean_text(row.get(col, ''))
            if raw.lower() in ('', 'nan', 'none', '<na>'):
                df.at[i, col] = ''
                continue
            try:
                n = float(raw)
                if not math.isfinite(n) or n < 0 or n != int(n) or n > 1000:
                    raise ValueError()
                df.at[i, col] = int(n)
            except (TypeError, ValueError, OverflowError):
                errors.append(f'{name}：{col}は0以上の整数にしてください（空欄なら制限なし）。')
        for col, ng in [(NG_COLUMN, True), (REQUEST_COLUMN, False)]:
            try:
                items = period_checked_items(row[col], P, ng)
                df.at[i, col] = period_items_text(items, P, ng)
            except ValueError as e:
                errors.append(f'{name}：{col.split("(")[0]} {e}')
        weekdays = clean_text(row['翌日PM duty']).replace('，', ',')
        if any(w.strip() not in WEEKDAYS_JA for w in weekdays.split(',') if w.strip()):
            errors.append(f'{name}：翌日PM dutyは「水,木」のように曜日を入力してください。')
    return df, errors


def period_validate_fixed_inputs(df, doctors, P, holidays=None):
    year, month = P.year, P.month
    """確定当直の入力チェック。holidays を渡すと、対象月の平日に日直が入っていないかも確認する。"""
    errors = []
    if df is None:
        return None, errors
    df = df.copy()
    if '日付' not in df:
        return df, ['確定当直CSVに「日付」の列がありません。']
    unknown = [c for c in df if c not in ['日付', '平日/休日'] + ALL_SHIFT_TYPES]
    for c in unknown:
        if df[c].map(clean_text).ne('').any():
            errors.append(f'確定当直：未対応の列「{c}」に入力があります。枠名を確認してください。')
    for i, row in df.iterrows():
        raw = clean_text(row['日付'])
        populated = any(clean_text(row.get(s, '')) not in ('', '-') for s in ALL_SHIFT_TYPES)
        if not raw and not populated:
            continue
        dt = parse_shift_date(raw, year, month)
        if dt is None:
            errors.append(f'確定当直 {i+1}行目：日付「{raw}」を確認してください。「11/1」または「2026/11/1」の形で入力します。')
            continue
        df.at[i, '日付'] = dt.isoformat()
        for s in ALL_SHIFT_TYPES:
            val = clean_text(row.get(s, ''))
            if val in ('', '-'):
                continue
            if (holidays is not None and s in DAY_SHIFTS and P.index(dt) is not None
                    and not period_is_holiday_date(P, dt, holidays)):
                errors.append(f'確定当直 {dt.month}/{dt.day} {s}：平日には日直がありません。この日に日直を設ける場合は、上のカレンダーで「休日にする」にチェックを入れてください。')
            names = list(dict.fromkeys(n.strip() for n in re.split('[、,]', val)))
            for name in names:
                if name not in doctors:
                    errors.append(f'確定当直 {dt.month}/{dt.day} {s}：医師名「{name}」が名簿と一致しません。')
            df.at[i, s] = '、'.join(names)
    return df, errors


def period_audit_schedule(result, staff, P, holidays, multi, past, future, next_month_special_holiday=False):
    year, month = P.year, P.month
    warnings = []
    for _, row in staff.iterrows():
        name = row['先生の名前']
        worked = []; counts = dict.fromkeys(ALL_SHIFT_TYPES, 0); hol_count = 0
        ng = period_parse_ng(row.get(NG_COLUMN, ''), P)
        pm_weekdays = pm_duty_weekdays(row.get('翌日PM duty', ''))
        for _, r in result.iterrows():
            dt = P.date(int(r['_idx']))
            for slot in ALL_SHIFT_TYPES:
                if name not in [n.strip() for n in re.split('[、,]', str(r[slot]))]:
                    continue
                worked.append(dt); counts[slot] += 1
                hol_count += int(r['平日/休日'] == '休日')
                kind = ng.get(P.index(dt), 'OK')
                night = slot in NIGHT_SHIFTS
                if kind == '全NG' or (kind == '宿NG' and night) or (kind == '日NG' and not night):
                    warnings.append(f'{name}：{dt.month}/{dt.day} {slot}はNG指定より確定指定を優先しました。')
                if night and period_pm_restricted(P, P.index(dt), pm_weekdays, holidays, next_month_special_holiday):
                    warnings.append(f'{name}：{dt.month}/{dt.day} {slot}は翌日PM dutyの制限より確定指定を優先しました。')
        total = sum(counts.values())
        if total < int(row['月間最小回数']):
            warnings.append(f'{name}：月間最小回数の目標{int(row["月間最小回数"])}回に対し、{total}回です。')
        group_items = [(col, sum(counts[s] for s in slots), int(row[col])) for col, slots in GROUP_CAP_COLUMNS.items()
                       if clean_text(row.get(col, '')) not in ('', 'nan')]
        for label, value, cap in [('月間最大回数', total, int(row['月間最大回数'])), ('休日最大回数', hol_count, int(row['休日最大回数']))] + group_items + [(s + '上限', counts[s], int(row[s + '上限'])) for s in ALL_SHIFT_TYPES]:
            if value > cap:
                warnings.append(f'{name}：確定指定に伴い{label}を超えています（設定{cap}回／結果{value}回）。')
        dates = sorted(set(worked) | set((past or {}).get(name, [])) | set((future or {}).get(name, [])))
        for a, b in zip(dates, dates[1:]):
            if P.index(a) is None and P.index(b) is None: continue
            gap = (b - a).days - 1
            if gap < int(row['最低空ける日数']):
                warnings.append(f'{name}：確定勤務同士の間隔が設定未満です（{a.month}/{a.day}→{b.month}/{b.day}、空き{gap}日／設定{int(row["最低空ける日数"])}日）。確定勤務を維持しました。')
    for _, r in result.iterrows():
        d = int(r['_idx'])
        for s in ALL_SHIFT_TYPES:
            names = [n.strip() for n in re.split('[、,]', str(r[s])) if n.strip() not in ('', '-') and '⚠️不足' not in n]
            need = multi.get((d, s), 1)
            if len(names) > need:
                warnings.append(f'{P.label(d)} {s}：設定{need}名に対して{len(names)}名を配置しています。')
    return warnings


def period_generate_shift(P, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None, locked_days=()):
    staff_df, errors = period_validate_staff_inputs(staff_df, P)
    fixed_df, fixed_errors = period_validate_fixed_inputs(fixed_df, staff_df.get('先生の名前', pd.Series(dtype=str)).tolist(), P, custom_holidays)
    errors += fixed_errors
    if staff_df.empty: errors.append('医師を1名以上入力してください。')
    for (d, s), count in multi_slots_dict.items():
        if not isinstance(d, int) or not 1 <= d <= P.n or s not in ALL_SHIFT_TYPES or not isinstance(count, int) or not 2 <= count <= 10:
            errors.append(f'増員設定「{d} {s} {count}名」を確認してください。')
        elif s in DAY_SHIFTS:
            if not period_is_holiday(P, d, custom_holidays):
                errors.append(f'{P.label(d)}の日直増員：先にカレンダーでその日を「休日にする」にしてください。')
    if errors: return None, False, errors, None, None
    result, success, warnings, past, future = _period_generate_core(P, staff_df, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority, locked_days)
    if result is not None:
        result_idx = result.assign(_idx=range(1, len(result) + 1))
        warnings += period_audit_schedule(result_idx, staff_df, P, custom_holidays, multi_slots_dict, past, future, next_month_special_holiday)
        success = not any('⚠️不足' in str(v) for s in ALL_SHIFT_TYPES for v in result[s])
    return result, success, list(dict.fromkeys(warnings)), past, future


def period_build_duty_matrix(df_result, staff_df, P, custom_holidays, next_month_special_holiday=False, comments=None):
    year, month = P.year, P.month
    """
    医師ごと・日ごとのマスの内容を作る。
    優先順：決まった当直 ＞ ×（当直NG）＞ PM（翌日PM duty）。
    """
    num_days = len(df_result)
    days = []
    for d in range(1, num_days + 1):
        dt = P.date(d)
        hol = str(df_result.iloc[d - 1]['平日/休日']) == '休日'
        kind = 'holiday' if (hol and (dt.weekday() == 6 or jpholiday.is_holiday(dt) or d in custom_holidays)) else ('saturday' if dt.weekday() == 5 else 'weekday')
        # 当直が決まっていない枠（⚠️不足）
        shortage = []
        for s in RESULT_COLUMNS[2:]:
            m = re.search(r'⚠️不足\((\d+)名\)', str(df_result.iloc[d - 1][s]))
            if m:
                shortage.append({'slot': s, 'count': int(m.group(1))})
        days.append({'day': d, 'weekday': WEEKDAYS_JA[dt.weekday()], 'holiday': hol, 'kind': kind, 'shortage': shortage,
                     'ng_count': 0, 'label': P.label(d)})   # その日にNG（翌日PM dutyで宿直を外す先生を含む）の先生の数

    # 備考は「備考」で始まる列から読む（列名の書き方の違いに対応）
    note_columns = [c for c in staff_df.columns if str(c).strip().startswith("備考")]
    rows = []
    for _, staff in staff_df.iterrows():
        name = str(staff['先生の名前'])
        ng = period_parse_ng(staff.get(NG_COLUMN, ''), P)
        pm_days = pm_duty_weekdays(staff.get('翌日PM duty', ''))
        wish_days, wish_specific = period_parse_requests(staff.get(REQUEST_COLUMN, ''), P)
        cells = []
        weekday_count = holiday_count = ward_count = outpatient_count = night_count = day_count = 0
        for info in days:
            d = info['day']
            r = df_result.iloc[d - 1]
            # 日ごとのNG人数：その日にNG（平日は宿NG・全NG、休日は日NGも）か、翌日PM dutyで宿直を外す先生を数える
            if effective_ng(ng.get(d, 'OK'), info['holiday']) != 'OK' or period_pm_restricted(P, d, pm_days, custom_holidays, next_month_special_holiday):
                info['ng_count'] += 1
            assigned = [s for s in RESULT_COLUMNS[2:] if name in [x.strip() for x in re.split(r'[、,]', str(r[s]))]]
            if assigned:
                if info['holiday']:
                    holiday_count += len(assigned)
                else:
                    weekday_count += len(assigned)
                night_count += sum(1 for s in assigned if s in NIGHT_SHIFTS)
                day_count += sum(1 for s in assigned if s in DAY_SHIFTS)
                ward_count += sum(1 for s in assigned if s in MATRIX_WARD_SHIFTS)
                outpatient_count += sum(1 for s in assigned if s in MATRIX_OUTPATIENT_SHIFTS)
                # 希望が通った日（日付だけの希望はどの枠でも、枠指定の希望はその枠に入ったとき）
                wish = d in wish_days or any((d, s) in wish_specific for s in assigned)
                text = '/'.join(assigned)
                cells.append({'type': 'shift', 'text': text, 'shift': assigned[0], 'wish': wish})
                continue
            kind = ng.get(d)
            pm = period_pm_restricted(P, d, pm_days, custom_holidays, next_month_special_holiday)
            # 翌日PM dutyの日にNGも付いている場合は、斜線の上に×を出す（斜線だけ＝NGは付いていない）
            if info['holiday'] and (pm or kind in ('全NG', '宿NG', '日NG')):
                # 休日は、上半分＝日直、下半分＝宿直に分けて、NGの側に×を出す（全NGは上下とも×）。
                # 翌日PM dutyで宿直を外す日は、下半分を斜線にする
                top = '×' if kind in ('全NG', '日NG') else ''
                bottom = '×' if kind in ('全NG', '宿NG') else ''
                cells.append({'type': 'pmhalf' if pm else 'half', 'top': top, 'bottom': bottom, 'text': '×' if (top or bottom) else ''})
            elif kind in ('全NG', '宿NG'):
                cells.append({'type': 'ng', 'text': '×', 'pm': pm})
            elif pm:
                cells.append({'type': 'pm', 'text': ''})
            else:
                cells.append({'type': 'empty', 'text': ''})
        rows.append({
            'name': name,
            'pm_duty': '・'.join(WEEKDAYS_JA[i] for i in pm_days),
            'note': "\n".join(v for v in (clean_text(staff.get(c, "")) for c in note_columns) if v and v.lower() not in ("nan", "none")),
            'comment': (comments or {}).get(name, ("", ""))[0],
            'comment_at': (comments or {}).get(name, ("", ""))[1],
            'cells': cells,
            'night': night_count, 'day': day_count, 'ward': ward_count, 'outpatient': outpatient_count,
            'weekday': weekday_count, 'holiday': holiday_count, 'total': weekday_count + holiday_count,
        })
    return days, rows



# ==========================================
# ページ設定
# ==========================================
st.set_page_config(page_title="当直作成アプリ", layout="wide")

# 先生用の専用URL（…/?key=XXXX）で開かれたときは、その先生のNG日入力画面だけを表示する
try:
    _remote_key = st.query_params.get("key")
except AttributeError:   # 古いStreamlit
    _remote_key = (st.experimental_get_query_params().get("key") or [None])[0]
if _remote_key:
    doctor_entry_page(str(_remote_key))
    st.stop()

# 期間指定版（年末年始など）：…/?mode=period で開かれたときは、そのページだけを表示する
try:
    _mode = st.query_params.get("mode")
except AttributeError:
    _mode = None
if _mode == "period":
    _pp = pathlib.Path(__file__).resolve().parent / "period_page.py"
    if _pp.exists():
        exec(compile(_pp.read_text(encoding="utf-8"), str(_pp), "exec"), globals())
    else:
        st.error("期間指定版のファイル（period_page.py）が見つかりません。GitHubに period_page.py を入れてください。")
        st.markdown("<a href='./' target='_self'>← 通常版に戻る</a>", unsafe_allow_html=True)
    st.stop()

st.title("当直作成アプリ")
st.caption("年月・勤務条件・NG日を入力して、当直表の案を作成します。")
with st.expander("📖 使い方（初めての方・毎月の流れ・結果の見方）", expanded=False):
    tab_intro, tab_flow, tab_ng, tab_result, tab_trouble = st.tabs(["はじめに", "毎月の流れ", "NG日の集め方", "結果の見方・Excel", "困ったとき"])
    with tab_intro:
        st.markdown("""
**このアプリでできること**

先生方の条件（回数・間隔・希望日・NG日など）をもとに、毎月の当直表の案を自動で作ります。
先生方には専用のURL（QRコード）から、スマートフォンでNG日を入力してもらえます。

**画面のつくり**

- **いちばん上の紺色の帯**：「💾 作業を途中保存する」と「📂 前回の作業の続きから始める」があります。スクロールしても常に表示されます。
- **いちばん下の紺色の帯**：「🚀 当直案を作成」ボタンと、「医師条件」「NG日」「結果」へ移動するボタン、作成結果（未決定の枠の数）があります。
- **真ん中**：上から順に、紺色の見出しの帯で4つの項目に分かれています。上から順に進めれば、1か月分の作業が終わります。
    1. 📅 年月・休日・必要人数の設定
    2. 👩‍⚕️ 医師条件の読み込み・入力
    3. 🚫 先生ごとのNG日設定
    4. 🎯 当直案の作成・確認

**年末年始など、月をまたいで作るとき**

年月の設定の下にある「📆 期間指定版」から、今月1日〜翌月の指定日までを通しで作る別のページに移れます。先生方の専用URLは同じものを使えます。

**管理用パスワード**

途中保存・続きからの読み込み・前回の医師条件の引き継ぎ・専用URLの管理（受付・取り込みなど）には、管理用パスワードが必要です。
一度入れると、画面を閉じるまでほかの場所でも入力不要です。

**大事なこと**

- 画面に入力した内容は、しばらく使わないと消えます。作業の区切りごとに、上の帯の「💾 作業を途中保存する」を押してください。
- 作った当直案は途中保存に入りません。気に入った案は、Excelでダウンロードしておいてください。
        """)
    with tab_flow:
        st.markdown("""
**毎月の作業は、画面の上から順に進めれば完了します。**

1. **年月を選ぶ**：いちばん上で、作る年月を選びます。
   同じ月の作業の続きなら、上の帯の「📂 前回の作業の続きから始める」で、NG日なども含めてそのまま戻せます（この場合、2〜5は確認だけで済みます）。
2. **休日を設定**：カレンダーで、次の2つを必要な日だけチェックします。土日祝は自動で休日になります。
    - **休日にする**：病院独自の休みなど、平日でも土日祝と同じ扱いにする日。日直の枠ができます。
    - **別途決定**：年末年始など、別の方法で当番を決める日。このアプリではその日の当直を決めず、先生もNGを入力できなくなります。
3. **人数を増やす枠（任意）**：1つの枠を2名以上にする日があれば、行を足して設定します。
4. **確定当直（任意）**：「📅 先月今月来月の確定当直」の帯を押して開き、すでに決まっている当直や、先月末・来月初めの当直を入れます。勤務の間隔の計算に使われます。
5. **医師条件**：新しい月なら、医師条件の欄の「📋 前回の医師条件を引き継ぐ」を押します。途中保存してある医師条件（回数・間隔・翌日PM duty・備考など）が入り、NG日・希望日・先生のコメントは空になります。
   初めての場合は、ひな形（CSV）をダウンロードして先生の名前や回数の上限などを入れ、アップロードします。表の中で直接直すこともできます。希望日もここに入れます。
   入力したら、すぐ下の「⚖️ 当直の枠は足りそう？」で、全体と枠の種類ごとに足りているか（赤い棒がないか）を確認します。
6. **NG日を集める**：先生方に専用URLから入力してもらい、取り込みます（「NG日の集め方」のタブを参照）。
7. **不足時の優先順位（任意）**：人数が足りない月に、どの枠から埋めるかを決めます。
8. **当直案を作る**：下の帯の「🚀 当直案を作成」を押します（計算に最大1分少々かかります）。押すたびに違うパターンの案ができるので、気に入るまで押し直して構いません。
   未決定の枠が出たら、条件を直して作り直します。下の帯の移動ボタンで、条件と結果を行き来できます。
9. **確認してダウンロード**：「医師別の当直一覧」をExcelでダウンロードします（「結果の見方・Excel」のタブを参照）。
10. **途中保存**：最後に上の帯の「💾 作業を途中保存する」を押しておくと、翌月は「📋 前回の医師条件を引き継ぐ」で、この医師条件から始められます。

途中保存には、年月・休日・別途決定の日・増員設定・確定当直・医師条件（NG日を含む）・取り込んだ先生のコメント・不足時の優先順位・アラートの基準が入ります。
        """)
    with tab_ng:
        st.markdown("""
**方法A：先生方に専用URLから入力してもらう（おすすめ）**

「🚫 先生ごとのNG日設定」の中の「📨 先生方からのNG日入力（専用URL）」を開き、管理用パスワードを入れて使います。

1. **受付を開始**：先に、いちばん上のカレンダーで「休日にする」「別途決定」を済ませておきます。締め切り日を書いて「受付を開始する」を押します。
   受付中にカレンダーを変えた場合も、パスワードを入れた状態なら自動で先生側に反映されます。
2. **URLを配る（初回だけ）**：「配布用の案内書類（QRコード付き）」をダウンロードして印刷し、点線で三つ折りにして封筒で各先生に渡します。URLは毎月同じなので、翌月からは「入力が始まりました」と伝えるだけです。
3. **先生が入力**：先生はQRコードを読み取り、入れない日の「日直」「宿直」のボタンを押してNGにし、「NG日とコメントを保存する」を押します。「全日NGにする」「すべてOKに戻す」のボタンもあります。締め切りまでは何度でも直せて、最後に保存した内容が採用されます。
4. **提出状況を確認**：管理画面の表で、提出済み・未提出・コメントを確認できます。「🔄 最新の状態に更新」で最新になります。
5. **締め切る**：「🔒 締め切る」を押すと、先生側は編集できなくなります。
6. **取り込む**：「取り込むと変わる先生」の表で内容とコメントを確認し、取り込む先生を選んで「⬇️ 選んだ先生のNG日を取り込む」を押します。その先生のNG日は提出内容に置き換わり、コメントも当直一覧のメモ欄に表示されるようになります。

**やり直したいとき**

- **取り込みを取り消す**：「↩️ 取り込みを取り消す」で、取り込む前のNG日に戻せます（先生ごとのNG日の欄にもボタンがあります）。取り込んだあとにアプリを開き直したり、途中保存から読み込んだりした場合は取り消せないので、カレンダーで直してください。
- **先生の提出をリセットする**：入力をやり直してもらう場合は、「先生の提出をリセットする」で先生を選んでリセットします。未提出に戻り、先生の画面も空になります（これまでの提出は履歴に残ります）。
- **URLを作り直す**：案内書類をなくした・URLが漏れた場合に使います。古いURLは使えなくなります。

**方法B：管理者が直接入力する**

「NG日を入力する医師」で先生を選び、カレンダーの「日直」「宿直」のボタンを押してNGにし、「NG日を反映する」を押します。紙やメールで集めた場合はこちらを使います。
反映していない変更があると、ボタンがオレンジに光ります。そのまま別の先生に切り替えようとすると注意が出ます。

| NGの種類 | 意味 |
| --- | --- |
| 宿NG | 宿直（夜）に入れない。「宿直」のボタンだけ押した状態 |
| 日NG | 日直（日中）に入れない。休日だけ。「日直」のボタンだけ押した状態 |
| 全NG | 日直・宿直とも入れない。休日だけ。両方のボタンを押した状態 |
| ⚠ | 翌日PMにdutyがあるため、原則その日の宿直を外す日 |
| 灰色の「別途決定」 | 別の方法で当番を決める日。NGは入力できません |
        """)
    with tab_result:
        st.markdown("""
**🗓️ 医師別の当直一覧**（先生ごとに1行）

- 枠名のマス：決まった当直（色は枠の種類）
- **茶色の太い枠**：希望日が、希望どおりに通った当直
- **薄い赤の列**：土日祝と「休日にする」にした日
- **×**：当直NGの日（平日は宿直NG）
- **休日のマス**：上半分＝日直、下半分＝宿直です。上に×＝日直NG、下に×＝宿直NG、上下とも×＝終日NG
- **黄色の斜線**：翌日PM dutyのため宿直を外す日。×があれば、NGも付いている日です
- **下半分だけ黄色の斜線**：休日の翌日PM dutyの日。宿直は外しますが、日直には入れます（×の意味は休日のマスと同じ）
- **NG人数の行**（見出しのすぐ下）：その日にNGの先生の数。翌日PM dutyで宿直を外す先生も数えます。多い日ほど濃いオレンジです
- **赤線の下の「⚠️ 未決定」**：人数が足りず、担当が決まっていない枠
- **右端の回数**：病棟・外来／宿直・日直／平日・休日／合計。回数が多い先生は**赤い枠**で表示されます。該当する先生は、作成ボタンの下の「計算状況・条件の注意事項と対処方法」にも一覧で出ます。基準は表のすぐ上の「⚠ アラートの基準」で変えられます（初期設定は宿直4回以上・休日3回以上。0にした項目はアラートなし）。
- **メモ欄**：💬は先生からのコメント（取り込んだ先生だけ）、📝は医師条件の備考です。カーソルを合わせると全文が出ます。全員分は「💬 先生方のコメント・備考の一覧」でまとめて見られます。
- 行の見間違いを防ぐため、先生の行は1行おきに薄いグレーになっています。
- マスにカーソルを合わせると、その先生の行と、その日の列が黄色く色付きます。

**📥 Excelでダウンロード（緑のボタン）**

印刷用の表（A3横・1枚）がダウンロードされます。Excelで最終調整もできます。

- 画面の表と同じ見た目ですが、翌日PM dutyの日は斜線ではなく黄色の塗り（休日は下半分だけ黄色）で表します。
- 右の回数は計算式になっているので、マスに「A宿直」「外来日直」などの枠名を書き足したり消したりすると、自動で変わります。赤い枠も自動で付きます。
- 「⚠️ 未決定」の行の枠を、表の中のマスへドラッグして移せます。移すと、未決定の「○枠」と、移した先の先生の回数が自動で変わります（移した元のマスは色のない白になります）。
- メモ欄にはメモの全文が入っています。マスを選ぶと上の入力欄で全文が読めます。「メモ全文」シートにも一覧があります。

**📅 作成した当直案**（日付ごとの表）

名前にカーソルを合わせると、同じ先生の当直が強調されます。
「🔍 特定の医師の当直を色別でハイライト」を開くと、選んだ先生を色分けできます。CSVでもダウンロードできます。
        """)
    with tab_trouble:
        st.markdown("""
- **「未決定」の枠が出る**：作成ボタンの下の「計算状況・条件の注意事項と対処方法」を見てください。医師条件の下の「⚖️ 当直の枠は足りそう？」で、赤い棒（足りない枠）がないかも確認します。NG日・回数の上限・勤務間隔の条件が厳しすぎる可能性があります。月間最大回数を増やす、勤務間隔を短くするなどして、下の帯の「🚀 当直案を作成」で作り直してください。少しだけなら、Excelで手で埋めることもできます。
- **アプリを開くと入力した内容が消えている**：しばらく使わないと画面がリセットされます。上の帯の「📂 前回の作業の続きから始める」で読み込んでください。先生方の提出内容と専用URLはスプレッドシートに残っているので消えません。
- **英語の画面（Zzzz…）が出る**：アプリが休止しています。青いボタン「Yes, get this app back up!」を押して1分ほど待つと開きます。
- **先生のNG日が当直案に入っていない**：管理画面で「⬇️ 選んだ先生のNG日を取り込む」を押したか確認してください。
- **先生側の休日・別途決定の表示が違う**：管理用パスワードを入れて、「締め切り日・休日設定を更新」を押してください。
- **先生から「URLが無効」と言われた**：URLを作り直していないか確認してください。案内書類をなくした場合は、管理画面の「URLを作り直す」で新しい案内書類を印刷します。
- **アプリが起動しない・エラーが出る**：表示されたエラーの内容を控えて、管理担当者に連絡してください。
        """)

# 見出しの装飾。Streamlitの内部名に依存するため、将来のバージョンで効かなくなることがあるが、
# その場合も標準の見た目に戻るだけで、動作やレイアウトは崩れない。
st.markdown("""
<style>
[data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h1,
.main [data-testid="stHeadingWithActionElements"] h1 {
    font-size: 1.9rem !important; line-height: 1.4 !important; font-weight: 750 !important; padding: 0 0 .6rem !important;
}
[data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h2,
.main [data-testid="stHeadingWithActionElements"] h2 {
    font-size: 1.3rem !important; line-height: 1.5 !important; font-weight: 700 !important;
    background: var(--secondary-background-color, #f0f3f7);
    border-left: 5px solid #47749b; border-radius: 0 8px 8px 0;
    padding: .85rem 1rem !important; margin: .4rem 0 .35rem !important;
}
[data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h3,
.main [data-testid="stHeadingWithActionElements"] h3 {
    font-size: 1.08rem !important; line-height: 1.55 !important; font-weight: 700 !important;
    padding: .45rem 0 .6rem !important; border-bottom: 1px solid rgba(128, 144, 160, .3);
    margin: .2rem 0 .15rem !important;
}
[data-testid="stMain"] [data-testid="stMarkdownContainer"] hr,
.main [data-testid="stMarkdownContainer"] hr {
    margin: 1.4rem 0 !important; border: 0; border-top: 1px solid rgba(128, 144, 160, .35);
}
[data-testid="stMetricLabel"] p {font-size: .9rem !important; line-height: 1.5 !important;}
[data-testid="stMetricValue"] {font-size: 1.8rem !important; line-height: 1.25 !important;}
@media (max-width:600px) {
    [data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h1 {font-size:1.6rem !important;}
    [data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h2 {font-size:1.15rem !important; padding:.7rem .75rem !important;}
    [data-testid="stMain"] [data-testid="stHeadingWithActionElements"] h3 {font-size:1rem !important;}
}
</style>
""", unsafe_allow_html=True)

# 大きな区切り（セクション）の見出し。紺色の帯で、どこからどこまでが1つの項目かを分かりやすくする
st.markdown("""
<style>
.sec-banner{display:flex;align-items:center;gap:14px;margin:2.6rem 0 .9rem;padding:14px 18px;border-radius:12px;
  background:linear-gradient(90deg,#2f4a6d 0%,#47749b 100%);color:#fff;box-shadow:0 2px 6px rgba(31,52,82,.18)}
.sec-banner .ic{font-size:1.6rem;line-height:1}
.sec-banner .t{font-size:1.35rem;font-weight:800;letter-spacing:.02em;line-height:1.3}
.sec-banner .s{font-size:.88rem;opacity:.9;margin-top:2px}
.sec-banner .tag{margin-left:auto;font-size:.8rem;font-weight:800;padding:3px 12px;border-radius:999px;background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.45);white-space:nowrap}
.sec-banner.first{margin-top:.6rem}
/* 画面上部の帯（ボタンの並ぶ所）：下の画面との境目が分かるよう、細い線とごく薄い影を付ける。
   すりガラス風にして、下を流れる内容が透けて重ならないようにする */
[data-testid="stHeader"]{background:linear-gradient(90deg,#24405f 0%,#35597f 100%) !important;
  border-bottom:3px solid #f2b84b;box-shadow:0 3px 12px rgba(16,24,40,.22)}
/* 紺色の帯の上でも見えるよう、右上のStreamlitのアイコンや「実行中」の表示を白にする */
[data-testid="stHeader"] button,[data-testid="stHeader"] a,[data-testid="stHeader"] [data-testid="stToolbar"],
[data-testid="stHeader"] [data-testid="stStatusWidget"],[data-testid="stHeader"] [data-testid="stStatusWidget"] *{color:#fff !important}
[data-testid="stHeader"] svg{fill:currentColor}
[data-testid="stDecoration"]{display:none}
/* 下の帯のリンクで移動したとき、画面上部の帯（ボタンの並ぶ所）に見出しが隠れないよう、少し上に余白を取って止める */
#sec-staff,#sec-ng,#sec-create,#sec-result{scroll-margin-top:5.5rem}
.sec-banner[id]{scroll-margin-top:5.5rem}
/* 折りたたみ式の項目（確定当直）も、同じ帯の見た目にする */
.st-key-fixed_section{margin:2.6rem 0 .4rem}
.st-key-fixed_section [data-testid="stExpander"] details{border:0;border-radius:12px;overflow:hidden;box-shadow:0 2px 6px rgba(31,52,82,.18)}
.st-key-fixed_section [data-testid="stExpander"] details > summary{background:linear-gradient(90deg,#2f4a6d 0%,#47749b 100%);color:#fff;padding:14px 18px}
.st-key-fixed_section [data-testid="stExpander"] details > summary p{color:#fff;font-size:1.35rem;font-weight:800}
.st-key-fixed_section [data-testid="stExpander"] details > summary svg{color:#fff;fill:#fff}
.st-key-fixed_section [data-testid="stExpanderDetails"]{padding-top:1rem}
@media (max-width:600px){.sec-banner{padding:11px 13px;gap:10px}.sec-banner .t,.st-key-fixed_section summary p{font-size:1.12rem !important}.sec-banner .s{display:none}}
</style>""", unsafe_allow_html=True)


def section_banner(icon, title, sub="", tag="", first=False, anchor=""):
    tag_html = f'<span class="tag">{tag}</span>' if tag else ""
    sub_html = f'<div class="s">{sub}</div>' if sub else ""
    id_html = f' id="{anchor}"' if anchor else ""
    st.markdown(f'<div{id_html} class="sec-banner{" first" if first else ""}"><span class="ic">{icon}</span>'
                f'<div><div class="t">{title}</div>{sub_html}</div>{tag_html}</div>', unsafe_allow_html=True)


# ==========================================
# 1. 上部ダッシュボード：年月と休日の設定
# ==========================================
section_banner("📅", "年月・休日・必要人数の設定", "作る月・休日と別途決定の日・1枠の人数を決めます", "必須", first=True)
st.info("作成する年月を選んでください。土日祝日は自動で休日扱いになります。下のカレンダーで、平日でも休日扱いにする日は「休日にする」、別の方法で当番を決める日は「別途決定」にチェックします。")
# 年末年始など、翌月にまたがって通しで作りたいときの入口（同じアプリの中で ?mode=period を開く）
try:
    _period_exists = (pathlib.Path(__file__).resolve().parent / "period_page.py").exists()
except Exception:
    _period_exists = False
if _period_exists:
    st.markdown("<a href='./?mode=period' target='_self' style='font-weight:700'>📆 年末年始など、翌月の数日まで含めて通しで作る場合はこちら（期間指定版）</a>",
                unsafe_allow_html=True)

today = datetime.date.today()
if today.month == 12:
    default_year = today.year + 1
    default_month = 1
else:
    default_year = today.year
    default_month = today.month + 1

# 画面右上の帯に「前回の作業の続きから始める」「作業を途中保存する」を置く
work_save_floating()

# --- 画面の下に固定する帯（フッター）：「当直案を作成」ボタンと結果、移動リンク ---
# 条件を直す → 作成し直す、を行ったり来たりせずに繰り返せるように、どこまでスクロールしていても押せるようにする。
# 上の帯は「保存」、下の帯は「作成」と役割を分ける。
st.markdown("""
<style>
.st-key-create_bar{position:fixed;left:0;right:0;bottom:0;z-index:1000001;width:auto !important;height:64px !important;min-height:0 !important;
  display:flex !important;flex-direction:row !important;align-items:center !important;gap:14px !important;
  padding:0 9rem 0 1rem;background:linear-gradient(90deg,#24405f 0%,#35597f 100%);
  border-top:3px solid #f2b84b;box-shadow:0 -3px 12px rgba(16,24,40,.22);overflow:visible !important}
.st-key-create_bar > div{width:auto !important;flex:0 0 auto !important;margin:0 !important}
.st-key-create_bar [data-testid="stElementContainer"],.st-key-create_bar [data-testid="stMarkdownContainer"],
.st-key-create_bar [data-testid="stMarkdown"]{margin:0 !important;width:auto !important}
/* 並び：いちばん左に作成ボタン、その右に移動リンク、右端に結果 */
.st-key-create_bar > div:has(button){order:1}
.st-key-create_bar > div:has(.cb-box){order:2;flex:1 1 auto !important}
.st-key-create_bar button{min-height:2.7rem;padding:0 1.5rem 0 1.15rem;font-size:1rem;font-weight:800;white-space:nowrap}
/* 絵文字の前に入る余白をなくし、ボタンの左右の余白を見た目でそろえる */
.st-key-create_bar button p{margin:0;letter-spacing:.01em}
.st-key-create_bar p{margin:0}
.cb-box{display:flex;flex-direction:row;align-items:center;gap:14px;width:100%}
.cb-res{margin-left:auto;font-size:.92rem;font-weight:800;line-height:1.4;white-space:nowrap;padding:4px 12px;border-radius:999px;background:rgba(255,255,255,.12)}
.cb-res.ok{color:#8ff0b6}.cb-res.bad{color:#ffd08a}.cb-res.wait{color:#e3e9f1}
.cb-nav{display:flex;gap:6px;font-size:.82rem;order:-1}
.cb-nav a{text-decoration:none;color:#fff !important;background:rgba(255,255,255,.14);border:1px solid rgba(255,255,255,.28);border-radius:6px;padding:4px 10px;white-space:nowrap}
.cb-nav a:hover{background:rgba(255,255,255,.26)}
/* 作成ボタンは紺の帯の上で目立つよう、白い縁取りを付ける */
.st-key-create_bar button{box-shadow:0 0 0 2px rgba(255,255,255,.85),0 2px 8px rgba(0,0,0,.25)}
/* 下の帯に最後の内容が隠れないよう、ページの下に余白を取る */
[data-testid="stMainBlockContainer"],.block-container{padding-bottom:6.5rem !important}
@media (max-width:600px){
  /* スマートフォン：右下にはStreamlitのマークが出るので、右側を空ける */
  .st-key-create_bar{height:58px !important;padding:0 7.5rem 0 .5rem;gap:8px !important}
  .st-key-create_bar button{padding:0 .9rem;font-size:.9rem}
  .cb-nav{display:none}.cb-res{margin-left:0;font-size:.75rem;white-space:normal}
}
</style>""", unsafe_allow_html=True)
create_bar_request = False


def create_bar_show(result_html=""):
    """下の帯の、移動リンクと結果の1行をまとめて表示する（1つの部品にして、表示のずれを防ぐ）。"""
    res_anchor = "#sec-result" if "generated_df" in st.session_state else "#sec-create"
    CREATE_BAR_SLOT.markdown(
        f"<div class='cb-box'><div class='cb-nav'><a href='#sec-staff' target='_self'>医師条件</a>"
        f"<a href='#sec-ng' target='_self'>NG日</a><a href='{res_anchor}' target='_self'>結果</a></div>{result_html}</div>",
        unsafe_allow_html=True)


with st.container(key="create_bar"):
    if st.button("🚀 当直案を作成", key="create_bar_button", type="primary"):
        create_bar_request = True
    CREATE_BAR_SLOT = st.empty()
    _last = st.session_state.get("create_bar_summary")
    if create_bar_request:
        create_bar_show("<div class='cb-res wait'>計算しています…（最大75秒）</div>")
    elif st.session_state.get("result_needs_refresh"):
        create_bar_show("<div class='cb-res wait'>条件が変わりました。もう一度作成してください。</div>")
    elif _last:
        create_bar_show(_last)
    else:
        create_bar_show()

if "sel_year" not in st.session_state:
    st.session_state["sel_year"] = default_year
if "sel_month" not in st.session_state:
    st.session_state["sel_month"] = default_month
col_y, col_m = st.columns(2)
year = int(col_y.number_input("年", min_value=2000, step=1, key="sel_year"))
month = int(col_m.number_input("月", min_value=1, max_value=12, step=1, key="sel_month"))
restored_state = st.session_state.get("restored_state")

st.divider()

st.subheader(f"📅 休日と別途決定の日の設定 - {month}月")
st.caption("**休日にする**：病院独自の休みなど、平日でも土日祝と同じ扱いにする日です。日直の枠ができ、日直・宿直とも休日回数に数えます。")
st.caption("**別途決定**：年末年始など、別の方法で当番を決める日です。このアプリでは当直を決めず、先生の専用URLの画面でも灰色になってNGを入力できません。確定当直に入れた分だけは当直案に表示されます。")

cal_matrix = calendar.monthcalendar(year, month)
custom_holidays = []
locked_days = []   # 先生がNGを入力できない日（年末年始など）

# 特別休日カレンダーの見た目。key を付けたコンテナ（st-key-〜）を起点に指定しているので、
# 他の部品には影響しない。
st.markdown("""
<style>
.st-key-special_holiday_calendar [data-testid="stHorizontalBlock"] {
 display:grid !important;
 grid-template-columns:repeat(7,minmax(0,1fr)) !important;
 gap:6px !important;
 width:100% !important;
}
.st-key-special_holiday_calendar [data-testid="stHorizontalBlock"] > :is([data-testid="stColumn"],[data-testid="column"]) {
 width:100% !important; min-width:0 !important;
 flex:none !important;
}
.st-key-special_holiday_calendar [class*="st-key-special_day_"] {
 height:138px !important; min-height:138px !important;
 border:1px solid #d6dce5; border-radius:7px;
 padding:8px 3px !important; gap:6px !important;
 box-sizing:border-box;
}
.st-key-special_holiday_calendar [data-testid="stElementContainer"]:has([data-testid="stCheckbox"]),
.st-key-special_holiday_calendar .element-container:has([data-testid="stCheckbox"]) {
 width:100% !important; align-self:stretch !important;
}
.st-key-special_holiday_calendar [data-testid="stCheckbox"] {
 width:100% !important; display:flex !important; justify-content:center !important;
}
.st-key-special_holiday_calendar [data-testid="stCheckbox"] label {
 display:flex !important; justify-content:center !important; gap:3px; width:fit-content !important; max-width:100%; margin-left:auto !important; margin-right:auto !important;
 min-width:0;
}
.st-key-special_holiday_calendar [data-testid="stCheckbox"] label p {
 font-size:12px; line-height:1.2; overflow-wrap:anywhere;
}
.st-key-special_holiday_calendar [data-testid="stCheckbox"] label > span {
 flex-shrink:0;
}
@media(max-width:600px) {
 .st-key-special_holiday_calendar [data-testid="stHorizontalBlock"] {gap:3px !important}
 .st-key-special_holiday_calendar [class*="st-key-special_day_"] {
  height:128px !important; min-height:128px !important; padding:6px 1px !important;
 }
 .st-key-special_holiday_calendar [data-testid="stCheckbox"] label p {font-size:10px}
 .st-key-special_holiday_calendar [data-testid="stCheckbox"] label {gap:1px}
}
</style>
""", unsafe_allow_html=True)

with st.container(key="special_holiday_calendar"):
    # 曜日のヘッダー行
    cols = st.columns(7)
    for i, w in enumerate(WEEKDAYS_JA):
        color = "#ff4b4b" if i == 6 else ("#1e90ff" if i == 5 else "inherit")
        cols[i].markdown(
            f"<div style='color: {color}; font-weight: bold; text-align: center; padding: 4px 0;'>{w}</div>",
            unsafe_allow_html=True
        )

    # 日付とチェックボックス（1日ごとの格子セルは平日・休日・空欄すべて同じ高さ）
    for week in cal_matrix:
        cols = st.columns(7)
        for i, day in enumerate(week):
            with cols[i]:
                with st.container(key=f"special_day_{year}_{month}_{week[0]}_{i}"):
                    if day == 0:
                        st.markdown("<div style='height: 72px;'></div>", unsafe_allow_html=True)
                        continue

                    date_obj = datetime.date(year, month, day)
                    is_weekend_or_hol = date_obj.weekday() >= 5 or jpholiday.is_holiday(date_obj)
                    day_color = "#ff4b4b" if is_weekend_or_hol else "inherit"

                    st.markdown(
                        f"<div style='width:100%;text-align:center;color:{day_color};font-weight:600;"
                        f"font-size:0.95rem;line-height:1.35;margin:0 0 8px 0;padding:0;'>{day}日</div>",
                        unsafe_allow_html=True
                    )

                    if is_weekend_or_hol:
                        st.markdown(
                            "<div style='width:100%;height:2.35rem;display:flex;align-items:center;"
                            "justify-content:center;color:#ff4b4b;font-size:0.85rem;line-height:1.2;"
                            "margin:0;padding:0;'>休</div>",
                            unsafe_allow_html=True
                        )
                    else:
                        if st.checkbox("休日にする", key=f"hol_{year}_{month}_{day}"):
                            custom_holidays.append(day)
                    # 年末年始など、別の方法で当番を決める日（先生側でNGを入力できなくする）
                    if st.checkbox("別途決定", key=f"lock_{year}_{month}_{day}",
                                   help="年末年始など、別の方法で当番を決める日です。この日は当直案を作らず、先生の専用URLの画面でもNGを入力できなくなります。"):
                        locked_days.append(day)

next_month_first = datetime.date(year, month, calendar.monthrange(year, month)[1]) + datetime.timedelta(days=1)
next_month_first_label = f"{next_month_first.year}/{next_month_first.month}/1（{WEEKDAYS_JA[next_month_first.weekday()]}）"
month_last = next_month_first - datetime.timedelta(days=1)
month_last_label = f"{month_last.month}/{month_last.day}（{WEEKDAYS_JA[month_last.weekday()]}）"
# 翌月1日が休みかどうか：月末の宿直に「翌日PM duty」の制限をかけるかどうかの判断に使う
st.markdown(f"**翌月1日の扱い**（月末 {month_last_label} の宿直に、翌日PM dutyの制限をかけるかどうか）")
if next_month_first.weekday() >= 5 or jpholiday.is_holiday(next_month_first):
    # 土日祝は自動で休日扱いになるため、チェックボックスは表示しない
    _kind = "祝日" if jpholiday.is_holiday(next_month_first) else "土日"
    st.caption(f"翌月1日 {next_month_first_label} は{_kind}なので、自動で休日として扱います（チェックは不要です）。"
               f"翌日が休みなので、月末 {month_last_label} の宿直は「翌日PM duty」の制限を受けません。")
    next_month_special_holiday = False
else:
    next_month_special_holiday = st.checkbox(
        f"翌月1日 {next_month_first_label} は病院独自の休み",
        key=f"next_month_special_holiday_{year}_{month}",
        help="平日でも病院が休みの場合だけチェックしてください。"
    )
    st.caption(f"翌月1日 {next_month_first_label} は平日です。病院が休みの場合だけチェックしてください。"
               f"チェックすると、月末 {month_last_label} の宿直は「翌日PM duty」の制限を受けません。")

holiday_total_placeholder = st.empty()

st.divider()

st.subheader("👥 1つの枠を2名以上にする設定（任意）")
st.info("通常は各枠1名です。増員する場合だけ行を追加し、日付・枠・合計人数を選んでください。2名体制にしたい場合は「2」を入力します。")
st.caption("日直の増員は休日扱いの日に設定してください。平日に日直を設ける場合は、先に上のカレンダーで「休日にする」にチェックを入れます。")

_, num_days = calendar.monthrange(year, month)
date_options = [f"{d}日" for d in range(1, num_days + 1)]

multi_df_template = pd.DataFrame(columns=["日付", "当直枠", "人数"])
if restored_state and restored_state.get("multi") and (int(restored_state["year"]), int(restored_state["month"])) == (year, month):
    multi_df_template = pd.DataFrame(restored_state["multi"], columns=["日付", "当直枠", "人数"])
def clear_multi_slots(y, m):
    """増員設定をすべて消す（表を空の状態で作り直す）。"""
    st.session_state["multi_editor_gen"] = st.session_state.get("multi_editor_gen", 0) + 1
    rs = st.session_state.get("restored_state")
    if rs and (int(rs.get("year", 0)), int(rs.get("month", 0))) == (y, m):
        st.session_state["restored_state"] = {**rs, "multi": []}


# 入力した内容を空欄に戻せるよう、必須の指定はしない（日付・枠・人数のどれかが空の行は使わない）
edited_multi_df = st.data_editor(
    multi_df_template,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    height=150,
    key=f"multi_editor_{year}_{month}_{st.session_state.get('multi_editor_gen', 0)}",
    column_config={
        "日付": st.column_config.SelectboxColumn("日付を選択", options=date_options),
        "当直枠": st.column_config.SelectboxColumn("増員する枠を選択", options=NIGHT_SHIFTS + DAY_SHIFTS),
        "人数": st.column_config.NumberColumn("合計人数", help="追加人数ではなく、その枠に配置する合計人数です。例：1名から2名体制にする場合は2。", min_value=2, max_value=10, step=1)
    }
)
st.caption("行を消すには、行の左端をクリックして選び、表の右上のごみ箱のマークを押します。セルを選んで Delete キーで空欄にすることもできます（空欄のある行は使いません）。")
st.button("増員設定をすべて消す", key="clear_multi", on_click=clear_multi_slots, args=(year, month))

multi_slots_dict = {}
for _, row in edited_multi_df.iterrows():
    d_str = clean_text(row.get("日付", ""))
    s_val = clean_text(row.get("当直枠", ""))
    c_val = row.get("人数")
    if d_str and s_val and s_val in ALL_SHIFT_TYPES and pd.notna(c_val):
        try:
            d_val = int(re.sub(r'\D', '', d_str))
            multi_slots_dict[(d_val, s_val)] = int(c_val)
        except (ValueError, TypeError):
            pass


# ==========================================
# 2. 枠数の集計
# ==========================================
shift_counts = {s: 0 for s in NIGHT_SHIFTS + DAY_SHIFTS}
for d in range(1, num_days + 1):
    if d in locked_days:
        continue   # 「別途決定」の日は別の方法で決めるので、このアプリの枠数に数えない
    is_hol = is_holiday_date(datetime.date(year, month, d), year, month, custom_holidays)
    for s in NIGHT_SHIFTS:
        shift_counts[s] += multi_slots_dict.get((d, s), 1)
    if is_hol:
        for s in DAY_SHIFTS:
            shift_counts[s] += multi_slots_dict.get((d, s), 1)

total_slots = sum(shift_counts.values())
holiday_total_placeholder.metric("🏥 埋める必要がある枠", f"{total_slots} 枠")

# ==========================================
# 3. メイン画面：データの読み込み＆画面入力
# ==========================================
# あまり使わない項目なので、ふだんは折りたたんでおく（途中保存から読み込んだ確定当直があるときは開く）
def _fixed_csv_has_entries(text):
    """途中保存の確定当直に、空でない入力が1つでもあるか。"""
    try:
        df = parse_fixed_csv(text.encode("utf-8"))
    except Exception:
        return False
    return any(str(v).strip() not in ("", "nan", "None") for v in df.drop(columns=["平日/休日"], errors="ignore").values.ravel())


_has_fixed = bool(restored_state and restored_state.get("fixed_csv") and _fixed_csv_has_entries(restored_state["fixed_csv"]))
with st.container(key="fixed_section"), st.expander("📅 先月今月来月の確定当直（押すと開きます）", expanded=_has_fixed):
    st.info("先月・今月・来月の、担当者が決まっている当直を入力してください。先月末の勤務は、月初の勤務間隔の確認に使います。入力がなければ、この項目は飛ばせます。")
    with (st.popover("確定当直の入力例と扱い") if hasattr(st, "popover") else st.container()):
        st.markdown("""
- CSVのアップロードと、下の表への直接入力のどちらでも入力できます。
- 日付は `2026/10/1` または `10/1` の形式で入力します。
- 担当する枠の欄に、医師条件と同じ名前を入力します。複数人の場合は `佐藤、鈴木` のように「、」で区切ります。
- 医師名の区切りには「、」を使ってください。名前に含まれる空白は区切りません。医師条件と同じ表記にしてください。
- 未確定の枠は空欄で構いません。
- 今月の平日に日直を入力する場合は、先に上のカレンダーでその日を「休日にする」にしてください。
- 「別途決定」にした日（年末年始など）の当番が決まったら、ここに入れておくと、当直案と一覧に表示され、前後の勤務間隔の計算にも使われます。
- 今月の確定勤務はNG日・曜日制限より優先され、回数上限も必要に応じて緩められます。確定勤務を基準に追加勤務の間隔を守ります。確定勤務同士が近すぎる場合は、確定内容を維持して警告します。

**今月の同じ日・同じ枠を指定した場合、確定当直への入力と「希望優先度100以上」での指定は、どちらも固定され、計算上の扱いは基本的に同じです。両方に入力する必要はありません。**
    """)

    fixed_columns = ["日付"] + ALL_SHIFT_TYPES   # 平日か休日かは、上のカレンダーで自動的に決まるので入力しない
    fixed_template_df = pd.DataFrame(columns=fixed_columns)
    fixed_csv_template = fixed_template_df.to_csv(index=False).encode('utf-8-sig')

    col_dl_fixed, col_ul_fixed = st.columns(2)
    with col_dl_fixed:
        st.write("▼ Excelで一括入力したい場合")
        st.download_button(
            label="📥 ひな形（CSV）をダウンロード",
            data=fixed_csv_template,
            file_name="確定当直_ひな形.csv",
            mime="text/csv",
        )
    with col_ul_fixed:
        fixed_file = st.file_uploader("確定当直（CSV）をアップロード", type="csv", key=f"fixed_csv_{st.session_state.get('uploader_gen', 0)}")

    if fixed_file is not None:
        try:
            base_fixed_df = parse_fixed_csv(fixed_file.getvalue())
        except Exception as e:
            st.error(f"確定当直ファイルの読み込みに失敗しました。詳細: {e}")
            st.stop()
    elif restored_state and restored_state.get("fixed_csv"):
        base_fixed_df = parse_fixed_csv(restored_state["fixed_csv"].encode("utf-8"))
    else:
        base_fixed_df = pd.DataFrame(columns=fixed_columns)
        base_fixed_df.loc[0] = ["" for _ in range(len(fixed_columns))]

    if "日付" not in base_fixed_df.columns:
        st.error("確定当直CSVに「日付」の列がありません。ひな形の列名を確認してください。")
        st.stop()
    # 日付も普通の列として編集できるようにする（行の見出し＝インデックスにすると、既存の行の日付が直せないため）。
    # 空の列は数字の列とみなされて文字が入らないことがあるので、すべて文字の列にそろえる
    # 以前のひな形にあった「平日/休日」の列は使わないので外す（入っていても読み込める）
    base_fixed_df = base_fixed_df.drop(columns=["平日/休日"], errors="ignore").copy()
    for _c in base_fixed_df.columns:
        base_fixed_df[_c] = base_fixed_df[_c].apply(
            lambda x: "" if pd.isna(x) or str(x).strip().lower() in ("nan", "none", "<na>") else str(x)).astype(str)
    _date_first = ["日付"] + [c for c in base_fixed_df.columns if c != "日付"]
    base_fixed_df = base_fixed_df[_date_first].reset_index(drop=True)

    st.write("表のセルをクリックして、日付（例：11/1）と担当医師名を入力・編集できます。行は下の「＋」で追加できます。")
    edited_fixed_df = st.data_editor(
        base_fixed_df, num_rows="dynamic", use_container_width=True, height=200, hide_index=True,
        column_config={c: st.column_config.TextColumn(c, help="例：11/1 または 2026/11/1" if c == "日付" else None)
                       for c in base_fixed_df.columns},
    )
    edited_fixed_df = edited_fixed_df.reset_index(drop=True)

section_banner("👩‍⚕️", "医師条件の読み込み・入力", "先生ごとの回数・間隔・希望日を入力します", "必須", anchor="sec-staff")
st.info("下の表に医師ごとの条件を入力してください。CSVを使う場合は、ひな形をダウンロードして編集し、アップロードします。NG日は、この下の「🚫 先生ごとのNG日設定」で設定します。")
st.caption("最初に表示される「型1」〜「型9」は、よくある条件の入力例です。名前の欄が型の中身です（例：型4＝A宿直・A日直・外来日直に入る先生）。実際の先生の名前に置き換え、回数などを調整して使ってください。行は「＋」で増やせます。")

# --- 前回の医師条件を引き継ぐ（管理用パスワードが必要） ---
if remote_is_configured():
    _co_msg = st.session_state.pop("carry_over_message", None)
    _co_label = "📋 前回の医師条件を引き継ぐ（NG日・希望日・コメントは空にします）" + "\u200b" * st.session_state.get("carry_over_gen", 0)
    with (st.popover(_co_label) if hasattr(st, "popover") else st.expander(_co_label, expanded=False)):
        if admin_logged_in("carry_over_pw"):
            _co_at, _co_data = work_state_info()
            if _co_data and _co_data.get("staff_csv"):
                st.write(f"途中保存：**{_co_at}**（{_co_data.get('year')}年{_co_data.get('month')}月の作業）")
                st.caption(f"この医師条件（先生の名前・回数・勤務間隔・翌日PM duty・希望優先度・備考など）を、{year}年{month}月用に引き継ぎます。"
                           "月ごとに変わる NG日・希望日・先生からのコメントは空になります。今の画面の医師条件は置き換わります。")
                if (int(_co_data.get("year", 0)), int(_co_data.get("month", 0))) == (year, month):
                    st.warning(f"途中保存は{year}年{month}月（今選んでいる月と同じ）の作業です。NG日なども含めて続きから作業する場合は、"
                               "上の帯の「📂 前回の作業の続きから始める」を使ってください。")
                st.button(f"{year}年{month}月用に医師条件を引き継ぐ", key="carry_over_do", type="primary",
                          on_click=carry_over_staff, args=(year, month))
            else:
                st.caption("まだ途中保存された医師条件がありません。上の帯の「💾 作業を途中保存する」で保存すると、翌月に引き継げます。")
        else:
            st.caption("引き継ぎには管理用パスワードが必要です。")
    if _co_msg:
        (st.success if _co_msg[0] == "success" else st.error)(_co_msg[1])
with st.expander("入力例：曜日・希望日・備考", expanded=False):
    st.markdown("""
| 項目 | 入力方法・意味 |
| --- | --- |
| 翌日PM duty | `水,木` のように半角カンマで区切ります。翌日PMにdutyがある曜日を入力します（例：木曜PMにdutyがある場合は「水」）。原則としてその曜日の宿直を外しますが、翌日が休日なら宿直に入る場合があります。日直は対象外です。 |
| 希望日 | `10,15` はその日のいずれかの枠、`10:A宿直` はその枠を希望します。複数の希望は半角カンマで区切ります。日直は休日（または「休日にする」にした日）だけ指定できます。「別途決定」の日の希望は使いません。 |
| 指定できる枠 | A宿直・B宿直・外来宿直・A日直・B日直・外来日直。表記を一致させてください。 |
| 備考 | 管理用のメモです。「学会」などと書いても計算条件には反映されません。休みはNG日で指定してください。 |

曜日にかかわらず勤務できない日は、下の「🚫 先生ごとのNG日設定」のカレンダーでNGにします。
平日は「宿直」のボタン、休日は「日直」「宿直」のボタンを押してNGにします（休日に終日入れない場合は両方押します）。
    """)
with st.expander("回数・勤務間隔の数え方", expanded=False):
    st.markdown("""
| 項目 | 意味・入力例 |
| --- | --- |
| 最低空ける日数 | 勤務と次の勤務の間に空ける日数です。5日なら、10日の次は16日以降です。 |
| 月間最小回数 | できるだけ確保したい回数です。条件によっては、この回数に届かないことがあります。 |
| 月間最大回数 | 日直と宿直を合わせた月間の上限です。 |
| 休日最大回数 | 土日祝日と「休日にする」にした日に担当する、日直と宿直の合計上限です。 |
| 宿直上限・日直上限 | A・B・外来を合わせた宿直（日直）の月間の上限です。空欄なら制限しません。 |
| 各枠の上限 | A宿直など、それぞれの枠を担当する月間の上限です。 |

1つの枠を1回と数えます。確定指定により同じ日に日直と宿直を担当する場合は2回です。
月間最小回数は、月間最大回数以下に設定してください。
確定当直や希望優先度100以上の希望がある場合は、上限の例外があります。自動追加勤務の間隔は守り、確定勤務同士の間隔違反は警告します。
    """)
with st.expander("希望優先度：通常の希望と、100以上の特別な設定", expanded=False):
    st.markdown("""
- **通常は「1」**を使用します。1〜99は、数字が大きいほど希望を優先しますが、NG日・回数上限・勤務間隔などの範囲内で割り当てます。
- **100以上は、その医師の希望日すべてを確定扱いにする設定**です。「できれば入りたい」という用途には使わないでください。
- 確定扱いの日はNG日・曜日制限より優先され、回数上限が必要に応じて緩められます。その日と自動追加勤務の間隔は守ります。確定勤務同士の間隔違反は警告します。
- 一部の勤務だけを確定させたい場合は、上の「📅 先月今月来月の確定当直」へ入力し、希望優先度は通常の値にしてください。
- 指定の誤りや条件の組み合わせによっては作成できない場合があります。作成後に確定勤務が反映されているか確認してください。

**今月の同じ日・同じ枠を指定した場合、確定当直への入力と「希望優先度100以上」での指定は、どちらも固定され、計算上の扱いは基本的に同じです。両方に入力する必要はありません。**
    """)

# 入力例：よくある9つの型（名前の欄に型の中身が分かる名前を入れる）。実際の先生に置き換えて使う
# 型2（A日直だけ・希望日を確定）の希望日は、選んでいる月の最初の休日2日にする（平日の日直は指定できないため）
# （「休日にする」のチェックで例の表が作り直されないよう、土日祝だけで決める）
_tmpl_hols = [d for d in range(1, calendar.monthrange(year, month)[1] + 1)
              if datetime.date(year, month, d).weekday() >= 5 or jpholiday.is_holiday(datetime.date(year, month, d))][:2]
_tmpl_wish2 = ",".join(f"{d}:A日直" for d in _tmpl_hols)
_tmpl_rows = [
    # 名前, 備考, 希望日, 希望優先度, 最低空ける日数, 月間最小, 月間最大, 休日最大, 宿直上限, 日直上限, A宿直, B宿直, 外来宿直, A日直, B日直, 外来日直
    ("型1_A宿直・A日直", "Aの宿直・日直だけ。月1回", "", 1, 5, 1, 1, 1, 3, 3, 1, 0, 0, 1, 0, 0),
    ("型2_A日直のみ・希望日確定", "決めた日のA日直だけ（希望優先度100で確定）", _tmpl_wish2, 100, 0, 0, 2, 2, 3, 3, 0, 0, 0, 2, 0, 0),
    ("型3_外来日直のみ", "外来日直だけ。月1回まで", "", 1, 5, 0, 1, 4, 3, 3, 0, 0, 0, 0, 0, 1),
    ("型4_A宿直・A日直・外来日直", "Aの宿直・日直と外来日直。月2〜3回", "", 1, 5, 2, 3, 1, 3, 3, 3, 0, 0, 1, 0, 1),
    ("型5_A・B宿直・A日直・外来日直", "型4にB宿直を追加。月2〜3回", "", 1, 5, 2, 3, 1, 3, 3, 3, 3, 0, 1, 0, 1),
    ("型6_外来宿直以外すべて", "型5にB日直を追加。月3〜4回", "", 1, 5, 3, 4, 2, 3, 3, 3, 3, 0, 1, 1, 2),
    ("型7_全枠", "6枠すべて。月3〜4回", "", 1, 5, 3, 4, 2, 3, 3, 3, 3, 3, 1, 1, 2),
    ("型8_日直のみ", "A・B・外来の日直だけ。月1〜3回", "", 1, 4, 1, 3, 3, 3, 3, 0, 0, 0, 3, 3, 3),
    ("型9_A以外", "B・外来の宿直と日直。月3〜4回", "", 1, 4, 3, 4, 2, 3, 3, 0, 3, 3, 0, 3, 3),
]
_tmpl_cols = ["先生の名前", "備考（メモ・説明など自由記入）", REQUEST_COLUMN, "希望優先度(数字が大きいほど優先)",
              "最低空ける日数", "月間最小回数", "月間最大回数", "休日最大回数", "宿直上限", "日直上限",
              "A宿直上限", "B宿直上限", "外来宿直上限", "A日直上限", "B日直上限", "外来日直上限"]
template_data = {c: [r[k] for r in _tmpl_rows] for k, c in enumerate(_tmpl_cols)}
template_data["翌日PM duty"] = [""] * len(_tmpl_rows)
template_data[NG_COLUMN] = [""] * len(_tmpl_rows)
# 列の並びは、これまでのひな形と同じにする
template_data = {c: template_data[c] for c in ["先生の名前", "備考（メモ・説明など自由記入）", "翌日PM duty", NG_COLUMN, REQUEST_COLUMN,
                                               "希望優先度(数字が大きいほど優先)", "最低空ける日数", "月間最小回数", "月間最大回数",
                                               "休日最大回数", "宿直上限", "日直上限", "A宿直上限", "B宿直上限", "外来宿直上限",
                                               "A日直上限", "B日直上限", "外来日直上限"]}
df_template = pd.DataFrame(template_data)
# 保存用CSVと同じUTF-8（BOM付き）に統一。Excelでもそのまま開ける。
csv_template = df_template.to_csv(index=False).encode('utf-8-sig')

col_dl, col_ul = st.columns(2)
with col_dl:
    st.write("▼ Excelで一括入力したい場合")
    st.download_button(
        label="📥 ひな形（CSV）をダウンロード",
        data=csv_template,
        file_name="医師条件_ひな形.csv",
        mime="text/csv",
    )
with col_ul:
    uploaded_file = st.file_uploader("医師条件（以前ダウンロードしたCSVも可）をアップロード", type="csv", key=f"staff_csv_{st.session_state.get('uploader_gen', 0)}")

if uploaded_file is not None:
    if st.session_state.get('last_uploaded_file_id') != uploaded_file.file_id:
        for key in list(st.session_state.keys()):
            if key.startswith("ng_"):
                del st.session_state[key]
        st.session_state['last_uploaded_file_id'] = uploaded_file.file_id
    try:
        base_df = parse_staff_csv(uploaded_file.getvalue())
    except Exception as e:
        st.error(f"医師条件CSVを読み込めませんでした。詳細: {e}")
        st.stop()
elif restored_state and restored_state.get("staff_csv"):
    base_df = parse_staff_csv(restored_state["staff_csv"].encode("utf-8"))
else:
    base_df = df_template.copy()

if "先生の名前" not in base_df.columns:
    st.error("医師条件CSVに「先生の名前」の列がありません。ひな形の列名を確認してください。")
    st.stop()
# 宿直上限・日直上限の列がない（以前の）CSVでも、「休日最大回数」の右に空欄の列を足す（空欄＝制限なし）
for _gc in reversed(list(GROUP_CAP_COLUMNS)):
    if _gc not in base_df.columns:
        _pos = list(base_df.columns).index("休日最大回数") + 1 if "休日最大回数" in base_df.columns else len(base_df.columns)
        base_df.insert(_pos, _gc, None)
    base_df[_gc] = pd.to_numeric(base_df[_gc], errors="coerce")
base_df = base_df.set_index("先生の名前")

if "希望優先度(数字が大きいほど優先)" in base_df.columns:
    base_df["希望優先度(数字が大きいほど優先)"] = pd.to_numeric(base_df["希望優先度(数字が大きいほど優先)"], errors='coerce')

text_cols = ["翌日PM duty", NG_COLUMN, REQUEST_COLUMN, "備考（メモ・説明など自由記入）"]
for c in text_cols:
    if c in base_df.columns:
        base_df[c] = base_df[c].apply(lambda x: "" if pd.isna(x) or str(x).lower() in ["nan", "none", "<na>"] else str(x))

st.markdown("### 👩‍⚕️ 医師条件の入力・編集")
st.write("セルをクリックして編集できます。列名にマウスを合わせると、入力例や説明が表示されます。医師名は重複しない表記にしてください。")

cap_help = "を担当する月間の上限です。確定指定がある場合は例外があります。"
edited_df = st.data_editor(
    base_df,
    num_rows="dynamic",
    use_container_width=True,
    height=300,
    column_config={
        "翌日PM duty": st.column_config.TextColumn(
            "翌日PM duty",
            help="例：木曜PMにdutyがある場合は水。複数は水,木のように入力。翌日が休日なら宿直に入る場合があります。日直は対象外です。確実に外す日はカレンダーでNGを指定してください。"
        ),
        "最低空ける日数": st.column_config.NumberColumn("最低空ける日数", help="勤務間の空き日数。5日なら10日の次は16日以降。確定勤務と追加勤務の間隔も守ります。確定勤務同士の違反は警告します。"),
        "月間最小回数": st.column_config.NumberColumn("月間最小回数（目標）", help="できるだけ確保したい回数です。条件によっては未達になります。月間最大回数以下にしてください。"),
        "月間最大回数": st.column_config.NumberColumn("月間最大回数", help="日直・宿直を合わせた上限です。確定指定がある場合は例外があります。"),
        "休日最大回数": st.column_config.NumberColumn("休日最大回数", help="土日祝と「休日にする」にした日の、日直と宿直の合計上限です。1枠を1回と数えます。"),
        "宿直上限": st.column_config.NumberColumn("宿直上限", help="A宿直・B宿直・外来宿直を合わせた月間の上限です。空欄なら制限なし。確定指定がある場合は例外があります。", min_value=0, step=1),
        "日直上限": st.column_config.NumberColumn("日直上限", help="A日直・B日直・外来日直を合わせた月間の上限です。空欄なら制限なし。確定指定がある場合は例外があります。", min_value=0, step=1),
        **{s + "上限": st.column_config.NumberColumn(s + "上限", help=s + cap_help) for s in ALL_SHIFT_TYPES},
        NG_COLUMN: None,
        REQUEST_COLUMN: st.column_config.TextColumn(
            "希望日",
            help="例：10,15 または 10:A宿直。複数は半角カンマ区切り。日直は休日（または「休日にする」にした日）だけ指定できます。通常の希望は各種条件の範囲内で割り当てます。"
        ),
        "希望優先度(数字が大きいほど優先)": st.column_config.NumberColumn(
            "希望優先度",
            help="通常は1。1〜99は数字が大きいほど優先。100以上はこの医師の希望日すべてが確定扱いになり、NG・上限の例外になります。自動追加勤務との間隔は守ります。"
        ),
        "備考（メモ・説明など自由記入）": st.column_config.TextColumn(
            "備考",
            help="管理用メモです。ここに書いた内容は計算に反映されません。セルをダブルクリックすると全文が表示され、編集できます。表の下の「備考の全文を見る」でも確認できます。",
        )
    },
    # 備考を名前のすぐ右に置き、幅を広くして読みやすくする
    column_order=[c for c in base_df.columns if str(c).startswith("備考")] + [c for c in base_df.columns if not str(c).startswith("備考")],

)

staff_df = edited_df.reset_index()

# 備考が長くて表の中で読み切れない場合のために、全文の一覧を出す
_note_cols = [c for c in staff_df.columns if str(c).strip().startswith("備考")]
_notes = [(clean_text(r.get("先生の名前", "")), "　".join(clean_text(r.get(c, "")) for c in _note_cols if clean_text(r.get(c, ""))))
          for _, r in staff_df.iterrows()]
_notes = [(n, t) for n, t in _notes if n and t]
if _notes:
    with st.expander(f"📝 備考の全文を見る（{len(_notes)}名）", expanded=False):
        # st.table は長い文章を折り返して全部表示する
        st.table(pd.DataFrame(_notes, columns=["先生の名前", "備考"]).set_index("先生の名前"))
staff_df, staff_input_errors = validate_staff_inputs(staff_df, year, month)
if staff_input_errors:
    show_input_errors(staff_input_errors)
    if create_bar_request:
        create_bar_show("<div class='cb-res bad'>⚠ 医師条件の入力に誤りがあります。赤い表示を確認してください。</div>")
    st.stop()

st.markdown("### ⚖️ 当直の枠は足りそう？（目安）")
st.caption("埋める必要がある枠の数と、先生方が入れる回数の合計を比べています。先生ごとの回数は、「月間最大回数」と各枠の上限・休日最大回数のうち、"
           "いちばん少ないものを使います。足りていても、NG日・勤務の間隔などによっては埋まらないことがあります。"
           "確定当直で追加される枠や上限の例外は、この目安に含まれません。")

st.markdown(CAPACITY_CSS, unsafe_allow_html=True)
cap_needs = slot_needs(year, month, num_days, custom_holidays, locked_days, multi_slots_dict)
cap_rows, total_max_capacity = capacity_rows(staff_df, cap_needs)
margin = total_max_capacity - total_slots
c1, c2, c3 = st.columns(3)
c1.metric("🏥 埋める必要がある枠", f"{total_slots} 枠")
c2.metric("👩‍⚕️ 先生方が入れる回数（上限の合計）", f"{total_max_capacity} 回")
if margin >= 0:
    c3.metric("✨ 差し引き", f"{margin} 回分の余裕")
else:
    c3.metric("🚨 差し引き", f"{-margin} 枠 足りません")
cap_short = [r for r in cap_rows[1:] if capacity_status(r[1], r[2]) == "short"]   # 1行目は全体
cap_tight = [r for r in cap_rows if capacity_status(r[1], r[2]) == "tight"]
if margin < 0:
    st.error(f"🚨 **全体で{-margin}枠足りません。** 先生方の「月間最大回数」や各枠の上限を増やすか、担当できる先生を追加してください。")
elif cap_short:
    st.warning(f"⚠️ **全体では足りています（あと{margin}回分の余裕）が、枠の種類別では"
               + "・".join(f"「{r[0]}」" for r in cap_short) + "が足りません。** 下の横棒で確認してください。")
elif cap_tight:
    _others = [r for r in cap_tight if r is not cap_rows[0]]
    if cap_rows and cap_tight[0] is cap_rows[0]:
        _msg = f"足りていますが、全体の余裕があと{margin}回分しかありません。"
        if _others:
            _msg += "・".join(f"「{r[0]}」" for r in _others) + "も余裕が少なめです。"
    else:
        _msg = f"全体（あと{margin}回分の余裕）・枠の種類別ともに足りていますが、" + "・".join(f"「{r[0]}」" for r in _others) + "は余裕が少なめです。"
    st.warning("🟧 " + _msg + "NG日が多いと埋まらない可能性があります。")
else:
    st.success(f"✅ 全体でも枠の種類別でも足りています（あと{margin}回分の余裕）。")

st.markdown("##### 枠の種類ごとの見通し")
st.markdown(capacity_bars_html(cap_rows), unsafe_allow_html=True)
if cap_short:
    st.error("🚨 " + "、".join(f"{r[0]}が{r[1] - r[2]}枠" for r in cap_short) + "足りません。"
             + "該当する先生の" + "・".join(dict.fromkeys(f"「{r[3]}」" for r in cap_short))
             + "を増やすか、担当できる先生を追加してください（「月間最大回数」が小さいと、上限を増やしても回数は増えません）。")
section_banner("🚫", "先生ごとのNG日設定", "先生を選んでカレンダーでNG日を指定します。専用URLで集めたNG日の取り込みもここで行います", anchor="sec-ng")
st.info("先生方に専用URLから入力してもらう場合は、下の「📨 先生方からのNG日入力（専用URL）」を使います。"
        "管理者が直接入力する場合は、医師を選んでカレンダーの「日直」「宿直」のボタンを押してNGにし、最後に「NG日を反映する」を押してください。")
with st.expander("NGの種類・一括操作について", expanded=False):
    st.markdown("""
カレンダーの各日の「日直」「宿直」のボタンを押すと、その勤務だけ「可」⇔「NG」が切り替わります。

| 表示 | 意味 |
| --- | --- |
| 日直 可・宿直 可 | この日についてNGを指定しません（OK）。曜日・回数など、ほかの条件は適用されます。 |
| 日直NG・宿直NG | 日直・宿直ともに勤務できません（全NG）。休日だけです。 |
| 日直NGのみ | 日直に勤務できません。宿直は候補になります（日NG）。休日だけです。 |
| 宿直NGのみ | 宿直に勤務できません。休日なら日直は候補になります（宿NG）。 |

平日は日直がないので、「宿直」のボタンだけあります。
「全日NGにする」は平日を宿NG、休日を全NGにし、「すべてOKに戻す」はNG指定を解除します。
一括操作は押すとすぐ反映されます。確定当直・希望優先度100以上の希望は、NGより優先されます。
灰色の「別途決定」の日は、いちばん上のカレンダーで「別途決定」にした日です。NGは入力できず、当直案でもその日は決めません。

「休日にする」のチェックを外した日は、平日として表示されます（全NG→宿NG として扱い、日NGは平日には関係しません）。
チェックを戻すと、元のNG指定がそのまま復活します。

**「NG日を反映する」は、今開いている画面の計算条件に入れる操作です。**
次回も続きから使う場合は、画面左上の「💾 作業を途中保存する」を押してください。
    """)

ng_layout = st.radio("NGカレンダーの表示", ["月間カレンダー", "1日〜月末を横一列"], horizontal=True, key="ng_calendar_layout")
ng_horizontal = ng_layout == "1日〜月末を横一列"
st.caption("表示を切り替える前に、押したNG日を「反映する」で反映してください。反映済みの内容は、どちらの表示でも共通です。")
if ng_horizontal:
    st.caption("左右へスクロールして日付を選べます。押した欄の色はすぐに変わります。最後に「NG日を反映する」を押してください。")

day_is_holiday = {d: is_holiday_date(datetime.date(year, month, d), year, month, custom_holidays) for d in range(1, num_days + 1)}

valid_staff = staff_df[staff_df["先生の名前"].astype(str).str.strip() != ""]
if not valid_staff.empty:
    doctor_names = valid_staff["先生の名前"].astype(str).tolist()
    remote_admin_panel(
        year, month,
        [(str(r["先生の名前"]), clean_text(r.get("翌日PM duty", ""))) for _, r in valid_staff.iterrows()],
        custom_holidays, next_month_special_holiday, day_is_holiday, locked_days,
    )
    # 医師の選択はタブではなく選択欄にする。
    # （タブは表示の切り替えなどで画面が作り直されると先頭の医師に戻ってしまい、入力ミスの原因になるため）
    # 選択中の医師は「ng_current_doctor」に覚えておく（選択欄そのものの値には頼らない）。
    # 選択欄の表示（NG日数など）が変わると、Streamlit が選択欄を別物とみなして先頭に戻すことがあるため。
    if st.session_state.get("ng_current_doctor") not in doctor_names:
        st.session_state["ng_current_doctor"] = doctor_names[0]

    # 先生方のコメント（専用URLから）と、医師条件の備考。選択欄の💬マークと、カレンダーの上の表示に使う
    doctor_comments = imported_comments(year, month)
    # 備考の列名は「備考」「備考（メモ…）」「備考(メモ…)」など、CSVによって少し違うことがあるので、
    # 「備考」で始まる列をすべて見る
    note_columns = [c for c in valid_staff.columns if str(c).strip().startswith("備考")]
    doctor_notes = {
        str(r["先生の名前"]): "\n".join(clean_text(r.get(c, "")) for c in note_columns if clean_text(r.get(c, "")) and clean_text(r.get(c, "")).lower() not in ("nan", "none"))
        for _, r in valid_staff.iterrows()
    }

    def ng_doctor_label(name):
        count = sum(
            1 for d in range(1, num_days + 1)
            if d not in locked_days and effective_ng(st.session_state.get(f"ng_{name}_{year}_{month}_{d}", "OK"), day_is_holiday[d]) != "OK"
        )
        mark = "　💬" if name in doctor_comments or doctor_notes.get(name) else ""
        return (f"{name}（NG {count}日）" if count else f"{name}（NGなし）") + mark

    # --- 反映していない変更があるまま先生を切り替えようとしたときの注意 ---
    # カレンダーで押しただけの変更は、別の先生に切り替えると消えてしまう。
    # 1回目の切り替えは止めて注意を出し、もう一度切り替えたら（変更を取り消して）移る。
    def ng_is_dirty(name):
        rev = st.session_state.get(f"ng_rev::{name}")
        return rev is not None and st.session_state.get(f"ng_dirty::{name}") == rev

    def ng_leave_allowed(current):
        if current and ng_is_dirty(current) and st.session_state.get("ng_leave_ack") != current:
            st.session_state["ng_leave_ack"] = current
            st.session_state["ng_leave_warning"] = current
            return False
        if current:
            st.session_state.pop(f"ng_dirty::{current}", None)   # 反映せずに移る＝変更は取り消し
        st.session_state.pop("ng_leave_ack", None)
        st.session_state.pop("ng_leave_warning", None)
        return True

    def move_ng_doctor(step, names):
        # 名前の一覧は押したときの画面のものを受け取る。
        # 選択中の医師が一覧にない場合（名前を書き換えた・CSVを読み込み直した直後など）は先頭の医師にする
        if not names:
            return
        current = st.session_state.get("ng_current_doctor")
        if not ng_leave_allowed(current):
            return
        i = names.index(current) if current in names else -step
        st.session_state["ng_current_doctor"] = names[(i + step) % len(names)]

    def on_ng_doctor_select(widget_key):
        prev = st.session_state.get("ng_current_doctor")
        new = st.session_state.get(widget_key)
        if prev and prev != new and not ng_leave_allowed(prev):
            st.session_state[widget_key] = prev   # 切り替えを取りやめて、元の先生に戻す
            return
        st.session_state["ng_current_doctor"] = new

    sel_col, prev_col, next_col = st.columns([6, 1.2, 1.2])
    # 選択欄は、表示が変わるたびに作り直し、いつも ng_current_doctor の先生を選んだ状態で表示する
    _labels = [ng_doctor_label(n) for n in doctor_names]
    _sel_key = "ng_doctor_select_" + hashlib.sha256("\n".join(_labels).encode("utf-8")).hexdigest()[:12]
    for _k in [k for k in st.session_state.keys() if str(k).startswith("ng_doctor_select_") and k != _sel_key]:
        del st.session_state[_k]
    st.session_state[_sel_key] = st.session_state["ng_current_doctor"]
    sel_col.selectbox("NG日を入力する医師", doctor_names, key=_sel_key, format_func=ng_doctor_label,
                      on_change=on_ng_doctor_select, args=(_sel_key,))
    selected_doc = st.session_state["ng_current_doctor"]
    # ボタンの高さを選択欄にそろえるための余白
    prev_col.markdown("<div style='height:1.75rem'></div>", unsafe_allow_html=True)
    next_col.markdown("<div style='height:1.75rem'></div>", unsafe_allow_html=True)
    prev_col.button("◀ 前の医師", key="ng_prev_doctor", on_click=move_ng_doctor, args=(-1, list(doctor_names)), use_container_width=True)
    next_col.button("次の医師 ▶", key="ng_next_doctor", on_click=move_ng_doctor, args=(1, list(doctor_names)), use_container_width=True)
    if st.session_state.get("ng_leave_warning") == selected_doc and ng_is_dirty(selected_doc):
        st.warning(f"**{selected_doc}先生のカレンダーに、まだ反映していない変更があります。** "
                   "カレンダーの下の「反映する」ボタンを押してください。"
                   "反映せずにほかの先生へ移る場合は、もう一度切り替えてください（変更は取り消されます）。")

    for t_idx, doc_name in enumerate(doctor_names):
        original_idx = valid_staff.index[t_idx]
        doc_row = valid_staff.loc[original_idx]
        hard_days = pm_duty_weekdays(doc_row.get("翌日PM duty", ""))
        current_ng_dict = parse_ng_dict(doc_row.get(NG_COLUMN, ""), year, month)

        # 初回だけ、表（CSV）のNG指定を読み込む。以後は画面で保存した内容を使う。
        for d in range(1, num_days + 1):
            chk_key = f"ng_{doc_name}_{year}_{month}_{d}"
            if chk_key not in st.session_state:
                st.session_state[chk_key] = current_ng_dict.get(d, "OK")

        if doc_name == selected_doc:
            with st.container(border=True):
                st.markdown(f"#### 🩺 {doc_name} 先生のNG日")
                note = doctor_notes.get(doc_name, "")
                comment, comment_at = doctor_comments.get(doc_name, ("", ""))
                if note or comment:
                    with st.container(border=True):
                        if comment:
                            st.markdown(f"💬 **先生からのコメント**（専用URLから・{comment_at}）")
                            st.text(comment)
                        if note:
                            st.markdown("📝 **医師条件の備考**")
                            st.text(note)
                # 保存値は書き換えず、その日に実際に効く内容だけを表示する
                # 「別途決定」の日は、別の方法で当番を決めるのでNGとして扱わない
                shown = {d: "OK" if d in locked_days else effective_ng(st.session_state[f"ng_{doc_name}_{year}_{month}_{d}"], day_is_holiday[d]) for d in range(1, num_days + 1)}

                saved_strs = []
                for d, val in shown.items():
                    if val == "全NG": saved_strs.append(f"{d}日")
                    elif val == "日NG": saved_strs.append(f"{d}日(日直NG)")
                    elif val == "宿NG": saved_strs.append(f"{d}日(宿直NG)")
                if saved_strs:
                    st.success(f"✅ **反映済みのNG日:** {', '.join(saved_strs)}")
                else:
                    st.info("💡 **現在、反映されているNG日はありません**")
                if import_undo_doctors(year, month, [doc_name]):
                    cu1, cu2 = st.columns([3, 1.6])
                    cu1.caption("このNG日には、先生がURLから提出した内容が取り込まれています。"
                                "取り消すと、取り込む前のNG日に戻ります（取り込み後に手直しした分も戻ります）。")
                    cu2.button("↩️ 取り込みを取り消す", key=f"undo_import_{doc_name}_{year}_{month}", use_container_width=True,
                               on_click=undo_remote_import, args=(year, month, [doc_name]))

                component_days = []
                for d in range(1, num_days + 1):
                    dt = datetime.date(year, month, d)
                    hol = day_is_holiday[d]
                    options = ["OK", "全NG", "日NG", "宿NG"] if hol else ["OK", "宿NG"]
                    component_days.append({
                        "day": d, "weekday": WEEKDAYS_JA[dt.weekday()],
                        "options": options, "value": shown[d],
                        "warning": pm_duty_restricted(dt, hard_days, year, month, custom_holidays, next_month_special_holiday),
                        "kind": "saturday" if dt.weekday() == 5 and not jpholiday.is_holiday(dt) and d not in custom_holidays else ("holiday" if hol else "weekday"),
                        "locked": d in locked_days,
                    })
                revision = hashlib.sha256(json.dumps([year, month, doc_name, ng_layout, component_days], ensure_ascii=False).encode()).hexdigest()
                component_key = f"ng_editor_{'row' if ng_horizontal else 'month'}_{doc_name}_{year}_{month}"
                st.markdown("<div style='color:#bf5700;background:#fff0c2;border:1px solid #ef9b20;border-radius:6px;padding:8px 10px;font-size:0.9rem;font-weight:600;'>⚠は、翌日PMにdutyがあるため、原則としてその日の宿直を外すことを示します。翌日が休日の場合は⚠を表示せず、この制限の対象外になります。</div>", unsafe_allow_html=True)
                if locked_days:
                    st.caption("灰色の日（" + "・".join(f"{d}日" for d in sorted(locked_days)) + "）は、いちばん上のカレンダーで「別途決定」にした日です。NGは入力できず、当直案でもこの日の当直は決めません（確定当直に入れた分だけ表示されます）。")
                response = horizontal_ng_component(HORIZONTAL_NG_HTML)(
                    days=component_days, mode="row" if ng_horizontal else "month",
                    offset=datetime.date(year, month, 1).weekday(), doctor=doc_name,
                    version=revision, key=component_key, default=None, report_dirty=True,
                )
                st.session_state[f"ng_rev::{doc_name}"] = revision
                seen_key = component_key + "_last_token"
                if isinstance(response, dict) and response.get("token") != st.session_state.get(seen_key):
                    st.session_state[seen_key] = response.get("token")
                    if response.get("kind") == "dirty":
                        # カレンダーで押したが、まだ反映していない変更があるかどうか
                        if response.get("dirty") and response.get("version") == revision:
                            st.session_state[f"ng_dirty::{doc_name}"] = revision
                        else:
                            st.session_state.pop(f"ng_dirty::{doc_name}", None)
                            st.session_state.pop("ng_leave_warning", None)
                            st.session_state.pop("ng_leave_ack", None)
                        response = None
                if isinstance(response, dict) and response.get("kind") != "dirty" and response.get("token") == st.session_state.get(seen_key) \
                        and st.session_state.get(seen_key + "_applied") != response.get("token"):
                    st.session_state[seen_key + "_applied"] = response.get("token")
                    st.session_state.pop(f"ng_dirty::{doc_name}", None)
                    st.session_state.pop("ng_leave_warning", None)
                    st.session_state.pop("ng_leave_ack", None)
                    values = response.get("values")
                    if response.get("version") == revision and isinstance(values, list) and len(values) == num_days:
                        if all(v in item["options"] for v, item in zip(values, component_days)):
                            changed = False
                            for d, (value, item) in enumerate(zip(values, component_days), 1):
                                # 画面で実際に変更した日だけ保存値を更新する
                                # （休日設定の切り替えで隠れているNG指定を消さないため）
                                if value != item["value"]:
                                    st.session_state[f"ng_{doc_name}_{year}_{month}_{d}"] = value
                                    changed = True
                            if changed:
                                st.rerun()

                _, col_btn1, col_btn2 = st.columns([6, 1.5, 1.5])
                with col_btn1:
                    st.button("全日NGにする", key=f"btn_all_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, "全NG", custom_holidays), use_container_width=True)
                with col_btn2:
                    st.button("すべてOKに戻す", key=f"btn_clear_{doc_name}_{year}_{month}", on_click=set_all_ng, args=(doc_name, year, month, num_days, "OK", custom_holidays), use_container_width=True)

        # 保存値（指定したとおりの内容）を医師条件へ書き込む。
        # 平日の全NGは計算上は宿NGと同じ、平日の日NGは計算に影響しない。
        ng_items = []
        for d in range(1, num_days + 1):
            val = st.session_state.get(f"ng_{doc_name}_{year}_{month}_{d}", "OK")
            if d in locked_days:
                continue   # 「別途決定」の日のNGは計算に使わない
            if val == "全NG":
                ng_items.append(str(d))
            elif val != "OK":
                ng_items.append(f"{d}:{val}")
        staff_df.at[original_idx, NG_COLUMN] = ",".join(ng_items)


st.divider()
st.markdown("### 📂 医師条件をCSVでダウンロード（手元にファイルで残す場合）")
st.write("医師名・回数・勤務間隔・希望日・NG日・備考をCSVファイルとしてパソコンなどにダウンロードします。このファイルは「医師条件」のアップロード欄から読み込めます。")
st.caption("ダウンロードの前に、各医師の「NG日を反映する」を押してください。対象年月・休日と別途決定の日・増員設定・確定当直・作成した当直案・色分けとそのメモは、このCSVには含まれません。"
           "続きから作業するだけなら、CSVではなく「💾 途中保存」を使うほうが簡単です。")

# ダウンロードするCSVも、画面と同じく「先生の名前 → 備考 → その他」の順にそろえる（読み込みは列名で行うので、古い並びのCSVもそのまま使える）
_first = ["先生の名前"] + [c for c in staff_df.columns if str(c).startswith("備考")]
current_csv = staff_df[_first + [c for c in staff_df.columns if c not in _first]].to_csv(index=False).encode('utf-8-sig')
st.download_button(
    label="📥 医師条件をCSVでダウンロード",
    data=current_csv,
    file_name=f"医師条件_{year}年{month}月.csv",
    mime="text/csv",
    use_container_width=True
)


# ==========================================
# 4. 実行ボタンと結果表示
# ==========================================
st.markdown("""
<style>
.create-guide{border:2px dashed #ff4b4b;background:#fff5f5;color:#9b1c1c;border-radius:10px;padding:14px 16px;font-size:1.05rem;text-align:center}
.create-guide span{display:block;font-size:.85rem;color:#7a4a4a;margin-top:2px}
.st-key-create_duty_action button {
    width: 100%;
    min-height: 60px;
}
.st-key-create_duty_action button p {
    font-size: 1.15rem;
    font-weight: 700;
}
</style>
""", unsafe_allow_html=True)
section_banner("🎯", "当直案の作成・確認", "不足時の優先順位を確認して、当直案を作ります", anchor="sec-create")

# --- 先生方のコメント・備考の一覧（当直案を作る前後にまとめて見られるように） ---
def comment_note_rows(staff, year, month):
    comments = imported_comments(year, month)
    note_cols = [c for c in staff.columns if str(c).strip().startswith("備考")]
    rows = []
    for _, r in staff.iterrows():
        name = clean_text(r.get("先生の名前", ""))
        if not name:
            continue
        note = "\n".join(clean_text(r.get(c, "")) for c in note_cols
                         if clean_text(r.get(c, "")) and clean_text(r.get(c, "")).lower() not in ("nan", "none"))
        comment, at = comments.get(name, ("", ""))
        if comment or note:
            rows.append((name, comment, at, note))
    return rows

_cn_rows = comment_note_rows(staff_df, year, month)
# 見出しの文字は固定にする（文字が変わると、開いている欄が閉じてしまうため）
with st.expander("💬 先生方のコメント・備考の一覧", expanded=False):
    if _cn_rows:
        st.caption(f"{len(_cn_rows)}名分。コメントは、URLから提出されたNG日を取り込んだ先生の分だけ表示します。備考は医師条件の「備考」列です。")
        _parts = []
        for n, c, t, no in _cn_rows:
            empty = "<span class='e'>—</span>"
            c_html = html_escape(c).replace("\n", "<br>") if c else empty
            if c and t:
                c_html += "<div class='t'>" + html_escape(t) + "</div>"
            no_html = html_escape(no).replace("\n", "<br>") if no else empty
            _parts.append("<tr><td class='n'>" + html_escape(n) + "</td><td>" + c_html + "</td><td>" + no_html + "</td></tr>")
        _body = "".join(_parts)
        st.markdown(
            "<style>.cn-table{width:100%;border-collapse:collapse;font-size:.92rem}"
            ".cn-table th,.cn-table td{border:1px solid rgba(128,144,160,.35);padding:6px 9px;vertical-align:top;text-align:left;overflow-wrap:anywhere}"
            ".cn-table th{background:rgba(128,144,160,.12);white-space:nowrap}.cn-table td.n{white-space:nowrap;font-weight:700}"
            ".cn-table .t{font-size:.78rem;color:#7a8494;margin-top:2px}.cn-table .e{color:#b0b7c3}</style>"
            "<table class='cn-table'><thead><tr><th>先生</th><th>💬 先生からのコメント</th><th>📝 備考</th></tr></thead>"
            f"<tbody>{_body}</tbody></table>", unsafe_allow_html=True)
    else:
        st.caption("表示するコメント・備考はありません。コメントは、URLから提出されたNG日を取り込むと表示されます。")

st.subheader("🎯 人数が足りないときに優先して埋める枠（任意）")
st.info("全員の条件を満たす当直表が作れない場合に、どの枠から優先して埋めるかを決めます。数字が小さいほど優先し、同じ数字の枠は同順位として扱います。")
st.caption("人数が足りている月は、この設定は結果に影響しません。初期設定は、A宿直・B宿直・A日直・B日直が1、外来宿直が2、外来日直が3です。")

priority_options = list(range(1, len(ALL_SHIFT_TYPES) + 1))
shortage_priority = {}
priority_cols = st.columns(len(ALL_SHIFT_TYPES))
for col, s in zip(priority_cols, NIGHT_SHIFTS + DAY_SHIFTS):
    if f"shortage_priority_{s}" not in st.session_state:
        st.session_state[f"shortage_priority_{s}"] = DEFAULT_SHORTAGE_PRIORITY[s]
    shortage_priority[s] = col.selectbox(s, priority_options, key=f"shortage_priority_{s}")


def reset_shortage_priority():
    for s, rank in DEFAULT_SHORTAGE_PRIORITY.items():
        st.session_state[f"shortage_priority_{s}"] = rank

st.button("優先順位を初期設定に戻す", key="reset_shortage_priority", on_click=reset_shortage_priority)
priority_order = shortage_priority_groups(shortage_priority)
st.caption("現在の優先順位：" + " ＞ ".join("・".join(g) for g in priority_order))

# --- 作業の途中保存 ---
work_snapshot = {
    "year": year, "month": month,
    "holidays": sorted(int(d) for d in custom_holidays),
    "locked": sorted(int(d) for d in locked_days),
    "next_special": bool(next_month_special_holiday),
    "multi": [[f"{d}日", s, int(c)] for (d, s), c in multi_slots_dict.items()],
    "fixed_csv": edited_fixed_df.to_csv(index=False),
    "staff_csv": staff_df.to_csv(index=False),
    "priority": {k: int(v) for k, v in shortage_priority.items()},
    "comments": {n: list(v) for n, v in imported_comments(year, month).items()},
    "alerts": {k: int(st.session_state.get(f"alert_{k}", v) or 0) for k, v in MATRIX_COUNT_ALERTS.items()},
}
# 右上の途中保存ボタンが押されていたら、ここで（全部の入力を読み終えてから）保存する
if work_save_request:
    _slot = WORK_SAVE_SLOT or st.empty()
    _slot.markdown("<span class='work-save-note busy fresh'>💾 途中保存しています…</span>", unsafe_allow_html=True)
    try:
        work_state_save(work_snapshot)
        st.session_state["work_save_note"] = remote_now()[-8:-3]
        st.session_state.pop("work_restore_note", None)
        st.session_state["work_note_fresh"] = True   # スマートフォンでは、保存の直後に画面上部へ数秒だけ表示する
        # カレンダーで押したまま反映していない変更は、途中保存に入らないので知らせる
        st.session_state["work_save_dirty"] = [
            k.split("::", 1)[1] for k in list(st.session_state.keys())
            if k.startswith("ng_dirty::") and st.session_state[k] == st.session_state.get("ng_rev::" + k.split("::", 1)[1])]
        _saved_ok = True
    except Exception as e:
        _saved_ok = False
        _slot.markdown("<span class='work-save-note bad fresh'>⚠ 途中保存できませんでした</span>", unsafe_allow_html=True)
        st.error(f"途中保存できませんでした。通信状態を確認して、もう一度押してください。詳細：{e}")
    if _saved_ok:
        # 画面右上の「前回の作業の続きから始める」に新しい保存日時が出るよう、画面を作り直す
        st.rerun()

create_button_container = st.container(key="create_duty_action")
result_notice_container = st.container()

staff_df = staff_df[staff_df['先生の名前'].astype(str).str.strip() != '']
staff_df = staff_df.dropna(subset=['先生の名前']).reset_index(drop=True)

fixed_df, fixed_input_errors = validate_fixed_inputs(edited_fixed_df, staff_df['先生の名前'].tolist(), year, month, custom_holidays)
if fixed_input_errors:
    show_input_errors(fixed_input_errors)
    if create_bar_request:
        create_bar_show("<div class='cb-res bad'>⚠ 確定当直の入力に誤りがあります。赤い表示を確認してください。</div>")
    st.stop()
current_signature = (input_signature(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority), tuple(sorted(locked_days)))
if 'generated_df' in st.session_state and st.session_state.get('generated_signature') != current_signature:
    for key in ['generated_df', 'past_worked_dates', 'future_worked_dates', 'generated_warnings', 'generated_signature', 'generated_year', 'generated_month']:
        st.session_state.pop(key, None)
    st.session_state['result_needs_refresh'] = True
    st.session_state.pop("create_bar_summary", None)
    if not create_bar_request:
        create_bar_show("<div class='cb-res wait'>条件が変わりました。もう一度作成してください。</div>")
if st.session_state.get('result_needs_refresh'):
    st.warning("条件が変更されています。画面下の帯の「🚀 当直案を作成」で作り直してください。前回の結果表示とダウンロードは停止しました。")

if len(staff_df) > 0:
    with create_button_container:
        # 作成ボタンは画面の端の「🚀 当直案を作成」1つだけにする（同じ働きのボタンが2つあると迷うため）
        create_clicked = create_bar_request
        st.markdown("<div class='create-guide'>🚀 当直案は、画面のいちばん下の帯にある<b>「当直案を作成」</b>ボタンで作ります"
                    "<span>（どこまでスクロールしていても押せます）</span></div>", unsafe_allow_html=True)
        st.caption("※ 押すたびに、同じ条件でも違うパターンの当直表を作ります。気に入った案は、結果の緑のボタン「📥 この一覧をExcelでダウンロード」で保存してください（作成した当直案は途中保存には含まれません）。")
    if create_clicked:
        with st.spinner("当直案を計算中…（通常は最大60秒、不足枠の確認を含む場合は計算時間が最大75秒です）"):
            try:
                df_result, success, error_reasons, past_worked_dates, future_worked_dates = generate_shift(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority, locked_days)

                if df_result is not None:
                    st.session_state['generated_signature'] = current_signature
                    st.session_state['generated_year'] = year
                    st.session_state['generated_month'] = month
                    st.session_state['generated_warnings'] = error_reasons
                    st.session_state['result_needs_refresh'] = False
                if df_result is not None and not df_result.empty:
                    st.session_state['generated_df'] = df_result
                    st.session_state['past_worked_dates'] = past_worked_dates or {}
                    st.session_state['future_worked_dates'] = future_worked_dates or {}
                else:
                    st.session_state.pop('generated_df', None)
                    with result_notice_container:
                        show_schedule_notice(None, error_reasons)
            except Exception as e:
                st.session_state.pop('generated_df', None)
                with result_notice_container:
                    show_schedule_notice(None, [f"当直計算中にエラーが発生しました。詳細: {e}"])
        if 'generated_df' not in st.session_state:
            st.session_state.pop("create_bar_summary", None)
            create_bar_show("<div class='cb-res bad'>⚠ 当直案を作れませんでした。「結果」で理由を確認してください。</div>")

    if 'generated_df' in st.session_state:
        year = st.session_state['generated_year']
        month = st.session_state['generated_month']
        df_result = st.session_state['generated_df'].reindex(columns=RESULT_COLUMNS)
        past_worked_dates = st.session_state.get('past_worked_dates', {})
        future_worked_dates = st.session_state.get('future_worked_dates', {})

        shift_columns = ALL_SHIFT_TYPES
        doctors_list = staff_df['先生の名前'].astype(str).tolist()

        def split_names(value):
            return [x.strip() for x in re.split(r'[、,]', str(value))]

        st.markdown("<div id='sec-result'></div>", unsafe_allow_html=True)
        st.subheader("🗓️ 医師別の当直一覧")
        st.caption("決まった当直は枠名、当直NGは×、翌日PM dutyで宿直を外す日は斜線で表示します。休日のマスは上半分が日直、下半分が宿直です。"
                   "希望日が通った当直は茶色の太い枠で囲み、回数が多い先生は右端の回数を赤い枠で示します。マスにカーソルを合わせると、その先生の行とその日の列が黄色く色付きます。")
        matrix_days, matrix_rows = build_duty_matrix(df_result, staff_df, year, month, custom_holidays, next_month_special_holiday,
                                                     comments=imported_comments(year, month))
        missing_total = sum(len(info['shortage']) for info in matrix_days)
        # 下の帯の「当直案を作成」ボタンの横に、結果のまとめを出す
        _t = remote_now()[-8:-3]
        _summary = (f"<div class='cb-res bad'>⚠ {_t} 作成：未決定 {missing_total}枠</div>" if missing_total
                    else f"<div class='cb-res ok'>✅ {_t} 作成：全枠決定</div>")
        if create_clicked or not st.session_state.get("create_bar_summary"):
            st.session_state["create_bar_summary"] = _summary
        create_bar_show(st.session_state["create_bar_summary"])
        # アラートの基準（この回数以上で赤く表示）。0はアラートなし
        for _k, _v in MATRIX_COUNT_ALERTS.items():
            if f"alert_{_k}" not in st.session_state:
                st.session_state[f"alert_{_k}"] = _v

        def _reset_alerts():
            for _k2, _v2 in MATRIX_COUNT_ALERTS.items():
                st.session_state[f"alert_{_k2}"] = _v2

        with st.expander("⚠ アラートの基準（回数が多い先生を赤く表示）", expanded=False):
            st.caption("それぞれの回数がこの数以上の先生を、表の右端で赤く表示します。0にすると、その項目はアラートを出しません。")
            _acols = st.columns(len(MATRIX_COUNT_ALERTS))
            for _col, _k in zip(_acols, MATRIX_COUNT_ALERTS):
                _col.number_input(f"{MATRIX_ALERT_LABELS[_k]}（回以上）", min_value=0, max_value=31, step=1, key=f"alert_{_k}")
            st.button("初期設定に戻す（宿直4回以上・休日3回以上）", key="alert_reset", on_click=_reset_alerts)
        alert_limits = {_k: int(st.session_state.get(f"alert_{_k}", _v) or 0) for _k, _v in MATRIX_COUNT_ALERTS.items()}
        _alerts = count_alert_messages(matrix_rows, alert_limits)
        # 計算結果のまとめ（不足・注意事項・当直回数のアラート）は、作成ボタンの下の1か所にまとめて出す
        with result_notice_container:
            show_schedule_notice(st.session_state['generated_df'], st.session_state.get('generated_warnings', []), _alerts)
        st.markdown(duty_matrix_html(matrix_days, matrix_rows, year, month, alert_limits), unsafe_allow_html=True)
        try:
            # 上下の帯の赤いボタン（途中保存・当直案を作成）と見分けやすいよう、Excelらしい緑にする
            st.markdown("<style>.st-key-xlsx_download button{background:#1d6f42 !important;border:1px solid #1d6f42 !important;color:#fff !important;font-weight:700}"
                        ".st-key-xlsx_download button:hover{background:#155a35 !important;border-color:#155a35 !important}"
                        ".st-key-xlsx_download button p{color:#fff !important}</style>", unsafe_allow_html=True)
            with st.container(key="xlsx_download"):
                st.download_button(
                    label="📥 この一覧をExcelでダウンロード（印刷用）",
                    data=duty_matrix_xlsx(matrix_days, matrix_rows, year, month, alert_limits),
                    file_name=f"当直一覧_{year}年{month}月.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
        except ImportError:
            st.error("Excelでダウンロードするには、Excelを扱う部品（openpyxl）が必要です。requirements.txt に「openpyxl」の1行を追加してください。")

        st.divider()
        st.subheader("📅 作成した当直案")
        table_container = st.container()

        st.divider()
        # (見出し, 選択欄のキー, メモ欄のキー, セルの色)
        highlight_colors = [
            ("🟨 **黄色**", "hl_yellow", "memo_y", 'background-color: #fff200; color: #000000; font-weight: bold; border: 2px solid #ffcc00;'),
            ("🟥 **赤色**", "hl_red", "memo_r", 'background-color: #ffcccc; color: #000000; font-weight: bold; border: 2px solid #ff6666;'),
            ("🟦 **水色**", "hl_blue", "memo_b", 'background-color: #cce5ff; color: #000000; font-weight: bold; border: 2px solid #66b2ff;'),
            ("🟩 **緑色**", "hl_green", "memo_g", 'background-color: #ccffcc; color: #000000; font-weight: bold; border: 2px solid #66ff66;'),
            ("🟧 **オレンジ**", "hl_orange", "memo_o", 'background-color: #ffe5b4; color: #000000; font-weight: bold; border: 2px solid #ffb347;'),
            ("🟫 **茶色**", "hl_brown", "memo_br", 'background-color: #e6ccb3; color: #000000; font-weight: bold; border: 2px solid #c68c53;'),
            ("🟪 **紫色**", "hl_purple", "memo_p", 'background-color: #e6ccff; color: #000000; font-weight: bold; border: 2px solid #b366ff;'),
            ("💗 **ピンク**", "hl_pink", "memo_pi", 'background-color: #ffccff; color: #000000; font-weight: bold; border: 2px solid #ff66ff;'),
        ]
        # 見出しの文字は固定にする（文字が変わると、選んでいる途中で欄が閉じてしまうため）
        with st.expander("🔍 特定の医師の当直を色別でハイライト（任意・クリックで開く）", expanded=False):
            st.write("※各色のすぐ下にあるメモ欄に「神経内科」「呼吸器内科」など自由に書き込めます。")

            selected_by_color = []
            for pair_start in range(0, len(highlight_colors), 2):
                columns = st.columns(2)
                for column, (title, select_key, memo_key, style) in zip(columns, highlight_colors[pair_start:pair_start + 2]):
                    with column:
                        st.markdown(title)
                        label = title.replace("*", "").split(" ", 1)[-1]
                        selected = st.multiselect(label, options=doctors_list, default=[], key=select_key, label_visibility="collapsed")
                        st.text_input(label + "メモ", key=memo_key, placeholder="自由記入欄", label_visibility="collapsed", autocomplete="off")
                        selected_by_color.append((set(selected), style))
                st.write("")

        def color_highlighted_doctor(val):
            val_str = str(val)
            if val_str in ("-", ""):
                return ''
            if "⚠️不足" in val_str:
                return 'background-color: #ffe6e6; color: #cc0000; font-weight: bold; border: 2px solid #cc0000;'
            for doc in split_names(val_str):
                for selected, style in selected_by_color:
                    if doc in selected:
                        return style
            return ''

        with table_container:
            if locked_days:
                st.caption("「別途決定」にした日（" + "・".join(f"{month}/{d}" for d in sorted(locked_days)) + "）は、このアプリでは当直を決めていません（確定当直に入れた分だけ表示）。")
            # 表の中身はデータとして渡す。部品（一時ファイル）は1つのまま増えない。
            horizontal_ng_component(HOVER_TABLE_HTML)(
                table=hover_schedule_data(df_result, shift_columns, doctors_list, color_highlighted_doctor),
                key="hover_schedule_table",
                default=None,
            )

            csv_result = df_result.to_csv(index=False).encode('utf-8-sig')
            st.download_button(
                label="📥 表示中の当直案をCSVでダウンロード",
                data=csv_result,
                file_name=f"当直案_{year}年{month}月.csv",
                mime="text/csv",
            )


else:
    st.warning("☝️ 表に先生の名前を入力するか、CSVファイルをアップロードしてください。")
