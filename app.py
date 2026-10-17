import streamlit as st
import pandas as pd
import datetime
import calendar
import io
import math
import random
import re
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
*{box-sizing:border-box}body{margin:0;font-family:system-ui,sans-serif;color:#243247;background:white;font-size:14px}.strip{display:flex;gap:8px;overflow-x:auto;width:100%;padding:6px 2px 16px;align-items:stretch;scrollbar-width:auto}.cell{flex:0 0 112px;width:112px;min-width:112px;border:2px solid #dfe3ea;border-radius:9px;padding:6px;background:#f8fafc}.head{height:78px;display:flex;flex-direction:column;justify-content:center;align-items:center;border-radius:5px;gap:5px;font-weight:750;white-space:nowrap}.day{font-size:16px}.state{font-size:14px}.day.holiday{color:#c92336}.day.saturday{color:#1670c5}select{width:100%;height:36px;margin-top:6px;font-size:14px;font-weight:650;border:1px solid #8b95a5;border-radius:5px;background:white;color:#243247;padding:2px}button{background:#ff4b4b;color:white;border:0;border-radius:7px;padding:12px 18px;font:600 14px system-ui;cursor:pointer}button:focus-visible,select:focus-visible{outline:3px solid #4789ff;outline-offset:2px}.note{margin:8px 0;font-size:13px;color:#566174;min-height:20px}

.strip.month{display:grid;grid-template-columns:repeat(7,minmax(0,1fr));gap:6px;overflow:visible;padding-bottom:8px}
.month .cell{width:auto;min-width:0;padding:5px;flex:none}
.month .head{height:72px}.weekday{text-align:center;font-weight:700;padding:4px}.blank{min-height:126px;background:#f5f6f8;border-radius:9px}
@media(max-width:600px){.strip.month{gap:3px}.month .cell{padding:2px;border-width:1px}.month .day{font-size:12px}.month .state{font-size:11px}.month select{font-size:11px;padding:0;height:30px}.month .head{height:66px}.month .blank{min-height:108px}.weekday{font-size:12px}}

/* 上段が日直、下段が宿直。不可の勤務帯だけ塗る。 */
.cell,.month .cell{background:#fff;border-color:#cbd2dc}
.head,.month .head{height:124px;gap:4px;justify-content:flex-start;padding-top:2px;color:#243247;background:transparent}
.state{width:100%;display:grid;grid-template-rows:repeat(2,36px);gap:0;border:1px solid #d6dce5;border-radius:5px;overflow:hidden;order:3}
.band{display:flex;align-items:center;justify-content:center;font-size:13px;font-weight:750;background:#fff;color:#425268;white-space:nowrap}
.band + .band{border-top:1px solid #d6dce5}.band.off{background:#fcfcfd;color:#d8dde5;font-weight:400}
/* 日直・宿直の段そのものを押して、可⇔NG を切り替える */
.band.tap{cursor:pointer;user-select:none;-webkit-tap-highlight-color:transparent;gap:3px}
.band.tap:hover{filter:brightness(.95)}
.band.tap:focus-visible{outline:3px solid #4789ff;outline-offset:-3px}
.band .mark{font-size:12px;font-weight:900}
/* 「可」の段は緑で目立たせる（休日に日直の指定を忘れていないか一目で分かるように） */
.band.ok{background:#e3f5e9 !important;color:#12692f !important;font-weight:800;box-shadow:inset 0 0 0 1px #8fd0a5}
.cell[data-state="全NG"] .band{background:#b42332;color:#fff}
.cell[data-state="日NG"] .day-band{background:#9a4700;color:#fff}
.cell[data-state="宿NG"] .night-band{background:#1856a4;color:#fff}
.warning-slot{height:22px;min-height:22px;display:flex;align-items:center;justify-content:center}
.weekday-warning{color:#bf5700;background:#fff0c2;border:1px solid #ef9b20;border-radius:4px;padding:0 4px;font-size:14px;font-weight:900;line-height:20px}
.blank{min-height:176px}
@media(max-width:600px){.month .head{height:120px}.month .band{font-size:10px}.month .weekday-warning{font-size:15px;padding:0 3px}.month .blank{min-height:160px}}

.weekday-warning{display:inline-flex;align-items:center;justify-content:center;gap:3px;max-width:100%}
.duty-label{font-size:9px;font-weight:650;line-height:1.15;white-space:nowrap}
@media(max-width:600px){.month .weekday-warning{gap:1px;padding:0 1px;font-size:12px}.month .duty-label{font-size:8px;white-space:normal;max-width:30px;overflow-wrap:anywhere}}

/* 日付の文字・休日色・警告の有無によらず各段の位置を固定する。 */
.head,.month .head{
 display:grid;
 grid-template-columns:minmax(0,1fr);
 grid-template-rows:26px 30px 74px;
 justify-content:stretch;
 width:100%;
 min-width:0;
 align-content:end;
 align-items:center;
 justify-items:center;
 gap:4px;
 padding-top:0;
 padding-bottom:0;
}
.head > .day{line-height:24px;margin:0;align-self:center}
.head > .warning-slot{height:30px;min-height:30px;width:100%;margin:0}
.head > .state{height:74px;min-height:74px;margin:0;align-self:end}
.head,.month .head{height:142px}
.cell{padding-bottom:6px}.blank{min-height:158px !important}
@media(max-width:600px){.month .head{height:138px}.month .band{font-size:10px;letter-spacing:-.02em}.month .band .mark{display:none}}
@media(max-width:600px){
 .month .head{grid-template-rows:22px 30px 74px}
 .month .blank{min-height:146px !important}
}
/* 休日はNGの勤務帯をすべて同じ赤にする。平日の宿NGは従来の青。 */
.cell[data-holiday="true"][data-state="日NG"] .day-band,
.cell[data-holiday="true"][data-state="宿NG"] .night-band,
.cell[data-holiday="true"][data-state="全NG"] .band{background:#b42332;color:#fff}
</style><body><div id="strip" class="strip" aria-label="NG日カレンダー"></div><div id="note" class="note">左右へスクロールして選択できます。</div><button id="apply" type="button">NG日を反映する</button><script>
let version=null,days=[],draft=[];let mode='row';let frameHeight=0;
function send(type,data={}){window.parent.postMessage({isStreamlitMessage:true,type,...data},'*')}
function resize(){let h=document.body.scrollHeight+10;if(h!==frameHeight){frameHeight=h;send('streamlit:setFrameHeight',{height:h})}}
function paint(cell,value){
  cell.dataset.state=value;
  let holiday=cell.dataset.holiday==='true';
  let dayNg=value==='全NG'||value==='日NG', nightNg=value==='全NG'||value==='宿NG';
  let db=cell.querySelector('.day-band'), nb=cell.querySelector('.night-band');
  if(holiday){db.innerHTML=dayNg?'<span class="mark">✕</span>日直NG':'<span class="mark">○</span>日直 可';db.setAttribute('aria-pressed',String(dayNg));db.classList.toggle('ok',!dayNg);}
  else{db.textContent='日直なし';db.classList.remove('ok');}
  nb.innerHTML=nightNg?'<span class="mark">✕</span>宿直NG':'<span class="mark">○</span>宿直 可';nb.setAttribute('aria-pressed',String(nightNg));nb.classList.toggle('ok',!nightNg);
}
function combine(dayNg,nightNg){return dayNg&&nightNg?'全NG':dayNg?'日NG':nightNg?'宿NG':'OK'}
window.addEventListener('message',(event)=>{
  if(!event.data||event.data.type!=='streamlit:render')return;
  let a=event.data.args;
  document.getElementById('apply').textContent=a.button_label||(a.doctor+'先生のNG日を反映する');
  if(a.version!==version){
    version=a.version;mode=a.mode||'row';days=a.days;draft=days.map(d=>d.value);
    let strip=document.getElementById('strip');
    strip.replaceChildren();
    strip.className='strip '+(mode==='month'?'month':'');
    if(mode==='month'){
      ['月','火','水','木','金','土','日'].forEach((w,i)=>{let el=document.createElement('div');el.className='weekday';el.textContent=w;el.style.color=i===6?'#c92336':i===5?'#1670c5':'#243247';strip.append(el)});
      for(let i=0;i<a.offset;i++){let el=document.createElement('div');el.className='blank';strip.append(el)}
    }
    days.forEach((d,i)=>{
      let cell=document.createElement('div');cell.className='cell';cell.dataset.holiday=String(d.options.includes('日NG'));
      let head=document.createElement('div');head.className='head';
      let day=document.createElement('div');day.className='day '+d.kind;day.textContent=mode==='month'?d.day+'日':d.day+'日（'+d.weekday+'）';
      let state=document.createElement('div');state.className='state';
      let dayBand=document.createElement('div');dayBand.className='band day-band'+(d.options.includes('日NG')?'':' off');
      let nightBand=document.createElement('div');nightBand.className='band night-band';
      state.append(dayBand,nightBand);
      let warningSlot=document.createElement('div');warningSlot.className='warning-slot';
      if(d.warning){
        let warning=document.createElement('span');warning.className='weekday-warning';warning.textContent='⚠︎';
        let duty=document.createElement('small');duty.className='duty-label';duty.textContent='翌日PM duty';
        warning.append(duty);warning.title='翌日PM duty（翌日は平日）';warning.setAttribute('aria-label',warning.title);
        warningSlot.append(warning);
      }
      head.append(day,warningSlot,state);
      // 日直・宿直の段を押すと、その勤務だけ 可⇔NG が切り替わる（休日は両方を別々に指定できる）
      const toggle=(which)=>{
        let v=draft[i];let dayNg=v==='全NG'||v==='日NG',nightNg=v==='全NG'||v==='宿NG';
        if(which==='day')dayNg=!dayNg;else nightNg=!nightNg;
        let nv=combine(dayNg,nightNg);if(!d.options.includes(nv))return;
        draft[i]=nv;paint(cell,nv);
        document.getElementById('note').textContent=a.dirty_note||'まだ反映していない変更があります。選び終わったら下のボタンを押してください。';
      };
      const tappable=(el,which,label)=>{el.classList.add('tap');el.tabIndex=0;el.setAttribute('role','button');
        el.setAttribute('aria-label',d.day+'日の'+label);el.onclick=()=>toggle(which);
        el.onkeydown=(e)=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();toggle(which)}}};
      if(d.options.includes('日NG'))tappable(dayBand,'day','日直');
      tappable(nightBand,'night','宿直');
      cell.append(head);paint(cell,d.value);strip.append(cell);
    });
    if(mode==='month'){let blanks=(7-(a.offset+days.length)%7)%7;for(let i=0;i<blanks;i++){let el=document.createElement('div');el.className='blank';strip.append(el)}}
    document.getElementById('note').textContent='「日直」「宿直」の欄を押すと、可⇔NGが切り替わります（休日は日直と宿直を別々に指定します）。選び終わったら下のボタンを押してください。';
  }
  resize();
});
document.getElementById('apply').onclick=()=>{send('streamlit:setComponentValue',{value:{version,values:draft,token:Date.now().toString()+'-'+Math.random()},dataType:'json'});document.getElementById('note').textContent='処理しています…';};
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
# 右端の回数の列（宿直・日直・病棟・外来・平日・休日・合計）
MATRIX_WARD_SHIFTS = ('A日直', 'B日直', 'A宿直', 'B宿直')
MATRIX_OUTPATIENT_SHIFTS = ('外来宿直', '外来日直')
MATRIX_COUNT_COLUMNS = [('night', '宿直'), ('day', '日直'), ('ward', '病棟'), ('outpatient', '外来'), ('weekday', '平日'), ('holiday', '休日'), ('total', '合計')]
# 2列ずつのセットがひと目でわかるように色分けする（見出しの色, 数字欄の色）
MATRIX_COUNT_GROUP_COLORS = {
    'night': ('#cfe0f5', '#eef4fc'), 'day': ('#cfe0f5', '#eef4fc'),            # 宿直・日直：青
    'ward': ('#d5ead0', '#eef7ec'), 'outpatient': ('#d5ead0', '#eef7ec'),      # 病棟・外来：緑
    'weekday': ('#f8dcc4', '#fdf1e7'), 'holiday': ('#f8dcc4', '#fdf1e7'),      # 平日・休日：橙
    'total': ('#d9dde3', '#f3f4f6'),                                           # 合計：灰
}
MATRIX_COUNT_GROUP_STARTS = ('night', 'ward', 'weekday', 'total')  # セットの左端（区切り線を太くする）
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
        days.append({'day': d, 'weekday': WEEKDAYS_JA[dt.weekday()], 'holiday': hol, 'kind': kind, 'shortage': shortage})

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
            if kind == '全NG' or (kind == '宿NG' and not info['holiday']):
                cells.append({'type': 'ng', 'text': '×'})
            elif kind == '宿NG':
                cells.append({'type': 'ng', 'text': '宿×'})
            elif kind == '日NG' and info['holiday']:
                # 日直NGに加えて翌日PM dutyで宿直も外れる日は、両方わかるように表示
                cells.append({'type': 'ng', 'text': '日×', 'pm': pm})
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


def memo_cell_html(row):
    """メモ欄：ふだんは短く表示し、押すと全文が開く（スマートフォンでも使える）。"""
    if not row.get("comment") and not row.get("note"):
        return '<td class="memo"></td>'
    parts = []
    if row.get("comment"):
        parts.append(f'<b>💬 コメント</b>（{html_escape(row.get("comment_at", ""))}）<br>' + html_escape(row["comment"]).replace("\n", "<br>"))
    if row.get("note"):
        parts.append('<b>📝 備考</b><br>' + html_escape(row["note"]).replace("\n", "<br>"))
    # ポップアップ用の全文（カーソルを合わせると表示される）
    tip = []
    if row.get("comment"):
        tip.append(f"💬 コメント（{row.get('comment_at', '')}）\n{row['comment']}")
    if row.get("note"):
        tip.append(f"📝 備考\n{row['note']}")
    tip_text = html_escape("\n\n".join(tip)).replace("\n", "&#10;")
    # パソコン：カーソルを合わせるとポップアップ。スマートフォン（カーソルがない）：タップでその場に開く
    return (f'<td class="memo"><details><summary title="{tip_text}">{html_escape(memo_preview(row))}</summary>'
            f'<div class="full">{"<br><br>".join(parts)}</div></details></td>')


def duty_matrix_html(days, rows, year, month):
    def head_color(info):
        return MATRIX_HOLIDAY_HEAD if info['kind'] == 'holiday' else (MATRIX_SATURDAY_HEAD if info['kind'] == 'saturday' else '#243247')

    css = """
<style>
.dm-wrap{overflow-x:auto;border:1px solid #cfd6e0;border-radius:6px;margin:4px 0 8px}
.dm{border-collapse:collapse;font:12px system-ui,sans-serif;white-space:nowrap;background:#fff;color:#243247}
.dm th,.dm td{border:1px solid #d8dde5;padding:0;text-align:center;height:26px}
.dm thead th{background:#f3f5f8;font-weight:700;padding:2px 3px;min-width:30px}
.dm .nm{position:sticky;left:0;background:#fff;text-align:left;padding:0 8px;font-weight:600;z-index:1;min-width:90px}
.dm thead .nm{background:#f3f5f8;z-index:2}
.dm .pmcol{padding:0 6px;color:#566174}
.dm td.hol{background:#fdf1f1}
.dm .cell{display:flex;align-items:center;justify-content:center;height:100%;min-height:26px;font-weight:700;font-size:10px;padding:1px 2px;width:38px;margin:0 auto;white-space:normal;line-height:1.15;text-align:center;word-break:keep-all}
.dm .ng{color:#c0392b;font-size:14px}
.dm .pm{background:repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px)}
.dm tr.short.first td{border-top:2px solid #c0392b}
.dm tr.short td.nm{color:#c0392b}
.dm th.miss{background:#fde2e0}
.dm .sum{padding:0 6px;font-weight:700;min-width:34px}
.dm .memo{min-width:56px;text-align:left;padding:0 4px;font-size:11px;vertical-align:top}
.dm .memo details summary{cursor:pointer;list-style:none;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:116px;color:#425268}
.dm .memo details summary::-webkit-details-marker{display:none}
.dm .memo details[open] summary{color:#1f3c74;font-weight:700}
.dm .memo .full{white-space:normal;width:220px;padding:4px 6px;margin:2px 0 4px;background:#f6f8fb;border:1px solid #d8dde5;border-radius:4px;line-height:1.45}
@media (hover:hover){.dm .memo details .full{display:none}.dm .memo details summary{cursor:help}}
.dm-legend{display:flex;flex-wrap:wrap;gap:10px;font-size:12px;margin:2px 0 6px;color:#243247}
.dm-legend span{display:inline-flex;align-items:center;gap:4px}
.dm-legend i{display:inline-block;min-width:34px;padding:1px 4px;border:1px solid #d8dde5;border-radius:3px;font-style:normal;font-weight:700;text-align:center;font-size:11px}
</style>"""
    legend = ['<div class="dm-legend">']
    for s, (bg, fg) in MATRIX_SHIFT_COLORS.items():
        legend.append(f'<span><i style="background:{bg};color:{fg}">{s}</i></span>')
    legend.append(f'<span><i style="box-shadow:inset 0 0 0 3px {MATRIX_WISH_FRAME}">&nbsp;&nbsp;&nbsp;</i>茶色の枠：希望日が通った当直</span>')
    legend.append('<span><i style="color:#c0392b">×</i>当直NG（宿×＝宿直NG、日×＝日直NG）</span>')
    legend.append('<span><i>💬📝</i>メモ欄：先生からのコメント・備考（カーソルを合わせると全文）</span>')
    legend.append('<span><i style="background:repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px)">&nbsp;</i>翌日PM duty（宿直を外す日）</span>')
    legend.append('<span><i style="color:#c0392b;border-color:#c0392b">⚠️ 未決定</i>赤線の下：当直が決まっていない枠</span>')
    legend.append('</div>')

    head = ['<tr><th class="nm">名前</th><th class="pmcol">翌日PM duty</th><th class="memo">メモ</th>']
    for info in days:
        miss_class = ' class="miss"' if info['shortage'] else ''
        head.append(f'<th{miss_class} style="color:{head_color(info)}">{info["day"]}<br>{info["weekday"]}</th>')
    head.append(''.join(f'<th class="sum" style="{count_cell_css(key, True)}">{label}</th>' for key, label in MATRIX_COUNT_COLUMNS) + '</tr>')

    body = []
    for row in rows:
        tr = [f'<tr><td class="nm">{html_escape(row["name"])}</td><td class="pmcol">{html_escape(row["pm_duty"])}</td>{memo_cell_html(row)}']
        for info, cell in zip(days, row['cells']):
            td_class = ' class="hol"' if info['holiday'] else ''
            if cell['type'] == 'shift':
                bg, fg = MATRIX_SHIFT_COLORS.get(cell['shift'], ('#eee', '#000'))
                wish_css = f'box-shadow:inset 0 0 0 3px {MATRIX_WISH_FRAME},inset 0 0 0 4px #ffffff;' if cell.get('wish') else ''
                tr.append(f'<td><div class="cell" style="background:{bg};color:{fg};{wish_css}">{slot_label(cell["text"], "<br>")}</div></td>')
            elif cell['type'] == 'ng':
                pm_class = ' pm' if cell.get('pm') else ''
                tr.append(f'<td{td_class}><div class="cell ng{pm_class}">{cell["text"]}</div></td>')
            elif cell['type'] == 'pm':
                tr.append('<td><div class="cell pm"></div></td>')
            else:
                tr.append(f'<td{td_class}></td>')
        tr.append(''.join(f'<td class="sum" style="{count_cell_css(key, False)}">{row[key]}</td>' for key, _ in MATRIX_COUNT_COLUMNS) + '</tr>')
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
                tr.append(f'<td><div class="cell" style="background:{bg};color:{fg};white-space:normal;line-height:1.15;font-size:10px">{label}</div></td>')
            else:
                tr.append('<td></td>')
        tr.append(f'<td class="sum" colspan="{len(MATRIX_COUNT_COLUMNS)}" style="color:#c0392b">{slot_total}枠</td></tr>')
        body.append(''.join(tr))
    return css + ''.join(legend) + f'<div class="dm-wrap"><table class="dm"><thead>{"".join(head)}</thead><tbody>{"".join(body)}</tbody></table></div>'


def html_escape(text):
    import html
    return html.escape(str(text))


def duty_matrix_xlsx(days, rows, year, month):
    """一覧表をExcelファイル（印刷向けの書式付き）にする。"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.comments import Comment
    from openpyxl.utils import get_column_letter

    wb = Workbook()
    ws = wb.active
    ws.title = f"{year}年{month}月"
    thin = Side(style='thin', color='B8C0CC')
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    center = Alignment(horizontal='center', vertical='center', wrap_text=True)
    def fill(color):
        return PatternFill('solid', fgColor=color.lstrip('#'))
    PM_FILL = PatternFill('lightUp', fgColor='E0A800', bgColor='FFF5D9')

    ws.cell(row=1, column=1, value=f"{year}年{month}月分 当直表").font = Font(bold=True, size=14)
    header_row, first_day_col = 3, 4
    ws.cell(row=header_row, column=1, value='名前')
    ws.cell(row=header_row, column=2, value='翌日PM duty')
    ws.cell(row=header_row, column=3, value='メモ')
    for i, info in enumerate(days):
        c = ws.cell(row=header_row, column=first_day_col + i, value=f"{info['day']}\n{info['weekday']}")
        color = MATRIX_HOLIDAY_HEAD if info['kind'] == 'holiday' else (MATRIX_SATURDAY_HEAD if info['kind'] == 'saturday' else '#243247')
        c.font = Font(bold=True, color=color.lstrip('#'))
    sum_col = first_day_col + len(days)
    n_counts = len(MATRIX_COUNT_COLUMNS)
    for j, (_, label) in enumerate(MATRIX_COUNT_COLUMNS):
        ws.cell(row=header_row, column=sum_col + j, value=label)
    for col in range(1, sum_col + n_counts):
        c = ws.cell(row=header_row, column=col)
        c.alignment = center; c.border = border; c.fill = fill('#E6EBF2')
        if c.font is None or not c.font.bold:
            c.font = Font(bold=True)
    group_side = Side(style='medium', color='7B8698')
    wish_side = Side(style='thick', color=MATRIX_WISH_FRAME.lstrip('#').upper())
    wish_border = Border(left=wish_side, right=wish_side, top=wish_side, bottom=wish_side)

    def style_count_cell(c, key, is_head):
        head_bg, body_bg = MATRIX_COUNT_GROUP_COLORS[key]
        c.fill = fill(head_bg if is_head else body_bg)
        if key in MATRIX_COUNT_GROUP_STARTS:
            c.border = Border(left=group_side, right=thin, top=c.border.top, bottom=c.border.bottom)

    for j, (key, _) in enumerate(MATRIX_COUNT_COLUMNS):
        style_count_cell(ws.cell(row=header_row, column=sum_col + j), key, True)

    for r_idx, row in enumerate(rows, header_row + 1):
        ws.cell(row=r_idx, column=1, value=row['name']).font = Font(bold=True)
        ws.cell(row=r_idx, column=2, value=row['pm_duty'])
        if row.get('comment') or row.get('note'):
            # セルには短く表示し、全文はセルのコメント（マウスを乗せると表示）と「メモ全文」シートに入れる
            memo = ws.cell(row=r_idx, column=3, value=memo_preview(row, length=4))
            memo.font = Font(size=8)
            full = []
            if row.get('comment'):
                full.append(f"【コメント】{row['comment']}")
            if row.get('note'):
                full.append(f"【備考】{row['note']}")
            memo.comment = Comment("\n".join(full), "当直作成アプリ", width=260, height=120)
        for i, (info, cell) in enumerate(zip(days, row['cells'])):
            c = ws.cell(row=r_idx, column=first_day_col + i, value=slot_label(cell['text'], '\n') or None)
            if cell['type'] == 'shift':
                bg, fg = MATRIX_SHIFT_COLORS.get(cell['shift'], ('#EEEEEE', '#000000'))
                c.fill = fill(bg); c.font = Font(bold=True, size=8, color=fg.lstrip('#'))
            elif cell['type'] == 'ng':
                c.font = Font(bold=True, color='C0392B')
                if cell.get('pm'):
                    c.fill = PM_FILL
                elif info['holiday']:
                    c.fill = fill('#FDF1F1')
            elif cell['type'] == 'pm':
                c.fill = PM_FILL
            elif info['holiday']:
                c.fill = fill('#FDF1F1')
        for j, (key, _) in enumerate(MATRIX_COUNT_COLUMNS):
            ws.cell(row=r_idx, column=sum_col + j, value=row[key]).font = Font(bold=True)
        for col in range(1, sum_col + n_counts):
            c = ws.cell(row=r_idx, column=col)
            c.border = border
            c.alignment = center if col > 1 else Alignment(vertical='center')
        # メモ欄は1行で表示（折り返すと行が高くなるため、はみ出す分は縮小）
        ws.cell(row=r_idx, column=3).alignment = Alignment(horizontal='left', vertical='center', shrink_to_fit=True)
        # 希望日が通った当直をオレンジの太線で囲む
        for i, cell in enumerate(row['cells']):
            if cell.get('wish'):
                ws.cell(row=r_idx, column=first_day_col + i).border = wish_border
        for j, (key, _) in enumerate(MATRIX_COUNT_COLUMNS):
            style_count_cell(ws.cell(row=r_idx, column=sum_col + j), key, False)
        # 同じ日に2つ当直がある医師の行は、2段で表示できる高さにする
        max_lines = max([slot_label(cell['text'], '\n').count('\n') + 1 for cell in row['cells'] if cell['type'] == 'shift'] or [1])
        ws.row_dimensions[r_idx].height = max(20, 11 * max_lines + 6)

    # 当直が決まっていない枠：枠ごとに1行、その枠の色で表示
    last_row = header_row + len(rows)
    red = Side(style='medium', color='C0392B')
    for k, slot in enumerate(shortage_slots(days)):
        last_row += 1
        bg, fg = MATRIX_SHIFT_COLORS[slot]
        ws.cell(row=last_row, column=1, value='⚠️ 未決定').font = Font(bold=True, color='C0392B')
        slot_total = 0
        for i, info in enumerate(days):
            count = sum(x['count'] for x in info['shortage'] if x['slot'] == slot)
            c = ws.cell(row=last_row, column=first_day_col + i)
            c.border = Border(left=thin, right=thin, bottom=thin, top=red if k == 0 else thin)
            if count:
                slot_total += 1
                c.value = slot_label(slot, '\n') + ('' if count == 1 else f'×{count}')
                c.fill = fill(bg); c.font = Font(bold=True, size=8, color=fg.lstrip('#'))
                ws.cell(row=header_row, column=first_day_col + i).fill = fill('#FDE2E0')
            c.alignment = center
        ws.cell(row=last_row, column=sum_col + n_counts - 1, value=f"{slot_total}枠").font = Font(bold=True, color='C0392B')
        for col in list(range(1, first_day_col)) + list(range(sum_col, sum_col + n_counts)):
            c = ws.cell(row=last_row, column=col)
            c.border = Border(left=thin, right=thin, bottom=thin, top=red if k == 0 else thin)
            c.alignment = center if col > 1 else Alignment(vertical='center')
        ws.row_dimensions[last_row].height = 24

    legend_row = last_row + 2
    ws.cell(row=legend_row, column=1, value='凡例').font = Font(bold=True)
    items = [(s, bg, fg) for s, (bg, fg) in MATRIX_SHIFT_COLORS.items()]
    for k, (label, bg, fg) in enumerate(items):
        c = ws.cell(row=legend_row, column=2 + k * 2, value=label)
        c.fill = fill(bg); c.font = Font(bold=True, size=8, color=fg.lstrip('#')); c.alignment = center; c.border = border
    ws.cell(row=legend_row + 1, column=2, value='×').font = Font(bold=True, color='C0392B')
    wish_legend = ws.cell(row=legend_row + 4, column=2)
    wish_legend.border = wish_border; wish_legend.alignment = center; wish_legend.font = Font(bold=True)
    ws.cell(row=legend_row + 4, column=3, value='茶色の枠：希望日が通った当直')
    ws.cell(row=legend_row + 1, column=3, value='当直NG（宿×＝宿直NG、日×＝日直NG）')
    ws.cell(row=legend_row + 2, column=2).fill = PM_FILL
    ws.cell(row=legend_row + 2, column=3, value='翌日PM duty（宿直を外す日）')
    ws.cell(row=legend_row + 3, column=2, value='⚠️ 未決定').font = Font(bold=True, color='C0392B')
    ws.cell(row=legend_row + 3, column=3, value='赤線の下：当直が決まっていない枠')

    ws.column_dimensions['A'].width = 14
    ws.column_dimensions['B'].width = 12
    ws.column_dimensions['C'].width = 11
    for i in range(len(days)):
        ws.column_dimensions[get_column_letter(first_day_col + i)].width = 5.2
    for j in range(n_counts):
        ws.column_dimensions[get_column_letter(sum_col + j)].width = 5
    ws.row_dimensions[header_row].height = 30
    ws.freeze_panes = ws.cell(row=header_row + 1, column=first_day_col)
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
        ws2.cell(row=1, column=1, value=f"{year}年{month}月 先生からのコメント・備考").font = Font(bold=True, size=13)
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
    for col in [NG_COLUMN, REQUEST_COLUMN, '翌日PM duty']:
        if col not in df:
            df[col] = ''
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
            errors.append(f'確定当直 {i+1}行目：日付「{raw}」を確認してください。年付きの年月日で指定できます。')
            continue
        df.at[i, '日付'] = dt.isoformat()
        for s in ALL_SHIFT_TYPES:
            val = clean_text(row.get(s, ''))
            if val in ('', '-'):
                continue
            if (holidays is not None and s in DAY_SHIFTS and (dt.year, dt.month) == (year, month)
                    and not is_holiday_date(dt, year, month, holidays)):
                errors.append(f'確定当直 {dt} {s}：平日には日直がありません。この日に日直を設ける場合は、上のカレンダーで「休日にする」にチェックを入れてください。')
            names = list(dict.fromkeys(n.strip() for n in re.split('[、,]', val)))
            for name in names:
                if name not in doctors:
                    errors.append(f'確定当直 {dt} {s}：医師名「{name}」が名簿と一致しません。')
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
                    warnings.append(f'{name}：{dt} {slot}はNG指定より確定指定を優先しました。')
                if night and pm_duty_restricted(dt, pm_weekdays, year, month, holidays, next_month_special_holiday):
                    warnings.append(f'{name}：{dt} {slot}は翌日PM dutyの制限より確定指定を優先しました。')
        total = sum(counts.values())
        if total < int(row['月間最小回数']):
            warnings.append(f'{name}：月間最小回数の目標{int(row["月間最小回数"])}回に対し、{total}回です。')
        for label, value, cap in [('月間最大回数', total, int(row['月間最大回数'])), ('休日最大回数', hol_count, int(row['休日最大回数']))] + [(s + '上限', counts[s], int(row[s + '上限'])) for s in ALL_SHIFT_TYPES]:
            if value > cap:
                warnings.append(f'{name}：確定指定に伴い{label}を超えています（設定{cap}回／結果{value}回）。')
        dates = sorted(set(worked) | set((past or {}).get(name, [])) | set((future or {}).get(name, [])))
        for a, b in zip(dates, dates[1:]):
            if (a.year, a.month) != (year, month) and (b.year, b.month) != (year, month): continue
            gap = (b - a).days - 1
            if gap < int(row['最低空ける日数']):
                warnings.append(f'{name}：確定勤務同士の間隔が設定未満です（{a}→{b}、空き{gap}日／設定{int(row["最低空ける日数"])}日）。確定勤務を維持しました。')
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


def _generate_shift_core(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None):
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
        req_priority[doc] = safe_int(row.get('希望優先度(数字が大きいほど優先)'), 1)
        min_intervals[doc] = safe_int(row.get('最低空ける日数'), 5)
        min_shifts_total[doc] = safe_int(row.get('月間最小回数'), 0)
        max_shifts_total[doc] = safe_int(row.get('月間最大回数'), 5)
        max_hol_shifts_per_doc[doc] = safe_int(row.get('休日最大回数'), 4)
        max_shifts_per_type[doc] = {s: safe_int(row.get(s + '上限'), 2) for s in ALL_SHIFT_TYPES}

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
            warnings.append("⚠️ **【重要】以下の枠は「決定済み当直」や「優先度100」が重なったため、AIが自動的に定員を拡張（2名以上配置）して当直を完成させました:**")
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


def generate_shift(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None):
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
                errors.append(f'{d}日の日直増員：先に特別休日を指定してください。')
    if errors: return None, False, errors, None, None
    result, success, warnings, past, future = _generate_shift_core(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority)
    if result is not None:
        warnings += audit_schedule(result, staff_df, target_year, target_month, custom_holidays, multi_slots_dict, past, future, next_month_special_holiday)
        success = not any('⚠️不足' in str(v) for s in ALL_SHIFT_TYPES for v in result[s])
    return result, success, list(dict.fromkeys(warnings)), past, future


def show_input_errors(messages):
    messages = list(dict.fromkeys(str(m).strip() for m in messages if str(m).strip()))
    st.error(f"入力内容を確認してください（{len(messages)}件）")
    st.dataframe(pd.DataFrame({"修正が必要な項目": messages}), hide_index=True, use_container_width=True)


def show_schedule_notice(result, messages):
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
        st.warning(f"不足が残っています：{len(shortage_rows)}枠、合計{sum(shortage_rows)}名分")
    else:
        st.success("必要人数を満たす当直案ができました。")

    details = []
    for message in messages:
        if message.startswith(('🚨 **以下の枠', '📊 **【不足している枠')):
            continue
        if re.fullmatch(r'・\d+/\d+ の「.+」', message) or re.fullmatch(r'・.+： 計 \d+ 枠不足', message):
            continue
        details.append(message)
    if details or result is None or shortage_rows:
        with st.container(border=True):
            st.markdown("**計算状況・条件の注意事項と対処方法**")
            for message in details:
                st.markdown(message)
            if result is None or shortage_rows:
                st.markdown("**条件を見直す場合**")
                st.write("入力内容を確認し、実際に調整できる範囲で、月間・休日・枠別の上限や勤務間隔を見直して再作成してください。")
            if shortage_rows:
                st.write("不足箇所は当直案の「⚠️不足」で確認できます。手動調整する場合はCSVをダウンロードし、Excelなどで編集してください。")


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
    "submissions": ["年月", "先生", "NG日", "保存日時", "コメント"],
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
    先生ごとの専用URLを、A4・1人1枚の三つ折り用紙にした印刷用HTMLを作る。
    片面印刷のまま、下1/3を内側に折り（谷折り）、上1/3を裏側に折る（山折り）と、
    宛名だけが外に出て、QRコードと説明は内側に隠れる。左右をホチキスで留めて配る。
    cards: [(先生, 専用URL)]
    """
    import html as _html
    pages = []
    for name, url in cards:
        n = _html.escape(name)
        pages.append(f"""
<section class="sheet">
  <div class="panel cover">
    <div class="cover-title">当直NG日 入力用カード</div>
    <div class="cover-name">{n} 先生</div>
    <div class="staple left"><svg class="staple-svg" viewBox="0 0 20 60" xmlns="http://www.w3.org/2000/svg" aria-label="ホチキス止めの位置"><path d="M15 6 H6 Q4 6 4 8 V52 Q4 54 6 54 H15" fill="none" stroke="#6b7380" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"/><path d="M5.2 10 V50" stroke="#d5dae1" stroke-width="1" stroke-linecap="round"/></svg></div><div class="staple right"><svg class="staple-svg" viewBox="0 0 20 60" xmlns="http://www.w3.org/2000/svg" aria-label="ホチキス止めの位置"><path d="M15 6 H6 Q4 6 4 8 V52 Q4 54 6 54 H15" fill="none" stroke="#6b7380" stroke-width="3.2" stroke-linecap="round" stroke-linejoin="round"/><path d="M5.2 10 V50" stroke="#d5dae1" stroke-width="1" stroke-linecap="round"/></svg></div>
  </div>
  <div class="fold"><span>② ここを山折り（上の部分を裏側へ）</span></div>
  <div class="panel qr">
    <div class="qr-name">{n} 先生 専用</div>
    <div class="qr-body">
      {qr_svg(url, size_mm=58)}
      <div class="qr-side">
        <div class="big">スマートフォンの<br>カメラで読み取る</div>
        <div class="small">QRコードにカメラを向けると、<br>NG日の入力ページが開きます。</div>
        <div class="wake"><b>英語の画面（Zzzz…）が出たら</b><br>青いボタン「<b>Yes, get this app back up!</b>」を押して、<br>1分ほどお待ちください。</div>
      </div>
    </div>
    <div class="url">{_html.escape(url)}</div>
  </div>
  <div class="fold"><span>① ここを谷折り（下の部分を内側へ）</span></div>
  <div class="panel howto">
    <div class="howto-title">入力のしかた</div>
    <ol>
      <li>QRコードを読み取り、入力ページを開く</li>
      <li>入れない日の「日直」「宿直」の欄を押して、NGにする</li>
      <li>「NG日とコメントを保存する」を押す</li>
      <li>「✅ 保存しました」が出れば完了</li>
    </ol>
    <div class="change">休日は「日直」と「宿直」を<b>別々に</b>指定します。終日入れない日は<b>両方</b>押してください。</div>
    <ul>
      <li><b>このURLは毎月同じです。</b>ブックマークしておくと便利です。</li>
      <li>締め切りまでは何度でも修正でき、<b>最後に保存した内容</b>が採用されます。</li>
      <li>このカードは<b>先生専用</b>です。他の方に渡さないでください。なくした場合は担当者へご連絡ください。</li>
    </ul>
  </div>
</section>""")
    return f"""<!doctype html><html lang="ja"><head><meta charset="utf-8"><title>当直NG日入力 配布用カード</title>
<style>
@page {{ size: A4 portrait; margin: 0; }}
* {{ box-sizing: border-box; }}
body {{ margin: 0; font-family: "Hiragino Kaku Gothic ProN", "Yu Gothic", "Meiryo", sans-serif; color: #1f2a37; background: #e9edf2; }}
.bar {{ padding: 12px 16px; background: #eef3f8; border-bottom: 1px solid #cfd8e3; font-size: 14px; line-height: 1.6; }}
.bar button {{ font-size: 15px; padding: 8px 18px; margin-right: 12px; cursor: pointer; }}
.sheet {{ width: 210mm; height: 297mm; margin: 8mm auto; background: #fff; display: flex; flex-direction: column;
         box-shadow: 0 1px 4px rgba(0,0,0,.15); break-after: page; page-break-after: always; overflow: hidden; }}
.panel {{ height: 98mm; padding: 10mm 16mm; position: relative; }}
.fold {{ height: 1.5mm; border-top: 1.2px dashed #8a94a6; position: relative; }}
.fold span {{ position: absolute; right: 12mm; top: -3.2mm; background: #fff; padding: 0 2mm; font-size: 9pt; color: #6b7686; }}
.cover {{ display: flex; flex-direction: column; align-items: center; justify-content: center; text-align: center; }}
.seal {{ border: 2.5px solid #b42332; color: #b42332; font-weight: 700; font-size: 20pt; padding: 1mm 6mm; border-radius: 2mm; letter-spacing: .3em; }}
.cover-title {{ margin-top: 6mm; font-size: 14pt; color: #4a5568; }}
.cover-name {{ margin-top: 3mm; font-size: 30pt; font-weight: 700; }}
.cover-note {{ margin-top: 4mm; font-size: 11pt; color: #b42332; }}
/* ホチキスは左右の端・高さの真ん中に縦向きで留める。
   折りたたむと、内側のQRコード・URL・説明文のどれにも重ならない位置（端から約5〜10mm。印刷できない端の余白も避ける） */
.staple {{ position: absolute; top: 50%; transform: translateY(-50%); width: 5mm; height: 18mm; }}
.staple.right svg {{ transform: scaleX(-1); }}
.staple-svg {{ width: 100%; height: 100%; display: block; }}
.staple.left {{ left: 5mm; }} .staple.right {{ right: 5mm; }}
.qr-name {{ font-size: 15pt; font-weight: 700; margin-bottom: 4mm; text-align: center; }}
.qr-body {{ display: flex; align-items: center; justify-content: center; gap: 7mm; }}
.qr-side {{ max-width: 92mm; }}
.qr-side .big {{ font-size: 17pt; font-weight: 700; line-height: 1.4; }}
.qr-side .small {{ margin-top: 4mm; font-size: 11.5pt; line-height: 1.6; color: #374151; }}
.qr-side .wake {{ margin-top: 4mm; padding: 2mm 3mm; border: 1px solid #e3b341; background: #fff8e1; border-radius: 1.5mm;
                 font-size: 9.5pt; line-height: 1.55; color: #5b4300; }}
.url {{ position: absolute; left: 30mm; right: 30mm; bottom: 6mm; text-align: center; font-size: 7.5pt; color: #6b7280; word-break: break-all; font-family: monospace; }}
.howto-title {{ font-size: 15pt; font-weight: 700; margin-bottom: 2mm; }}
.howto ol {{ margin: 0 0 3mm 6mm; padding-left: 5mm; font-size: 12.5pt; line-height: 1.75; }}
.howto .change {{ margin: 0 0 2.5mm; padding: 1.5mm 3mm; border: 1.5px solid #e8590c; background: #fff4e6; border-radius: 1.5mm; font-size: 10.5pt; line-height: 1.5; color: #3b2a1a; }}
.howto ul {{ margin: 0; padding-left: 5mm; font-size: 10.5pt; line-height: 1.65; color: #374151; }}
@media print {{ body {{ background: #fff; }} .bar {{ display: none; }} .sheet {{ margin: 0; box-shadow: none; }} }}
</style></head><body>
<div class="bar"><button onclick="window.print()">🖨️ 印刷する</button>
A4・縦・<b>片面</b>で印刷してください（{len(cards)}名分・1人1枚）。印刷の設定で「倍率：100%（実際のサイズ）」「余白：なし」を選ぶと、折り線の位置が正しくなります。<br>
折り方：① 下の1/3を内側に折る（谷折り）→ ② 上の1/3（宛名）を裏側に折る（山折り）→ 宛名が表に出たら、左右の端の印の位置（2か所）をホチキスで留めます。</div>
{''.join(pages)}
</body></html>"""

def remote_latest_submissions(month_key):
    """その年月の、先生ごとの最新の提出 {先生: {"NG日":…, "保存日時":…}}"""
    latest = {}
    for r in remote_read("submissions"):
        if r.get("年月") == month_key and r.get("先生"):
            latest[r["先生"]] = r   # 後ろの行ほど新しい
    return latest


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


def ng_calendar_days(year, month, holidays, next_month_special_holiday, pm_weekdays, values_by_day):
    """NGカレンダー部品に渡す日ごとのデータを作る。"""
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
        })
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
const colors = {'全NG': ['#b42332', '全日NG'], '日NG': ['#9a4700', '日直NG'], '宿NG': ['#1856a4', '宿直NG']};
let count = 0;
D.days.forEach((day, i) => {
  const p = D.offset + i, col = p % 7, row = Math.floor(p / 7);
  const x = x0 + col * cw, y = y0 + 60 + row * cellH;
  g.fillStyle = day.hol ? '#fdf1f1' : '#ffffff'; g.fillRect(x, y, cw, cellH);
  g.strokeStyle = '#cbd2dc'; g.lineWidth = 2; g.strokeRect(x, y, cw, cellH);
  g.fillStyle = day.sat ? '#1670c5' : day.hol ? '#c92336' : '#243247';
  g.font = F(30, 700); g.fillText(String(day.d), x + 12, y + 40);
  const info = colors[day.ng];
  if (info) {
    count++;
    const color = (day.hol && day.ng !== 'OK') ? '#b42332' : info[0];
    g.fillStyle = color; g.fillRect(x + 8, y + 62, cw - 16, 60);
    g.fillStyle = '#fff'; g.font = F(26, 700); g.textAlign = 'center';
    g.fillText(info[1], x + cw / 2, y + 102); g.textAlign = 'left';
  }
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
def work_state_save(snapshot):
    text = json.dumps(snapshot, ensure_ascii=False)
    if len(text) > 45000:
        raise ValueError("保存する内容が大きすぎます（医師の人数や確定当直の行数を確認してください）。")
    remote_rewrite("saved_state", [["保存日時", remote_now()], ["内容", text]])


def work_state_info():
    """保存済みの作業 (保存日時, 内容) を返す。なければ (None, None)。"""
    try:
        rows = {r.get("項目", ""): r.get("値", "") for r in remote_read("saved_state")}
    except Exception:
        return None, None
    if not rows.get("内容"):
        return None, None
    try:
        return rows.get("保存日時", ""), json.loads(rows["内容"])
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
        if key.startswith(("ng_", "hol_", "next_month_special_holiday_", "shortage_priority_")):
            del st.session_state[key]
    st.session_state["sel_year"] = y
    st.session_state["sel_month"] = m
    for d in data.get("holidays", []):
        st.session_state[f"hol_{y}_{m}_{int(d)}"] = True
    st.session_state[f"next_month_special_holiday_{y}_{m}"] = bool(data.get("next_special"))
    for slot, rank in (data.get("priority") or {}).items():
        st.session_state[f"shortage_priority_{slot}"] = int(rank)
    st.session_state["restored_state"] = data
    # ファイルのアップロード欄を空にして、復元した内容を使う
    st.session_state["uploader_gen"] = st.session_state.get("uploader_gen", 0) + 1
    st.session_state.pop("last_uploaded_file_id", None)
    st.session_state["work_state_message"] = ("success", f"{saved_at} に保存した作業（{y}年{m}月）を読み込みました。")


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
    position: fixed; top: 0.45rem; right: 8.5rem; z-index: 1000002;
    width: auto !important; max-width: 60vw;
    display: flex !important; flex-direction: row !important; align-items: center; gap: 10px;
    background: transparent;
}
.st-key-work_save_float > div { width: auto !important; flex: 0 0 auto !important; }
.st-key-work_save_float button { min-height: 2.2rem; padding: 0.2rem 0.9rem; font-weight: 700; white-space: nowrap; }
.st-key-work_save_float p { margin: 0; white-space: nowrap; }
.work-save-note { font-size: 0.85rem; color: #1b7f3b; font-weight: 600; }
.work-save-note.busy { color: #566174; }
.work-save-note.bad { color: #c92336; }
@media (max-width: 640px) {
    .st-key-work_save_float { right: 5.5rem; top: 0.5rem; max-width: 75vw; }
    .st-key-work_save_float button { padding: 0.15rem 0.6rem; font-size: 0.8rem; }
    .work-save-note { display: none; }
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
    if msg:
        if msg[0] == "success":
            st.toast(msg[1], icon="📂")
        else:
            st.error(msg[1])
    has_popover = hasattr(st, "popover")
    with st.container(key="work_save_float"):
        # --- 前回の作業の続きから始める（押すと小さな枠が開く） ---
        box = st.popover("📂 前回の作業の続きから始める") if has_popover \
            else st.expander("📂 前回の作業の続きから始める", expanded=False)
        with box:
            if admin_logged_in("work_restore_pw"):
                saved_at, saved = work_state_info()
                if saved:
                    st.write(f"保存日時：**{saved_at}**")
                    st.write(f"{saved['year']}年{saved['month']}月の作業")
                    st.button("この作業を読み込む", key="work_restore", type="primary", on_click=work_state_restore)
                    st.caption("年月・特別休日・増員設定・確定当直・医師条件（NG日を含む）・不足時の優先順位が、"
                               "途中保存したときの状態に戻ります。今の画面の入力内容は置き換わります。")
                else:
                    st.caption("まだ途中保存された作業はありません。隣の「💾 作業を途中保存する」で保存できます。")
            else:
                st.caption("読み込みには管理用パスワードが必要です。")

        # --- 作業を途中保存する ---
        if st.session_state.get("remote_admin_ok"):
            if st.button("💾 作業を途中保存する", key="work_save_button", type="primary"):
                work_save_request = True
            WORK_SAVE_SLOT = st.empty()
            if work_save_request:
                WORK_SAVE_SLOT.markdown("<span class='work-save-note busy'>保存しています…</span>", unsafe_allow_html=True)
            elif msg and msg[0] == "success":
                WORK_SAVE_SLOT.markdown("<span class='work-save-note'>✅ 前回の作業を読み込みました</span>", unsafe_allow_html=True)
        else:
            # パスワード未入力のときは、押すと入力欄が開くボタンにする
            box = st.popover("💾 作業を途中保存する") if has_popover \
                else st.expander("💾 作業を途中保存する", expanded=False)
            with box:
                st.caption("途中保存には管理用パスワードが必要です。")
                admin_logged_in("work_save_pw")


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
                       "変更したい場合は、カレンダーで選び直して、もう一度保存してください。")
            # コメントは、アプリの案内文と混ざらないように別の枠で表示する
            comment = latest.get("コメント", "").strip()
            if comment and show_comment:
                with st.container(border=True):
                    st.markdown("💬 **保存されている管理者へのコメント**")
                    st.text(f"「{comment}」")
        else:
            st.warning("**まだ提出されていません。** 当直に入れない日を選び、保存ボタンを押してください。"
                       "NG日がない場合も、そのまま保存ボタンを押してください（「NG日なし」として提出されます）。")

    saved_status_box()

    with st.expander("NGの種類について", expanded=False):
        st.markdown("""
- **宿NG**：宿直（夜）に入れない日。平日はこれを選びます。
- **日NG**：日直（日中）に入れない日。土日祝だけ選べます。
- **全NG**：日直・宿直とも入れない日。土日祝だけ選べます。
- ⚠ の日は、翌日PMにdutyがあるため、原則として宿直を外します。
        """)

    # --- 管理者へのコメント（任意）。NG日の保存と一緒に送られる ---
    comment_key = f"remote_comment_{year}_{month}"
    if comment_key not in st.session_state:
        st.session_state[comment_key] = latest.get("コメント", "") if latest else ""
    st.text_area(
        "管理者へのコメント・要望（任意）",
        key=comment_key, max_chars=500, height=90,
        placeholder="例：第2週は学会のため、できれば平日の宿直を避けたいです。／ 月4回までを希望します。",
        help="カレンダーの下の保存ボタンを押すと、NG日と一緒に保存されます。",
    )
    st.caption("下のカレンダーで「日直」「宿直」の欄を押してNGにし（休日は日直と宿直が別々です）、カレンダーの下の「保存する」ボタンを押してください。コメントも一緒に保存されます。"
               "コメントだけを直したいときも、同じボタンを押してください（NG日はそのまま）。")

    # 以前の紙の用紙（休日は日直・宿直まとめて×）から変わった点を、カレンダーのすぐ上で強調する
    st.markdown("""
<div style="border:2px solid #e8590c;background:#fff4e6;border-radius:10px;padding:12px 14px;margin:6px 0 10px">
 <div style="font-weight:800;font-size:1.05rem;color:#c2410c">⚠ 休日は「日直」と「宿直」を別々に指定します</div>
 <div style="margin-top:4px;font-size:.95rem;color:#3b2a1a">終日入れない日は、<b>両方</b>押してください。</div>
 <div style="display:flex;flex-wrap:wrap;gap:10px;margin-top:10px;font-size:.85rem;font-weight:700">
  <div style="text-align:center;width:100px">
   <div style="width:96px;border:1px solid #d6dce5;border-radius:5px;overflow:hidden">
    <div style="background:#b42332;color:#fff;padding:5px 0">✕ 日直NG</div>
    <div style="background:#b42332;color:#fff;padding:5px 0;border-top:1px solid #fff">✕ 宿直NG</div></div>
   <div style="margin-top:4px;color:#3b2a1a">終日NG</div></div>
  <div style="text-align:center;width:100px">
   <div style="width:96px;border:1px solid #d6dce5;border-radius:5px;overflow:hidden">
    <div style="background:#e3f5e9;color:#12692f;padding:5px 0">○ 日直 可</div>
    <div style="background:#b42332;color:#fff;padding:5px 0">✕ 宿直NG</div></div>
   <div style="margin-top:4px;color:#3b2a1a">宿直だけNG<br><span style="color:#c2410c">日直は入れる</span></div></div>
  <div style="text-align:center;width:100px">
   <div style="width:96px;border:1px solid #d6dce5;border-radius:5px;overflow:hidden">
    <div style="background:#b42332;color:#fff;padding:5px 0">✕ 日直NG</div>
    <div style="background:#e3f5e9;color:#12692f;padding:5px 0">○ 宿直 可</div></div>
   <div style="margin-top:4px;color:#3b2a1a">日直だけNG<br><span style="color:#c2410c">宿直は入れる</span></div></div>
 </div>
</div>""", unsafe_allow_html=True)

    layout = st.radio("表示", ["月間カレンダー", "1日〜月末を横一列"], horizontal=True, key="remote_layout")
    pm_weekdays = pm_duty_weekdays(tokens[doctor].get("翌日PM duty", ""))
    values = parse_ng_dict(saved_text, year, month)
    days = ng_calendar_days(year, month, holidays, next_special, pm_weekdays, values)
    revision = hashlib.sha256(json.dumps([year, month, doctor, layout, days], ensure_ascii=False).encode()).hexdigest()
    component_key = f"remote_ng_{'row' if layout != '月間カレンダー' else 'month'}_{year}_{month}"
    response = horizontal_ng_component(HORIZONTAL_NG_HTML)(
        days=days, mode="row" if layout != "月間カレンダー" else "month",
        offset=datetime.date(year, month, 1).weekday(), doctor=doctor,
        button_label="NG日とコメントを保存する",
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
            merged = dict(values)
            for d, (v, item) in enumerate(zip(new_values, days), 1):
                if v != item["value"]:
                    if v == "OK":
                        merged.pop(d, None)
                    else:
                        merged[d] = v
            try:
                now = remote_now()
                remote_append("submissions", [f"{year}-{month}", doctor, ng_dict_to_text(merged), now,
                                              st.session_state.get(comment_key, "").strip()])
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
        for d in range(1, num_days + 1):
            st.session_state[f"ng_{doctor}_{year}_{month}_{d}"] = result.get(d, "OK")
        st.session_state[f"ng_import_last_{doctor}_{year}_{month}"] = (result, base)
    how = "管理者の入力と合わせて" if mode == "merge" else "先生の提出内容に置き換えて"
    st.session_state["remote_import_message"] = f"{len(submissions)}名分のNG日を、{how}取り込みました。"


def remote_admin_panel(year, month, doctor_rows, custom_holidays, next_month_special_holiday, day_is_holiday):
    """管理者用：受付の開始・締め切り、専用URL、提出状況、取り込み。"""
    with st.expander("📨 先生方からのNG日入力（専用URL）", expanded=False):
        if not remote_is_configured():
            st.info("この機能を使うには、最初に保存用のGoogleスプレッドシートとの接続設定が必要です。設定方法は担当者向けの手順書を参照してください。")
            return
        admin_password = remote_secret("admin_password")
        if not admin_password:
            st.error("管理用パスワード（admin_password）が設定されていません。Secretsに追加してください。")
            return
        if not st.session_state.get("remote_admin_ok"):
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

        st.caption("毎月の流れ：① 受付を開始する → ② URL・配布用カードを配る（初回のみ）→ ③ 提出状況を確認 → ④ 締め切る → ⑤ 提出されたNG日を取り込む")
        month_key = f"{year}-{month}"
        active = settings.get("受付中の年月", "")
        closed = settings.get("締め切り", "") == "True"
        if active == month_key:
            status = "🔒 締め切り済み" if closed else "🟢 受付中"
            st.markdown(f"**{year}年{month}月：{status}**（締め切り日の表示：{settings.get('締め切り日', '') or '未設定'}）")
        elif active:
            st.warning(f"現在の受付は「{active.replace('-', '年')}月」です。上で選んでいる{year}年{month}月とは異なります。")
        else:
            st.markdown("**現在、受付していません。**")

        # --- 受付の開始・締め切り ---
        deadline = st.text_input("先生方に表示する締め切り日（例：10月25日）", value=settings.get("締め切り日", "") if active == month_key else "", key="remote_deadline")
        c1, c2, c3 = st.columns(3)
        if c1.button(f"{year}年{month}月の受付を開始する", key="remote_open", use_container_width=True,
                     help="上で設定した年月・特別休日で受付を始めます。先生の名簿と専用URLも更新します。"):
            remote_sync_doctors(doctor_rows)
            remote_save_settings(**{"受付中の年月": month_key, "特別休日": ",".join(map(str, sorted(custom_holidays))),
                                    "翌月1日特別休日": bool(next_month_special_holiday), "締め切り": False, "締め切り日": deadline})
            st.rerun()
        if active == month_key:
            if not closed and c2.button("🔒 締め切る（先生側の編集を止める）", key="remote_close", use_container_width=True):
                remote_save_settings(**{"締め切り": True})
                st.rerun()
            if closed and c2.button("受付を再開する", key="remote_reopen", use_container_width=True):
                remote_save_settings(**{"締め切り": False})
                st.rerun()
            if c3.button("締め切り日・休日設定を更新", key="remote_update", use_container_width=True,
                         help="受付中に「休日にする」や締め切り日を変えた場合に押します。"):
                remote_save_settings(**{"特別休日": ",".join(map(str, sorted(custom_holidays))),
                                        "翌月1日特別休日": bool(next_month_special_holiday), "締め切り日": deadline})
                st.rerun()
            saved_holidays = parse_holiday_list(settings.get("特別休日", ""))
            if saved_holidays != sorted(custom_holidays):
                st.warning("上のカレンダーの「休日にする」が、受付開始時の設定と違います。先生側の画面に反映するには「締め切り日・休日設定を更新」を押してください。")

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
            # 紙で配る用：QRコード付きカード（A4に8人分、点線で切り離す）
            card_targets = st.multiselect("配布用カードを作る先生（URLを作り直した先生だけ印刷する場合などは絞り込めます）",
                                          [r["先生"] for r in url_rows], default=[r["先生"] for r in url_rows], key="remote_card_targets")
            if card_targets:
                url_by_name = {r["先生"]: r["専用URL"] for r in url_rows}
                try:
                    st.download_button(f"🖨️ 配布用カード（QRコード付き・{len(card_targets)}名分）をダウンロード",
                                       remote_qr_cards_html([(n, url_by_name[n]) for n in card_targets]).encode("utf-8"),
                                       file_name="NG日入力_配布用カード.html", mime="text/html", key="remote_qr_cards", type="primary")
                    st.caption("ダウンロードしたファイルを開くとブラウザに表示されるので、「印刷する」を押してA4・片面で印刷します（1人1枚）。"
                               "三つ折りにすると宛名だけが外に出て、QRコードは内側に隠れます。左右をホチキスで留めて各先生に渡してください。")
                except ImportError:
                    st.error("カードを作るには、requirements.txt に「reportlab」の1行を追加してください。")
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
        importable = {name: r for name, r in submissions.items() if name in {n for n, _ in doctor_rows}}
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
                                    "取り込むNG日（提出）": ng_text_summary(r.get("NG日", ""), year, month, custom_holidays) or "なし"})
            if changes:
                st.markdown("**取り込むと変わる先生**")
                st.dataframe(pd.DataFrame(changes), hide_index=True, use_container_width=True)
                st.button(f"⬇️ 提出されたNG日を取り込む（{len(changes)}名分）", key="remote_import", type="primary",
                          on_click=import_remote_ng, args=(year, month, {n: importable[n] for n in [c["先生"] for c in changes]}, mode))
                st.caption("取り込むと、その先生のNG日は提出内容に置き換わります（CSVなどで入れていたNGは消えます）。"
                           "提出していない先生のNG日は変わりません。取り込んだあとも、下のカレンダーで個別に直せます。")
            else:
                st.caption("提出内容は、すべてアプリに取り込み済みです。")
        unknown = [name for name in submissions if name not in {n for n, _ in doctor_rows}]
        if unknown:
            st.caption(f"名簿にない先生の提出（取り込み対象外）：{'、'.join(unknown)}")

        # （この欄自体が折りたたみの中にあるため、履歴はチェックボックスで開く）
        if st.checkbox("変更の履歴を表示（だれが・いつ・どう変えたか）", key="remote_history_open"):
            history = [r for r in remote_read("submissions") if r.get("年月") == month_key]
            if history:
                st.dataframe(pd.DataFrame([{"保存日時": r.get("保存日時", ""), "先生": r.get("先生", ""),
                                            "NG日": ng_text_summary(r.get("NG日", ""), year, month, custom_holidays) or "NG日なし",
                                            "コメント": r.get("コメント", "")}
                                           for r in reversed(history)]), hide_index=True, use_container_width=True)
            else:
                st.caption("まだ履歴はありません。")


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

st.title("当直作成アプリ")
st.caption("年月・勤務条件・NG日を入力して、当直表の案を作成します。")
with st.expander("📖 使い方（初めての方・毎月の流れ・結果の見方）", expanded=False):
    tab_flow, tab_ng, tab_result, tab_trouble = st.tabs(["毎月の流れ", "NG日の集め方", "結果の見方", "困ったとき"])
    with tab_flow:
        st.markdown("""
**毎月の作業は、画面の上から順に進めれば完了します。**

1. **年月・休日を設定**：作る年月を選びます。平日に日直を置く日（年末年始など）は「休日にする」にチェックします。
2. **確定当直を入力（任意）**：担当が決まっている当直や、先月末の当直を入れます。月初の勤務間隔の確認に使われます。
3. **医師条件を読み込む**：途中保存から続ける場合は不要です。初めての場合や、手元の「医師条件CSV」を使う場合はアップロードします。
4. **NG日を集める**：先生方に専用URLから入力してもらい、管理画面で取り込みます（「NG日の集め方」のタブを参照）。
5. **不足時の優先順位を確認（任意）**：人数が足りない月に、どの枠から埋めるかを決めます。
6. **「この条件で当直案を作成する」を押す**：押すたびに違うパターンの案ができます。気に入った案が出るまで押し直して構いません。
7. **結果を確認してダウンロード**：「医師別の当直一覧」をExcelで、「作成した当直案」をCSVでダウンロードできます。

**途中で中断するとき**は、画面右上の「💾 作業を途中保存する」を押してください。スクロールしても右上に付いてくるので、いつでも押せます（管理用パスワードが必要です）。
年月・特別休日・増員設定・確定当直・医師条件（NG日を含む）・不足時の優先順位がまとめて途中保存され、
次回は画面右上の「📂 前回の作業の続きから始める」で元に戻せます。途中保存は、ボタンを押したときだけ行われます（当直案の作成では保存されません）。

        """)
    with tab_ng:
        st.markdown("""
**方法A：先生方に専用URLから入力してもらう（おすすめ）**

NG日の欄にある「📨 先生方からのNG日入力（専用URL）」を開き、管理用パスワードを入れて使います。

1. **受付を開始**：締め切り日を書いて「受付を開始する」を押します。
2. **URLを配る（初回だけ）**：「配布用カード（QRコード付き）」を印刷して、各先生に渡します。URLは毎月同じなので、翌月からは「入力が始まりました」と伝えるだけです。
3. **先生が入力**：先生はスマートフォンのカメラでQRコードを読み取り、カレンダーでNG日を選んで保存します。管理者へのコメントも書けます。締め切りまでは何度でも直せて、最後に保存した内容が採用されます。
   アプリは12時間ほど誰も開かないと休止し、英語の画面（Zzzz…）が出ます。青いボタン「Yes, get this app back up!」を押して1分ほど待つと開きます（配布用カードにも書いてあります）。
4. **提出状況を確認**：管理画面の表で、提出済み・未提出・コメントを確認できます。
5. **締め切る**：「締め切る」を押すと、先生側は編集できなくなります。
6. **取り込む**：「提出されたNG日を取り込む」を押すと、全員分が下のNGカレンダーに入ります。その先生のNG日は、提出内容に置き換わります。取り込んだあとも個別に直せます。

**方法B：管理者が直接入力する**

「NG日を入力する医師」で先生を選び、カレンダーでNG日を選んで「NG日を反映する」を押します。紙やメールで集めた場合はこちらを使います。

カレンダーの「日直」「宿直」の欄を押すと、その勤務だけNGになります（もう一度押すと可に戻ります）。

| NGの種類 | 意味 |
| --- | --- |
| 宿NG | 宿直（夜）に入れない。「宿直」の欄だけ押した状態 |
| 日NG | 日直（日中）に入れない。休日だけ。「日直」の欄だけ押した状態 |
| 全NG | 日直・宿直とも入れない。休日だけ。両方の欄を押した状態 |
| ⚠ | 翌日PMにdutyがあるため、原則その日の宿直を外す日 |
        """)
    with tab_result:
        st.markdown("""
**🗓️ 医師別の当直一覧**（先生ごとに1行、Excelで印刷できる表）

- 枠名のマス：決まった当直（色は枠の種類）
- **茶色の太い枠**：希望日が、希望どおりに通った当直
- **×**：当直NGの日（宿×＝宿直NG、日×＝日直NG）
- **黄色の斜線**：翌日PM dutyのため宿直を外す日
- **赤線の下の「⚠️ 未決定」**：人数が足りず、担当が決まっていない枠
- 右端の回数：宿直・日直／病棟・外来／平日・休日／合計

**📅 作成した当直案**（日付ごとの表）

名前にカーソルを合わせる（スマートフォンはタップ）と、同じ先生の当直が強調されます。
「特定の医師の当直を色別でハイライト」を開くと、選んだ先生を色分けできます。
        """)
    with tab_trouble:
        st.markdown("""
- **「不足」や「未決定」が出る**：NG日・回数の上限・勤務間隔の条件が厳しすぎる可能性があります。月間最大回数を増やす、勤務間隔を短くするなどして、もう一度作成してください。
- **アプリを開くと医師の名前が消えている**：しばらく使わないとアプリがリセットされます。医師条件CSVを読み込み直してください。先生方の提出内容と専用URLはスプレッドシートに残っているので、消えません。
- **先生から「URLが無効」と言われた**：URLを作り直していないか確認してください。カードをなくした場合は、管理画面の「URLを作り直す」で新しいカードを印刷します。
- **先生のNG日が当直案に反映されていない**：管理画面で「提出されたNG日を取り込む」を押したか確認してください。
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

# ==========================================
# 1. 上部ダッシュボード：年月と休日の設定
# ==========================================
st.header("📅 年月・休日・必要人数の設定")
st.info("作成する年月を選んでください。土日祝日は自動で休日扱いになります。平日にも日直を設けたい場合は、その日の「休日にする」にチェックを入れます。")

today = datetime.date.today()
if today.month == 12:
    default_year = today.year + 1
    default_month = 1
else:
    default_year = today.year
    default_month = today.month + 1

# 画面右上の帯に「前回の作業の続きから始める」「作業を途中保存する」を置く
work_save_floating()

if "sel_year" not in st.session_state:
    st.session_state["sel_year"] = default_year
if "sel_month" not in st.session_state:
    st.session_state["sel_month"] = default_month
col_y, col_m = st.columns(2)
year = int(col_y.number_input("年", min_value=2000, step=1, key="sel_year"))
month = int(col_m.number_input("月", min_value=1, max_value=12, step=1, key="sel_month"))
restored_state = st.session_state.get("restored_state")

st.divider()

st.subheader(f"📅 平日に日直を設ける日（特別休日） - {month}月")
st.caption("年末年始やお盆など、平日でも日直が必要な日を指定します。チェックした日は、日直・宿直ともに休日回数の集計対象になります。")

cal_matrix = calendar.monthcalendar(year, month)
custom_holidays = []

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
 height:112px !important; min-height:112px !important;
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
  height:100px !important; min-height:100px !important; padding:6px 1px !important;
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

next_month_first = datetime.date(year, month, calendar.monthrange(year, month)[1]) + datetime.timedelta(days=1)
next_month_first_label = f"{next_month_first.year}/{next_month_first.month}/1（{WEEKDAYS_JA[next_month_first.weekday()]}）"
month_last = next_month_first - datetime.timedelta(days=1)
month_last_label = f"{month_last.month}/{month_last.day}（{WEEKDAYS_JA[month_last.weekday()]}）"
duty_note = f"月末 {month_last_label} の宿直で「翌日PM duty」を判定するために使います。"
if next_month_first.weekday() >= 5 or jpholiday.is_holiday(next_month_first):
    # 土日祝は自動で休日扱いになるため、チェックボックスは表示しない
    st.caption(f"翌月1日 {next_month_first_label} は休日のため、設定不要です。（{duty_note}）")
    next_month_special_holiday = False
else:
    next_month_special_holiday = st.checkbox(
        f"翌月1日 {next_month_first_label} は病院独自の休み（特別休日）",
        key=f"next_month_special_holiday_{year}_{month}",
        help="平日でも病院が休みの場合だけチェックしてください。休みなら、月末の宿直は「翌日PM duty」の制限を受けません。"
    )
    st.caption(duty_note)

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
edited_multi_df = st.data_editor(
    multi_df_template,
    num_rows="dynamic",
    use_container_width=True,
    hide_index=True,
    height=150,
    column_config={
        "日付": st.column_config.SelectboxColumn("日付を選択", options=date_options, required=True),
        "当直枠": st.column_config.SelectboxColumn("増員する枠を選択", options=NIGHT_SHIFTS + DAY_SHIFTS, required=True),
        "人数": st.column_config.NumberColumn("合計人数", help="追加人数ではなく、その枠に配置する合計人数です。例：1名から2名体制にする場合は2。", min_value=2, max_value=10, step=1, required=True)
    }
)

multi_slots_dict = {}
for _, row in edited_multi_df.iterrows():
    d_str = str(row.get("日付", ""))
    s_val = str(row.get("当直枠", ""))
    c_val = row.get("人数")
    if d_str and s_val and pd.notna(c_val):
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
    is_hol = is_holiday_date(datetime.date(year, month, d), year, month, custom_holidays)
    for s in NIGHT_SHIFTS:
        shift_counts[s] += multi_slots_dict.get((d, s), 1)
    if is_hol:
        for s in DAY_SHIFTS:
            shift_counts[s] += multi_slots_dict.get((d, s), 1)

total_slots = sum(shift_counts.values())
holiday_total_placeholder.metric("🏥 必要な総当直枠数", f"{total_slots} 枠")

st.divider()

# ==========================================
# 3. メイン画面：データの読み込み＆画面入力
# ==========================================
st.header("1. 先月今月来月の確定当直を入力（任意）")
st.info("先月・今月・来月の、担当者が決まっている当直を入力してください。先月末の勤務は、月初の勤務間隔の確認に使います。入力がなければ、この項目は飛ばせます。")
with st.expander("確定済み当直の入力例と扱い", expanded=False):
    st.markdown("""
- CSVのアップロードと、下の表への直接入力のどちらでも入力できます。
- 日付は `2026/10/1` または `10/1` の形式で入力します。
- 担当する枠の欄に、医師条件と同じ名前を入力します。複数人の場合は `佐藤、鈴木` のように「、」で区切ります。
- 医師名の区切りには「、」を使ってください。名前に含まれる空白は区切りません。医師条件と同じ表記にしてください。
- 未確定の枠は空欄で構いません。「平日/休日」欄ではなく、上部のカレンダー設定で休日を判定します。
- 今月の平日に日直を入力する場合は、先に上のカレンダーでその日を「休日にする」にしてください。
- 今月の確定勤務はNG日・曜日制限より優先され、回数上限も必要に応じて緩められます。確定勤務を基準に追加勤務の間隔を守ります。確定勤務同士が近すぎる場合は、確定内容を維持して警告します。

**今月の同じ日・同じ枠を指定した場合、確定当直表への入力と「希望優先度100以上」での指定は、どちらも固定され、計算上の扱いは基本的に同じです。両方に入力する必要はありません。**
    """)

fixed_columns = ["日付", "平日/休日"] + ALL_SHIFT_TYPES
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
    fixed_file = st.file_uploader("決定済み当直表（CSV）をアップロード", type="csv", key=f"fixed_csv_{st.session_state.get('uploader_gen', 0)}")

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
base_fixed_df = base_fixed_df.set_index("日付")

st.markdown("### 📅 先月今月来月の確定当直")
st.write("表のセルをクリックして、日付と担当医師名を入力・編集できます。")
edited_fixed_df_raw = st.data_editor(base_fixed_df, num_rows="dynamic", use_container_width=True, height=200)
edited_fixed_df = edited_fixed_df_raw.reset_index()

st.divider()

st.header("2. 医師条件の読み込み・入力（必須）")
st.info("下の表に医師ごとの条件を入力してください。CSVを使う場合は、ひな形をダウンロードして編集し、アップロードします。NG日は、この後の医師別カレンダーで設定します。")
st.caption("最初に表示される5名は入力例です。実際の医師名・条件に置き換えてください。希望優先度は通常「1」を使用します。")
with st.expander("入力例：曜日・希望日・備考", expanded=False):
    st.markdown("""
| 項目 | 入力方法・意味 |
| --- | --- |
| 翌日PM duty | `水,木` のように半角カンマで区切ります。翌日PMにdutyがある曜日を入力します（例：木曜PMにdutyがある場合は「水」）。原則としてその曜日の宿直を外しますが、翌日が休日なら宿直に入る場合があります。日直は対象外です。 |
| 希望日 | `10,15` はその日のいずれかの枠、`10:A宿直` はその枠を希望します。複数の希望は半角カンマで区切ります。日直は休日（または「休日にする」にした日）だけ指定できます。 |
| 指定できる枠 | A宿直・B宿直・外来宿直・A日直・B日直・外来日直。表記を一致させてください。 |
| 備考 | 管理用のメモです。「学会」などと書いても計算条件には反映されません。休みはNG日で指定してください。 |

曜日にかかわらず勤務できない日は、カレンダーでNGを指定します。
平日は「宿NG」、休日に日直・宿直とも勤務できない場合は「全NG」を選んでください。
    """)
with st.expander("回数・勤務間隔の数え方", expanded=False):
    st.markdown("""
| 項目 | 意味・入力例 |
| --- | --- |
| 最低空ける日数 | 勤務と次の勤務の間に空ける日数です。5日なら、10日の次は16日以降です。 |
| 月間最小回数 | できるだけ確保したい回数です。条件によっては、この回数に届かないことがあります。 |
| 月間最大回数 | 日直と宿直を合わせた月間の上限です。 |
| 休日最大回数 | 土日祝日・特別休日に担当する日直と宿直の合計上限です。 |
| 各枠の上限 | A宿直など、それぞれの枠を担当する月間の上限です。 |

1つの枠を1回と数えます。確定指定により同じ日に日直と宿直を担当する場合は2回です。
月間最小回数は、月間最大回数以下に設定してください。
確定済み当直や優先度100以上の希望がある場合は、上限の例外があります。自動追加勤務の間隔は守り、確定勤務同士の間隔違反は警告します。
    """)
with st.expander("希望優先度：通常の希望と、100以上の特別な設定", expanded=False):
    st.markdown("""
- **通常は「1」**を使用します。1〜99は、数字が大きいほど希望を優先しますが、NG日・回数上限・勤務間隔などの範囲内で割り当てます。
- **100以上は、その医師の希望日すべてを確定扱いにする設定**です。「できれば入りたい」という用途には使わないでください。
- 確定扱いの日はNG日・曜日制限より優先され、回数上限が必要に応じて緩められます。その日と自動追加勤務の間隔は守ります。確定勤務同士の間隔違反は警告します。
- 一部の勤務だけを確定させたい場合は、上の「確定済み当直」へ入力し、希望優先度は通常の値にしてください。
- 指定の誤りや条件の組み合わせによっては作成できない場合があります。作成後に確定勤務が反映されているか確認してください。

**今月の同じ日・同じ枠を指定した場合、確定当直表への入力と「希望優先度100以上」での指定は、どちらも固定され、計算上の扱いは基本的に同じです。両方に入力する必要はありません。**
    """)

template_data = {
    "先生の名前": ["Dr. A", "Dr. B", "Dr. C", "Dr. D", "Dr. E"],
    "備考（メモ・説明など自由記入）": ["学会のため休み多め", "15日は午後休", "", "当直明け休み希望", ""],
    "翌日PM duty": ["水,木", "", "土,日", "", ""],
    NG_COLUMN: ["", "15:宿NG", "10:宿NG", "", ""],
    REQUEST_COLUMN: ["10:A宿直, 15", "", "8", "20", ""],
    "希望優先度(数字が大きいほど優先)": [1, 1, 1, 1, 1],
    "最低空ける日数": [5, 4, 6, 5, 3],
    "月間最小回数": [1, 2, 0, 1, 0],
    "月間最大回数": [5, 6, 4, 5, 7],
    "休日最大回数": [2, 2, 2, 2, 2],
    "A宿直上限": [2, 2, 2, 2, 2],
    "B宿直上限": [2, 2, 2, 2, 2],
    "外来宿直上限": [2, 2, 2, 2, 2],
    "A日直上限": [2, 2, 2, 2, 2],
    "B日直上限": [2, 2, 2, 2, 2],
    "外来日直上限": [2, 2, 2, 2, 2],
}
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
        "休日最大回数": st.column_config.NumberColumn("休日最大回数", help="土日祝・特別休日の日直と宿直の合計上限です。1枠を1回と数えます。"),
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
    st.stop()

st.markdown("### ⚖️ 必要枠数と担当可能回数の目安")
st.caption("月間最大回数の合計と必要枠数を比較しています。プラスでも、NG日・勤務間隔・枠別上限などによっては埋まらない場合があります。確定指定で追加される枠や上限の例外は、この目安に含まれません。")

if "月間最大回数" in staff_df.columns:
    total_max_capacity = int(pd.to_numeric(staff_df["月間最大回数"], errors='coerce').fillna(0).sum())
    margin = total_max_capacity - total_slots
    c1, c2, c3 = st.columns(3)
    c1.metric("🏥 必要な総当直枠数", f"{total_slots} 枠")
    c2.metric("👩‍⚕️ 医師の月間最大回数の合計", f"{total_max_capacity} 回分")
    if margin >= 0:
        c3.metric("✨ 担当可能回数 − 必要枠数", f"+{margin} 回分")
    else:
        c3.metric("🚨 担当可能回数 − 必要枠数", f"{margin} 回分", delta_color="inverse")
st.divider()

st.markdown("### 🚫 先生ごとのNG日設定（カレンダーで詳細選択）")
st.info("先生方に専用URLから入力してもらう場合は、下の「📨 先生方からのNG日入力（専用URL）」を使います。"
        "管理者が直接入力する場合は、医師を選んでカレンダーの「日直」「宿直」の欄を押してNGにし、最後に「NG日を反映する」を押してください。")
with st.expander("NGの種類・一括操作について", expanded=False):
    st.markdown("""
カレンダーの各日の「日直」「宿直」の欄を押すと、その勤務だけ「可」⇔「NG」が切り替わります。

| 表示 | 意味 |
| --- | --- |
| 日直 可・宿直 可 | この日についてNGを指定しません（OK）。曜日・回数など、ほかの条件は適用されます。 |
| 日直NG・宿直NG | 日直・宿直ともに勤務できません（全NG）。休日だけです。 |
| 日直NGのみ | 日直に勤務できません。宿直は候補になります（日NG）。休日だけです。 |
| 宿直NGのみ | 宿直に勤務できません。休日なら日直は候補になります（宿NG）。 |

平日は日直がないので、「宿直」の欄だけ押せます。
「全日NGにする」は平日を宿NG、休日を全NGにし、「すべてOKに戻す」はNG指定を解除します。
一括操作は押すとすぐ反映されます。確定済み当直・優先度100以上の希望は、NGより優先されます。

「休日にする」のチェックを外した日は、平日として表示されます（全NG→宿NG として扱い、日NGは平日には関係しません）。
チェックを戻すと、元のNG指定がそのまま復活します。

**「NG日を反映する」は、今開いている画面の計算条件に入れる操作です。**
次回も続きから使う場合は、画面右上の「💾 作業を途中保存する」を押してください。
    """)

ng_layout = st.radio("NGカレンダーの表示", ["月間カレンダー", "1日〜月末を横一列"], horizontal=True, key="ng_calendar_layout")
ng_horizontal = ng_layout == "1日〜月末を横一列"
st.caption("表示を切り替える前に、選択中のNG日を反映してください。反映済みの内容は、どちらの表示でも共通です。")
if ng_horizontal:
    st.caption("左右へスクロールして日付を選べます。選択中の色はすぐに変わります。最後に「NG日を反映する」を押してください。")

day_is_holiday = {d: is_holiday_date(datetime.date(year, month, d), year, month, custom_holidays) for d in range(1, num_days + 1)}

valid_staff = staff_df[staff_df["先生の名前"].astype(str).str.strip() != ""]
if not valid_staff.empty:
    doctor_names = valid_staff["先生の名前"].astype(str).tolist()
    remote_admin_panel(
        year, month,
        [(str(r["先生の名前"]), clean_text(r.get("翌日PM duty", ""))) for _, r in valid_staff.iterrows()],
        custom_holidays, next_month_special_holiday, day_is_holiday,
    )
    # 医師の選択はタブではなく選択欄にする。
    # （タブは表示の切り替えなどで画面が作り直されると先頭の医師に戻ってしまい、入力ミスの原因になるため）
    if st.session_state.get("ng_selected_doctor") not in doctor_names:
        st.session_state["ng_selected_doctor"] = doctor_names[0]

    # 先生方のコメント（専用URLから）と、医師条件の備考。選択欄の💬マークと、カレンダーの上の表示に使う
    doctor_comments = remote_comments_safe(year, month)
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
            if effective_ng(st.session_state.get(f"ng_{name}_{year}_{month}_{d}", "OK"), day_is_holiday[d]) != "OK"
        )
        mark = "　💬" if name in doctor_comments or doctor_notes.get(name) else ""
        return (f"{name}（NG {count}日）" if count else f"{name}（NGなし）") + mark

    def move_ng_doctor(step):
        i = doctor_names.index(st.session_state["ng_selected_doctor"])
        st.session_state["ng_selected_doctor"] = doctor_names[(i + step) % len(doctor_names)]

    sel_col, prev_col, next_col = st.columns([6, 1.2, 1.2])
    selected_doc = sel_col.selectbox("NG日を入力する医師", doctor_names, key="ng_selected_doctor", format_func=ng_doctor_label)
    # ボタンの高さを選択欄にそろえるための余白
    prev_col.markdown("<div style='height:1.75rem'></div>", unsafe_allow_html=True)
    next_col.markdown("<div style='height:1.75rem'></div>", unsafe_allow_html=True)
    prev_col.button("◀ 前の医師", key="ng_prev_doctor", on_click=move_ng_doctor, args=(-1,), use_container_width=True)
    next_col.button("次の医師 ▶", key="ng_next_doctor", on_click=move_ng_doctor, args=(1,), use_container_width=True)

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
                shown = {d: effective_ng(st.session_state[f"ng_{doc_name}_{year}_{month}_{d}"], day_is_holiday[d]) for d in range(1, num_days + 1)}

                saved_strs = []
                for d, val in shown.items():
                    if val == "全NG": saved_strs.append(f"{d}日")
                    elif val == "日NG": saved_strs.append(f"{d}日(日直NG)")
                    elif val == "宿NG": saved_strs.append(f"{d}日(宿直NG)")
                if saved_strs:
                    st.success(f"✅ **反映済みのNG日:** {', '.join(saved_strs)}")
                else:
                    st.info("💡 **現在、反映されているNG日はありません**")

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
                    })
                revision = hashlib.sha256(json.dumps([year, month, doc_name, ng_layout, component_days], ensure_ascii=False).encode()).hexdigest()
                component_key = f"ng_editor_{'row' if ng_horizontal else 'month'}_{doc_name}_{year}_{month}"
                st.markdown("<div style='color:#bf5700;background:#fff0c2;border:1px solid #ef9b20;border-radius:6px;padding:8px 10px;font-size:0.9rem;font-weight:600;'>⚠は、翌日PMにdutyがあるため、原則としてその日の宿直を外すことを示します。翌日が休日の場合は⚠を表示せず、この制限の対象外になります。</div>", unsafe_allow_html=True)
                response = horizontal_ng_component(HORIZONTAL_NG_HTML)(
                    days=component_days, mode="row" if ng_horizontal else "month",
                    offset=datetime.date(year, month, 1).weekday(), doctor=doc_name,
                    version=revision, key=component_key, default=None,
                )
                seen_key = component_key + "_last_token"
                if isinstance(response, dict) and response.get("token") != st.session_state.get(seen_key):
                    st.session_state[seen_key] = response.get("token")
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
            if val == "全NG":
                ng_items.append(str(d))
            elif val != "OK":
                ng_items.append(f"{d}:{val}")
        staff_df.at[original_idx, NG_COLUMN] = ",".join(ng_items)



st.divider()
st.markdown("### 📂 医師条件をCSVでダウンロード（手元にファイルで残す場合）")
st.write("医師名・回数・勤務間隔・希望日・NG日・備考をCSVファイルとしてパソコンなどにダウンロードします。このファイルは「医師条件」のアップロード欄から読み込めます。")
st.caption("ダウンロードの前に、各医師の「NG日を反映する」を押してください。対象年月・特別休日・増員設定・確定済み当直・生成結果・色分けとそのメモは、このCSVには含まれません。"
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
st.divider()
st.markdown("""
<style>
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
    "next_special": bool(next_month_special_holiday),
    "multi": [[f"{d}日", s, int(c)] for (d, s), c in multi_slots_dict.items()],
    "fixed_csv": edited_fixed_df.to_csv(index=False),
    "staff_csv": staff_df.to_csv(index=False),
    "priority": {k: int(v) for k, v in shortage_priority.items()},
}
# 右上の途中保存ボタンが押されていたら、ここで（全部の入力を読み終えてから）保存する
if work_save_request:
    _slot = WORK_SAVE_SLOT or st.empty()
    try:
        work_state_save(work_snapshot)
        _slot.markdown(f"<span class='work-save-note'>✅ {remote_now()[-8:-3]} 途中保存しました</span>", unsafe_allow_html=True)
        st.toast("途中保存しました。次回は画面右上の「📂 前回の作業の続きから始める」から読み込めます。", icon="💾")
    except Exception as e:
        _slot.markdown("<span class='work-save-note bad'>⚠ 途中保存できませんでした</span>", unsafe_allow_html=True)
        st.error(f"途中保存できませんでした。通信状態を確認して、もう一度押してください。詳細：{e}")

create_button_container = st.container(key="create_duty_action")
st.header("3. 当直案の作成・確認")
result_notice_container = st.container()

staff_df = staff_df[staff_df['先生の名前'].astype(str).str.strip() != '']
staff_df = staff_df.dropna(subset=['先生の名前']).reset_index(drop=True)

fixed_df, fixed_input_errors = validate_fixed_inputs(edited_fixed_df, staff_df['先生の名前'].tolist(), year, month, custom_holidays)
if fixed_input_errors:
    show_input_errors(fixed_input_errors)
    st.stop()
current_signature = input_signature(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority)
if 'generated_df' in st.session_state and st.session_state.get('generated_signature') != current_signature:
    for key in ['generated_df', 'past_worked_dates', 'future_worked_dates', 'generated_warnings', 'generated_signature', 'generated_year', 'generated_month']:
        st.session_state.pop(key, None)
    st.session_state['result_needs_refresh'] = True
if st.session_state.get('result_needs_refresh'):
    st.warning("条件が変更されています。再作成してください。前回の結果表示とダウンロードを停止しました。")

if len(staff_df) > 0:
    with create_button_container:
        create_clicked = st.button("🚀 この条件で当直案を作成する", type="primary", use_container_width=True)
        st.caption("※ 押すたびに、同じ条件でも違うパターンの当直表を作ります。気に入った案は、下のボタンからダウンロードしてください（作成した当直案は途中保存には含まれません）。")
    if create_clicked:
        with st.spinner("当直案を計算中…（通常は最大60秒、不足枠の確認を含む場合は計算時間が最大75秒です）"):
            try:
                df_result, success, error_reasons, past_worked_dates, future_worked_dates = generate_shift(year, month, staff_df, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority)

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

    if 'generated_df' in st.session_state:
        year = st.session_state['generated_year']
        month = st.session_state['generated_month']
        with result_notice_container:
            show_schedule_notice(st.session_state['generated_df'], st.session_state.get('generated_warnings', []))
        df_result = st.session_state['generated_df'].reindex(columns=RESULT_COLUMNS)
        past_worked_dates = st.session_state.get('past_worked_dates', {})
        future_worked_dates = st.session_state.get('future_worked_dates', {})

        shift_columns = ALL_SHIFT_TYPES
        doctors_list = staff_df['先生の名前'].astype(str).tolist()

        def split_names(value):
            return [x.strip() for x in re.split(r'[、,]', str(value))]

        st.subheader("🗓️ 医師別の当直一覧")
        st.caption("決まった当直は枠名、当直NGは×、翌日PM dutyで宿直を外す日は斜線で表示します。希望日が通った当直は茶色の太い枠で囲みます。")
        matrix_days, matrix_rows = build_duty_matrix(df_result, staff_df, year, month, custom_holidays, next_month_special_holiday,
                                                     comments=remote_comments_safe(year, month))
        missing_total = sum(len(info['shortage']) for info in matrix_days)
        if missing_total:
            st.warning(f"当直が決まっていない枠が{missing_total}枠あります。表の一番下（赤線の下）で確認できます。")
        st.markdown(duty_matrix_html(matrix_days, matrix_rows, year, month), unsafe_allow_html=True)
        try:
            st.download_button(
                label="📥 この一覧をExcelでダウンロード（印刷用）",
                data=duty_matrix_xlsx(matrix_days, matrix_rows, year, month),
                file_name=f"当直一覧_{year}年{month}月.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
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
