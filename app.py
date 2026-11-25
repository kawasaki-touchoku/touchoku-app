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
NUMERIC_DEFAULTS = {'希望優先度(数字が大きいほど優先)': 1, '宿⇔宿 間隔下限': 5, '月下限': 0, '月上限': 5, '休日上限': 4, **{s + '上限': 2 for s in ALL_SHIFT_TYPES}}
NG_COLUMN = 'NG日(半角カンマ区切り)'
# 宿直・日直それぞれの月間上限（A・B・外来を合わせた回数）。空欄なら制限なし
GROUP_CAP_COLUMNS = {'宿直上限': ['A宿直', 'B宿直', '外来宿直'], '日直上限': ['A日直', 'B日直', '外来日直']}
REQUEST_COLUMN = '希望日(半角カンマ区切り)'
# 日直がからむ勤務の間隔（日直→宿直・宿直→日直・日直→日直）。宿直→宿直は「宿⇔宿 間隔下限」を使う。
# この列がない・空欄のときは「宿⇔宿 間隔下限」と同じ値にする（以前の動きのまま）
DAY_GAP_COLUMN = '宿⇔日・日⇔日 間隔下限'
DAY_GAP_TEMPLATE_DEFAULT = 2   # 新しいひな形の初めの値
# 多め希望（できれば多めに入れる種類）。{列名: (選べる値1, 選べる値2)}。反対向きどうしを同じ列にして、両方は選べないようにする
DOC_PREF_COLUMNS = {"病棟/外来 多め": ("病棟多め", "外来多め"), "宿直/日直 多め": ("宿直多め", "日直多め"), "平日/休日 多め": ("平日多め", "休日多め")}
DOC_PREF_WEIGHT = 60   # 多め希望の強さ（希望日の優先度1＝100より弱い「できれば」）


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


def _read_table_any(file_bytes):
    """
    Excel（.xlsx）でもCSVでも、表として読む。Excelは1枚目のシートを読む。
    ・Excelで日付になってしまったマスは「月/日」の文字に戻す（NG日・確定当直の日付が読めるように）
    ・見出しのない空の列（Excelで書式だけ残った列など）と、空の行は捨てる
    """
    if not bytes(file_bytes[:4]) == b"PK\x03\x04":
        return _read_csv_any_encoding(file_bytes)
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=0, engine="openpyxl")
    except Exception:
        raise ValueError("Excelファイルとして読めませんでした。ファイルが壊れていないか、パスワードが付いていないかを確認してください。")
    df = df[[c for c in df.columns if not (str(c).startswith("Unnamed:") and df[c].isna().all())]]
    df = df.dropna(how="all").reset_index(drop=True)

    def _fix(v):
        if isinstance(v, (datetime.datetime, datetime.date)) and not pd.isna(v):
            return f"{v.month}/{v.day}"
        return v
    return df.apply(lambda col: col.map(_fix) if col.dtype == object or str(col.dtype).startswith("datetime") else col)


def table_xlsx(df, sheet_name, text_cols=None, blank_rows=0):
    """
    表をExcelにする（見出しは色つき・固定、列の幅を整える）。
    ・text_cols の列は「文字」の形にして、入れた「11/3」などが勝手に日付に変わらないようにする（None なら数字以外の列すべて）
    ・blank_rows：ひな形のように、下に入力用の空の行を「文字」の形で用意する行数
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name[:31]
    cols = [str(c) for c in df.columns]
    if text_cols is None:
        text_cols = [c for c in df.columns if not pd.api.types.is_numeric_dtype(df[c])]
    text_idx = {i for i, c in enumerate(df.columns, 1) if c in set(text_cols)}
    for i, c in enumerate(cols, 1):
        cell = ws.cell(1, i, c)
        cell.font = Font(bold=True)
        cell.fill = PatternFill("solid", fgColor="DFE6EF")
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    for r, row in enumerate(df.itertuples(index=False), 2):
        for i, v in enumerate(row, 1):
            if v is None or (not isinstance(v, str) and pd.isna(v)):
                v = None
            elif isinstance(v, float) and v.is_integer():
                v = int(v)
            elif hasattr(v, "item") and not isinstance(v, str):
                v = v.item()
            if i in text_idx and v is not None:
                v = str(v)
            cell = ws.cell(r, i, v)
            if i in text_idx:
                cell.number_format = "@"
    for r in range(len(df) + 2, len(df) + 2 + blank_rows):
        for i in text_idx:
            ws.cell(r, i).number_format = "@"
    for i, c in enumerate(cols, 1):
        longest = max([len(c)] + [len(str(v)) for v in df.iloc[:, i - 1].tolist() if v is not None and not (not isinstance(v, str) and pd.isna(v))])
        ws.column_dimensions[get_column_letter(i)].width = max(8, min(40, longest * 1.8 + 2))
    ws.row_dimensions[1].height = 30
    ws.freeze_panes = "B2"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"


def staff_xlsx_with_groups(df, text_cols, rules):
    """医師条件のExcel（1枚目：医師条件、2枚目：グループの設定）。"""
    from openpyxl import load_workbook
    wb = load_workbook(io.BytesIO(table_xlsx(df, "医師条件", text_cols=text_cols)))
    group_rules_xlsx_sheet(wb, rules)
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


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


# 以前の列名 → 今の列名（以前のCSV・途中保存も読み込めるように）
STAFF_OLD_COLUMN_NAMES = {"月間" + "最小回数": "月下限", "月間" + "最大回数": "月上限",
                          "休日" + "最大回数": "休日上限", "最低" + "空ける日数": "宿⇔宿 間隔下限",
                          "間隔" + "下限": "宿⇔宿 間隔下限", "宿直" + "間隔": "宿⇔宿 間隔下限", "宿直" + "間隔下限": "宿⇔宿 間隔下限",
                          "日直" + "間隔": "宿⇔日・日⇔日 間隔下限", "日直" + "間隔下限": "宿⇔日・日⇔日 間隔下限"}


def rename_old_staff_columns(df):
    """
    以前の列名を今の列名に読み替える。同じ列の以前の名前が2つ以上ある（版の途中で保存した）ときは、
    今の名前の列があればそれを、なければいちばん左の列を使い、残りは捨てる（同じ名前の列が2つあると表が出せないため）。
    """
    for new in dict.fromkeys(STAFF_OLD_COLUMN_NAMES.values()):
        olds = [c for c in df.columns if STAFF_OLD_COLUMN_NAMES.get(c) == new]
        if not olds:
            continue
        if new in df.columns:
            df = df.drop(columns=olds)
        else:
            df = df.drop(columns=olds[1:]).rename(columns={olds[0]: new})
    return df.loc[:, ~df.columns.duplicated()]


@st.cache_data
def parse_staff_csv(file_bytes):
    """医師条件のファイル（Excel・CSV）を読む。"""
    df = normalize_shift_names(_read_table_any(file_bytes))
    # 旧CSVの列名も受け付け、画面・出力CSVでは新名称に統一する。
    # 以前の列名のCSV・途中保存も、今の列名に読み替える
    df = rename_old_staff_columns(df)
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
    return add_day_gap_column(df)


def add_day_gap_column(df):
    """「宿⇔日・日⇔日 間隔下限」の列がない（以前の）CSV・途中保存では、「宿⇔宿 間隔下限」の右に列を足し、宿⇔宿 間隔下限と同じ値を入れる。
    空欄のマスも宿⇔宿 間隔下限と同じ値にする（今までの動きを変えないため）。"""
    if "宿⇔宿 間隔下限" not in df.columns:
        return df
    df = df.copy()
    base = pd.to_numeric(df["宿⇔宿 間隔下限"], errors="coerce")
    if DAY_GAP_COLUMN not in df.columns:
        df.insert(list(df.columns).index("宿⇔宿 間隔下限") + 1, DAY_GAP_COLUMN, base)
    else:
        blank = df[DAY_GAP_COLUMN].map(lambda v: clean_text(v).lower() in ("", "nan", "none", "<na>"))
        if blank.any():
            df[DAY_GAP_COLUMN] = df[DAY_GAP_COLUMN].astype(object)
            df.loc[blank, DAY_GAP_COLUMN] = base[blank]
    return df


@st.cache_data
def parse_fixed_csv(file_bytes):
    """確定当直のファイル（Excel・CSV）を読む。"""
    df = normalize_shift_names(_read_table_any(file_bytes))
    if '区分' in df.columns:
        df = df.rename(columns={'区分': '平日/休日'})
    return df


# 確定当直の表の列：下の「作成した当直案」と同じ並び（平日か休日かは上のカレンダーで決まるので入力しない）
FIXED_SLOT_COLUMNS = [c for c in RESULT_COLUMNS[2:]]


def fixed_rows_for_editor(df):
    """
    確定当直の表を、画面で編集しやすい形にそろえる。
    ・列は「日付」＋当直案と同じ並びの枠（ほかの列は後ろに残す。入力チェックで知らせるため）
    ・1つのマスに「佐藤、鈴木」のように複数人いる場合は、同じ日付の行を増やして1マス1人にする
      （名前は一覧から選ぶので、1マスに1人だけ入れる形にする）
    """
    df = df.drop(columns=["平日/休日", FIXED_KUBUN_COLUMN], errors="ignore").copy()
    for c in ["日付"] + FIXED_SLOT_COLUMNS:
        if c not in df.columns:
            df[c] = ""
    extra = [c for c in df.columns if c not in ["日付"] + FIXED_SLOT_COLUMNS]
    df = df[["日付"] + FIXED_SLOT_COLUMNS + extra]

    def _text(v):
        return "" if v is None or (not isinstance(v, str) and pd.isna(v)) or str(v).strip().lower() in ("nan", "none", "<na>") else str(v).strip()

    rows = []
    for _, r in df.iterrows():
        base = {c: _text(r[c]) for c in df.columns}
        names = {s: [n.strip() for n in re.split(r"[、,]+", base[s]) if n.strip()] for s in FIXED_SLOT_COLUMNS}
        k = max([len(v) for v in names.values()] + [1])
        for i in range(k):
            row = dict(base)
            for s_ in FIXED_SLOT_COLUMNS:
                row[s_] = names[s_][i] if i < len(names[s_]) else ""
            if i > 0:
                for c in extra:
                    row[c] = ""
            rows.append(row)
    out = pd.DataFrame(rows, columns=df.columns)
    if out.empty:
        out.loc[0] = ["" for _ in out.columns]
    return out.astype(str).reset_index(drop=True)


# 確定当直の表には、対象の月（期間）の前後この日数分の日付を、最初から並べておく
FIXED_PAD_DAYS = 7
# 確定当直の表の「区分」の列（日付から自動で表示する。入力はしない・保存もしない）
FIXED_KUBUN_COLUMN = "区分"


def fixed_date_text(dt, near):
    """確定当直の表に出す日付の文字（例：11/3(月)）。対象の月から離れた日付は、年も付ける。"""
    youbi = WEEKDAYS_JA[dt.weekday()]
    if abs((dt - near).days) <= 150:
        return f"{dt.month}/{dt.day}({youbi})"
    return f"{dt.year}/{dt.month}/{dt.day}({youbi})"


def fixed_merge_into_days(df, start, end, year, month, pad=FIXED_PAD_DAYS):
    """
    確定当直の表に、start の pad 日前〜 end の pad 日後の日付を最初から並べ、そこへ df の内容を入れる。
    ・df は fixed_rows_for_editor でそろえた表（1マス1人）
    ・名前は、同じ日付の行の空いている枠に入れる。空いていなければ、その日付の行を増やす
    ・並べた範囲の外の日付（前月1か月分など）の行も、日付順に残す
    ・日付が読めない行は、入力チェックで知らせるため、いちばん下に残す
    ・名前の入っていない行（日付だけの行）は、並べた範囲の外なら消す
    """
    cols = list(df.columns)
    extra = [c for c in cols if c not in ["日付"] + FIXED_SLOT_COLUMNS]

    def blank(dt):
        row = {c: "" for c in cols}
        row["日付"] = fixed_date_text(dt, start)
        return row

    buckets = {}
    dt = start - datetime.timedelta(days=pad)
    while dt <= end + datetime.timedelta(days=pad):
        buckets[dt] = [blank(dt)]
        dt += datetime.timedelta(days=1)
    leftovers = []
    for _, r in df.iterrows():
        row = {c: clean_text(r.get(c, "")) for c in cols}
        row = {c: ("" if v.lower() in ("nan", "none", "<na>") else v) for c, v in row.items()}
        has_names = any(row[s] for s in FIXED_SLOT_COLUMNS)
        has_extra = any(row[c] for c in extra)
        if not has_names and not has_extra:
            continue
        dt = parse_shift_date(row["日付"], year, month) if row["日付"] else None
        if dt is None:
            leftovers.append(row)
            continue
        rows_d = buckets.setdefault(dt, [])
        if has_extra:
            new = dict(row)
            new["日付"] = fixed_date_text(dt, start)
            rows_d.append(new)
            continue
        for s in FIXED_SLOT_COLUMNS:
            if not row[s] or row[s] in [x[s] for x in rows_d]:
                continue
            target = next((x for x in rows_d if not x[s] and not any(x[c] for c in extra)), None)
            if target is None:
                target = blank(dt)
                rows_d.append(target)
            target[s] = row[s]
    rows = [x for d in sorted(buckets) for x in buckets[d]] + leftovers
    return pd.DataFrame(rows, columns=cols).astype(str).reset_index(drop=True)


def fixed_kubun(date_text, start, end, year, month, labels):
    """確定当直の表の「区分」：日付が対象の前なら labels[0]、対象の中なら labels[1]、後なら labels[2]。"""
    dt = parse_shift_date(date_text, year, month) if clean_text(date_text) else None
    if dt is None:
        return ""
    return labels[0] if dt < start else labels[2] if dt > end else labels[1]


def commit_fixed_current():
    """
    確定当直の表の今の内容（画面で入力した分）を、もとの表として確定し、表を出し直す。
    先生の名前の一覧が変わったときや、「表で選べるようにする」を押したときだけ使う（入力のたびには表を作り直さない）。
    """
    cur = st.session_state.get("fixed_current")
    args = st.session_state.get("fixed_window_args")
    if cur is not None and args is not None:
        st.session_state["fixed_rows"] = fixed_merge_into_days(fixed_rows_for_editor(cur), *args)
    st.session_state["fixed_editor_gen"] = st.session_state.get("fixed_editor_gen", 0) + 1


def add_fixed_extra_name():
    """「表で選べるようにする」を押したとき：入力した名前を、確定当直の表で選べる名前に加える。"""
    name = clean_text(st.session_state.get("fixed_extra_input", ""))
    if not name:
        return
    names = list(st.session_state.get("fixed_extra_names", []))
    if name not in names:
        names.append(name)
    st.session_state["fixed_extra_names"] = names
    st.session_state["fixed_extra_input"] = ""
    commit_fixed_current()


def fixed_from_posting_xlsx(data, file_label):
    """
    掲示用Excel（手で直したものでもよい）を、確定当直の表の形にする。読めないときは ValueError。
    ・ICU病棟の列は読まない（当直案の計算には使わないため）。「未決定」（以前の「未定」）のマスは空欄にする
    ・戻り値：(確定当直の表, 画面に出す知らせの文)
    """
    got = annual_read_posting(data, file_label)
    rows = []
    for dt in sorted(got["days"]):
        row = {"日付": f"{dt.year}/{dt.month}/{dt.day}"}
        for s in FIXED_SLOT_COLUMNS:
            row[s] = "、".join(n for n, slot in got["days"][dt]["duties"] if slot == s)
        rows.append(row)
    first, last = min(got["days"]), max(got["days"])
    note = (f"✅ 掲示用Excel「{file_label}」を読み込みました：{first.month}/{first.day}〜{last.month}/{last.day}の{len(got['days'])}日分"
            "（確定当直の表の同じ日付のところに入れました。ICU病棟の列と「未決定」の枠は読みません）")
    return pd.DataFrame(rows, columns=["日付"] + FIXED_SLOT_COLUMNS), note


def fixed_duty_table(fixed_file, restored_state, fixed_columns, start, end, year, month, labels, example):
    """
    確定当直の表を出し、入力した内容（「区分」の列を除く）を返す。
    ・表の中身は画面の裏（session_state）に持っておき、CSVや途中保存を読み込んだとき、
      対象の月（期間）が変わったとき、名前の一覧が変わったときだけ作り直す（入力のたびには作り直さない）
    ・前後 FIXED_PAD_DAYS 日分の日付は、最初から並べておく
    ・labels は「区分」の表示（前・中・後）、example は日付の入力例
    """
    if fixed_file is not None:
        src = ("file", fixed_file.name, fixed_file.size)
    elif restored_state and restored_state.get("fixed_csv"):
        src = ("restored", hashlib.sha256(restored_state["fixed_csv"].encode("utf-8")).hexdigest())
    else:
        src = ("empty",)
    args = (start, end, year, month)
    if st.session_state.get("fixed_src") != src or "fixed_rows" not in st.session_state:
        st.session_state.pop("fixed_load_note", None)
        if src[0] == "file" and str(fixed_file.name).lower().endswith(".xlsx"):
            try:
                base, note = fixed_from_posting_xlsx(fixed_file.getvalue(), fixed_file.name)
            except ValueError as e:
                st.error(f"確定当直ファイルの読み込みに失敗しました。{e}")
                st.stop()
            st.session_state["fixed_load_note"] = note
        elif src[0] == "file":
            try:
                base = parse_fixed_csv(fixed_file.getvalue())
            except Exception as e:
                st.error(f"確定当直ファイルの読み込みに失敗しました。詳細: {e}")
                st.stop()
            if "日付" not in base.columns:
                st.error("確定当直のファイルに「日付」の列がありません。ひな形の列名を確認してください。")
                st.stop()
        elif src[0] == "restored":
            base = parse_fixed_csv(restored_state["fixed_csv"].encode("utf-8"))
        else:
            base = pd.DataFrame(columns=fixed_columns)
        st.session_state["fixed_rows"] = fixed_merge_into_days(fixed_rows_for_editor(base), *args)
        st.session_state["fixed_src"] = src
        st.session_state["fixed_window_args"] = args
        st.session_state.pop("fixed_current", None)
        st.session_state["fixed_editor_gen"] = st.session_state.get("fixed_editor_gen", 0) + 1
    elif st.session_state.get("fixed_window_args") != args:
        # 月（期間）を変えたとき：入力済みの内容はそのままで、前後の日付を並べ直す
        st.session_state["fixed_window_args"] = args
        commit_fixed_current()
    base = st.session_state["fixed_rows"]
    if st.session_state.get("fixed_load_note") and src[0] == "file":
        st.success(st.session_state["fixed_load_note"])

    # 担当医師は、医師条件の名前（と「表で選べるようにする」で入れた名前）から選ぶ
    names = list(dict.fromkeys(
        list(st.session_state.get("fixed_name_options", []))
        + list(st.session_state.get("fixed_extra_names", []))
        + [v for s_ in FIXED_SLOT_COLUMNS for v in base[s_].tolist() if v]))
    st.markdown("<div class='fx-table-h'>✏️ 確定当直の表<span>方法1はここに入力します。方法2で読み込んだ内容も、ここに入ります。</span></div>",
                unsafe_allow_html=True)
    st.caption(f"ほかの日付は、表の下の「＋」で行を足して日付（例：{example}）を入力します。")
    if not st.session_state.get("fixed_name_options"):
        st.caption("先に下の「医師条件」に先生の名前を入れると、ここで名前を選べるようになります。")

    view = base.copy()
    view.insert(1, FIXED_KUBUN_COLUMN, [fixed_kubun(v, start, end, year, month, labels) for v in view["日付"]])
    for s_ in FIXED_SLOT_COLUMNS:
        view[s_] = view[s_].astype(object).where(view[s_] != "", None)
    config = {
        "日付": st.column_config.TextColumn("日付", help=f"例：{example}"),
        FIXED_KUBUN_COLUMN: st.column_config.TextColumn(FIXED_KUBUN_COLUMN, disabled=True,
                                                        help=f"日付から自動で表示します（{labels[0]}／{labels[1]}／{labels[2]}）。"),
    }
    config.update({s_: st.column_config.SelectboxColumn(s_, options=names, help="担当する先生を一覧から選びます。消すときはマスを選んで Delete キー。")
                   for s_ in FIXED_SLOT_COLUMNS})
    config.update({c: st.column_config.TextColumn(c) for c in view.columns if c not in config})
    with st.container(key="fixed_table_box"):   # 方法1と同じ青い枠で囲む（手で入れる場所だと分かるように）
        cur = st.data_editor(
            view, num_rows="dynamic", use_container_width=True, height=420, hide_index=True,
            column_config=config, key=f"fixed_editor_{st.session_state.get('fixed_editor_gen', 0)}",
        )
    st.caption("同じ日・同じ枠を2人以上で担当する場合は、表の下の「＋」で同じ日付の行を足して、もう1人を選んでください。")

    # 前月・翌月（期間の前後）だけ勤務した先生：困ったときに開く欄にする（医師条件の「➕ 先生を追加する」と混ざらないように）
    rng = "前月・翌月" if labels[1] == "今月" else "期間の前後"
    who = "前月だけ勤務した先生など" if labels[1] == "今月" else "期間の前だけ勤務した先生など"
    with st.expander(f"選びたい先生が、名前の一覧に出てこないとき（{who}）", expanded=False):
        st.caption(f"医師条件にいない先生の名前を入れると、表で選べるようになります。**{rng}の日付だけ**に使えます（{labels[1]}の当直案には入りません）。")
        c_in, c_btn = st.columns([3, 2], vertical_alignment="center")
        with c_in:
            st.text_input("先生の名前", key="fixed_extra_input", placeholder="例：佐藤", label_visibility="collapsed")
        with c_btn:
            st.button("表で選べるようにする", key="fixed_extra_add", on_click=add_fixed_extra_name, use_container_width=True)
        st.caption(f"{labels[1]}の先生を増やすときは、下の医師条件の「➕ 先生を追加する」を使います。")

    cur = cur.drop(columns=[FIXED_KUBUN_COLUMN], errors="ignore").reset_index(drop=True)
    for c in cur.columns:
        cur[c] = cur[c].map(lambda v: "" if v is None or (not isinstance(v, str) and pd.isna(v)) else str(v).strip())
    cur = cur.astype(str)
    st.session_state["fixed_current"] = cur
    return cur.copy()


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
def duty_kind(slot):
    """枠名から 'day'（日直）か 'night'（宿直）かを返す。"""
    return 'day' if slot in DAY_SHIFTS else 'night'


def worked_items(items):
    """前後の月の確定当直の一覧を [(日付, 'day'|'night')] にそろえる（日付だけの古い形は宿直として扱う）。"""
    out = []
    for it in items or []:
        if isinstance(it, (tuple, list)) and len(it) == 2:
            out.append((it[0], it[1] if it[1] in ('day', 'night') else 'night'))
        else:
            out.append((it, 'night'))
    return out


def gap_needed(kind_a, kind_b, night_gap, day_gap):
    """2つの勤務の間に空ける日数。宿直→宿直は宿⇔宿 間隔下限、日直がからむ組み合わせは宿⇔日・日⇔日 間隔下限。"""
    return night_gap if (kind_a == 'night' and kind_b == 'night') else day_gap


def external_gap_block(dt, external, night_gap, day_gap):
    """前後の月の確定当直のため、その日に (日直に入れない, 宿直に入れない) を返す。"""
    day_block = night_block = False
    for ext, kind in worked_items(external):
        dist = abs((dt - ext).days)
        if dist <= gap_needed(kind, 'day', night_gap, day_gap):
            day_block = True
        if dist <= gap_needed(kind, 'night', night_gap, day_gap):
            night_block = True
    return day_block, night_block


def external_gap_blocked(staff_df, past, future, num_days, date_of):
    """
    医師別の当直一覧の薄紫のマス用：前後の月の確定当直のため入れない日。
    {先生: {日: (日直に入れない, 宿直に入れない)}}（どちらも入れる日は入れない）
    """
    blocked = {}
    for _, r in staff_df.iterrows():
        name = str(r['先生の名前'])
        gv = pd.to_numeric(pd.Series([r.get('宿⇔宿 間隔下限')]), errors='coerce').iloc[0]
        night_gap = 0 if pd.isna(gv) else int(gv)
        dv = pd.to_numeric(pd.Series([r.get(DAY_GAP_COLUMN)]), errors='coerce').iloc[0]
        day_gap = night_gap if pd.isna(dv) else int(dv)
        ext = worked_items((past or {}).get(name, [])) + worked_items((future or {}).get(name, []))
        if not ext:
            continue
        marks = {}
        for d in range(1, num_days + 1):
            b = external_gap_block(date_of(d), ext, night_gap, day_gap)
            if any(b):
                marks[d] = b
        if marks:
            blocked[name] = marks
    return blocked


def gap_marks(gap_blocked, name, d, holiday):
    """その日のマスの (日直側が薄紫, 宿直側が薄紫)。平日は日直がないので宿直側だけを見る。"""
    v = (gap_blocked or {}).get(name, {})
    if isinstance(v, dict):
        day_b, night_b = v.get(d, (False, False))
    else:   # 以前の形（日の集まり）
        day_b = night_b = d in v
    return (bool(day_b) and holiday), bool(night_b)


def _add_typed_interval_constraints(model, shifts, doctors, daily_active_shifts, num_days, date_of, min_intervals, day_intervals,
                                    past_worked_dates, future_worked_dates, absolute_req_days, absolute_req_specific):
    """
    勤務間隔の条件（通常版・期間指定版で共通）。
    宿直→宿直は「宿⇔宿 間隔下限」、日直→宿直・宿直→日直・日直→日直は「宿⇔日・日⇔日 間隔下限」の日数を空ける。
    確定同士は維持し、自動割当と全勤務（前後の月の確定当直を含む）の間隔を制限する。
    宿⇔日・日⇔日 間隔下限と宿⇔宿 間隔下限が同じ先生は、以前とまったく同じ条件になる（短いほうの日数までは「どの勤務どうしも」、
    そこから長いほうの日数までは、該当する組み合わせだけを制限する）。
    """
    for doc in doctors:
        night_gap = min_intervals[doc]
        day_gap = (day_intervals or {}).get(doc, night_gap)
        short_gap, long_gap = min(night_gap, day_gap), max(night_gap, day_gap)
        fixed_days = set(absolute_req_days[doc]) | {d for d, s in absolute_req_specific[doc]}
        external = worked_items(past_worked_dates.get(doc, [])) + worked_items(future_worked_dates.get(doc, []))
        worked = {}                       # その日に何か勤務があるか（以前と同じ）
        typed = {}                        # 宿⇔宿 間隔下限と宿⇔日・日⇔日 間隔下限が違うときだけ：その日に宿直（または日直）があるか
        typed_kind = 'night' if night_gap > day_gap else 'day'
        typed_group = NIGHT_SHIFTS if typed_kind == 'night' else DAY_SHIFTS
        for d in range(1, num_days + 1):
            variables = [shifts[d, doc, s] for s in daily_active_shifts.get(d, []) if (d, doc, s) in shifts]
            if not variables:
                continue
            worked[d] = model.NewBoolVar(f'working_{doc}_{d}')
            model.AddMaxEquality(worked[d], variables)
            if long_gap > short_gap:
                tv = [shifts[d, doc, s] for s in daily_active_shifts.get(d, []) if s in typed_group and (d, doc, s) in shifts]
                if tv:
                    typed[d] = model.NewBoolVar(f'working_{typed_kind}_{doc}_{d}')
                    model.AddMaxEquality(typed[d], tv)
            # 前後の月の確定当直との間隔（今月・期間中の確定の日は対象外）
            if d not in fixed_days and external:
                day_block, night_block = external_gap_block(date_of(d), external, night_gap, day_gap)
                if day_block and night_block:
                    model.Add(worked[d] == 0)
                else:
                    for s in daily_active_shifts.get(d, []):
                        if (d, doc, s) in shifts and ((day_block and s in DAY_SHIFTS) or (night_block and s in NIGHT_SHIFTS)):
                            model.Add(shifts[d, doc, s] == 0)
        for d in worked:
            for other in range(d + 1, min(num_days, d + long_gap) + 1):
                if other not in worked or (d in fixed_days and other in fixed_days):
                    continue
                if other - d <= short_gap:
                    model.Add(worked[d] + worked[other] <= 1)
                elif typed_kind == 'night':
                    # 宿直→宿直だけ、宿⇔宿 間隔下限まで空ける
                    if d in typed and other in typed:
                        model.Add(typed[d] + typed[other] <= 1)
                else:
                    # 日直がからむ組み合わせだけ、宿⇔日・日⇔日 間隔下限まで空ける（日直→何か、何か→日直）
                    if d in typed:
                        model.Add(typed[d] + worked[other] <= 1)
                    if other in typed:
                        model.Add(worked[d] + typed[other] <= 1)


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
    day_intervals=None,
):
    """確定同士は維持し、自動割当と全勤務の間隔を制限する（宿直→宿直は宿⇔宿 間隔下限、日直がからむ組み合わせは宿⇔日・日⇔日 間隔下限）。"""
    _add_typed_interval_constraints(model, shifts, doctors, daily_active_shifts, num_days,
                                    lambda d: datetime.date(target_year, target_month, d), min_intervals, day_intervals,
                                    past_worked_dates, future_worked_dates, absolute_req_days, absolute_req_specific)


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


def build_duty_matrix(df_result, staff_df, year, month, custom_holidays, next_month_special_holiday=False, comments=None, gap_blocked=None):
    """
    医師ごと・日ごとのマスの内容を作る。
    優先順：決まった当直 ＞ ×（当直NG）＞ 翌日PM duty で宿直を外す日。
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
                     'ng_count': 0, 'ng_day': 0, 'ng_night': 0})   # その日にNG（翌日PM dutyで宿直を外す先生を含む）の先生の数

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
            # 前後の月の確定当直との勤務間隔で入れない日（薄紫のマス）も、実質NGとして数える
            _kind = effective_ng(ng.get(d, 'OK'), info['holiday'])
            _pmr = pm_duty_restricted(datetime.date(year, month, d), pm_days, year, month, custom_holidays, next_month_special_holiday)
            # （日直と宿直で空ける日数が違うので、日直側・宿直側を分けて見る。平日は宿直側だけ）
            _gday, _gnight = gap_marks(gap_blocked, name, d, info['holiday'])
            _gap = _gday or _gnight
            if _kind != 'OK' or _pmr or _gap:
                info['ng_count'] += 1
            # 休日は、日直NGと宿直NGを分けても数える（日直NG＝全NG・日NG、宿直NG＝全NG・宿NG・翌日PM duty）
            if info['holiday']:
                info['ng_day'] += int(_kind in ('全NG', '日NG') or _gday)
                info['ng_night'] += int(_kind in ('全NG', '宿NG') or _pmr or _gnight)
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
            # 前月末・翌月初めの確定当直との勤務間隔のため入れない日：マスを薄紫にする（×の表示はそのまま）
            # （今月の確定当直の前後は、色を付けない）
            gday, gnight = gap_marks(gap_blocked, name, d, info['holiday'])
            if info['holiday'] and gday != gnight:
                # 休日で、日直か宿直の片方だけ入れない日：入れない側の半分だけを薄紫にする
                # （上半分＝日直、下半分＝宿直。宿直だけ入れる日の翌日PM dutyは、下半分を斜線にする）
                top = '×' if kind in ('全NG', '日NG') else ''
                bottom = '×' if kind in ('全NG', '宿NG') else ''
                cells.append({'type': 'pmhalf' if (pm and not gnight) else 'half', 'top': top, 'bottom': bottom,
                              'text': '×' if (top or bottom) else '', 'gap_top': gday, 'gap_bottom': gnight})
                continue
            if gnight:
                if info['holiday'] and kind in ('全NG', '宿NG', '日NG'):
                    top = '×' if kind in ('全NG', '日NG') else ''
                    bottom = '×' if kind in ('全NG', '宿NG') else ''
                    cells.append({'type': 'half', 'top': top, 'bottom': bottom, 'text': '×', 'gap': True})
                elif kind in ('全NG', '宿NG'):
                    cells.append({'type': 'ng', 'text': '×', 'pm': False, 'gap': True})
                else:
                    cells.append({'type': 'empty', 'text': '', 'gap': True})
                continue
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


def shortage_lines(days):
    """「⚠️ 未決定」の行の並び [(枠, 何段目)]。同じ日に同じ枠が2名以上足りないときは、行を分けて1マス＝1枠にする
    （「B日直×2」のように1マスにまとめると、枠の数え方やExcelでのドラッグがずれるため）。"""
    lines = []
    for slot in shortage_slots(days):
        most = max(sum(x['count'] for x in info['shortage'] if x['slot'] == slot) for info in days)
        lines.extend((slot, j) for j in range(most))
    return lines


def shortage_total(days):
    """未決定の枠の数（1名分＝1枠）。"""
    return sum(x['count'] for info in days for x in info['shortage'])


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


# 医師別の当直一覧の凡例の見た目（当直一覧と、使い方のページで同じものを使う）
# 改行を入れない（使い方のページでは文の途中に入るため、改行があると画面の文として扱われ、そのあとが表示されなくなる）
MATRIX_LEGEND_CSS = ("<style>.dm-legend{display:flex;flex-wrap:wrap;gap:10px;font-size:12px;margin:2px 0 6px;color:#243247}"
                     ".dm-legend span{display:inline-flex;align-items:center;gap:4px}"
                     ".dm-legend i{display:inline-block;min-width:34px;padding:1px 4px;border:1px solid #d8dde5;border-radius:3px;"
                     "font-style:normal;font-weight:700;text-align:center;font-size:11px}</style>")


def matrix_legend_html(alerts=None):
    """医師別の当直一覧の凡例（色・記号の見本と意味）。alerts は回数のアラートの基準（None なら初期設定）。"""
    legend = ['<div class="dm-legend">']
    for s, (bg, fg) in MATRIX_SHIFT_COLORS.items():
        legend.append(f'<span><i style="background:{bg};color:{fg}">{s}</i></span>')
    legend.append(f'<span><i style="box-shadow:inset 0 0 0 3px {MATRIX_WISH_FRAME}">&nbsp;&nbsp;&nbsp;</i>茶色の枠：希望日が通った当直</span>')
    legend.append('<span><i style="background:#fdf1f1">&nbsp;</i>薄い赤の列：土日祝・休日にした日</span>')
    legend.append('<span><i style="background:#e2daf2">&nbsp;</i>薄紫：前後の月（前月末・翌月初めなど）の確定当直との勤務間隔のため、入れない日</span>')
    legend.append('<span><i style="background:linear-gradient(#e2daf2,#e2daf2) left top/100% 50% no-repeat,#fdf1f1;line-height:1">&nbsp;<br>&nbsp;</i>'
                  '上半分だけ薄紫：休日の日直だけ入れない日（下半分だけ薄紫＝宿直だけ入れない日）。日直がからむ組み合わせは「宿⇔日・日⇔日 間隔下限」で空けるため</span>')
    legend.append('<span><i style="color:#c0392b">×</i>当直NG（平日は宿直NG）</span>')
    legend.append('<span><i style="color:#c0392b;background:#fdf1f1;line-height:1;padding:0 4px">×<br>&nbsp;</i>'
                  '休日のマス：上半分＝日直、下半分＝宿直。上の×＝日直NG、下の×＝宿直NG（上下とも×＝全NG）</span>')
    legend.append('<span><i style="background:repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px)">&nbsp;</i>斜線：翌日PM dutyで宿直を外す日（×があればNGも付いている）</span>')
    legend.append('<span><i style="background:repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px) left bottom/100% 50% no-repeat,#fdf1f1">&nbsp;</i>'
                  '下半分だけ斜線：休日の翌日PM dutyの日。宿直は外し、日直には入れる（×の意味は休日のマスと同じ）</span>')
    legend.append('<span><i>💬📝</i>メモ欄：先生からのコメント・備考（カーソルを合わせると全文）</span>')
    legend.append('<span><i style="color:#c0392b;border-color:#c0392b">⚠️ 未決定</i>赤線の下：当直が決まっていない枠</span>')
    legend.append('<span><i style="background:#ffd2a1;color:#7a3a00">3</i>NG人数の行：その日にNGの先生の数（翌日PM dutyで宿直を外す先生・薄紫のマスの先生を含む）。休日は上＝日直NG・下＝宿直NGの人数</span>')
    if count_alert_text(alerts):
        legend.append('<span><i style="color:#c0392b;background:#ffd8d8;border-color:#e03131">⚠</i>右端の赤いマス：' + count_alert_text(alerts) + '</span>')
    legend.append('</div>')
    return "".join(legend)


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
/* 前月末・翌月初めの確定当直との勤務間隔のため入れない日 */
.dm td.gap{background-color:#e2daf2}
.dm .cell{display:flex;align-items:center;justify-content:center;height:100%;min-height:26px;font-weight:700;font-size:10px;padding:1px 2px;width:38px;margin:0 auto;white-space:normal;line-height:1.15;text-align:center;word-break:keep-all}
.dm .ng{color:#c0392b;font-size:14px}
.dm .pm{background:repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px)}
.dm td.pmhalf{position:relative;min-width:38px;background:repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px) left bottom/100% 50% no-repeat,#fdf1f1}
.dm td.pmhalf span{position:absolute;left:0;right:0;height:50%;display:flex;align-items:center;justify-content:center;color:#c0392b;font-weight:700;font-size:13px;line-height:1}
.dm td.pmhalf span.t{top:0} .dm td.pmhalf span.b{bottom:0}
.dm td.half{position:relative;min-width:38px}
.dm td.half span{position:absolute;left:0;right:0;height:50%;display:flex;align-items:center;justify-content:center;color:#c0392b;font-weight:700;font-size:13px;line-height:1}
.dm td.half span.t{top:0} .dm td.half span.b{bottom:0}
.dm td.ngh{position:relative;min-width:38px;height:34px;color:#7a3a00;font-weight:700}
.dm td.ngh span{position:absolute;left:0;right:0;height:50%;display:flex;align-items:center;justify-content:center;font-size:11px;line-height:1}
.dm td.ngh span.t{top:0}.dm td.ngh span.b{bottom:0}
.dm td.ngh::after{content:"";position:absolute;left:5px;right:5px;top:50%;border-top:1px dashed #d9a066}
.dm td.half::after{content:"";position:absolute;left:5px;right:5px;top:50%;border-top:1px dashed #dcb8b8}
/* 行の境目を見やすくする：先生の行を1行おきに薄いグレーにする（休日の色・斜線などの上に薄く重ねる） */
/* 右端の回数の列（色をマスごとに指定している）まで同じ縞にするため !important を付ける */
.dm tbody tr:not(.ngcount):nth-child(odd) td{background-image:linear-gradient(rgba(36,64,95,.075),rgba(36,64,95,.075)) !important}
/* 休日の翌日PM dutyのマス（下半分が斜線）は、縞の行でも斜線を消さない */
.dm tbody tr:not(.ngcount):nth-child(odd) td.pmhalf{background-image:linear-gradient(rgba(36,64,95,.075),rgba(36,64,95,.075)),
  repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px) !important;
  background-size:100% 50%,100% 50% !important;background-position:left top,left bottom !important;background-repeat:no-repeat !important}
/* 休日の、日直か宿直の片方だけが前後の月の確定当直のため入れない日：入れない側の半分だけ薄紫（上＝日直、下＝宿直） */
.dm td.half.gtop{background:linear-gradient(#e2daf2,#e2daf2) left top/100% 50% no-repeat,#fdf1f1}
.dm td.half.gbot{background:linear-gradient(#e2daf2,#e2daf2) left bottom/100% 50% no-repeat,#fdf1f1}
.dm td.pmhalf.gtop{background:linear-gradient(#e2daf2,#e2daf2) left top/100% 50% no-repeat,
  repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px) left bottom/100% 50% no-repeat,#fdf1f1}
.dm tbody tr:not(.ngcount):nth-child(odd) td.half.gtop,.dm tbody tr:not(.ngcount):nth-child(odd) td.half.gbot{
  background-image:linear-gradient(rgba(36,64,95,.075),rgba(36,64,95,.075)),linear-gradient(#e2daf2,#e2daf2) !important;
  background-size:100% 100%,100% 50% !important;background-repeat:no-repeat !important}
.dm tbody tr:not(.ngcount):nth-child(odd) td.half.gtop{background-position:0 0,left top !important}
.dm tbody tr:not(.ngcount):nth-child(odd) td.half.gbot{background-position:0 0,left bottom !important}
.dm tbody tr:not(.ngcount):nth-child(odd) td.pmhalf.gtop{background-image:linear-gradient(rgba(36,64,95,.075),rgba(36,64,95,.075)),
  linear-gradient(#e2daf2,#e2daf2),repeating-linear-gradient(45deg,#f2c75c 0 2px,#fff5d9 2px 7px) !important;
  background-size:100% 100%,100% 50%,100% 50% !important;background-position:0 0,left top,left bottom !important;background-repeat:no-repeat !important}
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
    legend = [matrix_legend_html(alerts)]

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
    def ng_bg(cnt):
        ratio = cnt / n_doctors
        return '#ffffff' if ratio == 0 else ('#ffefdc' if ratio < 0.25 else ('#ffd2a1' if ratio < 0.5 else '#ffad5c'))
    for info in days:
        if info['holiday']:
            # 休日：上半分＝日直NGの人数、下半分＝宿直NGの人数（当直案の休日のマスと同じ並び）
            top, bot = info.get('ng_day', 0), info.get('ng_night', 0)
            ng_tr.append(f'<td class="c{info["day"]} ngh" style="background:linear-gradient({ng_bg(top)} 0 50%,{ng_bg(bot)} 50% 100%)" '
                         f'title="日直NG {top}名・宿直NG {bot}名"><span class="t">{top}</span><span class="b">{bot}</span></td>')
        else:
            ng_tr.append(f'<td class="c{info["day"]}" style="background:{ng_bg(info.get("ng_count", 0))};text-align:center;font-weight:700;color:#7a3a00">{info.get("ng_count", 0)}</td>')
    ng_tr.append(f'<td class="sum" colspan="{len(MATRIX_COUNT_COLUMNS)}" style="font-size:11px;background:#ffe2c2;color:#7a3a00;font-weight:800">／{len(rows)}名中</td></tr>')
    body.append(''.join(ng_tr))
    for idx, row in enumerate(rows):
        tip_up = idx >= max(len(rows) - 6, 6)
        tr = [f'<tr><td class="nm">{html_escape(row["name"])}</td><td class="pmcol">{html_escape(row["pm_duty"])}</td>{memo_cell_html(row, tip_up)}']
        for info, cell in zip(days, row['cells']):
            col = f'c{info["day"]}'   # 列（日付）ごとの印。カーソルを合わせた列を色付けするのに使う
            gap_class = ' gap' if cell.get('gap') else ''
            td_class = f' class="{col} hol{gap_class}"' if info['holiday'] else f' class="{col}{gap_class}"'
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
                hol_class = (' hol' if cell['type'] == 'half' else '') + (' gap' if cell.get('gap') else '') \
                    + (' gtop' if cell.get('gap_top') else '') + (' gbot' if cell.get('gap_bottom') else '')
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
    for k, (slot, j) in enumerate(shortage_lines(days)):
        first = ' first' if k == 0 else ''
        tr = [f'<tr class="short{first}"><td class="nm">⚠️ 未決定</td><td></td><td></td>']
        bg, fg = MATRIX_SHIFT_COLORS[slot]
        slot_total = 0
        for info in days:
            count = sum(x['count'] for x in info['shortage'] if x['slot'] == slot)
            if count > j:
                slot_total += 1
                label = slot_label(slot, "<br>")
                tr.append(f'<td class="c{info["day"]}"><div class="cell" style="background:{bg};color:{fg};white-space:normal;line-height:1.15;font-size:10px">{label}</div></td>')
            else:
                tr.append(f'<td class="c{info["day"]}"></td>')
        tr.append(f'<td class="sum" colspan="{len(MATRIX_COUNT_COLUMNS)}" style="color:#c0392b">{slot_total}枠</td></tr>')
        body.append(''.join(tr))
    # カーソルを合わせたセルの「行」と「列」を、オレンジの線で囲んで示す（縦横に長い表での見間違いを防ぐ）
    # マスの色は変えない（薄紫・休日の薄い赤・斜線などがそのまま見える）。線はマスの中身より手前に、::before で描く
    C = "#e8890c"
    line = f"content:'';position:absolute;pointer-events:none;"
    hover_css = ("<style>"
                 # 行：上下の線（左端は名前のマス、右端は最後のマス）
                 ".dm tbody tr:hover td:not(.nm){position:relative}"
                 f".dm tbody tr:hover td:not(.nm)::before{{{line}left:0;right:0;top:-1px;bottom:-1px;border-top:2px solid {C};border-bottom:2px solid {C}}}"
                 f".dm tbody tr:hover td:last-child::before{{border-right:2px solid {C}}}"
                 f".dm tbody tr:hover td.nm{{background-color:#ffe08a !important;box-shadow:inset 2px 0 0 {C},inset 0 2px 0 {C},inset 0 -2px 0 {C}}}"
                 # 列：左右の線（上端は日付の見出し、下端は最後の行）
                 + "".join(
                     f".dm:has(.c{d}:hover) td.c{d}{{position:relative}}"
                     f".dm:has(.c{d}:hover) td.c{d}::before{{{line}top:0;bottom:0;left:-1px;right:-1px;border-left:2px solid {C};border-right:2px solid {C}}}"
                     f".dm:has(.c{d}:hover) tbody tr:last-child td.c{d}::before{{border-bottom:2px solid {C}}}"
                     f".dm:has(.c{d}:hover) th.c{d}{{background-color:#ffe08a !important;box-shadow:inset 2px 0 0 {C},inset -2px 0 0 {C},inset 0 2px 0 {C}}}"
                     for d in (info['day'] for info in days))
                 # カーソルのあるマス：4辺を太めに囲む
                 + f".dm tbody td:hover{{outline:3px solid {C};outline-offset:-2px}}"
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


# 掲示用Excelの見出し（当直案の枠 → 掲示用の病棟名）。ICU病棟（循環器）は別に作られるので、空けておいて手で入れる
def _posting_color(slot):
    """先生ごとの一覧と同じ色（背景色, 文字色）を、Excel用の形（#なしの大文字）にする。"""
    bg, fg = MATRIX_SHIFT_COLORS[slot]
    return bg.lstrip("#").upper(), fg.lstrip("#").upper()


POSTING_GROUPS = [
    ("病棟", "14階、13階、12階、8階病棟", [("A日直", "A日直", *_posting_color("A日直")), ("A宿直", "A宿直", *_posting_color("A宿直"))]),
    ("病棟", "11階、10階、9階、救急病棟", [("B日直", "B日直", *_posting_color("B日直")), ("B宿直", "B宿直", *_posting_color("B宿直"))]),
    ("外来", "", [("外来日直", "外来日直", *_posting_color("外来日直")), ("外来宿直", "外来宿直", *_posting_color("外来宿直"))]),
    ("循環器", "ICU病棟", [(None, "ICU日直", "FFFFFF", "000000"), (None, "ICU宿直", "FFFFFF", "000000")]),
]


POSTING_DATA_SHEET = "集計用"   # 掲示用Excelの見えないシート（年間集計で読む：作成日時・先生の一覧・休日）


def posting_xlsx(df_result, year, month, title, roster=None, created=None):
    """
    掲示用のExcel（病院に貼る形）を作る。
    ・日付と曜日、病棟（A・B）・外来・ICU病棟（循環器）の日直と宿直を並べる
    ・休日の行はだいだい色、平日の日直のマスは灰色（平日は日直がない）
    ・ICU病棟の列と、右上の作成日は空けておき、あとから手で入れる
    ・まだ決まっていない枠（不足）は「未決定」と書く（画面・当直一覧と同じ言葉）
    ・色は先生ごとの一覧と同じ。A3横1枚に収まる大きさにする
    ・年間集計のために、見えないシートに作成日時・その月の先生の一覧・休日を入れておく
      （roster はその月の先生の名前、created は当直案を作った日時）
    """
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "掲示用"
    font_name = "游ゴシック"
    side = Side(style="thin", color="7F7F7F")
    border = Border(left=side, right=side, top=side, bottom=side)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    fill = lambda c: PatternFill("solid", fgColor=c)
    slots = [s for _, _, items in POSTING_GROUPS for s in items]
    last_col = 2 + len(slots)
    last_letter = get_column_letter(last_col)

    def box(r1, c1, r2, c2, text, bg="F2F2F2", color="000000", bold=False, size=13):
        if (r1, c1) != (r2, c2):
            ws.merge_cells(start_row=r1, start_column=c1, end_row=r2, end_column=c2)
        cell = ws.cell(r1, c1, text)
        cell.font = Font(name=font_name, bold=bold, color=color, size=size)
        cell.alignment = center
        for r in range(r1, r2 + 1):
            for c in range(c1, c2 + 1):
                ws.cell(r, c).border = border
                ws.cell(r, c).fill = fill(bg)

    ws.merge_cells(f"A1:{last_letter}1")
    ws["A1"] = title
    ws["A1"].font = Font(name=font_name, size=22, bold=True)
    ws["A1"].alignment = center
    ws.row_dimensions[1].height = 36
    ws.merge_cells(start_row=2, start_column=last_col - 3, end_row=2, end_column=last_col)
    note = ws.cell(2, last_col - 3, "　　　　年　　月　　日（掲示時点での情報です）")
    note.font = Font(name=font_name, size=11)
    note.alignment = Alignment(horizontal="right", vertical="center")

    # 見出し（3〜5行目）
    box(3, 1, 5, 2, "日付", bold=True)
    col = 3
    groups = []
    for top, sub, items in POSTING_GROUPS:
        if groups and groups[-1][0] == top:
            groups[-1][2] += len(items)
        else:
            groups.append([top, col, len(items)])
        if sub:
            box(4, col, 4, col + len(items) - 1, sub, bg="FFFFFF")
        for i, (_, label, bg, color) in enumerate(items):
            box(5, col + i, 5, col + i, label, bg=bg, color=color, bold=True)
        col += len(items)
    for top, c1, n in groups:
        has_sub = any(sub for t, sub, _ in POSTING_GROUPS if t == top)
        box(3, c1, 3 if has_sub else 4, c1 + n - 1, top, bold=True)

    for r in (3, 4, 5):
        ws.row_dimensions[r].height = 24

    # 日ごとの行
    r = 6
    for _, row in df_result.iterrows():
        dt = parse_shift_date(row.get("日付", ""), year, month)
        holiday = clean_text(row.get("平日/休日", "")) == "休日"
        ws.cell(r, 1, f"{dt.year}/{dt.month}/{dt.day}" if dt else clean_text(row.get("日付", "")))
        ws.cell(r, 2, WEEKDAYS_JA[dt.weekday()] if dt else "")
        for i, (slot, label, _, _) in enumerate(slots):
            value = clean_text(row.get(slot, "")) if slot else ""
            value = re.sub(r"⚠️?不足\(\d+名\)", "未決定", value)
            ws.cell(r, 3 + i, "" if value in ("-", "nan", "None") else value)
        for c in range(1, last_col + 1):
            cell = ws.cell(r, c)
            cell.border = border
            cell.alignment = center
            cell.font = Font(name=font_name, size=14, color="C00000" if (c == 2 and holiday) else "000000")
            if holiday:
                cell.fill = fill("F8CBAD")
            elif c >= 3 and slots[c - 3][1].endswith("日直"):
                cell.fill = fill("EDEDED")   # 平日は日直がない
        # 名前が長くて（2人以上の枠など）1行に入らないときは、2行になるので行を高くする
        many = any(len(str(ws.cell(r, c).value or "")) > 10 for c in range(3, last_col + 1))
        ws.row_dimensions[r].height = 40 if many else 23
        r += 1

    ws.column_dimensions["A"].width = 18
    ws.column_dimensions["B"].width = 5
    for c in range(3, last_col + 1):
        ws.column_dimensions[get_column_letter(c)].width = 25
    ws.freeze_panes = "C6"
    ws.print_title_rows = "1:5"
    ws.page_setup.orientation = "landscape"
    ws.page_setup.paperSize = ws.PAPERSIZE_A3
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    # 年間集計用の見えないシート（当直のない先生も「在籍」として数えるため、先生の一覧も入れる）
    ds = wb.create_sheet(POSTING_DATA_SHEET)
    ds["A1"] = "このシートは、アプリの年間集計で使います。消したり書きかえたりしないでください。"
    ds["A2"], ds["B2"] = "作成日時", created or remote_now()
    ds["A4"], ds["B4"], ds["D4"] = "日付", "平日/休日", "先生の一覧"
    r = 5
    for _, row in df_result.iterrows():
        dt = parse_shift_date(row.get("日付", ""), year, month)
        if dt:
            ds.cell(r, 1, dt.isoformat())
            ds.cell(r, 2, clean_text(row.get("平日/休日", "")))
            r += 1
    for i, name in enumerate(n for n in (roster or []) if clean_text(n)):
        ds.cell(5 + i, 4, clean_text(name))
    ds.sheet_state = "hidden"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


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

    def half_fill(top_color, bottom_color):
        """上半分と下半分で色が変わる塗り（Excelはマスの半分だけを塗れないため）。"""
        return GradientFill(type='linear', degree=90, stop=(Stop(top_color, 0), Stop(top_color, 0.45),
                                                            Stop(bottom_color, 0.6), Stop(bottom_color, 1)))

    ws.cell(row=1, column=1, value=f"{year}年{month}月分 当直表" if title is None else f"{title} 当直表").font = Font(bold=True, size=14)
    header_row, first_day_col = 3, 4
    ws.cell(row=header_row, column=1, value='名前')
    ws.cell(row=header_row, column=2, value='翌日PM duty')
    ws.cell(row=header_row, column=3, value='メモ')
    for i, info in enumerate(days):
        c = ws.cell(row=header_row, column=first_day_col + i, value=f"{info.get('label', info['day'])}\n{info['weekday']}")
        color = MATRIX_HOLIDAY_HEAD if info['kind'] == 'holiday' else (MATRIX_SATURDAY_HEAD if info['kind'] == 'saturday' else '#243247')
        # 期間指定版の「12/10」のような長い日付は、列の幅に入るよう文字を小さくする（入らないと3段に折れるため）
        label_len = len(str(info.get('label', info['day'])))
        c.font = Font(bold=True, color=color.lstrip('#'), size=7.5 if label_len >= 5 else (9 if label_len == 4 else 11))
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
        if info['holiday']:
            # 休日は「日直NG／宿直NG」の人数を2段で書く（色は多いほうに合わせる）
            c = ws.cell(row=ng_row, column=first_day_col + i, value=f"{info.get('ng_day', 0)}\n{info.get('ng_night', 0)}")
            cnt = max(info.get('ng_day', 0), info.get('ng_night', 0))
        else:
            c = ws.cell(row=ng_row, column=first_day_col + i, value=cnt)
        ratio = cnt / n_doctors
        c.fill = fill('#FFFFFF' if ratio == 0 else ('#FFEFDC' if ratio < 0.25 else ('#FFD2A1' if ratio < 0.5 else '#FFAD5C')))
        c.font = Font(bold=True, color='7A3A00', **({'size': 8} if info['holiday'] else {}))
        c.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True) if info['holiday'] else center
        c.border = border
    for col in range(1, first_day_col):
        ws.cell(row=ng_row, column=col).border = border
    ws.cell(row=ng_row, column=sum_col, value=f"／{len(rows)}名中").font = Font(bold=True, size=9, color='7A3A00')
    for j in range(n_counts):
        ws.cell(row=ng_row, column=sum_col + j).fill = fill('#FFE2C2')
    ws.row_dimensions[ng_row].height = 26 if any(info['holiday'] for info in days) else 20

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
            if cell.get('gap'):
                c.fill = fill('#E2DAF2')   # 前月末・翌月初めの確定当直との勤務間隔のため入れない日
            elif cell.get('gap_top') or cell.get('gap_bottom'):
                # 休日の、日直か宿直の片方だけ入れない日：入れない側（上＝日直、下＝宿直）の半分だけ薄紫にする
                other = PM_COLOR if cell['type'] == 'pmhalf' else 'FDF1F1'
                top_c, bottom_c = ('E2DAF2', other) if cell.get('gap_top') else ('FDF1F1', 'E2DAF2')
                c.fill = half_fill(top_c, bottom_c)
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
                if getattr(c.fill, 'tagname', '') == 'gradientFill':   # 途中で色が変わる塗り（c.fill は包まれているので型では判定できない）
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
    for k, (slot, j) in enumerate(shortage_lines(days)):
        last_row += 1
        bg, fg = MATRIX_SHIFT_COLORS[slot]
        ws.cell(row=last_row, column=1, value='⚠️ 未決定').font = Font(bold=True, color='C0392B')
        slot_total = 0
        for i, info in enumerate(days):
            count = sum(x['count'] for x in info['shortage'] if x['slot'] == slot)
            c = ws.cell(row=last_row, column=first_day_col + i)
            c.border = border
            if count > j:
                slot_total += 1
                c.value = slot
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

    # 動かない列（名前・翌日PM duty・メモ）と、横にスクロールすると動く日付の列の境目に、太い紺の線を引く
    # （メモの列の右側に付ける。日付のマスに付けると、Excelでマスをドラッグして移したときに線までついて行くため）
    edge = Side(style='thick', color='24405F')
    for r in range(header_row, last_row + 1):
        c = ws.cell(row=r, column=first_day_col - 1)
        b0 = c.border
        c.border = Border(left=b0.left, top=b0.top, bottom=b0.bottom, right=edge)
    # 縦にスクロールしても動かない行（題・日付の見出し・NG人数）と先生の行の境目にも、同じ線を引く
    # （NG人数の行の下側に付ける。先生の行のマスに付けると、ドラッグしたときに線までついて行くため）
    for col in range(1, sum_col + n_counts):
        c = ws.cell(row=header_row + 1, column=col)
        b0 = c.border
        c.border = Border(left=b0.left, right=b0.right, top=b0.top, bottom=edge)

    # 凡例は別のシート「凡例」（当直表の右隣）に置き、当直表の下には案内を1行だけ書く
    # （当直表のシートに置くと、横にスクロールしたときに説明が切れて見えるため）
    note = ws.cell(row=last_row + 2, column=1, value='※ 色や記号の意味は「凡例」シート')
    note.font = Font(bold=True, color='1D4E89'); note.alignment = Alignment(vertical='center')
    lgs = wb.create_sheet("凡例")
    lgs.sheet_view.showGridLines = False
    one_line = Alignment(horizontal='center', vertical='center', wrap_text=False, shrink_to_fit=True)
    middle = Alignment(vertical='center')   # 説明の文字は、見本のマスと高さをそろえて上下中央にする
    lgs.cell(row=1, column=1, value='凡例（当直表の色・記号の意味）').font = Font(bold=True, size=15)
    lgs.row_dimensions[1].height = 30
    lgs.cell(row=3, column=1, value='当直の枠の色').font = Font(bold=True)
    # 当直の枠の色：宿直3つ・日直3つの2段で並べる
    for k, (label, (bg, fg)) in enumerate(MATRIX_SHIFT_COLORS.items()):
        c = lgs.cell(row=4 + k // 3, column=1 + k % 3, value=label)
        c.fill = fill(bg); c.font = Font(bold=True, size=10, color=fg.lstrip('#')); c.alignment = one_line; c.border = border
    for r in (4, 5):
        lgs.row_dimensions[r].height = 22
    legend_row = 6   # この下に「見本のマス＋説明」を1行ずつ並べる

    def legend_line(offset, text):
        r = legend_row + offset
        lgs.cell(row=r, column=2, value=text).alignment = middle
        lgs.row_dimensions[r].height = 24
        sample = lgs.cell(row=r, column=1)
        sample.border = border; sample.alignment = one_line
        return sample

    s1 = legend_line(1, '×：当直NG（平日は宿直NG）。休日のマスは上＝日直、下＝宿直で、上の×＝日直NG、下の×＝宿直NG（上下とも×＝全NG）')
    s1.value = '×'; s1.font = Font(bold=True, color='C0392B')
    legend_line(2, '黄色：翌日PM dutyで宿直を外す日（×があればNGも付いている）').fill = PM_FILL
    legend_line(4, '薄い赤の列：土日祝・休日にした日').fill = fill('#FDF1F1')
    s4 = legend_line(5, '茶色の太い枠：希望日が通った当直')
    # 枠が見えやすいよう、淡い色の枠名を見本にする
    sample_slot = 'B日直' if 'B日直' in MATRIX_SHIFT_COLORS else next(iter(MATRIX_SHIFT_COLORS))
    first_bg, first_fg = MATRIX_SHIFT_COLORS[sample_slot]
    s4.value = sample_slot; s4.fill = fill(first_bg)
    s4.font = Font(bold=True, size=9, color=first_fg.lstrip('#')); s4.border = wish_border
    s5 = legend_line(6, '赤線より下の「⚠️ 未決定」の行：当直が決まっていない枠（その枠の色で表示）')
    s5.border = Border(left=thin, right=thin, bottom=thin, top=red)
    s5.value = '未決定'; s5.font = Font(bold=True, size=9, color='C0392B')
    s6 = legend_line(7, 'メモ：💬＝先生からのコメント、📝＝備考（セルにカーソルを乗せると全文。「メモ全文」シートにも一覧）')
    s6.value = '💬📝'; s6.border = Border()
    s8 = legend_line(8, 'NG人数の行（見出しのすぐ下）：その日にNGの先生の数（翌日PM dutyで宿直を外す先生・薄紫のマスの先生を含む）。休日は「日直NG／宿直NG」の人数')
    s8.value = 3; s8.fill = fill('#FFD2A1'); s8.font = Font(bold=True, color='7A3A00')
    legend_line(10, '薄紫：前後の月（前月末・翌月初めなど）の確定当直との勤務間隔のため、入れない日').fill = fill('#E2DAF2')
    legend_line(11, '上半分だけ薄紫：休日の日直だけ入れない日（下半分だけ薄紫＝宿直だけ入れない日）。日直がからむ組み合わせは「宿⇔日・日⇔日 間隔下限」で空けるため').fill = half_fill('E2DAF2', 'FDF1F1')
    legend_line(3, '下半分だけ黄色：休日の翌日PM dutyの日。宿直は外し、日直には入れる（×の意味は休日のマスと同じ）').fill = PM_HALF_FILL
    if count_alert_text(alerts):
        s7 = legend_line(9, '右端の赤い枠の回数：' + count_alert_text(alerts))
        s7.value = next(v for k, v in alerts.items() if v) if alerts else 4
        s7.fill = fill('#FFD8D8'); s7.font = Font(bold=True, color='C0392B')
        s7.border = Border(left=red_side, right=red_side, top=red_side, bottom=red_side)
    for col in 'ABC':
        lgs.column_dimensions[col].width = 13
    # 凡例のシートは、A4横向き・1ページに収めて印刷
    lgs.page_setup.orientation = 'landscape'
    lgs.page_setup.paperSize = lgs.PAPERSIZE_A4
    lgs.sheet_properties.pageSetUpPr.fitToPage = True
    lgs.page_setup.fitToWidth = 1
    lgs.page_setup.fitToHeight = 1

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
        return df, ['医師条件のファイルに「先生の名前」の列がありません。']
    df['先生の名前'] = df['先生の名前'].map(clean_text)
    blank = df['先生の名前'].eq('')
    for i, row in df[blank].iterrows():
        if any(clean_text(v) for k, v in row.items() if k != '先生の名前'):
            errors.append(f'医師条件 {i+1}行目：条件が入力されていますが先生の名前が空欄です。')
    df = df[~blank].reset_index(drop=True)
    for name in df.loc[df['先生の名前'].duplicated(), '先生の名前'].unique():
        errors.append(f'先生の名前「{name}」が重複しています。')
    df = mark_soft_slot_caps(df)   # 枠ごとの上限の空欄（できれば2回まで）を、2が入る前に覚えておく
    for col, default in NUMERIC_DEFAULTS.items():
        if col not in df:
            df[col] = default
    for col in [NG_COLUMN, REQUEST_COLUMN, '翌日PM duty', DAY_GAP_COLUMN] + list(GROUP_CAP_COLUMNS) + list(DOC_PREF_COLUMNS):
        if col not in df:
            df[col] = ''
    errors += doc_pref_errors(df)
    for col in list(GROUP_CAP_COLUMNS) + [DAY_GAP_COLUMN]:
        df[col] = df[col].astype(object)   # 空欄（制限なし／宿⇔宿 間隔下限と同じ）と数字が混ざるため
    for i, row in df.iterrows():
        name = row['先生の名前']
        if any(c in name for c in ',、') or name == '-' or '⚠️不足' in name:
            errors.append(f'先生の名前「{name}」：カンマ・読点・不足表示・単独の「-」は使用できません。')
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
            if float(df.at[i, '月下限']) > float(df.at[i, '月上限']):
                errors.append(f'{name}：月下限は月上限以下にしてください。')
        except (ValueError, TypeError):
            pass
        # 宿⇔日・日⇔日 間隔下限は空欄なら「宿⇔宿 間隔下限」と同じ値にする
        raw = clean_text(row.get(DAY_GAP_COLUMN, ''))
        if raw.lower() in ('', 'nan', 'none', '<na>'):
            df.at[i, DAY_GAP_COLUMN] = df.at[i, '宿⇔宿 間隔下限']
        else:
            try:
                n = float(raw)
                if not math.isfinite(n) or n < 0 or n != int(n) or n > 1000000:
                    raise ValueError()
                df.at[i, DAY_GAP_COLUMN] = int(n)
            except (TypeError, ValueError, OverflowError):
                errors.append(f'{name}：宿⇔日・日⇔日 間隔下限は0以上の整数にしてください（空欄なら宿⇔宿 間隔下限と同じ）。')
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
        return df, ['確定当直のファイルに「日付」の列がありません。']
    unknown = [c for c in df if c not in ['日付', '平日/休日', FIXED_KUBUN_COLUMN] + ALL_SHIFT_TYPES]
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
            errors.append(f'確定当直 {i+1}行目：日付「{raw}」を確認してください。「{month}/1」または「{year}/{month}/1」の形で入力します。')
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
                # 前月・翌月の日付は、医師条件にいない先生（前月だけ勤務した先生など）でもよい（勤務間隔の計算には使わない）
                if name not in doctors and (dt.year, dt.month) == (year, month):
                    errors.append(f'確定当直 {dt.month}/{dt.day} {s}：先生の名前「{name}」が医師条件の名前と一致しません。今月の日付には、医師条件の先生を入れてください。')
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
def close_fixed_duty_warnings(name, worked, past, future, night_gap, day_gap, in_range):
    """
    確定当直同士の間隔が短すぎるときの注意。宿直→宿直は「宿⇔宿 間隔下限」、日直がからむ組み合わせは「宿⇔日・日⇔日 間隔下限」と比べる。
    worked：作った当直案の (日付, 'day'|'night')。past・future：前後の月の確定当直。in_range：作る月（期間）の日なら True。
    """
    items = sorted(set(worked_items(worked)) | set(worked_items((past or {}).get(name, []))) | set(worked_items((future or {}).get(name, []))))
    label = {'day': '日直', 'night': '宿直'}
    out = []
    # となり合う勤務どうし、と、宿直だけを並べたときのとなり合う宿直どうしを比べる（同じ組は1回だけ）
    nights = [x for x in items if x[1] == 'night']
    pairs = list(dict.fromkeys(list(zip(items, items[1:])) + list(zip(nights, nights[1:]))))
    pairs.sort()
    for (a, ka), (b, kb) in pairs:
        if b == a:
            continue   # 同じ日の日直と宿直（確定で1日2回）は、間隔の注意の対象にしない
        need = gap_needed(ka, kb, night_gap, day_gap)
        if (b - a).days > need or not (in_range(a) or in_range(b)):
            continue
        col = '宿⇔宿 間隔下限' if (ka, kb) == ('night', 'night') else '宿⇔日・日⇔日 間隔下限'
        out.append(f'{name}：確定当直同士の間隔が「{col}」より短くなっています（{a.month}/{a.day} {label[ka]}→{b.month}/{b.day} {label[kb]}、'
                   f'空き{(b - a).days - 1}日／設定{need}日）。確定当直のまま残しました。')
    return out


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
                worked.append((dt, duty_kind(slot))); counts[slot] += 1
                hol_count += int(r['平日/休日'] == '休日')
                kind = ng.get(dt.day, 'OK')
                night = slot in NIGHT_SHIFTS
                if kind == '全NG' or (kind == '宿NG' and night) or (kind == '日NG' and not night):
                    warnings.append(f'{name}：{dt.month}/{dt.day} {slot}は確定当直のため、NG日より優先して入れました。')
                if night and pm_duty_restricted(dt, pm_weekdays, year, month, holidays, next_month_special_holiday):
                    warnings.append(f'{name}：{dt.month}/{dt.day} {slot}は確定当直のため、翌日PM dutyより優先して入れました。')
        total = sum(counts.values())
        if total < int(row['月下限']):
            warnings.append(f'{name}：月下限の目標{int(row["月下限"])}回に対し、{total}回です。')
        group_items = [(col, sum(counts[s] for s in slots), int(row[col])) for col, slots in GROUP_CAP_COLUMNS.items()
                       if clean_text(row.get(col, '')) not in ('', 'nan')]
        for label, value, cap in [('月上限', total, int(row['月上限'])), ('休日上限', hol_count, int(row['休日上限']))] + group_items + [(s + '上限', counts[s], int(row[s + '上限'])) for s in ALL_SHIFT_TYPES if s not in soft_cap_slots(row)]:
            if value > cap:
                warnings.append(f'{name}：確定当直のため{staff_column_specs().get(label, (label,))[0]}を超えています（設定{cap}回／結果{value}回）。')
        warnings += close_fixed_duty_warnings(name, worked, past, future, int(row['宿⇔宿 間隔下限']), int(row.get(DAY_GAP_COLUMN, row['宿⇔宿 間隔下限'])),
                                              lambda dt: (dt.year, dt.month) == (year, month))
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


def _generate_shift_core(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None, locked_days=(), annual_balance=None):
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
    day_intervals = {}   # 日直がからむ勤務の間隔
    min_shifts_total = {}
    max_shifts_total = {}
    max_hol_shifts_per_doc = {}
    max_shifts_per_type = {}
    soft_caps = {}   # 枠ごとの上限が空欄の枠（できれば2回まで）
    group_caps = {}

    absolute_req_days = {doc: [] for doc in doctors}
    absolute_req_specific = {doc: [] for doc in doctors}
    past_worked_dates = {doc: [] for doc in doctors}
    future_worked_dates = {doc: [] for doc in doctors}

    # --- 確定当直（前月・今月・翌月）の読み取り ---
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
                            past_worked_dates[doc_val].append((date_obj, duty_kind(s_type)))   # 日付と、日直か宿直か
                        else:
                            future_worked_dates[doc_val].append((date_obj, duty_kind(s_type)))

    # --- 医師条件の読み取り ---
    for _, row in staff_df.iterrows():
        doc = str(row['先生の名前'])
        hard_weekdays[doc] = pm_duty_weekdays(row.get('翌日PM duty', ''))
        ng_days_dict[doc] = parse_ng_dict(row.get(NG_COLUMN, ''), target_year, target_month)
        req_days[doc], req_specific[doc] = parse_requests(row.get(REQUEST_COLUMN, ''), target_year, target_month)
        req_days[doc] = [d for d in req_days[doc] if d not in locked_days]
        req_specific[doc] = [(d, s) for d, s in req_specific[doc] if d not in locked_days]
        req_priority[doc] = safe_int(row.get('希望優先度(数字が大きいほど優先)'), 1)
        min_intervals[doc] = safe_int(row.get('宿⇔宿 間隔下限'), 5)
        day_intervals[doc] = safe_int(row.get(DAY_GAP_COLUMN), min_intervals[doc])
        min_shifts_total[doc] = safe_int(row.get('月下限'), 0)
        max_shifts_total[doc] = safe_int(row.get('月上限'), 5)
        max_hol_shifts_per_doc[doc] = safe_int(row.get('休日上限'), 4)
        max_shifts_per_type[doc] = {s: safe_int(row.get(s + '上限'), 2) for s in ALL_SHIFT_TYPES}
        soft_caps[doc] = soft_cap_slots(row)
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
                if not active_days or s_type in soft_caps[doc]:   # 空欄の枠は決まりにせず、目的の減点で「できれば2回まで」にする
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
            absolute_req_specific=absolute_req_specific, day_intervals=day_intervals,
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

    # 月下限（目標。届かない分は減点）
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

    # グループの設定で△（できれば入れない）の枠は、入れるたびに減点する（人が足りないときだけ入る）
    objective_terms += group_avoid_terms(doctors, days, daily_active_shifts, shifts, staff_df)
    # 枠ごとの上限が空欄の枠は、できれば2回まで（1つの枠に偏らないように）
    objective_terms += soft_cap_terms(model, doctors, days, daily_active_shifts, shifts, soft_caps)
    # 回数の勾配：枠に余裕があるとき、グループごとに回数を多め・少なめにする（対象外のグループは入れない）
    objective_terms += group_level_terms(model, doctors, worked_all, staff_df)
    # 多め希望（病棟/外来・宿直/日直・平日/休日）：希望の種類に寄せる（月の回数は増やさない）
    objective_terms += doc_pref_terms(model, doctors, days, daily_active_shifts, shifts, staff_df, is_holiday, required_count,
                                      lambda doc: min_shifts_total[doc], total_cap)
    # 年間の回数（1か月あたり）をそろえる（年間の当直回数を読み込み、反映するにしたときだけ）
    objective_terms += add_annual_balance_terms(model, doctors, days, daily_active_shifts, is_holiday, required_count,
                                                worked_all, holiday_worked, shifts, annual_balance)
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
        return None, False, ["当直案の計算を始められませんでした。医師条件の数字（回数・間隔など）に、大きすぎる値やおかしな値がないか確認してください。"], None, None
    if status == cp_model.UNKNOWN:
        reasons.append("すべての枠を埋める案は、時間内に見つかりませんでした（条件の上で、ぜったいに無理と決まったわけではありません）。")
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
            reasons.append("空き枠を残しても、条件どうしがぶつかって当直案を作れませんでした。確定当直と回数の上限を確認してください。")
        elif relax_status == cp_model.MODEL_INVALID:
            reasons.append("当直案の計算を始められませんでした。医師条件の数字（回数・間隔など）に、大きすぎる値やおかしな値がないか確認してください。")
        else:
            reasons.append("空き枠を残す案も、時間内に見つかりませんでした（条件の上で、ぜったいに無理と決まったわけではありません）。")
    except Exception as e:
        reasons.append(f"⚠️ 不足を含む当直案の作成中にもエラーが発生しました。詳細: {e}")

    return None, False, reasons, None, None


def generate_shift(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None, locked_days=(), annual_balance=None):
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
    result, success, warnings, past, future = _generate_shift_core(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority, locked_days, annual_balance)
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
        st.warning(f"不足が残っています：未決定の枠が{sum(shortage_rows)}枠あります（医師別の当直一覧のいちばん下、赤線の下で確認できます）")
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
    "saved_state": ["項目", "値"],   # 以前の途中保存（1回分だけ）。今は読み込みと、履歴への引き継ぎにだけ使う
    "saved_history": ["保存ID", "保存日時", "年月", "番号", "内容"],   # 作業の途中保存の履歴（新しいものから一定回数分）
    # 期間指定版の途中保存（期間指定版のページでも同じものを登録する。閲覧用の画面は通常版の中で開くため、ここにも書いておく）
    "saved_state_period": ["項目", "値"],
    "saved_history_period": ["保存ID", "保存日時", "年月", "番号", "内容"],
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


WORK_HISTORY_KEEP = 20     # 途中保存を残す回数（これより古いものは、保存のときに消す）
# 保存先のシート。期間指定版では、通常版とは別のシートに切り替える
WORK_STATE_SHEETS = {"legacy": "saved_state", "history": "saved_history"}


def _work_state_legacy(sheet):
    """以前の形（1回分だけ）の途中保存 (保存日時, 内容) を読む。なければ (None, None)。"""
    try:
        rows = {r.get("項目", ""): r.get("値", "") for r in remote_read(sheet)}
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


def _work_history_rows(sheets):
    try:
        return [r for r in remote_read(sheets["history"]) if r.get("保存ID")]
    except Exception:
        return []


def work_state_entries(sheets=None):
    """途中保存の一覧 [(保存ID, 保存日時, 内容)] を新しい順に返す。"""
    sheets = sheets or WORK_STATE_SHEETS
    groups = {}
    for r in _work_history_rows(sheets):
        g = groups.setdefault(r["保存ID"], {"at": r.get("保存日時", ""), "parts": {}})
        try:
            g["parts"][int(r.get("番号") or 1)] = r.get("内容", "")
        except ValueError:
            continue
    entries = []
    for sid, g in groups.items():
        try:
            data = json.loads("".join(g["parts"][k] for k in sorted(g["parts"])))
        except ValueError:
            continue
        entries.append((sid, g["at"], data))
    if not entries:
        # 履歴がまだないときは、以前の形の途中保存を1件として見せる
        at, data = _work_state_legacy(sheets["legacy"])
        if data:
            entries.append(("legacy", at, data))
    entries.sort(key=lambda e: str(e[1]), reverse=True)   # 保存日時は「2026/10/30 15:25:00」の形なので文字の順で並ぶ
    return entries


def work_state_title(data):
    """途中保存の中身の名前（「2026年11月」、期間指定版なら「2026年12/1〜2027年1/5」）。"""
    y, m = int(data.get("year", 0)), int(data.get("month", 0))
    if "extra" in data:
        return TouchokuPeriod(y, m, int(data.get("extra") or 0)).title()
    return f"{y}年{m}月"


def work_state_save(snapshot, sheets=None):
    """途中保存を履歴に1件足す（古いものは WORK_HISTORY_KEEP 回分を残して消す）。"""
    sheets = sheets or WORK_STATE_SHEETS
    text = json.dumps(snapshot, ensure_ascii=False)
    if len(text) > WORK_STATE_CHUNK * 20:
        raise ValueError("保存する内容が大きすぎます（医師の人数や確定当直の行数を確認してください）。")
    old_rows = _work_history_rows(sheets)
    rows = [[r.get("保存ID", ""), r.get("保存日時", ""), r.get("年月", ""), r.get("番号", ""), r.get("内容", "")] for r in old_rows]
    if not rows:
        # 以前の形の途中保存があれば、履歴の1件目として残す
        at, data = _work_state_legacy(sheets["legacy"])
        if data:
            old = json.dumps(data, ensure_ascii=False)
            for k, i in enumerate(range(0, len(old), WORK_STATE_CHUNK), 1):
                rows.append(["legacy", at, work_state_title(data), k, old[i:i + WORK_STATE_CHUNK]])
    now = remote_now()
    sid = hashlib.sha256(f"{now}{random.random()}".encode()).hexdigest()[:12]
    # 長いときは複数の行（番号1, 2, …）に分けて保存する
    for k, i in enumerate(range(0, max(len(text), 1), WORK_STATE_CHUNK), 1):
        rows.append([sid, now, work_state_title(snapshot), k, text[i:i + WORK_STATE_CHUNK]])
    # 新しいものから WORK_HISTORY_KEEP 回分だけ残す
    order = sorted({(str(r[1]), str(r[0])) for r in rows}, reverse=True)
    keep = {sid_ for _, sid_ in order[:WORK_HISTORY_KEEP]}
    remote_rewrite(sheets["history"], [r for r in rows if str(r[0]) in keep])


def work_state_info(sheets=None):
    """いちばん新しい途中保存 (保存日時, 内容) を返す。なければ (None, None)。"""
    entries = work_state_entries(sheets)
    return (entries[0][1], entries[0][2]) if entries else (None, None)


def work_state_for_carry(year, month, sheets=None):
    """医師条件の引き継ぎに使う途中保存：前月の分があればその最新、なければいちばん新しいもの。"""
    entries = [e for e in work_state_entries(sheets) if e[2].get("staff_csv")]
    if not entries:
        return None, None
    prev = (year, month - 1) if month > 1 else (year - 1, 12)
    for _, at, data in entries:
        if (int(data.get("year", 0)), int(data.get("month", 0))) == prev:
            return at, data
    return entries[0][1], entries[0][2]


def work_state_restore():
    """保存した作業を画面に戻す（ボタンのコールバック。画面を作る前に実行される）。一覧で選んだ回を読み込む。"""
    entries = work_state_entries()
    pick = st.session_state.get("work_restore_pick")
    chosen = next((e for e in entries if e[0] == pick), entries[0] if entries else None)
    saved_at, data = (chosen[1], chosen[2]) if chosen else (None, None)
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
    if data.get("paper_notes") is not None:   # 紙で配る当直不可日の用紙の注意書き
        st.session_state[PAPER_NOTES_KEY] = str(data["paper_notes"])
        st.session_state.pop("paper_form_notes_edit", None)
    st.session_state["restored_state"] = data
    st.session_state.pop("fixed_src", None)   # 確定当直の表も、読み込んだ内容で作り直す
    st.session_state["multi_editor_gen"] = st.session_state.get("multi_editor_gen", 0) + 1   # 増員設定の表を読み込んだ内容で作り直す
    # ファイルのアップロード欄を空にして、復元した内容を使う
    st.session_state["uploader_gen"] = st.session_state.get("uploader_gen", 0) + 1
    st.session_state.pop("last_uploaded_file_id", None)
    st.session_state["work_state_message"] = ("success", f"{saved_at} に保存した作業（{y}年{m}月）を読み込みました。")
    st.session_state["staff_source_label"] = f"{str(saved_at)[5:16]} の途中保存から"
    st.session_state.pop("staff_added_names", None)
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
    saved_at, data = work_state_for_carry(year, month)
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
    if data.get("group_rules"):
        state["group_rules"] = data["group_rules"]   # グループの設定も引き継ぐ
    if data.get("paper_notes") is not None:   # 紙で配る用紙の注意書きも引き継ぐ
        st.session_state[PAPER_NOTES_KEY] = str(data["paper_notes"])
        st.session_state.pop("paper_form_notes_edit", None)
    st.session_state["restored_state"] = state
    st.session_state["uploader_gen"] = st.session_state.get("uploader_gen", 0) + 1
    st.session_state.pop("last_uploaded_file_id", None)
    st.session_state["carry_over_gen"] = st.session_state.get("carry_over_gen", 0) + 1   # 小さな枠を閉じる
    st.session_state["staff_source_label"] = f"{data.get('year')}年{data.get('month')}月の医師条件（引き継ぎ）から"
    st.session_state.pop("staff_added_names", None)
    st.session_state["carry_over_message"] = (
        "success", f"{saved_at} に途中保存した{data.get('year')}年{data.get('month')}月の医師条件（{len(df)}名）を引き継ぎました。"
                   "NG日・希望日・先生のコメントは空にしてあります。")


# ==========================================
# 先生の追加・外す（①先生の名簿）
# 今の医師条件の表（前回の画面の内容）に行を足したり消したりして、表を作り直す。
# ==========================================
STAFF_KEEP_EMPTY_ON_COPY = ("備考（メモ・説明など自由記入）", NG_COLUMN, REQUEST_COLUMN, *DOC_PREF_COLUMNS)   # 多め希望も先生ごとのものなので写さない


def _staff_apply_new_table(df, year, month, message):
    """作り直した医師条件の表を、次の画面の表として使う（ほかの入力はそのまま）。"""
    state = dict(st.session_state.get("restored_state") or {"year": year, "month": month})
    state["staff_csv"] = df.to_csv(index=False)
    st.session_state["restored_state"] = state
    # 医師条件のCSVのアップロード欄だけを空にする（確定当直のアップロード欄はそのまま）
    st.session_state["staff_uploader_gen"] = st.session_state.get("staff_uploader_gen", 0) + 1
    st.session_state.pop("last_uploaded_file_id", None)
    st.session_state["staff_edit_gen"] = st.session_state.get("staff_edit_gen", 0) + 1   # 開いている小さな枠を閉じる
    st.session_state["staff_edit_message"] = message


def _pm_duty_changed(doc_name, key, year, month):
    """②の「翌日PM duty」の曜日ボタンを押したとき（コールバック）。医師条件の表のその先生の「翌日PM duty」を書き換える。"""
    csv_text = st.session_state.get("staff_current_csv")
    if not csv_text:
        return
    df = parse_staff_csv(csv_text.encode("utf-8"))
    hit = df["先生の名前"].map(clean_text) == doc_name
    if not hit.any():
        return
    picked = st.session_state.get(key) or []
    if "翌日PM duty" not in df.columns:
        df["翌日PM duty"] = ""
    df["翌日PM duty"] = df["翌日PM duty"].astype(object)
    df.loc[hit, "翌日PM duty"] = ",".join(w for w in WEEKDAYS_JA if w in picked)
    _staff_apply_new_table(df, year, month, None)


def pm_duty_setting_ui(doc_name, value, locked, year, month):
    """
    ②（NG日を入れる）の先生ごとの枠に出す「翌日PM duty」の曜日ボタン。押すとすぐ医師条件の表（③）にも反映する。
    locked（カレンダーに反映していないNGがある）ときは押せない（押すとカレンダーが描き直され、反映前のNGが消えるため）。
    """
    key = f"pm_duty_pick::{doc_name}"
    st.session_state[key] = [WEEKDAYS_JA[i] for i in pm_duty_weekdays(value)]   # いつも医師条件の表の今の値に合わせる
    with st.container(border=True):
        st.markdown("<div style='font-weight:800;color:#8a4b08'>🕑 翌日PM duty"
                    "<span style='font-weight:500;color:#6a5a40;font-size:.85rem;margin-left:6px'>（その曜日の宿直を原則外す。例：木曜PMにdutyがあれば「水」）</span></div>",
                    unsafe_allow_html=True)
        kwargs = dict(key=key, on_change=_pm_duty_changed, args=(doc_name, key, year, month), disabled=locked, label_visibility="collapsed")
        try:
            st.pills("翌日PM duty", WEEKDAYS_JA, selection_mode="multi", **kwargs)
        except (AttributeError, TypeError):   # 古いStreamlitで pills が使えない場合
            st.multiselect("翌日PM duty", WEEKDAYS_JA, **kwargs)
        if locked:
            st.caption("カレンダーにまだ反映していないNGがあるので、押せません。先に「NG日を反映する」を押してください。")
        else:
            st.caption("押すとすぐ、③ 医師条件の表の「翌日PM duty」にも反映され、下のカレンダーの⚠も変わります（複数選べます。もう一度押すと外れます）。")


def staff_add_doctor(year, month):
    """「この先生を追加する」（ボタンのコールバック）。"""
    csv_text = st.session_state.get("staff_current_csv")
    name = clean_text(st.session_state.get("add_doc_name", ""))
    if not csv_text:
        st.session_state["staff_edit_message"] = ("error", "医師条件の表が読み込めませんでした。画面を開き直してください。")
        return
    df = parse_staff_csv(csv_text.encode("utf-8"))
    names = [clean_text(n) for n in df["先生の名前"].tolist()]
    if not name:
        st.session_state["staff_edit_message"] = ("error", "追加する先生の名前を入れてください。")
        return
    if name in names:
        st.session_state["staff_edit_message"] = ("error", f"「{name}」先生は、すでに名簿にいます。")
        return
    if any(c in name for c in ",、") or name == "-" or "⚠️不足" in name:
        st.session_state["staff_edit_message"] = ("error", "先生の名前に、カンマ・読点・単独の「-」は使えません。")
        return
    if st.session_state.get("add_doc_source") == "型から選ぶ":
        src_name = st.session_state.get("add_doc_template")
        tmpl = (st.session_state.get("staff_templates") or {}).get(src_name)
        if not tmpl:
            st.session_state["staff_edit_message"] = ("error", "型を選んでください。")
            return
        row = {c: tmpl.get(c, "") for c in df.columns}
        src_text = f"{src_name.split('_')[0]}の条件"
    else:
        src_name = st.session_state.get("add_doc_copy")
        if src_name not in names:
            st.session_state["staff_edit_message"] = ("error", "条件を写す先生を選んでください。")
            return
        row = df.iloc[names.index(src_name)].to_dict()
        src_text = f"{src_name}先生と同じ条件"
    for c in STAFF_KEEP_EMPTY_ON_COPY:
        if c in row:
            row[c] = ""
    if "希望優先度(数字が大きいほど優先)" in row:
        row["希望優先度(数字が大きいほど優先)"] = 1
    row["先生の名前"] = name
    if "翌日PM duty" in df.columns:
        row["翌日PM duty"] = clean_text(st.session_state.get("add_doc_pm", ""))
    after = st.session_state.get("add_doc_after")
    pos = 0 if after == "（いちばん上）" else (names.index(after) + 1 if after in names else len(df))
    new_df = pd.concat([df.iloc[:pos], pd.DataFrame([row], columns=df.columns), df.iloc[pos:]], ignore_index=True)
    for k in ("add_doc_name", "add_doc_pm"):
        st.session_state[k] = ""
    where = "いちばん上" if pos == 0 else f"{new_df.iloc[pos - 1]['先生の名前']}先生の下"
    st.session_state["staff_added_names"] = list(st.session_state.get("staff_added_names") or []) + [name]
    _staff_apply_new_table(new_df, year, month, ("success", f"{name}先生を、{where}に追加しました（{src_text}）。②のNG日と③の表に表示されます。"))


def staff_remove_doctors(year, month):
    """「選んだ先生を外す」（ボタンのコールバック）。"""
    csv_text = st.session_state.get("staff_current_csv")
    targets = list(st.session_state.get("remove_doc_names") or [])
    if not csv_text or not targets:
        st.session_state["staff_edit_message"] = ("error", "外す先生を選んでください。")
        return
    df = parse_staff_csv(csv_text.encode("utf-8"))
    new_df = df[~df["先生の名前"].astype(str).str.strip().isin(targets)].reset_index(drop=True)
    st.session_state["remove_doc_names"] = []
    st.session_state["staff_added_names"] = [n for n in (st.session_state.get("staff_added_names") or []) if n not in targets]
    _staff_apply_new_table(new_df, year, month, ("success", f"{'、'.join(targets)}先生を名簿から外しました。"
                                                            "間違えた場合は、途中保存の履歴から読み込み直せます。"))


def _staff_summary(row):
    """「A宿直・A日直・外来日直／2〜3回」のような、条件の短い説明。"""
    def num(c):
        v = pd.to_numeric(pd.Series([row.get(c)]), errors="coerce").iloc[0]
        return None if pd.isna(v) else int(v)
    slots = [s for s in RESULT_COLUMNS[2:] if (num(s + "上限") or 0) > 0]
    lo, hi = num("月下限"), num("月上限")
    return ("・".join(slots) or "担当枠なし") + (f"／{lo}〜{hi}回" if lo is not None and hi is not None else "")


def staff_roster_html(names, source, added=()):
    """①に出す「今の名簿」：番号付きの小さな札を、表と同じ順番で並べる。追加した先生には「追加」の印。"""
    chips = "".join(
        f"<span class='rs-chip{' new' if n in added else ''}'><b>{i}</b>{html_escape(n)}{'<em>追加</em>' if n in added else ''}</span>"
        for i, n in enumerate(names, 1))
    return ("<style>.rs-box{border:1px solid #d6e0ec;background:#f7f9fc;border-radius:10px;padding:10px 12px;margin:10px 0 4px}"
            ".rs-hd{font-weight:800;color:#24405f;margin-bottom:6px}.rs-hd small{font-weight:500;color:#6a7686;margin-left:8px}"
            ".rs-chip{display:inline-flex;align-items:center;gap:4px;border:1px solid #d6dde8;background:#fff;border-radius:999px;"
            "padding:2px 9px 2px 3px;margin:3px;font-size:.85rem;color:#31333f}"
            ".rs-chip b{display:inline-flex;align-items:center;justify-content:center;min-width:1.3rem;height:1.3rem;border-radius:999px;"
            "background:#e6ecf4;color:#425268;font-size:.7rem;padding:0 3px}"
            ".rs-chip.new{border-color:#2e8a5b;background:#effaf3}"
            ".rs-chip em{font-style:normal;font-size:.68rem;background:#2e8a5b;color:#fff;border-radius:4px;padding:0 4px}</style>"
            f"<div class='rs-box'><div class='rs-hd'>👥 今の名簿（{len(names)}名・この順番で表や当直案に並びます）"
            f"<small>{html_escape(source)}</small></div>{chips}</div>")


# 医師条件の表の列名（短い名前）・幅・説明。①（見るだけ）と③（編集用）で同じものを使う
STAFF_SLOT_SHORT = {"A宿直": "A宿", "B宿直": "B宿", "外来宿直": "外宿", "A日直": "A日", "B日直": "B日", "外来日直": "外日"}
STAFF_ROW_HEIGHT = 26


def staff_column_specs():
    """{列名: (表に出す名前, 幅px, 種類, 説明)}"""
    cap = "を担当する月間の上限です。空欄なら、できるだけ2回まで（ほかの先生で埋まらないときだけ3回以上も担当します）。確定当直がある場合は、超えることがあります。"
    specs = {
        "先生の名前": ("名前", 72, "text", "先生の名前です。ほかの先生と重ならない表記にしてください。"),
        "備考（メモ・説明など自由記入）": ("備考", 110, "text",
            "管理用メモです。ここに書いた内容は計算に反映されません。セルをダブルクリックすると全文が表示され、編集できます。表の下の「備考の全文を見る」でも確認できます。"),
        DOC_GROUP_COLUMN: ("グループ", 96, "group",
            "グループを選ぶと、上の「👥 グループの設定」で決めた枠の決まり（◯担当する・△できるだけ担当しない・×担当しない）が使われます。空欄の先生には、グループの決まりを使いません（医師条件の表の値だけで決めます）。"),
        NG_COLUMN: ("NG日", 150, "text",
            "当直に入れない日です（例：3,10:宿NG）。日にちだけなら日直・宿直ともNG、「:宿NG」は宿直だけNG、「:日NG」は日直だけNGです。"),
        "翌日PM duty": ("翌日PM duty", 92, "text",
            "例：木曜PMにdutyがある場合は水。複数は水,木のように入力。翌日が休日なら宿直に入る場合があります。日直は対象外です。確実に外す日はカレンダーでNGを指定してください。"),
        "病棟/外来 多め": ("病棟/外来 多め", 120, "pref",
            "病棟多め：病棟（A・B）の日直・宿直を、できれば多めに入れます。外来多め：外来日直・外来宿直を、できれば多めに入れます。"
            "月の回数は増えません（回数の中で、種類の割合を寄せるだけ）。空欄なら希望なし。"),
        "宿直/日直 多め": ("宿直/日直 多め", 120, "pref",
            "宿直多め：宿直（A・B・外来）を、できれば多めに入れます。日直多め：日直（A・B・外来）を、できれば多めに入れます。"
            "月の回数は増えません（回数の中で、種類の割合を寄せるだけ）。空欄なら希望なし。"),
        "平日/休日 多め": ("平日/休日 多め", 120, "pref",
            "平日多め：平日の宿直を、できれば多めに入れます。休日多め：休日（土日祝と「休日にする」にした日）の日直・宿直を、できれば多めに入れます。"
            "月の回数は増えません（回数の中で、種類の割合を寄せるだけ）。休日回数の偏りを減らす決まりの方が優先です。空欄なら希望なし。"),
        STAFF_COMMENT_COLUMN: ("💬 先生からのコメント", 200, "comment",
            "先生が専用URLから入れた「管理者へのコメント」です。取り込むと自動で入ります（この表では直せません）。"
            "長いコメントは途中まで表示されます。マスをクリックすると全文が出ます。"),
        STAFF_ANNUAL_COLUMN: ("📊 年間（1か月あたり）", 190, "comment",
            "「📊 年間の当直回数」で読み込んだ集計の、1か月あたりの回数です（計＝当直合計、外＝外来、日＝日直、休＝休日）。"
            "病棟＝計−外来、宿直＝計−日直、平日＝計−休日で分かります。見るだけで、この表では直せません。"),
        REQUEST_COLUMN: ("希望日", 100, "text",
            "例：10,15 または 10:A宿直。複数は半角カンマ区切り。日直は休日（または「休日にする」にした日）だけ指定できます。通常の希望は各種条件の範囲内で割り当てます。"),
        "希望優先度(数字が大きいほど優先)": ("優先度", 56, "num",
            "希望優先度。通常は1。1〜99は数字が大きいほど優先。100以上はこの先生の希望日すべてが確定当直と同じ扱いになり、NG日・上限より優先されます。ほかの勤務との間隔は空けます。"),
        "宿⇔宿 間隔下限": ("宿⇔宿", 76, "days",
            "【宿⇔宿 間隔下限】宿直と次の宿直の間に空ける日数の下限。5日なら10日の宿直の次の宿直は16日以降。日直がからむ組み合わせは右の「宿⇔日・日⇔日 間隔下限」を使います。確定当直とほかの勤務の間も空けます。確定当直同士が近すぎる場合は注意を表示します。"),
        DAY_GAP_COLUMN: ("宿⇔日・日⇔日", 118, "days",
            f"【{DAY_GAP_COLUMN}】日直がからむ組み合わせ（日直→宿直・宿直→日直・日直→日直）で、勤務の間に空ける日数の下限。2日なら10日の次は13日以降。"
            "宿直→宿直は「宿⇔宿 間隔下限」を使います。空欄なら宿⇔宿 間隔下限と同じ日数にします。"),
        "月下限": ("月下限", 60, "count", "月下限（目標）：できるだけ確保したい回数です。条件によっては届きません。月上限以下にしてください。"),
        "月上限": ("月上限", 60, "count", "月上限：日直・宿直を合わせた月間の上限です。確定当直がある場合は、超えることがあります。"),
        "休日上限": ("休日上限", 66, "count", "休日上限：土日祝と「休日にする」にした日の、日直と宿直の合計の上限です。1枠を1回と数えます。"),
        "宿直上限": ("宿直上限", 66, "count0", "A宿直・B宿直・外来宿直を合わせた月間の上限です。空欄なら制限なし。確定当直がある場合は、超えることがあります。"),
        "日直上限": ("日直上限", 66, "count0", "A日直・B日直・外来日直を合わせた月間の上限です。空欄なら制限なし。確定当直がある場合は、超えることがあります。"),
    }
    for slot, short in STAFF_SLOT_SHORT.items():
        specs[slot + "上限"] = (short + "上限", 62 if len(short) == 2 and short[0] in "AB" else 66, "count", f"{short}上限＝{slot}" + cap)
    return specs


def header_fit_width(label, width):
    """列の見出しが全部見える幅（px）。見出しの左には型の印（≡）が付くので、その分も足す。"""
    need = sum(14 if ord(ch) > 255 else 8 for ch in str(label)) + 46
    return max(int(width), need)


def _column(maker, label, width, help_text, pinned=False, **kw):
    """列の設定を作る。名前の列は左に固定する。数字の列も含めて、すべて左揃えにする（どの列の値か分かりやすいように）。
    固定・左揃えの機能がない古いStreamlitでは、その機能を使わずに作る。列の幅は、見出しが全部見える幅にする。"""
    width = header_fit_width(label, width)
    tries = ([dict(pinned=True, alignment="left"), dict(pinned=True)] if pinned else []) + [dict(alignment="left")]
    for extra in tries:
        try:
            return maker(label, width=width, help=help_text, **extra, **kw)
        except TypeError:
            continue
    return maker(label, width=width, help=help_text, **kw)


def staff_column_config(columns, as_index=False, show_ng=False, ng_width=None):
    """st.data_editor / st.dataframe に渡す column_config。as_index=True なら「先生の名前」は表の見出し列（index）。
    NG日の列はふつう出さない（NG日はカレンダーで入れるため）。show_ng=True（閲覧用の画面）のときだけ出す。
    ng_width を渡すと（閲覧用の画面）、NG日の列をその幅にし、ほかの列は列名が見える最小限の幅にする（名前の列はそのまま）。"""
    specs = staff_column_specs()
    config = {}
    for c in columns:
        if c == NG_COLUMN and not show_ng:
            config[c] = None
            continue
        label, width, kind, help_text = specs.get(c, (str(c), 80, "text", None))
        if ng_width is not None and c != "先生の名前":
            width = ng_width if c == NG_COLUMN else 0   # 0 → 列名が見える最小限の幅（_column で列名に合わせる）
        if kind == "text":
            config[c] = _column(st.column_config.TextColumn, label, width, help_text, pinned=(c == "先生の名前"))
        elif kind == "group":
            config[c] = _column(st.column_config.SelectboxColumn, label, width, help_text, options=doc_group_names())
        elif kind == "comment":
            config[c] = _column(st.column_config.TextColumn, label, width, help_text, disabled=True)
        elif kind == "pref":
            config[c] = _column(st.column_config.SelectboxColumn, label, width, help_text, options=list(DOC_PREF_COLUMNS[c]))
        elif kind == "num0":
            config[c] = _column(st.column_config.NumberColumn, label, width, help_text, min_value=0, step=1)
        elif kind == "days":
            # 日数の列は「5日間」のように単位を付けて見せる（中身は数字のまま。入力は数字だけでよい）
            config[c] = _column(st.column_config.NumberColumn, label, width, help_text, min_value=0, step=1, format="%d日間")
        elif kind in ("count", "count0"):
            # 回数の列は「3回」のように見せる（count0 は空欄＝制限なし を使える列）
            config[c] = _column(st.column_config.NumberColumn, label, width, help_text, min_value=0, step=1, format="%d回")
        else:
            config[c] = _column(st.column_config.NumberColumn, label, width, help_text)
    if as_index:
        label, width, _, help_text = specs["先生の名前"]
        config["_index"] = _column(st.column_config.TextColumn, label, width, help_text, pinned=True)
    return config


def staff_legend_text():
    return ("列名の意味：優先度＝希望優先度　宿⇔宿＝宿直と宿直の間に空ける日数（宿⇔宿 間隔下限）　宿⇔日・日⇔日＝日直がからむ勤務（日直→宿直・宿直→日直・日直→日直）の間に空ける日数（宿⇔日・日⇔日 間隔下限）　"
            "月下限＝月の回数の目標（できるだけ確保）　"
            "月上限＝月の回数の上限　休日上限＝休日の回数の上限　宿直上限・日直上限＝A・B・外来を合わせた上限　"
            "A宿上限〜外日上限＝各枠（A宿直・B宿直・外来宿直・A日直・B日直・外来日直）の月の上限（空欄はできるだけ2回まで）　"
            "病棟/外来 多め・宿直/日直 多め・平日/休日 多め＝その種類の当直を、できれば多めに入れる（月の回数は増えない）")


# 表のすぐ上に出す、青い案内の帯（列名にカーソルを合わせると説明が出ることを知らせる）
STAFF_COLUMN_HINT_HTML = (
    "<style>.col-hint{display:flex;align-items:center;flex-wrap:wrap;gap:4px 10px;background:#eaf3ff;border:1.5px solid #6ea2e6;"
    "border-radius:10px;padding:9px 14px;margin:2px 0 8px;font-size:1rem;color:#1d4e89;font-weight:700;line-height:1.5}"
    ".col-hint .ic{font-size:1.25rem}.col-hint small{font-weight:500;color:#41618a;font-size:.85rem}</style>"
    "<div class='col-hint'><span class='ic'>💡</span>列名にカーソルを合わせると、その列の説明が出ます"
    "<small>（例：「宿⇔日・日⇔日」に合わせると、どの組み合わせに使うかが出ます）</small></div>")


# 医師条件の表（①・③）に、見るだけで出す「先生からのコメント」の列（専用URLから取り込んだもの。医師条件のデータには入れない）
STAFF_COMMENT_COLUMN = "💬 先生からのコメント"


STAFF_ANNUAL_COLUMN = "📊 年間（1か月あたり）"


def with_annual_column(df, rates, as_index):
    """表に出すために、コメント（なければ備考）の右へ「📊 年間（1か月あたり）」の列を足したコピーを返す。rates は {名前: 文字}。"""
    df = df.copy()
    names = list(df.index) if as_index else list(df.get("先生の名前", pd.Series(dtype=str)))
    values = [rates.get(clean_text(n), "") for n in names]
    cols = list(df.columns)
    anchor = next((i for i, c in enumerate(cols) if c == STAFF_COMMENT_COLUMN), None)
    if anchor is None:
        anchor = next((i for i, c in enumerate(cols) if str(c).startswith("備考")), None)
    pos = anchor + 1 if anchor is not None else (0 if as_index else min(1, len(cols)))
    df.insert(pos, STAFF_ANNUAL_COLUMN, values)
    return df


def with_comment_column(df, comments, as_index):
    """表に出すために、備考の右へ「先生からのコメント」の列を足したコピーを返す。comments は {名前: (コメント, 保存日時)}。"""
    df = df.copy()
    names = list(df.index) if as_index else list(df.get("先生の名前", pd.Series(dtype=str)))
    values = [clean_text((comments.get(clean_text(n)) or ("",))[0]) for n in names]
    cols = list(df.columns)
    note = next((i for i, c in enumerate(cols) if str(c).startswith("備考")), None)
    pos = note + 1 if note is not None else (0 if as_index else min(1, len(cols)))
    df.insert(pos, STAFF_COMMENT_COLUMN, values)
    return df


# 表の上に出す「この表の列」の地図：列の仲間（名前, 札の背景色, 文字色）
STAFF_MAP_GROUPS = [("基本", "#e9eef6", "#425268"), ("希望", "#fff1dc", "#8a4b08"), ("間隔・回数", "#e6f4ea", "#2e6b3f"), ("枠ごとの上限", "#efe8fb", "#4f3a8c")]


def staff_column_group(column):
    """医師条件の列が、地図のどの仲間か（STAFF_MAP_GROUPS の番号）。"""
    if column in DOC_PREF_COLUMNS or column in (REQUEST_COLUMN, "希望優先度(数字が大きいほど優先)"):
        return 1
    if column in ("宿⇔宿 間隔下限", DAY_GAP_COLUMN, "月下限", "月上限", "休日上限", *GROUP_CAP_COLUMNS):
        return 2
    if column in [s + "上限" for s in STAFF_SLOT_SHORT]:
        return 3
    return 0


SLOT_CAP_LABEL = "枠ごとの上限（A宿・B宿・外宿・A日・B日・外日）"
SLOT_CAP_SHOW_KEY = "show_slot_caps"   # 「枠ごとの上限の列も出す」のスイッチ


def slot_cap_columns(columns):
    """医師条件の列のうち、枠ごとの上限（A宿上限〜外日上限）の列。"""
    names = [s + "上限" for s in STAFF_SLOT_SHORT]
    return [c for c in columns if c in names]


def slot_cap_values(df):
    """枠ごとの上限に値が入っている先生と値。[(名前, 列の表示名, 値)]（空欄は「できれば2回まで」で、値ではないので数えない）。"""
    specs = staff_column_specs()
    names = list(df.index) if "先生の名前" not in df.columns else list(df["先生の名前"])
    found = []
    for c in slot_cap_columns(df.columns):
        for n, v in zip(names, df[c].tolist()):
            t = clean_text(v)
            if t and t.lower() not in ("nan", "none", "<na>"):
                try:
                    t = f"{int(float(t))}回"
                except ValueError:
                    pass
                found.append((clean_text(n), specs.get(c, (c,))[0], t))
    return found


def staff_column_map_html(columns, hidden_slot_caps=False):
    """表のすぐ上に出す「この表の列」の地図。全部の列名を、仲間ごとに色分けした札で、表と同じ順（左から右）に並べる。"""
    specs = staff_column_specs()
    parts, last = [], None
    for c in columns:
        g = staff_column_group(c)
        name, bg, fg = STAFF_MAP_GROUPS[g]
        if g != last:
            if last is not None:
                parts.append("<span class='cm-ar'>→</span>")
            parts.append(f"<span class='cm-gl' style='color:{fg}'>{name}</span>")
            last = g
        parts.append(f"<span class='cm-c' style='background:{bg};color:{fg}'>{html_escape(specs.get(c, (str(c),))[0])}</span>")
    return ("<style>.cm-map{display:flex;flex-wrap:wrap;align-items:center;gap:4px;margin:0 0 8px;font-size:.8rem;line-height:1.5}"
            ".cm-map .cm-t{font-weight:800;color:#24405f;margin-right:4px}.cm-map .cm-gl{font-weight:800;font-size:.75rem;margin:0 1px 0 6px}"
            ".cm-map .cm-c{border-radius:999px;padding:0 8px;border:1px solid rgba(0,0,0,.08);white-space:nowrap}"
            ".cm-map .cm-ar,.cm-map .cm-n{color:#8a94a3}</style>"
            f"<div class='cm-map'><span class='cm-t'>🗺 この表の列</span>{''.join(parts)}"
            + ("<span class='cm-ar'>→</span><span style='border:1.5px dashed #7a5cc4;color:#4f3a8c;background:#f6f2fd;border-radius:999px;"
               f"padding:0 9px;font-weight:800;white-space:nowrap'>＋ {SLOT_CAP_LABEL}：通常は不要。各枠の上限を個別に決めたいときに表示</span>"
               if hidden_slot_caps else "")
            + "<span class='cm-n'>（左から右の順。右へスクロールすると続きが見えます）</span></div>")


def _slot_cap_switch_changed(base):
    """「枠ごとの上限の列も出す」を切り替えたとき。表が作り直されるので、③で直しかけの内容を先に今の表として残す。"""
    csv_text = st.session_state.get("staff_current_csv")
    if not csv_text or csv_text == base.reset_index().to_csv(index=False):
        return
    year, month = st.session_state.get("staff_table_ym") or (None, None)
    _staff_apply_new_table(parse_staff_csv(csv_text.encode("utf-8")), year, month, None)


def keep_blank_slot_caps(df):
    """
    途中保存・Excelのダウンロード用。枠ごとの上限（A宿上限〜外日上限）は、③の表で空欄だった先生は空欄のまま残す。
    入力チェックで空欄に計算用の初めの値（2）が入るため、そのまま保存すると、読み込んだときに全員に値が入ってしまう。
    計算用の列（空欄だった枠の印）は保存しない。
    """
    caps = slot_cap_columns(df.columns)
    if SLOT_CAP_SOFT_COLUMN not in df.columns or not caps:
        return df.drop(columns=[SLOT_CAP_SOFT_COLUMN], errors="ignore")
    out = df.copy()
    for c in caps:
        out[c] = out[c].astype(object)
    for i, r in out.iterrows():
        for s in soft_cap_slots(r):
            if s + "上限" in out.columns:
                out.at[i, s + "上限"] = None
    return out.drop(columns=[SLOT_CAP_SOFT_COLUMN])


def _slot_cap_clear_all(base):
    """「🧹 枠ごとの上限を全員分空欄にする」を押したとき。③で直しかけの内容は残したまま、枠ごとの上限だけを全員空欄にする。"""
    csv_text = st.session_state.get("staff_current_csv")
    df = parse_staff_csv(csv_text.encode("utf-8")) if csv_text else base.reset_index()
    for c in slot_cap_columns(df.columns):
        df[c] = None
    year, month = st.session_state.get("staff_table_ym") or (None, None)
    _staff_apply_new_table(df, year, month, None)


def slot_cap_switch(base, locked=False):
    """③の表の下の、紫の枠の「枠ごとの上限の列も出す」スイッチ。値が入っているあいだは、列を隠せないので押せない。"""
    with st.container(key="slot_cap_box"):
        st.markdown("<style>.st-key-slot_cap_box{border:2px solid #7a5cc4;background:#f6f2fd;border-radius:10px;padding:8px 14px 4px;margin:6px 0}"
                    ".st-key-slot_cap_box p{color:#4f3a8c}</style>", unsafe_allow_html=True)
        st.toggle(f"🟪 {SLOT_CAP_LABEL}の列も出す", key=SLOT_CAP_SHOW_KEY, disabled=locked,
                  on_change=_slot_cap_switch_changed, args=(base,))
        st.caption("通常はなくても大丈夫なことが多いです（担当する枠はグループの設定で決めます）。"
                   "空欄のときは、1つの枠はできるだけ月2回までにします（ほかの先生で埋まらないときだけ3回以上も担当します）。"
                   "先生ごとに、各枠の上限をより個別に決めたいとき（例：外来宿直は月1回まで）に使ってください。")
        if locked:
            # 値が入っているあいだだけ：人数と、全員分を1回で空欄にするボタン（表の上のほうの行に残った値も消せる）
            n = len({n for n, _, _ in slot_cap_values(base)})
            with st.container(key="slot_cap_clear_box"):
                st.markdown("<style>.st-key-slot_cap_clear_box{background:#fff;border:1.5px solid #e0a36a;border-radius:8px;padding:6px 12px;margin:2px 0 6px}"
                            ".st-key-slot_cap_clear_box p{color:#7a4a0a !important}"
                            ".st-key-slot_cap_clear_btn button{border:1.5px solid #c0392b !important;color:#c0392b !important;background:#fff !important;font-weight:700}"
                            ".st-key-slot_cap_clear_btn button p{color:#c0392b !important;font-weight:700}</style>", unsafe_allow_html=True)
                c1, c2 = st.columns([3, 2], vertical_alignment="center")
                c1.markdown(f"※ 値が入っている先生が **{n}名** いるので、列を出したままにしています。")
                with c2:
                    st.button("🧹 枠ごとの上限を全員分空欄にする", key="slot_cap_clear_btn",
                              on_click=_slot_cap_clear_all, args=(base,))


def staff_table(df, editable, show_ng=False, max_height=300, ng_width=None, comments=None):
    """医師条件の表。③（編集用）と①（見るだけ）で、列名・幅・行の高さ・並びを同じにする。
    show_ng=True なら、見るだけの表にNG日の列も出す（閲覧用の画面）。max_height は見るだけの表の高さの上限。
    ng_width を渡すと、NG日の列をその幅にし、ほかの列は列名が見える最小限の幅にする（閲覧用の画面）。
    comments（{名前: (コメント, 保存日時)}）を渡すと、備考の右に「先生からのコメント」の列を見るだけで出す（返す表には入れない）。"""
    st.markdown(STAFF_COLUMN_HINT_HTML, unsafe_allow_html=True)
    # 枠ごとの上限（A宿〜外日）の列：ふだんは出さない（入れる枠はグループで決める）。スイッチを入れたときと、値が入っているときだけ出す
    cap_values = slot_cap_values(df)
    show_caps = bool(st.session_state.get(SLOT_CAP_SHOW_KEY)) or bool(cap_values)
    hidden_caps = [] if show_caps else slot_cap_columns(df.columns)
    if editable and cap_values:
        shown = "、".join(f"{html_escape(n)}：{html_escape(l)} {html_escape(v)}" for n, l, v in cap_values[:6]) + ("…" if len(cap_values) > 6 else "")
        st.markdown("<div style='background:#fff8e6;border:1.5px solid #f0b46a;border-radius:10px;padding:8px 14px;margin:0 0 8px;color:#7a4a0a'>"
                    f"⚠ 枠ごとの上限に値が入っている先生がいるので、列を出しています（{shown}）。不要なら空欄にしてください。</div>",
                    unsafe_allow_html=True)
    if comments is not None:
        df = with_comment_column(df, comments, as_index=editable)
        # 年間の集計を読み込んでいれば、「📊 年間（1か月あたり）」の列も見るだけで出す（①・③だけ。閲覧用には出さない）
        rates = st.session_state.get("annual_doc_rates") or {}
        st.session_state["annual_doc_rates_used"] = rates
        if rates:
            df = with_annual_column(df, rates, as_index=editable)
    if editable:
        cols = [c for c in df.columns if str(c).startswith("備考")] + [c for c in df.columns if not str(c).startswith("備考")]
        cols = [c for c in cols if c not in hidden_caps]   # 隠した列も、返す表にはそのまま残る
        st.markdown(staff_column_map_html(["先生の名前"] + [c for c in cols if c != NG_COLUMN and not str(c).startswith("_")], bool(hidden_caps)),
                    unsafe_allow_html=True)
        kwargs = dict(num_rows="dynamic", use_container_width=True, height=300, column_order=cols,
                      column_config=staff_column_config(df.columns, as_index=True))
        try:
            try:
                out = st.data_editor(df, row_height=STAFF_ROW_HEIGHT, **kwargs)
            except TypeError:   # 古いStreamlitで row_height が使えない場合
                out = st.data_editor(df, **kwargs)
        except Exception as e:
            # 表を出せないときは、原因の文をそのまま画面に出す（ふつうのエラー表示では中身が隠されるため）
            st.error(f"医師条件の表を表示できませんでした。この画面の写真を管理担当者に送ってください。\n\n詳細：{type(e).__name__}: {e}")
            st.stop()
        if isinstance(out, pd.DataFrame):
            out = out.drop(columns=[STAFF_COMMENT_COLUMN, STAFF_ANNUAL_COLUMN], errors="ignore")   # 見るだけの列は、医師条件のデータには入れない
        # 表で枠ごとの上限を空欄にした（または入れた）ら、その内容を今の表として残して描き直す
        # （列を出す・隠す判断は表を出す前に行うので、描き直さないと、空欄にしても列が出たまま・スイッチも押せないままになる）
        if isinstance(out, pd.DataFrame) and bool(slot_cap_values(out)) != bool(cap_values):
            year, month = st.session_state.get("staff_table_ym") or (None, None)
            _staff_apply_new_table(out.reset_index(), year, month, None)
            st.rerun()
        slot_cap_switch(df.drop(columns=[STAFF_COMMENT_COLUMN, STAFF_ANNUAL_COLUMN], errors="ignore"), locked=bool(cap_values))
    else:
        cols = ["先生の名前"] + [c for c in df.columns if str(c).startswith("備考")] \
            + [c for c in df.columns if c != "先生の名前" and not str(c).startswith("備考") and (show_ng or c != NG_COLUMN)]
        view = df[[c for c in cols if c in df.columns and c not in hidden_caps]].copy()
        st.markdown(staff_column_map_html([c for c in view.columns if not str(c).startswith("_")], bool(hidden_caps)), unsafe_allow_html=True)
        # 「5.0」ではなく「5」と見せる（整数だけの列は整数にする。空欄は空欄のまま）
        for c in view.columns:
            num = pd.to_numeric(view[c], errors="coerce")
            if c != "先生の名前" and num.notna().any() and num.notna().sum() == view[c].replace("", None).notna().sum() \
                    and (num.dropna() % 1 == 0).all():
                view[c] = num.astype("Int64")
        kwargs = dict(hide_index=True, use_container_width=True, column_config=staff_column_config(view.columns, show_ng=show_ng, ng_width=ng_width),
                      height=min(max_height, 38 + (STAFF_ROW_HEIGHT + 1) * max(len(view), 1)))
        try:
            out = st.dataframe(view, row_height=STAFF_ROW_HEIGHT, **kwargs)
        except TypeError:
            out = st.dataframe(view, **kwargs)
    st.caption(staff_legend_text())
    return out


def staff_readonly_table(staff_df, comments=None):
    """①に出す「読み込んだ医師条件」：③の表と同じ列名・幅・行の高さで、見るだけの薄い表にする。"""
    st.markdown("""<style>
.st-key-staff_readonly{position:relative;border:1px solid #e3e7ee;border-radius:10px;background:#f6f7f9;padding:8px 8px 6px}
.st-key-staff_readonly [data-testid="stDataFrame"]{opacity:.55}
.st-key-staff_readonly::after{content:"🔒 見るだけ ｜ 直すときは ③の表で";position:absolute;top:45%;left:50%;transform:translate(-50%,-50%);
  background:#fff;border:2px solid #8a94a6;color:#4a5568;border-radius:999px;padding:6px 16px;font-weight:800;font-size:.95rem;
  box-shadow:0 2px 8px rgba(0,0,0,.08);pointer-events:none;white-space:nowrap;z-index:5}
.ro-head{font-weight:800;color:#24405f;margin:10px 0 4px}.ro-head small{font-weight:500;color:#6a7686;margin-left:6px}
.ro-link{display:inline-block;margin:6px 0 2px;border:1px solid #2e8a5b;color:#1f6b45 !important;background:#fff;border-radius:8px;
  padding:4px 12px;font-weight:700;font-size:.85rem;text-decoration:none !important}
</style>""", unsafe_allow_html=True)
    st.markdown("<div class='ro-head'>📋 読み込んだ医師条件<small>③の表と同じ内容・ここでは直せません</small></div>", unsafe_allow_html=True)
    with st.container(key="staff_readonly"):
        staff_table(staff_df, editable=False, comments=comments)
    st.markdown("<a class='ro-link' href='#sec-cond' target='_self'>⬇ ③で医師条件を直す</a>", unsafe_allow_html=True)


def staff_add_remove_ui(staff_df, year, month, is_template=False):
    """①先生の名簿：「今の名簿」の一覧と、「➕ 先生を追加する」「➖ 先生を外す」。"""
    st.session_state["staff_table_ym"] = (year, month)   # 「枠ごとの上限の列も出す」を切り替えたときに使う
    names = [n for n in (clean_text(v) for v in staff_df["先生の名前"].tolist()) if n and n.lower() not in ("nan", "none")]
    if is_template:
        st.markdown(staff_roster_html(names, "入力例の型1〜型9を表示中"), unsafe_allow_html=True)
        st.warning("まだ先生の名簿が読み込まれていません。"
                   + ("上の「📋 前月の医師条件を引き継ぐ」か、" if remote_is_configured() else "")
                   + "下のファイルのアップロードで読み込んでください。「➕ 先生を追加する」で1人ずつ足すこともできます。")
    else:
        st.markdown(staff_roster_html(names, st.session_state.get("staff_source_label", ""),
                                      set(st.session_state.get("staff_added_names") or [])), unsafe_allow_html=True)
    rows = {clean_text(r["先生の名前"]): r for _, r in staff_df.iterrows()}
    msg = st.session_state.pop("staff_edit_message", None)
    gen = "\u200b" * st.session_state.get("staff_edit_gen", 0)
    has_popover = hasattr(st, "popover")
    c_add, c_rm = st.columns(2)
    with c_add:
        box = st.popover("➕ 先生を追加する" + gen, use_container_width=True) if has_popover else st.expander("➕ 先生を追加する" + gen)
        with box:
            st.text_input("先生の名前", key="add_doc_name", placeholder="例：新田")
            st.selectbox("入れる位置（名簿の並び順）", ["（いちばん上）"] + names, index=len(names),
                         format_func=lambda n: n if n == "（いちばん上）" else f"{n}先生の下", key="add_doc_after")
            st.radio("条件のもと", ["ほかの先生と同じ条件", "型から選ぶ"], horizontal=True, key="add_doc_source")
            if st.session_state.get("add_doc_source") == "型から選ぶ":
                tmpls = st.session_state.get("staff_templates") or {}
                pick = st.selectbox("型", list(tmpls), key="add_doc_template",
                                    format_func=lambda n: f"{n.split('_')[0]}：{tmpls[n].get('備考（メモ・説明など自由記入）', '')}")
                src_row = tmpls.get(pick, {})
            else:
                pick = st.selectbox("条件を写す先生", names, key="add_doc_copy",
                                    format_func=lambda n: f"{n}（{_staff_summary(rows[n])}）" if n in rows else n)
                src_row = rows.get(pick, {})
            if src_row is not None and len(src_row):
                def v(c):
                    x = pd.to_numeric(pd.Series([src_row.get(c)]), errors="coerce").iloc[0]
                    return "なし" if pd.isna(x) else str(int(x))
                caps = "・".join(f"{s}{v(s + '上限')}" for s in RESULT_COLUMNS[2:] if v(s + "上限") not in ("0", "なし"))
                st.caption(f"写す内容：宿⇔宿 間隔下限 {v('宿⇔宿 間隔下限')}日、宿⇔日・日⇔日 間隔下限 {v(DAY_GAP_COLUMN) if v(DAY_GAP_COLUMN) != 'なし' else v('宿⇔宿 間隔下限')}日、"
                           f"月間 {v('月下限')}〜{v('月上限')}回、"
                           f"休日 {v('休日上限')}回まで、宿直上限 {v('宿直上限')}・日直上限 {v('日直上限')}、枠 {caps or 'なし'}。"
                           "備考・希望日・NG日は写しません。")
            st.text_input("翌日PM dutyの曜日（任意）", key="add_doc_pm", placeholder="例：水　（複数は 水,木）")
            st.button("この先生を追加する", key="add_doc_do", type="primary", on_click=staff_add_doctor, args=(year, month))
            if remote_is_configured():
                st.caption("専用URLを使っている場合：新しい先生のURLは、管理画面の「締め切り日・休日設定を更新」（受付中）または「受付を開始する」で作られます。")
    with c_rm:
        box = st.popover("➖ 先生を外す" + gen, use_container_width=True) if has_popover else st.expander("➖ 先生を外す" + gen)
        with box:
            st.multiselect("外す先生（複数選べます）", names, key="remove_doc_names")
            st.caption("外した先生は、医師条件の表・NG日の一覧から消えます。途中保存の履歴には残るので、必要なら読み込み直せます。")
            st.button("選んだ先生を外す", key="remove_doc_do", type="primary", on_click=staff_remove_doctors, args=(year, month))
    if msg:
        (st.success if msg[0] == "success" else st.error)(msg[1])



# ==========================================
# CSVのアップロード欄（ドラッグ＆ドロップできることが分かる見た目）。アプリ内のCSV欄はすべてこれを使う
# ==========================================
CSV_UPLOAD_CSS = """<style>
/* 見た目の指定だけの欄（中身が<style>だけ）は場所を取らないようにする（設定欄の上に大きな空きができないように） */
[data-testid="stElementContainer"]:has(> [data-testid="stMarkdown"] [data-testid="stMarkdownContainer"] > style:only-child),
.element-container:has(> [data-testid="stMarkdown"] [data-testid="stMarkdownContainer"] > style:only-child){display:none}
.csv-card-head{display:flex;align-items:center;gap:8px;font-weight:800;font-size:1.02rem;margin:0 0 4px}
.csv-card-head .lab{display:inline-block;color:#fff;border-radius:6px;padding:1px 9px;font-size:.8rem}
.csv-card-head.green{color:#1f6b45}.csv-card-head.green .lab{background:#2e8a5b}
.csv-card-head.blue{color:#1d4e89}.csv-card-head.blue .lab{background:#2f6db5}
[class*="st-key-csvcard_"]{border:2.5px solid #2e8a5b;border-radius:12px;padding:12px 14px 10px;background:#eef8f1}
/* ドラッグ＆ドロップの枠：緑の点線で大きくし、「ここにドラッグ＆ドロップ」と書く */
[data-testid="stFileUploaderDropzone"]{border:2.5px dashed #2e8a5b !important;background:#fff !important;border-radius:12px !important;
  min-height:7.5rem;padding:1rem !important;display:flex !important;flex-direction:column !important;align-items:center !important;
  justify-content:center !important;gap:.45rem;text-align:center}
[data-testid="stFileUploaderDropzone"]::before{content:"📄⬇　ここにファイルをドラッグ＆ドロップ";order:-1;display:block;
  font-weight:800;color:#1f6b45;font-size:1rem}
[data-testid="stFileUploaderDropzone"] button{border:0 !important;color:#fff !important;background:#2e8a5b !important;font-weight:700}
[data-testid="stFileUploaderDropzone"] button p{color:#fff}
[class*="st-key-csvcard_"] [data-testid="stDownloadButton"] button{border:1.5px solid #2e8a5b !important;color:#1f6b45 !important;background:#fff !important;font-weight:700}
[data-testid="stFileUploaderDropzoneInstructions"]{font-size:.8rem;color:#8a94a6}
.imp-choose{font-weight:800;color:#24405f;margin:4px 0 8px;font-size:1rem}
.imp-or{display:flex;justify-content:center}
.imp-or span{background:#eef1f5;border-radius:999px;padding:6px 9px;font-weight:800;color:#6a7686;font-size:.85rem;white-space:nowrap;line-height:1.3}
/* 「または」の欄は、文字欄の下に付く負の余白を消して、上下の真ん中に置く（縦書きのように折り返さない） */
[data-testid="stColumn"]:has(.imp-or) [data-testid="stMarkdown"], [data-testid="stColumn"]:has(.imp-or) [data-testid="stMarkdownContainer"],
[data-testid="stColumn"]:has(.imp-or) [data-testid="stMarkdownContainer"] > *{margin:0 !important;padding:0 !important}
/* 中くらいの画面幅でも、方法1・または・方法2 を折り返さずに横1列に並べる（Streamlitは幅が足りないと列を次の行に送るため） */
@media (min-width:641px){
  [data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)){flex-wrap:nowrap !important}
  [data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)) > [data-testid="stColumn"]{min-width:0 !important;flex:1 1 0 !important}
  [data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)) > [data-testid="stColumn"]:has(.imp-or){flex:0 0 4.6rem !important;width:4.6rem !important}
}
.imp-note{background:#fff7e0;border:1px solid #f0c36d;border-radius:8px;padding:6px 10px;font-size:.85rem;color:#7a5a00;margin:10px 0 2px}
@media (max-width:640px){.imp-or{margin:0}}
/* 方法1（引き継ぎ）の枠は青 */
:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand){border-color:#2f6db5 !important;background:#eef5fd !important}
:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand) [data-testid="stPopover"] button, :is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand) [data-testid="stExpander"] summary{
  background:#2f6db5 !important;color:#fff !important;border:0 !important;min-height:2.8rem;font-weight:800}
:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand) [data-testid="stPopover"] button p{font-size:1.05rem;font-weight:800;color:#fff}
:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand) [data-testid="stPopover"] button svg{color:#fff;fill:#fff}
/* 方法1と方法2の枠の上端と高さをそろえる（横に並んだ列を同じ高さにして、枠を列いっぱいに伸ばす） */
[data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)){align-items:stretch !important}
[data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)) > [data-testid="stColumn"] > [data-testid="stVerticalBlock"],
[data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)) > [data-testid="column"] > div > [data-testid="stVerticalBlock"]{height:100%}
[data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)) [data-testid="stLayoutWrapper"]:has(> [class*="st-key-csvcard_"]){
  flex:1 1 auto;display:flex;flex-direction:column}
[data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)) [class*="st-key-csvcard_"]{flex:1 1 auto;box-sizing:border-box}
@media (max-width:640px){[data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)) [class*="st-key-csvcard_"]{flex:0 0 auto}}
/* 「または」は2つの枠の上下の真ん中に置く */
[data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)) > [data-testid="stColumn"]:has(.imp-or) > [data-testid="stVerticalBlock"],
[data-testid="stHorizontalBlock"]:has(:is(.st-key-csvcard_carry,.st-key-csvcard_fixedhand)) > [data-testid="column"]:has(.imp-or) > div > [data-testid="stVerticalBlock"]{justify-content:center}
/* 見出しと説明は上でそろえ、操作する場所（方法1の白い枠・方法2のドロップ枠）は枠の下でそろえる */
:is(.st-key-carry_act,.st-key-fixedhand_act){background:#fff;border:2.5px solid #b9cfee;border-radius:12px;min-height:7.5rem;padding:1rem;box-sizing:border-box;
  justify-content:center;align-items:center;text-align:center;gap:.45rem}
[class*="st-key-csvcard_"] [data-testid="stLayoutWrapper"]:has(> :is(.st-key-carry_act,.st-key-fixedhand_act)), [class*="st-key-csvcard_"] > :is(.st-key-carry_act,.st-key-fixedhand_act){margin-top:auto}
.st-key-csvcard_staff [data-testid="stElementContainer"]:has([data-testid="stFileUploader"]),
.st-key-csvcard_staff .element-container:has([data-testid="stFileUploader"]){margin-top:auto}
/* 「初めての場合：ひな形をダウンロード」は2つの枠の下に1行で出す */
.st-key-imp_tmpl{background:#f7f9fc;border:1px solid #e1e6ee;border-radius:10px;padding:8px 14px;margin-top:6px;
  flex-direction:row !important;flex-wrap:wrap;align-items:center;gap:.3rem .6rem}
.st-key-imp_tmpl > div{width:auto !important;flex:0 1 auto;align-self:center}
/* 文字欄の下に付く負の余白を消して、文字とボタンの高さ（上下の真ん中）をそろえる */
.st-key-imp_tmpl [data-testid="stMarkdown"], .st-key-imp_tmpl [data-testid="stMarkdownContainer"],
.st-key-imp_tmpl [data-testid="stMarkdownContainer"] > *{margin:0 !important;padding:0 !important}
.st-key-imp_tmpl p{margin:0 !important;font-size:.9rem;line-height:1.4}
.st-key-imp_tmpl [data-testid="stDownloadButton"] button{min-height:0 !important;padding:.25rem .8rem !important;line-height:1.4}
.st-key-imp_tmpl [data-testid="stDownloadButton"] button p{font-size:.88rem}
.st-key-csv_ng_ask{margin-top:14px;border:3px solid #d98a00;background:#fff4dc;border-radius:14px;padding:14px 18px 18px;
  box-shadow:0 3px 10px rgba(217,138,0,.18);gap:.6rem}
.csvng-h{display:flex;align-items:center;gap:10px;font-size:1.12rem;font-weight:800;color:#5a3d00;line-height:1.5}
.csvng-h .badge{flex:0 0 auto;background:#d98a00;color:#fff;border-radius:6px;padding:2px 10px;font-size:.85rem;font-weight:800}
.csvng-h b{color:#b34700}
.csvng-sub{font-size:.88rem;color:#7a5a00;margin-top:4px;padding-bottom:14px}
.st-key-csv_ng_ask [data-testid="stRadio"], .st-key-csv_ng_ask [data-testid="stRadio"] > div,
.st-key-csv_ng_ask [data-testid="stElementContainer"]:has([data-testid="stRadio"]){width:100% !important}
.csvng-wait{font-size:.9rem;color:#7a5a00;font-weight:700;padding:2px 0 4px}
.st-key-csv_ng_ask [role="radiogroup"]{gap:.5rem;width:100%;display:flex;flex-direction:column}
.st-key-csv_ng_ask [role="radiogroup"] label{background:#fff;border:1.5px solid #e9c97a;border-radius:10px;padding:10px 14px;
  width:100%;max-width:none;margin:0;box-sizing:border-box}
.st-key-csv_ng_ask [role="radiogroup"] label:hover{border-color:#d98a00;background:#fffaf0}
.st-key-csv_ng_ask [role="radiogroup"] label:has(input:checked){border:2px solid #d98a00;background:#fff1cc}
.st-key-csv_ng_ask [role="radiogroup"] label p{font-size:.95rem;font-weight:400;color:#5f6b7a}
.st-key-csv_ng_ask [role="radiogroup"] label p strong{font-size:1rem;color:#31333f}
.csvng-ok{margin-top:10px;background:#e8f5ec;border:1px solid #9fd0ad;border-radius:10px;padding:9px 14px;line-height:1.8;color:#1f4d33}
.csvng-ok a{color:#1d4e89 !important;text-decoration:underline;margin-left:6px;font-weight:700}
.st-key-imp_tmpl [data-testid="stDownloadButton"] button{border:1.5px solid #2e8a5b !important;color:#1f6b45 !important;background:#fff !important;font-weight:700}
/* 確定当直：方法1（表に直接入力）の「下の表へ」と、表の見出し・青い枠 */
.fx-go{font-weight:800;color:#1d4e89;font-size:1rem}.fx-go small{display:block;font-weight:400;color:#566174;font-size:.8rem;margin-top:2px}
.fx-table-h{font-weight:800;color:#1d4e89;font-size:1.05rem;margin:10px 0 0}.fx-table-h span{font-weight:400;color:#566174;font-size:.82rem;margin-left:10px}
.st-key-fixed_table_box [data-testid="stDataFrame"], .st-key-fixed_table_box [data-testid="stDataFrameResizable"]{border:2px solid #2f6db5;border-radius:8px}
/* 確定当直の項目：見出しの帯と左の太い線だけオレンジにする（医師条件の項目と、ぱっと見で見分けるため） */
.st-key-fixed_section{border-left:6px solid #e08a1e;background:#fff3e2;border-radius:0 12px 12px 0;padding:14px 16px 14px 14px;margin-top:2.6rem}
/* 方法1の白い枠は下まで伸ばして、説明との間にすき間ができないようにする（文字は枠の上下の真ん中） */
.st-key-csvcard_fixedhand [data-testid="stLayoutWrapper"]:has(> .st-key-fixedhand_act){flex:1 1 auto;display:flex;flex-direction:column;margin-top:0 !important}
.st-key-csvcard_fixedhand .st-key-fixedhand_act{flex:1 1 auto;margin-top:0 !important}
.st-key-fixed_section .sec-banner{margin-top:0;background:linear-gradient(90deg,#b5651d 0%,#e08a1e 100%) !important;box-shadow:0 2px 6px rgba(150,80,20,.2)}
@media (max-width:600px){.st-key-fixed_section{border-left-width:4px;padding:10px 8px 10px 8px}}
</style>
"""


def csv_upload_card(card_key, tag, title, desc, uploader_key, template_bytes, template_name, template_label, template_note="",
                    file_types=("csv",), drop_text=None):
    """CSVのアップロード欄（見出し・説明・ドラッグ＆ドロップの枠・ひな形のダウンロード）。読み込んだファイルを返す。
    見た目の指定（CSV_UPLOAD_CSS）は、画面の最初に1回だけ入れる（ここで入れると、枠の上に余白ができて高さがずれるため）。
    file_types で読み込めるファイルの種類を、drop_text で枠の中の文字（CSV以外も読めるとき）を変えられる。"""
    with st.container(key=f"csvcard_{card_key}"):
        drop_css = (f"<style>.st-key-csvcard_{card_key} [data-testid='stFileUploaderDropzone']::before{{content:'{drop_text}'}}</style>"
                    if drop_text else "")
        st.markdown(f"{drop_css}<div class='csv-card-head green'><span class='lab'>{tag}</span>{title}</div>", unsafe_allow_html=True)
        if desc:
            st.caption(desc)
        up = st.file_uploader(title, type=list(file_types), key=uploader_key, label_visibility="collapsed")
        if template_note:
            st.caption(template_note)
        if template_bytes is not None:   # 方法2（医師条件）は、ひな形のダウンロードを枠の下に出すので、ここには出さない
            st.download_button(label=template_label, data=template_bytes, file_name=template_name,
                               mime=XLSX_MIME if template_name.endswith(".xlsx") else "text/csv")
    return up


def _csv_cells(df, col):
    """CSVの列の、空でない値の一覧（先生ごと）。"""
    if col not in df.columns:
        return []
    vals = [clean_text(v) for v in df[col].tolist()]
    return [v for v in vals if v and v.lower() not in ("nan", "none", "<na>")]


def _csv_ng_choice_changed():
    """「使う／空にする」を選び直したら、②のカレンダーをCSVの内容から作り直す。"""
    for key in list(st.session_state.keys()):
        if key.startswith("ng_"):
            del st.session_state[key]


def csv_ng_wish_choice(df, uploaded_file, box, year, month):
    """①方法2：CSVにNG日・希望日が入っているときは、今月の分として使うか空にするかを選んでもらう。
    CSVには日にちしか入っていないため、前月のCSVだと前月のNG日が今月の同じ日に入ってしまうため。読み込んだ結果も表示する。"""
    ng_vals, wish_vals = _csv_cells(df, NG_COLUMN), _csv_cells(df, REQUEST_COLUMN)
    ng_days = sum(len([t for t in str(v).replace("、", ",").split(",") if t.strip()]) for v in ng_vals)
    n_doc = len(df)
    fname = html_escape(uploaded_file.name)
    with box:
        if not ng_vals and not wish_vals:
            st.success(f"ファイル「{uploaded_file.name}」を読み込みました：**{n_doc}名**（NG日・希望日は入っていません）")
            return df
        parts = []
        if ng_vals:
            parts.append(f"<b>NG日（{len(ng_vals)}名分）</b>")
        if wish_vals:
            parts.append(f"<b>希望日（{len(wish_vals)}名分）</b>")
        with st.container(key="csv_ng_ask"):
            st.markdown(f"<div class='csvng-h'><span class='badge'>確認</span><span>このファイルには {'と'.join(parts)}が入っています。どうしますか？</span></div>"
                        "<div class='csvng-sub'>ファイルには日にちしか入っていないため、前月のファイルだと前月のNG日・希望日が今月の同じ日に入ってしまいます。</div>",
                        unsafe_allow_html=True)
            opts = ["use", "clear"]
            labels = {"use": f"**今月（{year}年{month}月）のNG日・希望日として使う**　―　このファイルを今月用に作った・今月ダウンロードしたとき",
                      "clear": "**空にする（NG日・希望日は読み込まない）**　―　前月などのファイルを使うとき。NG日は②で入れ直します"}
            choice = st.radio("ファイルのNG日・希望日", opts, index=None, format_func=lambda o: labels[o],
                              key=f"csvng_choice_{uploaded_file.file_id}", label_visibility="collapsed",
                              on_change=_csv_ng_choice_changed)
            if choice is None:
                st.markdown("<div class='csvng-wait'>⏳ どちらかを選ぶまで、NG日・希望日は入れていません（回数などの医師条件は読み込み済みです）。</div>",
                            unsafe_allow_html=True)
        if choice == "use":
            lines = []
            if ng_vals:
                lines.append(f"・NG日：<b>{len(ng_vals)}名・のべ{ng_days}日</b>を ② のカレンダーに入れました　<a href='#sec-ng'>→ ② NG日を見る</a>")
            if wish_vals:
                lines.append(f"・希望日：<b>{len(wish_vals)}名</b>を ③ の表の「希望日」に入れました")
            st.markdown(f"<div class='csvng-ok'>✅ ファイル「{fname}」を読み込みました：<b>{n_doc}名</b><br>{'<br>'.join(lines)}</div>",
                        unsafe_allow_html=True)
            return df
    df = df.copy()
    for col in (NG_COLUMN, REQUEST_COLUMN):
        if col in df.columns:
            df[col] = ""
    if choice == "clear":
        with box:
            st.markdown(f"<div class='csvng-ok'>✅ ファイル「{fname}」を読み込みました：<b>{n_doc}名</b><br>"
                        "・NG日・希望日は空にしました（② のカレンダーで入れてください）　<a href='#sec-ng'>→ ② NG日を入れる</a></div>",
                        unsafe_allow_html=True)
    return df


# ==========================================
# 作る月を選ぶ（大きな「2026年12月」＋前の月／次の月＋今月・来月・再来月のボタン）
# ==========================================
YM_PICK_CSS = """<style>
.st-key-ym_pick{flex-direction:row !important;flex-wrap:wrap;align-items:center;gap:.6rem 1rem !important;
  border:2px solid #b9cfee;background:#f3f8fe;border-radius:12px;padding:10px 16px}
.st-key-ym_pick > div{width:auto !important;flex:0 0 auto}
.st-key-ym_pick > div:has(.st-key-ym_quick), .st-key-ym_pick > .st-key-ym_quick{margin-left:auto}
.st-key-ym_quick{flex-direction:row !important;flex-wrap:wrap;gap:.4rem !important}
.st-key-ym_quick > div{width:auto !important}
/* Streamlitの文字欄は下に負の余白が付くため、そのままだと上下の真ん中より下にずれる。余白を消してボタンと高さをそろえる */
.st-key-ym_pick > div{align-self:center}
.st-key-ym_pick [data-testid="stMarkdown"], .st-key-ym_pick [data-testid="stMarkdownContainer"],
.st-key-ym_pick [data-testid="stMarkdownContainer"] > *{margin:0 !important;padding:0 !important}
.st-key-ym_inputs{flex-direction:row !important;align-items:center;gap:.35rem !important}
.st-key-ym_inputs > div{width:auto !important;flex:0 0 auto;align-self:center}
.st-key-ym_inputs [data-testid="stMarkdown"], .st-key-ym_inputs [data-testid="stMarkdownContainer"],
.st-key-ym_inputs [data-testid="stMarkdownContainer"] > *{margin:0 !important;padding:0 !important}
.st-key-ym_year_in{width:8.2rem !important}.st-key-ym_month_in{width:5.6rem !important}
.st-key-ym_year_in [data-baseweb="select"] > div, .st-key-ym_month_in [data-baseweb="select"] > div{
  background:#fff !important;border:1.5px solid #9fb6d6 !important;border-radius:8px !important;min-height:3rem}
.st-key-ym_year_in [data-baseweb="select"] *, .st-key-ym_month_in [data-baseweb="select"] *{font-size:1.65rem;font-weight:800;color:#1d4e89}
.st-key-ym_year_in [data-baseweb="select"] svg, .st-key-ym_month_in [data-baseweb="select"] svg{width:1.3rem;height:1.3rem;color:#1d4e89}
.ym-unit{font-size:1.05rem;font-weight:800;color:#1d4e89}
.ym-big{font-size:2rem;font-weight:800;color:#1d4e89;min-width:12rem;text-align:center;line-height:1;white-space:nowrap;margin:0 !important}
.ym-big small{font-size:1rem;margin:0 .25rem 0 .1rem}
.st-key-ym_prev button,.st-key-ym_next button{border:1.5px solid #2f6db5 !important;color:#1d4e89 !important;background:#fff !important;font-weight:700}
.st-key-ym_quick button{border-radius:999px !important;min-height:2.1rem;padding:0 .9rem}
.st-key-ym_quick button[kind="primary"],.st-key-ym_quick button[data-testid="stBaseButton-primary"]{background:#2f6db5 !important;border-color:#2f6db5 !important;color:#fff !important}
.ym-sum{margin-top:6px;background:#f7f9fc;border:1px solid #e1e6ee;border-radius:10px;padding:8px 12px;font-weight:700;color:#24405f}
.ym-sum span{font-weight:400;font-size:.82rem;color:#8a94a6;margin-left:.6rem}
.ym-per{display:flex;align-items:center;gap:14px;flex-wrap:wrap;margin-top:10px;border:1px solid #e1e6ee;border-left:8px solid var(--c);
  background:var(--bg);border-radius:12px;padding:12px 18px;color:var(--fg)}
.ym-per.to-period{--c:#7a52c2;--bg:#f1eafb;--fg:#3f2370}
.ym-per.to-normal{--c:#2f6db5;--bg:#e8f1fb;--fg:#173f73}
.ym-per .ic{display:inline-flex;flex:0 0 auto;width:2.6rem;height:2.6rem;border-radius:50%;background:var(--c);align-items:center;justify-content:center;font-size:1.35rem}
.ym-per .tx{flex:1 1 16rem;min-width:0}
.ym-per .tt{font-weight:800;font-size:1.05rem;line-height:1.4}
.ym-per .ps{font-size:.84rem;opacity:.85;margin-top:2px}
.ym-per .pbtn{margin-left:auto;background:var(--c);color:#fff !important;border-radius:9px;padding:.55rem 1.1rem;font-weight:800;
  font-size:.95rem;text-decoration:none !important;white-space:nowrap}
.ym-per .pbtn:hover{filter:brightness(1.08)}
@media (max-width:640px){.ym-big{font-size:1.6rem;min-width:0}.st-key-ym_pick > div:has(.st-key-ym_quick){margin-left:0}}
</style>"""


def _ym_set(y, m):
    st.session_state["sel_year"], st.session_state["sel_month"] = int(y), int(m)


def _ym_from_inputs():
    """年・月の入力欄で選んだとき。"""
    _ym_set(st.session_state.get("ym_year_in", st.session_state["sel_year"]),
            st.session_state.get("ym_month_in", st.session_state["sel_month"]))


def _ym_shift(delta):
    t = int(st.session_state["sel_year"]) * 12 + int(st.session_state["sel_month"]) - 1 + delta
    _ym_set(t // 12, t % 12 + 1)


def month_picker(today):
    """作る年月を選ぶ。選んだ年月は session_state の sel_year / sel_month に入る（途中保存の読み込みもここを書き換える）。"""
    st.markdown(YM_PICK_CSS, unsafe_allow_html=True)
    y, m = int(st.session_state["sel_year"]), int(st.session_state["sel_month"])
    with st.container(key="ym_pick"):
        st.button("◀ 前の月", key="ym_prev", on_click=_ym_shift, args=(-1,))
        # 年と月は、大きな入力欄で直接選べる（年は数字を打ち込んで絞り込める）
        if st.session_state.get("ym_year_in") != y:
            st.session_state["ym_year_in"] = y
        if st.session_state.get("ym_month_in") != m:
            st.session_state["ym_month_in"] = m
        _years = sorted(set(range(today.year - 3, today.year + 6)) | {y})
        with st.container(key="ym_inputs"):
            st.selectbox("年", _years, key="ym_year_in", label_visibility="collapsed", on_change=_ym_from_inputs)
            st.markdown("<span class='ym-unit'>年</span>", unsafe_allow_html=True)
            st.selectbox("月", list(range(1, 13)), key="ym_month_in", label_visibility="collapsed", on_change=_ym_from_inputs)
            st.markdown("<span class='ym-unit'>月</span>", unsafe_allow_html=True)
        st.button("次の月 ▶", key="ym_next", on_click=_ym_shift, args=(1,))
        with st.container(key="ym_quick"):
            for i, name in enumerate(("今月", "来月", "再来月")):
                t = today.year * 12 + today.month - 1 + i
                qy, qm = t // 12, t % 12 + 1
                label = f"{name}（{qm}月）" if qy == today.year else f"{name}（{qy}年{qm}月）"
                st.button(label, key=f"ym_q{i}", type="primary" if (qy, qm) == (y, m) else "secondary",
                          on_click=_ym_set, args=(qy, qm))
    return y, m


def mode_switch_box(to):
    """通常版⇔期間指定版を切りかえる入口（どちらの画面でも「① 作る月／期間を選ぶ」の下の同じ場所に、同じ形で出す）。
    色は行き先に合わせる（期間指定版へ＝紫、通常版へ＝青）。"""
    if to == "period":
        cls, icon, title, sub, href, btn = ("to-period", "📆", "年末年始など、翌月の数日まで通しで作るとき",
                                            "今月1日〜翌月の指定日を、1つの当直案として作ります", "./?mode=period", "期間指定版を開く →")
    else:
        cls, icon, title, sub, href, btn = ("to-normal", "🏠", "ふだんの月（1日〜末日）の当直を作るとき",
                                            "1か月分を、通常版で作ります", "./", "通常版に戻る →")
    st.markdown(f"<div class='ym-per {cls}'><span class='ic'>{icon}</span><div class='tx'><div class='tt'>{title}</div>"
                f"<div class='ps'>{sub}</div></div><a class='pbtn' href='{href}' target='_self'>{btn}</a></div>", unsafe_allow_html=True)


def ym_summary(year, month):
    """選んだ月の期間・日数・土日祝の数（選び間違いに気づけるように）。"""
    W = "月火水木金土日"
    n = calendar.monthrange(year, month)[1]
    first, last = datetime.date(year, month, 1), datetime.date(year, month, n)
    hol = sum(1 for d in range(1, n + 1)
              if datetime.date(year, month, d).weekday() >= 5 or jpholiday.is_holiday(datetime.date(year, month, d)))
    st.markdown(f"<div class='ym-sum'>🗓 {year}/{month}/1（{W[first.weekday()]}）〜 {month}/{n}（{W[last.weekday()]}）・{n}日間・土日祝 {hol}日"
                "<span>（土日祝は自動で休日になります）</span></div>", unsafe_allow_html=True)


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
                _entries = work_state_entries()
                if _entries:
                    _labels = {sid: f"{str(at)[:16]}　{work_state_title(d)}の作業" for sid, at, d in _entries}
                    # 途中保存した直後は、いちばん新しい保存（今保存したもの）を選んだ状態にする
                    # （前に選んでいた古い保存が選ばれたままだと、「この作業を読み込む」で古い内容が戻ってしまう）
                    if st.session_state.pop("work_restore_pick_newest", False) or st.session_state.get("work_restore_pick") not in _labels:
                        st.session_state["work_restore_pick"] = _entries[0][0]
                    st.selectbox(f"読み込む途中保存（新しい順・最大{WORK_HISTORY_KEEP}回分）", list(_labels),
                                 format_func=lambda sid: _labels[sid], key="work_restore_pick")
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
  <div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">日直<span>NG</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">宿直<span>NG</span></div></div></div><div style="margin-top:4px;color:#243247">全NG</div></div><div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#fff;color:#243041;">日直<span style="color:#2f8a5b">可</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">宿直<span>NG</span></div></div></div><div style="margin-top:4px;color:#243247">宿直だけNG<br><span style="color:#566174">日直は入れる</span></div></div><div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">日直<span>NG</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#fff;color:#243041;border-top:1px solid #e3e7ed;">宿直<span style="color:#2f8a5b">可</span></div></div></div><div style="margin-top:4px;color:#243247">日直だけNG<br><span style="color:#566174">宿直は入れる</span></div></div>
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
  <div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">日直<span>NG</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">宿直<span>NG</span></div></div></div><div style="margin-top:4px;color:#243247">全NG</div></div><div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#fff;color:#243041;">日直<span style="color:#2f8a5b">可</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">宿直<span>NG</span></div></div></div><div style="margin-top:4px;color:#243247">宿直だけNG<br><span style="color:#566174">日直は入れる</span></div></div><div style="width:104px;text-align:center"><div style="background:#fff6f5;border:1px solid #f0dcdc;border-radius:10px;padding:6px"><div style="border:1px solid #d8dee6;border-radius:7px;overflow:hidden"><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#cf6a35;color:#fff">日直<span>NG</span></div><div style="display:flex;justify-content:center;gap:6px;padding:5px 0;background:#fff;color:#243041;border-top:1px solid #e3e7ed;">宿直<span style="color:#2f8a5b">可</span></div></div></div><div style="margin-top:4px;color:#243247">日直だけNG<br><span style="color:#566174">宿直は入れる</span></div></div>
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


# ==========================================
# 紙で配る「当直不可日」の用紙（Excel・A4縦1枚）。先生方の専用URLの入力画面に近い見た目にする
# ==========================================
PAPER_NOTES_DEFAULT = "記名なき者は無効とさせていただきます。\n働き方改革を考慮ください。\n遅延は一切認めません。"
PAPER_NOTES_KEY = "paper_form_notes"


def paper_ng_form_xlsx(title, days, notes, show_locked=False):
    """
    紙で配る「当直不可日」の用紙（A4縦1枚のExcel）を作る。
    days は [(日付, 休日か, 別途決定か, 表に出す日付の文字)]。氏名・締め切り・提出先・管理者へのコメントは空欄（印刷してから手で書く）。
    休日は「日直」「宿直」の2つの欄、平日は「宿直」の欄だけ（日直の欄は斜線）。別途決定の日は、show_locked でなければ灰色の点線のマスにする。
    白黒で印刷しても分かるように、色だけに頼らない（休日は日付の帯を灰色にして「休日」と書き、枠を太くする。平日の日直は濃い斜線）。
    """
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "当直不可日"
    ws.sheet_view.showGridLines = False
    font = "Meiryo"
    line, hol_line, pale = "7F7F7F", "262626", "8C8C8C"   # 白黒で印刷しても見分けられる濃さ
    # 列：曜日ごとに「マスの列」と「すき間の列」を交互に並べる（B,D,F,…がマス）
    ws.column_dimensions["A"].width = 1.5
    day_cols = []
    for i in range(7):
        c = 2 + i * 2
        day_cols.append(c)
        ws.column_dimensions[get_column_letter(c)].width = 12.2
        if i < 6:
            ws.column_dimensions[get_column_letter(c + 1)].width = 1.0
    first, last = get_column_letter(day_cols[0]), get_column_letter(day_cols[-1])

    def put(row, col, value=None, size=11, bold=False, color="1F2937", align="left", valign="center"):
        cell = ws.cell(row=row, column=col, value=value)
        cell.font = Font(name=font, size=size, bold=bold, color=color)
        cell.alignment = Alignment(horizontal=align, vertical=valign, wrap_text=False)
        return cell

    def underline(row, c1, c2):
        for c in range(c1, c2 + 1):
            ws.cell(row=row, column=c).border = Border(bottom=Side(style="thin", color="111111"))

    r = 2
    ws.row_dimensions[r].height = 34
    put(r, day_cols[0], title, size=18, bold=True)
    r += 2
    ws.row_dimensions[r].height = 24
    put(r, day_cols[0], "氏名：", size=13, bold=True)
    underline(r, day_cols[1], day_cols[4])
    r += 2
    ws.row_dimensions[r].height = 22
    put(r, day_cols[0], "締め切り：", size=11)
    underline(r, day_cols[1], day_cols[3])
    put(r, day_cols[3] + 1, "まで", size=11)
    r += 1
    ws.row_dimensions[r].height = 22
    put(r, day_cols[0], "（消化器内科は", size=11)
    underline(r, day_cols[1], day_cols[1] + 1)
    put(r, day_cols[2], "まで、その他は", size=11)
    underline(r, day_cols[3], day_cols[3] + 1)
    put(r, day_cols[4], "まで）", size=11)
    r += 1
    for text in [t for t in str(notes or "").splitlines() if t.strip()]:
        ws.row_dimensions[r].height = 18
        put(r, day_cols[0], text.strip(), size=10.5)
        r += 1
    r += 1
    ws.row_dimensions[r].height = 34
    ws.merge_cells(f"{first}{r}:{last}{r}")
    how = put(r, day_cols[0], "入れない日の「日直」「宿直」の欄に×を書いてください（休日は日直・宿直の両方、平日は宿直だけ。斜線の欄は書かなくてよい）。", size=10)
    how.fill = PatternFill("solid", fgColor="EDEDED")
    how.alignment = Alignment(horizontal="left", vertical="center", wrap_text=True)
    r += 2
    for i, w in enumerate(WEEKDAYS_JA):
        put(r, day_cols[i], w, size=12, bold=True, align="center",
            color="1F4E8C" if i == 5 else ("A61B29" if i == 6 else "262626"))
    r += 1

    # カレンダー：月曜はじまり。範囲の外の日は空ける
    start = days[0][0] - datetime.timedelta(days=days[0][0].weekday())
    end = days[-1][0] + datetime.timedelta(days=6 - days[-1][0].weekday())
    info = {d: (hol, locked, label) for d, hol, locked, label in days}
    cur = start
    while cur <= end:
        r_date, r_day, r_night = r, r + 1, r + 2
        ws.row_dimensions[r_date].height = 22
        ws.row_dimensions[r_day].height = 30
        ws.row_dimensions[r_night].height = 30
        ws.row_dimensions[r + 3].height = 6
        for i in range(7):
            d = cur + datetime.timedelta(days=i)
            c = day_cols[i]
            if d not in info:
                continue
            hol, locked, label = info[d]
            if locked and not show_locked:
                side = Side(style="dashed", color=line)
                for rr, top, bottom in ((r_date, True, False), (r_day, False, False), (r_night, False, True)):
                    cell = ws.cell(row=rr, column=c)
                    cell.fill = PatternFill("solid", fgColor="BFBFBF")
                    cell.border = Border(left=side, right=side, top=side if top else None, bottom=side if bottom else None)
                put(r_date, c, label, size=13, bold=True, color="595959", align="center")
                continue
            outline = Side(style="medium" if hol else "thin", color=hol_line if hol else line)   # 休日は太い枠
            inner = Side(style="thin", color=line)
            color = "1F4E8C" if (hol and d.weekday() == 5) else ("A61B29" if hol else "1F2937")
            dc = put(r_date, c, f"{label}　休日" if hol else label, size=13, bold=True, color=color, align="center")
            if hol:
                # 日付の後ろの「休日」は小さく、黒で（白黒でも休日と分かるように）
                try:
                    from openpyxl.cell.rich_text import CellRichText, TextBlock
                    from openpyxl.cell.text import InlineFont
                    dc.value = CellRichText([TextBlock(InlineFont(rFont=font, sz=13, b=True, color=color), label),
                                             TextBlock(InlineFont(rFont=font, sz=8, b=True, color="262626"), "  休日")])
                except ImportError:
                    pass
            dc.fill = PatternFill("solid", fgColor="D9D9D9" if hol else "FFFFFF")
            dc.border = Border(left=outline, right=outline, top=outline, bottom=inner)
            day_cell = ws.cell(row=r_day, column=c)
            if hol:
                put(r_day, c, "日直", size=7.5, bold=True, color="404040", valign="top")
            else:
                day_cell.fill = PatternFill("lightUp", fgColor=pale, bgColor="FFFFFF")   # 平日の日直は書かなくてよい（斜線）
            day_cell.border = Border(left=outline, right=outline, bottom=inner)
            put(r_night, c, "宿直", size=7.5, bold=True, color="404040", valign="top")
            ws.cell(row=r_night, column=c).border = Border(left=outline, right=outline, bottom=outline)
        cur += datetime.timedelta(days=7)
        r += 4

    # 専用URLの画面と同じく、カレンダーの下に「管理者へのコメント・要望（任意）」の欄
    r += 1
    ws.row_dimensions[r].height = 20
    put(r, day_cols[0], "管理者へのコメント・要望（任意）", size=11, bold=True)
    r += 1
    box = Side(style="thin", color=line)
    for k in range(4):
        ws.row_dimensions[r + k].height = 22
        for c in range(day_cols[0], day_cols[-1] + 1):
            ws.cell(row=r + k, column=c).border = Border(left=box if c == day_cols[0] else None, right=box if c == day_cols[-1] else None,
                                                         top=box if k == 0 else None, bottom=box if k == 3 else None)
    r += 4

    ws.print_area = f"A1:{last}{r}"
    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_margins.left = ws.page_margins.right = 0.5
    ws.page_margins.top = ws.page_margins.bottom = 0.5
    ws.print_options.horizontalCentered = True
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def paper_ng_form_ui(title, days, file_label):
    """
    ②（NG日を入れる）の「🖨️ 紙で配る当直不可日の用紙」の欄。ふだんは閉じておける。
    days は paper_ng_form_xlsx と同じ形。注意書きはスイッチを入れたときだけ書き換えられる（途中保存に入る）。
    """
    if PAPER_NOTES_KEY not in st.session_state:
        st.session_state[PAPER_NOTES_KEY] = PAPER_NOTES_DEFAULT
    # 管理用の欄（紺の枠）と同じ形で、一歩控えめに：細く薄い紺の枠と、白抜きのしるし
    st.markdown("<style>.st-key-paper_form_box{border:1.5px solid #9fb0cc;background:#f7f9fc;border-radius:12px;padding:10px 12px 6px}"
                ".st-key-paper_form_box [data-testid=\"stExpander\"] details{background:#fff}"
                ".paper-badge{display:inline-flex;align-items:center;gap:6px;font-size:.85rem;font-weight:800;color:#3b5b92;background:#fff;"
                "border:1.5px solid #3b5b92;border-radius:999px;padding:1px 11px}"
                ".paper-note{font-size:.85rem;color:#3b4a63;margin-left:8px}</style>", unsafe_allow_html=True)
    box = st.container(key="paper_form_box")
    with box:
        st.markdown("<span class='paper-badge'>🖨️ 紙で集める</span>"
                    "<span class='paper-note'>URLを使わない先生に、印刷して配る用紙です（パスワードなしで使えます）</span>",
                    unsafe_allow_html=True)
    with box, st.expander("🖨️ 紙で配る当直不可日の用紙", expanded=False):
        st.caption("紙でNG日を集めるときの用紙です（A4縦1枚のExcel）。休日と別途決定の日は、上の「年月・休日」の設定のとおりです。"
                   "氏名・締め切り・提出先は空欄なので、印刷してから手で書きます。")
        show_locked = st.checkbox("別途決定の日も書いてもらう", key="paper_form_show_locked",
                                  help="チェックなし：別途決定の日は灰色のマスにします（書かなくてよい）。"
                                       "チェックあり：別途決定の日も、ほかの日と同じ書く欄を出します（書いてもらった×は、別の方法で当番を決めるときの参考用。アプリのカレンダーには入れられません）。")
        try:
            data = paper_ng_form_xlsx(title, days, st.session_state[PAPER_NOTES_KEY], show_locked)
            st.download_button("🖨️ 当直不可日の用紙（Excel）をダウンロード", data, file_name=file_label, mime=XLSX_MIME, key="paper_form_download")
        except ImportError:
            st.error("Excelで作るには、Excelを扱う部品（openpyxl）が必要です。requirements.txt に「openpyxl」の1行を追加してください。")
        if st.toggle("✏️ 用紙の注意書きを変える", key="paper_form_edit_notes"):
            def _save_notes():
                st.session_state[PAPER_NOTES_KEY] = st.session_state.get("paper_form_notes_edit", "")

            def _reset_notes():
                st.session_state[PAPER_NOTES_KEY] = PAPER_NOTES_DEFAULT
                st.session_state["paper_form_notes_edit"] = PAPER_NOTES_DEFAULT
            if "paper_form_notes_edit" not in st.session_state:
                st.session_state["paper_form_notes_edit"] = st.session_state[PAPER_NOTES_KEY]
            st.text_area("注意書き（1行ずつ。用紙の締め切りの下に出ます）", key="paper_form_notes_edit", height=100, on_change=_save_notes)
            st.button("初めの文に戻す", key="paper_form_reset_notes", on_click=_reset_notes)
            st.caption("書き換えた注意書きは途中保存に入るので、翌月も同じ文が使えます。")


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
            st.download_button("📥 URL一覧をExcelでダウンロード", table_xlsx(pd.DataFrame(url_rows), "URL一覧"),
                               file_name="NG日入力_専用URL一覧.xlsx", mime=XLSX_MIME, key="remote_url_csv")
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
                st.caption("取り込むと、その先生のNG日は提出内容に置き換わります（ファイルなどで入れていたNGは消えます）。"
                           "選ばなかった先生や、提出していない先生のNG日は変わりません。取り込んだあとも、下のカレンダーで個別に直せます。")
            else:
                st.caption("提出内容は、すべてアプリに取り込み済みです。")
        # --- 取り込みの取り消し ---
        undoable = import_undo_doctors(year, month, [n for n, _ in doctor_rows])
        if undoable:
            with st.container(border=True):
                st.markdown("**↩️ 取り込みを取り消す**")
                st.caption("URLから取り込んだNG日を取り消して、取り込む前のNG日（ファイルや手入力の分）に戻します。"
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
    月上限・休日上限・宿直上限・日直上限・各枠の上限のうち、いちばん厳しいものが効く。
    """
    total = _cap_num(row.get("月上限"), 0)
    hol = _cap_num(row.get("休日上限"), 0)
    soft = soft_cap_slots(row)   # 空欄の枠（できれば2回まで）は、足りないときは3回目以上も入るので、枠の上限では頭打ちにしない
    per = {s: (total if s in soft else _cap_num(row.get(s + "上限"), 0)) for s in ALL_SHIFT_TYPES}
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
    rows = [("全体（日直＋宿直）", needs["total"], sum(c["total"] for c in caps), "月上限"),
            ("宿直（A・B・外来の合計）", needs["night"], sum(c["night"] for c in caps), "宿直上限"),
            ("日直（A・B・外来の合計）", needs["day"], sum(c["day"] for c in caps), "日直上限"),
            ("休日の枠（日直＋宿直）", needs["holiday"], sum(c["holiday"] for c in caps), "休日上限")]
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
    day_intervals=None,
):
    """確定同士は維持し、自動割当と全勤務の間隔を制限する（宿直→宿直は宿⇔宿 間隔下限、日直がからむ組み合わせは宿⇔日・日⇔日 間隔下限）。"""
    _add_typed_interval_constraints(model, shifts, doctors, daily_active_shifts, num_days, period.date, min_intervals, day_intervals,
                                    past_worked_dates, future_worked_dates, absolute_req_days, absolute_req_specific)


def _period_generate_core(P, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None, locked_days=(), annual_balance=None):
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
    day_intervals = {}   # 日直がからむ勤務の間隔
    min_shifts_total = {}
    max_shifts_total = {}
    max_hol_shifts_per_doc = {}
    max_shifts_per_type = {}
    soft_caps = {}   # 枠ごとの上限が空欄の枠（できれば2回まで）
    group_caps = {}

    absolute_req_days = {doc: [] for doc in doctors}
    absolute_req_specific = {doc: [] for doc in doctors}
    past_worked_dates = {doc: [] for doc in doctors}
    future_worked_dates = {doc: [] for doc in doctors}

    # --- 確定当直（前月・今月・翌月）の読み取り ---
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
                            past_worked_dates[doc_val].append((date_obj, duty_kind(s_type)))   # 日付と、日直か宿直か
                        else:
                            future_worked_dates[doc_val].append((date_obj, duty_kind(s_type)))

    # --- 医師条件の読み取り ---
    for _, row in staff_df.iterrows():
        doc = str(row['先生の名前'])
        hard_weekdays[doc] = pm_duty_weekdays(row.get('翌日PM duty', ''))
        ng_days_dict[doc] = period_parse_ng(row.get(NG_COLUMN, ''), P)
        req_days[doc], req_specific[doc] = period_parse_requests(row.get(REQUEST_COLUMN, ''), P)
        req_days[doc] = [d for d in req_days[doc] if d not in locked_days]
        req_specific[doc] = [(d, s) for d, s in req_specific[doc] if d not in locked_days]
        req_priority[doc] = safe_int(row.get('希望優先度(数字が大きいほど優先)'), 1)
        min_intervals[doc] = safe_int(row.get('宿⇔宿 間隔下限'), 5)
        day_intervals[doc] = safe_int(row.get(DAY_GAP_COLUMN), min_intervals[doc])
        min_shifts_total[doc] = safe_int(row.get('月下限'), 0)
        max_shifts_total[doc] = safe_int(row.get('月上限'), 5)
        max_hol_shifts_per_doc[doc] = safe_int(row.get('休日上限'), 4)
        max_shifts_per_type[doc] = {s: safe_int(row.get(s + '上限'), 2) for s in ALL_SHIFT_TYPES}
        soft_caps[doc] = soft_cap_slots(row)
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
                if not active_days or s_type in soft_caps[doc]:   # 空欄の枠は決まりにせず、目的の減点で「できれば2回まで」にする
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
            absolute_req_specific=absolute_req_specific, day_intervals=day_intervals,
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

    # 月下限（目標。届かない分は減点）
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

    # グループの設定で△（できれば入れない）の枠は、入れるたびに減点する（人が足りないときだけ入る）
    objective_terms += group_avoid_terms(doctors, days, daily_active_shifts, shifts, staff_df)
    # 枠ごとの上限が空欄の枠は、できれば2回まで（1つの枠に偏らないように）
    objective_terms += soft_cap_terms(model, doctors, days, daily_active_shifts, shifts, soft_caps)
    # 回数の勾配：枠に余裕があるとき、グループごとに回数を多め・少なめにする（対象外のグループは入れない）
    objective_terms += group_level_terms(model, doctors, worked_all, staff_df)
    # 多め希望（病棟/外来・宿直/日直・平日/休日）：希望の種類に寄せる（月の回数は増やさない）
    objective_terms += doc_pref_terms(model, doctors, days, daily_active_shifts, shifts, staff_df, is_holiday, required_count,
                                      lambda doc: min_shifts_total[doc], total_cap)
    # 年間の回数（1か月あたり）をそろえる（年間の当直回数を読み込み、反映するにしたときだけ）
    objective_terms += add_annual_balance_terms(model, doctors, days, daily_active_shifts, is_holiday, required_count,
                                                worked_all, holiday_worked, shifts, annual_balance)
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
        return None, False, ["当直案の計算を始められませんでした。医師条件の数字（回数・間隔など）に、大きすぎる値やおかしな値がないか確認してください。"], None, None
    if status == cp_model.UNKNOWN:
        reasons.append("すべての枠を埋める案は、時間内に見つかりませんでした（条件の上で、ぜったいに無理と決まったわけではありません）。")
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
            reasons.append("空き枠を残しても、条件どうしがぶつかって当直案を作れませんでした。確定当直と回数の上限を確認してください。")
        elif relax_status == cp_model.MODEL_INVALID:
            reasons.append("当直案の計算を始められませんでした。医師条件の数字（回数・間隔など）に、大きすぎる値やおかしな値がないか確認してください。")
        else:
            reasons.append("空き枠を残す案も、時間内に見つかりませんでした（条件の上で、ぜったいに無理と決まったわけではありません）。")
    except Exception as e:
        reasons.append(f"⚠️ 不足を含む当直案の作成中にもエラーが発生しました。詳細: {e}")

    return None, False, reasons, None, None


def period_validate_staff_inputs(df, P):
    df = df.copy()
    errors = []
    if '先生の名前' not in df:
        return df, ['医師条件のファイルに「先生の名前」の列がありません。']
    df['先生の名前'] = df['先生の名前'].map(clean_text)
    blank = df['先生の名前'].eq('')
    for i, row in df[blank].iterrows():
        if any(clean_text(v) for k, v in row.items() if k != '先生の名前'):
            errors.append(f'医師条件 {i+1}行目：条件が入力されていますが先生の名前が空欄です。')
    df = df[~blank].reset_index(drop=True)
    for name in df.loc[df['先生の名前'].duplicated(), '先生の名前'].unique():
        errors.append(f'先生の名前「{name}」が重複しています。')
    df = mark_soft_slot_caps(df)   # 枠ごとの上限の空欄（できれば2回まで）を、2が入る前に覚えておく
    for col, default in NUMERIC_DEFAULTS.items():
        if col not in df:
            df[col] = default
    for col in [NG_COLUMN, REQUEST_COLUMN, '翌日PM duty', DAY_GAP_COLUMN] + list(GROUP_CAP_COLUMNS) + list(DOC_PREF_COLUMNS):
        if col not in df:
            df[col] = ''
    errors += doc_pref_errors(df)
    for col in list(GROUP_CAP_COLUMNS) + [DAY_GAP_COLUMN]:
        df[col] = df[col].astype(object)   # 空欄（制限なし／宿⇔宿 間隔下限と同じ）と数字が混ざるため
    for i, row in df.iterrows():
        name = row['先生の名前']
        if any(c in name for c in ',、') or name == '-' or '⚠️不足' in name:
            errors.append(f'先生の名前「{name}」：カンマ・読点・不足表示・単独の「-」は使用できません。')
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
            if float(df.at[i, '月下限']) > float(df.at[i, '月上限']):
                errors.append(f'{name}：月下限は月上限以下にしてください。')
        except (ValueError, TypeError):
            pass
        # 宿⇔日・日⇔日 間隔下限は空欄なら「宿⇔宿 間隔下限」と同じ値にする
        raw = clean_text(row.get(DAY_GAP_COLUMN, ''))
        if raw.lower() in ('', 'nan', 'none', '<na>'):
            df.at[i, DAY_GAP_COLUMN] = df.at[i, '宿⇔宿 間隔下限']
        else:
            try:
                n = float(raw)
                if not math.isfinite(n) or n < 0 or n != int(n) or n > 1000000:
                    raise ValueError()
                df.at[i, DAY_GAP_COLUMN] = int(n)
            except (TypeError, ValueError, OverflowError):
                errors.append(f'{name}：宿⇔日・日⇔日 間隔下限は0以上の整数にしてください（空欄なら宿⇔宿 間隔下限と同じ）。')
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
    """確定当直の入力チェック。holidays を渡すと、対象月の平日に日直が入っていないかも確認する。"""
    year, month = P.year, P.month
    errors = []
    if df is None:
        return None, errors
    df = df.copy()
    if '日付' not in df:
        return df, ['確定当直のファイルに「日付」の列がありません。']
    unknown = [c for c in df if c not in ['日付', '平日/休日', FIXED_KUBUN_COLUMN] + ALL_SHIFT_TYPES]
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
            errors.append(f'確定当直 {i+1}行目：日付「{raw}」を確認してください。「{P.example()}」または「{P.year}/{P.example()}」の形で入力します。')
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
                # 期間の前後の日付は、医師条件にいない先生（前月だけ勤務した先生など）でもよい（勤務間隔の計算には使わない）
                if name not in doctors and P.index(dt) is not None:
                    errors.append(f'確定当直 {dt.month}/{dt.day} {s}：先生の名前「{name}」が医師条件の名前と一致しません。期間中の日付には、医師条件の先生を入れてください。')
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
                worked.append((dt, duty_kind(slot))); counts[slot] += 1
                hol_count += int(r['平日/休日'] == '休日')
                kind = ng.get(P.index(dt), 'OK')
                night = slot in NIGHT_SHIFTS
                if kind == '全NG' or (kind == '宿NG' and night) or (kind == '日NG' and not night):
                    warnings.append(f'{name}：{dt.month}/{dt.day} {slot}は確定当直のため、NG日より優先して入れました。')
                if night and period_pm_restricted(P, P.index(dt), pm_weekdays, holidays, next_month_special_holiday):
                    warnings.append(f'{name}：{dt.month}/{dt.day} {slot}は確定当直のため、翌日PM dutyより優先して入れました。')
        total = sum(counts.values())
        if total < int(row['月下限']):
            warnings.append(f'{name}：月下限の目標{int(row["月下限"])}回に対し、{total}回です。')
        group_items = [(col, sum(counts[s] for s in slots), int(row[col])) for col, slots in GROUP_CAP_COLUMNS.items()
                       if clean_text(row.get(col, '')) not in ('', 'nan')]
        for label, value, cap in [('月上限', total, int(row['月上限'])), ('休日上限', hol_count, int(row['休日上限']))] + group_items + [(s + '上限', counts[s], int(row[s + '上限'])) for s in ALL_SHIFT_TYPES if s not in soft_cap_slots(row)]:
            if value > cap:
                warnings.append(f'{name}：確定当直のため{staff_column_specs().get(label, (label,))[0]}を超えています（設定{cap}回／結果{value}回）。')
        warnings += close_fixed_duty_warnings(name, worked, past, future, int(row['宿⇔宿 間隔下限']), int(row.get(DAY_GAP_COLUMN, row['宿⇔宿 間隔下限'])),
                                              lambda dt: P.index(dt) is not None)
    for _, r in result.iterrows():
        d = int(r['_idx'])
        for s in ALL_SHIFT_TYPES:
            names = [n.strip() for n in re.split('[、,]', str(r[s])) if n.strip() not in ('', '-') and '⚠️不足' not in n]
            need = multi.get((d, s), 1)
            if len(names) > need:
                warnings.append(f'{P.label(d)} {s}：設定{need}名に対して{len(names)}名を配置しています。')
    return warnings


def period_generate_shift(P, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None, locked_days=(), annual_balance=None):
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
    result, success, warnings, past, future = _period_generate_core(P, staff_df, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority, locked_days, annual_balance)
    if result is not None:
        result_idx = result.assign(_idx=range(1, len(result) + 1))
        warnings += period_audit_schedule(result_idx, staff_df, P, custom_holidays, multi_slots_dict, past, future, next_month_special_holiday)
        success = not any('⚠️不足' in str(v) for s in ALL_SHIFT_TYPES for v in result[s])
    return result, success, list(dict.fromkeys(warnings)), past, future


def period_build_duty_matrix(df_result, staff_df, P, custom_holidays, next_month_special_holiday=False, comments=None, gap_blocked=None):
    """
    医師ごと・日ごとのマスの内容を作る。
    優先順：決まった当直 ＞ ×（当直NG）＞ 翌日PM duty で宿直を外す日。
    """
    year, month = P.year, P.month
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
                     'ng_count': 0, 'ng_day': 0, 'ng_night': 0, 'label': P.label(d)})   # その日にNG（翌日PM dutyで宿直を外す先生を含む）の先生の数

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
            # 前後の月の確定当直との勤務間隔で入れない日（薄紫のマス）も、実質NGとして数える
            _kind = effective_ng(ng.get(d, 'OK'), info['holiday'])
            _pmr = period_pm_restricted(P, d, pm_days, custom_holidays, next_month_special_holiday)
            # （日直と宿直で空ける日数が違うので、日直側・宿直側を分けて見る。平日は宿直側だけ）
            _gday, _gnight = gap_marks(gap_blocked, name, d, info['holiday'])
            _gap = _gday or _gnight
            if _kind != 'OK' or _pmr or _gap:
                info['ng_count'] += 1
            # 休日は、日直NGと宿直NGを分けても数える（日直NG＝全NG・日NG、宿直NG＝全NG・宿NG・翌日PM duty）
            if info['holiday']:
                info['ng_day'] += int(_kind in ('全NG', '日NG') or _gday)
                info['ng_night'] += int(_kind in ('全NG', '宿NG') or _pmr or _gnight)
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
            # 前月末・翌月初めの確定当直との勤務間隔のため入れない日：マスを薄紫にする（×の表示はそのまま）
            # （今月の確定当直の前後は、色を付けない）
            gday, gnight = gap_marks(gap_blocked, name, d, info['holiday'])
            if info['holiday'] and gday != gnight:
                # 休日で、日直か宿直の片方だけ入れない日：入れない側の半分だけを薄紫にする
                # （上半分＝日直、下半分＝宿直。宿直だけ入れる日の翌日PM dutyは、下半分を斜線にする）
                top = '×' if kind in ('全NG', '日NG') else ''
                bottom = '×' if kind in ('全NG', '宿NG') else ''
                cells.append({'type': 'pmhalf' if (pm and not gnight) else 'half', 'top': top, 'bottom': bottom,
                              'text': '×' if (top or bottom) else '', 'gap_top': gday, 'gap_bottom': gnight})
                continue
            if gnight:
                if info['holiday'] and kind in ('全NG', '宿NG', '日NG'):
                    top = '×' if kind in ('全NG', '日NG') else ''
                    bottom = '×' if kind in ('全NG', '宿NG') else ''
                    cells.append({'type': 'half', 'top': top, 'bottom': bottom, 'text': '×', 'gap': True})
                elif kind in ('全NG', '宿NG'):
                    cells.append({'type': 'ng', 'text': '×', 'pm': False, 'gap': True})
                else:
                    cells.append({'type': 'empty', 'text': '', 'gap': True})
                continue
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
# 先生のグループ（グループごとに、入れる枠を ◯・△・× で決める）
# ==========================================
DOC_GROUP_COLUMN = "グループ"   # 医師条件の列（空欄＝グループなし。今までどおりの動き）
DOC_GROUP_MARKS = ["◯", "△", "×"]   # ◯入れる　△できれば入れない（人が足りないときだけ）　×入れない
DOC_GROUP_MEMO = "メモ"   # グループの設定の、いちばん右のメモ欄（計算には使わない）
DOC_GROUP_AVOID_COLUMN = "_できれば入れない枠"   # 計算用の列（表には出さない）
DOC_GROUP_AVOID_PENALTY = 3000   # △の枠に1回入れるごとの減点（人が足りない枠を空けておく減点より小さくして、足りないときは入れる）
DOC_GROUP_SHEET = "グループの設定"   # 医師条件のExcelの2枚目のシート
# 回数の勾配：枠に余裕があるとき、グループごとに当直の回数を多め・少なめにする（対象外のグループは差をつける計算に入れない）
DOC_GROUP_LEVEL = "回数（余裕があるとき）"   # グループの設定の列
DOC_GROUP_LEVEL_OFF = "対象外"
DOC_GROUP_LEVEL_PCT = {"⬆⬆ 多め": 125, "⬆ やや多め": 112, "ふつう": 100, "⬇ やや少なめ": 88, "⬇⬇ 少なめ": 75}   # 「ふつう」を100としたときの回数の割合
DOC_GROUP_LEVELS = list(DOC_GROUP_LEVEL_PCT) + [DOC_GROUP_LEVEL_OFF]
DOC_GROUP_LEVEL_COLUMN = "_回数の勾配"   # 計算用の列（表には出さない）。割合（％）、対象外は空欄
DOC_GROUP_LEVEL_PENALTY = 1000   # 目安の回数から1回ずれるごとの減点（空き枠・月下限・枠の偏り・△より弱い）
# 最初に並べる6つのグループ（名前, A日直・A宿直・B日直・B宿直・外来日直・外来宿直 の順の◯△×）。年次下位の外来宿直は△
DOC_GROUP_DEFAULT = [("年次上位", "◯◯××◯×"), ("年次中位", "◯◯△◯◯×"), ("年次下位", "△△◯◯◯△"), ("レジ", "××◯◯◯◯"),
                     ("外来日直のみ", "××××◯×"), ("日直のみ", "◯×◯×◯×")]
# 最初のグループの「回数（余裕があるとき）」。ここにないグループは対象外
DOC_GROUP_LEVEL_DEFAULT = {"年次上位": "⬇⬇ 少なめ", "年次中位": "⬇ やや少なめ", "年次下位": "ふつう", "レジ": "⬆⬆ 多め"}


def doc_group_slots():
    """グループの設定の表の列（枠）。"""
    return list(RESULT_COLUMNS[2:])


def doc_group_next_name(used):
    """足した行に付ける名前（グループ6 のように、まだ使っていない番号）。"""
    n = 1
    while f"グループ{n}" in used:
        n += 1
    return f"グループ{n}"


def default_group_rules():
    """最初のグループの設定（年次上位・年次中位・年次下位・レジ・外来日直のみ・日直のみ）。{グループ名: {枠: ◯△×, "メモ": 文字}}"""
    slots = doc_group_slots()
    return {name: {**{s: marks[j] for j, s in enumerate(slots)}, DOC_GROUP_LEVEL: DOC_GROUP_LEVEL_DEFAULT.get(name, DOC_GROUP_LEVEL_OFF),
                   DOC_GROUP_MEMO: ""}
            for name, marks in DOC_GROUP_DEFAULT}


def doc_group_level_of(value):
    """「回数（余裕があるとき）」のマスを、決まった値（⬆⬆ 多め〜⬇⬇ 少なめ・対象外）にする。矢印なしの「多め」なども受け付ける。空欄は対象外。"""
    v = clean_text(value)
    if v in DOC_GROUP_LEVELS:
        return v
    plain = {re.sub(r"[⬆⬇\s]", "", k): k for k in DOC_GROUP_LEVELS}
    return plain.get(re.sub(r"[⬆⬇↑↓\s]", "", v), DOC_GROUP_LEVEL_OFF)


def normalize_group_rules(rules):
    """
    途中保存・Excelから読んだグループの設定を、決まった形にそろえる（並び順はそのまま、全部の枠、◯△×のどれか、メモ）。
    rules が None なら最初の設定。グループが1つもない設定もそのまま（全部消した場合）。名前が空欄・重なる行は使わない。
    """
    if rules is None:
        return default_group_rules()
    out = {}
    for g, row in rules.items():
        g = doc_group_of(g)
        if not g or g in out:
            continue
        row = row if isinstance(row, dict) else {}
        marks = {}
        for s in doc_group_slots():
            v = clean_text(row.get(s, "")).replace("○", "◯").replace("〇", "◯").replace("✕", "×").replace("x", "×").replace("X", "×")
            marks[s] = v if v in DOC_GROUP_MARKS else "◯"
        memo = clean_text(row.get(DOC_GROUP_MEMO, ""))
        # この列ができる前の途中保存・Excel（列そのものがない）は、最初のグループと同じ名前なら最初の値にする
        level = doc_group_level_of(row[DOC_GROUP_LEVEL]) if DOC_GROUP_LEVEL in row else DOC_GROUP_LEVEL_DEFAULT.get(g, DOC_GROUP_LEVEL_OFF)
        out[g] = {**marks, DOC_GROUP_LEVEL: level,
                  DOC_GROUP_MEMO: "" if memo.lower() in ("nan", "none") else memo}
    return out


def doc_group_names():
    """今あるグループの名前（医師条件の「グループ」の列で選べるもの）。"""
    return list((st.session_state.get("group_rules_current") or default_group_rules()).keys())


def doc_group_of(value):
    """医師条件の「グループ」のマス・グループ名から、グループの名前を返す（前後の空白は除く。空欄は ""）。"""
    v = clean_text(value)
    return "" if v.lower() in ("nan", "none", "<na>") else v


def add_group_column(df):
    """医師条件に「グループ」の列がなければ、備考（なければ先生の名前）の右に空欄の列を足す。中身はグループの名前にそろえる。"""
    df = df.copy()
    if DOC_GROUP_COLUMN not in df.columns:
        cols = list(df.columns)
        notes = [c for c in cols if str(c).startswith("備考")]
        pos = cols.index(notes[-1]) + 1 if notes else (cols.index("先生の名前") + 1 if "先生の名前" in cols else 0)
        df.insert(pos, DOC_GROUP_COLUMN, "")
    df[DOC_GROUP_COLUMN] = df[DOC_GROUP_COLUMN].map(lambda v: doc_group_of(v) or None).astype(object)   # 空欄は None（選ぶ欄で空欄に見せる）
    return df


def doc_pref_value(column, value):
    """多め希望のマスの値をそろえる（「病棟」「病棟多め希望」なども「病棟多め」にする）。空欄は ""。分からない値はそのまま返す（入力の確認でエラーにする）。"""
    v = clean_text(value)
    if v.lower() in ("nan", "none", "<na>"):
        return ""
    for option in DOC_PREF_COLUMNS[column]:
        if v and v.startswith(option[:2]):
            return option
    return v


def add_pref_columns(df):
    """医師条件に「多め希望」の3つの列がなければ、翌日PM duty（なければグループ・名前）の右に空欄で足す。中身は選べる値にそろえる。"""
    df = df.copy()
    cols = list(df.columns)
    anchor = next((c for c in ("翌日PM duty", DOC_GROUP_COLUMN, "先生の名前") if c in cols), None)
    pos = cols.index(anchor) + 1 if anchor else len(cols)
    for c in DOC_PREF_COLUMNS:
        if c not in df.columns:
            df.insert(pos, c, "")
        pos = list(df.columns).index(c) + 1
        df[c] = df[c].map(lambda v, c=c: doc_pref_value(c, v) or None).astype(object)   # 空欄は None（選ぶ欄で空欄に見せる）
    return df


def doc_pref_errors(df):
    """多め希望の列に、選べない値が入っていないか確かめる（Excelで手入力した場合など）。"""
    errors = []
    for _, row in df.iterrows():
        for c, options in DOC_PREF_COLUMNS.items():
            if c in df.columns:
                v = doc_pref_value(c, row.get(c, ""))
                if v and v not in options:
                    errors.append(f"{row['先生の名前']}：{c}は「{options[0]}」「{options[1]}」か空欄にしてください（今の値：{v}）。")
    return errors


def doc_pref_terms(model, doctors, days, daily_active_shifts, shifts, staff_df, is_holiday, required_count, min_total, cap_total):
    """
    多め希望の加点（目的関数に足す項）。
    その先生の希望の種類に入れると加点、反対の種類に入れると減点する。点の大きさは、月全体の枠の数の割合で決める。
    多め希望の先生の回数は、全体の枠の数を月上限の比で分けた回数（月下限〜月上限の範囲）に保つよう減点する
    （減点は種類の加点より大きいので、多め希望のために回数が増えたり減ったりしない。種類の割合だけを寄せる）。
    """
    kinds = {"病棟多め": lambda d, s: s not in ANNUAL_OUT_SLOTS, "外来多め": lambda d, s: s in ANNUAL_OUT_SLOTS,
             "宿直多め": lambda d, s: s in NIGHT_SHIFTS, "日直多め": lambda d, s: s not in NIGHT_SHIFTS,
             "平日多め": lambda d, s: not is_holiday(d), "休日多め": lambda d, s: bool(is_holiday(d))}
    need = {k: sum(required_count(d, s) for d in days for s in daily_active_shifts[d] if f(d, s)) for k, f in kinds.items()}
    demand = sum(required_count(d, s) for d in days for s in daily_active_shifts[d])
    cap_sum = sum(cap_total(doc) for doc in doctors) or 1
    terms = []
    # 回数の勾配（グループの「回数（余裕があるとき）」）の対象の先生は、回数を勾配の目安に合わせるので、ここでは回数の目安を使わない
    # （両方の目安があると、多め希望の先生だけ「月上限の比で分けた回数」に引っぱられ、勾配とずれるため）
    def _in_gradient(r):
        return clean_text(r.get(DOC_GROUP_LEVEL_COLUMN, "")).lower() not in ("", "nan", "none")
    gradient_on = DOC_GROUP_LEVEL_COLUMN in staff_df.columns and sum(
        1 for _, r in staff_df.iterrows() if str(r["先生の名前"]) in doctors and _in_gradient(r)) >= 2
    for _, row in staff_df.iterrows():
        doc = str(row["先生の名前"])
        if doc not in doctors:
            continue
        wants = [doc_pref_value(c, row.get(c, "")) for c in DOC_PREF_COLUMNS]
        if any(w in pair for w, pair in zip(wants, DOC_PREF_COLUMNS.values())) and not (gradient_on and _in_gradient(row)):
            # この先生の回数の目安（全体の枠の数を、月上限の比で分ける）。ここからのずれを減点する
            target = min(max(round(demand * cap_total(doc) / cap_sum), min_total(doc)), cap_total(doc))
            mine = [shifts[(d, doc, s)] for d in days for s in daily_active_shifts[d] if (d, doc, s) in shifts]
            dev = model.NewIntVar(0, len(mine) + target, f"pref_dev_{doc}")
            model.Add(dev >= sum(mine) - target)
            model.Add(dev >= target - sum(mine))
            terms.append(-(DOC_PREF_WEIGHT * len(DOC_PREF_COLUMNS) + 1) * dev)
        for c, (a, b) in DOC_PREF_COLUMNS.items():
            want = doc_pref_value(c, row.get(c, ""))
            if want not in (a, b):
                continue
            other = b if want == a else a
            total = need[want] + need[other]
            if not need[want] or not need[other]:
                continue
            plus = max(1, round(DOC_PREF_WEIGHT * need[other] / total))
            minus = max(1, round(DOC_PREF_WEIGHT * need[want] / total))
            for d in days:
                for s in daily_active_shifts[d]:
                    if (d, doc, s) not in shifts:
                        continue
                    if kinds[want](d, s):
                        terms.append(plus * shifts[(d, doc, s)])
                    elif kinds[other](d, s):
                        terms.append(-minus * shifts[(d, doc, s)])
    return terms


# 枠ごとの上限（A宿上限〜外日上限）が空欄の枠は「できれば2回まで」（どうしても足りないときだけ3回目以上も入れる）
SLOT_CAP_SOFT_COLUMN = "_枠ごとの上限が空欄"   # 計算用の列（表には出さない）。空欄だった枠の名前をカンマ区切りで入れる
SLOT_CAP_SOFT_LIMIT = 2
SLOT_CAP_SOFT_PENALTY = 4000   # 2回を超えた1回ごとの減点（空き枠・月下限より弱く、グループの△・休日の偏り・ふつうの希望日より強い）


def mark_soft_slot_caps(df):
    """入力チェックの前に、枠ごとの上限が空欄の枠を計算用の列に書いておく（入力チェックで空欄に2が入るため）。
    すでに列がある表（計算用に作り直した表）は、そのままにする。"""
    if SLOT_CAP_SOFT_COLUMN in df.columns:
        return df
    df = df.copy()
    df[SLOT_CAP_SOFT_COLUMN] = [",".join(s for s in ALL_SHIFT_TYPES
                                         if s + "上限" not in df.columns or clean_text(r[s + "上限"]).lower() in ("", "nan", "none", "<na>"))
                                for _, r in df.iterrows()]
    return df


def soft_cap_slots(row):
    """その先生の「枠ごとの上限が空欄（できれば2回まで）」の枠。"""
    return {x for x in clean_text(row.get(SLOT_CAP_SOFT_COLUMN, "")).split(",") if x in ALL_SHIFT_TYPES}


def soft_cap_terms(model, doctors, days, daily_active_shifts, shifts, soft_caps):
    """枠ごとの上限が空欄の枠：2回を超えた分を減点する（入れないと空き枠が出るときだけ、3回目以上が入る）。"""
    terms = []
    for doc in doctors:
        for s in soft_caps.get(doc, ()):
            worked = [shifts[(d, doc, s)] for d in days if s in daily_active_shifts[d]]
            if len(worked) <= SLOT_CAP_SOFT_LIMIT:
                continue
            excess = model.NewIntVar(0, len(worked), f"soft_cap_{doc}_{s}")
            model.Add(excess >= sum(worked) - SLOT_CAP_SOFT_LIMIT)
            terms.append(-SLOT_CAP_SOFT_PENALTY * excess)
    return terms


def group_level_terms(model, doctors, worked_all, staff_df):
    """
    回数の勾配：グループの「回数（余裕があるとき）」が対象外でない先生の回数を、割合（ふつう＝100）に近づける。
    みんなに共通の「ふつうの人の回数」（0.1回きざみ）を計算で選び、先生ごとの目安＝割合×ふつうの回数（四捨五入して、月下限〜月上限に収める）
    からのずれ（回）を減点する。目安を整数の回数にしておくと、計算が速い。
    枠を埋めること・月下限〜月上限などの決まりが先なので、効くのは枠に余裕があるときだけ。同じ段階の先生は、なるべく同じ回数になる。
    """
    pct, bounds = {}, {}
    if DOC_GROUP_LEVEL_COLUMN in staff_df.columns:
        for _, row in staff_df.iterrows():
            v = clean_text(row.get(DOC_GROUP_LEVEL_COLUMN, ""))
            if v and v.lower() not in ("nan", "none"):
                name = clean_text(row["先生の名前"])
                pct[name] = int(float(v))
                lo = int(float(clean_text(row.get("月下限", 0)) or 0))
                hi = int(float(clean_text(row.get("月上限", 0)) or 0))
                bounds[name] = (min(lo, hi), hi)
    docs = [d for d in doctors if d in pct and d in worked_all]
    if len(docs) < 2:
        return []
    caps = {d: (min(bounds[d][0], min(bounds[d][1], len(worked_all[d]))), min(bounds[d][1], len(worked_all[d]))) for d in docs}
    most = max(hi for _, hi in caps.values())

    def targets(sc):
        """ふつうの人の回数が sc（0.1回単位）のときの、先生ごとの目安の回数（四捨五入して、月下限〜月上限に収める）。"""
        return tuple(min(caps[d][1], max(caps[d][0], int(pct[d] * sc / 1000 + 0.5))) for d in docs)
    # 候補（0.1回きざみ）のうち、目安の組み合わせが同じになるものは1つにまとめる（計算を軽くする）
    choices = list(dict.fromkeys(targets(sc) for sc in range(0, 10 * most * 2 + 1)))
    pick = [model.NewBoolVar(f"group_level_choice_{k}") for k in range(len(choices))]   # どれか1つを選ぶ
    model.AddExactlyOne(pick)
    terms = []
    for j, d in enumerate(docs):
        table = [c[j] for c in choices]
        target = sum(t * z for t, z in zip(table, pick))
        dev = model.NewIntVar(0, len(worked_all[d]) + max(table), f"group_level_dev_{d}")
        model.Add(dev >= sum(worked_all[d]) - target)
        model.Add(dev >= target - sum(worked_all[d]))
        terms.append(-DOC_GROUP_LEVEL_PENALTY * dev)
    return terms


def apply_group_rules(staff_df, rules):
    """
    計算に使う医師条件を作る（画面の表・保存する医師条件は変えない）。
    ・×の枠：その枠の上限を0にする（確定当直は、今までどおり上限より優先）
    ・△の枠：「できれば入れない枠」の列に入れる（計算で、入れるたびに減点する）
    ・もうないグループ（消したグループ）の先生は、グループなしとして扱う
    ・回数（余裕があるとき）：「回数の勾配」の列に割合（％）を入れる（対象外は空欄）
    """
    df = staff_df.copy()
    rules = normalize_group_rules(rules)
    avoid, levels = [], []
    for i, row in df.iterrows():
        g = doc_group_of(row.get(DOC_GROUP_COLUMN, ""))
        marks = {s: m for s, m in rules.get(g, {}).items() if s in doc_group_slots()}
        lv = doc_group_level_of(rules.get(g, {}).get(DOC_GROUP_LEVEL, "")) if g in rules else DOC_GROUP_LEVEL_OFF
        levels.append(DOC_GROUP_LEVEL_PCT.get(lv, ""))
        for s, mark in marks.items():
            if mark == "×" and s + "上限" in df.columns:
                df.at[i, s + "上限"] = 0
                if SLOT_CAP_SOFT_COLUMN in df.columns:   # ×の枠は「できれば」ではなく、必ず0回
                    df.at[i, SLOT_CAP_SOFT_COLUMN] = ",".join(x for x in clean_text(df.at[i, SLOT_CAP_SOFT_COLUMN]).split(",") if x and x != s)
        avoid.append(",".join(s for s, mark in marks.items() if mark == "△"))
    df[DOC_GROUP_AVOID_COLUMN] = avoid
    df[DOC_GROUP_LEVEL_COLUMN] = levels   # 回数の勾配（％。対象外・グループなしは空欄）
    return df


def group_orphan_notice(staff_df, rules):
    """消したグループに入っている先生がいれば、画面で知らせる（その先生はグループなしとして扱う）。"""
    rules = normalize_group_rules(rules)
    gone = {}
    for _, row in staff_df.iterrows():
        g = doc_group_of(row.get(DOC_GROUP_COLUMN, ""))
        if g and g not in rules:
            gone.setdefault(g, []).append(clean_text(row.get("先生の名前", "")))
    for g, names in gone.items():
        st.warning(f"「{g}」はグループの設定にないため、{'・'.join(names)}先生は、グループが空欄の先生と同じように扱います（グループの決まりは使いません）。"
                   "表の「グループ」の列で選び直すか、グループの設定に行を足してください。")


# グループ内で、ほかの先生と違う設定を知らせる（比べるのは休日上限・宿⇔宿・宿⇔日・日⇔日・枠ごとの上限だけ）
GROUP_MISMATCH_MIN_MEMBERS = 3   # この人数以上のグループだけ比べる（2人だと、どちらが正しいか決められない）


def group_mismatches(staff_df, rules):
    """
    グループ内で、ほかの先生と違う設定を探す。[(名前, グループ, 列の表示名, その先生の値, 多い値, 多い値の人数)]
    ・3人以上のグループで、半分より多くの先生が同じ値の項目だけを比べる（値がばらばらの項目は比べない）
    ・×の枠の上限は計算で0になるので比べない
    """
    rules = normalize_group_rules(rules)
    specs = staff_column_specs()
    cols = ["休日上限", "宿⇔宿 間隔下限", DAY_GAP_COLUMN] + [s + "上限" for s in STAFF_SLOT_SHORT]   # 月下限・月上限・宿直上限・日直上限は比べない
    cols = [c for c in cols if c in staff_df.columns]

    def value_of(v):
        if v is None or (isinstance(v, float) and math.isnan(v)) or clean_text(v) == "":
            return None
        try:
            return int(float(v))
        except (TypeError, ValueError):
            return clean_text(v)

    def word(v, c):
        if v is None:
            return "空欄（制限なし）" if specs.get(c, ("", 0, ""))[2] == "count0" else "空欄"
        return f"{v}日間" if specs.get(c, ("", 0, ""))[2] == "days" else f"{v}回"

    found = []
    groups = {}
    for _, row in staff_df.iterrows():
        g, name = doc_group_of(row.get(DOC_GROUP_COLUMN, "")), clean_text(row.get("先生の名前", ""))
        if g and g in rules and name and name.lower() not in ("nan", "none"):
            groups.setdefault(g, []).append(row)
    for g, members in groups.items():
        if len(members) < GROUP_MISMATCH_MIN_MEMBERS:
            continue
        for c in cols:
            if c.endswith("上限") and rules[g].get(c[:-2]) == "×":
                continue
            vals = [value_of(r.get(c)) for r in members]
            common = max(set(vals), key=vals.count)
            n = vals.count(common)
            if n * 2 <= len(vals):
                continue
            label = specs.get(c, (c,))[0]
            for r, v in zip(members, vals):
                if v != common:
                    found.append((clean_text(r.get("先生の名前", "")), g, label, word(v, c), word(common, c), n))
    return found


def group_mismatch_notice(staff_df, rules):
    """グループ内で、ほかの先生と違う設定があれば、表の下に黄色の欄で、1人1行で知らせる（知らせるだけ。当直案はそのまま作れる）。"""
    found = group_mismatches(staff_df, rules)
    if not found:
        return
    order = {clean_text(n): i for i, n in enumerate(staff_df.get("先生の名前", pd.Series(dtype=str)).tolist())}   # 表と同じ順に並べる
    by_doc = {}
    for name, g, label, v, common, n in found:
        by_doc.setdefault((name, g), []).append(f"{html_escape(label)} <b style='color:#8a4b08'>{html_escape(v)}</b>（ほかは{html_escape(common)}）")
    items = "".join(f"<li><b style='color:#8a4b08'>{html_escape(name)}</b><span style='color:#6a5a40'>（{html_escape(g)}）</span>：{'、'.join(parts)}</li>"
                    for (name, g), parts in sorted(by_doc.items(), key=lambda kv: order.get(kv[0][0], 0)))
    st.markdown(
        "<div style='background:#fff8e6;border:1.5px solid #f0b46a;border-radius:10px;padding:10px 14px;margin:6px 0 10px'>"
        "<div style='font-weight:800;color:#8a4b08;margin-bottom:4px'>⚠ グループ内で、ほかの先生と違う設定があります（念のため確認してください）</div>"
        f"<ul style='margin:0 0 4px;padding-left:20px;line-height:1.75'>{items}</ul>"
        "<div style='font-size:0.85rem;color:#6a5a40'>比べる項目：休日上限・宿⇔宿・宿⇔日・日⇔日・枠ごとの上限（3人以上のグループで、半分より多くの先生が同じ値の項目だけ）。"
        "わざと変えている場合は、そのままで大丈夫です（当直案はそのまま作れます）。</div></div>",
        unsafe_allow_html=True)


def group_avoid_terms(doctors, days, daily_active_shifts, shifts, staff_df):
    """△（できれば入れない）の枠に入れるたびの減点（目的関数に足す項）。"""
    terms = []
    for _, row in staff_df.iterrows():
        doc = str(row["先生の名前"])
        if doc not in doctors:
            continue
        avoid = {s for s in clean_text(row.get(DOC_GROUP_AVOID_COLUMN, "")).split(",") if s}
        terms += [-DOC_GROUP_AVOID_PENALTY * shifts[(d, doc, s)] for d in days for s in daily_active_shifts[d]
                  if s in avoid and (d, doc, s) in shifts]
    return terms


def group_rules_xlsx_sheet(wb, rules):
    """医師条件のExcel（wb）に「グループの設定」のシートを足す。"""
    from openpyxl.styles import Font, PatternFill, Alignment
    ws = wb.create_sheet(DOC_GROUP_SHEET)
    rules = normalize_group_rules(rules)
    slots = doc_group_slots()
    ws.append(["グループ名"] + slots + [DOC_GROUP_LEVEL, DOC_GROUP_MEMO])
    for c in ws[1]:
        c.font, c.fill, c.alignment = Font(bold=True), PatternFill("solid", fgColor="DFE6EF"), Alignment(horizontal="center")
    for g, row in rules.items():
        ws.append([g] + [row[s] for s in slots] + [row.get(DOC_GROUP_LEVEL, DOC_GROUP_LEVEL_OFF), row.get(DOC_GROUP_MEMO, "")])
        for i in range(2, len(slots) + 3):
            ws.cell(ws.max_row, i).alignment = Alignment(horizontal="center")
    ws.append([])
    ws.append(["◯：担当する　△：できるだけ担当しない（ほかの先生で埋まらないときだけ担当する）　×：担当しない（ただし、確定当直で決めた分は×でも担当します）"])
    ws.append([f"{DOC_GROUP_LEVEL}：" + "／".join(DOC_GROUP_LEVELS) + "（枠に余裕があるとき、グループごとに回数を多め・少なめにします）"])
    ws.column_dimensions["A"].width = 16
    for i in range(2, len(slots) + 2):
        ws.column_dimensions[chr(64 + i)].width = 10
    ws.column_dimensions[chr(64 + len(slots) + 2)].width = 20
    ws.column_dimensions[chr(64 + len(slots) + 3)].width = 30


def group_rules_from_xlsx(data):
    """医師条件のExcelに「グループの設定」のシートがあれば読む（なければ None）。空の行のところで読むのをやめる（下の説明の行は読まない）。"""
    if not bytes(data[:4]) == b"PK\x03\x04":
        return None
    from openpyxl import load_workbook
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        return None
    if DOC_GROUP_SHEET not in wb.sheetnames:
        return None
    rows = [list(r) for r in wb[DOC_GROUP_SHEET].iter_rows(values_only=True)]
    if not rows:
        return None
    head = [clean_text(v) for v in rows[0]]
    rules = {}
    for r in rows[1:]:
        if not any(clean_text(v) for v in r):
            break
        g = doc_group_of(r[0] if r else "")
        if g:
            rules[g] = {head[j]: r[j] for j in range(1, min(len(head), len(r))) if head[j]}
    return normalize_group_rules(rules)


def set_group_rules_base(rules, ids=None):
    """
    グループの設定の表を、rules で作り直す（読み込んだとき・途中保存から戻したとき・名前を変えた／行を足し引きしたときだけ）。
    行ごとに見えない番号（_id）を付けておき、名前を変えたときに、どのグループの名前が変わったかを分かるようにする。
    """
    rules = normalize_group_rules(rules)
    nxt = st.session_state.get("group_rules_next_id", 0)
    base = []
    for i, (g, row) in enumerate(rules.items()):
        rid = ids[i] if ids and i < len(ids) and ids[i] is not None else None
        if rid is None:
            rid, nxt = nxt, nxt + 1
        base.append({"_id": rid, "name": g, **row})
    st.session_state["group_rules_next_id"] = max([nxt] + [b["_id"] + 1 for b in base])
    st.session_state["group_rules_base"] = base
    st.session_state["group_rules_current"] = rules
    st.session_state["group_rules_gen"] = st.session_state.get("group_rules_gen", 0) + 1


def group_rules_ui(restored_state, year, month, staff_is_template=False):
    """
    「👥 グループの設定（任意）」の表を出し、今の設定を返す。
    ・表の中身は画面の裏に持っておき、途中保存から戻したとき・Excelを読み込んだときだけ作り直す（入力のたびには作り直さない）
    ・グループ名を変えたときは、そのグループの先生の「グループ」の列も新しい名前にする
    ・名前を変えた・行を足した・消したときは、医師条件の表で選べるグループが変わるので、医師条件の表を今の内容のまま出し直す
    ・名前が空欄・ほかのグループと同じになったときは、元の名前（足した行は「グループ6」など）にして知らせる
    """
    saved = (restored_state or {}).get("group_rules")
    src = json.dumps(saved, ensure_ascii=False) if saved is not None else ""
    if st.session_state.get("group_rules_src") != src:
        st.session_state["group_rules_src"] = src
        if saved is not None or "group_rules_base" not in st.session_state:
            set_group_rules_base(saved)
    slots = doc_group_slots()
    base = st.session_state["group_rules_base"]
    view = pd.DataFrame([[b["_id"], b["name"]] + [b[s] for s in slots] + [b.get(DOC_GROUP_LEVEL, DOC_GROUP_LEVEL_OFF), b.get(DOC_GROUP_MEMO, "")] for b in base],
                        columns=["_id", "グループ名"] + slots + [DOC_GROUP_LEVEL, DOC_GROUP_MEMO])
    with st.container(key="group_rules_box", border=True):
        st.markdown("<div style='font-weight:800;color:#24405f;font-size:1.02rem'>👥 グループの設定（任意）</div>", unsafe_allow_html=True)
        st.caption("同じ扱いにしたい先生をグループにまとめ、グループごとに担当する枠を決めます。**グループ名はマスをダブルクリックして変えられます**"
                   "（変えると、そのグループの先生の「グループ」の列も自動で変わります）。先生のグループは、下の表の「グループ」の列で選びます（空欄の先生には、グループの決まりを使いません）。"
                   "グループは表の下の「＋」で足せます（いくつでも）。いらないグループは、行を選んで消せます。")
        for msg in st.session_state.pop("group_rules_notes", []):
            st.warning(msg)
        config = {"グループ名": st.column_config.TextColumn("グループ名", width=120, help="ダブルクリックで名前を変えられます。例：年次上位")}
        config.update({s: st.column_config.SelectboxColumn(s, options=DOC_GROUP_MARKS, required=True, default="◯",
                                                           width=header_fit_width(s, 80),   # 見出しが切れない幅
                                                           help="◯ 担当する　△ できるだけ担当しない（ほかの先生で埋まらないときだけ担当する）　× 担当しない")
                       for s in slots})
        config[DOC_GROUP_LEVEL] = st.column_config.SelectboxColumn(
            DOC_GROUP_LEVEL, options=DOC_GROUP_LEVELS, required=True, default=DOC_GROUP_LEVEL_OFF,
            width=header_fit_width(DOC_GROUP_LEVEL, 150),   # 見出しが切れない幅
            help="枠に余裕があるとき、グループごとに当直の回数を多め・少なめにします（足りない月は全員で埋めるので差はつきません）。"
                 "月下限〜月上限は必ず守ります。同じ段階のグループの先生は、なるべく同じ回数にします。"
                 "対象外のグループの先生には、回数の差をつけません（月下限〜月上限の中で決まります）。")
        config[DOC_GROUP_MEMO] = st.column_config.TextColumn(DOC_GROUP_MEMO, width=240, help="自由に書けるメモです（計算には使いません）。")
        cur = st.data_editor(view, hide_index=True, use_container_width=False, num_rows="dynamic", column_config=config,
                             column_order=["グループ名"] + slots + [DOC_GROUP_LEVEL, DOC_GROUP_MEMO],   # 見えない番号（_id）は出さない
                             key=f"group_rules_editor_{st.session_state.get('group_rules_gen', 0)}")
        st.caption("◯：担当する　△：できるだけ担当しない（ほかの先生で埋まらないときだけ担当する）　×：担当しない（ただし、確定当直で決めた分は×でも担当します）")
        st.caption(f"{DOC_GROUP_LEVEL}：枠に余裕があるとき、グループごとに回数を ⬆⬆ 多め〜⬇⬇ 少なめ にします（ふつうに比べて 多め 約1.25倍・やや多め 約1.1倍・"
                   "やや少なめ 約0.9倍・少なめ 約0.75倍）。月下限〜月上限は必ず守ります。対象外のグループは差をつけません。")

    old_name = {b["_id"]: b["name"] for b in base}
    all_names = set(old_name.values())
    items = []
    for _, r in cur.iterrows():
        rid = r.get("_id")
        rid = int(rid) if rid is not None and not pd.isna(rid) and int(rid) in old_name else None
        items.append((rid, doc_group_of(r.get("グループ名")), r))
    # 名前を変えていないグループの名前は、先に確保しておく（変えた側を元に戻す）
    kept = {name for rid, name, _ in items if rid is not None and name == old_name[rid]}
    rows, ids, used, notes, renames = [], [], set(), [], {}
    for rid, name, r in items:
        old = old_name.get(rid)
        if name and old and name == old:
            pass
        elif not name:
            name = old or doc_group_next_name(used | kept | all_names)
            if old:
                notes.append(f"グループ名を空欄にはできないので、「{old}」のままにしました。")
        elif name in used or name in kept:
            fallback = old if (old and old not in used and old not in kept) else doc_group_next_name(used | kept | all_names)
            notes.append(f"「{name}」はほかのグループと同じ名前なので、" + (f"元の名前「{fallback}」に戻しました。" if fallback == old else f"「{fallback}」にしました。"))
            name = fallback
        used.add(name)
        if old and name != old:
            renames[old] = name
        row = {s: r.get(s) for s in slots}
        row[DOC_GROUP_LEVEL] = r.get(DOC_GROUP_LEVEL)
        row[DOC_GROUP_MEMO] = r.get(DOC_GROUP_MEMO)
        rows.append((name, row))
        ids.append(rid)
    rules = {name: row for name, row in rows}
    if [n for n, _ in rows] != [b["name"] for b in base] or notes:
        # 名前を変えた・行を足した・消した：グループの表と、医師条件の表（今の内容のまま）を出し直す
        set_group_rules_base(rules, ids)
        st.session_state["group_rules_notes"] = notes
        csv_text = st.session_state.get("staff_current_csv")
        if csv_text and not staff_is_template:
            df = parse_staff_csv(csv_text.encode("utf-8"))
            if DOC_GROUP_COLUMN in df.columns and renames:
                df[DOC_GROUP_COLUMN] = df[DOC_GROUP_COLUMN].map(lambda v: renames.get(doc_group_of(v), doc_group_of(v)) or None)
            _staff_apply_new_table(df, year, month, None)
        st.rerun()
    st.session_state["group_rules_current"] = normalize_group_rules(rules)
    return st.session_state["group_rules_current"]


# ==========================================
# 年間の当直回数（掲示用Excelから集計する。Excelそのものは保存しない）
# ==========================================
ANNUAL_KINDS = ["total", "night", "day", "hol", "ward", "out", "weekday"]   # 当直合計・宿直・日直・休日・病棟（A・B）・外来・平日（回数のリストはこの順）
ANNUAL_WARD_SLOTS = [s for s in ALL_SHIFT_TYPES if s[:1] in ("A", "B")]   # 病棟（A・B）
ANNUAL_OUT_SLOTS = [s for s in ALL_SHIFT_TYPES if s.startswith("外来")]   # 外来
ANNUAL_STRENGTHS = ["しない", "弱め", "ふつう", "強め"]
ANNUAL_STRENGTH_WEIGHT = {"しない": 0, "弱め": 100, "ふつう": 300, "強め": 1000}   # 当直1回分のずれに付ける減点
# 当直案に反映して「そろえる」回数。組（休日と平日など）は足すと当直合計になるので、片方をそろえればもう片方もそろう
ANNUAL_BALANCE_ITEMS = [("total", "当直合計", "すべての当直"),
                        ("hol", "休日・平日", "土日祝・休日にした日の当直"),
                        ("night", "宿直・日直", "A宿直・B宿直・外来宿直（夜に寝られない）"),
                        ("out", "外来・病棟", "外来日直・外来宿直")]
ANNUAL_KEEP_MONTHS = 24   # 途中保存に残す月の数（これより古い月は、保存のときに消す）


def annual_split_names(value):
    """マスの名前を1人ずつに分ける（「未決定」「未定」や不足の印は除く）。"""
    names = []
    for n in re.split(r"[、,，・/／\n\s]+", clean_text(value)):
        n = n.strip()
        if n and n not in ("-", "－", "―", "ー", "未定", "未決定", "⚠️未決定", "nan", "None") and "不足" not in n:
            names.append(n)
    return names


def annual_read_posting(data, file_label):
    """
    掲示用Excelを読んで、日ごとの当直を返す。手で直したあとのファイルでもよい。
    戻り値：{"label", "created", "roster", "days": {日付: {"hol": 休日か, "duties": [(先生, 枠)]}}}
    読めないときは ValueError（画面に出す文）。
    """
    from openpyxl import load_workbook
    try:
        wb = load_workbook(io.BytesIO(data), data_only=True)
    except Exception:
        raise ValueError(f"「{file_label}」はExcelファイルとして読めませんでした。")
    ws = wb["掲示用"] if "掲示用" in wb.sheetnames else wb.worksheets[0]
    # 見出しの行（A日直・A宿直…）を探す
    header_row, slot_cols = None, {}
    for r in range(1, min(ws.max_row, 15) + 1):
        found = {c: clean_text(ws.cell(r, c).value) for c in range(1, ws.max_column + 1)
                 if clean_text(ws.cell(r, c).value) in ALL_SHIFT_TYPES}
        if len(found) >= 2:
            header_row, slot_cols = r, found
            break
    if header_row is None:
        raise ValueError(f"「{file_label}」に「A日直」「A宿直」などの見出しが見つかりません。アプリでダウンロードした掲示用Excelを読み込んでください。")
    # 見えないシート（作成日時・休日・先生の一覧）
    created, hol_map, roster = None, {}, []
    if POSTING_DATA_SHEET in wb.sheetnames:
        ds = wb[POSTING_DATA_SHEET]
        created = clean_text(ds["B2"].value) or None
        for r in range(5, ds.max_row + 1):
            d_txt, h_txt, n_txt = clean_text(ds.cell(r, 1).value), clean_text(ds.cell(r, 2).value), clean_text(ds.cell(r, 4).value)
            if d_txt:
                try:
                    hol_map[datetime.date.fromisoformat(d_txt[:10])] = (h_txt == "休日")
                except ValueError:
                    pass
            if n_txt:
                roster.append(n_txt)
    if not created:
        c = wb.properties.created
        created = c.strftime("%Y/%m/%d %H:%M:%S") if c else "0000"
    days = {}
    for r in range(header_row + 1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, datetime.datetime):
            dt = v.date()
        elif isinstance(v, datetime.date):
            dt = v
        else:
            txt = clean_text(v)
            m = re.fullmatch(r"(\d{4})[/.-](\d{1,2})[/.-](\d{1,2})", txt)
            if not m:
                continue
            try:
                dt = datetime.date(*map(int, m.groups()))
            except ValueError:
                continue
        hol = hol_map.get(dt, dt.weekday() >= 5 or bool(jpholiday.is_holiday(dt)))
        duties = [(n, slot) for c, slot in slot_cols.items() for n in annual_split_names(ws.cell(r, c).value)]
        days[dt] = {"hol": hol, "duties": duties}
    if not days:
        raise ValueError(f"「{file_label}」に日付の行が見つかりません。")
    names_in = [n for d in days.values() for n, _ in d["duties"]]
    roster = list(dict.fromkeys(roster + names_in))
    return {"label": file_label, "created": created, "roster": roster, "days": days}


# 集計だけのファイル（当直表が見つからない月の分）の列：見出し → 回数のリストの番号
ANNUAL_SUMMARY_COLUMNS = {"合計": 0, "当直合計": 0, "宿直": 1, "日直": 2, "休日": 3, "病棟（A・B）": 4, "病棟": 4, "外来": 5, "平日": 6}
ANNUAL_SUMMARY_PAIRS = [(4, 5), (6, 3), (1, 2)]   # 足すと合計になる組（病棟・外来、平日・休日、宿直・日直）


def annual_month_key(value):
    """「2026/4」「2026年4月」「2026-04」や日付のマスを "2026-04" にする。読めなければ None。"""
    if isinstance(value, (datetime.datetime, datetime.date)):
        return f"{value.year}-{value.month:02d}"
    m = re.search(r"(\d{4})\s*[年/.-]\s*(\d{1,2})", clean_text(value).translate(str.maketrans("０１２３４５６７８９／－", "0123456789/-")))
    if not m or not 1 <= int(m.group(2)) <= 12:
        return None
    return f"{int(m.group(1))}-{int(m.group(2)):02d}"


def annual_month_span(start_key, end_key):
    """"2026-04"〜"2026-09" の月の一覧。"""
    y, m = int(start_key[:4]), int(start_key[5:])
    ey, em = int(end_key[:4]), int(end_key[5:])
    out = []
    while (y, m) <= (ey, em) and len(out) < 60:
        out.append(f"{y}-{m:02d}")
        y, m = (y + 1, 1) if m == 12 else (y, m + 1)
    return out


def annual_fill_pairs(c):
    """分かっている回数から、分かっていない回数を補う（例：合計と休日が分かれば、平日は 合計−休日）。"""
    c = list(c)
    for _ in range(2):
        if c[0] is None:
            for a, b in ANNUAL_SUMMARY_PAIRS:
                if c[a] is not None and c[b] is not None:
                    c[0] = c[a] + c[b]
                    break
        if c[0] is not None:
            for a, b in ANNUAL_SUMMARY_PAIRS:
                if c[a] is None and c[b] is not None:
                    c[a] = c[0] - c[b]
                elif c[b] is None and c[a] is not None:
                    c[b] = c[0] - c[a]
    return c


def annual_read_summary(data, file_label):
    """
    集計だけのファイル（Excel・CSV）を読む。読めないときは ValueError（画面に出す文）。
    ・いちばん上の「集計の期間（はじめ）」「集計の期間（おわり）」の右のマスから期間を読む
      （アプリでダウンロードした「年間の当直回数」のExcelは、題名の「2026年4月〜2026年10月」から読む）
    ・「先生」の見出しの行から下を、先生ごとの回数として読む。「1か月あたり」の列は読まない
    ・空欄の回数は None（分からない）。在籍月数が空欄なら、あとで期間の月数にする
    戻り値：{"label", "start", "end"（"YYYY-MM" または None）, "docs": {先生: {"m": 在籍月数 or None, "c": 回数のリスト}}}
    """
    rows = []
    if bytes(data[:4]) == b"PK\x03\x04":
        from openpyxl import load_workbook
        try:
            wb = load_workbook(io.BytesIO(data), data_only=True)
        except Exception:
            raise ValueError(f"「{file_label}」はExcelファイルとして読めませんでした。")
        rows = [list(r) for r in wb.worksheets[0].iter_rows(values_only=True)]
    else:
        import csv
        for enc in ("utf-8-sig", "cp932"):
            try:
                rows = list(csv.reader(io.StringIO(bytes(data).decode(enc))))
                break
            except UnicodeDecodeError:
                continue
    start = end = None
    head_i = None
    for i, r in enumerate(rows[:15]):
        cells = [clean_text(v) for v in r]
        for j, v in enumerate(cells):
            if "集計の期間" in v and "はじめ" in v and j + 1 < len(r):
                start = annual_month_key(r[j + 1]) or start
            if "集計の期間" in v and "おわり" in v and j + 1 < len(r):
                end = annual_month_key(r[j + 1]) or end
            t = re.search(r"(\d{4})年(\d{1,2})月〜(\d{4})年(\d{1,2})月", v)
            if t and start is None:
                start, end = f"{t.group(1)}-{int(t.group(2)):02d}", f"{t.group(3)}-{int(t.group(4)):02d}"
        if head_i is None and "先生" in cells:
            head_i = i
    if head_i is None:
        raise ValueError(f"「{file_label}」に「先生」の見出しが見つかりません。集計のひな形の形で書いてください。")
    head = [clean_text(v) for v in rows[head_i]]
    name_col = head.index("先生")
    month_col = head.index("在籍月数") if "在籍月数" in head else None
    cols = {j: ANNUAL_SUMMARY_COLUMNS[h] for j, h in enumerate(head) if h in ANNUAL_SUMMARY_COLUMNS and "あたり" not in h}
    if not cols:
        raise ValueError(f"「{file_label}」に回数の列（合計・休日・宿直など）が見つかりません。")

    def num(v):
        t = clean_text(v)
        if t in ("", "―", "-", "nan", "None"):
            return None
        try:
            return float(t)
        except ValueError:
            return None
    docs = {}
    for r in rows[head_i + 1:]:
        name = clean_text(r[name_col]) if name_col < len(r) else ""
        if not name or name in ("平均", "合計"):
            continue
        c = [None] * len(ANNUAL_KINDS)
        for j, k in cols.items():
            v = num(r[j]) if j < len(r) else None
            if v is not None:
                c[k] = int(round(v))
        mm = num(r[month_col]) if (month_col is not None and month_col < len(r)) else None
        docs[name] = {"m": mm, "c": annual_fill_pairs(c)}
    if not docs:
        raise ValueError(f"「{file_label}」に先生の行がありません。")
    if start and end and start > end:
        start, end = end, start
    return {"label": file_label, "start": start, "end": end, "docs": docs}


def annual_summary_template_xlsx():
    """集計だけのファイルのひな形（期間の行・先生ごとの回数の見出し。期間のマスは「文字」の形にして、日付に変わらないようにする）。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill, Alignment
    wb = Workbook()
    ws = wb.active
    ws.title = "集計"
    head_fill, in_fill = PatternFill("solid", fgColor="DFE6EF"), PatternFill("solid", fgColor="FFFBE6")
    ws["A1"], ws["C1"] = "集計の期間（はじめ）", "集計の期間（おわり）"
    ws["E1"] = "← 年/月 で書きます（例：2026/4）"
    for ref in ("A1", "C1"):
        ws[ref].font, ws[ref].fill = Font(bold=True), head_fill
    for ref in ("B1", "D1"):
        ws[ref].number_format, ws[ref].fill = "@", in_fill
    head = ["先生", "在籍月数", "病棟（A・B）", "外来", "宿直", "日直", "平日", "休日", "合計"]
    for j, h in enumerate(head, 1):
        c = ws.cell(2, j, h)
        c.font, c.fill, c.alignment = Font(bold=True), head_fill, Alignment(horizontal="center")
    for r in range(3, 63):
        for j in range(1, len(head) + 1):
            ws.cell(r, j).fill = in_fill
    ws["K2"] = "・分かる欄だけで大丈夫です（空欄は「分からない」として扱います）"
    ws["K3"] = "・在籍月数が空欄なら、期間の月数として数えます"
    ws["K4"] = "・期間は、アプリの「集計の区切り」（4月〜3月など）の中に収めてください"
    for col, w in zip("ABCDEFGHI", [20, 10, 20, 8, 8, 8, 8, 8, 8]):
        ws.column_dimensions[col].width = w
    ws.freeze_panes = "B3"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def annual_counts_of(day_items):
    """日ごとの当直から、先生ごとの回数 [当直合計, 宿直, 日直, 休日, 病棟, 外来, 平日] を数える。"""
    counts = {}
    for info in day_items:
        for name, slot in info["duties"]:
            c = counts.setdefault(name, [0] * len(ANNUAL_KINDS))
            c[0] += 1
            c[1] += int(slot in NIGHT_SHIFTS)
            c[2] += int(slot in DAY_SHIFTS)
            c[3] += int(bool(info["hol"]))
            c[4] += int(slot in ANNUAL_WARD_SLOTS)
            c[5] += int(slot in ANNUAL_OUT_SLOTS)
            c[6] += int(not info["hol"])
    return counts


def annual_merge(saved_pieces, files):
    """
    途中保存から戻した分（saved_pieces：月ごとの回数）と、読み込んだExcel（files：日ごとの当直）をまとめる。
    ・同じ日が重なったときは、あとから作ったほう（作成日時が新しいほう）を使う
    ・戻り値：(月ごとの回数のリスト, 注意の文のリスト)。読み込んだ分には "loaded": True を付ける
    """
    notes = []
    # 読み込んだExcel同士：日ごとに新しいほうを使う
    by_day = {}
    for f in sorted(files, key=lambda f: f["created"]):
        for dt, info in f["days"].items():
            by_day[dt] = (f, info)
    # 途中保存の分と重なる日
    kept = []
    for p in saved_pieces:
        s, e = datetime.date.fromisoformat(p["start"]), datetime.date.fromisoformat(p["end"])
        overlap = [dt for dt in by_day if s <= dt <= e]
        if not overlap:
            kept.append(p)
            continue
        newer = [dt for dt in overlap if by_day[dt][0]["created"] >= p["created"]]
        if len(newer) == len(overlap) and len(overlap) >= int(p.get("ndays", 0)):
            continue   # 読み込んだExcelのほうが新しく、同じ日をすべて含む → 保存の分は使わない
        if newer:
            notes.append(f"{p['month'][:4]}年{int(p['month'][5:])}月：途中保存の分と、読み込んだExcelの日付が一部だけ重なっています。重なった日は途中保存の分を使いました。")
        for dt in overlap:
            by_day.pop(dt, None)   # 保存の分のほうが新しい（または一部だけ重なる）→ 保存の分を使う
        kept.append(p)
    # 読み込んだ分を、ファイルごと・月ごとにまとめる
    groups = {}
    for dt, (f, info) in by_day.items():
        groups.setdefault((f["label"], f["created"], f"{dt.year}-{dt.month:02d}"), (f, []))[1].append((dt, info))
    for (label, created, month), (f, items) in groups.items():
        items.sort(key=lambda x: x[0])
        kept.append({"month": month, "start": items[0][0].isoformat(), "end": items[-1][0].isoformat(),
                     "ndays": len(items), "created": created, "roster": list(f["roster"]),
                     "counts": annual_counts_of(info for _, info in items), "loaded": True})
    kept.sort(key=lambda p: (p["month"], p["start"]))
    return kept, notes


def annual_fy_months(year, month, fy_start):
    """集計の区切り（fy_start 月はじまり）で、year年month月を含む1年分の月（"YYYY-MM"）。"""
    y0 = year if month >= fy_start else year - 1
    out = []
    for i in range(12):
        m = (fy_start - 1 + i) % 12 + 1
        y = y0 + (fy_start - 1 + i) // 12
        out.append(f"{y}-{m:02d}")
    return out


def annual_summary_months(summaries, months):
    """集計だけのファイルのうち、期間がすべて months（集計の区切り）に入っているものと、その月の一覧。"""
    used, covered = [], set()
    for sm in summaries or []:
        if not (sm.get("start") and sm.get("end")):
            continue
        span = annual_month_span(sm["start"], sm["end"])
        if span and all(m in months for m in span):
            used.append((sm, span))
            covered |= set(span)
    return used, covered


def annual_totals(pieces, months, exclude=None, summaries=()):
    """
    months の月の分を、先生ごとに合計する。
    集計だけのファイル（summaries）の期間に入る月は、そのファイルの回数を使い、掲示用Excelの分は使わない（二重に数えないため）。
    戻り値：{先生: {"m": 在籍月数（日数の割合で数える）, "c": [当直合計, 宿直, 日直, 休日, 病棟, 外来, 平日]}}
    以前の途中保存（回数が4つだけ）の月がある先生は、病棟・外来が分からないので None にする。
    exclude=(開始日, 終了日) を渡すと、その期間と重なる分は数えない（これから作る月の分）。
    """
    out = {}
    used, covered = annual_summary_months(summaries, months)
    for sm, span in used:
        s = datetime.date(int(span[0][:4]), int(span[0][5:]), 1)
        e = datetime.date(int(span[-1][:4]), int(span[-1][5:]), _REAL_MONTHRANGE(int(span[-1][:4]), int(span[-1][5:]))[1])
        if exclude and not (e < exclude[0] or s > exclude[1]):
            continue
        for name, d in sm.get("docs", {}).items():
            o = out.setdefault(name, {"m": 0.0, "c": [0] * len(ANNUAL_KINDS)})
            o["m"] += float(d.get("m") if d.get("m") is not None else len(span))
            for i, x in enumerate(list(d.get("c", []))[:len(ANNUAL_KINDS)]):
                o["c"][i] = None if (x is None or o["c"][i] is None) else o["c"][i] + int(x)
    for p in pieces:
        if p["month"] not in months or p["month"] in covered:
            continue
        s, e = datetime.date.fromisoformat(p["start"]), datetime.date.fromisoformat(p["end"])
        if exclude and not (e < exclude[0] or s > exclude[1]):
            continue
        y, m = int(p["month"][:4]), int(p["month"][5:])
        frac = int(p.get("ndays", 0)) / _REAL_MONTHRANGE(y, m)[1]
        for name in list(dict.fromkeys(list(p.get("roster", [])) + list(p.get("counts", {}).keys()))):
            o = out.setdefault(name, {"m": 0.0, "c": [0] * len(ANNUAL_KINDS)})
            o["m"] += frac
            v = list(p.get("counts", {}).get(name, [0] * len(ANNUAL_KINDS)))
            if len(v) < len(ANNUAL_KINDS):   # 以前の形（当直合計・宿直・日直・休日の4つ）
                v = v[:4] + [None, None, int(v[0]) - int(v[3])]
            for i, x in enumerate(v):
                o["c"][i] = None if (x is None or o["c"][i] is None) else o["c"][i] + int(x)
    return out


def annual_result_file(df_result, year, month, roster, created):
    """作成した当直案を、読み込んだExcelと同じ形（日ごとの当直）にする（年間の回数に今月分を足すため）。"""
    days = {}
    for _, row in df_result.iterrows():
        dt = parse_shift_date(row.get("日付", ""), year, month)
        if dt is None:
            continue
        duties = [(n, s) for s in ALL_SHIFT_TYPES for n in annual_split_names(row.get(s, ""))]
        days[dt] = {"hol": clean_text(row.get("平日/休日", "")) == "休日", "duties": duties}
    return {"label": "作成した当直案", "created": created or "9999", "roster": list(roster), "days": days}


def annual_balance_info(pieces, months, target_start, target_end, frac, weights, doctors, summaries=()):
    """
    当直案の計算に渡す、年間の回数（これから作る期間を除く）。weights がすべて0なら None。
    frac：これから作る期間が何か月分か（通常は1、期間指定は 日数÷その月の日数）
    """
    if not (pieces or summaries) or not any(weights.values()):
        return None
    tot = annual_totals(pieces, months, exclude=(target_start, target_end), summaries=summaries)
    if not tot:
        return None
    docs = {d: {"m": round(tot.get(d, {}).get("m", 0.0), 4), "c": list(tot.get(d, {}).get("c", [0] * len(ANNUAL_KINDS)))} for d in doctors}
    weights = dict(weights)
    # 回数が分からない先生がいる種類は、そろえない（以前の途中保存や、空欄のある集計だけのファイル）
    for kind, i in {"total": 0, "night": 1, "hol": 3, "out": 5}.items():
        if any(docs[d]["c"][i] is None for d in doctors):
            weights[kind] = 0
    return {"weights": weights, "frac": round(frac, 4), "docs": docs}


def add_annual_balance_terms(model, doctors, days, daily_active_shifts, is_holiday, required_count, worked_all, holiday_worked, shifts, info):
    """
    年間の回数を、1か月あたりでそろえるための減点（目的関数に足す項）を返す。
    先生ごとに「これまでの回数＋今回の回数」と「在籍月数×全体の1か月あたりの回数」のずれ（0.1回単位）を減点する。
    月上限・NG日・間隔などの決まりは、ほかの制約のとおり守る（その範囲の中でそろえる）。
    """
    if not info:
        return []
    terms = []
    frac = float(info.get("frac", 1.0)) or 1.0
    docs = info.get("docs", {})
    slot_need = {"total": 0, "hol": 0, "night": 0, "out": 0}
    for d in days:
        for s in daily_active_shifts[d]:
            need = required_count(d, s)
            slot_need["total"] += need
            slot_need["night"] += need if s in NIGHT_SHIFTS else 0
            slot_need["hol"] += need if is_holiday(d) else 0
            slot_need["out"] += need if s in ANNUAL_OUT_SLOTS else 0
    idx = {"total": 0, "night": 1, "hol": 3, "out": 5}
    months_sum = sum(float(docs.get(doc, {}).get("m", 0.0)) + frac for doc in doctors)
    for kind, weight in info.get("weights", {}).items():
        weight = int(weight)
        if weight <= 0 or kind not in idx or months_sum <= 0:
            continue
        past = {doc: int((docs.get(doc, {}).get("c") or [0] * len(ANNUAL_KINDS))[idx[kind]] or 0) for doc in doctors}
        rate = (sum(past.values()) + slot_need[kind]) / months_sum
        for doc in doctors:
            if kind == "total":
                now = worked_all[doc]
            elif kind == "hol":
                now = holiday_worked[doc]
            elif kind == "out":
                now = [shifts[(d, doc, s)] for d in days for s in daily_active_shifts[d] if s in ANNUAL_OUT_SLOTS]
            else:
                now = [shifts[(d, doc, s)] for d in days for s in daily_active_shifts[d] if s in NIGHT_SHIFTS]
            target10 = int(round(10 * rate * (float(docs.get(doc, {}).get("m", 0.0)) + frac)))
            dev = model.NewIntVar(0, 10 * (past[doc] + len(now)) + target10 + 10, f"annual_dev_{kind}_{doc}")
            model.Add(dev >= 10 * past[doc] + 10 * sum(now) - target10)
            model.Add(dev >= target10 - 10 * past[doc] - 10 * sum(now))
            terms.append(dev * -max(1, weight // 10))
    return terms


def annual_month_label(ym):
    return f"{int(ym[5:])}月"


def annual_fmt_months(m):
    return f"{m:.0f}か月" if abs(m - round(m)) < 0.05 else f"{m:.1f}か月"


def annual_rate_text(r):
    """医師条件の表の「📊 年間（1か月あたり）」のマスの文字（計・外来・日直・休日の1か月あたり）。"""
    m = r["m"]

    def f(v):
        return "―" if (v is None or m <= 0) else f"{v / m:.1f}"
    c = r["c"]
    return f"計{f(c[0])}｜外{f(c[5])}｜日{f(c[2])}｜休{f(c[3])}"


def annual_publish_rates(rows):
    """
    年間の集計（1か月あたり）を、医師条件の表で使えるように残す。
    医師条件の表は年間の項目より上にあるので、この回の表で使った値と違えば、もう一度描き直して表にも反映する（同じになれば描き直さない）。
    """
    rates = {r["name"]: annual_rate_text(r) for r in rows if r.get("here")}
    st.session_state["annual_doc_rates"] = rates
    if rates != st.session_state.get("annual_doc_rates_used", {}):
        st.session_state["annual_doc_rates_used"] = rates
        st.rerun()


def annual_rows(tot, staff_names):
    """
    集計の表の行を作る。今月の先生を先に、いない先生（年間の分だけある先生）を後に並べる。
    平均との差は、今月の先生のうち在籍月数のある先生の「1か月あたりの当直合計」の平均とくらべる。
    """
    now = [n for n in staff_names if n in tot]
    rates = [tot[n]["c"][0] / tot[n]["m"] for n in now if tot[n]["m"] > 0 and tot[n]["c"][0] is not None]
    avg = sum(rates) / len(rates) if rates else 0.0
    rows = []
    for n in now + [n for n in tot if n not in staff_names]:
        o = tot[n]
        rate = o["c"][0] / o["m"] if (o["m"] > 0 and o["c"][0] is not None) else None
        diff = (rate - avg) if (rate is not None and n in staff_names and rates) else None
        rows.append({"name": n, "here": n in staff_names, "m": o["m"], "c": o["c"], "diff": diff})
    return rows, avg


ANNUAL_TABLE_CSS = (
    "<style>.an-t{width:100%;border-collapse:collapse;font-size:.85rem}"
    ".an-t th,.an-t td{border-bottom:1px solid rgba(128,144,160,.3);padding:3px 8px;text-align:right;white-space:nowrap}"
    ".an-t th{background:#eef1f5;font-weight:600}.an-t td:first-child,.an-t th:first-child{text-align:left}"
    ".an-box{max-height:420px;overflow:auto;border:1px solid rgba(128,144,160,.35);border-radius:8px;margin:2px 0 6px}"
    ".an-box thead{position:sticky;top:0;z-index:2}.an-box .an-t{border-collapse:separate;border-spacing:0}"
    ".an-t .pm{display:inline;font-size:.74rem;opacity:.6;margin-left:3px}.an-t .lo{color:#1f6fbf;font-weight:700}.an-t .hi{color:#c0392b;font-weight:700}"
    ".an-t tr.gone td{opacity:.45}.an-t .l{text-align:left}"
    ".an-m{display:flex;gap:6px;flex-wrap:wrap;margin:4px 0}.an-m div{border-radius:8px;padding:4px 0;width:66px;text-align:center;font-size:.8rem;border:1px solid rgba(128,144,160,.35)}"
    ".an-m .ok{background:#e9f6ee;border-color:#9fd3b0;color:#1d6b3a;font-weight:700}.an-m .sv{background:#eaf0f7;border-color:#b7c7dc;color:#24405f;font-weight:700}"
    ".an-m .no{opacity:.45}.an-m .sm{background:#f1ecfb;border-color:#c9b9ee;color:#4f3a8c;font-weight:700}.an-t th.grp{text-align:center;border-left:1px solid rgba(128,144,160,.3)}.an-t .gl{border-left:1px solid rgba(128,144,160,.3)}"
    ".an-t th.grp.l{text-align:left}</style>")


def annual_cell(v, m, bold=False):
    if v is None:   # 以前の途中保存で、内訳が分からない
        return "―"
    pm = f"({v / m:.1f})" if m > 0 else "(―)"   # （ ）内は1か月あたり。1人1段に収めるため横に並べる
    return (f"<b>{v}回</b><span class='pm'><b>{pm}</b></span>" if bold else f"{v}回<span class='pm'>{pm}</span>")


def annual_diff_html(diff):
    if diff is None:
        return "<td class='l'>―</td>"
    return f"<td class='l'>{diff:+.1f}回/月</td>"   # 差の数字だけを出す（多い・少ないの評価や色分けはしない）


# 集計の表の列（回数のリストの番号, 見出し）。区切りごとに足すと当直合計になる
ANNUAL_TABLE_GROUPS = [("枠の種類", [(4, "病棟（A・B）"), (5, "外来")]),
                       ("宿直・日直", [(1, "宿直"), (2, "日直")]),
                       ("平日・休日", [(6, "平日"), (3, "休日")])]


def annual_head_html(first_cols, suffix=""):
    """集計の表の見出し（2段）。first_cols は左の列の見出し、suffix は「（年間）」など。"""
    top = "".join(f"<th rowspan=2>{c}</th>" for c in first_cols)
    top += "".join(f"<th class='grp' colspan={len(cols)}>{g}{suffix}</th>" for g, cols in ANNUAL_TABLE_GROUPS)
    top += (f"<th class='grp' rowspan=2>合計{suffix}</th>"
            "<th class='grp l' rowspan=2>平均との差<br><span style='font-weight:400;font-size:.72rem'>（1か月あたりの合計）</span></th>")
    sub = "".join(f"<th class='{'gl' if i == 0 else ''}'>{label}</th>" for _, cols in ANNUAL_TABLE_GROUPS for i, (_, label) in enumerate(cols))
    return f"<tr>{top}</tr><tr>{sub}</tr>"


def annual_body_cells(r):
    """1人分の回数のマス（区切りごとの列・合計・平均との差）。"""
    cells = "".join(f"<td class='{'gl' if i == 0 else ''}'>{annual_cell(r['c'][k], r['m'])}</td>"
                    for _, cols in ANNUAL_TABLE_GROUPS for i, (k, _) in enumerate(cols))
    return cells + f"<td class='gl'>{annual_cell(r['c'][0], r['m'], True)}</td>" + annual_diff_html(r["diff"])


def annual_table_html(rows):
    body = ""
    for r in rows:
        name = html_escape(r["name"]) + ("" if r["here"] else "（今月はいない先生）")
        body += f"<tr class='{'' if r['here'] else 'gone'}'><td>{name}</td><td>{annual_fmt_months(r['m'])}</td>{annual_body_cells(r)}</tr>"
    return ANNUAL_TABLE_CSS + f"<div class='an-box'><table class='an-t'><thead>{annual_head_html(['先生', '在籍月数'])}</thead><tbody>{body}</tbody></table></div>"


def annual_totals_xlsx(rows, title):
    """集計の表をExcelにする（合計と1か月あたり）。"""
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment
    from openpyxl.utils import get_column_letter
    wb = Workbook()
    ws = wb.active
    ws.title = "年間の当直回数"
    ws["A1"] = title
    ws["A1"].font = Font(size=14, bold=True)
    items = [(k, label) for _, cols in ANNUAL_TABLE_GROUPS for k, label in cols] + [(0, "合計")]
    head = ["先生", "在籍月数"] + [h for _, label in items for h in (label, f"{label}（1か月あたり）")] + ["平均との差（1か月あたりの合計）", "備考"]
    ws.append([])
    ws.append(head)
    for c in ws[3]:
        c.font = Font(bold=True)
        c.alignment = Alignment(wrap_text=True, vertical="center")
    for r in rows:
        m = r["m"]
        vals = []
        for k, _ in items:
            v = r["c"][k]
            vals += [v, round(v / m, 1) if (v is not None and m > 0) else None]
        ws.append([r["name"], round(m, 1)] + vals + [None if r["diff"] is None else round(r["diff"], 1), "" if r["here"] else "今月はいない先生"])
    for i in range(1, len(head) + 1):
        ws.column_dimensions[get_column_letter(i)].width = 16 if i == 1 else 11
    ws.row_dimensions[3].height = 32
    ws.freeze_panes = "B4"
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def annual_clear_saved():
    """「🗑 保存した集計を消す」：途中保存から戻した分を消す（次に途中保存すると、保存の中からも消える）。"""
    st.session_state["annual_saved"] = []
    st.session_state["annual_summaries_saved"] = []


def annual_section(year, month, target_start, target_end, frac, restored_state, staff_names, now_word="今月"):
    """
    「📊 年間の当直回数」の項目を出す（集計して見るだけ。当直案の計算には使わないので、いつも None を返す）。
    ・掲示用Excelは画面の裏で読むだけで保存しない。途中保存には、月ごとの回数だけを入れる（annual_snapshot）
    ・target_start〜target_end はこれから作る期間（反映の計算では、この期間と重なる分は数えない）
    """
    # 途中保存から戻したとき：回数と設定を戻す（同じ保存を何度も戻さない）
    saved = (restored_state or {}).get("annual") or {}
    src = hashlib.sha256(json.dumps(saved, ensure_ascii=False, sort_keys=True).encode("utf-8")).hexdigest() if saved else ""
    if st.session_state.get("annual_src") != src:
        st.session_state["annual_src"] = src
        st.session_state["annual_saved"] = list(saved.get("pieces", []))
        st.session_state["annual_summaries_saved"] = list(saved.get("summaries", []))
        if saved:
            st.session_state["annual_fy"] = "1月〜12月" if int(saved.get("fy", 4)) == 1 else "4月〜3月（年度）"
            st.session_state["annual_on"] = bool(saved.get("on", True))
            for k, _, _ in ANNUAL_BALANCE_ITEMS:
                v = (saved.get("w") or {}).get(k, "ふつう")
                st.session_state[f"annual_w_{k}"] = v if v in ANNUAL_STRENGTHS else "ふつう"
    st.session_state.setdefault("annual_saved", [])
    st.session_state.setdefault("annual_summaries_saved", [])
    st.session_state.setdefault("annual_fy", "4月〜3月（年度）")
    st.session_state.setdefault("annual_on", True)
    for k, _, _ in ANNUAL_BALANCE_ITEMS:
        st.session_state.setdefault(f"annual_w_{k}", "ふつう")

    section_banner("📊", "年間の当直回数", "毎月の掲示用Excel、または集計だけのファイルから、先生ごとの回数を集計します", "任意", anchor="sec-annual")
    st.info("アプリからダウンロードした「当直表_掲示用_〇年〇月.xlsx」を、まとめて読み込んでください（手で直したあとのファイルでも大丈夫です）。"
            "当直表が見つからない月は、右の「集計だけのファイル」で先生ごとの回数を入れられます。\n\n"
            "**ファイルそのものは、どこにも保存しません。**「💾 途中保存」を押したときだけ、先生ごとの**回数**を保存します（どの日に誰が当直したかは保存しません）。")
    st.markdown("<style>.st-key-annual_upbox [data-testid='stFileUploaderDropzone']::before{content:'📄⬇　ここに掲示用Excelをドラッグ＆ドロップ（何か月分でも）'}"
                ".st-key-annual_sumbox [data-testid='stFileUploaderDropzone']{border-color:#7b5cc4 !important}"
                ".st-key-annual_sumbox [data-testid='stFileUploaderDropzone']::before{content:'📋⬇　ここに集計のファイルをドラッグ＆ドロップ';color:#4f3a8c}"
                ".st-key-annual_sumbox [data-testid='stFileUploaderDropzone'] button{background:#7b5cc4 !important}"
                ".st-key-annual_sumbox [data-testid='stDownloadButton'] button{border:1.5px solid #7b5cc4 !important;color:#4f3a8c !important;background:#fff !important;font-weight:700}"
                ".an-up-h{font-weight:800;color:#24405f;font-size:1rem;margin-bottom:2px}</style>", unsafe_allow_html=True)
    up_a, up_b = st.columns(2)
    with up_a, st.container(key="annual_upbox", border=True):
        st.markdown("<div class='an-up-h'>📄 掲示用Excel（日ごとの当直がわかる月）</div>", unsafe_allow_html=True)
        st.caption("アプリでダウンロードした掲示用Excel。何か月分でもまとめて読み込めます。")
        files = st.file_uploader("掲示用Excel（何か月分でも、まとめて選べます）", type=["xlsx"], accept_multiple_files=True,
                                 key=f"annual_upload_{st.session_state.get('uploader_gen', 0)}", label_visibility="collapsed")
    with up_b, st.container(key="annual_sumbox", border=True):
        st.markdown("<div class='an-up-h'>📋 集計だけのファイル（当直表が見つからない月）</div>", unsafe_allow_html=True)
        st.caption("先生ごとの回数だけが分かる場合は、ひな形に期間と回数を書いて読み込みます。分からない欄は空欄で大丈夫です。"
                   "アプリでダウンロードした「年間の当直回数」のExcelも読み込めます。")
        sum_files = st.file_uploader("集計だけのファイル", type=["xlsx", "csv"], accept_multiple_files=True,
                                     key=f"annual_sum_upload_{st.session_state.get('uploader_gen', 0)}", label_visibility="collapsed")
        st.download_button("📥 集計のひな形（Excel）をダウンロード", annual_summary_template_xlsx(), file_name="年間の当直回数_集計のひな形.xlsx",
                           mime=XLSX_MIME, key="annual_sum_template")
    cache = st.session_state.setdefault("annual_parsed", {})
    parsed = []
    for f in files or []:
        data = f.getvalue()
        key = hashlib.sha256(data).hexdigest()
        if key not in cache:
            try:
                cache[key] = annual_read_posting(data, f.name)
            except ValueError as e:
                cache[key] = str(e)
        if isinstance(cache[key], str):
            st.warning(cache[key])
        else:
            parsed.append(cache[key])
    st.session_state["annual_loaded"] = parsed

    pieces, notes = annual_merge(st.session_state["annual_saved"], parsed)
    st.session_state["annual_pieces"] = pieces

    c_fy, c_del = st.columns([2, 1], vertical_alignment="bottom")
    with c_fy:
        fy_label = st.selectbox("集計の区切り", ["4月〜3月（年度）", "1月〜12月"], key="annual_fy")
    fy_start = 1 if fy_label.startswith("1月") else 4
    months = annual_fy_months(year, month, fy_start)

    # 集計だけのファイル：期間が書いていなければ、ここで選んでもらう
    sum_cache = st.session_state.setdefault("annual_sum_parsed", {})
    loaded_sums = []
    month_choices = annual_month_span(f"{year - 3}-{month:02d}", f"{year}-{month:02d}")
    month_label = lambda k: f"{k[:4]}年{int(k[5:])}月"
    for f in sum_files or []:
        data = f.getvalue()
        key = hashlib.sha256(data).hexdigest()
        if key not in sum_cache:
            try:
                sum_cache[key] = annual_read_summary(data, f.name)
            except ValueError as e:
                sum_cache[key] = str(e)
        got = sum_cache[key]
        if isinstance(got, str):
            st.warning(got)
            continue
        sm = dict(got)
        if not (sm.get("start") and sm.get("end")):
            st.warning(f"「{f.name}」に集計の期間が書かれていません。下で、はじめの月とおわりの月を選んでください。")
            c1, c2 = st.columns(2)
            with c1:
                sm["start"] = st.selectbox(f"「{f.name}」の集計の期間（はじめ）", [None] + month_choices, format_func=lambda k: "選んでください" if k is None else month_label(k),
                                           key=f"annual_sum_start_{key[:12]}")
            with c2:
                sm["end"] = st.selectbox(f"「{f.name}」の集計の期間（おわり）", [None] + month_choices, format_func=lambda k: "選んでください" if k is None else month_label(k),
                                         key=f"annual_sum_end_{key[:12]}")
            if not (sm["start"] and sm["end"]):
                continue
            if sm["start"] > sm["end"]:
                sm["start"], sm["end"] = sm["end"], sm["start"]
        span = annual_month_span(sm["start"], sm["end"])
        if not all(m in months for m in span):
            st.warning(f"「{f.name}」の期間（{month_label(sm['start'])}〜{month_label(sm['end'])}）が、集計の区切り（{month_label(months[0])}〜{month_label(months[-1])}）からはみ出しています。"
                       "期間を区切りの中に収めてください（はみ出しているあいだは、集計に入れません）。")
        sm["created"] = remote_now()
        loaded_sums.append(sm)
    # 途中保存から戻した分と合わせる（同じ期間なら、今回読み込んだほうを使う）
    spans = {(sm["start"], sm["end"]) for sm in loaded_sums}
    summaries = [sm for sm in st.session_state["annual_summaries_saved"] if (sm.get("start"), sm.get("end")) not in spans] + loaded_sums
    st.session_state["annual_summaries"] = summaries
    used_sums, sum_months = annual_summary_months(summaries, months)

    with c_del:
        if st.session_state["annual_saved"] or st.session_state["annual_summaries_saved"]:
            st.button("🗑 保存した集計を消す", key="annual_clear", on_click=annual_clear_saved, use_container_width=True,
                      help="途中保存から戻した回数を消します。消したあとに途中保存すると、保存の中からも消えます。")
    status = {}
    for p in pieces:
        if p["month"] in months:
            status[p["month"]] = "ok" if p.get("loaded") or status.get(p["month"]) == "ok" else "sv"
    hidden = sorted(m for m in sum_months if status.get(m) == "ok")   # 集計だけのファイルを優先したため、使わなかった掲示用Excelの月
    for m in sum_months:
        status[m] = "sm"
    marks = {"ok": " ✓", "sv": " 💾", "sm": " 📋"}
    chips = "".join(f"<div class='{status.get(m, 'no')}'>{annual_month_label(m)}{marks.get(status.get(m), '')}</div>" for m in months)
    st.markdown(ANNUAL_TABLE_CSS + f"<div class='an-m'>{chips}</div>", unsafe_allow_html=True)
    st.caption("📋 集計だけのファイルから入れた月　✓ 今回掲示用Excelを読み込んだ月　💾 途中保存から戻した月。"
               "同じ日が重なったときは、あとから作ったほうのExcelを使います。集計だけのファイルの期間に入る月は、その月の掲示用Excelを使いません（二重には数えません）。")
    if hidden:
        st.caption("・" + "・".join(annual_month_label(m) for m in hidden) + "は、集計だけのファイルの回数を使いました（その月の掲示用Excelは使っていません）。")
    for n in notes:
        st.warning(n)

    tot = annual_totals(pieces, months, summaries=summaries)
    if not tot:
        st.caption("まだ集計する回数がありません。掲示用Excelか集計だけのファイルを読み込むと、ここに先生ごとの回数が出ます。")
        annual_publish_rates([])
        return None
    got = [m for m in months if m in status]
    span = f"{got[0][:4]}年{int(got[0][5:])}月〜{got[-1][:4]}年{int(got[-1][5:])}月" if got else ""
    rows, avg = annual_rows(tot, staff_names)
    annual_publish_rates(rows)   # 医師条件の表の「📊 年間（1か月あたり）」の列に使う
    st.markdown(f"**先生ごとの回数（{span}）**　<span style='font-size:.8rem;opacity:.7'>回数の右の（ ）内は1か月あたり。表の中でスクロールできます</span>", unsafe_allow_html=True)
    st.markdown(annual_table_html(rows), unsafe_allow_html=True)
    st.caption(f"在籍月数：読み込んだ月のうち、その先生の名前がのっている月の数です（日数の割合で数えます。集計だけのファイルの月は、そのファイルの在籍月数）。"
               f"平均（1か月あたり{avg:.1f}回）とくらべます。「―」は分からない回数です。")
    st.download_button("📥 集計をExcelでダウンロード", annual_totals_xlsx(rows, f"年間の当直回数（{span}）"),
                       file_name=f"年間の当直回数_{span}.xlsx", key="annual_xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # 年間の回数は見るだけ（当直案の計算には使わない）。偏りは、医師条件の「多め希望」で先生ごとに調整する
    st.markdown("<div style='background:#eaf3ff;border:1.5px solid #6ea2e6;border-radius:10px;padding:9px 14px;color:#1d4e89;"
                "line-height:1.7;margin:4px 0 8px'>💡 <b>この集計は、見るだけです（当直案の計算には使いません）。</b><br>"
                "偏りが気になる先生は、医師条件の「病棟/外来 多め」「宿直/日直 多め」「平日/休日 多め」で調整してください。</div>",
                unsafe_allow_html=True)
    return None


def annual_snapshot(year, month):
    """途中保存に入れる年間の回数（月ごとの回数と、集計だけのファイルの回数。日ごとの当直は入れない）。古い月は消す。"""
    oldest = (year * 12 + month - 1) - ANNUAL_KEEP_MONTHS
    pieces = []
    for p in st.session_state.get("annual_pieces", []):
        y, m = int(p["month"][:4]), int(p["month"][5:])
        if y * 12 + m - 1 >= oldest:
            pieces.append({k: p[k] for k in ("month", "start", "end", "ndays", "created", "roster", "counts")})
    sums = [{k: sm.get(k) for k in ("label", "start", "end", "created", "docs")}
            for sm in st.session_state.get("annual_summaries", []) if sm.get("start") and sm.get("end")
            and int(sm["end"][:4]) * 12 + int(sm["end"][5:]) - 1 >= oldest]
    return {"pieces": pieces, "summaries": sums, "fy": 1 if str(st.session_state.get("annual_fy", "")).startswith("1月") else 4,
            "on": bool(st.session_state.get("annual_on", True)),
            "w": {k: st.session_state.get(f"annual_w_{k}") or "しない" for k, _, _ in ANNUAL_BALANCE_ITEMS}}


def annual_result_view(df_result, year, month, staff_names, created, now_word="今月"):
    """作成した当直案の下に、先生ごとの「年間（今月込み）」の回数を出す（年間の回数を読み込んでいるときだけ）。"""
    if not (st.session_state.get("annual_pieces") or st.session_state.get("annual_summaries")):
        return
    fy_start = 1 if str(st.session_state.get("annual_fy", "")).startswith("1月") else 4
    months = annual_fy_months(year, month, fy_start)
    result_file = annual_result_file(df_result, year, month, staff_names, created)
    pieces, _ = annual_merge(st.session_state.get("annual_saved", []), st.session_state.get("annual_loaded", []) + [result_file])
    months = sorted(set(months) | {f"{d.year}-{d.month:02d}" for d in result_file["days"]})
    tot = annual_totals(pieces, months, summaries=st.session_state.get("annual_summaries", []))
    if not tot:
        return
    now_counts = annual_counts_of(result_file["days"].values())
    rows, _ = annual_rows(tot, staff_names)
    body = ""
    for r in rows:
        if not r["here"]:
            continue
        body += f"<tr><td>{html_escape(r['name'])}</td><td>{now_counts.get(r['name'], [0])[0]}回</td>{annual_body_cells(r)}</tr>"
    head = annual_head_html(["先生", f"{now_word}の当直"], "（年間）")
    st.subheader(f"📊 年間の回数（{now_word}込み）")
    st.caption("回数の右の（ ）内は1か月あたり。この当直案を入れた場合の、年間の回数です（表の中でスクロールできます）。")
    st.markdown(ANNUAL_TABLE_CSS + f"<div class='an-box'><table class='an-t'><thead>{head}</thead><tbody>{body}</tbody></table></div>", unsafe_allow_html=True)


# 期間指定版の途中保存のシート（閲覧用の画面で、通常版の途中保存と一緒に選べるようにする）
WORK_STATE_SHEETS_PERIOD = {"legacy": "saved_state_period", "history": "saved_history_period"}


def staff_view_page():
    """
    閲覧専用の画面（…/?mode=view）。閲覧用パスワード（view_password）を入れた人だけが、
    管理者が途中保存した医師条件を見られる（通常版・期間指定版の途中保存から選ぶ）。備考は出さない。ここでは何も直せない・保存しない。
    """
    st.title("📋 当直の条件と回数（閲覧用）")
    st.caption("医師条件・グループの設定・年間の当直回数を見られます。ここでは直せません。")
    if not remote_is_configured():
        st.error("保存用のGoogleスプレッドシートが設定されていないため、表示できません。管理者に連絡してください。")
        return
    view_password = remote_secret("view_password")
    if not view_password:
        st.warning("閲覧用パスワード（view_password）が設定されていません。管理者に連絡してください。")
        return
    if not st.session_state.get("view_ok"):
        with st.form("view_login"):
            pw = st.text_input("閲覧用パスワード", type="password")
            ok = st.form_submit_button("開く")
        if ok:
            if pw == str(view_password):
                st.session_state["view_ok"] = True
                st.rerun()
            st.error("パスワードが違います。")
        st.caption("管理用パスワードとは別のパスワードです。")
        return
    entries = [("", e) for e in work_state_entries(WORK_STATE_SHEETS)] + [("・期間指定版", e) for e in work_state_entries(WORK_STATE_SHEETS_PERIOD)]
    entries = [(kind, e) for kind, e in entries if e[2].get("staff_csv")]
    if not entries:
        st.info("まだ途中保存された医師条件がありません。")
        return
    entries.sort(key=lambda x: str(x[1][1]), reverse=True)
    labels = [f"{work_state_title(e[2])}{kind}（{e[1]} の途中保存）" for kind, e in entries]
    pick = st.selectbox("どの保存を見るか", range(len(entries)), format_func=lambda i: labels[i], key="view_pick")
    data = entries[pick][1][2]
    st.info("管理者が途中保存した医師条件です（備考は表示しません）。最新の内容は、管理者が途中保存したときに変わります。")
    df = parse_staff_csv(str(data["staff_csv"]).encode("utf-8"))
    df = df[[c for c in df.columns if not str(c).startswith("備考") and not str(c).startswith("_")]]
    df = add_pref_columns(add_day_gap_column(df.loc[:, ~df.columns.duplicated()]))   # 多め希望の列（ない保存では空欄で出す）
    for c in (NG_COLUMN, REQUEST_COLUMN, "翌日PM duty", DOC_GROUP_COLUMN, *DOC_PREF_COLUMNS):
        if c in df.columns:
            df[c] = df[c].map(lambda v: "" if clean_text(v).lower() in ("nan", "none", "<na>") else clean_text(v))
    # NG日の列は、いちばん長い先生のNG日が入る幅にする（ほかの列は列名が見える最小限）
    _ng_len = max([len(clean_text(v)) for v in df.get(NG_COLUMN, pd.Series(dtype=str)).tolist()] + [10])
    staff_table(df, editable=False, show_ng=True, max_height=900, ng_width=min(900, 8 * _ng_len + 30))
    # グループの設定（見るだけ）：グループごとに担当する枠（◯△×）とメモ
    rules = normalize_group_rules(data.get("group_rules")) if data.get("group_rules") is not None else {}
    st.markdown("<div style='font-weight:800;color:#24405f;font-size:1.02rem;margin-top:14px'>👥 グループの設定（見るだけ）</div>",
                unsafe_allow_html=True)
    if data.get("group_rules") is None:
        st.info("この保存には、グループの設定が入っていません（グループの設定を使い始める前に保存したものです）。")
    elif not rules:
        st.info("この保存では、グループは設定されていません。")
    if rules:
        slots = doc_group_slots()
        gdf = pd.DataFrame([[g] + [r[s] for s in slots] + [r.get(DOC_GROUP_LEVEL, DOC_GROUP_LEVEL_OFF), r.get(DOC_GROUP_MEMO, "")] for g, r in rules.items()],
                           columns=["グループ名"] + slots + [DOC_GROUP_LEVEL, DOC_GROUP_MEMO])
        gcfg = {"グループ名": st.column_config.TextColumn("グループ名", width=120)}
        gcfg.update({s: st.column_config.TextColumn(s, width=header_fit_width(s, 0)) for s in slots})
        gcfg[DOC_GROUP_LEVEL] = st.column_config.TextColumn(DOC_GROUP_LEVEL, width=header_fit_width(DOC_GROUP_LEVEL, 150))
        gcfg[DOC_GROUP_MEMO] = st.column_config.TextColumn(DOC_GROUP_MEMO, width=240)
        st.dataframe(gdf, hide_index=True, use_container_width=False, column_config=gcfg)
        st.caption("◯：担当する　△：できるだけ担当しない（ほかの先生で埋まらないときだけ担当する）　×：担当しない（ただし、確定当直で決めた分は×でも担当します）")
    # 年間の当直回数（見るだけ）：途中保存に入っている回数（先生ごと・月ごとの回数と、集計だけのファイルの回数）を、その保存の集計の区切りで合計する
    annual = data.get("annual") or {}
    a_pieces, a_sums = list(annual.get("pieces") or []), list(annual.get("summaries") or [])
    shown_annual = False
    if a_pieces or a_sums:
        y, m = int(data.get("year", 0)), int(data.get("month", 0))
        months = annual_fy_months(y, m, 1 if int(annual.get("fy", 4) or 4) == 1 else 4)
        tot = annual_totals(a_pieces, months, summaries=a_sums)
        if tot:
            _, sum_months = annual_summary_months(a_sums, months)
            got = sorted({p["month"] for p in a_pieces if p["month"] in months} | set(sum_months))
            span = f"{got[0][:4]}年{int(got[0][5:])}月〜{got[-1][:4]}年{int(got[-1][5:])}月" if got else ""
            names = [clean_text(n) for n in df["先生の名前"].tolist() if clean_text(n)]
            rows, avg = annual_rows(tot, names)
            st.markdown(f"<div style='font-weight:800;color:#24405f;font-size:1.02rem;margin-top:14px'>📊 年間の当直回数（{span}・見るだけ）</div>"
                        "<span style='font-size:.8rem;opacity:.7'>回数の右の（ ）内は1か月あたり。表の中でスクロールできます</span>", unsafe_allow_html=True)
            st.markdown(annual_table_html(rows), unsafe_allow_html=True)
            st.caption(f"在籍月数：その先生の名前がのっている月の数です（日数の割合で数えます）。平均（1か月あたり{avg:.1f}回）とくらべます。"
                       "「―」は分からない回数です。途中保存したときの集計なので、作成中の当直案の分は入っていません。")
            shown_annual = True
    if not shown_annual:
        st.markdown("<div style='font-weight:800;color:#24405f;font-size:1.02rem;margin-top:14px'>📊 年間の当直回数（見るだけ）</div>",
                    unsafe_allow_html=True)
        st.info("この保存には、年間の当直回数の集計が入っていません（管理者がアプリの「📊 年間の当直回数」で掲示用Excelなどを読み込んでから途中保存すると、ここに出ます）。")


# ==========================================
# 使い方（…/?mode=manual）：別のタブで開いて、アプリと並べて見ながら操作できるページ。パスワードなしで開ける（データは出さない）
# ==========================================
MANUAL_LINK_HTML = ("<a href='./?mode=manual' target='_blank' style='display:inline-block;margin:2px 0 8px;border:1.5px solid #47749b;"
                    "color:#24405f !important;background:#fff;border-radius:8px;padding:5px 14px;font-weight:700;text-decoration:none !important'>"
                    "📖 使い方を別のタブで開く</a>"
                    "<span style='color:#6a7686;font-size:.85rem;margin-left:8px'>アプリと並べて、見ながら操作できます（入力中の内容は消えません）</span>")

USAGE_CSS = """<style>
.ug{line-height:1.75;font-size:.97rem}
.ug .lead{background:#fff8e6;border:1.5px solid #f0b46a;border-radius:10px;padding:8px 14px;margin:4px 0 12px;color:#7a4a0a;font-size:.93rem}
.ug .step{display:flex;gap:14px;border:1px solid #e3e7ee;border-radius:10px;padding:10px 14px;margin:0 0 10px;background:#fff}
.ug .num{flex:none;width:34px;height:34px;border-radius:50%;background:#24405f;color:#fff;font-weight:800;text-align:center;line-height:34px;font-size:1rem}
.ug .ck{flex:none;width:22px;height:22px;border:2px solid #24405f;border-radius:5px;margin:5px 6px 0}
.ug .step.opt .ck{border-style:dashed}
.ug .band{background:#24405f;color:#fff;border-radius:8px;padding:7px 14px;font-weight:800;margin:16px 0 8px;font-size:1rem}
.ug .band.o{background:#c2410c}.ug .band small{font-weight:500;opacity:.85;margin-left:8px;font-size:.8rem}
.ug .t{font-weight:800;font-size:1.02rem;color:#24405f}.ug .t small{font-weight:500;color:#6a7686;margin-left:8px;font-size:.8rem}
.ug .where{display:inline-block;background:#eef3fa;color:#24405f;border-radius:6px;padding:0 8px;font-size:.8rem;margin:2px 0}
.ug ul{margin:4px 0 0;padding-left:20px}.ug li{margin:1px 0}
.ug .tip{background:#eaf3ff;border-left:4px solid #6ea2e6;border-radius:6px;padding:4px 10px;color:#1d4e89;font-size:.88rem;margin-top:6px}
.ug .warn{background:#fff1f0;border-left:4px solid #e5534b;border-radius:6px;padding:4px 10px;color:#8a1f17;font-size:.88rem;margin-top:6px}
.ug h4{font-size:1.05rem;color:#24405f;border-left:5px solid #47749b;padding-left:8px;margin:18px 0 8px}
.ug table{border-collapse:collapse;font-size:.88rem;margin:4px 0 8px}.ug th,.ug td{border:1px solid #d6dde8;padding:4px 10px;text-align:left}
.ug th{background:#eef1f5}
.ug dl dt{font-weight:800;color:#24405f;margin-top:12px}.ug dl dd{margin:2px 0 0 0}
.ug .gbtn{display:inline-block;background:#1d6f42;color:#fff;border-radius:8px;padding:5px 14px;font-weight:700;font-size:.9rem;margin:2px 0 4px}
.ug .part{border:1px solid #e3e7ee;border-radius:10px;padding:10px 14px;margin:0 0 10px;background:#fff}
.ug .part .pt{font-weight:800;color:#24405f;margin-bottom:4px}.ug .part .pw{font-size:.85rem;color:#6a7686}
</style>"""


def _ug_step(num, title, where, items, tip="", optional=False, warn=""):
    """
    使い方の手順を1つの枠にする（印・見出し・操作する場所・箇条書き・ヒント）。
    num は、アプリの見出しにある番号（1〜4など。アプリと同じ見た目の丸）。アプリに番号がない手順は None（チェック欄）にする。
    """
    sub = "<small>必要なときだけ</small>" if optional else ""
    w = f"<span class='where'>場所：{where}</span>" if where else ""
    lis = "".join(f"<li>{x}</li>" for x in items)
    extra = (f"<div class='tip'>💡 {tip}</div>" if tip else "") + (f"<div class='warn'>⚠ {warn}</div>" if warn else "")
    mark = "<div class='ck'></div>" if num is None else f"<div class='num'>{num}</div>"
    return (f"<div class='step{' opt' if optional else ''}'>{mark}<div style='flex:1'>"
            f"<div class='t'>{title}{sub}</div>{w}<ul>{lis}</ul>{extra}</div></div>")


def _ug_band(title, orange=False, tag=""):
    """アプリの見出しの帯と同じ名前・同じ色の帯（使い方の見出し）。"""
    t = f"<small>{tag}</small>" if tag else ""
    return f"<div class='band{' o' if orange else ''}'>{title}{t}</div>"


def _ug(html):
    st.markdown(USAGE_CSS + f"<div class='ug'>{html}</div>", unsafe_allow_html=True)


NG_KIND_TABLE = ("<table><tr><th>表示</th><th>意味</th></tr>"
                 "<tr><td>宿NG</td><td>宿直（夜）に入れない（「宿直」のボタンだけ押した状態）</td></tr>"
                 "<tr><td>日NG</td><td>日直（日中）に入れない。休日だけ（「日直」のボタンだけ押した状態）</td></tr>"
                 "<tr><td>全NG</td><td>日直・宿直とも入れない。休日だけ（両方のボタンを押した状態）</td></tr>"
                 "<tr><td>⚠</td><td>翌日PMにdutyがあるため、原則その日の宿直を外す日</td></tr>"
                 "<tr><td>灰色の「別途決定」</td><td>別の方法で当番を決める日。NGは入れられません</td></tr></table>")


# 使い方の目次（左の帯に縦に並べる）
USAGE_PAGES = ["🔰 はじめに", "📅 毎月の流れ", "📨 NG日の集め方", "🗓️ 結果の見方・Excel", "📆 期間指定版", "🆘 困ったとき"]


def manual_page():
    """使い方のページ（…/?mode=manual）。別のタブで開いて、アプリと並べて見ながら操作できる。パスワードなしで開ける。"""
    st.title("📖 当直作成アプリ　使い方")
    st.markdown("<a href='./' target='_self' style='display:inline-block;border:1px solid #2e8a5b;color:#1f6b45 !important;border-radius:8px;"
                "padding:2px 12px;font-weight:700;text-decoration:none !important'>← アプリを開く</a>"
                "<span style='color:#6a7686;font-size:.85rem;margin-left:8px'>このページはアプリとは別のタブです。アプリのタブと行き来しながら操作できます。</span>",
                unsafe_allow_html=True)
    # 目次は左の帯（サイドバー）に縦に並べる。押した項目だけを右に出す（スクロールしても目次は動かない）
    st.markdown("<style>.st-key-usage_nav [role='radiogroup']{gap:2px}"
                ".st-key-usage_nav [role='radiogroup'] label{padding:7px 10px;border-radius:8px;width:100%;margin:0}"
                ".st-key-usage_nav [role='radiogroup'] label > div:first-child{display:none}"
                ".st-key-usage_nav [role='radiogroup'] label:has(input:checked){background:#fff;box-shadow:0 1px 3px rgba(0,0,0,.1)}"
                ".st-key-usage_nav [role='radiogroup'] label:has(input:checked) p{color:#c2410c;font-weight:800}"
                ".st-key-usage_nav [role='radiogroup'] label p{font-size:1rem}</style>", unsafe_allow_html=True)
    with st.sidebar:
        with st.container(key="usage_nav"):
            page = st.radio("📖 使い方", list(USAGE_PAGES), key="usage_page")
    st.markdown(f"### {page}")
    t_intro, t_flow, t_ng, t_result, t_period, t_trouble = (page == p for p in USAGE_PAGES)
    if t_intro:
        _ug("<h4>このアプリでできること</h4>"
            "<ul><li>先生方の条件（回数・間隔・希望日・NG日など）をもとに、毎月の当直案を自動で作ります。</li>"
            "<li>NG日は、管理者が画面のカレンダーで入れます。先生方に専用のURL（QRコード）から、スマートフォンで入れてもらうこともできます（左の目次の「NG日の集め方」）。</li></ul>"
            "<h4>画面のつくり</h4>"
            + _ug_step("上", "いちばん上の紺色の帯", "", ["「💾 作業を途中保存する」と「📂 前回の作業の続きから始める」があります。スクロールしても、いつも見えています。"])
            + _ug_step("中", "真ん中（上から順に5つの項目）", "",
                       ["📅 年月・休日・必要人数の設定", "📅 前月・今月・翌月の確定当直（任意・オレンジの見出し）",
                        "👩‍⚕️ 先生ごとの条件（① 名簿と医師条件 → ② NG日 → ③ 医師条件の調整 → ④ 当直の枠は足りそう？）",
                        "📊 年間の当直回数（任意）", "🎯 当直案の作成・確認"], tip="上から順に進めれば、1か月分の作業が終わります。")
            + _ug_step("下", "いちばん下の紺色の帯", "", ["「🚀 当直案を作成」ボタンと、「医師条件」「NG日」「結果」へ移動するボタン、作成結果（未決定の枠の数）があります。"])
            + "<h4>管理用パスワード</h4><ul><li>途中保存・続きからの読み込み・前月の医師条件の引き継ぎ・専用URLの管理（受付・取り込みなど）に必要です。</li>"
              "<li>一度入れると、画面を閉じるまでほかの場所では入れなくて大丈夫です。</li></ul>"
            "<h4>大事なこと</h4><ul>"
            "<li>画面に入れた内容は、しばらく使わないと消えます。作業の区切りごとに、上の帯の<b>「💾 作業を途中保存する」</b>を押してください。</li>"
            "<li>作った当直案は途中保存に入りません。気に入った案は、Excel（医師別の当直一覧と掲示用のExcel）でダウンロードしてください。</li>"
            "<li><b>手直しした掲示用Excelは、毎月手元に残してください。</b>翌月の確定当直と、年間の当直回数の集計に、そのまま読み込めます。</li></ul>")
    if t_flow:
        _ug("<div class='lead'>毎月の作業は、アプリの画面を<b>上から順に</b>進めれば終わります。紺色・オレンジの帯は、アプリの見出しの帯と同じ名前です。</div>"
            + _ug_band("📅 年月・休日・必要人数の設定")
            + _ug_step(1, "作る月を選ぶ", "",
                       ["ふつうは<b>「来月」</b>を押すだけです。「◀ 前の月」「次の月 ▶」でも動かせます。",
                        "同じ月の続きなら、いちばん上の紺色の帯の<b>「📂 前回の作業の続きから始める」</b>で戻せます（NG日も含めて戻るので、ほかの手順は確認だけで済みます）。"])
            + _ug_step(2, "休日と別途決定の日の設定", "",
                       ["土日祝は自動で休日になります。",
                        "<b>休日にする</b>：病院独自の休みなど、平日でも休日と同じ扱いにする日（日直の枠ができます）。",
                        "<b>別途決定</b>：年末年始など、別の方法で当番を決める日（この日の当直は作らず、先生もNGを入れられません）。"])
            + _ug_step(None, "人数を増やす枠", "", ["1つの枠を2名以上にする日があれば、行を足して設定します。"], optional=True)
            + _ug_band("📅 前月・今月・翌月の確定当直", orange=True, tag="任意")
            + _ug_step(None, "確定当直を入れる", "",
                       ["すでに決まっている当直や、前月末・翌月初めの当直を入れます（勤務間隔の計算に使います）。",
                        "<b>方法1</b>：「✏️ 確定当直の表」に直接入れる（前後7日分の日付は並んでいるので、先生を選ぶだけ）。",
                        "<b>方法2</b>：前月の掲示用Excel（手直ししたもの）を読み込む。重なった日は、あとから読み込んだ方が使われます。"],
                       tip="確定当直に入れるのは、ふつうは<b>前月</b>の掲示用Excelです。", optional=True)
            + _ug_band("👩‍⚕️ 先生ごとの条件（NG日・回数など）")
            + _ug_step(1, "以前の医師条件の取り込み・医師の追加（外す）", "",
                       ["<b>方法1</b>：新しい月なら<b>「📋 前月の医師条件を引き継ぐ」</b>を押します。前月の途中保存の医師条件（回数・間隔・翌日PM duty・グループ・多め希望・備考など）が入り、NG日・希望日・コメントは空になります。",
                        "<b>方法2</b>：手元のExcel（ひな形は「📥 ひな形（Excel）をダウンロード」）を、緑の点線の枠にドラッグ＆ドロップします。NG日・希望日が入っているときは「今月の分として使う」か「空にする」かを選びます（前月のファイルなら「空にする」）。",
                        "「今の名簿」と「読み込んだ医師条件」（見るだけの表）で、正しく読み込めたかを確認します。",
                        "新しく来た先生は<b>「➕ 先生を追加する」</b>（名前・入れる位置・条件のもと）、いなくなった先生は<b>「➖ 先生を外す」</b>で直します。"])
            + _ug_step(2, "NG日を入れる", "",
                       ["「NG日を入力する医師」で先生を選びます（「◀」「▶」で前後の先生に移れます）。",
                        "<b>翌日PM duty</b>の曜日も、ここのボタンで選べます（例：木曜PMにdutyがあれば「水」）。③の表にもすぐ反映され、カレンダーに⚠が付きます。",
                        "入れない日の「日直」「宿直」を押してNGにし、最後に<b>「NG日を反映する」</b>を押します。",
                        "専用URLで集める場合は、ここで受付・取り込みを行います（左の目次の「NG日の集め方」）。"])
            + _ug_step(3, "医師条件を調整する", "",
                       ["当直案の枠ができるだけ埋まるように、この表で回数・上限・勤務間隔・希望日などを直します（①の表は見るだけ。直すのは③の表）。",
                        "表は横にスクロールできます。表の上の<b>「🗺 この表の列」</b>で、右にどの列があるかを確認できます。列名にカーソルを合わせると、その列の説明が出ます。",
                        "<b>グループ</b>：表の上の「👥 グループの設定」で、グループごとに担当する枠（◯担当する・△できるだけ担当しない・×担当しない）を決め、表の「グループ」の列で先生のグループを選びます。"
                        "「回数（余裕があるとき）」で、枠に余裕がある月に、グループごとに回数を多め・少なめにできます（対象外のグループは差をつけません）。",
                        "<b>多め希望</b>：「病棟/外来 多め」「宿直/日直 多め」「平日/休日 多め」で、できれば多めに入れたい種類を選べます（月の回数は増えず、種類の割合を寄せるだけ）。",
                        "灰色の列（💬 先生からのコメント・📊 年間（1か月あたり））は見るだけで、直せません。",
                        "<b>枠ごとの上限（A宿・B宿・外宿・A日・B日・外日）</b>の列は、ふだんは出しません（担当する枠はグループで決めます）。空欄のときは、1つの枠はできるだけ月2回までにします。各枠の上限を先生ごとに個別に決めたいときは、表の下の紫のスイッチで出します。値が入っている先生がいると、自動で出して黄色で知らせます。",
                        "グループ内でほかの先生と違う設定（休日上限・勤務間隔・枠ごとの上限）があると、表の下に黄色の欄で知らせます。わざとならそのままで大丈夫です。"])
            + _ug_step(4, "当直の枠は足りそう？（目安）", "", ["赤い棒（足りない枠）がないかを確認します。条件を直すと、すぐに変わります。"])
            + _ug_band("📊 年間の当直回数", tag="任意")
            + _ug_step(None, "年間の当直回数を見る", "",
                       ["これまでの月の掲示用Excelを、まとめて読み込みます。当直表が見つからない月は、右の「📋 集計だけのファイル」に、ひな形で期間と先生ごとの回数を書いて読み込みます。",
                        "先生ごとの回数（病棟・外来、宿直・日直、平日・休日、合計）が出ます。回数の右の（ ）内は1か月あたりです。",
                        "集計は見るだけで、当直案の計算には使いません。偏りが気になる先生は、③の表の「📊 年間（1か月あたり）」の列を見ながら、多め希望で調整します。"], optional=True)
            + _ug_band("🎯 当直案の作成・確認")
            + _ug_step(None, "不足時の優先順位を決める", "", ["人数が足りない月に、どの枠から埋めるかを決めます。"], optional=True)
            + _ug_step(None, "当直案を作る", "いちばん下の紺色の帯",
                       ["<b>「🚀 当直案を作成」</b>を押します（最大1分少々かかります）。",
                        "押すたびに違う案ができるので、気に入るまで押し直して構いません。",
                        "未決定の枠が出たら、条件を直して作り直します。帯の移動ボタンで、条件と結果を行き来できます。"])
            + _ug_step(None, "確認してダウンロードする", "",
                       ["🗓️ 医師別の当直一覧の下の「📥 この一覧をExcelでダウンロード（印刷用）」と、📅 作成した当直案の下の「📥 この当直案を掲示用のExcelでダウンロード（A3・印刷用）」でダウンロードします（左の目次の「結果の見方・Excel」）。",
                        "掲示用Excelは手直しして掲示し、手元に残しておきます。"])
            + _ug_band("💾 いちばん上の紺色の帯")
            + _ug_step(None, "途中保存する", "",
                       ["最後に<b>「💾 作業を途中保存する」</b>を押します。翌月は「📋 前月の医師条件を引き継ぐ」で、この医師条件から始められます。",
                        "途中保存は、保存した回ごとに新しいものから20回分まで残ります。「📂 前回の作業の続きから始める」で、どの回から再開するかを選べます。"],
                       tip="途中保存に入るもの：年月・休日・別途決定の日・増員設定・確定当直・医師条件（NG日・グループの設定を含む）・取り込んだ先生のコメント・不足時の優先順位・アラートの基準・年間の当直回数（回数だけ。Excelそのものは入りません）。作った当直案は入りません。"))
    if t_ng:
        _ug("<div class='lead'>NG日は「👩‍⚕️ 先生ごとの条件」の <b>② NG日を入れる</b> で入れます。</div>"
            "<h4>方法1：管理者が直接入れる</h4>"
            + _ug_step(None, "先生を選ぶ", "② NG日を入れる", ["「NG日を入力する医師」で先生を選びます。紙やメールで集めたNG日は、こちらで入れます。"])
            + _ug_step(None, "NGを押す", "先生ごとのカレンダー",
                       ["入れない日の「日直」「宿直」のボタンを押してNGにします。もう一度押すと戻ります。",
                        "「全日NGにする」「すべてOKに戻す」で、まとめて指定・解除もできます。"])
            + _ug_step(None, "反映する", "カレンダーの下",
                       ["<b>「NG日を反映する」</b>を押します。反映していない変更があると、ボタンがオレンジに光ります。",
                        "反映せずに別の先生に切り替えようとすると、注意が出ます。"])
            + "<h4>方法2：先生方に専用URLから入れてもらう</h4>"
            "<div class='lead'>② の中の「📨 先生方からのNG日入力（専用URL）」を開き、管理用パスワードを入れて使います。</div>"
            + _ug_step(None, "受付を開始する", "📨 先生方からのNG日入力（専用URL）",
                       ["先に、いちばん上のカレンダーで「休日にする」「別途決定」を済ませておきます。",
                        "締め切り日を書いて<b>「受付を開始する」</b>を押します。受付中にカレンダーを変えても、パスワードを入れた状態なら自動で先生側に反映されます。"])
            + _ug_step(None, "URLを配る", "同じ欄の「配布用の案内書類（QRコード付き）」",
                       ["案内書類をダウンロードして印刷し、点線で三つ折りにして、封筒で各先生に渡します。",
                        "URLは毎月同じです。翌月からは「入力が始まりました」と伝えるだけです。"], optional=True)
            + _ug_step(None, "先生が入力する", "先生のスマートフォン",
                       ["先生はQRコードを読み取り、入れない日の「日直」「宿直」を押してNGにし、「NG日とコメントを保存する」を押します。",
                        "締め切りまでは何度でも直せます。最後に保存した内容が使われます。締め切ったあとも、自分の提出内容は見られます。"])
            + _ug_step(None, "提出状況を確かめる", "管理画面の表", ["提出済み・未提出・コメントを確認できます。「🔄 最新の状態に更新」で最新になります。"])
            + _ug_step(None, "締め切る", "管理画面", ["<b>「🔒 締め切る」</b>を押すと、先生側は直せなくなります。"])
            + _ug_step(None, "取り込む", "管理画面の「取り込むと変わる先生」",
                       ["内容とコメントを確かめ、取り込む先生を選んで<b>「⬇️ 選んだ先生のNG日を取り込む」</b>を押します。",
                        "その先生のNG日は提出内容に置き換わります。コメントは、③の表の「💬 先生からのコメント」の列や、当直一覧のメモ欄に出ます。"])
            + "<h4>方法2でやり直したいとき</h4><ul>"
            "<li><b>取り込みを取り消す</b>：「↩️ 取り込みを取り消す」で、取り込む前のNG日に戻せます（先生ごとのNG日の欄にもボタンがあります）。アプリを開き直したり、途中保存から読み込んだりしたあとは取り消せないので、カレンダーで直してください。</li>"
            "<li><b>先生の提出をリセットする</b>：入力をやり直してもらうときに使います。未提出に戻り、先生の画面も空になります（これまでの提出は履歴に残ります）。</li>"
            "<li><b>URLを作り直す</b>：案内書類をなくした・URLが漏れたときに使います。古いURLは使えなくなります。</li></ul>"
            "<h4>NGの種類（どちらの方法でも共通）</h4>" + NG_KIND_TABLE)
    if t_result:
        # アプリの結果の画面と同じ順・同じ名前にする（ダウンロードのボタンは、それぞれの表のすぐ下）
        _ug("<div class='lead'>結果は、作成ボタンを押すと次の順に出ます。ダウンロードのボタンは、それぞれの表のすぐ下にあります。</div>"
            + _ug_band("🗓️ 医師別の当直一覧")
            + "<div class='part'><div class='pt'>⚠ アラートの基準</div><ul><li>表のすぐ上で、右端の回数がこの数以上になった枠を赤く囲む基準を決めます"
              "（初期設定は宿直4回以上・休日3回以上。0にした項目はアラートなし）。</li></ul></div>"
            + "<div class='part'><div class='pt'>表の見方（表のすぐ上の凡例と同じ）</div>" + MATRIX_LEGEND_CSS + matrix_legend_html()
            + "<ul><li><b>右端の回数</b>：病棟・外来／宿直・日直／平日・休日／合計。</li>"
              "<li>マスにカーソルを合わせると、その先生の行と、その日の列がオレンジの線で囲まれます。</li></ul></div>"
            + "<div class='part'><span class='gbtn'>📥 この一覧をExcelでダウンロード（印刷用）</span><div class='pw'>表のすぐ下の緑のボタン</div><ul>"
              "<li>印刷用の表（A3横・1枚）です。翌日PM dutyの日は、斜線ではなく黄色の塗りで表します。</li>"
              "<li>右の回数は計算式なので、マスに枠名を書き足したり消したりすると自動で変わります。赤い枠も自動で付きます。</li>"
              "<li>「⚠️ 未決定」の行の枠を、表の中のマスへドラッグして移せます。</li>"
              "<li>メモの全文は、マスを選ぶと上の入力欄で読めます。「メモ全文」シートにも一覧があります。</li></ul></div>"
            + _ug_band("📅 作成した当直案")
            + "<div class='part'><div class='pt'>🔍 特定の医師の当直を色別でハイライト</div><ul><li>開くと、選んだ先生を色分けできます。</li>"
              "<li>表の名前にカーソルを合わせると、同じ先生の当直が強調されます。</li></ul></div>"
            + "<div class='part'><span class='gbtn'>📥 この当直案を掲示用のExcelでダウンロード（A3・印刷用）</span><div class='pw'>表のすぐ下の緑のボタン</div><ul>"
              "<li>病院に貼る形の表（A3横・1枚）です。ICU病棟（循環器）の列と右上の作成日は空けてあるので、ダウンロードしたあとに入れてください。</li>"
              "<li>交代などの手直しは、このExcelで直接して構いません。人数が足りない枠は「未決定」と書かれています。</li>"
              "<li>見えないシート「集計用」は年間の集計に使うので、消さないでください。</li></ul>"
              "<div class='warn'><b>手直しした掲示用Excelは、毎月手元に残してください。</b>翌月の確定当直（方法2）と、「📊 年間の当直回数」の集計に、そのまま読み込めます。</div></div>"
            + _ug_band("📊 年間の回数（今月込み）", tag="年間の当直回数を読み込んでいるときだけ")
            + "<div class='part'><ul><li>この当直案を入れた場合の、先生ごとの年間の回数と、平均との差（1か月あたり）です。回数の右の（ ）内は1か月あたりです。</li></ul></div>")
    if t_period:
        _ug("<div class='lead'>年末年始など、月をまたいで通しで当直を決めたいときに使います。今月の1日から、翌月の指定した日までを1つの当直案として作ります（例：12/1〜1/5）。</div>"
            + _ug_step(None, "期間指定版を開く", "通常版の、作る月を選ぶ欄の下の紫の枠",
                       ["「期間指定版を開く →」から移ります。ふだんの月は、期間指定版の青い枠の「通常版に戻る →」から、通常版で作ってください。"])
            + _ug_step(1, "作る期間を選ぶ", "期間指定版の「📅 年月・休日・必要人数の設定」",
                       ["年月（この月の1日から）と、翌月の何日まで含めるかを選びます。",
                        "カレンダーや表では、翌月の日は「1/1」「1/2」のように月も付けて出ます。"])
            + _ug_step(None, "あとは通常版と同じ", "",
                       ["休日・確定当直・医師条件・NG日・当直案の作成は、通常版と同じ手順です（左の目次の「毎月の流れ」）。"])
            + "<h4>通常版と違うところ</h4><ul>"
            "<li>医師条件の回数（月上限・月下限・休日上限・各枠の上限など）は、<b>期間全体での回数</b>として扱います。期間が長い分、必要に応じて多めにしてください。</li>"
            "<li>希望日は、翌月の日を「1/3」「1/3:A宿直」のように月も付けて書きます（今月の日は「30」「30:A宿直」のままで構いません）。</li>"
            "<li>「📋 前月の医師条件を引き継ぐ」は、<b>通常版で途中保存した医師条件</b>を読み込みます。</li>"
            "<li>途中保存は、通常版とは別の場所に残ります（新しいものから20回分）。</li>"
            "<li>先生方の専用URLと案内書類は、通常版と同じものを使います。期間指定版で受付を開始すると、先生方の画面が期間のカレンダーに切り替わります（通常版の受付は終わります）。</li></ul>")
    if t_trouble:
        _ug("<dl>"
            "<dt>「未決定」の枠が出る</dt><dd>作成ボタンの下の「計算状況・条件の注意事項と対処方法」と、④の赤い棒を確認します。NG日・回数の上限・勤務間隔が厳しすぎるかもしれません。月上限を増やす、勤務間隔を短くするなどして作り直してください。少しだけなら、Excelで手で埋めても構いません。</dd>"
            "<dt>アプリを開くと、入れた内容が消えている</dt><dd>しばらく使わないと画面がリセットされます。「📂 前回の作業の続きから始める」で読み込んでください。先生方の提出内容と専用URLは残っています。</dd>"
            "<dt>英語の画面（Zzzz…）が出る</dt><dd>アプリが休止しています。青いボタン「Yes, get this app back up!」を押して1分ほど待つと開きます。</dd>"
            "<dt>先生のNG日が当直案に入っていない</dt><dd>カレンダーで入れた場合は「NG日を反映する」、専用URLの場合は「⬇️ 選んだ先生のNG日を取り込む」を押したか確認してください。</dd>"
            "<dt>先生側の休日・別途決定の表示が違う</dt><dd>管理用パスワードを入れて、「締め切り日・休日設定を更新」を押してください。</dd>"
            "<dt>先生から「URLが無効」と言われた</dt><dd>URLを作り直していないか確認してください。案内書類をなくした場合は「URLを作り直す」で新しい案内書類を印刷します。</dd>"
            "<dt>医師条件の表の右の方の列が見えない</dt><dd>表は横にスクロールできます。表の上の「🗺 この表の列」で、右にどの列があるかを確認できます。</dd>"
            "<dt>年間の当直回数で、在籍月数が少なく出る先生がいる</dt><dd>見えないシート「集計用」がない掲示用Excel（シートを消したものや、少し前のアプリで作ったもの）では、その月に当直がなかった先生を在籍に数えられません。アプリでもう一度ダウンロードした掲示用Excelを使うと、正しく数えられます。</dd>"
            "<dt>確定当直に掲示用Excelを読み込んだら、作る月の分まで入った</dt><dd>作る月（または期間）の掲示用Excelを読み込んでいないか確認してください。確定当直に入れるのは、ふつうは前月の掲示用Excelです。</dd>"
            "<dt>アプリが起動しない・エラーが出る</dt><dd>表示されたエラーの内容を控えて、管理担当者に連絡してください。</dd></dl>")


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
# 閲覧用の画面：…/?mode=view（閲覧用パスワードで、医師条件だけを見られる）
if _mode == "view":
    staff_view_page()
    st.stop()
# 使い方：…/?mode=manual（別のタブで開く使い方のページ。パスワードなしで開ける）
if _mode == "manual":
    manual_page()
    st.stop()
if _mode == "period":
    _pp = pathlib.Path(__file__).resolve().parent / "period_page.py"
    if _pp.exists():
        exec(compile(_pp.read_text(encoding="utf-8"), str(_pp), "exec"), globals())
    else:
        st.error("期間指定版のファイル（period_page.py）が見つかりません。GitHubに period_page.py を入れてください。")
        st.markdown("<a href='./' target='_self'>← 通常版に戻る</a>", unsafe_allow_html=True)
    st.stop()

st.title("当直作成アプリ")
st.caption("年月・勤務条件・NG日を入力して、当直案を作成します。")
st.markdown(MANUAL_LINK_HTML, unsafe_allow_html=True)   # 使い方は別のタブで開く（アプリと並べて見ながら操作できる。入力中の内容は消えない）
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
#sec-staff,#sec-ng,#sec-cond,#sec-create,#sec-result{scroll-margin-top:5.5rem}
.sec-banner[id]{scroll-margin-top:5.5rem}
/* 折りたたみ式の項目（確定当直）も、同じ帯の見た目にする */
/* 「先生ごとの条件」の中の小見出し（①〜④） */
/* 番号の丸とタイトルは必ず横に並べる（幅がせまいときは、タイトルの中で折り返す）。説明はタイトルの下に出す */
.step-h{display:grid;grid-template-columns:auto minmax(0,1fr);align-items:center;gap:4px 10px;margin:1.8rem 0 .6rem;padding-bottom:6px;border-bottom:2px solid #d6e0ec}
.step-h .tt{min-width:0;line-height:1.4}
.step-h .ss{grid-column:2}
.step-h .no{display:inline-flex;align-items:center;justify-content:center;width:1.9rem;height:1.9rem;border-radius:50%;background:#24405f;color:#fff;font-weight:800;font-size:1rem;flex:0 0 auto}
.step-h .tt{font-size:1.2rem;font-weight:800;color:#24405f}
.step-h .ss{font-size:.85rem;color:#5a6a80}
@media (max-width:600px){.sec-banner{padding:11px 13px;gap:10px}.sec-banner .t{font-size:1.12rem !important}.sec-banner .s{display:none}.step-h .tt{font-size:1.05rem}}
</style>""", unsafe_allow_html=True)


def step_header(num, title, sub="", anchor=""):
    """「先生ごとの条件」の欄の中の、番号付きの小見出し。"""
    id_html = f' id="{anchor}"' if anchor else ""
    sub_html = f'<span class="ss">{sub}</span>' if sub else ""
    st.markdown(f'<div{id_html} class="step-h"><span class="no">{num}</span><span class="tt">{title}</span>{sub_html}</div>',
                unsafe_allow_html=True)


def section_banner(icon, title, sub="", tag="", first=False, anchor=""):
    tag_html = f'<span class="tag">{tag}</span>' if tag else ""
    sub_html = f'<div class="s">{sub}</div>' if sub else ""
    id_html = f' id="{anchor}"' if anchor else ""
    st.markdown(f'<div{id_html} class="sec-banner{" first" if first else ""}"><span class="ic">{icon}</span>'
                f'<div><div class="t">{title}</div>{sub_html}</div>{tag_html}</div>', unsafe_allow_html=True)


# ==========================================
# 1. 上部ダッシュボード：年月と休日の設定
# ==========================================
st.markdown(CSV_UPLOAD_CSS, unsafe_allow_html=True)   # CSVのアップロード欄・①の取り込み欄の見た目（画面全体で1回だけ）
section_banner("📅", "年月・休日・必要人数の設定", "作る月・休日と別途決定の日・1枠の人数を決めます", "必須", first=True)
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
        f"<div class='cb-box'><div class='cb-nav'><a href='#sec-cond' target='_self'>医師条件</a>"
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
step_header(1, "作る月を選ぶ")
year, month = month_picker(today)
ym_summary(year, month)
# 年末年始など、翌月にまたがって通しで作りたいときの入口（同じアプリの中で ?mode=period を開く）
try:
    _period_exists = (pathlib.Path(__file__).resolve().parent / "period_page.py").exists()
except Exception:
    _period_exists = False
if _period_exists:
    mode_switch_box("period")
# （期間指定版への入口ここまで）
restored_state = st.session_state.get("restored_state")

st.divider()

step_header(2, f"休日と別途決定の日の設定（{month}月）", "土日祝は自動で休日です。下のカレンダーで、平日でも休日にする日・別途決定の日にチェックします")
st.caption("**休日にする**：病院独自の休みなど、平日でも土日祝と同じ扱いにする日です。日直の枠ができ、日直・宿直とも休日回数に数えます。")
st.caption("**別途決定**：年末年始など、別の方法で当番を決める日です。このアプリでは当直を決めず、先生の専用URLの画面でも灰色になってNGを入力できません。確定当直に入れた分だけは当直案に表示されます。")
holiday_total_placeholder = st.empty()   # 🏥 埋める必要がある枠（カレンダーの上に出す。数は休日・増員の設定を読んだあとで入れる）

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
# 確定当直：いつも開いた状態で表示する（任意の項目）
with st.container(key="fixed_section"):
    section_banner("📅", "前月・今月・翌月の確定当直", "すでに決まっている当直や、前月末・翌月初めの当直を入れます", "任意", anchor="sec-fixed")
    st.info("前月・今月・翌月の、担当者が決まっている当直を入れます。入れ方は2つあります。**どちらか片方で大丈夫です**（両方使うこともできます）。"
            "前月末の勤務は、月初の勤務間隔の確認に使います。入力がなければ、この項目は飛ばせます。")
    with st.expander("入力例：確定当直の書き方と扱い", expanded=False):
        st.markdown("""
- 掲示用Excel（手直ししたものでも可）・CSVのアップロードと、下の表への直接入力のどれでも入力できます。掲示用ExcelのICU病棟の列は読みません。
- 表には、前月末7日分・今月・翌月はじめ7日分の日付が最初から並んでいます。「区分」の列に ◀前月／今月／翌月▶ を表示します。
- それ以外の日付（前月1か月分など）は、表の下の「＋」で行を足して `月/日`（例：`11/1`）または `年/月/日` の形で入力するか、ファイル（Excel・CSV）で読み込みます。ファイルの行は、表の同じ日付のところに入ります。
- 担当する先生は、枠の欄をクリックすると出る一覧（医師条件の名前）から選びます。
- 前月だけ勤務した先生など、医師条件にいない先生は、表の下の「選びたい先生が、名前の一覧に出てこないとき」を開いて名前を入れ、「表で選べるようにする」を押すと選べるようになります（前月・翌月の日付だけに使えます。当直案には入りません）。
- 同じ日・同じ枠を2人以上で担当する場合は、同じ日付の行を追加して、もう1人を選びます。ファイルで `佐藤、鈴木` のように「、」で区切って入れた場合は、読み込むと自動で2行に分かれます。
- 未確定の枠は空欄で構いません。
- 今月の平日に日直を入力する場合は、先に上のカレンダーでその日を「休日にする」にしてください。
- 「別途決定」にした日（年末年始など）の当番が決まったら、ここに入れておくと、当直案と一覧に表示され、前後の勤務間隔の計算にも使われます。
- 今月の確定当直は、NG日や翌日PM duty（宿直を外す曜日）よりも優先されます。回数の上限（月上限・休日上限・各枠の上限など）を超える場合も、確定した分はそのまま入ります。
- ほかの勤務は、確定当直から「宿⇔宿 間隔下限」（日直がからむ組み合わせは「宿⇔日・日⇔日 間隔下限」）以上あけて自動で入れます。確定当直同士の間隔が近すぎる場合は、確定内容をそのまま残し、注意を表示します。

**今月の同じ日・同じ枠を指定した場合、確定当直への入力と「希望優先度100以上」での指定は、どちらも固定され、計算上の扱いは基本的に同じです。両方に入力する必要はありません。**
        """)

    fixed_columns = ["日付"] + FIXED_SLOT_COLUMNS   # 当直案と同じ並び。平日か休日かは、上のカレンダーで自動的に決まるので入力しない
    fixed_template_df = pd.DataFrame(columns=fixed_columns)
    fixed_csv_template = table_xlsx(fixed_template_df, "確定当直", text_cols=fixed_columns, blank_rows=200)   # Excelのひな形（日付が勝手に変わらないよう「文字」の形）

    # 入れ方は2つ：方法1（下の表に直接入力）・方法2（ファイルから読み込む）。医師条件の取り込みと同じ形にする
    fx_a, fx_or, fx_b = st.columns([10, 1.4, 10])
    fx_or.markdown("<div class='imp-or'><span>または</span></div>", unsafe_allow_html=True)
    with fx_a, st.container(key="csvcard_fixedhand"):
        st.markdown("<div class='csv-card-head blue'><span class='lab'>方法1</span>✏️ 下の表に直接入力する</div>", unsafe_allow_html=True)
        st.caption(f"前後{FIXED_PAD_DAYS}日分の日付は最初から並んでいます。先生の欄をクリックして、名前を選ぶだけです。")
        with st.container(key="fixedhand_act"):
            st.markdown("<div class='fx-go'>⬇ 下の「✏️ 確定当直の表」へ<small>表の先生の欄をクリックすると、名前の一覧が出ます</small></div>",
                        unsafe_allow_html=True)
    with fx_b:
        fixed_file = csv_upload_card(
            "fixed", "方法2", "📂 ファイルから読み込む",
            "前月の掲示用Excel（手直ししたものでも可）か、CSVを読み込みます。読み込んだ内容は下の表に入ります。",
            f"fixed_csv_{st.session_state.get('uploader_gen', 0)}",
            fixed_csv_template, "確定当直_ひな形.xlsx", "📥 確定当直のひな形（Excel）をダウンロード",
            file_types=("csv", "xlsx"), drop_text="📄⬇　ここに掲示用ExcelかCSVファイルをドラッグ＆ドロップ")

    edited_fixed_df = fixed_duty_table(fixed_file, restored_state, fixed_columns, datetime.date(year, month, 1),
                                       datetime.date(year, month, calendar.monthrange(year, month)[1]), year, month,
                                       ("◀前月", "今月", "翌月▶"), f"{month}/1")

# ==========================================
# 先生ごとの条件：① 名簿と医師条件（見るだけ） → ② NG日 → ③ 医師条件の調整 → ④ 足りそう？ の順に、1つの欄にまとめて表示する。
# 計算の都合（NG日の欄は医師条件の表の名前を使う）で、プログラムの順番は表示の順番と違う。
# 先に入れ物（container）を表示の順に作っておき、あとから中身を入れる。
# ==========================================
section_banner("👩‍⚕️", "先生ごとの条件（NG日・回数など）", "① 先生の名簿と医師条件 → ② NG日 → ③ 医師条件の調整、の順に上から進めます", "必須", anchor="sec-staff")
staff_box_roster = st.container()
staff_box_ng = st.container()
staff_box_cond = st.container()
staff_box_cap = st.container()
staff_box_dl = st.container()

with staff_box_roster:
    step_header(1, "以前の医師条件の取り込み（引き継ぎ or ファイルのアップロード）、医師を追加（or外す）")
    st.markdown("<div class='imp-choose'>📥 以前の医師条件を取り込む（どちらか1つの方法を使います）</div>", unsafe_allow_html=True)
    imp_a, imp_or, imp_b = st.columns([10, 1.4, 10])   # 方法1（引き継ぎ）・または・方法2（CSV）
    imp_or.markdown("<div class='imp-or'><span>または</span></div>", unsafe_allow_html=True)
    imp_tmpl = st.container(key="imp_tmpl")   # 「初めての場合：ひな形をダウンロード」（中身はひな形を作ったあとで入れる）
    imp_ask = st.container()   # CSVにNG日・希望日が入っているときの「使う／空にする」と、読み込んだ結果（中身はCSVを読んだあとで入れる）
    # --- 方法1：途中保存から前月の医師条件を引き継ぐ（管理用パスワードが必要） ---
    with imp_a, st.container(key="csvcard_carry"):
        st.markdown("<div class='csv-card-head blue'><span class='lab'>方法1</span>途中保存から前月の医師条件を引き継ぐ</div>", unsafe_allow_html=True)
        if not remote_is_configured():
            st.caption("保存用のGoogleスプレッドシートの設定をすると使えます。今はファイルのアップロード（方法2）で読み込んでください。")
        else:
            _co_msg = st.session_state.pop("carry_over_message", None)
            # ボタンの名前：ログイン中は、途中保存してある月を出す（ふつうは前月）
            _co_name = "前月"
            if st.session_state.get("remote_admin_ok"):
                _co_at, _co_data = work_state_for_carry(year, month)
                if _co_data and _co_data.get("staff_csv"):
                    _co_ym = (int(_co_data.get("year", 0)), int(_co_data.get("month", 0)))
                    _co_prev = (year, month - 1) if month > 1 else (year - 1, 12)
                    _co_name = f"前月（{_co_ym[1]}月）" if _co_ym == _co_prev else f"{_co_ym[0]}年{_co_ym[1]}月"
            _co_label = f"📋 {_co_name}の医師条件を引き継ぐ" + "\u200b" * st.session_state.get("carry_over_gen", 0)
            st.caption("「💾 作業を途中保存する」で保存した医師条件（回数・勤務間隔・翌日PM duty・備考など）を読み込みます。前月分の保存があればその最新、なければいちばん新しい保存を使います。NG日・希望日・先生のコメントは空になります。")
            _co_act = st.container(key="carry_act")   # ボタンを入れる白い枠（方法2のドロップ枠と下でそろえる）
            with _co_act:
                _co_box = st.popover(_co_label) if hasattr(st, "popover") else st.expander(_co_label, expanded=False)
                if not st.session_state.get("remote_admin_ok"):
                    st.caption("押すと管理用パスワードの入力欄が開きます")
            with _co_box:
                if admin_logged_in("carry_over_pw"):
                    _co_at, _co_data = work_state_for_carry(year, month)
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
with staff_box_cond:
    step_header(3, "医師条件を調整する", "当直案の枠ができるだけ埋まるように、ここで医師条件を直します（直すと①の表示も同じ内容になります）", anchor="sec-cond")
    st.info("下の表に先生ごとの条件を入力してください。セルをクリックして編集できます。")
    st.caption("最初に表示される「型1」〜「型9」は、よくある条件の入力例です。名前の欄が型の中身です（例：型4＝A宿直・A日直・外来日直に入る先生）。実際の先生の名前に置き換え、回数などを調整して使ってください。行は「＋」で増やせます。")
    with st.expander("入力例：曜日・希望日・備考", expanded=False):
        st.markdown("""
| 項目 | 入力方法・意味 |
| --- | --- |
| 翌日PM duty | `水,木` のように半角カンマで区切ります。翌日PMにdutyがある曜日を入力します（例：木曜PMにdutyがある場合は「水」）。原則としてその曜日の宿直を外しますが、翌日が休日なら宿直に入る場合があります。日直は対象外です。 |
| 希望日 | `10,15` はその日のいずれかの枠、`10:A宿直` はその枠を希望します。複数の希望は半角カンマで区切ります。日直は休日（または「休日にする」にした日）だけ指定できます。「別途決定」の日の希望は使いません。 |
| 指定できる枠 | A宿直・B宿直・外来宿直・A日直・B日直・外来日直。表記を一致させてください。 |
| 備考 | 管理用のメモです。「学会」などと書いても計算条件には反映されません。休みはNG日で指定してください。 |

曜日にかかわらず勤務できない日は、上の「② NG日を入れる」のカレンダーでNGにします。
平日は「宿直」のボタン、休日は「日直」「宿直」のボタンを押してNGにします（休日に終日入れない場合は両方押します）。
    """)
    with st.expander("回数・勤務間隔の数え方", expanded=False):
        st.markdown("""
| 項目 | 意味・入力例 |
| --- | --- |
| 宿⇔宿 間隔下限 | 宿直と次の宿直の間に空ける日数です。5日なら、10日の宿直の次の宿直は16日以降です。 |
| 宿⇔日・日⇔日 間隔下限 | 日直がからむ組み合わせ（日直→宿直、宿直→日直、日直→日直）で空ける日数です。2日なら、10日の次は13日以降です。宿直どうしより負担が小さいので、宿⇔宿 間隔下限より短くできます。空欄なら宿⇔宿 間隔下限と同じです。 |
| 月下限 | できるだけ確保したい回数です。条件によっては、この回数に届かないことがあります。 |
| 月上限 | 日直と宿直を合わせた月間の上限です。 |
| 休日上限 | 土日祝日と「休日にする」にした日に担当する、日直と宿直の合計上限です。 |
| 宿直上限・日直上限 | A・B・外来を合わせた宿直（日直）の月間の上限です。空欄なら制限しません。 |
| 各枠の上限（A宿上限〜外日上限） | A宿直など、それぞれの枠を担当する月間の上限です。空欄なら、できるだけ2回まで（ほかの先生で埋まらないときだけ3回以上も担当します）。表では「A宿上限」「外日上限」のように短く表示します。 |

1つの枠を1回と数えます。確定当直で同じ日に日直と宿直を担当する場合は2回です。
月下限は、月上限以下に設定してください。
確定当直や希望優先度100以上の希望がある場合は、上限を超えることがあります。ほかの勤務は確定当直から「宿⇔宿 間隔下限」（日直がからむ組み合わせは「宿⇔日・日⇔日 間隔下限」）以上あけて入れ、確定当直同士が近すぎる場合は注意を表示します。
    """)
    with st.expander("希望優先度：通常の希望と、100以上の特別な設定", expanded=False):
        st.markdown("""
- **通常は「1」**を使用します。1〜99は、数字が大きいほど希望を優先しますが、NG日・回数上限・勤務間隔などの範囲内で割り当てます。
- **100以上は、その先生の希望日すべてを確定当直と同じ扱いにする設定**です。「できれば入りたい」という用途には使わないでください。
- 確定にした日は、NG日や翌日PM duty（宿直を外す曜日）よりも優先されます。回数の上限を超える場合も、確定した分はそのまま入ります。ほかの勤務は、確定の日から「宿⇔宿 間隔下限」（日直がからむ組み合わせは「宿⇔日・日⇔日 間隔下限」）以上あけて入れます。確定当直同士の間隔が近すぎる場合は、注意を表示します。
- 一部の勤務だけを確定させたい場合は、上の「📅 前月・今月・翌月の確定当直」へ入力し、希望優先度は通常の値にしてください。
- 指定の誤りや条件の組み合わせによっては作成できない場合があります。作成後に、確定にした日が当直案に入っているか確認してください。

**今月の同じ日・同じ枠を指定した場合、確定当直への入力と「希望優先度100以上」での指定は、どちらも固定され、計算上の扱いは基本的に同じです。両方に入力する必要はありません。**
    """)

with staff_box_roster:
    staff_box_addrm = st.container()   # 「➕ 先生を追加する」「➖ 先生を外す」（中身は医師条件の表を読んだあとで入れる）
    # 入力例：よくある9つの型（名前の欄に型の中身が分かる名前を入れる）。実際の先生に置き換えて使う
    # 型2（A日直だけ・希望日を確定）の希望日は、選んでいる月の最初の休日2日にする（平日の日直は指定できないため）
    # （「休日にする」のチェックで例の表が作り直されないよう、土日祝だけで決める）
    _tmpl_hols = [d for d in range(1, calendar.monthrange(year, month)[1] + 1)
                  if datetime.date(year, month, d).weekday() >= 5 or jpholiday.is_holiday(datetime.date(year, month, d))][:2]
    _tmpl_wish2 = ",".join(f"{d}:A日直" for d in _tmpl_hols)
    _tmpl_rows = [
        # 名前, 備考, 希望日, 希望優先度, 宿⇔宿 間隔下限, 宿⇔日・日⇔日 間隔下限, 月下限, 月上限, 休日上限, 宿直上限, 日直上限, A宿直, B宿直, 外来宿直, A日直, B日直, 外来日直
        ("型1_A宿直・A日直", "Aの宿直・日直だけ。月1回", "", 1, 5, 2, 1, 1, 1, 3, 3, 1, 0, 0, 1, 0, 0),
        ("型2_A日直のみ・希望日確定", "決めた日のA日直だけ（希望優先度100で確定）", _tmpl_wish2, 100, 0, 0, 0, 2, 2, 3, 3, 0, 0, 0, 2, 0, 0),
        ("型3_外来日直のみ", "外来日直だけ。月1回まで", "", 1, 5, 2, 0, 1, 4, 3, 3, 0, 0, 0, 0, 0, 1),
        ("型4_A宿直・A日直・外来日直", "Aの宿直・日直と外来日直。月2〜3回", "", 1, 5, 2, 2, 3, 1, 3, 3, 3, 0, 0, 1, 0, 1),
        ("型5_A・B宿直・A日直・外来日直", "型4にB宿直を追加。月2〜3回", "", 1, 5, 2, 2, 3, 1, 3, 3, 3, 3, 0, 1, 0, 1),
        ("型6_外来宿直以外すべて", "型5にB日直を追加。月3〜4回", "", 1, 5, 2, 3, 4, 2, 3, 3, 3, 3, 0, 1, 1, 2),
        ("型7_全枠", "6枠すべて。月3〜4回", "", 1, 5, 2, 3, 4, 2, 3, 3, 3, 3, 3, 1, 1, 2),
        ("型8_日直のみ", "A・B・外来の日直だけ。月1〜3回", "", 1, 4, 2, 1, 3, 3, 3, 3, 0, 0, 0, 3, 3, 3),
        ("型9_A以外", "B・外来の宿直と日直。月3〜4回", "", 1, 4, 2, 3, 4, 2, 3, 3, 0, 3, 3, 0, 3, 3),
    ]
    _tmpl_cols = ["先生の名前", "備考（メモ・説明など自由記入）", REQUEST_COLUMN, "希望優先度(数字が大きいほど優先)",
                  "宿⇔宿 間隔下限", DAY_GAP_COLUMN, "月下限", "月上限", "休日上限", "宿直上限", "日直上限",
                  "A宿直上限", "B宿直上限", "外来宿直上限", "A日直上限", "B日直上限", "外来日直上限"]
    template_data = {c: [r[k] for r in _tmpl_rows] for k, c in enumerate(_tmpl_cols)}
    template_data["翌日PM duty"] = [""] * len(_tmpl_rows)
    template_data[NG_COLUMN] = [""] * len(_tmpl_rows)
    # 列の並びは、これまでのひな形と同じにする
    template_data = {c: template_data[c] for c in ["先生の名前", "備考（メモ・説明など自由記入）", "翌日PM duty", NG_COLUMN, REQUEST_COLUMN,
                                                   "希望優先度(数字が大きいほど優先)", "宿⇔宿 間隔下限", DAY_GAP_COLUMN, "月下限", "月上限",
                                                   "休日上限", "宿直上限", "日直上限", "A宿直上限", "B宿直上限", "外来宿直上限",
                                                   "A日直上限", "B日直上限", "外来日直上限"]}
    df_template = pd.DataFrame(template_data)
    # 「➕ 先生を追加する」で型から選べるように、型の中身を覚えておく
    st.session_state["staff_templates"] = {r["先生の名前"]: r for r in df_template.to_dict("records")}
    # ひな形はExcel。数字の列以外は「文字」の形にして、NG日・希望日の「11/3」などが勝手に日付に変わらないようにする
    xlsx_template = table_xlsx(df_template, "医師条件",
                               text_cols=[c for c in df_template.columns if c not in NUMERIC_DEFAULTS and c not in GROUP_CAP_COLUMNS and c != DAY_GAP_COLUMN],
                               blank_rows=60)

    # --- 方法2：CSVファイルをアップロード ---
    with imp_b:
        uploaded_file = csv_upload_card(
            "staff", "方法2", "ファイルをアップロード（Excel・CSV）",
            "手元のExcel（以前ダウンロードした医師条件、またはひな形に入力したもの）を読み込みます。CSVファイルも読み込めます。NG日・希望日が入っているときは、使うかどうかを選べます。",
            f"staff_csv_{st.session_state.get('uploader_gen', 0)}_{st.session_state.get('staff_uploader_gen', 0)}",
            None, "", "", file_types=("xlsx", "csv"), drop_text="📄⬇　ここにExcelかCSVファイルをドラッグ＆ドロップ")
    with imp_tmpl:
        st.markdown("📝 初めての場合：")
        st.download_button(label="📥 ひな形（Excel）をダウンロード", data=xlsx_template, file_name="医師条件_ひな形.xlsx", mime=XLSX_MIME)
        st.markdown("して、Excelで入力してから方法2で読み込みます。")
    staff_box_ro = st.container()   # ①のいちばん下：読み込んだ医師条件（見るだけ。中身は医師条件の表を読んだあとで入れる）

with staff_box_cond:
    if uploaded_file is not None:
        if st.session_state.get('last_uploaded_file_id') != uploaded_file.file_id:
            for key in list(st.session_state.keys()):
                if key.startswith("ng_"):
                    del st.session_state[key]
            st.session_state['last_uploaded_file_id'] = uploaded_file.file_id
            st.session_state["staff_source_label"] = f"ファイル「{uploaded_file.name}」から"
            _gr = group_rules_from_xlsx(uploaded_file.getvalue())
            if _gr:
                set_group_rules_base(_gr)   # 医師条件のExcelに「グループの設定」のシートがあれば、その設定にする
            st.session_state.pop("staff_added_names", None)
        try:
            base_df = parse_staff_csv(uploaded_file.getvalue())
        except Exception as e:
            st.error(f"医師条件のファイルを読み込めませんでした。詳細: {e}")
            st.stop()
        base_df = csv_ng_wish_choice(base_df, uploaded_file, imp_ask, year, month)
    elif restored_state and restored_state.get("staff_csv"):
        base_df = parse_staff_csv(restored_state["staff_csv"].encode("utf-8"))
    else:
        base_df = df_template.copy()
    staff_is_template = uploaded_file is None and not (restored_state and restored_state.get("staff_csv"))

    if "先生の名前" not in base_df.columns:
        st.error("医師条件のファイルに「先生の名前」の列がありません。ひな形の列名を確認してください。")
        st.stop()
    # 宿直上限・日直上限の列がない（以前の）CSVでも、「休日上限」の右に空欄の列を足す（空欄＝制限なし）
    for _gc in reversed(list(GROUP_CAP_COLUMNS)):
        if _gc not in base_df.columns:
            _pos = list(base_df.columns).index("休日上限") + 1 if "休日上限" in base_df.columns else len(base_df.columns)
            base_df.insert(_pos, _gc, None)
        base_df[_gc] = pd.to_numeric(base_df[_gc], errors="coerce")
    # 宿⇔日・日⇔日 間隔下限の列がない（以前の）CSV・途中保存でも、「宿⇔宿 間隔下限」の右に、宿⇔宿 間隔下限と同じ値で列を足す
    base_df = add_day_gap_column(base_df.loc[:, ~base_df.columns.duplicated()])   # 同じ名前の列が2つあると表が出せないので1つにする
    base_df = add_group_column(base_df.drop(columns=[DOC_GROUP_AVOID_COLUMN], errors="ignore"))   # 「グループ」の列（ないファイルでは空欄で足す）
    base_df = add_pref_columns(base_df)   # 「多め希望」の3つの列（ないファイルでは空欄で足す）
    if DAY_GAP_COLUMN in base_df.columns:
        base_df[DAY_GAP_COLUMN] = pd.to_numeric(base_df[DAY_GAP_COLUMN], errors="coerce")
    base_df = base_df.set_index("先生の名前")

    if "希望優先度(数字が大きいほど優先)" in base_df.columns:
        base_df["希望優先度(数字が大きいほど優先)"] = pd.to_numeric(base_df["希望優先度(数字が大きいほど優先)"], errors='coerce')

    text_cols = ["翌日PM duty", NG_COLUMN, REQUEST_COLUMN, "備考（メモ・説明など自由記入）"]
    for c in text_cols:
        if c in base_df.columns:
            base_df[c] = base_df[c].apply(lambda x: "" if pd.isna(x) or str(x).lower() in ["nan", "none", "<na>"] else str(x))

    group_rules = group_rules_ui(restored_state, year, month, staff_is_template)   # 👥 グループの設定（任意）
    st.caption("先生の名前は、ほかの先生と重ならない表記にしてください。")

    edited_df = staff_table(base_df, editable=True, comments=imported_comments(year, month))   # 先生からのコメントは見るだけで出す

    staff_df = edited_df.reset_index()
    # 先生の追加・外すは、この表（今の画面の内容）をもとに行う
    st.session_state["staff_current_csv"] = staff_df.to_csv(index=False)

    # 確定当直の表で選べる先生の名前を、医師条件に合わせる（名前が変わったときは、確定当直の表の一覧も作り直す）
    _names_now = list(dict.fromkeys(n for n in (clean_text(v) for v in staff_df.get("先生の名前", pd.Series(dtype=str)).tolist())
                                    if n and n.lower() not in ("nan", "none")))
    if st.session_state.get("fixed_name_options") != _names_now:
        st.session_state["fixed_name_options"] = _names_now
        commit_fixed_current()   # 確定当直の表に入力した分を残したまま、名前の一覧を新しくして出し直す
        if not (create_bar_request or work_save_request):
            st.rerun()
    # 「➕ 先生を追加する」「➖ 先生を外す」（名前が変わったときの再表示のあとに出す。知らせが消えないように）
    with staff_box_addrm:
        staff_add_remove_ui(staff_df, year, month, is_template=staff_is_template)
    with staff_box_ro:
        staff_readonly_table(staff_df, comments=imported_comments(year, month))

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
    staff_df_calc = apply_group_rules(staff_df, group_rules)   # 計算用：グループの×の枠は上限0、△の枠は「できれば入れない」
    group_orphan_notice(staff_df, group_rules)   # 消したグループに入っている先生がいれば知らせる
    if not staff_input_errors:
        group_mismatch_notice(staff_df, group_rules)   # グループ内で、ほかの先生と違う設定があれば知らせる
    if staff_input_errors:
        show_input_errors(staff_input_errors)
        if create_bar_request:
            create_bar_show("<div class='cb-res bad'>⚠ 医師条件の入力に誤りがあります。赤い表示を確認してください。</div>")
        st.stop()

with staff_box_cap:
    step_header(4, "当直の枠は足りそう？（目安）", "条件を直すと、すぐに変わります")
    st.caption("埋める必要がある枠の数と、先生方が入れる回数の合計を比べています。先生ごとの回数は、「月上限」と各枠の上限・休日上限のうち、"
               "いちばん少ないものを使います。足りていても、NG日・勤務間隔などによっては埋まらないことがあります。"
               "確定当直で追加される枠や上限の例外は、この目安に含まれません。")

    st.markdown(CAPACITY_CSS, unsafe_allow_html=True)
    cap_needs = slot_needs(year, month, num_days, custom_holidays, locked_days, multi_slots_dict)
    cap_rows, total_max_capacity = capacity_rows(staff_df_calc, cap_needs)
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
        st.error(f"🚨 **全体で{-margin}枠足りません。** 先生方の「月上限」や各枠の上限を増やすか、担当できる先生を追加してください。")
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
                 + "該当する先生の" + "・".join(dict.fromkeys(f"「{staff_column_specs().get(r[3], (r[3],))[0]}」" for r in cap_short))
                 + "を増やすか、担当できる先生を追加してください（「月上限」が小さいと、上限を増やしても回数は増えません）。")
with staff_box_ng:
    step_header(2, "NG日を入れる", "先生を選んで、カレンダーでNG日を指定します。専用URLで集めたNG日の取り込みもここで行います", anchor="sec-ng")
    st.info("先生を選んでカレンダーの「日直」「宿直」のボタンを押してNGにし、最後に「NG日を反映する」を押してください。"
            "先生方に専用URLから入力してもらう場合は、下の「📨 先生方からのNG日入力（専用URL）」を使います。")
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
        # 紙で配る「当直不可日」の用紙（専用URLのすぐ下）
        paper_ng_form_ui(f"当直不可日（{year}年{month}月）",
                         [(datetime.date(year, month, d), bool(day_is_holiday.get(d)), d in set(locked_days), str(d)) for d in range(1, num_days + 1)],
                         f"当直不可日_{year}年{month}月.xlsx")
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

                    pm_duty_setting_ui(doc_name, doc_row.get("翌日PM duty", ""), ng_is_dirty(doc_name), year, month)   # 翌日PM dutyもここで設定できる

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


with staff_box_dl:
    st.divider()
    st.markdown("### 📂 医師条件をExcelでダウンロード（手元にファイルで残す場合）")
    st.write("先生の名前・回数・勤務間隔・希望日・NG日・備考をExcelファイルとしてパソコンなどにダウンロードします。このファイルは「医師条件」のアップロード欄から読み込めます。")
    st.caption("ダウンロードの前に、各先生の「NG日を反映する」を押してください。作る年月・休日と別途決定の日・増員設定・確定当直・作成した当直案・色分けとそのメモは、このファイルには含まれません。"
               "続きから作業するだけなら、ファイルではなく「💾 途中保存」を使うほうが簡単です。")

    # ダウンロードするExcelも、画面と同じく「先生の名前 → 備考 → その他」の順にそろえる（読み込みは列名で行うので、古い並びのファイルもそのまま使える）
    _first = ["先生の名前"] + [c for c in staff_df.columns if str(c).startswith("備考")]
    current_xlsx = staff_xlsx_with_groups(keep_blank_slot_caps(staff_df)[_first + [c for c in staff_df.columns if c not in _first and not str(c).startswith("_")]],
                                          [c for c in staff_df.columns if c not in NUMERIC_DEFAULTS and c not in GROUP_CAP_COLUMNS and c != DAY_GAP_COLUMN],
                                          st.session_state.get("group_rules_current"))
    st.download_button(
        label="📥 医師条件をExcelでダウンロード",
        data=current_xlsx,
        file_name=f"医師条件_{year}年{month}月.xlsx",
        mime=XLSX_MIME,
        use_container_width=True
    )


# ==========================================
# 年間の当直回数（任意）：掲示用Excelから先生ごとの回数を集計する（見るだけ。当直案の計算には使わない）
# ==========================================
_annual_names = [n for n in (clean_text(v) for v in staff_df.get("先生の名前", pd.Series(dtype=str)).tolist()) if n and n.lower() not in ("nan", "none")]
annual_balance = annual_section(year, month, datetime.date(year, month, 1), datetime.date(year, month, calendar.monthrange(year, month)[1]), 1.0,
                                restored_state, _annual_names, now_word="今月")

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
st.info("全員の条件を満たす当直案が作れない場合に、どの枠から優先して埋めるかを決めます。数字が小さいほど優先し、同じ数字の枠は同順位として扱います。")
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
    "staff_csv": keep_blank_slot_caps(staff_df).to_csv(index=False),   # 枠ごとの上限の空欄は空欄のまま保存する
    "priority": {k: int(v) for k, v in shortage_priority.items()},
    "comments": {n: list(v) for n, v in imported_comments(year, month).items()},
    "alerts": {k: int(st.session_state.get(f"alert_{k}", v) or 0) for k, v in MATRIX_COUNT_ALERTS.items()},
    "annual": annual_snapshot(year, month),   # 年間の当直回数（月ごとの回数だけ）
    "group_rules": st.session_state.get("group_rules_current", default_group_rules()),   # 👥 グループの設定（メモを含む）
    "paper_notes": st.session_state.get(PAPER_NOTES_KEY, PAPER_NOTES_DEFAULT),   # 紙で配る当直不可日の用紙の注意書き
}
# 右上の途中保存ボタンが押されていたら、ここで（全部の入力を読み終えてから）保存する
if work_save_request:
    _slot = WORK_SAVE_SLOT or st.empty()
    _slot.markdown("<span class='work-save-note busy fresh'>💾 途中保存しています…</span>", unsafe_allow_html=True)
    try:
        work_state_save(work_snapshot)
        st.session_state["work_save_note"] = remote_now()[-8:-3]
        st.session_state["work_restore_pick_newest"] = True   # 「前回の作業の続きから始める」で、今保存したものを選んだ状態にする
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
current_signature = (input_signature(year, month, staff_df_calc, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority), tuple(sorted(locked_days)),
                     json.dumps(annual_balance, ensure_ascii=False, sort_keys=True))
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
        st.caption("※ 押すたびに、同じ条件でも違うパターンの当直案を作ります。気に入った案は、結果の緑のボタン（医師別の当直一覧のExcelと、掲示用のExcel）でダウンロードしてください（作成した当直案は途中保存には含まれません）。")
    if create_clicked:
        with st.spinner("当直案を計算中…（通常は最大60秒、不足枠の確認を含む場合は計算時間が最大75秒です）"):
            try:
                df_result, success, error_reasons, past_worked_dates, future_worked_dates = generate_shift(year, month, staff_df_calc, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority, locked_days, annual_balance=annual_balance)

                if df_result is not None:
                    st.session_state['generated_signature'] = current_signature
                    st.session_state['generated_year'] = year
                    st.session_state['generated_month'] = month
                    st.session_state['generated_warnings'] = error_reasons
                    st.session_state['generated_at'] = remote_now()   # 掲示用Excelの作成日時（年間集計で、あとから作ったほうを使うため）
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
                   "希望日が通った当直は茶色の太い枠で囲み、回数が多い先生は右端の回数を赤い枠で示します。マスにカーソルを合わせると、その先生の行とその日の列がオレンジの線で囲まれます。")
        # 前月末・翌月初めの確定当直から「宿⇔宿 間隔下限」（日直がからむときは「宿⇔日・日⇔日 間隔下限」）以内で、入れない日
        # （今月の確定当直の前後は含めない）。日ごとに (日直に入れない, 宿直に入れない)
        gap_blocked = external_gap_blocked(staff_df, past_worked_dates, future_worked_dates, num_days,
                                           lambda d: datetime.date(year, month, d))
        matrix_days, matrix_rows = build_duty_matrix(df_result, staff_df, year, month, custom_holidays, next_month_special_holiday,
                                                     comments=imported_comments(year, month), gap_blocked=gap_blocked)
        missing_total = shortage_total(matrix_days)
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

        # アラートの基準は、折りたたまずに常に表示する（控えめに：枠なし・上下の細い線だけ）
        st.markdown("<style>.st-key-alert_box{border-top:1px solid #e3e7ee;border-bottom:1px solid #e3e7ee;padding:12px 4px 14px;gap:.5rem;margin:6px 0 10px}"
                    ".alert-h{font-weight:700;color:#31333f;font-size:1rem}"
                    ".alert-d{font-size:.84rem;color:#6a7686;line-height:1.6;margin:4px 0 6px}"
                    ".st-key-alert_box [data-testid='stNumberInput'] input{padding-left:.6rem}</style>", unsafe_allow_html=True)
        with st.container(key="alert_box"):
            st.markdown("<div class='alert-h'>⚠ アラートの基準（回数が多い枠を赤く表示）</div>"
                        "<div class='alert-d'>医師別の当直一覧の右端の回数が、この数以上になった枠を赤く囲みます。0にした項目はアラートを出しません。</div>",
                        unsafe_allow_html=True)
            _spec = [1] * len(MATRIX_COUNT_ALERTS) + [0.5, 1.2]
            try:
                _acols = st.columns(_spec, vertical_alignment="bottom")
            except TypeError:   # 古いStreamlitで vertical_alignment が使えない場合
                _acols = st.columns(_spec)
            for _col, _k in zip(_acols, MATRIX_COUNT_ALERTS):
                _col.number_input(f"{MATRIX_ALERT_LABELS[_k]}（回以上）", min_value=0, max_value=31, step=1, key=f"alert_{_k}")
            _acols[-1].button("初期設定に戻す", key="alert_reset", on_click=_reset_alerts,
                              help="宿直4回以上・休日3回以上に戻します（ほかは0＝アラートなし）")
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

            # 掲示用のExcel（病院に貼る形）。ICU病棟（循環器）の列と作成日は空けておき、あとから手で入れる
            # 医師別の当直一覧のExcelボタンと同じ緑のボタンにする
            st.markdown("<style>.st-key-posting_download button{background:#1d6f42 !important;border:1px solid #1d6f42 !important;color:#fff !important;font-weight:700}"
                        ".st-key-posting_download button:hover{background:#155a35 !important;border-color:#155a35 !important}"
                        ".st-key-posting_download button p{color:#fff !important}</style>", unsafe_allow_html=True)
            with st.container(key="posting_download"):
                st.download_button(
                    label="📥 この当直案を掲示用のExcelでダウンロード（A3・印刷用）",
                    data=posting_xlsx(df_result, year, month, f"内科当直表（{year}年{month}月）",
                                      roster=doctors_list, created=st.session_state.get("generated_at")),
                    file_name=f"当直表_掲示用_{year}年{month}月.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            st.caption("ICU病棟（循環器）の列と、右上の作成日は空けてあります。ダウンロードしたあとに入れてください。")
            # 年間の当直回数を読み込んでいるときは、この当直案を入れた場合の年間の回数も出す
            annual_result_view(df_result, year, month, doctors_list, None, now_word="今月")


else:
    st.warning("☝️ 表に先生の名前を入力するか、医師条件のファイルをアップロードしてください。")
