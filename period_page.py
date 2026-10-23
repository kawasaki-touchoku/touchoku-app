# -*- coding: utf-8 -*-
"""
期間指定版の当直作成（年末年始など）。app.py から「?mode=period」で開かれる。
今月の1日から、翌月の指定した日までを1つの当直表として作る。

通常版（app.py）の計算・表示の部品をそのまま使い、通常版の画面には手を入れない。
入力内容のキー（st.session_state）は「p_」で始め、通常版の入力と混ざらないようにする。
"""
import ast
import pathlib

import streamlit as st

# ------------------------------------------------------------
# 通常版（app.py）の「部品」（関数・定数）を使う。
# app.py の中から呼ばれたとき（?mode=period）は、すでに読み込まれているのでそのまま使う。
# 単独で動かしたときだけ、app.py から部品を読み込む。
# ------------------------------------------------------------
if "TouchokuPeriod" not in globals():
    _APP_PATH = pathlib.Path(__file__).resolve().parent / "app.py"
    _tree = ast.parse(_APP_PATH.read_text(encoding="utf-8"))
    _lib = []
    for _node in _tree.body:
        if isinstance(_node, ast.Expr) and "set_page_config" in ast.unparse(_node):
            break
        _lib.append(_node)
    exec(compile(ast.Module(_lib, []), str(_APP_PATH), "exec"), globals())
    st.set_page_config(page_title="当直作成（期間指定）", layout="wide")

# 途中保存は通常版と別の場所に保存する
REMOTE_SHEETS["saved_state_period"] = ["項目", "値"]

st.markdown("""
<style>
.sec-banner{display:flex;align-items:center;gap:14px;margin:2.2rem 0 .9rem;padding:14px 18px;border-radius:12px;
  background:linear-gradient(90deg,#5b3a1e 0%,#a0662f 100%);color:#fff;box-shadow:0 2px 6px rgba(60,40,20,.2)}
.sec-banner .ic{font-size:1.6rem;line-height:1}
.sec-banner .t{font-size:1.3rem;font-weight:800;letter-spacing:.02em;line-height:1.3}
.sec-banner .s{font-size:.88rem;opacity:.92;margin-top:2px}
.sec-banner.first{margin-top:.4rem}
.p-cal-head{text-align:center;font-weight:800;padding:2px 0}
.st-key-p_calendar [data-testid="stVerticalBlockBorderWrapper"]{min-height:120px}
@media (max-width:600px){.sec-banner{padding:11px 13px}.sec-banner .t{font-size:1.1rem}.sec-banner .s{display:none}}
</style>""", unsafe_allow_html=True)


def banner(icon, title, sub="", first=False):
    sub_html = f'<div class="s">{sub}</div>' if sub else ""
    st.markdown(f'<div class="sec-banner{" first" if first else ""}"><span class="ic">{icon}</span>'
                f'<div><div class="t">{title}</div>{sub_html}</div></div>', unsafe_allow_html=True)


# ==========================================
# 途中保存（期間指定版専用の保存場所）
# ==========================================
def p_state_save(snapshot):
    text = json.dumps(snapshot, ensure_ascii=False)
    chunks = [text[i:i + WORK_STATE_CHUNK] for i in range(0, len(text), WORK_STATE_CHUNK)] or [""]
    rows = [["保存日時", remote_now()]] + [["内容" if i == 0 else f"内容{i + 1}", c] for i, c in enumerate(chunks)]
    remote_rewrite("saved_state_period", rows)


def p_state_info():
    try:
        rows = {r.get("項目", ""): r.get("値", "") for r in remote_read("saved_state_period")}
    except Exception:
        return None, None
    if not rows.get("内容"):
        return None, None
    text = rows["内容"]
    n = 2
    while rows.get(f"内容{n}"):
        text += rows[f"内容{n}"]
        n += 1
    try:
        return rows.get("保存日時", ""), json.loads(text)
    except ValueError:
        return None, None


def p_restore():
    """期間指定版で途中保存した作業を戻す（ボタンのコールバック）。"""
    saved_at, data = p_state_info()
    if not data:
        st.session_state["p_message"] = ("error", "保存された作業が見つかりませんでした。")
        return
    for k in list(st.session_state.keys()):
        if k.startswith(("p_hol_", "p_lock_", "p_ng_", "p_next_special_", "p_prio_", "p_comment_", "p_import_")):
            del st.session_state[k]
    st.session_state["p_year"] = int(data["year"])
    st.session_state["p_month"] = int(data["month"])
    st.session_state["p_extra"] = int(data.get("extra", 0))
    P0 = TouchokuPeriod(data["year"], data["month"], data.get("extra", 0))
    for i in data.get("holidays", []):
        st.session_state[f"p_hol_{P0.key}_{int(i)}"] = True
    for i in data.get("locked", []):
        st.session_state[f"p_lock_{P0.key}_{int(i)}"] = True
    st.session_state[f"p_next_special_{P0.key}"] = bool(data.get("next_special"))
    for s, rank in (data.get("priority") or {}).items():
        st.session_state[f"p_prio_{s}"] = int(rank)
    for name, cv in (data.get("comments") or {}).items():
        if cv:
            st.session_state[f"p_comment_{P0.key}_{name}"] = (str(cv[0]), str(cv[1]) if len(cv) > 1 else "")
    st.session_state["p_restored"] = data
    st.session_state["p_uploader_gen"] = st.session_state.get("p_uploader_gen", 0) + 1
    st.session_state["p_editor_gen"] = st.session_state.get("p_editor_gen", 0) + 1
    st.session_state.pop("p_result", None)
    st.session_state["p_message"] = ("success", f"{saved_at} に保存した作業（{P0.title()}）を読み込みました。")


def p_carry_over():
    """通常版で途中保存した医師条件を読み込む（NG日・希望日は空にする）。"""
    saved_at, data = work_state_info()
    if not data or not data.get("staff_csv"):
        st.session_state["p_message"] = ("error", "通常版で途中保存した医師条件が見つかりませんでした。")
        return
    df = parse_staff_csv(str(data["staff_csv"]).encode("utf-8"))
    for col in (NG_COLUMN, REQUEST_COLUMN):
        if col in df.columns:
            df[col] = ""
    for k in list(st.session_state.keys()):
        if k.startswith(("p_ng_", "p_comment_", "p_import_")):
            del st.session_state[k]
    restored = dict(st.session_state.get("p_restored") or {})
    restored["staff_csv"] = df.to_csv(index=False)
    st.session_state["p_restored"] = restored
    st.session_state["p_uploader_gen"] = st.session_state.get("p_uploader_gen", 0) + 1
    st.session_state["p_editor_gen"] = st.session_state.get("p_editor_gen", 0) + 1
    st.session_state["p_message"] = ("success", f"通常版で{saved_at}に途中保存した医師条件（{len(df)}名）を読み込みました。"
                                                "NG日・希望日は空にしてあります。")


# ==========================================
# 画面
# ==========================================
st.title("📆 当直作成（期間指定・年末年始など）")
st.caption("今月の1日から、翌月の指定した日までを1つの当直表として作ります。年末年始など、月をまたいで通しで決めたいときに使います。"
           "ふだんの月の当直は、通常版で作ってください。")
st.markdown("<a href='./' target='_self' style='font-weight:700'>🏠 ← 通常版の当直作成に戻る</a>", unsafe_allow_html=True)

msg = st.session_state.pop("p_message", None)
if msg:
    (st.success if msg[0] == "success" else st.error)(msg[1])

configured = remote_is_configured()
if configured:
    with st.expander("💾 途中保存・📂 続きから・👩‍⚕️ 通常版の医師条件を読み込む（管理用パスワードが必要）", expanded=False):
        if admin_logged_in("p_pw"):
            c1, c2, c3 = st.columns(3)
            p_save_clicked = c1.button("💾 作業を途中保存する", key="p_save", type="primary", use_container_width=True)
            _at, _saved = p_state_info()
            c2.button("📂 前回の作業の続きから始める", key="p_restore", on_click=p_restore, use_container_width=True,
                      disabled=not _saved)
            c3.button("👩‍⚕️ 通常版の医師条件を読み込む", key="p_carry", on_click=p_carry_over, use_container_width=True)
            st.caption((f"期間指定版の途中保存：{_at}（{TouchokuPeriod(_saved['year'], _saved['month'], _saved.get('extra', 0)).title()}）"
                        if _saved else "期間指定版の途中保存はまだありません。")
                       + "　／　通常版の医師条件を読み込むと、NG日・希望日は空になります（回数・間隔・備考などはそのまま）。")
        else:
            p_save_clicked = False
else:
    p_save_clicked = False

# ------------------------------------------
# 1. 期間・休日
# ------------------------------------------
banner("📅", "期間・休日・必要人数の設定", "作る期間（今月1日〜翌月の何日まで）と、休日・別途決定の日を決めます", first=True)
today = datetime.date.today()
default_y, default_m = (today.year + 1, 1) if today.month == 12 else (today.year, today.month + 1)
if "p_year" not in st.session_state:
    st.session_state["p_year"] = default_y
    st.session_state["p_month"] = default_m
    st.session_state["p_extra"] = 5
cy, cm, ce = st.columns(3)
year = int(cy.number_input("年", min_value=2000, step=1, key="p_year"))
month = int(cm.number_input("月（この月の1日から）", min_value=1, max_value=12, step=1, key="p_month"))
_next_first = datetime.date(year, month, calendar.monthrange(year, month)[1]) + datetime.timedelta(days=1)
extra = int(ce.number_input(f"翌月（{_next_first.month}月）の何日まで含めるか", min_value=0,
                            max_value=calendar.monthrange(_next_first.year, _next_first.month)[1], step=1, key="p_extra"))
P = TouchokuPeriod(year, month, extra)
st.info(f"**作る期間：{P.title()}（{P.n}日間）**　この期間全体を1つの当直表として作ります。"
        "医師条件の回数（月間最大回数・休日最大回数・各枠の上限など）は、この期間全体での回数として扱います。")

st.markdown("**休日にする**：平日でも土日祝と同じ扱いにする日。**別途決定**：このページでは当直を決めず、先生もNGを入力できない日。")
custom_holidays, locked_days = [], []
with st.container(key="p_calendar"):
    head = st.columns(7)
    for i, w in enumerate(WEEKDAYS_JA):
        color = "#c0392b" if i == 6 else ("#1f6fbf" if i == 5 else "#243247")
        head[i].markdown(f"<div class='p-cal-head' style='color:{color}'>{w}</div>", unsafe_allow_html=True)
    cells = [None] * P.start.weekday() + list(range(1, P.n + 1))
    cells += [None] * ((7 - len(cells) % 7) % 7)
    for wk in range(0, len(cells), 7):
        cols = st.columns(7)
        for col, i in zip(cols, cells[wk:wk + 7]):
            if i is None:
                continue
            dt = P.date(i)
            auto_hol = dt.weekday() >= 5 or jpholiday.is_holiday(dt)
            with col.container(border=True):
                color = "#c0392b" if (dt.weekday() == 6 or jpholiday.is_holiday(dt)) else ("#1f6fbf" if dt.weekday() == 5 else "#243247")
                st.markdown(f"<div style='text-align:center;font-weight:800;color:{color}'>{P.label(i)}</div>", unsafe_allow_html=True)
                if auto_hol:
                    st.markdown("<div style='text-align:center;color:#c0392b;font-size:.8rem'>休</div>", unsafe_allow_html=True)
                elif st.checkbox("休日にする", key=f"p_hol_{P.key}_{i}"):
                    custom_holidays.append(i)
                if st.checkbox("別途決定", key=f"p_lock_{P.key}_{i}"):
                    locked_days.append(i)

_after = P.end + datetime.timedelta(days=1)
_after_label = f"{_after.month}/{_after.day}（{WEEKDAYS_JA[_after.weekday()]}）"
if _after.weekday() >= 5 or jpholiday.is_holiday(_after):
    next_special = False
    st.caption(f"期間の翌日 {_after_label} は休日なので、最終日 {P.label(P.n)} の宿直は「翌日PM duty」の制限を受けません。")
else:
    next_special = st.checkbox(f"期間の翌日 {_after_label} は病院独自の休み", key=f"p_next_special_{P.key}",
                               help="チェックすると、最終日の宿直は「翌日PM duty」の制限を受けません。")

day_is_holiday = {i: period_is_holiday(P, i, custom_holidays) for i in range(1, P.n + 1)}
labels = [P.label(i) for i in range(1, P.n + 1)]
label_to_index = {lb: i for i, lb in enumerate(labels, 1)}

st.markdown("##### 👥 1つの枠を2名以上にする設定（任意）")
_restored = st.session_state.get("p_restored") or {}
_same_period = _restored and TouchokuPeriod(_restored.get("year", year), _restored.get("month", month), _restored.get("extra", 0)).key == P.key
multi_df0 = pd.DataFrame(_restored.get("multi", []) if _same_period else [], columns=["日付", "当直枠", "人数"])
multi_df = st.data_editor(
    multi_df0, num_rows="dynamic", use_container_width=True, hide_index=True, height=140,
    key=f"p_multi_{P.key}_{st.session_state.get('p_editor_gen', 0)}",
    column_config={
        "日付": st.column_config.SelectboxColumn("日付", options=labels),
        "当直枠": st.column_config.SelectboxColumn("増員する枠", options=NIGHT_SHIFTS + DAY_SHIFTS),
        "人数": st.column_config.NumberColumn("合計人数", min_value=2, max_value=10, step=1),
    })
multi_slots = {}
for _, r in multi_df.iterrows():
    lb, sl, cnt = clean_text(r.get("日付")), clean_text(r.get("当直枠")), r.get("人数")
    if lb in label_to_index and sl in ALL_SHIFT_TYPES and pd.notna(cnt):
        multi_slots[(label_to_index[lb], sl)] = int(cnt)

# ------------------------------------------
# 2. 確定当直
# ------------------------------------------
banner("📌", "確定当直（任意）", "すでに決まっている当直や、期間の前後の当直を入れます（勤務間隔の計算に使います）")
fixed_cols = ["日付"] + ALL_SHIFT_TYPES
if _same_period and _restored.get("fixed_csv"):
    fixed0 = parse_fixed_csv(_restored["fixed_csv"].encode("utf-8"))
else:
    fixed0 = pd.DataFrame([[""] * len(fixed_cols)], columns=fixed_cols)
fixed0 = fixed0.drop(columns=["平日/休日"], errors="ignore")
for c in fixed_cols:
    if c not in fixed0.columns:
        fixed0[c] = ""
fixed0 = fixed0[fixed_cols].apply(lambda col: col.map(lambda x: "" if pd.isna(x) or str(x).strip().lower() in ("nan", "none") else str(x)))
st.caption("日付は「12/30」「1/2」のように入力します。担当する枠の欄に先生の名前を入れます（複数は「、」で区切る）。")
fixed_df = st.data_editor(fixed0, num_rows="dynamic", use_container_width=True, hide_index=True, height=180,
                          key=f"p_fixed_{P.key}_{st.session_state.get('p_editor_gen', 0)}",
                          column_config={c: st.column_config.TextColumn(c) for c in fixed_cols})

# ------------------------------------------
# 3. 医師条件
# ------------------------------------------
banner("👩‍⚕️", "医師条件", "通常版と同じ条件の表です。回数は期間全体での回数として扱います")
st.caption("通常版で途中保存した医師条件は、上の「👩‍⚕️ 通常版の医師条件を読み込む」で読み込めます。CSVのアップロードもできます。"
           "希望日は「12/30」「1/2:A宿直」のように、月も入れて書きます。")
up = st.file_uploader("医師条件（CSV）をアップロード", type="csv", key=f"p_staff_csv_{st.session_state.get('p_uploader_gen', 0)}")
if up is not None:
    if st.session_state.get("p_last_upload") != up.file_id:
        for k in list(st.session_state.keys()):
            if k.startswith(("p_ng_", "p_comment_", "p_import_")):
                del st.session_state[k]
        st.session_state["p_last_upload"] = up.file_id
    base_df = parse_staff_csv(up.getvalue())
elif _restored.get("staff_csv"):
    base_df = parse_staff_csv(_restored["staff_csv"].encode("utf-8"))
else:
    base_df = pd.DataFrame(columns=["先生の名前", "備考（メモ・説明など自由記入）", "翌日PM duty", NG_COLUMN, REQUEST_COLUMN]
                           + list(NUMERIC_DEFAULTS) + list(GROUP_CAP_COLUMNS))
if "先生の名前" not in base_df.columns:
    st.error("医師条件CSVに「先生の名前」の列がありません。")
    st.stop()
for c in ["備考（メモ・説明など自由記入）", "翌日PM duty", NG_COLUMN, REQUEST_COLUMN]:
    if c not in base_df.columns:
        base_df[c] = ""
    base_df[c] = base_df[c].map(lambda x: "" if pd.isna(x) or str(x).lower() in ("nan", "none", "<na>") else str(x))
for c in GROUP_CAP_COLUMNS:
    if c not in base_df.columns:
        base_df[c] = None
    base_df[c] = pd.to_numeric(base_df[c], errors="coerce")
staff_df = st.data_editor(
    base_df, num_rows="dynamic", use_container_width=True, hide_index=True, height=320,
    key=f"p_staff_{P.key}_{st.session_state.get('p_editor_gen', 0)}",
    column_config={
        NG_COLUMN: None,
        REQUEST_COLUMN: st.column_config.TextColumn("希望日", help="例：12/30 または 1/2:A宿直。複数は半角カンマ区切り。"),
        "希望優先度(数字が大きいほど優先)": st.column_config.NumberColumn("希望優先度"),
        "備考（メモ・説明など自由記入）": st.column_config.TextColumn("備考"),
    },
    column_order=[c for c in base_df.columns if str(c).startswith("備考")] + [c for c in base_df.columns if not str(c).startswith("備考")],
)
staff_df = staff_df[staff_df["先生の名前"].map(clean_text) != ""].reset_index(drop=True)
doctor_names = [clean_text(n) for n in staff_df["先生の名前"]]

# ------------------------------------------
# 4. NG日（先生方の専用URL・管理者の入力）
# ------------------------------------------
banner("🚫", "先生ごとのNG日", "先生方に専用URLから入力してもらい、取り込みます。管理者が直接入力することもできます")


def ng_key(name, i):
    return f"p_ng_{P.key}_{name}_{i}"


# 医師条件の表にNG日が入っていれば、最初だけ読み込む
for _, r in staff_df.iterrows():
    nm = clean_text(r["先生の名前"])
    if not any(ng_key(nm, i) in st.session_state for i in range(1, P.n + 1)):
        vals = period_parse_ng(r.get(NG_COLUMN, ""), P)
        for i in range(1, P.n + 1):
            st.session_state[ng_key(nm, i)] = vals.get(i, "OK")


def p_current_ng(name):
    return {i: st.session_state.get(ng_key(name, i), "OK") for i in range(1, P.n + 1)
            if st.session_state.get(ng_key(name, i), "OK") != "OK"}


def p_import(submissions, names):
    for name in names:
        r = submissions.get(name)
        if not r:
            continue
        submitted = period_parse_ng(r.get("NG日", ""), P)
        undo = f"p_import_undo_{P.key}_{name}"
        if undo not in st.session_state:
            st.session_state[undo] = p_current_ng(name)
        for i in range(1, P.n + 1):
            st.session_state[ng_key(name, i)] = "OK" if i in locked_days else submitted.get(i, "OK")
        comment = clean_text(r.get("コメント", ""))
        if comment:
            st.session_state[f"p_comment_{P.key}_{name}"] = (comment, r.get("保存日時", ""))
        else:
            st.session_state.pop(f"p_comment_{P.key}_{name}", None)
    st.session_state["p_message"] = ("success", f"{len(names)}名分のNG日を取り込みました。")


def p_undo_import(name):
    before = st.session_state.pop(f"p_import_undo_{P.key}_{name}", None)
    if before is None:
        return
    for i in range(1, P.n + 1):
        st.session_state[ng_key(name, i)] = before.get(i, "OK")
    st.session_state.pop(f"p_comment_{P.key}_{name}", None)
    st.session_state["p_message"] = ("success", f"{name} 先生の取り込みを取り消しました。")


if configured:
    with st.container(border=True):
        st.markdown("**📨 先生方からのNG日入力（専用URL）**　🔒 管理者専用")
        if admin_logged_in("p_pw2"):
            try:
                settings = remote_settings()
            except Exception as e:
                st.error(f"スプレッドシートを読み込めませんでした。詳細：{e}")
                settings = {}
            active = settings.get("受付中の年月", "")
            closed = settings.get("締め切り", "") == "True"
            if active == P.key:
                st.markdown(f"**{P.title()}：{'🔒 締め切り済み' if closed else '🟢 受付中'}**（締め切り日の表示：{settings.get('締め切り日', '') or '未設定'}）")
            elif active:
                _ap = TouchokuPeriod.from_key(active) if str(active).startswith(PERIOD_KEY_PREFIX) else None
                st.warning(f"現在の受付は「{_ap.title() if _ap else str(active).replace('-', '年') + '月'}」です。"
                           f"ここで受付を開始すると、その受付は終わり、先生方の画面は{P.title()}のカレンダーに切り替わります。")
            else:
                st.markdown("**現在、受付していません。**")
            st.caption("先生方の専用URL・案内書類は、通常版と同じものをそのまま使えます（配り直しは不要です）。")
            deadline = st.text_input("先生方に表示する締め切り日（例：12月10日）",
                                     value=settings.get("締め切り日", "") if active == P.key else "", key="p_deadline")
            hol_text = ",".join(map(str, sorted(custom_holidays)))
            lock_text = ",".join(map(str, sorted(locked_days)))
            b1, b2, b3 = st.columns(3)
            if b1.button(f"{P.title()}の受付を開始する", key="p_open", use_container_width=True):
                remote_sync_doctors([(n, clean_text(r.get("翌日PM duty", ""))) for n, (_, r) in zip(doctor_names, staff_df.iterrows())])
                remote_save_settings(**{"受付中の年月": P.key, "特別休日": hol_text, "翌月1日特別休日": bool(next_special),
                                        "締め切り": False, "締め切り日": deadline, "入力不可日": lock_text})
                st.rerun()
            if active == P.key:
                if not closed and b2.button("🔒 締め切る", key="p_close", use_container_width=True):
                    remote_save_settings(**{"締め切り": True})
                    st.rerun()
                if closed and b2.button("受付を再開する", key="p_reopen", use_container_width=True):
                    remote_save_settings(**{"締め切り": False})
                    st.rerun()
                if b3.button("締め切り日・休日設定を更新", key="p_update", use_container_width=True):
                    remote_save_settings(**{"特別休日": hol_text, "翌月1日特別休日": bool(next_special),
                                            "締め切り日": deadline, "入力不可日": lock_text})
                    st.rerun()
                if (settings.get("特別休日", "") != hol_text or settings.get("入力不可日", "") != lock_text):
                    st.warning("上のカレンダーの「休日にする」「別途決定」が、先生側の画面と違います。「締め切り日・休日設定を更新」を押してください。")

                if st.button("🔄 最新の状態に更新", key="p_refresh"):
                    remote_clear_cache()
                    st.rerun()
                submissions = remote_latest_submissions(P.key)
                status = []
                for n in doctor_names:
                    r = submissions.get(n)
                    status.append({"先生": n, "状況": "✅ 提出済み" if r else "⏳ 未提出",
                                   "最終保存": r.get("保存日時", "") if r else "",
                                   "NG日": (period_ng_summary(r.get("NG日", ""), P, custom_holidays) or "NG日なし") if r else "",
                                   "コメント": r.get("コメント", "") if r else ""})
                st.caption(f"提出済み {sum(1 for x in status if x['状況'].startswith('✅'))} / {len(status)} 名")
                st.dataframe(pd.DataFrame(status), hide_index=True, use_container_width=True)
                importable = [n for n in doctor_names if n in submissions]
                if importable:
                    sel_key = "p_import_sel_" + hashlib.sha256("|".join(importable).encode()).hexdigest()[:12]
                    targets = [n for n in (st.multiselect("取り込む先生（×で外せます）", importable, default=importable, key=sel_key) or [])
                               if n in importable]
                    st.button(f"⬇️ 選んだ先生のNG日を取り込む（{len(targets)}名分）", key="p_import", type="primary",
                              disabled=not targets, on_click=p_import, args=(submissions, targets))
                    st.caption("取り込むと、その先生のNG日は提出内容に置き換わります。")
                undoable = [n for n in doctor_names if f"p_import_undo_{P.key}_{n}" in st.session_state]
                if undoable:
                    u1, u2 = st.columns([3, 2])
                    ut = u1.selectbox("取り込みを取り消す先生", undoable, key="p_undo_target")
                    u2.button("↩️ 取り込みを取り消す", key="p_undo", on_click=p_undo_import, args=(ut,), use_container_width=True)
                if submissions and st.checkbox("先生の提出をリセットする", key="p_reset_open"):
                    rs_key = "p_reset_sel_" + hashlib.sha256("|".join(submissions).encode()).hexdigest()[:12]
                    rs = [n for n in (st.multiselect("リセットする先生", list(submissions), key=rs_key) or []) if n in submissions]
                    if rs and st.checkbox(f"選んだ{len(rs)}名の提出をリセットしてよい", key="p_reset_confirm"):
                        if st.button(f"🗑️ {len(rs)}名の提出をリセットする", key="p_reset_do"):
                            for n in rs:
                                remote_reset_submission(P.key, n)
                            st.rerun()

st.markdown("##### ✍️ 管理者が直接入力する")
if doctor_names:
    sel = st.selectbox("NG日を入力する先生", doctor_names, key=f"p_ng_doctor_{P.key}")
    srow = staff_df.iloc[doctor_names.index(sel)]
    pm_weekdays = pm_duty_weekdays(srow.get("翌日PM duty", ""))
    cm_ = st.session_state.get(f"p_comment_{P.key}_{sel}")
    if cm_:
        st.info(f"💬 先生からのコメント（{cm_[1]}）：{cm_[0]}")
    vals = {i: st.session_state.get(ng_key(sel, i), "OK") for i in range(1, P.n + 1)}
    days = period_ng_calendar_days(P, custom_holidays, next_special, pm_weekdays, vals, set(locked_days))
    revision = hashlib.sha256(json.dumps([P.key, sel, days], ensure_ascii=False).encode()).hexdigest()
    comp_key = f"p_ng_editor_{P.key}_{sel}"
    response = horizontal_ng_component(HORIZONTAL_NG_HTML)(
        days=days, mode="month", offset=P.start.weekday(), doctor=sel, version=revision,
        key=comp_key, default=None, report_dirty=True, bulk_buttons=True)
    seen = comp_key + "_last"
    if isinstance(response, dict) and response.get("kind") == "apply" and response.get("token") != st.session_state.get(seen):
        st.session_state[seen] = response.get("token")
        new_values = response.get("values")
        if response.get("version") == revision and isinstance(new_values, list) and len(new_values) == len(days):
            changed = False
            for i, (v, item) in enumerate(zip(new_values, days), 1):
                if v in item["options"] and v != item["value"] and i not in locked_days:
                    st.session_state[ng_key(sel, i)] = v
                    changed = True
            if changed:
                st.rerun()
    if p_current_ng(sel):
        st.success("反映済みのNG日：" + period_ng_summary(period_ng_text(p_current_ng(sel), P), P, custom_holidays))

# NG日を医師条件に書き込む（計算・途中保存に使う）
staff_df = staff_df.copy()
staff_df[NG_COLUMN] = [period_ng_text({i: v for i, v in p_current_ng(n).items() if i not in locked_days}, P) for n in doctor_names]

# ------------------------------------------
# 5. 作成
# ------------------------------------------
banner("🎯", "当直案の作成・確認", "人数が足りないときに優先して埋める枠を確認して、当直案を作ります")
prio = {}
pcols = st.columns(len(ALL_SHIFT_TYPES))
for col, s in zip(pcols, NIGHT_SHIFTS + DAY_SHIFTS):
    if f"p_prio_{s}" not in st.session_state:
        st.session_state[f"p_prio_{s}"] = DEFAULT_SHORTAGE_PRIORITY[s]
    prio[s] = col.selectbox(s, list(range(1, len(ALL_SHIFT_TYPES) + 1)), key=f"p_prio_{s}")
st.caption("数字が小さい枠ほど、人数が足りないときに優先して埋めます。")

comments = {k[len(f"p_comment_{P.key}_"):]: v for k, v in st.session_state.items() if k.startswith(f"p_comment_{P.key}_")}

snapshot = {
    "year": year, "month": month, "extra": extra,
    "holidays": sorted(custom_holidays), "locked": sorted(locked_days), "next_special": bool(next_special),
    "multi": [[P.label(i), s, int(c)] for (i, s), c in multi_slots.items()],
    "fixed_csv": fixed_df.to_csv(index=False), "staff_csv": staff_df.to_csv(index=False),
    "priority": {k: int(v) for k, v in prio.items()},
    "comments": {n: list(v) for n, v in comments.items()},
}
if p_save_clicked:
    try:
        p_state_save(snapshot)
        st.success(f"途中保存しました（{remote_now()}）。")
    except Exception as e:
        st.error(f"途中保存できませんでした。詳細：{e}")

signature = hashlib.sha256(json.dumps(snapshot, ensure_ascii=False, sort_keys=True).encode()).hexdigest()
if st.button(f"🚀 {P.title()}の当直案を作成", key="p_create", type="primary"):
    with st.spinner("当直案を計算中…（最大75秒）"):
        try:
            res = period_generate_shift(P, staff_df, custom_holidays, multi_slots, fixed_df, next_special, prio, locked_days)
            st.session_state["p_result"] = (signature, P.key, res)
        except Exception as e:
            st.session_state.pop("p_result", None)
            st.error(f"当直計算中にエラーが発生しました。詳細：{e}")

stored = st.session_state.get("p_result")
if stored and stored[1] == P.key:
    if stored[0] != signature:
        st.warning("条件が変わりました。もう一度「当直案を作成」を押してください。下の結果は前回の条件のものです。")
    df_result, success, messages, past, future = stored[2]
    if df_result is None:
        show_schedule_notice(None, messages)
    else:
        matrix_days, matrix_rows = period_build_duty_matrix(df_result, staff_df, P, custom_holidays, next_special, comments=comments)
        alerts = count_alert_messages(matrix_rows, MATRIX_COUNT_ALERTS)
        show_schedule_notice(df_result, messages, alerts)
        st.subheader("🗓️ 医師別の当直一覧")
        st.caption("回数（右端）は、この期間全体での回数です。")
        st.markdown(duty_matrix_html(matrix_days, matrix_rows, year, month, MATRIX_COUNT_ALERTS), unsafe_allow_html=True)
        try:
            st.download_button("📥 この一覧をExcelでダウンロード（印刷用）",
                               data=duty_matrix_xlsx(matrix_days, matrix_rows, year, month, MATRIX_COUNT_ALERTS, title=P.title()),
                               file_name=f"当直一覧_{P.start.isoformat()}_{P.end.isoformat()}.xlsx",
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key="p_xlsx")
        except ImportError:
            st.error("Excelでダウンロードするには openpyxl が必要です。")
        st.subheader("📅 作成した当直案（日付ごと）")
        st.dataframe(df_result, hide_index=True, use_container_width=True)
        st.download_button("📥 当直案をCSVでダウンロード", df_result.to_csv(index=False).encode("utf-8-sig"),
                           file_name=f"当直案_{P.start.isoformat()}_{P.end.isoformat()}.csv", mime="text/csv", key="p_csv")
