# -*- coding: utf-8 -*-
"""
期間指定版（年末年始など）の当直作成。app.py から「?mode=period」で開かれる。
通常版（app.py）の画面と同じ作りで、「今月1日〜翌月の指定日」を1つの当直案として作る。

【このファイルは、app.py から自動で作ったものです】
通常版の画面をそのまま使い、日付の扱いだけを「期間」に読み替えています。
・日付は「期間の通し番号」（1＝今月1日、…）で扱う。今月の日は日付と同じ番号、翌月の日は月末の続きの番号。
・datetime.date / calendar.monthrange は、対象の年月のときだけ期間を指すように差し替える。
・休日・NG日・希望日・計算・表の作成は、期間用の関数に差し替える。
"""
import ast
import pathlib
import types as _types

import streamlit as st

if "TouchokuPeriod" not in globals():
    # 単独で動かしたとき（ふつうは app.py の中から呼ばれる）
    _APP_PATH = pathlib.Path(__file__).resolve().parent / "app.py"
    _tree = ast.parse(_APP_PATH.read_text(encoding="utf-8"))
    _lib = []
    for _node in _tree.body:
        if isinstance(_node, ast.Expr) and "set_page_config" in ast.unparse(_node):
            break
        _lib.append(_node)
    exec(compile(ast.Module(_lib, []), str(_APP_PATH), "exec"), globals())
    st.set_page_config(page_title="当直作成（期間指定）", layout="wide")

# ------------------------------------------------------------
# 期間への読み替え
# ------------------------------------------------------------
_REAL_CALENDAR = calendar
_REAL_DATETIME = datetime


def _install_period(period):
    """datetime.date と calendar を、対象の年月のときだけ期間を指すものに差し替える（このページの中だけ）。"""
    global datetime, calendar, P
    P = period
    real_last = _REAL_MONTHRANGE(period.year, period.month)[1]

    class _Meta(type):
        def __instancecheck__(cls, obj):
            return isinstance(obj, _REAL_DATE)

    class _PeriodDate(_REAL_DATE, metaclass=_Meta):
        def __new__(cls, y, m=None, d=None):
            if m is None:
                return _REAL_DATE.__new__(cls, y)
            if (int(y), int(m)) == (period.year, period.month) and int(d) > real_last:
                return period.date(int(d))
            return _REAL_DATE(int(y), int(m), int(d))

    def _monthrange(y, m):
        if (int(y), int(m)) == (period.year, period.month):
            return (period.start.weekday(), period.n)
        return _REAL_MONTHRANGE(y, m)

    def _monthcalendar(y, m):
        if (int(y), int(m)) != (period.year, period.month):
            return _REAL_CALENDAR.monthcalendar(y, m)
        cells = [0] * period.start.weekday() + list(range(1, period.n + 1))
        cells += [0] * ((7 - len(cells) % 7) % 7)
        return [cells[i:i + 7] for i in range(0, len(cells), 7)]

    dt_mod = _types.SimpleNamespace(date=_PeriodDate, datetime=_REAL_DATETIME.datetime, timedelta=_REAL_DATETIME.timedelta,
                                    timezone=_REAL_DATETIME.timezone, time=_REAL_DATETIME.time)
    cal_mod = _types.SimpleNamespace(**{k: getattr(_REAL_CALENDAR, k) for k in dir(_REAL_CALENDAR) if not k.startswith("__")})
    cal_mod.monthrange = _monthrange
    cal_mod.monthcalendar = _monthcalendar
    datetime, calendar = dt_mod, cal_mod


def _day_label(d):
    """今月の日は「5日」、翌月の日は「1/3」と表示する。"""
    dt = P.date(d)
    return f"{dt.day}日" if dt.month == P.month else f"{dt.month}/{dt.day}"


# 医師条件の表の列の説明を、期間指定版の言い方にする（「月間」→「期間全体」、希望日の書き方）
_normal_staff_column_specs = staff_column_specs


def staff_column_specs():
    specs = dict(_normal_staff_column_specs())
    for k, (label, width, kind, help_text) in list(specs.items()):
        if help_text:
            help_text = help_text.replace("月間の上限", "期間全体での上限").replace("できるだけ確保したい回数", "期間全体で、できるだけ確保したい回数")
            if k == REQUEST_COLUMN:
                help_text = help_text.replace("例：10,15 または 10:A宿直。", "例：10,15 または 10:A宿直。翌月の日は 1/3 や 1/3:A宿直 のように月も書きます。")
        specs[k] = (label, width, kind, help_text)
    return specs


_normal_staff_legend_text = staff_legend_text


def staff_legend_text():
    return _normal_staff_legend_text().replace("月の", "期間全体の")




# 途中保存は通常版と別のシートに保存する（通常版の途中保存とは混ざらない）
REMOTE_SHEETS["saved_state_period"] = ["項目", "値"]
REMOTE_SHEETS["saved_history_period"] = list(REMOTE_SHEETS["saved_history"])
_NORMAL_WORK_SHEETS = dict(WORK_STATE_SHEETS)
WORK_STATE_SHEETS = {"legacy": "saved_state_period", "history": "saved_history_period"}


def _normal_work_state_for_carry(year, month):
    """医師条件の引き継ぎは、通常版で途中保存した分から読む。"""
    return work_state_for_carry(year, month, _NORMAL_WORK_SHEETS)


# 日付・休日・NG日・希望日・計算を、期間用に差し替える（引数の年月は使わない）
def is_holiday_date(date_obj, target_year, target_month, custom_holidays):
    return period_is_holiday_date(P, date_obj, custom_holidays)


def pm_duty_restricted(date_obj, weekdays, target_year, target_month, custom_holidays, next_month_special_holiday=False):
    nd = date_obj + _REAL_DATETIME.timedelta(days=1)
    next_is_holiday = period_is_holiday_date(P, nd, custom_holidays) or (
        next_month_special_holiday and nd == P.end + _REAL_DATETIME.timedelta(days=1))
    return date_obj.weekday() in weekdays and not next_is_holiday


def checked_day_items(value, year, month, ng=False):
    return period_checked_items(value, P, ng)


def parse_ng_dict(value, year, month):
    return period_parse_ng(value, P)


def parse_requests(value, year, month):
    return period_parse_requests(value, P)


def ng_dict_to_text(values_by_day):
    return period_ng_text(values_by_day, P)


def ng_text_summary(text, year, month, holidays):
    return period_ng_summary(text, P, holidays)


def ng_calendar_days(year, month, holidays, next_month_special_holiday, pm_weekdays, values_by_day, locked_days=()):
    return period_ng_calendar_days(P, holidays, next_month_special_holiday, pm_weekdays, values_by_day, locked_days)


def validate_staff_inputs(df, year, month):
    return period_validate_staff_inputs(df, P)


def validate_fixed_inputs(df, doctors, year, month, holidays=None):
    return period_validate_fixed_inputs(df, doctors, P, holidays)


def generate_shift(target_year, target_month, staff_df, custom_holidays, multi_slots_dict, fixed_df=None, next_month_special_holiday=False, shortage_priority=None, locked_days=()):
    return period_generate_shift(P, staff_df, custom_holidays, multi_slots_dict, fixed_df, next_month_special_holiday, shortage_priority, locked_days)


def build_duty_matrix(df_result, staff_df, year, month, custom_holidays, next_month_special_holiday=False, comments=None, gap_blocked=None):
    return period_build_duty_matrix(df_result, staff_df, P, custom_holidays, next_month_special_holiday, comments, gap_blocked)


_normal_duty_matrix_xlsx = duty_matrix_xlsx


def duty_matrix_xlsx(days, rows, year, month, alerts=None, title=None):
    return _normal_duty_matrix_xlsx(days, rows, year, month, alerts, title=P.title())



# ---- 管理画面（受付）・途中保存の読み込み・医師条件の引き継ぎ（期間用） ----
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
                if _st.get("受付中の年月", "") == P.key and (
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

        st.caption("流れ：① 受付を開始する → ② URLの案内書類を配る（初回のみ）→ ③ 提出状況を確認 → ④ 締め切る → ⑤ 提出されたNG日を取り込む")
        month_key = P.key
        active = settings.get("受付中の年月", "")
        closed = settings.get("締め切り", "") == "True"
        if active == month_key:
            status = "🔒 締め切り済み" if closed else "🟢 受付中"
            st.markdown(f"**{P.title()}：{status}**（締め切り日の表示：{settings.get('締め切り日', '') or '未設定'}）")
        elif str(active).startswith(PERIOD_KEY_PREFIX):
            _ap = TouchokuPeriod.from_key(active)
            st.warning(f"現在は、別の期間「{_ap.title() if _ap else active}」の受付をしています。"
                       f"ここで{P.title()}の受付を開始すると、その受付は終わります。")
        elif active:
            st.warning(f"現在は、通常版で「{active.replace('-', '年')}月」の受付をしています。上で選んでいる{P.title()}とは異なります。ここで受付を開始すると、先生方の画面はこの期間のカレンダーに切り替わります。")
        else:
            st.markdown("**現在、受付していません。**")

        # --- 受付の開始・締め切り ---
        deadline = st.text_input("先生方に表示する締め切り日（例：10月25日）", value=settings.get("締め切り日", "") if active == month_key else "", key="remote_deadline")
        _n_days = calendar.monthrange(year, month)[1]
        _saved_locked = [d for d in parse_holiday_list(settings.get("入力不可日", "")) if d <= _n_days] if active == month_key else []
        locked_sel = sorted(locked_days)
        locked_text = ",".join(map(str, locked_sel))
        if locked_sel:
            st.caption("先生がNGを入力できない日（上のカレンダーの「別途決定」）：" + "・".join(P.label(d) for d in locked_sel if 1 <= d <= P.n))
        c1, c2, c3 = st.columns(3)
        if c1.button(f"{P.title()}の受付を開始する", key="remote_open", use_container_width=True,
                     help="上で設定した期間・休日・別途決定の日で受付を始めます。先生の名簿と専用URLも更新します。"):
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
        st.markdown(f"##### 📋 {P.title()}の提出状況")
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
    st.session_state["sel_extra"] = int(data.get("extra", 0))
    _tp = TouchokuPeriod(y, m, int(data.get("extra", 0)))
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
    st.session_state.pop("fixed_src", None)   # 確定当直の表も、読み込んだ内容で作り直す
    st.session_state["multi_editor_gen"] = st.session_state.get("multi_editor_gen", 0) + 1   # 増員設定の表を読み込んだ内容で作り直す
    # ファイルのアップロード欄を空にして、復元した内容を使う
    st.session_state["uploader_gen"] = st.session_state.get("uploader_gen", 0) + 1
    st.session_state.pop("last_uploaded_file_id", None)
    st.session_state["work_state_message"] = ("success", f"{saved_at} に保存した作業（{_tp.title()}）を読み込みました。")
    st.session_state["staff_source_label"] = f"{str(saved_at)[5:16]} の途中保存から"
    st.session_state.pop("staff_added_names", None)
    # 読み込んだことは、上の帯のボタンの右に1か所だけ出す（次に途中保存するまで表示しておく）
    st.session_state["work_restore_note"] = f"{str(saved_at)[5:16]} 保存分（{_tp.title()}）を読み込みました"
    st.session_state.pop("work_save_note", None)
    st.session_state["work_note_fresh"] = True   # スマートフォンでは、読み込んだ直後に画面上部へ数秒だけ表示する
    # 読み込んだあとは「続きから」の小さな枠を閉じる（見出しを目に見えない形で変えると、閉じた状態で作り直される）
    st.session_state["work_restore_gen"] = st.session_state.get("work_restore_gen", 0) + 1


def carry_over_staff(year, month):
    """
    途中保存してある医師条件（回数・間隔・翌日PM duty・備考など）を、今の年月に引き継ぐ（ボタンのコールバック）。
    月ごとに変わる NG日・希望日・先生のコメントは空にする。
    """
    saved_at, data = _normal_work_state_for_carry(year, month)   # 通常版で途中保存した医師条件を使う
    if not data or not data.get("staff_csv"):
        st.session_state["carry_over_message"] = ("error", "通常版で途中保存した医師条件が見つかりませんでした。先に通常版で「💾 作業を途中保存する」を押してください。")
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
    st.session_state["staff_source_label"] = f"{data.get('year')}年{data.get('month')}月の医師条件（引き継ぎ）から"
    st.session_state.pop("staff_added_names", None)
    st.session_state["carry_over_message"] = (
        "success", f"{saved_at} に途中保存した{data.get('year')}年{data.get('month')}月の医師条件（{len(df)}名）を引き継ぎました。"
                   "NG日・希望日・先生のコメントは空にしてあります。")


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
                    if st.session_state.get("work_restore_pick") not in _labels:
                        st.session_state["work_restore_pick"] = _entries[0][0]
                    st.selectbox(f"読み込む途中保存（新しい順・最大{WORK_HISTORY_KEEP}回分）", list(_labels),
                                 format_func=lambda sid: _labels[sid], key="work_restore_pick")
                    st.button("この作業を読み込む", key="work_restore", type="primary", on_click=work_state_restore)
                    st.caption("期間・休日と別途決定の日・増員設定・確定当直・医師条件（NG日を含む）・取り込んだ先生のコメント・不足時の優先順位・アラートの基準が、"
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


# ============================================================
# 画面（通常版と同じ。色は茶色）
# ============================================================
st.title("当直作成アプリ（期間指定版・年末年始など）")
st.caption("今月の1日から、翌月の指定した日までを1つの当直案として作ります。使い方は通常版と同じです。")
st.markdown("<a href='./' target='_self' style='font-weight:700'>🏠 ← 通常版の当直作成に戻る</a>", unsafe_allow_html=True)
with st.expander("📖 使い方（期間指定版・作業の流れ・結果の見方）", expanded=False):
    tab_intro, tab_flow, tab_ng, tab_result, tab_trouble = st.tabs(["はじめに（期間指定版）", "作業の流れ", "NG日の集め方", "結果の見方・Excel", "困ったとき"])
    with tab_intro:
        st.markdown("""
**この画面（期間指定版）でできること**

年末年始など、月をまたいで通しで当直を決めたいときに使います。今月の1日から、翌月の指定した日までを1つの当直案として作ります（例：12/1〜1/5）。
画面の作りと使い方は通常版と同じです。ふだんの月の当直は、上の「🏠 ← 通常版の当直作成に戻る」から通常版で作ってください。

**通常版と違うところ**

- 年・月の右の「翌月の何日まで含めるか」で、期間の終わりを決めます。
- カレンダーや表では、翌月の日は「1/1」「1/2」のように月も付けて表示します。
- 医師条件の回数（月上限・月下限・休日上限・各枠の上限など）は、**期間全体での回数**として扱います（列名は通常版と同じ「月上限」「月下限」のままです）。期間が長い分、必要に応じて多めにしてください。
- 希望日は、翌月の日を「1/3」「1/3:A宿直」のように月も付けて書きます（今月の日は「30」「30:A宿直」のままで構いません）。
- 「📋 先月の医師条件を引き継ぐ」は、**通常版で途中保存した医師条件**を読み込みます。
- 途中保存は、通常版とは別の場所に、保存した回ごとに新しいものから20回分まで残ります。「📂 前回の作業の続きから始める」で、どの回から再開するかを選べます。
- 先生方の専用URLと案内書類は、通常版と同じものを使います。期間指定版で受付を開始すると、先生方の画面が期間のカレンダーに切り替わります（通常版の受付は終わります）。

**管理用パスワード**

途中保存・続きからの読み込み・医師条件の引き継ぎ・専用URLの管理（受付・取り込みなど）には、管理用パスワードが必要です。
        """)
    with tab_flow:
        st.markdown("""
**作業は、画面の上から順に進めれば完了します。**

1. **期間を選ぶ**：年・月（この月の1日から）と、翌月の何日まで含めるかを選びます。作業の続きなら、上の帯の「📂 前回の作業の続きから始める」で戻せます。
2. **休日を設定**：期間全体のカレンダーで、「休日にする」「別途決定」を必要な日だけチェックします。
3. **人数を増やす枠（任意）**：1つの枠を2名以上にする日があれば設定します。
4. **確定当直（任意）**：すでに決まっている当直や、期間の前（先月末など）・後の当直を入れます。勤務間隔の計算に使われます。
5. **先生の名簿と医師条件**（「先生ごとの条件」の①）：「📋 先月の医師条件を引き継ぐ」で、通常版で途中保存した医師条件を読み込みます。NG日・希望日は空になります。
   読み込むと、①に「今の名簿」と「読み込んだ医師条件」（見るだけの表）が出るので、正しく読み込めたかを確認します。
   新しく来た先生は「➕ 先生を追加する」で、入れる位置と条件のもとを選んで追加します。いなくなった先生は「➖ 先生を外す」で外します。
6. **NG日**（②）：先生を選んでカレンダーでNG日を入れ、「NG日を反映する」を押します（紙やメールで集めたNG日を入力します）。専用URLで集める場合は、「📨 先生方からのNG日入力（専用URL）」で受付を開始し、提出されたNG日を取り込みます（「NG日の集め方」のタブを参照）。
7. **医師条件の調整**（③・④）：当直案の枠ができるだけ埋まるように、③の表で回数・上限・勤務間隔・希望日を調整します（①の表は見るだけなので、直すのは③です）。④の「⚖️ 当直の枠は足りそう？」で、赤い棒がないかを確認します。
8. **不足時の優先順位（任意）**
9. **当直案を作る**：下の帯の「🚀 当直案を作成」を押します（計算に最大1分少々かかります）。
10. **確認してダウンロード**：「医師別の当直一覧」をExcelでダウンロードします。
11. **途中保存**：上の帯の「💾 作業を途中保存する」を押しておくと、続きから再開できます。
        """)
    with tab_ng:
        st.markdown("""
**方法1：管理者が直接入力する（いまの運用）**

「先生ごとの条件」の②（NG日を入れる）で、「NG日を入力する医師」で先生を選び、カレンダーの「日直」「宿直」のボタンを押してNGにし、「NG日を反映する」を押します。紙やメールで集めたNG日は、こちらで入力します。
反映していない変更があると、ボタンがオレンジに光ります。そのまま別の先生に切り替えようとすると注意が出ます。
「全日NGにする」「すべてOKに戻す」のボタンで、まとめて指定・解除もできます。

**方法2：先生方に専用URLから入力してもらう（慣れてきたら）**

「先生ごとの条件」の②（NG日を入れる）の中の「📨 先生方からのNG日入力（専用URL）」を開き、管理用パスワードを入れて使います。

1. **受付を開始**：先に、いちばん上のカレンダーで「休日にする」「別途決定」を済ませておきます。締め切り日を書いて「受付を開始する」を押します。
   受付中にカレンダーを変えた場合も、パスワードを入れた状態なら自動で先生側に反映されます。
2. **URLを配る（初回だけ）**：「配布用の案内書類（QRコード付き）」をダウンロードして印刷し、点線で三つ折りにして封筒で各先生に渡します。URLは毎月同じなので、翌月からは「入力が始まりました」と伝えるだけです。
3. **先生が入力**：先生はQRコードを読み取り、入れない日の「日直」「宿直」のボタンを押してNGにし、「NG日とコメントを保存する」を押します。「全日NGにする」「すべてOKに戻す」のボタンもあります。締め切りまでは何度でも直せて、最後に保存した内容が採用されます。
4. **提出状況を確認**：管理画面の表で、提出済み・未提出・コメントを確認できます。「🔄 最新の状態に更新」で最新になります。
5. **締め切る**：「🔒 締め切る」を押すと、先生側は編集できなくなります。
6. **取り込む**：「取り込むと変わる先生」の表で内容とコメントを確認し、取り込む先生を選んで「⬇️ 選んだ先生のNG日を取り込む」を押します。その先生のNG日は提出内容に置き換わり、コメントも当直一覧のメモ欄に表示されるようになります。

**方法2でやり直したいとき**

- **取り込みを取り消す**：「↩️ 取り込みを取り消す」で、取り込む前のNG日に戻せます（先生ごとのNG日の欄にもボタンがあります）。取り込んだあとにアプリを開き直したり、途中保存から読み込んだりした場合は取り消せないので、カレンダーで直してください。
- **先生の提出をリセットする**：入力をやり直してもらう場合は、「先生の提出をリセットする」で先生を選んでリセットします。未提出に戻り、先生の画面も空になります（これまでの提出は履歴に残ります）。
- **URLを作り直す**：案内書類をなくした・URLが漏れた場合に使います。古いURLは使えなくなります。

**NGの種類（どちらの方法でも共通）**

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
- **薄紫のマス**：期間の前後の確定当直から「間隔下限」以内のため、入れない日（期間中の確定当直の前後には色を付けません）
- **×**：当直NGの日（平日は宿直NG）
- **休日のマス**：上半分＝日直、下半分＝宿直です。上に×＝日直NG、下に×＝宿直NG、上下とも×＝全NG
- **黄色の斜線**：翌日PM dutyのため宿直を外す日。×があれば、NGも付いている日です
- **下半分だけ黄色の斜線**：休日の翌日PM dutyの日。宿直は外しますが、日直には入れます（×の意味は休日のマスと同じ）
- **NG人数の行**（見出しのすぐ下）：その日にNGの先生の数。翌日PM dutyで宿直を外す先生と、薄紫のマス（勤務間隔のため入れない日）の先生も数えます。多い日ほど濃いオレンジです
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
- **「未決定」の枠が出る**：作成ボタンの下の「計算状況・条件の注意事項と対処方法」を見てください。「先生ごとの条件」の④「⚖️ 当直の枠は足りそう？」で、赤い棒（足りない枠）がないかも確認します。NG日・回数の上限・勤務間隔の条件が厳しすぎる可能性があります。月上限を増やす、勤務間隔を短くするなどして、下の帯の「🚀 当直案を作成」で作り直してください。少しだけなら、Excelで手で埋めることもできます。
- **アプリを開くと入力した内容が消えている**：しばらく使わないと画面がリセットされます。上の帯の「📂 前回の作業の続きから始める」で読み込んでください。先生方の提出内容と専用URLはスプレッドシートに残っているので消えません。
- **英語の画面（Zzzz…）が出る**：アプリが休止しています。青いボタン「Yes, get this app back up!」を押して1分ほど待つと開きます。
- **先生のNG日が当直案に入っていない**：カレンダーで入力した場合は、「NG日を反映する」を押したか確認してください。専用URLで集めた場合は、管理画面で「⬇️ 選んだ先生のNG日を取り込む」を押したか確認してください。
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
    border-left: 5px solid #a0662f; border-radius: 0 8px 8px 0;
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
  background:linear-gradient(90deg,#5b3a1e 0%,#a0662f 100%);color:#fff;box-shadow:0 2px 6px rgba(31,52,82,.18)}
.sec-banner .ic{font-size:1.6rem;line-height:1}
.sec-banner .t{font-size:1.35rem;font-weight:800;letter-spacing:.02em;line-height:1.3}
.sec-banner .s{font-size:.88rem;opacity:.9;margin-top:2px}
.sec-banner .tag{margin-left:auto;font-size:.8rem;font-weight:800;padding:3px 12px;border-radius:999px;background:rgba(255,255,255,.18);border:1px solid rgba(255,255,255,.45);white-space:nowrap}
.sec-banner.first{margin-top:.6rem}
/* 画面上部の帯（ボタンの並ぶ所）：下の画面との境目が分かるよう、細い線とごく薄い影を付ける。
   すりガラス風にして、下を流れる内容が透けて重ならないようにする */
[data-testid="stHeader"]{background:linear-gradient(90deg,#4a2e15 0%,#7a4a20 100%) !important;
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
.step-h{display:flex;align-items:center;flex-wrap:wrap;gap:4px 10px;margin:1.8rem 0 .6rem;padding-bottom:6px;border-bottom:2px solid #d6e0ec}
.step-h .no{display:inline-flex;align-items:center;justify-content:center;width:1.9rem;height:1.9rem;border-radius:50%;background:#4a2e15;color:#fff;font-weight:800;font-size:1rem;flex:0 0 auto}
.step-h .tt{font-size:1.2rem;font-weight:800;color:#4a2e15}
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
section_banner("📅", "期間・休日・必要人数の設定", "作る期間・休日と別途決定の日・1枠の人数を決めます", "必須", first=True)
st.info("作る期間（年・月・翌月の何日まで）を選んでください。土日祝日は自動で休日扱いになります。下のカレンダーで、平日でも休日扱いにする日は「休日にする」、別の方法で当番を決める日は「別途決定」にチェックします。")
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
  padding:0 9rem 0 1rem;background:linear-gradient(90deg,#4a2e15 0%,#7a4a20 100%);
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
if "sel_extra" not in st.session_state:
    st.session_state["sel_extra"] = 5
col_y, col_m, col_e = st.columns(3)
year = int(col_y.number_input("年", min_value=2000, step=1, key="sel_year"))
month = int(col_m.number_input("月（この月の1日から）", min_value=1, max_value=12, step=1, key="sel_month"))
_nf = _REAL_DATE(year, month, _REAL_MONTHRANGE(year, month)[1]) + datetime.timedelta(days=1)
extra = int(col_e.number_input(f"翌月（{_nf.month}月）の何日まで含めるか", min_value=0, max_value=_REAL_MONTHRANGE(_nf.year, _nf.month)[1], step=1, key="sel_extra"))
P = TouchokuPeriod(year, month, extra)
_install_period(P)
st.info(f"**作る期間：{P.title()}（{P.n}日間）**　この期間全体を1つの当直案として作ります。"
        "医師条件の回数（月上限・月下限・休日上限・各枠の上限など）は、この期間全体での回数として扱います。")
restored_state = st.session_state.get("restored_state")

st.divider()

st.subheader(f"📅 休日と別途決定の日の設定 - {P.title()}")
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
                        f"font-size:0.95rem;line-height:1.35;margin:0 0 8px 0;padding:0;'>{_day_label(day)}</div>",
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
next_month_first_label = f"{next_month_first.month}/{next_month_first.day}（{WEEKDAYS_JA[next_month_first.weekday()]}）"
st.markdown(f"**期間の翌日の扱い**（最終日 {month_last_label} の宿直に、翌日PM dutyの制限をかけるかどうか）")
if next_month_first.weekday() >= 5 or jpholiday.is_holiday(next_month_first):
    # 土日祝は自動で休日扱いになるため、チェックボックスは表示しない
    _kind = "祝日" if jpholiday.is_holiday(next_month_first) else "土日"
    st.caption(f"期間の翌日 {next_month_first_label} は{_kind}なので、自動で休日として扱います（チェックは不要です）。"
               f"翌日が休みなので、最終日 {month_last_label} の宿直は「翌日PM duty」の制限を受けません。")
    next_month_special_holiday = False
else:
    next_month_special_holiday = st.checkbox(
        f"期間の翌日 {next_month_first_label} は病院独自の休み",
        key=f"next_month_special_holiday_{year}_{month}",
        help="平日でも病院が休みの場合だけチェックしてください。"
    )
    st.caption(f"期間の翌日 {next_month_first_label} は平日です。病院が休みの場合だけチェックしてください。"
               f"チェックすると、最終日 {month_last_label} の宿直は「翌日PM duty」の制限を受けません。")

holiday_total_placeholder = st.empty()

st.divider()

st.subheader("👥 1つの枠を2名以上にする設定（任意）")
st.info("通常は各枠1名です。増員する場合だけ行を追加し、日付・枠・合計人数を選んでください。2名体制にしたい場合は「2」を入力します。")
st.caption("日直の増員は休日扱いの日に設定してください。平日に日直を設ける場合は、先に上のカレンダーで「休日にする」にチェックを入れます。")

_, num_days = calendar.monthrange(year, month)
date_options = [_day_label(d) for d in range(1, num_days + 1)]
_label_to_day = {_day_label(d): d for d in range(1, num_days + 1)}

multi_df_template = pd.DataFrame(columns=["日付", "当直枠", "人数"])
if restored_state and restored_state.get("multi") and (int(restored_state["year"]), int(restored_state["month"]), int(restored_state.get("extra", 0))) == (year, month, extra):
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
            d_val = _label_to_day[d_str]
            multi_slots_dict[(d_val, s_val)] = int(c_val)
        except (ValueError, TypeError, KeyError):
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
    section_banner("📅", "確定当直（期間の前後・期間中）", "すでに決まっている当直や、期間の前（先月末など）・後の当直を入れます", "任意", anchor="sec-fixed")
    st.info("期間の前（先月末など）・期間中・期間の後の、担当者が決まっている当直を入力してください。期間の前後の勤務は、勤務間隔の確認に使います。入力がなければ、この項目は飛ばせます。")
    with st.expander("入力例：確定当直の書き方と扱い", expanded=False):
        st.markdown("""
- CSVのアップロードと、下の表への直接入力のどちらでも入力できます。
- 日付は `12/30` や `1/2` のように月から入力します。期間の前後の日付も入れられます。
- 担当する先生は、枠の欄をクリックすると出る一覧（医師条件の名前）から選びます。
- 同じ日・同じ枠を2人以上で担当する場合は、同じ日付の行を追加して、もう1人を選びます。CSVで `佐藤、鈴木` のように「、」で区切って入れた場合は、読み込むと自動で2行に分かれます。
- 未確定の枠は空欄で構いません。
- 期間中の平日に日直を入力する場合は、先に上のカレンダーでその日を「休日にする」にしてください。
- 「別途決定」にした日（年末年始など）の当番が決まったら、ここに入れておくと、当直案と一覧に表示され、前後の勤務間隔の計算にも使われます。
- 期間中の確定当直は、NG日や翌日PM duty（宿直を外す曜日）よりも優先されます。回数の上限（月上限・休日上限・各枠の上限など）を超える場合も、確定した分はそのまま入ります。
- ほかの勤務は、確定当直から「間隔下限」以上あけて自動で入れます。確定当直同士の間隔が近すぎる場合は、確定内容をそのまま残し、注意を表示します。

**期間中の同じ日・同じ枠を指定した場合、確定当直への入力と「希望優先度100以上」での指定は、どちらも固定され、計算上の扱いは基本的に同じです。両方に入力する必要はありません。**
        """)

    fixed_columns = ["日付"] + FIXED_SLOT_COLUMNS   # 当直案と同じ並び。平日か休日かは、上のカレンダーで自動的に決まるので入力しない
    fixed_template_df = pd.DataFrame(columns=fixed_columns)
    fixed_csv_template = fixed_template_df.to_csv(index=False).encode('utf-8-sig')

    fixed_file = csv_upload_card(
        "fixed", "CSV", "確定当直のCSVファイルをアップロード（任意）",
        "下の表に直接入力することもできます。たくさんある場合は、ひな形に入力して読み込むと便利です。",
        f"fixed_csv_{st.session_state.get('uploader_gen', 0)}",
        fixed_csv_template, "確定当直_ひな形.csv", "📥 確定当直のひな形（CSV）をダウンロード")

    # 表の中身は画面の裏（session_state）に持っておき、CSVを読み込んだときや途中保存から読み込んだときだけ置き換える
    if fixed_file is not None:
        _fixed_src = ("file", fixed_file.name, fixed_file.size)
    elif restored_state and restored_state.get("fixed_csv"):
        _fixed_src = ("restored", hashlib.sha256(restored_state["fixed_csv"].encode("utf-8")).hexdigest())
    else:
        _fixed_src = ("empty",)
    if st.session_state.get("fixed_src") != _fixed_src or "fixed_rows" not in st.session_state:
        if _fixed_src[0] == "file":
            try:
                base_fixed_df = parse_fixed_csv(fixed_file.getvalue())
            except Exception as e:
                st.error(f"確定当直ファイルの読み込みに失敗しました。詳細: {e}")
                st.stop()
            if "日付" not in base_fixed_df.columns:
                st.error("確定当直CSVに「日付」の列がありません。ひな形の列名を確認してください。")
                st.stop()
        elif _fixed_src[0] == "restored":
            base_fixed_df = parse_fixed_csv(restored_state["fixed_csv"].encode("utf-8"))
        else:
            base_fixed_df = pd.DataFrame(columns=fixed_columns)
        st.session_state["fixed_rows"] = fixed_rows_for_editor(base_fixed_df)
        st.session_state["fixed_src"] = _fixed_src
        st.session_state["fixed_editor_gen"] = st.session_state.get("fixed_editor_gen", 0) + 1
    base_fixed_df = st.session_state["fixed_rows"]

    # 担当医師は、医師条件の名前から選ぶ（日本語を打たなくてよいので、入力の切り替えで文字が崩れない）
    _fixed_names = list(dict.fromkeys(
        list(st.session_state.get("fixed_name_options", []))
        + [v for s_ in FIXED_SLOT_COLUMNS for v in base_fixed_df[s_].tolist() if v]))
    st.write(f"日付（例：{month}/1）を入力し、担当する先生を枠の欄の一覧から選びます。行は下の「＋」で追加できます。")
    if not _fixed_names:
        st.caption("先に下の「医師条件」に先生の名前を入れると、ここで名前を選べるようになります。")
    _fixed_key = f"fixed_editor_{st.session_state.get('fixed_editor_gen', 0)}"
    _fixed_view = base_fixed_df.copy()
    for s_ in FIXED_SLOT_COLUMNS:
        _fixed_view[s_] = _fixed_view[s_].astype(object).where(_fixed_view[s_] != "", None)
    _fixed_config = {"日付": st.column_config.TextColumn("日付", help=f"例：{month}/1 または {year}/{month}/1")}
    _fixed_config.update({s_: st.column_config.SelectboxColumn(s_, options=_fixed_names, help="担当する先生を一覧から選びます。消すときはマスを選んで Delete キー。")
                          for s_ in FIXED_SLOT_COLUMNS})
    _fixed_config.update({c: st.column_config.TextColumn(c) for c in base_fixed_df.columns if c not in _fixed_config})
    st.data_editor(
        _fixed_view, num_rows="dynamic", use_container_width=True, height=200, hide_index=True,
        column_config=_fixed_config, key=_fixed_key, on_change=apply_fixed_edits, args=(_fixed_key,),
    )
    st.caption("同じ日・同じ枠を2人以上で担当する場合は、同じ日付の行をもう1行追加して、もう1人を選んでください。")
    edited_fixed_df = st.session_state["fixed_rows"].copy()

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
    step_header(1, "以前の医師条件の取り込み（引き継ぎ or CSVアップロード）、医師を追加（or外す）")
    st.markdown("<div class='imp-choose'>📥 以前の医師条件を取り込む（どちらか1つの方法を使います）</div>", unsafe_allow_html=True)
    imp_a, imp_or, imp_b = st.columns([10, 1.4, 10])   # 方法1（引き継ぎ）・または・方法2（CSV）
    imp_or.markdown("<div class='imp-or'><span>または</span></div>", unsafe_allow_html=True)
    st.markdown("<div class='imp-note'>どちらで取り込んでも、下の「今の名簿」と「読み込んだ医師条件」の表がその内容に置き換わります。</div>",
                unsafe_allow_html=True)
    # --- 方法1：途中保存から先月の医師条件を引き継ぐ（管理用パスワードが必要） ---
    with imp_a, st.container(key="csvcard_carry"):
        st.markdown("<div class='csv-card-head blue'><span class='lab'>方法1</span>通常版の途中保存から医師条件を引き継ぐ</div>", unsafe_allow_html=True)
        if not remote_is_configured():
            st.caption("保存用のGoogleスプレッドシートの設定をすると使えます。今はCSVのアップロード（方法2）で読み込んでください。")
        else:
            _co_msg = st.session_state.pop("carry_over_message", None)
            # ボタンの名前：ログイン中は、途中保存してある月を出す（ふつうは先月）
            _co_name = "先月"
            if st.session_state.get("remote_admin_ok"):
                _co_at, _co_data = _normal_work_state_for_carry(year, month)
                if _co_data and _co_data.get("staff_csv"):
                    _co_ym = (int(_co_data.get("year", 0)), int(_co_data.get("month", 0)))
                    _co_prev = (year, month - 1) if month > 1 else (year - 1, 12)
                    _co_name = f"先月（{_co_ym[1]}月）" if _co_ym == _co_prev else f"{_co_ym[0]}年{_co_ym[1]}月"
            _co_label = f"📋 {_co_name}の医師条件を引き継ぐ" + "\u200b" * st.session_state.get("carry_over_gen", 0)
            st.caption("通常版で「💾 作業を途中保存する」で保存した医師条件（回数・勤務間隔・翌日PM duty・備考など）を読み込みます。期間の前の月の保存があればその最新、なければいちばん新しい保存を使います。NG日・希望日・先生のコメントは空になります。")
            _co_box = st.popover(_co_label) if hasattr(st, "popover") else st.expander(_co_label, expanded=False)
            with _co_box:
                if admin_logged_in("carry_over_pw"):
                    _co_at, _co_data = _normal_work_state_for_carry(year, month)
                    if _co_data and _co_data.get("staff_csv"):
                        st.write(f"通常版の途中保存：**{_co_at}**（{_co_data.get('year')}年{_co_data.get('month')}月の作業）")
                        st.caption(f"この医師条件（先生の名前・回数・勤務間隔・翌日PM duty・希望優先度・備考など）を、{P.title()}用に引き継ぎます。"
                                   "月ごとに変わる NG日・希望日・先生からのコメントは空になります。今の画面の医師条件は置き換わります。")
                        if False:
                            st.warning(f"途中保存は{year}年{month}月（今選んでいる月と同じ）の作業です。NG日なども含めて続きから作業する場合は、"
                                       "上の帯の「📂 前回の作業の続きから始める」を使ってください。")
                        st.button(f"{P.title()}用に医師条件を引き継ぐ", key="carry_over_do", type="primary",
                                  on_click=carry_over_staff, args=(year, month))
                    else:
                        st.caption("通常版で途中保存された医師条件がありません。通常版で「💾 作業を途中保存する」を押してから、ここで読み込んでください。")
                else:
                    st.caption("引き継ぎには管理用パスワードが必要です。")
            if _co_msg:
                (st.success if _co_msg[0] == "success" else st.error)(_co_msg[1])
with staff_box_cond:
    step_header(3, "医師条件を調整する", "当直案の枠ができるだけ埋まるように、ここで医師条件を直します（直すと①の表示も同じ内容になります）", anchor="sec-cond")
    st.info("下の表に先生ごとの条件を入力してください。セルをクリックして編集できます。列名にカーソルを合わせると、入力例や説明が表示されます。")
    st.caption("最初に表示される「型1」〜「型9」は、よくある条件の入力例です。名前の欄が型の中身です（例：型4＝A宿直・A日直・外来日直に入る先生）。実際の先生の名前に置き換え、回数などを調整して使ってください。行は「＋」で増やせます。")
    with st.expander("入力例：曜日・希望日・備考", expanded=False):
        st.markdown("""
| 項目 | 入力方法・意味 |
| --- | --- |
| 翌日PM duty | `水,木` のように半角カンマで区切ります。翌日PMにdutyがある曜日を入力します（例：木曜PMにdutyがある場合は「水」）。原則としてその曜日の宿直を外しますが、翌日が休日なら宿直に入る場合があります。日直は対象外です。 |
| 希望日 | `10,15` はその日のいずれかの枠、`10:A宿直` はその枠を希望します。翌月の日は `1/3`、`1/3:A宿直` のように月も書きます。複数の希望は半角カンマで区切ります。日直は休日（または「休日にする」にした日）だけ指定できます。「別途決定」の日の希望は使いません。 |
| 指定できる枠 | A宿直・B宿直・外来宿直・A日直・B日直・外来日直。表記を一致させてください。 |
| 備考 | 管理用のメモです。「学会」などと書いても計算条件には反映されません。休みはNG日で指定してください。 |

曜日にかかわらず勤務できない日は、上の「② NG日を入れる」のカレンダーでNGにします。
平日は「宿直」のボタン、休日は「日直」「宿直」のボタンを押してNGにします（休日に終日入れない場合は両方押します）。
    """)
    with st.expander("回数・勤務間隔の数え方", expanded=False):
        st.markdown("""
| 項目 | 意味・入力例 |
| --- | --- |
| 間隔下限 | 勤務と次の勤務の間に空ける日数です。5日なら、10日の次は16日以降です。 |
| 月下限 | 期間全体で、できるだけ確保したい回数です。条件によっては、この回数に届かないことがあります。 |
| 月上限 | 日直と宿直を合わせた、**期間全体での**上限です（列名は通常版と共通）。 |
| 休日上限 | 土日祝日と「休日にする」にした日に担当する、日直と宿直の合計上限です。 |
| 宿直上限・日直上限 | A・B・外来を合わせた宿直（日直）の期間全体での上限です。空欄なら制限しません。 |
| 各枠の上限（A宿上限〜外日上限） | A宿直など、それぞれの枠を担当する期間全体での上限です。表では「A宿上限」「外日上限」のように短く表示します。 |

1つの枠を1回と数えます。確定当直で同じ日に日直と宿直を担当する場合は2回です。
月下限は、月上限以下に設定してください。
確定当直や希望優先度100以上の希望がある場合は、上限を超えることがあります。ほかの勤務は確定当直から「間隔下限」以上あけて入れ、確定当直同士が近すぎる場合は注意を表示します。
    """)
    with st.expander("希望優先度：通常の希望と、100以上の特別な設定", expanded=False):
        st.markdown("""
- **通常は「1」**を使用します。1〜99は、数字が大きいほど希望を優先しますが、NG日・回数上限・勤務間隔などの範囲内で割り当てます。
- **100以上は、その先生の希望日すべてを確定当直と同じ扱いにする設定**です。「できれば入りたい」という用途には使わないでください。
- 確定にした日は、NG日や翌日PM duty（宿直を外す曜日）よりも優先されます。回数の上限を超える場合も、確定した分はそのまま入ります。ほかの勤務は、確定の日から「間隔下限」以上あけて入れます。確定当直同士の間隔が近すぎる場合は、注意を表示します。
- 一部の勤務だけを確定させたい場合は、上の「📅 確定当直」へ入力し、希望優先度は通常の値にしてください。
- 指定の誤りや条件の組み合わせによっては作成できない場合があります。作成後に、確定にした日が当直案に入っているか確認してください。

**期間中の同じ日・同じ枠を指定した場合、確定当直への入力と「希望優先度100以上」での指定は、どちらも固定され、計算上の扱いは基本的に同じです。両方に入力する必要はありません。**
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
        # 名前, 備考, 希望日, 希望優先度, 間隔下限, 月下限, 月上限, 休日上限, 宿直上限, 日直上限, A宿直, B宿直, 外来宿直, A日直, B日直, 外来日直
        ("型1_A宿直・A日直", "Aの宿直・日直だけ。期間で1回", "", 1, 5, 1, 1, 1, 3, 3, 1, 0, 0, 1, 0, 0),
        ("型2_A日直のみ・希望日確定", "決めた日のA日直だけ（希望優先度100で確定）", _tmpl_wish2, 100, 0, 0, 2, 2, 3, 3, 0, 0, 0, 2, 0, 0),
        ("型3_外来日直のみ", "外来日直だけ。期間で1回まで", "", 1, 5, 0, 1, 4, 3, 3, 0, 0, 0, 0, 0, 1),
        ("型4_A宿直・A日直・外来日直", "Aの宿直・日直と外来日直。期間で2〜3回", "", 1, 5, 2, 3, 1, 3, 3, 3, 0, 0, 1, 0, 1),
        ("型5_A・B宿直・A日直・外来日直", "型4にB宿直を追加。期間で2〜3回", "", 1, 5, 2, 3, 1, 3, 3, 3, 3, 0, 1, 0, 1),
        ("型6_外来宿直以外すべて", "型5にB日直を追加。期間で3〜4回", "", 1, 5, 3, 4, 2, 3, 3, 3, 3, 0, 1, 1, 2),
        ("型7_全枠", "6枠すべて。期間で3〜4回", "", 1, 5, 3, 4, 2, 3, 3, 3, 3, 3, 1, 1, 2),
        ("型8_日直のみ", "A・B・外来の日直だけ。期間で1〜3回", "", 1, 4, 1, 3, 3, 3, 3, 0, 0, 0, 3, 3, 3),
        ("型9_A以外", "B・外来の宿直と日直。期間で3〜4回", "", 1, 4, 3, 4, 2, 3, 3, 0, 3, 3, 0, 3, 3),
    ]
    _tmpl_cols = ["先生の名前", "備考（メモ・説明など自由記入）", REQUEST_COLUMN, "希望優先度(数字が大きいほど優先)",
                  "間隔下限", "月下限", "月上限", "休日上限", "宿直上限", "日直上限",
                  "A宿直上限", "B宿直上限", "外来宿直上限", "A日直上限", "B日直上限", "外来日直上限"]
    template_data = {c: [r[k] for r in _tmpl_rows] for k, c in enumerate(_tmpl_cols)}
    template_data["翌日PM duty"] = [""] * len(_tmpl_rows)
    template_data[NG_COLUMN] = [""] * len(_tmpl_rows)
    # 列の並びは、これまでのひな形と同じにする
    template_data = {c: template_data[c] for c in ["先生の名前", "備考（メモ・説明など自由記入）", "翌日PM duty", NG_COLUMN, REQUEST_COLUMN,
                                                   "希望優先度(数字が大きいほど優先)", "間隔下限", "月下限", "月上限",
                                                   "休日上限", "宿直上限", "日直上限", "A宿直上限", "B宿直上限", "外来宿直上限",
                                                   "A日直上限", "B日直上限", "外来日直上限"]}
    df_template = pd.DataFrame(template_data)
    # 「➕ 先生を追加する」で型から選べるように、型の中身を覚えておく
    st.session_state["staff_templates"] = {r["先生の名前"]: r for r in df_template.to_dict("records")}
    # 保存用CSVと同じUTF-8（BOM付き）に統一。Excelでもそのまま開ける。
    csv_template = df_template.to_csv(index=False).encode('utf-8-sig')

    # --- 方法2：CSVファイルをアップロード ---
    with imp_b:
        uploaded_file = csv_upload_card(
            "staff", "方法2", "CSVファイルをアップロード",
            "手元のCSV（以前ダウンロードした医師条件、またはひな形に入力したもの）を読み込みます。",
            f"staff_csv_{st.session_state.get('uploader_gen', 0)}_{st.session_state.get('staff_uploader_gen', 0)}",
            csv_template, "医師条件_ひな形.csv", "📥 ひな形（CSV）をダウンロード",
            "初めての場合は、ひな形をダウンロードしてExcelで入力してから読み込みます。")
    staff_box_ro = st.container()   # ①のいちばん下：読み込んだ医師条件（見るだけ。中身は医師条件の表を読んだあとで入れる）

with staff_box_cond:
    if uploaded_file is not None:
        if st.session_state.get('last_uploaded_file_id') != uploaded_file.file_id:
            for key in list(st.session_state.keys()):
                if key.startswith("ng_"):
                    del st.session_state[key]
            st.session_state['last_uploaded_file_id'] = uploaded_file.file_id
            st.session_state["staff_source_label"] = f"CSV「{uploaded_file.name}」から"
            st.session_state.pop("staff_added_names", None)
        try:
            base_df = parse_staff_csv(uploaded_file.getvalue())
        except Exception as e:
            st.error(f"医師条件CSVを読み込めませんでした。詳細: {e}")
            st.stop()
    elif restored_state and restored_state.get("staff_csv"):
        base_df = parse_staff_csv(restored_state["staff_csv"].encode("utf-8"))
    else:
        base_df = df_template.copy()
    staff_is_template = uploaded_file is None and not (restored_state and restored_state.get("staff_csv"))

    if "先生の名前" not in base_df.columns:
        st.error("医師条件CSVに「先生の名前」の列がありません。ひな形の列名を確認してください。")
        st.stop()
    # 宿直上限・日直上限の列がない（以前の）CSVでも、「休日上限」の右に空欄の列を足す（空欄＝制限なし）
    for _gc in reversed(list(GROUP_CAP_COLUMNS)):
        if _gc not in base_df.columns:
            _pos = list(base_df.columns).index("休日上限") + 1 if "休日上限" in base_df.columns else len(base_df.columns)
            base_df.insert(_pos, _gc, None)
        base_df[_gc] = pd.to_numeric(base_df[_gc], errors="coerce")
    base_df = base_df.set_index("先生の名前")

    if "希望優先度(数字が大きいほど優先)" in base_df.columns:
        base_df["希望優先度(数字が大きいほど優先)"] = pd.to_numeric(base_df["希望優先度(数字が大きいほど優先)"], errors='coerce')

    text_cols = ["翌日PM duty", NG_COLUMN, REQUEST_COLUMN, "備考（メモ・説明など自由記入）"]
    for c in text_cols:
        if c in base_df.columns:
            base_df[c] = base_df[c].apply(lambda x: "" if pd.isna(x) or str(x).lower() in ["nan", "none", "<na>"] else str(x))

    st.caption("先生の名前は、ほかの先生と重ならない表記にしてください。")

    edited_df = staff_table(base_df, editable=True)

    staff_df = edited_df.reset_index()
    # 先生の追加・外すは、この表（今の画面の内容）をもとに行う
    st.session_state["staff_current_csv"] = staff_df.to_csv(index=False)

    # 確定当直の表で選べる先生の名前を、医師条件に合わせる（名前が変わったときは、確定当直の表の一覧も作り直す）
    _names_now = list(dict.fromkeys(n for n in (clean_text(v) for v in staff_df.get("先生の名前", pd.Series(dtype=str)).tolist())
                                    if n and n.lower() not in ("nan", "none")))
    if st.session_state.get("fixed_name_options") != _names_now:
        st.session_state["fixed_name_options"] = _names_now
        st.session_state["fixed_editor_gen"] = st.session_state.get("fixed_editor_gen", 0) + 1
        if not (create_bar_request or work_save_request):
            st.rerun()
    # 「➕ 先生を追加する」「➖ 先生を外す」（名前が変わったときの再表示のあとに出す。知らせが消えないように）
    with staff_box_addrm:
        staff_add_remove_ui(staff_df, year, month, is_template=staff_is_template)
    with staff_box_ro:
        staff_readonly_table(staff_df)

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

with staff_box_cap:
    step_header(4, "当直の枠は足りそう？（目安）", "条件を直すと、すぐに変わります")
    st.caption("埋める必要がある枠の数と、先生方が入れる回数の合計を比べています。先生ごとの回数は、期間全体での「月上限」と各枠の上限・休日上限のうち、"
               "いちばん少ないものを使います。足りていても、NG日・勤務間隔などによっては埋まらないことがあります。"
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

    ng_layout = st.radio("NGカレンダーの表示", ["月間カレンダー", "期間を横一列"], horizontal=True, key="ng_calendar_layout")
    ng_horizontal = ng_layout == "期間を横一列"
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
                        if val == "全NG": saved_strs.append(_day_label(d))
                        elif val == "日NG": saved_strs.append(f"{_day_label(d)}(日直NG)")
                        elif val == "宿NG": saved_strs.append(f"{_day_label(d)}(宿直NG)")
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
                        st.caption("灰色の日（" + "・".join(_day_label(d) for d in sorted(locked_days)) + "）は、いちばん上のカレンダーで「別途決定」にした日です。NGは入力できず、当直案でもこの日の当直は決めません（確定当直に入れた分だけ表示されます）。")
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
    st.markdown("### 📂 医師条件をCSVでダウンロード（手元にファイルで残す場合）")
    st.write("先生の名前・回数・勤務間隔・希望日・NG日・備考をCSVファイルとしてパソコンなどにダウンロードします。このファイルは「医師条件」のアップロード欄から読み込めます。")
    st.caption("ダウンロードの前に、各先生の「NG日を反映する」を押してください。対象期間・休日と別途決定の日・増員設定・確定当直・作成した当直案・色分けとそのメモは、このCSVには含まれません。"
               "続きから作業するだけなら、CSVではなく「💾 途中保存」を使うほうが簡単です。")

    # ダウンロードするCSVも、画面と同じく「先生の名前 → 備考 → その他」の順にそろえる（読み込みは列名で行うので、古い並びのCSVもそのまま使える）
    _first = ["先生の名前"] + [c for c in staff_df.columns if str(c).startswith("備考")]
    current_csv = staff_df[_first + [c for c in staff_df.columns if c not in _first]].to_csv(index=False).encode('utf-8-sig')
    st.download_button(
        label="📥 医師条件をCSVでダウンロード",
        data=current_csv,
        file_name=f"医師条件_{P.start.isoformat()}_{P.end.isoformat()}.csv",
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
st.info("全員の条件を満たす当直案が作れない場合に、どの枠から優先して埋めるかを決めます。数字が小さいほど優先し、同じ数字の枠は同順位として扱います。")
st.caption("人数が足りているときは、この設定は結果に影響しません。初期設定は、A宿直・B宿直・A日直・B日直が1、外来宿直が2、外来日直が3です。")

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
    "multi": [[_day_label(d), s, int(c)] for (d, s), c in multi_slots_dict.items()],
    "extra": extra,
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
        st.caption("※ 押すたびに、同じ条件でも違うパターンの当直案を作ります。気に入った案は、結果の緑のボタン「📥 この一覧をExcelでダウンロード」で保存してください（作成した当直案は途中保存には含まれません）。")
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
        # 先月末・来月初めの確定当直から「間隔下限」以内で、入れない日（今月の確定当直の前後は含めない）
        gap_blocked = {}
        for _, _r in staff_df.iterrows():
            _nm = str(_r['先生の名前'])
            _gv = pd.to_numeric(pd.Series([_r.get('間隔下限')]), errors='coerce').iloc[0]
            _gap = 0 if pd.isna(_gv) else int(_gv)
            _ext = list(past_worked_dates.get(_nm, [])) + list(future_worked_dates.get(_nm, []))
            if _ext and _gap > 0:
                gap_blocked[_nm] = {d for d in range(1, num_days + 1)
                                    if any(abs((datetime.date(year, month, d) - e).days) <= _gap for e in _ext)}
        matrix_days, matrix_rows = build_duty_matrix(df_result, staff_df, year, month, custom_holidays, next_month_special_holiday,
                                                     comments=imported_comments(year, month), gap_blocked=gap_blocked)
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
                    file_name=f"当直一覧_{P.start.isoformat()}_{P.end.isoformat()}.xlsx",
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
                st.caption("「別途決定」にした日（" + "・".join(P.label(d) for d in sorted(locked_days)) + "）は、このアプリでは当直を決めていません（確定当直に入れた分だけ表示）。")
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
                file_name=f"当直案_{P.start.isoformat()}_{P.end.isoformat()}.csv",
                mime="text/csv",
            )


else:
    st.warning("☝️ 表に先生の名前を入力するか、CSVファイルをアップロードしてください。")
