# -*- coding: utf-8 -*-
"""
نظام الحاجري المالي | ELHAGRY Financial Systems — واجهة سطح المكتب (Tkinter).

الواجهة فقط: كل المنطق المحاسبي في الحزمة elhagry_core (نفس المحرك الذي تستخدمه واجهة الويب).
التشغيل:  python foodzone_erp.py
"""
import io
import math
import os
import re
import sys
import tempfile
import pandas as pd

from PIL import Image
from datetime import date, datetime, timedelta

# واجهة سطح المكتب (Tkinter) اختيارية: على الخوادم بدون شاشة (مثل Streamlit Cloud) لا تتوفر مكتبة Tk،
# فيُحمَّل المنطق المحاسبي فقط ويعمل التطبيق عبر واجهة الويب streamlit_app.py.
try:
    import tkinter as tk
    from tkinter import ttk, messagebox, filedialog, simpledialog
    import tkinter.font as tkfont
    from PIL import ImageTk
    HAS_TK = True
except Exception:  # pragma: no cover - بيئة بدون Tk
    HAS_TK = False

    class _TkStub:
        """بديل خامل يسمح بتعريف أصناف الواجهة دون تشغيلها (لا يُستدعى فعلياً بدون Tk)."""
        def __getattr__(self, name):
            return object

    tk = ttk = messagebox = filedialog = simpledialog = tkfont = ImageTk = _TkStub()

from elhagry_core import engine as _engine
from elhagry_core.engine import save_voucher, set_current_user, invoice_payment_prefill, NO_REP  # noqa: E402
from elhagry_core import credit as _credit, audit as _audit  # noqa: E402
from elhagry_desktop.phase2 import Phase2Mixin  # noqa: E402
from elhagry_core.controls import user_role_id, party_advance_available  # noqa: E402
from elhagry_core.engine import voucher_rep_info  # noqa: E402
from elhagry_core.engine import (  # noqa: E402  — واجهة برمجة المحرك المستخدمة في الشاشات
    ACCOUNT_TYPES, ALL_PERMISSION_KEYS, ATTACH_DOCS, ATTACH_EXTS, BANK_ACCOUNTS, BRAND_AR, BRAND_EN, BRAND_FULL,
    CASH_OPENING_SETTING, CAT_ASSET, CAT_EMPLOYEE, CAT_EXPENSE, CAT_PARTNER, CHART_INK, CHART_PAL, CHART_PAL_DEFAULT,
    CHART_TEXT, DASH_REFRESH_MS, DB_FILE, DEFAULT_ADMIN_USER, DEFAULT_EXPENSES, DISC_ACCOUNT_DEFAULTS,
    DISC_KIND_LABEL, DISC_KIND_SHORT, DISC_METHOD_LABEL, DISC_PL, DISC_PURCHASE, DISC_SALES, DISC_STOCK,
    ELHAGRY_LOGO_FILES, EN_MONTHS, EPS, EXCEL_PATH, FLOW, HAS_AR_LIBS, IMPORT_KINDS, LOGIN_ACCENT, LOGIN_FORM_BG,
    LOGIN_PANEL_BG, NON_CASH_EXPENSE_PREFIXES, NegativeStockError, PERMISSIONS, PERMISSION_KEYS, PURCHASE_RETURN,
    REPORT_CATALOG, REPORT_FIN_REQUIRES, REPORT_GROUP_OF, REPORT_INDEX, ROLE_LABEL, Report, SALES_RETURN,
    SEQUENCE_DEFS, STALE_DAYS_DEFAULT, SUB_PARENT, SUB_PERMISSIONS, SUMMARY_HEADERS, SUMMARY_SCOPES, TYPE_ASSET,
    TYPE_EXP, UserError, _block_system_entry, _cur, _pct, _simple_report_pdf, account_statement, account_tree_rows,
    accumulated_depreciation, active_admin_count, add_attachment, all_unbilled, attachment_counts, authenticate,
    auto_backup_daily, backup_database, balance_text, build_balances_summary_pdf, build_fulfilment_pdf,
    build_income_report, build_invoice_pdf, build_item_card_pdf, build_order_pdf, build_post_discount_pdf,
    build_report_pdf, build_sales_invoice_pdf, build_statement_pdf, cancel_order, cash_account_names, cash_accounts,
    cash_balances, cash_movements, clean_name, close_order, company_info, create_order, create_sub_account, db,
    delete_attachment, delete_invoice, delete_item, delete_journal_entry, delete_partner, delete_post_discount,
    delete_rep, delete_return, delete_voucher, disc_accounts, discount_invoice_choices, employee_accounts,
    export_report_xlsx, extract_logo_image, fetch_dashboard_charts, fetch_dashboard_metrics, fetch_expiry_summary,
    fetch_item_card, fetch_partner_statement, find_logo_file, fmt, fmt_cell, fulfil_order, get_attachment,
    get_cash_opening, get_setting, hash_password, import_chart_of_accounts, import_items_from_df,
    import_partner_openings_from_df, import_partners_from_df, import_preview_columns, import_stats_text, init_db,
    integrity_check, invoice_number, invoice_unbilled_selection, item_details,
    item_has_movements, item_reference_count, linked_partner_codes, list_attachments, list_partners,
    list_post_discounts, list_reps, load_invoice, load_post_discount, make_line, next_child_account_code,
    next_serial, norm_name, normalize_digits, parse_date, parse_discount, parse_float, partner_rep_history,
    partners_balance_summary, peek_serial, period_is_closed, period_label, post_discount_register, read_table_file,
    rep_names, repair_inventory, repair_invoice_paid, run_import, run_report,
    save_company_info, save_journal, save_partner_record, save_post_discount, save_rep, save_return,
    save_user_account, search_norm, set_cash_opening, set_disc_accounts, set_setting, statement_parties,
    stock_batches, summary_table_rows, today, update_invoice, update_item, update_sequence, update_voucher, user_can,
    user_can_fin, validate_batch_dates, validate_new_password, verify_password, write_import_template,
    write_upload_template,
)

class SearchableCombobox(ttk.Combobox):
    """قائمة منسدلة بالبحث التلقائي:
    - النتائج تنفتح وتُصفّى تلقائياً أثناء الكتابة (بأي جزء من الكود أو الاسم، وبأكثر من كلمة).
    - الاختيار بالنقر أو Enter، والتنقل بالأسهم ↑↓، وEsc للإغلاق.
    - القيمة المختارة مطابقة تماماً للسطر الظاهر (لا تُغيَّر قائمة القيم أثناء فتح القائمة المنسدلة الأصلية)."""
    NAV_KEYS = {"Up", "Down", "Return", "KP_Enter", "Escape", "Tab", "Left", "Right", "Shift_L", "Shift_R",
                "Control_L", "Control_R", "Alt_L", "Alt_R", "Home", "End", "Prior", "Next"}
    MAX_ROWS = 10
    MAX_RESULTS = 300

    def __init__(self, master, **kwargs):
        super().__init__(master, postcommand=self._on_native_post, **kwargs)
        self._all = []
        self._map = {}
        self._norm = {}
        self._num = {}
        self._popup = None
        self._lb = None
        # وسم ربط خاص بكل قائمة: لا تلغيه ربطات الصفحات على نفس الأداة (bind بدون add)
        tag = f"SearchableCombobox{id(self)}"
        self.bindtags((tag,) + self.bindtags())
        self.bind_class(tag, "<KeyRelease>", self._on_key)
        self.bind_class(tag, "<Down>", lambda e: self._move(1))
        self.bind_class(tag, "<Up>", lambda e: self._move(-1))
        self.bind_class(tag, "<Prior>", lambda e: self._move(-self.MAX_ROWS))
        self.bind_class(tag, "<Next>", lambda e: self._move(self.MAX_ROWS))
        self.bind_class(tag, "<Return>", self._on_return)
        self.bind_class(tag, "<KP_Enter>", self._on_return)
        self.bind_class(tag, "<Escape>", self._on_escape)
        self.bind_class(tag, "<Tab>", self._on_tab)
        self.bind_class(tag, "<FocusOut>", lambda e: self.after(180, self._maybe_hide))
        self.bind_class(tag, "<Destroy>", lambda e: self._destroy_popup())
        self.bind_class(tag, "<Unmap>", lambda e: self._hide())

    # ---------------- البيانات
    def set_items(self, pairs):
        """pairs = [(النص المعروض، البيانات المرتبطة)]. الاختيار الحالي يبقى لو ما زال موجوداً،
        والنص الذي يكتبه المستخدم الآن لا يُمسح عند تحديث القائمة."""
        current = self.get()
        self._map = dict(pairs)
        self._all = [d for d, _ in pairs]
        # فهرس البحث: (النص الموحد كاملاً، الكود الموحد، الاسم الموحد)
        self._norm = {}
        self._num = {}  # الجزء الرقمي من الكود بدون أصفار بادئة: C0198 → 198 (للبحث بالرقم فقط)
        for d in self._all:
            code, _, name = d.partition(" - ")
            self._norm[d] = (search_norm(d), search_norm(code), search_norm(name))
            self._num[d] = re.sub(r"\D", "", search_norm(code)).lstrip("0")
        self.configure(values=self._all)
        try:
            typing = self.focus_get() is self
        except (KeyError, tk.TclError):
            typing = False
        if current not in self._map and not typing:
            self.set('')
        if not typing:
            self._hide()

    def _matches(self, typed):
        """كل الكلمات المكتوبة يجب أن توجد (بأي ترتيب). الترتيب: كود مطابق ← يبدأ بالكود ← يبدأ بالاسم ← يحتوي."""
        q = search_norm(typed)
        words = q.split()
        if not words:
            return list(self._all)
        qnum = q.lstrip("0") if q.isdigit() else None
        ranked = []
        for i, v in enumerate(self._all):
            full, code, name = self._norm.get(v) or (search_norm(v), "", "")
            if not all(w in full for w in words):
                continue
            num = self._num.get(v, "")
            if code == q or name == q or (qnum and num == qnum):
                rank = 0
            elif code.startswith(q) or (qnum and num.startswith(qnum)):
                rank = 1
            elif name.startswith(q):
                rank = 2
            elif any(part.startswith(words[0]) for part in name.split()):
                rank = 3
            else:
                rank = 4
            ranked.append((rank, i, v))
        ranked.sort()
        return [v for _, _, v in ranked]

    def selected(self):
        """البيانات المرتبطة بالسطر المختار بدقة:
        النص المطابق تماماً ← كود مطابق ← اسم مطابق ← نتيجة بحث وحيدة. أي غموض = لا اختيار (لا تخمين)."""
        text = self.get().strip()
        if text in self._map:
            return self._map[text]
        if not text:
            return None
        q = search_norm(text)
        for idx in (1, 2):  # الكود ثم الاسم
            exact = [v for v in self._all if (self._norm.get(v) or ("", "", ""))[idx] == q]
            if len(exact) == 1:
                return self._map[exact[0]]
        if q.isdigit():  # رقم الكود فقط (198 = C0198) بشرط ألا يطابق أكثر من كود
            exact = [v for v in self._all if self._num.get(v) == q.lstrip("0")]
            if len(exact) == 1:
                return self._map[exact[0]]
        found = [v for v in self._matches(text)]
        return self._map[found[0]] if len(found) == 1 else None

    # ---------------- القائمة المنبثقة
    def _popup_visible(self):
        return self._popup is not None and self._popup.winfo_exists() and self._popup.winfo_viewable()

    def _ensure_popup(self):
        if self._popup is not None and self._popup.winfo_exists():
            return
        top = tk.Toplevel(self)
        top.withdraw()
        top.wm_overrideredirect(True)
        try:
            top.wm_attributes("-topmost", True)
        except tk.TclError:
            pass
        frame = tk.Frame(top, bg="#0284c7", bd=0, padx=1, pady=1)
        frame.pack(fill="both", expand=True)
        lb = tk.Listbox(frame, activestyle="none", exportselection=False, takefocus=0, relief="flat", bd=0,
                        justify="left", font=("Arial", 10), selectbackground="#0284c7", selectforeground="white",
                        highlightthickness=0)
        sb = ttk.Scrollbar(frame, orient="vertical", command=lb.yview)
        lb.configure(yscrollcommand=sb.set)
        sb.pack(side="left", fill="y")
        lb.pack(side="right", fill="both", expand=True)
        lb.bind("<ButtonRelease-1>", self._on_click)
        lb.bind("<Motion>", self._on_hover)
        self._popup, self._lb, self._sb = top, lb, sb

    def _show(self, items):
        if not items:
            self._hide()
            return
        self._ensure_popup()
        lb = self._lb
        lb.delete(0, "end")
        for v in items[:self.MAX_RESULTS]:
            lb.insert("end", v)
        rows = min(len(items), self.MAX_ROWS)
        lb.configure(height=rows)
        if len(items) > self.MAX_ROWS:
            self._sb.pack(side="left", fill="y")
        else:
            self._sb.pack_forget()
        self._set_index(0)
        self.update_idletasks()
        self._popup.update_idletasks()
        try:
            f = tkfont.Font(font=lb.cget("font"))
            longest = int(max(f.measure(v) for v in items[:self.MAX_RESULTS]) * 1.3) + 40
        except (tk.TclError, ValueError):
            longest = 0
        w = min(max(self.winfo_width(), 320, longest), 720)
        h = lb.winfo_reqheight() + 2
        x = self.winfo_rootx() + self.winfo_width() - w  # محاذاة يمين القائمة مع يمين الحقل (عربي)
        x = max(0, x)
        y = self.winfo_rooty() + self.winfo_height()
        if y + h > self.winfo_screenheight() - 40:
            y = max(0, self.winfo_rooty() - h)
        self._popup.geometry(f"{w}x{h}+{x}+{y}")
        try:
            had_focus = self.focus_get() is self
        except (KeyError, tk.TclError):
            had_focus = False
        pos = self.index("insert")
        self._popup.deiconify()
        self._popup.lift()
        if had_focus:
            self.focus_force()
            self.icursor(pos)

    def _hide(self):
        if self._popup is not None and self._popup.winfo_exists():
            self._popup.withdraw()

    def _destroy_popup(self):
        if self._popup is not None and self._popup.winfo_exists():
            self._popup.destroy()
        self._popup = None

    def _maybe_hide(self):
        try:
            focus = self.focus_get()
        except (KeyError, tk.TclError):
            focus = None
        if focus is self or (self._lb is not None and focus is self._lb):
            return
        self._hide()

    def _set_index(self, i):
        lb = self._lb
        n = lb.size()
        if not n:
            return
        i = max(0, min(n - 1, i))
        lb.selection_clear(0, "end")
        lb.selection_set(i)
        lb.activate(i)
        lb.see(i)

    def _choose(self, text):
        self._hide()
        self.set(text)
        self.icursor("end")
        self.selection_clear()
        self.configure(values=self._all)
        self.focus_set()
        self.event_generate("<<ComboboxSelected>>")

    # ---------------- الأحداث
    def _on_native_post(self):
        """زر السهم: تعرض القائمة الأصلية نتائج النص المكتوب (أو الكل). القيم ثابتة طوال فتحها."""
        self._hide()
        text = self.get().strip()
        self.configure(values=self._all if (not text or text in self._map) else (self._matches(text) or self._all))

    def _on_key(self, event):
        if event.keysym in self.NAV_KEYS:
            return
        text = self.get().strip()
        if not text:
            self.configure(values=self._all)
            self._hide()
            return
        found = self._matches(text)
        self.configure(values=found or self._all)
        if len(found) == 1 and found[0] == text:
            self._hide()
            return
        self._show(found)

    def _move(self, step):
        if self._popup_visible():
            cur = self._lb.curselection()
            self._set_index((cur[0] if cur else -1) + step)
            return "break"
        if step > 0 and self.get().strip() and self.get().strip() not in self._map:
            self._show(self._matches(self.get().strip()))
            return "break"
        return None  # السلوك الأصلي (فتح القائمة الكاملة)

    def _on_return(self, event=None):
        if self._popup_visible():
            cur = self._lb.curselection()
            if cur:
                self._choose(self._lb.get(cur[0]))
                return "break"
        text = self.get().strip()
        if text and text not in self._map:
            found = self._matches(text)
            if len(found) == 1:
                self._choose(found[0])
        return None

    def _on_escape(self, event=None):
        if self._popup_visible():
            self._hide()
            return "break"
        return None

    def _on_tab(self, event=None):
        if self._popup_visible():
            cur = self._lb.curselection()
            if cur:
                self._choose(self._lb.get(cur[0]))
        return None

    def _on_click(self, event):
        i = self._lb.nearest(event.y)
        if i >= 0:
            self._choose(self._lb.get(i))

    def _on_hover(self, event):
        i = self._lb.nearest(event.y)
        if i >= 0:
            self._set_index(i)


class SortableTree(ttk.Treeview):
    """جدول تفاعلي: النقر على عنوان أي عمود يرتّب (أبجدياً للنصوص، تصاعدياً/تنازلياً للأرقام) مع إبقاء القيم
    الصفرية والفارغة في النهاية دائماً. الترتيب المختار يُعاد تطبيقه تلقائياً بعد كل تحديث للبيانات،
    وأعمدة الأسماء (الصنف/الاسم/البيان...) تتسع تلقائياً لعرض النص كاملاً بدون قص."""
    NAME_HINTS = ("الصنف", "اسم", "الاسم", "البيان", "الطرف", "العميل", "المورد", "الحساب", "الحركة")
    EMPTY = ("", "-", "—", "🔒", "None")

    def setup(self, columns):
        self._titles = {c[0]: c[1] for c in columns}
        self._base_width = {c[0]: c[2] for c in columns}
        self._anchors = {c[0]: (c[3] if len(c) > 3 else "center") for c in columns}
        self._sort = None
        self._pending = False
        for key, title, *_ in columns:
            # العنوان بنفس محاذاة بياناته: النص تحت عنوانه مباشرة (لا عنوان في الوسط وبيانات على الطرف)
            self.heading(key, text=title, anchor=self._anchors[key], command=lambda k=key: self.sort_by(k))
        self._tip = None
        self.bind("<Motion>", self._tip_motion, add="+")
        self.bind("<Leave>", lambda e: self._tip_hide(), add="+")

    def insert(self, *args, **kwargs):
        iid = super().insert(*args, **kwargs)
        if not getattr(self, "_pending", True):
            self._pending = True
            self.after_idle(self._after_fill)
        return iid

    def _after_fill(self):
        self._pending = False
        try:
            if self._sort:
                self._apply(*self._sort)
            self._autofit()
        except tk.TclError:
            pass

    def sort_by(self, col):
        desc = bool(self._sort and self._sort[0] == col and not self._sort[1])
        self._sort = (col, desc)
        self._apply(col, desc)

    @classmethod
    def _parse(cls, v):
        s_ = str(v).strip()
        if s_ in cls.EMPTY:
            return None, None
        tok = s_.replace(",", "").replace("%", "").replace("+", "").split()
        if tok:
            t = normalize_digits(tok[0])
            if t.startswith("(") and t.endswith(")"):
                t = "-" + t[1:-1]
            try:
                return float(t), None
            except ValueError:
                pass
        return None, search_norm(s_)

    def _apply(self, col, desc):
        items = list(self.get_children(""))
        parsed = [(self._parse(self.set(i, col)), i) for i in items]
        filled = [x for x in parsed if x[0] != (None, None)]
        is_num = bool(filled) and all(x[0][0] is not None for x in filled)
        if is_num:
            nonzero = sorted([x for x in filled if abs(x[0][0]) > 1e-9], key=lambda x: x[0][0], reverse=desc)
            tail = [x for x in filled if abs(x[0][0]) <= 1e-9]
        else:
            nonzero = sorted(filled, key=lambda x: (x[0][1] if x[0][1] is not None else f"{x[0][0]:020.4f}"), reverse=desc)
            tail = []
        ordered = nonzero + tail + [x for x in parsed if x[0] == (None, None)]
        for idx, (_, i) in enumerate(ordered):
            self.move(i, "", idx)
        for key, title in self._titles.items():
            arrow = (" ▼" if desc else " ▲") if key == col else ""
            self.heading(key, text=title + arrow)

    TEXT_MAX_WIDTH = 900   # أقصى عرض لعمود نصي؛ ما يزيد يظهر كاملاً في تلميح عند المرور بالماوس
    NUM_MAX_WIDTH = 260

    def _fonts(self):
        try:
            return tkfont.Font(family=UI_FONT, size=10), tkfont.Font(family=UI_FONT, size=10, weight="bold")
        except Exception:
            return None, None

    def _autofit(self):
        """كل عمود يتسع لأطول قيمة فيه ولعنوانه (لا قص للنصوص أو الأرقام أو العناوين)."""
        font, bold = self._fonts()
        if font is None:
            return
        items = self.get_children("")[:500]
        for key, title in self._titles.items():
            values = [str(self.set(i, key)) for i in items]
            longest = max((font.measure(v) for v in values), default=0)
            head = bold.measure(title + " ▼") + 24
            is_text = any(h in title for h in self.NAME_HINTS) or self._anchors.get(key) in ("e", "w")
            cap = self.TEXT_MAX_WIDTH if is_text else self.NUM_MAX_WIDTH
            width = min(max(self._base_width.get(key, 80), longest + 26, head), max(cap, head))
            self.column(key, width=width, minwidth=min(width, max(head, 60)), stretch=False)

    def _tip_hide(self):
        if self._tip is not None:
            try:
                self._tip.destroy()
            except tk.TclError:
                pass
            self._tip = None

    def _tip_motion(self, event):
        """تلميح بالنص الكامل للخلية عندما يكون أطول من عرض العمود."""
        try:
            row, col = self.identify_row(event.y), self.identify_column(event.x)
            if not row or not col:
                return self._tip_hide()
            cols = self["displaycolumns"]
            cols = self["columns"] if cols in ("#all", ("#all",)) else cols
            key = cols[int(col[1:]) - 1]
            text = str(self.set(row, key))
            font, _ = self._fonts()
            if not text or font is None or font.measure(text) + 20 <= int(self.column(key, "width")):
                return self._tip_hide()
            if self._tip is None:
                self._tip = tk.Toplevel(self)
                self._tip.wm_overrideredirect(True)
                tk.Label(self._tip, text="", justify="right", bg="#0f172a", fg="#ffffff", wraplength=520,
                         font=(UI_FONT, 10), padx=8, pady=5).pack()
            self._tip.winfo_children()[0].config(text=text)
            self._tip.wm_geometry(f"+{event.x_root + 14}+{event.y_root + 16}")
        except (tk.TclError, ValueError, IndexError):
            self._tip_hide()


def make_tree(parent, columns, height=10, expand=True, rtl=False):
    """columns = [(key, heading, width, anchor)] بالترتيب المنطقي.
    rtl=True: أول عمود يظهر في أقصى اليمين (ترتيب القراءة العربي) وشريط التمرير على اليسار."""
    keys = [c[0] for c in columns]
    tree = SortableTree(parent, columns=keys, show="headings", height=height)
    for key, text, width, anchor in columns:
        tree.heading(key, text=text, anchor=anchor)
        tree.column(key, width=width, anchor=anchor)
    tree.setup(columns)
    if rtl:
        tree["displaycolumns"] = list(reversed(keys))
    vsb = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
    hsb = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
    vsb.pack(side="left" if rtl else "right", fill="y")
    hsb.pack(side="bottom", fill="x")
    tree.pack(side="right" if rtl else "left", fill="both", expand=expand)
    return tree


def clear_tree(tree):
    tree.delete(*tree.get_children())


def rtl_row(parent, row, items, slots=4, pady=4):
    """يرص الحقول من اليمين لليسار: كل عنصر = (نص التسمية، الأداة)."""
    for i, (label, widget) in enumerate(items):
        col = (slots - 1 - i) * 2
        if label:
            ttk.Label(parent, text=label).grid(row=row, column=col + 1, sticky="e", padx=(10, 2), pady=pady)
        widget.grid(row=row, column=col, sticky="w", padx=(2, 4), pady=pady)


AR_MONTHS = ["يناير", "فبراير", "مارس", "أبريل", "مايو", "يونيو", "يوليو", "أغسطس", "سبتمبر", "أكتوبر", "نوفمبر", "ديسمبر"]
AR_WEEKDAYS = ["سبت", "أحد", "إثنين", "ثلاثاء", "أربعاء", "خميس", "جمعة"]


class CalendarPopup(tk.Toplevel):
    """تقويم منبثق خفيف (بدون مكتبات خارجية) يملأ حقل التاريخ المرتبط به."""

    def __init__(self, entry):
        super().__init__(entry.winfo_toplevel())
        self.entry = entry
        self.overrideredirect(True)
        self.configure(bg=UI["primary_dk"], padx=1, pady=1)
        cur = parse_date(entry.get()) if entry.get().strip() else None
        d = datetime.strptime(cur, "%Y-%m-%d").date() if cur else date.today()
        self.sel = d if cur else None
        self.year, self.month = d.year, d.month
        self.body = tk.Frame(self, bg="white")
        self.body.pack(fill="both", expand=True)
        self._draw()
        self.update_idletasks()
        x, y = entry.winfo_rootx(), entry.winfo_rooty() + entry.winfo_height() + 2
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        w, h = self.winfo_reqwidth(), self.winfo_reqheight()
        x = max(0, min(x + entry.winfo_width() - w, sw - w))
        if y + h > sh:
            y = entry.winfo_rooty() - h - 2
        self.geometry(f"+{x}+{y}")
        self.bind("<Escape>", lambda e: self.close())
        self.bind("<FocusOut>", self._focus_out)
        self.after(10, self._grab)

    def _grab(self):
        try:
            self.focus_force()
            self.grab_set()
        except tk.TclError:
            pass
        self.bind("<Button-1>", self._click_outside, add="+")

    def _click_outside(self, e):
        x, y = e.x_root, e.y_root
        if not (self.winfo_rootx() <= x <= self.winfo_rootx() + self.winfo_width()
                and self.winfo_rooty() <= y <= self.winfo_rooty() + self.winfo_height()):
            self.close()

    def _focus_out(self, e):
        if e.widget is self:
            self.after(50, lambda: self.focus_get() is None and self.close())

    def close(self):
        try:
            self.grab_release()
            self.destroy()
        except tk.TclError:
            pass

    def _nav(self, months=0, years=0):
        m = self.month - 1 + months
        self.year += years + m // 12
        self.month = m % 12 + 1
        self._draw()

    def _pick(self, d):
        self.entry.set_date(d)
        self.close()
        try:
            self.entry.focus_set()
            self.entry.event_generate("<<DateSelected>>")
        except tk.TclError:
            pass

    def _draw(self):
        import calendar
        for w in self.body.winfo_children():
            w.destroy()
        head = tk.Frame(self.body, bg=UI["primary"])
        head.pack(fill="x")
        btn = dict(bg=UI["primary"], fg="white", relief="flat", bd=0, cursor="hand2", font=(UI_FONT, 10, "bold"),
                   activebackground=UI["primary_dk"], activeforeground="white", padx=6)
        tk.Button(head, text="«", command=lambda: self._nav(years=-1), **btn).pack(side="left")
        tk.Button(head, text="‹", command=lambda: self._nav(months=-1), **btn).pack(side="left")
        tk.Button(head, text="»", command=lambda: self._nav(years=1), **btn).pack(side="right")
        tk.Button(head, text="›", command=lambda: self._nav(months=1), **btn).pack(side="right")
        tk.Label(head, text=f"{AR_MONTHS[self.month - 1]} {self.year}", bg=UI["primary"], fg="white",
                 font=(UI_FONT, 10, "bold")).pack(side="top", pady=4)
        grid = tk.Frame(self.body, bg="white", padx=4, pady=4)
        grid.pack()
        # الأسبوع يبدأ السبت، ويُعرض من اليمين لليسار
        cols = 7
        for i, wd in enumerate(AR_WEEKDAYS):
            tk.Label(grid, text=wd, bg="white", fg=UI["muted"], font=(UI_FONT, 8, "bold"), width=5).grid(
                row=0, column=cols - 1 - i)
        cal = calendar.Calendar(firstweekday=5)
        td = date.today()
        for r, week in enumerate(cal.monthdatescalendar(self.year, self.month), start=1):
            for i, d in enumerate(week):
                other = d.month != self.month
                is_sel, is_today = d == self.sel, d == td
                bg = UI["teal"] if is_sel else ("#dbeafe" if is_today else "white")
                fg = "white" if is_sel else ("#cbd5e1" if other else (UI["ink"] if i != 6 else "#dc2626"))
                b = tk.Button(grid, text=str(d.day), width=4, relief="flat", bd=0, bg=bg, fg=fg, cursor="hand2",
                              font=(UI_FONT, 9, "bold" if is_today or is_sel else "normal"),
                              activebackground=UI["teal2"], activeforeground="white",
                              command=lambda d=d: self._pick(d))
                b.grid(row=r, column=cols - 1 - i, padx=1, pady=1)
        foot = tk.Frame(self.body, bg="white")
        foot.pack(fill="x", padx=4, pady=(0, 4))
        tk.Button(foot, text="اليوم", command=lambda: self._pick(td), bg="#e2e8f0", relief="flat", bd=0,
                  cursor="hand2", font=(UI_FONT, 9), padx=8).pack(side="right")
        if self.entry.allow_empty:
            tk.Button(foot, text="مسح", command=lambda: (self.entry.delete(0, tk.END), self.close()), bg="#fee2e2",
                      relief="flat", bd=0, cursor="hand2", font=(UI_FONT, 9), padx=8).pack(side="left")


class DateEntry(ttk.Entry):
    """حقل تاريخ يقبل الكتابة اليدوية (YYYY-MM-DD) مع أيقونة تقويم 📅 داخل الحقل لاختيار التاريخ بنقرة.
    يعمل كـ ttk.Entry تماماً (get/insert/delete/grid) فلا يحتاج أي تعديل في الشاشات الحالية.
    فتح التقويم: النقر على الأيقونة، أو Alt+↓ / F4 داخل الحقل."""

    def __init__(self, parent, width=13, value=None, allow_empty=True, **kw):
        super().__init__(parent, width=width + 3, **kw)
        self.allow_empty = allow_empty
        if value:
            self.insert(0, value)
        self._icon = tk.Label(self, text="📅", cursor="hand2", bg="white", fg=UI["primary"], font=(UI_FONT, 9), bd=0)
        self._icon.place(relx=1.0, x=-2, rely=0.5, anchor="e")
        self._icon.bind("<Button-1>", lambda e: self.open_calendar())
        for seq in ("<Alt-Down>", "<F4>"):
            self.bind(seq, lambda e: (self.open_calendar(), "break")[1])
        self.bind("<FocusOut>", self._normalize, add="+")
        self._popup = None

    def configure(self, cnf=None, **kw):
        st_ = kw.get("state")
        if st_ is None and isinstance(cnf, dict):
            st_ = cnf.get("state")
        res = super().configure(cnf, **kw)
        if st_ is not None and hasattr(self, "_icon"):
            try:
                self._icon.configure(fg=UI["primary"] if str(st_) == "normal" else "#94a3b8")
            except tk.TclError:
                pass
        return res

    config = configure

    def open_calendar(self):
        if str(self.cget("state")) in ("disabled", "readonly"):
            return
        if self._popup is not None:
            try:
                if self._popup.winfo_exists():
                    self._popup.close()
                    self._popup = None
                    return
            except tk.TclError:
                pass
        self._popup = CalendarPopup(self)

    def set_date(self, d):
        v = d.strftime("%Y-%m-%d") if hasattr(d, "strftime") else str(d)
        restore = str(self.cget("state"))
        if restore != "normal":
            super().configure(state="normal")
        self.delete(0, tk.END)
        self.insert(0, v)
        if restore != "normal":
            super().configure(state=restore)

    def _normalize(self, e=None):
        """يحوّل الصيغ المقبولة (مثل 5/3/2026 أو ٢٠٢٦-٠٣-٠٥) إلى YYYY-MM-DD عند مغادرة الحقل."""
        v = self.get().strip()
        if v and str(self.cget("state")) == "normal":
            d = parse_date(v)
            if not d:
                t_ = normalize_digits(v).strip().replace(".", "/").replace("-", "/")
                for f_ in ("%Y/%m/%d", "%d/%m/%Y", "%d/%m/%y"):
                    try:
                        d = datetime.strptime(t_, f_).strftime("%Y-%m-%d")
                        break
                    except ValueError:
                        continue
            if d and d != v:
                self.delete(0, tk.END)
                self.insert(0, d)


def date_entry(parent, width=13, value=None):
    """حقل تاريخ بقيمة افتراضية = اليوم (أو value) مع منتقي تقويم."""
    return DateEntry(parent, width=width, value=value or today(), allow_empty=False)


# ============================================================
# الهوية البصرية لواجهة سطح المكتب: أزرق نابض + فيروزي + أردوازي داكن/فاتح
# ============================================================
UI = dict(bg="#edf2f9", card="#ffffff", border="#c9d6ea", ink="#0f172a", muted="#5b6b82",
          primary="#2563eb", primary_dk="#1e40af", teal="#0d9488", teal2="#14b8a6", cyan="#06b6d4",
          slate="#0b1220", slate2="#111c33", slate3="#1e293b", head="#1e3a8a", sel="#14b8a6")
# ألوان الأزرار القديمة ← لوحة الألوان الجديدة (مع الإبقاء على الأحمر للحذف والبرتقالي للتعديل)
BUTTON_REMAP = {"#16a34a": "#0d9488", "#0284c7": "#0891b2", "#ea580c": "#2563eb", "#64748b": "#475569"}
UI_FONT = "Segoe UI"


def _shade(color, t):
    try:
        return _hex_mix(color, "#000000", t)
    except (ValueError, TypeError, IndexError):
        return color


def money_button(parent, text, color, command, size=10):
    """زر مسطح حديث بحواف ناعمة وتأثير تحويم (يغمق اللون عند المرور)."""
    color = BUTTON_REMAP.get(str(color).lower(), color)
    hover = _shade(color, 0.18)
    b = tk.Button(parent, text=text, bg=color, fg="white", font=(UI_FONT, size, "bold"), command=command,
                  cursor="hand2", relief="flat", bd=0, padx=14, pady=6, highlightthickness=0,
                  activebackground=_shade(color, 0.3), activeforeground="white")
    b.bind("<Enter>", lambda e: b.configure(bg=hover) if str(b["state"]) != "disabled" else None, add="+")
    b.bind("<Leave>", lambda e: b.configure(bg=color), add="+")
    return b


def apply_theme(root):
    """أنماط ttk موحدة: بطاقات بحدود ناعمة، جداول بعناوين زرقاء داكنة وصفوف مريحة، تبويبات ملونة."""
    style = ttk.Style(root)
    style.theme_use("clam")
    for cls in ("Frame", "Label", "Checkbutton", "Radiobutton", "Labelframe"):
        root.option_add(f"*{cls}.background", UI["bg"])
    root.option_add("*Checkbutton.activeBackground", UI["bg"])
    root.option_add("*Radiobutton.activeBackground", UI["bg"])
    root.option_add("*Label.foreground", UI["ink"])
    base = (UI_FONT, 10)
    style.configure(".", background=UI["bg"], foreground=UI["ink"], font=base, bordercolor=UI["border"],
                    lightcolor=UI["bg"], darkcolor=UI["border"], troughcolor=UI["bg"], focuscolor=UI["teal2"])
    style.configure("TFrame", background=UI["bg"])
    style.configure("TLabel", background=UI["bg"], foreground=UI["ink"])
    style.configure("TLabelframe", background=UI["bg"], bordercolor=UI["border"], lightcolor="#ffffff",
                    darkcolor="#9fb3d1", relief="groove", borderwidth=2)
    style.configure("TLabelframe.Label", background=UI["bg"], foreground=UI["primary_dk"], font=(UI_FONT, 10, "bold"))
    style.configure("TNotebook", background=UI["bg"], borderwidth=0, tabmargins=(6, 6, 6, 0))
    style.configure("TNotebook.Tab", background="#d8e3f3", foreground=UI["slate3"], padding=(16, 7),
                    font=(UI_FONT, 10, "bold"), borderwidth=0)
    style.map("TNotebook.Tab", background=[("selected", UI["primary"]), ("active", "#bfdbfe")],
              foreground=[("selected", "#ffffff")], expand=[("selected", (1, 2, 1, 0))])
    style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground=UI["ink"], rowheight=28,
                    bordercolor=UI["border"], font=base)
    style.configure("Treeview.Heading", background=UI["head"], foreground="#ffffff", font=(UI_FONT, 10, "bold"),
                    relief="flat", padding=(6, 7), bordercolor=UI["head"])
    style.map("Treeview.Heading", background=[("active", UI["primary"])])
    style.map("Treeview", background=[("selected", UI["sel"])], foreground=[("selected", "#ffffff")])
    style.configure("TEntry", fieldbackground="#ffffff", bordercolor=UI["border"], lightcolor="#ffffff", padding=5)
    style.map("TEntry", bordercolor=[("focus", UI["teal2"])], lightcolor=[("focus", UI["teal2"])])
    style.configure("TCombobox", fieldbackground="#ffffff", bordercolor=UI["border"], padding=4,
                    arrowcolor=UI["primary"], background="#ffffff")
    style.map("TCombobox", bordercolor=[("focus", UI["teal2"])], fieldbackground=[("readonly", "#ffffff")],
              selectbackground=[("readonly", "#ffffff")], selectforeground=[("readonly", UI["ink"])])
    style.configure("TButton", background=UI["primary"], foreground="#ffffff", font=(UI_FONT, 10, "bold"),
                    borderwidth=0, padding=(12, 6))
    style.map("TButton", background=[("active", UI["primary_dk"]), ("disabled", "#94a3b8")])
    style.configure("TCheckbutton", background=UI["bg"], foreground=UI["ink"])
    style.map("TCheckbutton", background=[("active", UI["bg"])], indicatorcolor=[("selected", UI["teal"])])
    style.configure("TRadiobutton", background=UI["bg"], foreground=UI["ink"])
    style.map("TRadiobutton", background=[("active", UI["bg"])], indicatorcolor=[("selected", UI["teal"])])
    style.configure("Vertical.TScrollbar", background="#c3d0e4", troughcolor=UI["bg"], bordercolor=UI["bg"],
                    arrowcolor=UI["slate3"])
    style.configure("Horizontal.TScrollbar", background="#c3d0e4", troughcolor=UI["bg"], bordercolor=UI["bg"],
                    arrowcolor=UI["slate3"])
    style.layout("NoTabs.TNotebook.Tab", [])
    style.configure("NoTabs.TNotebook", background=UI["bg"], borderwidth=0)
    return style


def handle_error(e):
    if isinstance(e, UserError):
        messagebox.showwarning("تنبيه", str(e))
    else:
        messagebox.showerror("خطأ", f"{type(e).__name__}: {e}")


# ============================================================
# عناصر واجهة إضافية: سلة أوامر، منطقة تمرير، رسم بياني، جداول ديناميكية
# ============================================================
class ScrollFrame(tk.Frame):
    """إطار قابل للتمرير رأسياً؛ ضع العناصر داخل .inner"""

    def __init__(self, parent, height=220):
        super().__init__(parent)
        self.canvas = tk.Canvas(self, height=height, highlightthickness=0)
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.canvas.yview)
        self.inner = tk.Frame(self.canvas)
        self.inner.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.create_window((0, 0), window=self.inner, anchor="nw")
        self.canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)


# أعمدة بنود الفواتير بالترتيب المنطقي (تُعرض من اليمين لليسار عبر make_tree(..., rtl=True)):
# الكود ← الصنف ← الكمية ← الوحدة ← سعر الوحدة ← الإجمالي قبل الخصم ← الخصم الخاص ← المبلغ الخاضع للضريبة ← الضريبة ← الإجمالي
CART_COLUMNS = [("code", "الكود", 80, "center"), ("name", "الصنف", 250, "e"), ("qty", "الكمية", 75, "center"),
                ("unit", "الوحدة", 60, "center"), ("price", "سعر الوحدة", 90, "center"),
                ("subtotal", "المبلغ الخاضع للضريبة", 140, "center"), ("vat", "الضريبة 15%", 95, "center"),
                ("total", "الإجمالي", 105, "center")]
SALES_CART_COLUMNS = [("code", "الكود", 75, "center"), ("name", "الصنف", 210, "e"), ("qty", "الكمية", 65, "center"),
                      ("unit", "الوحدة", 55, "center"), ("price", "سعر الوحدة", 80, "center"),
                      ("gross", "الإجمالي قبل الخصم", 120, "center"), ("discount", "الخصم الخاص", 90, "center"),
                      ("subtotal", "المبلغ الخاضع للضريبة", 135, "center"), ("vat", "الضريبة 15%", 85, "center"),
                      ("total", "الإجمالي", 100, "center")]
PURCHASE_LINE_COLUMNS = [("code", "الكود", 75, "center"), ("name", "الصنف", 210, "e"), ("qty", "الكمية", 65, "center"),
                         ("unit", "الوحدة", 55, "center"), ("price", "تكلفة الوحدة", 85, "center"),
                         ("subtotal", "المبلغ الخاضع للضريبة", 135, "center"), ("vat", "الضريبة 15%", 85, "center"),
                         ("total", "الإجمالي", 100, "center"), ("production", "تاريخ الإنتاج", 95, "center"),
                         ("expiry", "تاريخ الانتهاء", 95, "center")]


def cart_row_values(columns, l):
    """قيم صف السلة بحسب أعمدة الجدول (يدعم السلة العادية وسلة المبيعات بالخصم)."""
    disc = l.get("discount", 0.0)
    cells = {"code": l["code"], "name": l["name"], "qty": f"{l['qty']:g}", "unit": l["unit"], "price": fmt(l["price"]),
             "gross": fmt(l.get("gross", l["subtotal"])), "discount": fmt(disc) if disc > 0.005 else "-",
             "subtotal": fmt(l["subtotal"]), "vat": fmt(l["vat"]), "total": fmt(l["total"]),
             "production": l.get("production") or "-", "expiry": l.get("expiry") or "-"}
    return [cells[c] for c in columns]


def cart_summary(lines):
    sub = sum(l["subtotal"] for l in lines)
    vat = sum(l["vat"] for l in lines)
    disc = sum(l.get("discount", 0.0) for l in lines)
    parts = [f"عدد الأصناف: {len(lines)}"]
    if disc > 0.005:
        parts += [f"قبل الخصم: {fmt(sub + disc)}", f"الخصم الخاص: {fmt(disc)}"]
    parts += [f"الخاضع للضريبة: {fmt(sub)}", f"الضريبة: {fmt(vat)}", f"الإجمالي: {fmt(sub + vat)} ر.س"]
    return "   |   ".join(parts)


class CartPanel:
    """سلة بنود قابلة لإعادة الاستخدام (أوامر البيع والشراء): اختيار صنف + كمية + سعر (+ خصم خاص للمبيعات) ثم إضافة."""

    def __init__(self, parent, price_label="السعر", warn_below_cost=False, allow_discount=False):
        self.lines = []
        self.warn_below_cost = warn_below_cost
        self.allow_discount = allow_discount
        f = tk.Frame(parent)
        f.pack(fill="x")
        self.item_cb = SearchableCombobox(f, width=38)
        self.item_cb.bind("<<ComboboxSelected>>", self.on_item)
        self.qty_e = ttk.Entry(f, width=10)
        self.price_e = ttk.Entry(f, width=12)
        self.disc_e = ttk.Entry(f, width=10)
        add_btn = money_button(f, "➕ إضافة للأمر", "#0284c7", self.add_line, 9)
        items = [("الصنف:", self.item_cb), ("الكمية:", self.qty_e), (price_label + ":", self.price_e)]
        if allow_discount:
            items.append(("خصم خاص (مبلغ أو %):", self.disc_e))
        items.append((None, add_btn))
        rtl_row(f, 0, items, slots=len(items))
        self.info_lbl = ttk.Label(f, text="", foreground="#0369a1")
        self.info_lbl.grid(row=1, column=0, columnspan=2 * len(items), sticky="e", padx=10)
        box = ttk.LabelFrame(parent, text=" بنود الأمر ", padding=6)
        box.pack(fill="x", pady=4)
        base = SALES_CART_COLUMNS if allow_discount else CART_COLUMNS
        cols = [(k, ("تكلفة الوحدة" if k == "price" and price_label == "التكلفة" else h), w, a) for k, h, w, a in base]
        self.tree = make_tree(box, cols, height=4, rtl=True)
        bar = tk.Frame(parent)
        bar.pack(fill="x")
        money_button(bar, "➖ حذف السطر المحدد", "#dc2626", self.remove_line, 9).pack(side="left", padx=4)
        self.total_lbl = tk.Label(bar, text="", font=("Arial", 10, "bold"), bg="#e2e8f0", pady=5)
        self.total_lbl.pack(side="right", fill="x", expand=True, padx=4)
        self.refresh()

    def set_items(self, pairs):
        self.item_cb.set_items(pairs)

    def on_item(self, event=None):
        item = self.item_cb.selected()
        if not item:
            return
        with db() as conn:
            row = conn.execute("SELECT quantity, cost_price FROM inventory WHERE item_code=?", (item[0],)).fetchone()
        if row:
            self.info_lbl.config(text=f"المتاح بالمخزن: {row[0]:g} {item[2] or ''}   |   متوسط التكلفة: {fmt(row[1])}")

    def add_line(self):
        try:
            item = self.item_cb.selected()
            if not item:
                raise UserError("اختر الصنف من القائمة.")
            qty, price = parse_float(self.qty_e.get()), parse_float(self.price_e.get())
            if qty is None or qty <= 0 or price is None or price <= 0:
                raise UserError("أدخل كمية وسعراً صحيحين أكبر من الصفر.")
            disc = parse_discount(self.disc_e.get(), round(qty * price, 2)) if self.allow_discount else 0.0
            line = make_line(item[0], item[1], item[2], qty, price, disc)
            if self.warn_below_cost:
                with db() as conn:
                    cost = conn.execute("SELECT cost_price FROM inventory WHERE item_code=?", (item[0],)).fetchone()[0] or 0.0
                net_unit = line["subtotal"] / line["qty"]
                if net_unit + EPS < cost and not messagebox.askyesno(
                        "تحذير: السعر أقل من التكلفة",
                        f"سعر البيع بعد الخصم ({fmt(net_unit)}) أقل من تكلفة الصنف ({fmt(cost)}).\n\nهل تريد الاستمرار؟"):
                    return
            self.lines.append(line)
            for e in (self.qty_e, self.price_e, self.disc_e):
                e.delete(0, tk.END)
            self.refresh()
        except Exception as e:
            handle_error(e)

    def remove_line(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر السطر المراد حذفه.")
            return
        del self.lines[int(sel[0])]
        self.refresh()

    def clear(self):
        self.lines.clear()
        self.info_lbl.config(text="")
        self.refresh()

    def refresh(self):
        clear_tree(self.tree)
        cols = self.tree["columns"]
        for i, l in enumerate(self.lines):
            self.tree.insert("", "end", iid=str(i), values=cart_row_values(cols, l))
        self.total_lbl.config(text=cart_summary(self.lines))


def configure_tree(tree, columns, highlight=()):
    """يعيد تهيئة أعمدة جدول موجود بحسب تعريف التقرير [(العنوان، النوع)].
    highlight: فهارس أعمدة مميزة (جدول Tk لا يلوّن عموداً منفرداً، فيُميَّز عنوانها بـ 🟨)."""
    widths = {"t": 190, "m": 115, "q": 95, "p": 85, "i": 80}
    keys = [f"c{i}" for i in range(len(columns))]
    tree.delete(*tree.get_children())
    tree.configure(columns=keys)
    spec = []
    for i, (key, (head, kind)) in enumerate(zip(keys, columns)):
        head = f"🟨 {head}" if i in highlight else head
        anchor = "e" if kind == "t" else "center"
        tree.heading(key, text=head, anchor=anchor)
        tree.column(key, width=widths[kind], anchor=anchor, stretch=True)
        spec.append((key, head, widths[kind], anchor))
    if isinstance(tree, SortableTree):   # ترتيب بالنقر + اتساع تلقائي + تلميح النص الكامل
        tree.setup(spec)


def draw_bar_chart(canvas, labels, values, title=""):
    """رسم أعمدة بسيط على Canvas (حتى 15 عموداً)."""
    canvas.delete("all")
    if not labels:
        return
    labels, values = labels[:15], values[:15]
    w = max(canvas.winfo_width(), 600)
    h = max(canvas.winfo_height(), 150)
    pad_l, pad_b, pad_t = 10, 38, 22
    vmax, vmin = max(max(values), 0), min(min(values), 0)
    span = (vmax - vmin) or 1
    plot_h = h - pad_b - pad_t
    zero_y = pad_t + plot_h * (vmax / span)
    slot = (w - 2 * pad_l) / len(values)
    canvas.create_text(w - 10, 10, text=title, anchor="e", font=("Arial", 9, "bold"), fill="#475569")
    canvas.create_line(pad_l, zero_y, w - pad_l, zero_y, fill="#94a3b8")
    for i, (lab, v) in enumerate(zip(labels, values)):
        x0 = pad_l + i * slot + slot * 0.2
        x1 = pad_l + (i + 1) * slot - slot * 0.2
        y = pad_t + plot_h * ((vmax - v) / span)
        canvas.create_rectangle(x0, min(y, zero_y), x1, max(y, zero_y), fill="#ea580c" if v >= 0 else "#dc2626", outline="")
        canvas.create_text((x0 + x1) / 2, min(y, zero_y) - 8, text=f"{v:,.0f}", font=("Arial", 7), fill="#334155")
        canvas.create_text((x0 + x1) / 2, h - pad_b + 12, text=str(lab)[:12], font=("Arial", 7), fill="#334155")


def short_money(v):
    a = abs(v)
    ar = CHART_PAL["lang"] == "ar"
    if a >= 1e6:
        return f"{v / 1e6:,.2f} " + ("مليون" if ar else "M")
    if a >= 1e4:
        return f"{v / 1e3:,.1f} " + ("ألف" if ar else "K")
    return f"{v:,.0f}"


def _chart_empty(cv, text=None):
    if text is None:
        text = "لا توجد بيانات بعد" if CHART_PAL["lang"] == "ar" else "No data yet"
    w, h = max(cv.winfo_width(), 200), max(cv.winfo_height(), 120)
    cv.create_text(w / 2, h / 2, text=text, font=("Arial", 10), fill=CHART_PAL["text"])


def _chart_tooltip(cv, item, text):
    """تلميح يظهر عند المرور بالماوس على عمود (بألوان الوضع الحالي وقت الرسم)."""
    bg, fg = CHART_PAL["tip_bg"], CHART_PAL["tip_fg"]

    def enter(e):
        cv.delete("tip")
        x = min(max(e.x, 90), max(cv.winfo_width() - 90, 90))
        t = cv.create_text(x, max(e.y - 12, 20), text=text, font=("Arial", 9, "bold"), fill=fg, anchor="s", tags="tip")
        x0, y0, x1, y1 = cv.bbox(t)
        cv.create_rectangle(x0 - 7, y0 - 4, x1 + 7, y1 + 4, fill=bg, outline="", tags="tip")
        cv.tag_raise(t)
    cv.tag_bind(item, "<Enter>", enter)
    cv.tag_bind(item, "<Leave>", lambda e: cv.delete("tip"))


def draw_grouped_bars(cv, labels, series):
    """أعمدة مجمعة بخطوط شبكة ومفتاح ألوان. series = [(الاسم، اللون، القيم)]."""
    P = CHART_PAL
    cv.delete("all")
    allv = [v for _, _, vals in series for v in vals]
    if not labels or not any(abs(v) > EPS for v in allv):
        _chart_empty(cv)
        return
    w, h = max(cv.winfo_width(), 260), max(cv.winfo_height(), 160)
    pad_l, pad_r, pad_t, pad_b = 62, 12, 34, 26
    vmax, vmin = max(max(allv), 0.0), min(min(allv), 0.0)
    span = (vmax - vmin) or 1.0
    ph, pw = h - pad_t - pad_b, w - pad_l - pad_r

    def y(v):
        return pad_t + ph * (vmax - v) / span

    for i in range(5):
        v = vmin + span * i / 4
        cv.create_line(pad_l, y(v), w - pad_r, y(v), fill=P["grid"])
        cv.create_text(pad_l - 6, y(v), text=short_money(v), anchor="e", font=("Arial", 7), fill=P["text"])
    n, ns = len(labels), len(series)
    slot = pw / n
    bw = slot * 0.72 / ns
    for i, lab in enumerate(labels):
        x0 = pad_l + i * slot + slot * 0.14
        for j, (name, color, vals) in enumerate(series):
            v = vals[i]
            xa = x0 + j * bw
            r = cv.create_rectangle(xa, min(y(v), y(0)), xa + bw - 3, max(y(v), y(0)),
                                    fill=color if v >= 0 else P["neg"], outline="")
            _chart_tooltip(cv, r, f"{name} - {lab}: {fmt(v)} {_cur()}")
        cv.create_text(pad_l + (i + 0.5) * slot, h - pad_b + 13, text=lab, font=("Arial", 8), fill=P["text"])
    cv.create_line(pad_l, y(0), w - pad_r, y(0), fill=P["axis"])
    if P["rtl"]:  # مفتاح الألوان من اليمين
        x = w - pad_r
        for name, color, _ in series:
            t = cv.create_text(x, 13, text=name, anchor="e", font=("Arial", 8, "bold"), fill=P["ink"])
            bx = cv.bbox(t)[0]
            cv.create_rectangle(bx - 14, 8, bx - 4, 18, fill=color, outline="")
            x = bx - 24
    else:  # مفتاح الألوان من اليسار
        x = pad_l
        for name, color, _ in series:
            cv.create_rectangle(x, 8, x + 10, 18, fill=color, outline="")
            t = cv.create_text(x + 14, 13, text=name, anchor="w", font=("Arial", 8, "bold"), fill=P["ink"])
            x = cv.bbox(t)[2] + 14


def draw_hbars(cv, items, color, max_items=7, colors=None):
    """أعمدة أفقية. بالعربي: التسمية يميناً والعمود يمتد يساراً؛ بالإنجليزي العكس. القيم السالبة بالأحمر."""
    P = CHART_PAL
    rtl = P["rtl"]
    cv.delete("all")
    items = items[:max_items]
    if not items:
        _chart_empty(cv)
        return
    w, h = max(cv.winfo_width(), 260), max(cv.winfo_height(), 160)
    label_w = min(170, w * 0.42)
    pad = 8
    vals = [v for _, v in items]
    vmax, vmin = max(max(vals), 0.0), min(min(vals), 0.0)
    span = (vmax - vmin) or 1.0
    if rtl:
        plot_l, plot_r = pad + 62, w - label_w - 10
    else:
        plot_l, plot_r = label_w + 10, w - pad - 62

    def x(v):
        frac = (v - vmin) / span
        return plot_r - (plot_r - plot_l) * frac if rtl else plot_l + (plot_r - plot_l) * frac

    zero = x(0)
    row_h = min(34, (h - 2 * pad) / len(items))
    for i, (lab, v) in enumerate(items):
        yc = pad + row_h * (i + 0.5)
        text = lab if len(lab) <= 28 else lab[:27] + "…"
        cv.create_text(w - pad if rtl else pad, yc, text=text, anchor="e" if rtl else "w", font=("Arial", 8),
                       fill=P["label"])
        xa, xb = sorted((zero, x(v)))
        r = cv.create_rectangle(xa, yc - row_h * 0.3, max(xb, xa + 2), yc + row_h * 0.3,
                                fill=(colors[i] if colors else color) if v >= 0 else P["neg"], outline="")
        _chart_tooltip(cv, r, f"{lab}: {fmt(v)} {_cur()}")
        # الرقم عند طرف العمود البعيد عن خط الصفر
        outer_left = (v >= 0) == rtl
        cv.create_text(xa - 4 if outer_left else xb + 4, yc, text=short_money(v), anchor="e" if outer_left else "w",
                       font=("Arial", 8, "bold"), fill=P["ink"] if v >= 0 else P["neg"])
    cv.create_line(zero, pad, zero, pad + row_h * len(items), fill=P["axis"])


# ---------------------------------------------------------- شعار الحاجري وشاشة الدخول
# هوية «الحاجري للأنظمة المالية»: تدرجات زرقاء حديثة (كحلي ← أزرق ملكي ← سماوي)
BRAND_NAVY, BRAND_NAVY2 = "#0a1a3f", "#123a8c"
BRAND_BLUE, BRAND_SKY, BRAND_CYAN = "#1d4ed8", "#38bdf8", "#7dd3fc"
BRAND_GOLD, BRAND_GOLD2 = BRAND_SKY, BRAND_CYAN  # أسماء قديمة تُستخدم في الواجهة = لون الإبراز الأزرق السماوي


def _hex_mix(c1, c2, t):
    a = [int(c1[i:i + 2], 16) for i in (1, 3, 5)]
    b = [int(c2[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02x}" for x, y in zip(a, b))


def draw_vertical_gradient(cv, x0, y0, x1, y1, stops, step=2):
    """تدرج رأسي متعدد الألوان: stops = [لون1، لون2، ...] من الأعلى للأسفل."""
    h = max(1, int(y1 - y0))
    seg = len(stops) - 1
    for y in range(0, h, step):
        t = y / h * seg
        i = min(int(t), seg - 1)
        cv.create_rectangle(x0, y0 + y, x1, y0 + y + step, fill=_hex_mix(stops[i], stops[i + 1], t - i), outline="")


def draw_elhagry_mark(cv, cx, cy, size):
    """رمز الحاجري: مربع دائري بتدرج أزرق، عموداً حرف H يؤطّران معداداً (3 قضبان بخرزات)،
    والقضيب الأوسط هو العارضة الأفقية لحرف H."""
    s = size / 2.0
    r = s * 0.32
    # ظل ناعم
    for k, col in enumerate(("#08142f", "#0b1d45")):
        o = s * (0.06 - k * 0.025)
        cv.create_oval(cx - s + o, cy + s * 0.78 + o, cx + s + o, cy + s * 1.08 + o, fill=col, outline="")
    # المربع الدائري بتدرج قطري تقريبي (شرائح أفقية مقصوصة بالزوايا)
    stops = ["#38bdf8", "#2563eb", "#1e3a8a"]
    n = max(24, int(size))
    for i in range(n):
        y = cy - s + 2 * s * i / n
        y2 = cy - s + 2 * s * (i + 1) / n + 1
        dy = min(y - (cy - s), (cy + s) - y2)
        inset = 0.0
        if dy < r:
            inset = r - math.sqrt(max(0.0, r * r - (r - max(dy, 0)) ** 2))
        t = i / n * (len(stops) - 1)
        j = min(int(t), len(stops) - 2)
        cv.create_rectangle(cx - s + inset, y, cx + s - inset, y2, fill=_hex_mix(stops[j], stops[j + 1], t - j), outline="")
    # لمعة علوية
    cv.create_arc(cx - s * 0.95, cy - s * 1.05, cx + s * 0.95, cy + s * 0.15, start=20, extent=140, style="arc",
                  outline="#bae6fd", width=max(1, int(size / 60)))
    # عمودا H
    pw = s * 0.2
    for side in (-1, 1):
        px = cx + side * s * 0.56
        cv.create_rectangle(px - pw / 2, cy - s * 0.66, px + pw / 2, cy + s * 0.66, fill="#f8fafc", outline="")
        cv.create_rectangle(px - pw / 2, cy - s * 0.66, px + pw / 2, cy - s * 0.58, fill="#e0f2fe", outline="")
    # قضبان المعداد (الأوسط = عارضة H أعرض)
    beads = {-0.36: (2, 1), 0.0: (1, 3), 0.36: (3, 0)}
    bead_cols = ["#facc15", "#7dd3fc", "#ffffff", "#7dd3fc"]
    for off, (left_n, right_n) in beads.items():
        yy = cy + off * s
        thick = s * (0.09 if off == 0 else 0.035)
        cv.create_rectangle(cx - s * 0.46, yy - thick / 2, cx + s * 0.46, yy + thick / 2,
                            fill="#f8fafc" if off == 0 else "#bfdbfe", outline="")
        br = s * 0.085
        xs = [cx - s * 0.36 + k * br * 2.1 for k in range(left_n)] + [cx + s * 0.36 - k * br * 2.1 for k in range(right_n)]
        for k, bx in enumerate(xs):
            col = bead_cols[(k + int(off * 10)) % len(bead_cols)] if off else "#facc15"
            cv.create_oval(bx - br, yy - br * 0.8, bx + br, yy + br * 0.8, fill=col, outline="#0b1d45", width=1)


def draw_elhagry_logo(cv, cx, cy, scale=1.0):
    """شعار «الحاجري / ELHAGRY Financial Systems»: رمز المعداد داخل إطار H بتدرجات زرقاء + الاسم بالعربية والإنجليزية."""
    draw_elhagry_mark(cv, cx, cy, 120 * scale)
    cv.create_text(cx, cy + 96 * scale, text="الحاجري", font=("Arial", int(24 * scale), "bold"), fill="white")
    t = cv.create_text(cx, cy + 126 * scale, text="E L H A G R Y", font=("Segoe UI", int(12 * scale), "bold"),
                       fill=BRAND_CYAN)
    x0, _, x1, _ = cv.bbox(t)
    ly = cy + 126 * scale
    cv.create_line(x0 - 44 * scale, ly, x0 - 10, ly, fill=BRAND_SKY, width=2)
    cv.create_line(x1 + 10, ly, x1 + 44 * scale, ly, fill=BRAND_SKY, width=2)
    cv.create_text(cx, cy + 146 * scale, text="FINANCIAL SYSTEMS  •  للأنظمة المالية",
                   font=("Segoe UI", int(8 * scale), "bold"), fill="#93c5fd")


# صورة شعار «الحاجري» توضع في مجلد البرنامج بأحد هذه الأسماء. لو وُجد أكثر من ملف يُستخدم الأحدث تعديلاً.
# الملف قد يكون الشعار وحده، أو تصميم شاشة الدخول كاملاً (لوحة كحلية بها الشعار) فيُقتطع الشعار منها تلقائياً.
def load_logo_photo(max_w, max_h, names=ELHAGRY_LOGO_FILES):
    """يرجع (PhotoImage، لون خلفية الشعار) أو (None، None) لو لم يوجد ملف."""
    path = find_logo_file(names)
    if not path:
        return None, None
    with Image.open(path) as src:
        im, bg = extract_logo_image(src)
    ratio = min(max_w / im.width, max_h / im.height, 1.0)
    size = (max(1, round(im.width * ratio)), max(1, round(im.height * ratio)))
    resample = getattr(Image, "Resampling", Image).LANCZOS
    return ImageTk.PhotoImage(im.resize(size, resample)), bg


class LoginWindow:
    """شاشة تسجيل الدخول (نفس تصميم الحاجري): لوحة الشعار الكحلية يميناً ونموذج الدخول يساراً.
    لا تُفتح الواجهة الرئيسية إلا بعد التحقق من المستخدم وكلمة المرور."""
    W, H = 1000, 625
    PANEL_W = 420

    def __init__(self, root, on_success):
        self.root, self.on_success = root, on_success
        w = self.win = tk.Toplevel(root)
        w.title(f"تسجيل الدخول - {BRAND_FULL}")
        w.configure(bg=LOGIN_FORM_BG)
        w.resizable(False, False)
        sx, sy = w.winfo_screenwidth(), w.winfo_screenheight()
        W, H = min(self.W, sx - 40), min(self.H, sy - 80)
        w.geometry(f"{W}x{H}+{(sx - W) // 2}+{max((sy - H) // 2 - 20, 0)}")
        w.protocol("WM_DELETE_WINDOW", self.cancel)

        # ---- اللوحة الجانبية (يمين): صورة الشعار من ملف المشروع
        self.logo_photo = None
        panel_bg = LOGIN_PANEL_BG
        try:
            self.logo_photo, img_bg = load_logo_photo(self.PANEL_W - 90, H - 200)
            panel_bg = img_bg or panel_bg
        except Exception as e:
            print(f"تعذر تحميل صورة الشعار: {e}")
        panel = tk.Frame(w, bg=panel_bg, width=self.PANEL_W)
        panel.pack(side="right", fill="y")
        panel.pack_propagate(False)
        if self.logo_photo is not None:
            tk.Label(panel, image=self.logo_photo, bg=panel_bg, bd=0).place(relx=0.5, rely=0.46, anchor="center")
            try:
                w.iconphoto(False, self.logo_photo)
            except tk.TclError:
                pass
        else:  # الصورة غير موجودة: شعار مرسوم احتياطي
            cv = tk.Canvas(panel, width=self.PANEL_W, height=360, bg=panel_bg, highlightthickness=0)
            cv.place(relx=0.5, rely=0.46, anchor="center")
            draw_elhagry_logo(cv, self.PANEL_W / 2, 110)
        tk.Label(panel, text=f"© {BRAND_FULL}", bg=panel_bg,
                 fg="#93a3bd", font=("Segoe UI", 9)).pack(side="bottom", pady=22)

        # ---- نموذج الدخول (يسار)
        side = tk.Frame(w, bg=LOGIN_FORM_BG)
        side.pack(side="left", fill="both", expand=True)
        form = tk.Frame(side, bg=LOGIN_FORM_BG, width=400)
        form.place(relx=0.5, rely=0.5, anchor="center")
        tk.Label(form, text="تسجيل الدخول", font=("Segoe UI", 22, "bold"), bg=LOGIN_FORM_BG,
                 fg=LOGIN_PANEL_BG, anchor="e").pack(fill="x")
        tk.Label(form, text=f"أهلاً بك في {BRAND_AR}", font=("Segoe UI", 11), bg=LOGIN_FORM_BG,
                 fg="#6b7280", anchor="e").pack(fill="x", pady=(4, 22))
        self.user_e = self._field(form, "اسم المستخدم")
        self.pass_e = self._field(form, "كلمة المرور", show="•")
        self.show_var = tk.IntVar(value=0)
        tk.Checkbutton(form, text="إظهار كلمة المرور", variable=self.show_var, bg=LOGIN_FORM_BG,
                       activebackground=LOGIN_FORM_BG, fg="#4b5563", font=("Segoe UI", 9),
                       command=self._toggle_show).pack(anchor="e")
        self.err = tk.Label(form, text="", fg="#dc2626", bg=LOGIN_FORM_BG, font=("Segoe UI", 9, "bold"),
                            wraplength=400, justify="center")
        self.err.pack(fill="x", pady=(2, 2))
        self.btn = tk.Button(form, text="دخول", font=("Segoe UI", 12, "bold"), bg="#2563eb", fg="white",
                             activebackground="#1e40af", activeforeground="white", bd=0, relief="flat",
                             cursor="hand2", command=self.submit, pady=9)
        self.btn.bind("<Enter>", lambda e: self.btn.configure(bg="#1d4ed8"))
        self.btn.bind("<Leave>", lambda e: self.btn.configure(bg="#2563eb"))
        self.btn.pack(fill="x", pady=(6, 16))
        forgot = tk.Label(form, text="نسيت كلمة المرور؟", font=("Segoe UI", 9, "underline"), bg=LOGIN_FORM_BG,
                          fg="#0d9488", cursor="hand2")
        forgot.pack()
        forgot.bind("<Button-1>", lambda e: messagebox.showinfo(
            "استعادة كلمة المرور", "لأمان النظام لا تُستعاد كلمة المرور تلقائياً.\n"
                                    "تواصل مع مدير النظام لإعادة تعيينها من شاشة «إدارة المستخدمين».", parent=w))
        tk.Frame(form, bg=LOGIN_FORM_BG, width=400, height=1).pack()  # يثبّت عرض النموذج

        w.bind("<Return>", self.submit)
        w.bind("<KP_Enter>", self.submit)
        w.after(50, lambda: (w.lift(), w.focus_force(), self.user_e.focus_set()))

    def _field(self, parent, label, show=""):
        tk.Label(parent, text=label, bg=LOGIN_FORM_BG, fg="#111827", font=("Segoe UI", 10, "bold"),
                 anchor="e").pack(fill="x")
        e = tk.Entry(parent, font=("Segoe UI", 12), relief="flat", bg="white", highlightthickness=1,
                     highlightbackground="#d1d5db", highlightcolor=LOGIN_ACCENT, show=show)
        e.pack(fill="x", pady=(5, 14), ipady=8)
        return e

    def _toggle_show(self):
        self.pass_e.configure(show="" if self.show_var.get() else "•")

    def submit(self, event=None):
        try:
            user = authenticate(self.user_e.get(), self.pass_e.get())
        except UserError as e:
            self.err.config(text=str(e))
            self.pass_e.delete(0, tk.END)
            self.pass_e.focus_set()
            return
        except Exception as e:
            self.err.config(text=f"خطأ: {type(e).__name__}: {e}")
            return
        self.win.destroy()
        self.on_success(user)

    def cancel(self):
        self.win.destroy()
        self.root.destroy()


# ============================================================
# التطبيق الرئيسي
# ============================================================
class ElhagryERP(Phase2Mixin):
    def __init__(self, root, user=None, on_logout=None):
        self.root = root
        # بدون مستخدم (استخدام برمجي/اختبارات) = مدير بكل الصلاحيات
        self.user = user or dict(user_id=None, username="admin", full_name="مدير النظام", is_admin=True,
                                 perms=set(ALL_PERMISSION_KEYS))
        self.on_logout = on_logout
        set_current_user(self.user.get("username"), "desktop", self.user)   # اسم المستخدم في سجل التدقيق
        self._dash_job = None
        self.root.title(f"{BRAND_FULL}   |   المستخدم: {self.user['full_name']}")
        self.root.geometry("1400x900")
        self.root.minsize(1200, 760)

        # هوية الحاجري: أزرق نابض + فيروزي على خلفية أردوازية
        self.clr_primary = UI["primary"]
        self.clr_hover = UI["slate3"]
        self.clr_accent_yellow = UI["cyan"]
        self.clr_accent_red = "#dc2626"
        self.clr_sidebar_bg = UI["slate2"]
        self.clr_light_bg = UI["bg"]
        self.clr_text_light = "#dbe4f3"
        self.root.configure(bg=UI["bg"])
        apply_theme(self.root)
        self._active_page = None

        self.editing_partner_code = None
        self.flow = {}

        self.build_header()

        main_layout_frame = tk.Frame(self.root)
        main_layout_frame.pack(fill="both", expand=True)

        self.sidebar_frame = tk.Frame(main_layout_frame, bg=self.clr_sidebar_bg, width=236)
        self.sidebar_frame.pack(side="right", fill="y")
        self.sidebar_frame.pack_propagate(False)

        # رأس الشريط الجانبي: شعار الحاجري الرسمي (elhagry.png)
        self.side_logo = None
        logo_bg = BRAND_NAVY
        try:
            self.side_logo, logo_bg = load_logo_photo(180, 78)
        except Exception:
            self.side_logo = None
        logo_box = tk.Frame(self.sidebar_frame, bg=logo_bg or BRAND_NAVY)
        logo_box.pack(fill="x")
        if self.side_logo is not None:
            tk.Label(logo_box, image=self.side_logo, bg=logo_box["bg"], bd=0).pack(pady=(6, 2))
        else:
            cv = tk.Canvas(logo_box, width=236, height=120, bg=logo_box["bg"], highlightthickness=0)
            cv.pack()
            draw_elhagry_mark(cv, 118, 52, 70)
        tk.Label(logo_box, text=BRAND_AR, bg=logo_box["bg"], fg="#ffffff", font=(UI_FONT, 11, "bold")).pack()
        tk.Label(logo_box, text=BRAND_EN.upper(), bg=logo_box["bg"], fg=UI["cyan"], font=(UI_FONT, 7, "bold")).pack(pady=(0, 6))
        accent = tk.Canvas(self.sidebar_frame, height=4, bg=self.clr_sidebar_bg, highlightthickness=0)
        accent.pack(fill="x", pady=(0, 4))
        accent.bind("<Configure>", lambda e, c=accent: (c.delete("all"), [c.create_rectangle(
            x, 0, x + 4, 4, outline="", fill=_hex_mix(UI["primary"], UI["teal2"], x / max(e.width, 1)))
            for x in range(0, e.width, 4)]))

        content_frame = tk.Frame(main_layout_frame, bg=self.clr_light_bg)
        content_frame.pack(side="left", fill="both", expand=True)

        tk.Label(
            content_frame,
            text=f"قاعدة البيانات النشطة: {os.path.abspath(DB_FILE)}",
            anchor="w",
            fg="#78716c",
            font=("Arial", 8),
            bg=self.clr_light_bg,
            pady=5
        ).pack(side="bottom", fill="x", padx=15)

        self.build_trail_bar(content_frame)
        self.notebook = ttk.Notebook(content_frame, style='NoTabs.TNotebook')
        self.notebook.pack(expand=True, fill="both", padx=15, pady=(6, 15))

        tabs_config = [
            ("dashboard", "🏠 لوحة المؤشرات",              self.setup_dashboard_tab),
            ("customers", "👥 العملاء",                    self.setup_customers_tab),
            ("suppliers", "🏭 الموردون",                   self.setup_suppliers_tab),
            ("reps",      "🧑‍💼 سجل المناديب",               self.setup_reps_tab),
            ("import",    "📥 مركز استيراد البيانات (Excel/CSV)", self.setup_import_tab),
            ("inventory", "🏭 المخزون وكارت الصنف",     self.setup_inventory_tab),
            ("expiry",    "⏳ الصلاحية والركود",          self.setup_expiry_tab),
            ("sales",     "🛒 المبيعات (أمر ← تسليم ← فاتورة)", self.setup_sales_tab),
            ("purchase",  "📦 المشتريات (أمر ← استلام ← فاتورة)", self.setup_purchase_tab),
            ("cash",      "💰 الخزينة والبنوك والسداد",     self.setup_cash_tab),
            ("returns",   "🔄 المرتجعات",                self.setup_returns_tab),
            ("assets",    "🏢 الأصول الثابتة والإهلاك",    self.setup_assets_tab),
            ("statement", "📜 كشف حساب موحد PDF",         self.setup_statement_tab),
            ("discounts", "🏷️ الخصومات اللاحقة والإشعارات",  self.setup_discounts_tab),
            ("income",    "📈 قائمة الدخل",              self.setup_income_tab),
            ("reports",   "📊 التقارير وتحليل القرار",      self.setup_reports_tab),
            ("ledger",    "📚 الدفتر العام والقوائم المالية", self.setup_ledger_tab),
            ("planning",  "🧭 التخطيط المالي والائتمان",     self.setup_planning_tab),
            ("journal",   "📓 القيود اليومية",            self.setup_journal_tab),
            ("coa",       "🗂️ دليل الحسابات",            self.setup_coa_tab),
            ("company",   "⚙️ بيانات الشركة والأمان",       self.setup_company_tab),
            ("users",     "👤 المستخدمون والصلاحيات",      self.setup_users_tab),
            ("audit",     "🛡️ سجل التدقيق والرقابة",        self.setup_audit_tab),
        ]

        self.frames = {}
        self.nav_buttons = {}

        # كل الصفحات تُبنى (لأن بعض العمليات تحدّث صفحات أخرى)، لكن التنقل يظهر فقط للصفحات المسموح بها
        for key, text, setup in tabs_config:
            frame = ttk.Frame(self.notebook)
            self.frames[key] = frame
            self.notebook.add(frame)
            setup(frame)

        group_titles = {"dashboard": "الرئيسية", "customers": "البيانات والمخزون", "sales": "العمليات",
                        "statement": "الحسابات", "company": "الإدارة"}
        self._nav_fg = {"import": UI["cyan"], "discounts": "#5eead4"}
        pending = None
        for key, text, _ in tabs_config:
            pending = group_titles.get(key, pending)
            if not self.can(key):
                continue
            if pending:
                tk.Label(self.sidebar_frame, text=f"— {pending} —", bg=self.clr_sidebar_bg, fg="#7b8db0",
                         font=(UI_FONT, 8, "bold"), anchor="e").pack(fill="x", padx=18, pady=(5, 0))
                pending = None
            btn = tk.Button(
                self.sidebar_frame,
                text=text,
                font=(UI_FONT, 9, "bold"),
                bg=self.sidebar_frame['bg'],
                fg=self.clr_text_light,
                activebackground=UI["primary_dk"],
                activeforeground="#ffffff",
                bd=0,
                relief="flat",
                padx=16,
                pady=3,
                anchor="e",
                cursor="hand2",
                command=lambda k=key: self.show_page(k)
            )
            btn.configure(fg=self._nav_fg.get(key, self.clr_text_light))
            btn.pack(fill="x", pady=0, padx=8)
            btn.bind("<Enter>", lambda e, k=key: self._nav_hover(k, True))
            btn.bind("<Leave>", lambda e, k=key: self._nav_hover(k, False))
            self.nav_buttons[key] = btn

        self.tab_loaders = {
            str(self.frames["dashboard"]): self.load_dashboard_data,
            str(self.frames["customers"]): lambda: self.load_partner_page("customer"),
            str(self.frames["suppliers"]): lambda: self.load_partner_page("supplier"),
            str(self.frames["reps"]):      self.load_reps,
            str(self.frames["import"]):    self.refresh_import_page,
            str(self.frames["inventory"]): self.load_inventory,
            str(self.frames["expiry"]):    self.load_expiry_page,
            str(self.frames["sales"]):     lambda: self.refresh_flow("sales"),
            str(self.frames["purchase"]):  lambda: self.refresh_flow("purchase"),
            str(self.frames["cash"]):      lambda: (self.refresh_cash_accounts(), self.load_cash_transactions(),
                                                    self.update_party_options()),
            str(self.frames["returns"]):   lambda: (self.update_returns_dropdowns(), self.load_returns_invoices()),
            str(self.frames["assets"]):    self.load_assets,
            str(self.frames["statement"]): self.update_statement_partners_dropdown,
            str(self.frames["discounts"]): self.load_discounts_page,
            str(self.frames["income"]):    self.generate_income_statement,
            str(self.frames["reports"]):   self.on_reports_shown,
            str(self.frames["journal"]):   lambda: (self.load_journal_accounts(), self.load_journal_register()),
            str(self.frames["coa"]):       self.load_coa_data,
            str(self.frames["company"]):   self.load_company,
            str(self.frames["users"]):     self.load_users,
            str(self.frames["ledger"]):    self.load_ledger_page,
            str(self.frames["planning"]):  self.load_planning_page,
            str(self.frames["audit"]):     self.load_audit_page,
        }

        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)
        allowed = [k for k, _, _ in tabs_config if self.can(k)]
        self._schedule_dashboard_refresh()
        if allowed:
            self.show_page(allowed[0])
        else:
            self.root.after(200, self._no_permissions)

    # ---------------------------------------------------------- الصلاحيات والجلسة
    def can(self, key):
        """صلاحية صفحة (مثل «purchase») أو إجراء/تبويب داخلها (مثل «purchase.receipt»)."""
        return user_can(self.user, key)

    def can_fin(self, role):
        return user_can_fin(self.user, role)

    def can_report(self, key):
        if not self.can(REPORT_GROUP_OF.get(key, "reports")):
            return False
        need = REPORT_FIN_REQUIRES.get(key)
        return need is None or self.can_fin("customer" if need.startswith("sales") else "supplier")

    def can_modify(self):
        """تعديل وحذف المستندات والأرصدة: المدير دائماً، أو مستخدم مُنح صلاحية «modify»."""
        return self.user["is_admin"] or "modify" in self.user["perms"]

    def require_modify(self):
        if self.can_modify():
            return True
        messagebox.showwarning("صلاحيات", "هذه العملية تحتاج صلاحية «تعديل وحذف المستندات والأرصدة» (المدير).")
        return False

    def with_stock_override(self, fn, *args, **kwargs):
        """ينفذ العملية؛ لو أدت لرصيد مخزون سالب: المدير يؤكد ويتابع، وغيره يُمنع."""
        try:
            return fn(*args, **kwargs)
        except NegativeStockError as e:
            if not self.can_modify():
                raise UserError(f"{e}\n\nلا يمكن المتابعة (تحتاج صلاحية المدير).")
            if not messagebox.askyesno("تأكيد إداري", f"{e}\n\nبصلاحية المدير: هل تريد المتابعة رغم الرصيد السالب؟",
                                       icon="warning"):
                raise UserError("تم إلغاء العملية.")
            return fn(*args, allow_negative=True, **kwargs)

    def _no_permissions(self):
        messagebox.showwarning("صلاحيات", "لا توجد صلاحيات مسندة لهذا المستخدم.\nراجع مدير النظام.")
        self.logout(confirm=False)

    def shutdown(self):
        if self._dash_job:
            try:
                self.root.after_cancel(self._dash_job)
            except tk.TclError:
                pass
            self._dash_job = None

    def logout(self, confirm=True):
        if confirm and not messagebox.askyesno("تسجيل الخروج", "هل تريد تسجيل الخروج؟"):
            return
        self.shutdown()
        for w in self.root.winfo_children():
            w.destroy()
        self.root.withdraw()
        if self.on_logout:
            self.on_logout()
        else:
            self.root.destroy()

    def show_page(self, page_key):
        """الانتقال بين الصفحات وتظليل الزر النشط (للصفحات المسموح بها فقط)."""
        if not self.can(page_key) or page_key not in self.nav_buttons:
            messagebox.showwarning("صلاحيات", "ليس لديك صلاحية لفتح هذه الصفحة.")
            return
        frame = self.frames[page_key]
        changed = self.notebook.select() != str(frame)
        self.notebook.select(frame)
        if changed:
            # التحميل يتم الآن (متزامناً) ويُتجاهل الحدث المؤجل <<NotebookTabChanged>> لهذا الانتقال،
            # حتى لا يعيد محمّل الصفحة ضبط الحقول بعد تعبئتها برمجياً (مثل «سداد هذه الفاتورة»).
            self._skip_tab_event = str(frame)
            loader = getattr(self, "tab_loaders", {}).get(str(frame))
            if loader:
                try:
                    loader()
                except Exception as e:
                    handle_error(e)
        for k, btn in self.nav_buttons.items():
            btn.configure(bg=self.clr_sidebar_bg, fg=self._nav_fg.get(k, self.clr_text_light))
        self.nav_buttons[page_key].configure(bg=self.clr_primary, fg="#ffffff")
        self._active_page = page_key

    def _nav_hover(self, key, entering):
        btn = self.nav_buttons.get(key)
        if btn is None or key == self._active_page:
            return
        btn.configure(bg=self.clr_hover if entering else self.clr_sidebar_bg,
                      fg="#ffffff" if entering else self._nav_fg.get(key, self.clr_text_light))

    def build_header(self):
        header = tk.Frame(self.root, bg=UI["slate"], height=84)
        header.pack(fill="x", side="top")
        header.pack_propagate(False)
        strip = tk.Canvas(self.root, height=4, bg=UI["slate"], highlightthickness=0)
        strip.pack(fill="x", side="top")
        strip.bind("<Configure>", lambda e, c=strip: (c.delete("all"), [c.create_rectangle(
            x, 0, x + 6, 4, outline="", fill=_hex_mix(UI["primary"], UI["cyan"], x / max(e.width, 1)))
            for x in range(0, e.width, 6)]))

        if self.user["is_admin"]:
            money_button(header, "🗑️ تصفير الحركات", "#dc2626", self.reset_system_data, 9).pack(side="left", padx=6, pady=24)
        if self.can("closure"):
            money_button(header, "🔒 الإغلاق اليومي", UI["primary"], self.perform_daily_closure, 9).pack(side="left", padx=6, pady=24)
        money_button(header, "💾 نسخة احتياطية", UI["teal"], self.manual_backup, 9).pack(side="left", padx=6, pady=24)

        brand = tk.Frame(header, bg=UI["slate"])
        brand.pack(side="right", padx=20)
        tk.Label(brand, text=BRAND_AR, font=(UI_FONT, 16, "bold"), bg=UI["slate"], fg="#ffffff").pack(anchor="e", pady=(12, 0))
        tk.Label(brand, text=BRAND_EN, font=(UI_FONT, 9, "bold"), bg=UI["slate"], fg=UI["cyan"]).pack(anchor="e")

        # بطاقة المستخدم الحالي
        ubox = tk.Frame(header, bg="#0b1f3a", highlightbackground=UI["teal2"], highlightthickness=1)
        ubox.pack(side="right", padx=10, pady=14)
        badge = tk.Canvas(ubox, width=42, height=42, bg="#0b1f3a", highlightthickness=0)
        badge.pack(side="right", padx=(6, 4), pady=4)
        draw_elhagry_mark(badge, 21, 19, 34)
        role = "مدير" if self.user["is_admin"] else "مستخدم"
        tk.Label(ubox, text=f"{self.user['full_name']}\n{role} • {self.user['username']}", bg="#0b1f3a", fg="white",
                 font=("Arial", 9, "bold"), justify="right").pack(side="right", padx=4)
        tk.Button(ubox, text="🔑", bg="#0b1f3a", fg=BRAND_GOLD2, bd=0, cursor="hand2", font=("Arial", 11),
                  activebackground=BRAND_NAVY2, command=self.change_own_password).pack(side="right", padx=2)
        tk.Button(ubox, text="⎋ خروج", bg="#0b1f3a", fg=BRAND_GOLD2, bd=0, cursor="hand2", font=("Arial", 9, "bold"),
                  activebackground=BRAND_NAVY2, command=self.logout).pack(side="right", padx=(2, 8))

    def change_own_password(self):
        if self.user.get("user_id") is None:
            return
        win = tk.Toplevel(self.root)
        win.title("تغيير كلمة المرور")
        win.geometry("440x250")
        win.resizable(False, False)
        win.grab_set()
        f = ttk.Frame(win, padding=16)
        f.pack(fill="both", expand=True)
        entries = []
        for i, label in enumerate(("كلمة المرور الحالية:", "كلمة المرور الجديدة:", "تأكيد كلمة المرور الجديدة:")):
            ttk.Label(f, text=label).grid(row=i, column=1, sticky="e", pady=6, padx=4)
            e = ttk.Entry(f, width=28, show="*")
            e.grid(row=i, column=0, sticky="w", pady=6)
            entries.append(e)

        def save():
            try:
                cur, new, new2 = (e.get() for e in entries)
                with db() as conn:
                    stored = conn.execute("SELECT password_hash FROM users WHERE user_id=?", (self.user["user_id"],)).fetchone()
                    if not stored or not verify_password(cur, stored[0]):
                        raise UserError("كلمة المرور الحالية غير صحيحة.")
                    if self.user.get("username") == DEFAULT_ADMIN_USER:
                        raise UserError(f"بيانات دخول المدير الرئيسي «{DEFAULT_ADMIN_USER}» ثابتة ولا يمكن تغييرها.")
                    if new != new2:
                        raise UserError("كلمتا المرور الجديدتان غير متطابقتين.")
                    validate_new_password(new)
                    conn.execute("UPDATE users SET password_hash=? WHERE user_id=?", (hash_password(new), self.user["user_id"]))
                messagebox.showinfo("نجاح", "تم تغيير كلمة المرور بنجاح.", parent=win)
                win.destroy()
            except Exception as ex:
                handle_error(ex)

        money_button(f, "حفظ كلمة المرور 💾", "#16a34a", save, 10).grid(row=3, column=0, columnspan=2, pady=14)
        entries[0].focus_set()

    # ---------------------------------------------------------- إدارة المستخدمين والصلاحيات (للمدير فقط)
    def setup_users_tab(self, frame):
        self.u_editing_id = None
        form = ttk.LabelFrame(frame, text=" بيانات المستخدم ", padding=10)
        form.pack(fill="x", padx=10, pady=(8, 4))
        self.u_username_e = ttk.Entry(form, width=20)
        self.u_fullname_e = ttk.Entry(form, width=28)
        self.u_pass_e = ttk.Entry(form, width=20, show="*")
        self.u_pass2_e = ttk.Entry(form, width=20, show="*")
        self.u_admin_var = tk.IntVar(value=0)
        self.u_active_var = tk.IntVar(value=1)
        chk_admin = ttk.Checkbutton(form, text="مدير النظام (كل الصلاحيات)", variable=self.u_admin_var,
                                    command=self.users_toggle_admin)
        chk_active = ttk.Checkbutton(form, text="نشط", variable=self.u_active_var)
        rtl_row(form, 0, [("اسم المستخدم:", self.u_username_e), ("الاسم الكامل:", self.u_fullname_e),
                          (None, chk_admin), (None, chk_active)])
        self.u_role_cb = ttk.Combobox(form, width=26, state="readonly")
        self.u_role_cb.bind("<<ComboboxSelected>>", lambda e: self.users_apply_role())
        rtl_row(form, 1, [("كلمة المرور:", self.u_pass_e), ("تأكيد كلمة المرور:", self.u_pass2_e),
                          ("الدور الوظيفي:", self.u_role_cb)])
        ttk.Label(form, text="كلمة المرور 8 أحرف على الأقل وتحتوي حروفاً إنجليزية وأرقاماً. "
                             "عند تعديل مستخدم اترك الحقلين فارغين للإبقاء على كلمة مروره.",
                  foreground="#78716c").grid(row=2, column=0, columnspan=8, sticky="e", pady=(2, 0))

        pbox = ttk.LabelFrame(frame, text=" الصلاحيات: الصفحات المسموح للمستخدم بفتحها ", padding=10)
        pbox.pack(fill="x", padx=10, pady=4)
        self.u_perm_vars, self.u_perm_checks = {}, []
        cols = 5
        for i, (key, label) in enumerate(PERMISSIONS):
            v = tk.IntVar(value=0)
            r, c = divmod(i, cols)
            cb = ttk.Checkbutton(pbox, text=label, variable=v)
            cb.grid(row=r, column=cols - 1 - c, sticky="e", padx=12, pady=3)
            self.u_perm_vars[key] = v
            self.u_perm_checks.append(cb)
        for c in range(cols):
            pbox.grid_columnconfigure(c, weight=1)
        for key, v in list(self.u_perm_vars.items()):
            if key in SUB_PERMISSIONS:
                v.trace_add("write", lambda *_a, k=key: self.users_parent_toggled(k))

        sbox = ttk.LabelFrame(frame, text=" صلاحيات تفصيلية داخل المديولات (تتفعّل عند تحديد الصفحة الأم أعلاه) ", padding=6)
        sbox.pack(fill="x", padx=10, pady=4)
        snb = ttk.Notebook(sbox)
        snb.pack(fill="x")
        self.u_sub_checks = {}
        titles = {"sales": "🧾 المبيعات", "purchase": "🛒 المشتريات", "cash": "💰 الخزينة والبنوك", "reports": "📊 التقارير"}
        sub_tabs = {}
        for parent in reversed(list(SUB_PERMISSIONS)):
            sub_tabs[parent] = tk.Frame(snb, padx=6, pady=4)
            snb.add(sub_tabs[parent], text=f"  {titles[parent]}  ")
        snb.select(sub_tabs["sales"])
        for parent, subs in SUB_PERMISSIONS.items():
            tab = sub_tabs[parent]
            scols = 3
            checks = []
            for i, (key, label) in enumerate(subs):
                v = tk.IntVar(value=0)
                r, c = divmod(i, scols)
                cb = ttk.Checkbutton(tab, text=label, variable=v)
                cb.grid(row=r, column=scols - 1 - c, sticky="e", padx=12, pady=2)
                self.u_perm_vars[key] = v
                checks.append(cb)
            for c in range(scols):
                tab.grid_columnconfigure(c, weight=1)
            self.u_sub_checks[parent] = checks
        ttk.Label(sbox, text="مثال: لأمين المخزن حدّد «المشتريات» ثم «أذون الاستلام / التحويل للمخزن» فقط "
                             "← تختفي أوامر الشراء والفواتير وأرصدة الموردين.",
                  foreground="#78716c").pack(anchor="e", pady=(4, 0))

        bar = tk.Frame(frame)
        bar.pack(fill="x", padx=10, pady=4)
        self.u_save_btn = money_button(bar, "➕ إضافة المستخدم", "#16a34a", self.save_user, 10)
        self.u_save_btn.pack(side="right", padx=4)
        money_button(bar, "✖ جديد / مسح الحقول", "#64748b", self.clear_user_form, 9).pack(side="right", padx=4)
        money_button(bar, "☑ تحديد كل الصلاحيات", "#0284c7", lambda: self.users_set_all(1), 9).pack(side="right", padx=4)
        money_button(bar, "☐ إلغاء كل الصلاحيات", "#0284c7", lambda: self.users_set_all(0), 9).pack(side="right", padx=4)
        money_button(bar, "🗑️ حذف المستخدم المحدد", "#dc2626", self.delete_user, 9).pack(side="left", padx=4)
        money_button(bar, "🔓 فك الإيقاف المؤقت", "#d97706", self.unlock_user, 9).pack(side="left", padx=4)
        money_button(bar, "🧑‍💼 حفظ الصلاحيات المحددة كدور", UI["primary"], self.save_role_from_form, 9).pack(side="left", padx=4)
        money_button(bar, "🗑️ حذف الدور المحدد", "#b91c1c", self.delete_selected_role, 9).pack(side="left", padx=4)

        box = ttk.LabelFrame(frame, text=" المستخدمون (اضغط على مستخدم لتعديل بياناته وصلاحياته) ", padding=8)
        box.pack(fill="both", expand=True, padx=10, pady=5)
        self.users_tree = make_tree(box, [
            ("user", "اسم المستخدم", 120, "center"), ("name", "الاسم الكامل", 180, "e"), ("role", "الدور", 130, "center"),
            ("status", "الحالة", 90, "center"), ("last", "آخر دخول", 140, "center"), ("perms", "الصلاحيات", 400, "e")])
        self.users_tree.bind("<<TreeviewSelect>>", self.on_user_selected)
        self.users_load_roles(None)
        self.load_users()

    def users_toggle_admin(self):
        admin = bool(self.u_admin_var.get())
        for cb in self.u_perm_checks:
            cb.state(["disabled"] if admin else ["!disabled"])
        if admin:
            self.users_set_all(1)
        for parent in SUB_PERMISSIONS:
            self.users_parent_toggled(parent, auto_fill=False)

    def users_parent_toggled(self, parent, auto_fill=True):
        """الصلاحيات الفرعية تُفعّل فقط مع صفحتها الأم؛ تحديد الصفحة الأم لأول مرة يحدد كل إجراءاتها."""
        if not getattr(self, "u_sub_checks", None):
            return
        on = bool(self.u_perm_vars[parent].get()) and not self.u_admin_var.get()
        subs = [k for k, _ in SUB_PERMISSIONS[parent]]
        for cb in self.u_sub_checks[parent]:
            cb.state(["!disabled"] if on else ["disabled"])
        if self.u_admin_var.get():
            for k in subs:
                self.u_perm_vars[k].set(1)
        elif auto_fill and on and not any(self.u_perm_vars[k].get() for k in subs):
            for k in subs:
                self.u_perm_vars[k].set(1)

    def users_set_all(self, value):
        for v in self.u_perm_vars.values():
            v.set(value)

    def clear_user_form(self):
        self.u_editing_id = None
        for e in (self.u_username_e, self.u_fullname_e, self.u_pass_e, self.u_pass2_e):
            e.delete(0, tk.END)
        self.u_admin_var.set(0)
        self.u_active_var.set(1)
        self.users_toggle_admin()
        self.users_set_all(0)
        self.u_save_btn.config(text="➕ إضافة المستخدم")
        self.users_load_roles(None)
        if self.users_tree.selection():
            self.users_tree.selection_remove(*self.users_tree.selection())

    def load_users(self):
        clear_tree(self.users_tree)
        labels = dict(PERMISSIONS)
        now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        with db() as conn:
            for uid, un, fn, adm, perms, act, last, locked, role in conn.execute(
                    "SELECT u.user_id, u.username, u.full_name, u.is_admin, u.permissions, u.active, u.last_login, "
                    "u.locked_until, r.name FROM users u LEFT JOIN user_roles r ON r.role_id=u.role_id ORDER BY u.user_id"):
                pset = set((perms or "").split(","))
                parts = []
                for p in (k for k in PERMISSION_KEYS if k in pset):
                    subs = [s for s, _ in SUB_PERMISSIONS.get(p, [])]
                    chosen = [s for s in subs if s in pset]
                    parts.append(f"{labels[p]} ({len(chosen)}/{len(subs)})" if chosen and len(chosen) < len(subs) else labels[p])
                summary = "كل الصلاحيات" if adm else ("، ".join(parts) or "بدون صلاحيات")
                status = "موقوف" if not act else ("مقفل مؤقتاً" if locked and locked > now else "نشط")
                self.users_tree.insert("", "end", iid=str(uid), values=(
                    un, fn or "-", "مدير" if adm else (role or "صلاحيات فردية"), status, last or "-", summary))

    def on_user_selected(self, event=None):
        sel = self.users_tree.selection()
        if not sel:
            return
        with db() as conn:
            r = conn.execute("SELECT user_id, username, full_name, is_admin, permissions, active FROM users WHERE user_id=?",
                             (int(sel[0]),)).fetchone()
        if not r:
            return
        self.u_editing_id = r[0]
        for e, val in ((self.u_username_e, r[1]), (self.u_fullname_e, r[2] or "")):
            e.delete(0, tk.END)
            e.insert(0, val)
        for e in (self.u_pass_e, self.u_pass2_e):
            e.delete(0, tk.END)
        self.u_admin_var.set(r[3])
        self.u_active_var.set(r[5])
        self.users_load_roles(user_role_id(r[0]))
        perms = set((r[4] or "").split(","))
        rec = dict(is_admin=bool(r[3]), perms=perms)
        for k, v in self.u_perm_vars.items():
            v.set(1 if (r[3] or (user_can(rec, k) if k in SUB_PARENT else k in perms)) else 0)
        self.users_toggle_admin()
        self.u_save_btn.config(text=f"💾 حفظ تعديلات {r[1]}")

    def save_user(self):
        try:
            uname = self.u_username_e.get().strip()
            full = self.u_fullname_e.get().strip()
            pw, pw2 = self.u_pass_e.get(), self.u_pass2_e.get()
            is_admin, active = int(self.u_admin_var.get()), int(self.u_active_var.get())
            perms = [k for k in ALL_PERMISSION_KEYS if k in self.u_perm_vars and self.u_perm_vars[k].get()]
            uid = self.u_editing_id
            msg = save_user_account(uid, uname, full, pw, pw2, is_admin, active, perms, self.user["user_id"],
                                    role_id=self.users_selected_role_id())
            if uid is not None and uid == self.user["user_id"]:
                self.user.update(username=uname, full_name=full or uname)
            messagebox.showinfo("نجاح", msg + "\nتسري الصلاحيات من الدخول التالي للمستخدم.")
            self.clear_user_form()
            self.load_users()
        except Exception as e:
            handle_error(e)

    def _selected_user_id(self):
        sel = self.users_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر مستخدماً من الجدول أولاً.")
            return None
        return int(sel[0])

    def delete_user(self):
        uid = self._selected_user_id()
        if uid is None:
            return
        try:
            if uid == self.user["user_id"]:
                raise UserError("لا يمكنك حذف حسابك الحالي.")
            with db() as conn:
                r = conn.execute("SELECT username, is_admin, active FROM users WHERE user_id=?", (uid,)).fetchone()
                if r is None:
                    return
                if r[0] == DEFAULT_ADMIN_USER:
                    raise UserError(f"لا يمكن حذف حساب المدير الرئيسي «{DEFAULT_ADMIN_USER}».")
                if r[1] and r[2] and active_admin_count(conn, uid) == 0:
                    raise UserError("لا يمكن حذف آخر مدير نشط.")
            if not messagebox.askyesno("حذف مستخدم", f"حذف المستخدم {r[0]} نهائياً؟\n(البديل: إلغاء علامة «نشط» لإيقافه)"):
                return
            with db() as conn:
                conn.execute("DELETE FROM users WHERE user_id=?", (uid,))
            self.clear_user_form()
            self.load_users()
        except Exception as e:
            handle_error(e)

    def unlock_user(self):
        uid = self._selected_user_id()
        if uid is None:
            return
        with db() as conn:
            conn.execute("UPDATE users SET failed_attempts=0, locked_until=NULL WHERE user_id=?", (uid,))
        self.load_users()
        messagebox.showinfo("تم", "تم فك الإيقاف المؤقت للمستخدم.")

    def manual_backup(self):
        try:
            path = backup_database("manual")
            messagebox.showinfo("نسخة احتياطية", f"تم حفظ النسخة الاحتياطية:\n{os.path.abspath(path)}")
        except Exception as e:
            handle_error(e)

    def reset_system_data(self):
        if not self.user["is_admin"]:
            messagebox.showwarning("صلاحيات", "تصفير الحركات متاح لمدير النظام فقط.")
            return
        if not messagebox.askyesno("تأكيد التصفير",
                                   "سيتم حذف كل الأوامر والفواتير والمرتجعات والخصومات اللاحقة وسندات الخزينة والقيود اليومية والإغلاقات،\n"
                                   "وإرجاع المخزون للرصيد الافتتاحي.\n(دليل الحسابات والشركاء والأصناف والأصول لا تُحذف)\n\n"
                                   "هل تريد المتابعة؟", icon="warning"):
            return
        word = simpledialog.askstring("تأكيد نهائي", 'للتأكيد اكتب كلمة: تصفير')
        if (word or "").strip() != "تصفير":
            messagebox.showinfo("إلغاء", "لم يتم التصفير.")
            return
        try:
            backup = backup_database("before_reset")
            with db() as conn:
                _audit.audit_event(conn, _audit.ACT_EVENT, "تصفير الحركات", "كل المستندات",
                                   new={"backup": os.path.basename(backup)})
                for table in ("sales_delivery_lines", "purchase_receipt_lines", "sales_invoice_lines", "purchase_invoice_lines",
                              "sales_deliveries", "purchase_receipts", "sales_invoices", "purchase_invoices",
                              "sales_order_lines", "purchase_order_lines", "sales_orders", "purchase_orders",
                              "returns_invoices", "cash_transactions", "daily_closures", "post_discount_lines",
                              "post_discounts", "journal_lines", "journal_entries"):
                    conn.execute(f"DELETE FROM {table}")
                # مرفقات المستندات المحذوفة: بدون حذفها ستظهر على المستندات الجديدة التي تأخذ نفس الأرقام
                conn.execute("DELETE FROM attachments WHERE doc_type IN ('purchase', 'journal')")
                conn.execute("UPDATE inventory SET quantity = opening_qty, cost_price = opening_cost")
                conn.execute("UPDATE sequences SET next_value=1 WHERE name IN ('sales_invoice','purchase_invoice',"
                             "'purchase_discount','sales_discount')")
                conn.execute("DELETE FROM app_settings WHERE key=?", (CASH_OPENING_SETTING,))
                conn.execute("DELETE FROM sqlite_sequence WHERE name IN ('sales_invoices','sales_invoice_lines',"
                             "'purchase_invoices','purchase_invoice_lines','returns_invoices','cash_transactions','daily_closures',"
                             "'sales_orders','sales_order_lines','sales_deliveries','sales_delivery_lines','purchase_orders',"
                             "'purchase_order_lines','purchase_receipts','purchase_receipt_lines','journal_entries','journal_lines',"
                             "'post_discounts','post_discount_lines')")
            self.refresh_all()
            messagebox.showinfo("تم التصفير", f"أصبح النظام جاهزاً للعمل.\nنسخة ما قبل التصفير:\n{os.path.abspath(backup)}")
        except Exception as e:
            handle_error(e)

    def refresh_all(self):
        self.refresh_flow("sales")
        self.refresh_flow("purchase")
        self.load_returns_invoices()
        self.load_inventory()
        self.load_cash_transactions()
        self.load_dashboard_data()

    def perform_daily_closure(self):
        if not self.can("closure"):
            messagebox.showwarning("صلاحيات", "ليس لديك صلاحية الإغلاق اليومي.")
            return
        d = today()
        try:
            with db() as conn:
                s_sum = conn.execute("SELECT SUM(total_amount) FROM sales_invoices WHERE invoice_date=?", (d,)).fetchone()[0] or 0.0
                p_sum = conn.execute("SELECT SUM(total_amount) FROM purchase_invoices WHERE purchase_date=?", (d,)).fetchone()[0] or 0.0
                moves = cash_movements(conn, d, d).values()
                c_in = sum(m[1] for m in moves)
                c_out = sum(m[2] for m in moves)
                conn.execute("INSERT OR REPLACE INTO daily_closures (closure_date, total_sales, total_purchases, net_cash_movement) "
                             "VALUES (?,?,?,?)", (d, s_sum, p_sum, c_in - c_out))
            messagebox.showinfo("الإغلاق اليومي",
                                f"تم الإغلاق اليومي لتاريخ {d}\n\nالمبيعات: {fmt(s_sum)}\nالمشتريات: {fmt(p_sum)}\n"
                                f"صافي حركة النقدية: {fmt(c_in - c_out)}")
        except Exception as e:
            handle_error(e)

    def on_tab_changed(self, event):
        skip, self._skip_tab_event = getattr(self, "_skip_tab_event", None), None
        if skip == self.notebook.select():
            return   # حُمِّلت بالفعل داخل show_page
        loader = getattr(self, "tab_loaders", {}).get(self.notebook.select())
        if loader:
            try:
                loader()
            except Exception as e:
                handle_error(e)

    # ---------------------------------------------------------- دليل الحسابات
    def setup_coa_tab(self, frame):
        top = tk.Frame(frame)
        top.pack(fill="x", padx=10, pady=(8, 0))
        money_button(top, "🔄 إعادة استيراد من الإكسيل", "#0284c7", self.reimport_chart, 9).pack(side="right")
        money_button(top, "➕ إضافة حساب جديد", "#16a34a", self.open_add_account_dialog, 9).pack(side="right", padx=6)
        money_button(top, "🌿 حساب فرعي تحت المحدد", UI["teal"], self.add_sub_account_selected, 9).pack(side="right", padx=6)
        money_button(top, "📒 كشف حساب تحليلي للمحدد", UI["primary"], self.open_account_statement, 9).pack(side="right", padx=6)

        box = ttk.LabelFrame(frame, text=" شجرة الحسابات (من ملف الإكسيل + الحسابات المضافة يدوياً) ", padding=10)
        box.pack(fill="both", expand=True, padx=10, pady=8)
        self.coa_tree = make_tree(box, [("code", "رقم الحساب", 120, "center"), ("name", "اسم الحساب", 400, "e"),
                                        ("type", "نوع الحساب", 120, "center"), ("level", "المستوى", 90, "center"),
                                        ("parent", "الحساب الرئيسي", 120, "center")])
        self.coa_tree.bind("<Double-1>", lambda e: self.open_account_statement())
        self.load_coa_data()

    def reimport_chart(self):
        if not os.path.exists(EXCEL_PATH):
            messagebox.showwarning("تنبيه", f"الملف غير موجود بجوار البرنامج:\n{EXCEL_PATH}")
            return
        try:
            n = import_chart_of_accounts()
        except Exception as e:
            handle_error(e)
            return
        self.load_coa_data()
        self.load_journal_accounts()
        if n is None:
            messagebox.showwarning("تنبيه", "تعذرت قراءة ملف الإكسيل (تأكد من وجود العمودين Account و Name). لم يتغير الدليل.")
        else:
            messagebox.showinfo("تم", f"تم تحديث {n} حساباً من ملف الإكسيل.\nالحسابات المضافة يدوياً لم تُحذف.")

    def load_coa_data(self):
        clear_tree(self.coa_tree)
        with db() as conn:
            parents = dict(conn.execute("SELECT account_code, COALESCE(parent_code,'') FROM chart_of_accounts"))
        for code, name, typ, lvl, grp in account_tree_rows():
            self.coa_tree.insert("", "end", iid=code, values=(code, ("      " * lvl) + ("▸ " if grp else "") + name, typ,
                                                              f"{'تجميعي' if grp else 'فرعي'} ({lvl + 1})", parents.get(code) or "-"))

    # ---------------------------------------------------------- سلة الفاتورة (مشتركة)
    CART_COLUMNS = CART_COLUMNS

    def refresh_cart(self, tree, cart, label):
        clear_tree(tree)
        cols = tree["columns"]
        for i, l in enumerate(cart):
            tree.insert("", "end", iid=str(i), values=cart_row_values(cols, l))
        label.config(text=cart_summary(cart))

    def remove_cart_line(self, tree, cart, label):
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر السطر المراد حذفه من الفاتورة.")
            return
        del cart[int(sel[0])]
        self.refresh_cart(tree, cart, label)

    def item_pairs(self):
        with db() as conn:
            return [(f"{c} - {n}", (c, n, u)) for c, n, u in
                    conn.execute("SELECT item_code, item_name, unit FROM inventory ORDER BY item_name")]

    def partner_pairs(self, flag):
        col = "is_customer" if flag == "customer" else "is_supplier"
        with db() as conn:
            return [(f"{c} - {n}", (c, n)) for c, n in
                    conn.execute(f"SELECT partner_code, name FROM business_partners WHERE {col}=1 ORDER BY name")]

    def print_sales_invoice_pdf(self, event=None):
        sel = self.sal_tree.selection()
        if sel:
            self.export_sales_invoice(int(sel[0]))

    def export_sales_invoice(self, inv_id):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                            title="حفظ الفاتورة", initialfile=f"فاتورة_{inv_id}.pdf")
        if not path:
            return
        try:
            build_sales_invoice_pdf(path, inv_id)
            messagebox.showinfo("نجاح", "تم إنشاء الفاتورة بنجاح!")
        except Exception as e:
            handle_error(e)

    def _read_batch_dates(self, prod_e, exp_e, label=""):
        raw_p, raw_e = prod_e.get().strip(), exp_e.get().strip()
        prod = parse_date(raw_p) if raw_p else None
        exp = parse_date(raw_e) if raw_e else None
        if (raw_p and not prod) or (raw_e and not exp):
            raise UserError(f"صيغة تاريخ الإنتاج/الانتهاء غير صحيحة{f' للصنف «{label}»' if label else ''}. استخدم YYYY-MM-DD")
        validate_batch_dates(prod, exp, label)
        return prod, exp

    def show_purchase_lines(self, event=None):
        sel = self.pur_tree.selection()
        if not sel:
            return
        win = tk.Toplevel(self.root)
        win.title(f"بنود فاتورة المشتريات {invoice_number('purchase', int(sel[0]))}")
        win.geometry("1000x320")
        box = tk.Frame(win)
        box.pack(fill="both", expand=True, padx=10, pady=10)
        tree = make_tree(box, PURCHASE_LINE_COLUMNS, rtl=True)
        with db() as conn:
            for r in conn.execute("SELECT item_code, item_name, quantity, unit, unit_cost, subtotal, vat_amount, total_amount, "
                                  "production_date, expiry_date FROM purchase_invoice_lines WHERE purchase_id=? ORDER BY line_id",
                                  (int(sel[0]),)):
                tree.insert("", "end", values=(r[0], r[1], f"{r[2]:g}", r[3] or "-", fmt(r[4]), fmt(r[5]), fmt(r[6]),
                                               fmt(r[7]), r[8] or "-", r[9] or "-"))

    # ---------------------------------------------------------- المرتجعات
    def setup_returns_tab(self, frame):
        self.ret_editing_id = None
        f = ttk.LabelFrame(frame, text=" تسجيل مرتجع (مبيعات أو مشتريات) ", padding=10)
        f.pack(fill="x", padx=10, pady=8)

        self.ret_type_cb = ttk.Combobox(f, values=[SALES_RETURN, PURCHASE_RETURN], width=18, state="readonly")
        self.ret_type_cb.set(SALES_RETURN)
        self.ret_type_cb.bind("<<ComboboxSelected>>", lambda e: self.update_returns_dropdowns())
        self.ret_date_e = date_entry(f)
        self.ret_party_cb = SearchableCombobox(f, width=34)
        rtl_row(f, 0, [("نوع المرتجع:", self.ret_type_cb), ("التاريخ:", self.ret_date_e),
                       ("الطرف (عميل/مورد):", self.ret_party_cb)])

        self.ret_item_cb = SearchableCombobox(f, width=34)
        self.ret_qty_e = ttk.Entry(f, width=12)
        self.ret_price_e = ttk.Entry(f, width=12)
        rtl_row(f, 1, [("الصنف:", self.ret_item_cb), ("الكمية المرتجعة:", self.ret_qty_e),
                       ("السعر / التكلفة:", self.ret_price_e)])

        self.ret_notes_e = ttk.Entry(f, width=60)
        rtl_row(f, 2, [("ملاحظات:", self.ret_notes_e)])
        bar = tk.Frame(f)
        bar.grid(row=3, column=0, columnspan=8, pady=8)
        self.ret_save_btn = money_button(bar, "حفظ وترحيل المرتجع للمخزن 🔄", "#dc2626", self.save_return_invoice, 11)
        self.ret_save_btn.pack(side="right", padx=4)
        money_button(bar, "✖ مسح / إلغاء التعديل", "#64748b", self.cancel_return_edit, 9).pack(side="right", padx=4)

        box = ttk.LabelFrame(frame, text=" سجل المرتجعات ", padding=10)
        box.pack(fill="both", expand=True, padx=10, pady=5)
        if self.can_modify():
            adm = tk.Frame(box)
            adm.pack(fill="x", pady=(0, 4))
            money_button(adm, "✏️ تعديل المرتجع المحدد", "#d97706", self.edit_selected_return, 9).pack(side="right", padx=3)
            money_button(adm, "🗑️ حذف المرتجع المحدد", "#dc2626", self.delete_selected_return, 9).pack(side="right", padx=3)
        holder = tk.Frame(box)
        holder.pack(fill="both", expand=True)
        self.ret_tree = make_tree(holder, [
            ("id", "الرقم", 60, "center"), ("date", "التاريخ", 90, "center"), ("type", "النوع", 110, "center"),
            ("party", "الطرف", 200, "e"), ("item", "الصنف", 200, "e"), ("qty", "الكمية", 70, "center"),
            ("price", "السعر", 90, "center"), ("vat", "الضريبة", 90, "center"), ("total", "الإجمالي", 100, "center"),
            ("notes", "ملاحظات", 150, "center")])
        self.ret_tree.bind("<Double-1>", lambda e: self.edit_selected_return() if self.can_modify() else None)
        self.load_returns_invoices()

    def update_returns_dropdowns(self):
        flag = "customer" if self.ret_type_cb.get() == SALES_RETURN else "supplier"
        self.ret_party_cb.set_items(self.partner_pairs(flag))
        self.ret_item_cb.set_items(self.item_pairs())

    def cancel_return_edit(self):
        self.ret_editing_id = None
        for e in (self.ret_qty_e, self.ret_price_e, self.ret_notes_e):
            e.delete(0, tk.END)
        self.ret_save_btn.config(text="حفظ وترحيل المرتجع للمخزن 🔄")

    def save_return_invoice(self):
        try:
            r_type = self.ret_type_cb.get()
            r_date = parse_date(self.ret_date_e.get())
            if not r_date:
                raise UserError("صيغة التاريخ غير صحيحة، استخدم YYYY-MM-DD")
            party = self.ret_party_cb.selected()
            item = self.ret_item_cb.selected()
            if not party or not item:
                raise UserError("اختر الطرف والصنف من القوائم.")
            qty, price = parse_float(self.ret_qty_e.get()), parse_float(self.ret_price_e.get())
            self.check_period_open(r_date)
            editing = self.ret_editing_id
            _, total = self.with_stock_override(save_return, r_type, r_date, party, (item[0], item[1]), qty, price,
                                                self.ret_notes_e.get().strip(), return_id=editing)
            messagebox.showinfo("نجاح", f"تم {'تعديل' if editing else 'تسجيل'} {r_type} بإجمالي {fmt(total)} وتحديث المخزن.")
            self.cancel_return_edit()
            self.load_returns_invoices()
            self.load_inventory()
        except Exception as e:
            handle_error(e)

    def edit_selected_return(self):
        if not self.require_modify():
            return
        sel = self.ret_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر مرتجعاً من السجل.")
            return
        with db() as conn:
            r = conn.execute("SELECT return_id, return_date, return_type, party_code, party_name, item_code, item_name, "
                             "quantity, unit_price, COALESCE(notes,'') FROM returns_invoices WHERE return_id=?",
                             (int(sel[0]),)).fetchone()
        if not r:
            return
        self.ret_type_cb.set(r[2])
        self.update_returns_dropdowns()
        self.ret_date_e.delete(0, tk.END)
        self.ret_date_e.insert(0, r[1])
        self.ret_party_cb.set(f"{r[3]} - {r[4]}")
        with db() as conn:
            iname = conn.execute("SELECT item_name FROM inventory WHERE item_code=?", (r[5],)).fetchone()
        self.ret_item_cb.set(f"{r[5]} - {iname[0] if iname else r[6]}")
        for e, v in ((self.ret_qty_e, f"{r[7]:g}"), (self.ret_price_e, f"{r[8]:g}"), (self.ret_notes_e, r[9])):
            e.delete(0, tk.END)
            e.insert(0, v)
        self.ret_editing_id = r[0]
        self.ret_save_btn.config(text=f"💾 حفظ تعديل المرتجع رقم {r[0]}")

    def delete_selected_return(self):
        if not self.require_modify():
            return
        sel = self.ret_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر مرتجعاً من السجل.")
            return
        if not messagebox.askyesno("حذف مرتجع", f"حذف المرتجع رقم {sel[0]}؟ سيُعاد احتساب رصيد الصنف."):
            return
        try:
            self.with_stock_override(delete_return, int(sel[0]))
            self.cancel_return_edit()
            self.load_returns_invoices()
            self.load_inventory()
        except Exception as e:
            handle_error(e)

    def load_returns_invoices(self):
        clear_tree(self.ret_tree)
        with db() as conn:
            for r in conn.execute("SELECT return_id, return_date, return_type, party_name, item_name, quantity, unit_price, "
                                  "vat_amount, total_amount, notes FROM returns_invoices ORDER BY return_id DESC"):
                self.ret_tree.insert("", "end", iid=str(r[0]), values=(
                    r[0], r[1], r[2], r[3], r[4], f"{r[5]:g}", fmt(r[6]), fmt(r[7]), fmt(r[8]), r[9] or "-"))

    # ---------------------------------------------------------- المخزون
    def setup_inventory_tab(self, frame):
        top = tk.Frame(frame, padx=10, pady=5)
        top.pack(fill="x")

        item_form = ttk.LabelFrame(top, text=" إضافة صنف جديد مع رصيد افتتاحي ", padding=10)
        item_form.pack(side="right", fill="x", expand=True, padx=5)
        self.inv_code_e = ttk.Label(item_form, text="", width=12, foreground=UI["primary_dk"], font=(UI_FONT, 10, "bold"))
        self.inv_name_e = ttk.Entry(item_form, width=26)
        self.inv_unit_e = ttk.Entry(item_form, width=8)
        self.inv_unit_e.insert(0, "KG")
        self.inv_open_e = ttk.Entry(item_form, width=10)
        self.inv_open_e.insert(0, "0")
        self.inv_open_cost_e = ttk.Entry(item_form, width=10)
        self.inv_open_cost_e.insert(0, "0")
        rtl_row(item_form, 0, [("الكود (تلقائي):", self.inv_code_e), ("الاسم:", self.inv_name_e), ("الوحدة:", self.inv_unit_e)])
        rtl_row(item_form, 1, [("رصيد افتتاحي:", self.inv_open_e), ("تكلفة الوحدة الافتتاحية:", self.inv_open_cost_e),
                               (None, money_button(item_form, "إضافة / تحديث صنف ➕", "#16a34a", self.save_new_item, 9))],
                slots=3)
        self.inv_code_hint = ttk.Label(item_form, text="", foreground="#0369a1")
        self.inv_code_hint.grid(row=2, column=0, columnspan=6, sticky="e", padx=6)

        search = ttk.LabelFrame(top, text=" بحث في المخزن ", padding=10)
        search.pack(side="left", fill="x", padx=5)
        ttk.Label(search, text="بحث بالاسم أو الكود:").pack(side="right", padx=5)
        self.inv_search_e = ttk.Entry(search, width=22)
        self.inv_search_e.pack(side="right", padx=5)
        self.inv_search_e.bind("<KeyRelease>", lambda e: self.load_inventory())
        tb = tk.Frame(frame)
        tb.pack(fill="x", padx=15, pady=(0, 2))
        if self.can("import"):
            money_button(tb, "📥 استيراد الأصناف (مركز استيراد البيانات)", UI["cyan"], lambda: self.show_page("import"),
                         9).pack(side="right", padx=3)
        money_button(tb, "✏️ رصيد افتتاحي للصنف المحدد", "#ea580c", self.edit_opening_balance, 9).pack(side="right", padx=3)
        if self.can_modify():
            money_button(tb, "🗑️ حذف الصنف المحدد", "#dc2626", self.delete_item_selected, 9).pack(side="left", padx=3)
            money_button(tb, "🛠️ تعديل الصنف المحدد", "#d97706", self.edit_item_dialog, 9).pack(side="left", padx=3)

        box = ttk.LabelFrame(frame, text=" أرصدة المخزن الحالية (اضغط مرتين لعرض تفاصيل الصنف) - الأحمر منتهي الصلاحية والبرتقالي قريب الانتهاء ",
                             padding=10)
        box.pack(fill="both", expand=True, padx=10, pady=5)
        self.inv_tree = make_tree(box, [
            ("code", "كود الصنف", 80, "center"), ("name", "اسم الصنف", 185, "e"), ("unit", "الوحدة", 50, "center"),
            ("opening", "رصيد افتتاحي", 80, "center"), ("added", "مشتريات", 75, "center"),
            ("ret_s", "مرتجع مبيعات", 80, "center"), ("issued", "مبيعات", 75, "center"),
            ("ret_p", "مرتجع مشتريات", 85, "center"), ("qty", "الرصيد الحالي", 85, "center"),
            ("cost", "متوسط التكلفة", 90, "center"), ("val", "القيمة الإجمالية", 100, "center"),
            ("expiry", "أقرب صلاحية", 90, "center")])
        self.inv_tree.tag_configure("expired", background="#fee2e2")
        self.inv_tree.tag_configure("near", background="#ffedd5")
        self.inv_tree.bind("<Double-1>", self.show_item_modal)
        self.load_inventory()

    def refresh_item_code_hint(self):
        with db() as conn:
            code = peek_serial(conn, "item")[0]
        self.inv_code_e.config(text=code)
        self.inv_code_hint.config(text="الكود يُولَّد تلقائياً بالصيغة الموحدة ELH-XXXX بعد آخر رقم مسجل.")

    def save_new_item(self):
        try:
            name = clean_name(self.inv_name_e.get())
            unit = self.inv_unit_e.get().strip() or "KG"
            open_qty = parse_float(self.inv_open_e.get() or "0")
            open_cost = parse_float(self.inv_open_cost_e.get() or "0")
            if not name:
                raise UserError("أدخل اسم الصنف!")
            if open_qty is None or open_qty < 0 or open_cost is None or open_cost < 0:
                raise UserError("الرصيد والتكلفة الافتتاحية يجب أن تكون أرقاماً صحيحة (صفر أو أكثر).")
            with db() as conn:
                dup = next((c for c, n in conn.execute("SELECT item_code, item_name FROM inventory")
                            if norm_name(n) == norm_name(name)), None)
                if dup:
                    raise UserError(f"يوجد صنف بنفس الاسم بالفعل (الكود {dup}). استخدم «تعديل الصنف» لتحديثه.")
                if True:
                    code = next_serial(conn, "item")
                    conn.execute("INSERT INTO inventory (item_code, item_name, unit, opening_qty, opening_cost, quantity, "
                                 "cost_price) VALUES (?,?,?,?,?,?,?)", (code, name, unit, open_qty, open_cost, open_qty, open_cost))
            messagebox.showinfo("نجاح", f"تم حفظ الصنف بنجاح (الكود {code})")
            self.inv_name_e.delete(0, tk.END)
            self.refresh_item_code_hint()
            self.load_inventory()
        except Exception as e:
            handle_error(e)

    def load_inventory(self):
        clear_tree(self.inv_tree)
        q = self.inv_search_e.get().strip().lower()
        with db() as conn:
            items = conn.execute("SELECT item_code, item_name, unit, opening_qty, quantity, cost_price FROM inventory "
                                 "ORDER BY item_name").fetchall()
            purch = dict(conn.execute("SELECT item_code, SUM(qty) FROM v_stock_in GROUP BY item_code"))
            sold = dict(conn.execute("SELECT item_code, SUM(qty) FROM v_stock_out GROUP BY item_code"))
            ret_s = dict(conn.execute("SELECT item_code, SUM(quantity) FROM returns_invoices WHERE return_type=? "
                                      "GROUP BY item_code", (SALES_RETURN,)))
            ret_p = dict(conn.execute("SELECT item_code, SUM(quantity) FROM returns_invoices WHERE return_type=? "
                                      "GROUP BY item_code", (PURCHASE_RETURN,)))
            nearest = {}
            for b in stock_batches(conn):
                if b["expiry"] and (b["code"] not in nearest or b["expiry"] < nearest[b["code"]]):
                    nearest[b["code"]] = b["expiry"]
        t = today()
        alert = (date.today() + timedelta(days=30)).strftime("%Y-%m-%d")
        for code, name, unit, opening, qty, cost in items:
            if q and q not in name.lower() and q not in code.lower():
                continue
            exp = nearest.get(code, "")
            tags = ("expired",) if exp and exp < t else (("near",) if exp and exp <= alert else ())
            self.inv_tree.insert("", "end", iid=code, tags=tags, values=(
                code, name, unit or "", f"{opening or 0:g}", f"{purch.get(code, 0) or 0:g}", f"{ret_s.get(code, 0) or 0:g}",
                f"{sold.get(code, 0) or 0:g}", f"{ret_p.get(code, 0) or 0:g}", f"{qty or 0:g}", fmt(cost),
                fmt((qty or 0) * (cost or 0)), exp or "-"))
        self.refresh_item_code_hint()

    def edit_item_dialog(self):
        if not self.require_modify():
            return
        sel = self.inv_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر صنفاً من الجدول أولاً.")
            return
        with db() as conn:
            r = conn.execute("SELECT item_code, item_name, unit, opening_qty, opening_cost FROM inventory WHERE item_code=?",
                             (sel[0],)).fetchone()
            refs = item_reference_count(conn, sel[0])
        if not r:
            return
        win = tk.Toplevel(self.root)
        win.title(f"تعديل الصنف {r[0]}")
        win.geometry("520x330")
        win.grab_set()
        f = ttk.Frame(win, padding=16)
        f.pack(fill="both", expand=True)
        fields = {}
        for i, (key, label, val) in enumerate((("code", "الكود:", r[0]), ("name", "الاسم:", r[1]), ("unit", "الوحدة:", r[2] or "KG"),
                                               ("oq", "الرصيد الافتتاحي (كمية):", f"{r[3] or 0:g}"),
                                               ("oc", "تكلفة الوحدة الافتتاحية:", f"{r[4] or 0:g}"))):
            ttk.Label(f, text=label).grid(row=i, column=1, sticky="e", pady=5, padx=4)
            e = ttk.Entry(f, width=34)
            e.insert(0, val)
            e.grid(row=i, column=0, sticky="w", pady=5)
            fields[key] = e
        ttk.Label(f, text=f"الصنف مستخدم في {refs} بند. تغيير الكود يُنقل لكل المستندات، وتغيير الرصيد الافتتاحي\n"
                          f"يعيد احتساب الرصيد ومتوسط التكلفة وتكلفة المبيعات من أول حركة.", foreground="#78716c",
                  justify="right").grid(row=5, column=0, columnspan=2, sticky="e", pady=6)

        def save():
            try:
                self.with_stock_override(update_item, r[0], fields["code"].get(), fields["name"].get(), fields["unit"].get(),
                                         parse_float(fields["oq"].get() or "0"), parse_float(fields["oc"].get() or "0"))
                win.destroy()
                self.load_inventory()
                messagebox.showinfo("نجاح", "تم تعديل الصنف وإعادة احتساب أرصدته.")
            except Exception as ex:
                handle_error(ex)

        money_button(f, "حفظ التعديلات 💾", "#16a34a", save, 10).grid(row=6, column=0, columnspan=2, pady=8)

    def delete_item_selected(self):
        if not self.require_modify():
            return
        sel = self.inv_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر صنفاً من الجدول أولاً.")
            return
        if not messagebox.askyesno("حذف صنف", f"حذف الصنف {sel[0]} نهائياً؟"):
            return
        try:
            delete_item(sel[0])
            self.load_inventory()
        except Exception as e:
            handle_error(e)

    def open_item_card(self, event=None):
        sel = self.inv_tree.selection()
        if not sel:
            return
        item_code = sel[0]
        with db() as conn:
            item_name = conn.execute("SELECT item_name FROM inventory WHERE item_code=?", (item_code,)).fetchone()[0]

        win = tk.Toplevel(self.root)
        win.title(f"كارت الصنف التفصيلي: {item_name} ({item_code})")
        win.geometry("1150x650")
        tk.Label(win, text=f"📊 كارت الصنف التفصيلي: {item_name} (الكود: {item_code})",
                 font=("Arial", 13, "bold"), fg="#1e293b", pady=10).pack()

        ff = ttk.LabelFrame(win, text=" فلترة الحركة بالفترة الزمنية ", padding=10)
        ff.pack(fill="x", padx=15, pady=5)
        d_from_e = date_entry(ff, 14, f"{date.today().year}-01-01")
        d_to_e = date_entry(ff, 14)
        state = {"rows": [], "final": ""}

        box = tk.Frame(win)
        box.pack(fill="both", expand=True, padx=15, pady=5)
        tree = make_tree(box, [("date", "التاريخ", 90, "center"), ("type", "نوع الحركة", 150, "center"),
                               ("party", "الطرف", 170, "e"), ("qty", "الكمية", 90, "center"),
                               ("bal", "الرصيد", 100, "center"), ("price", "السعر/التكلفة", 100, "center"),
                               ("val", "القيمة", 100, "center"), ("notes", "ملاحظات", 150, "center")])
        bottom = tk.Frame(win, pady=8)
        bottom.pack(fill="x")
        lbl = tk.Label(bottom, text="", font=("Arial", 11, "bold"), fg="#16a34a")
        lbl.pack(side="right", padx=20)

        def load():
            d1, d2 = parse_date(d_from_e.get()), parse_date(d_to_e.get())
            if not d1 or not d2:
                messagebox.showwarning("تنبيه", "صيغة التاريخ YYYY-MM-DD", parent=win)
                return
            rows, final = fetch_item_card(item_code, d1, d2)
            clear_tree(tree)
            for r in rows:
                tree.insert("", "end", values=r)
            state["rows"] = rows
            state["final"] = f"الرصيد في نهاية الفترة: {final:,.3f}"
            lbl.config(text=state["final"])

        def export():
            path = filedialog.asksaveasfilename(parent=win, defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                                title="حفظ كارت الصنف", initialfile=f"كارت_صنف_{item_code}.pdf")
            if not path:
                return
            try:
                build_item_card_pdf(path, item_code, item_name, state["rows"], state["final"])
                messagebox.showinfo("نجاح", "تم تصدير كارت الصنف إلى PDF", parent=win)
            except Exception as e:
                handle_error(e)

        rtl_row(ff, 0, [("من تاريخ:", d_from_e), ("إلى تاريخ:", d_to_e),
                        (None, money_button(ff, "تطبيق الفلتر 🔍", "#0284c7", load, 9))], slots=3)
        money_button(bottom, "🖨️ طباعة كارت الصنف (PDF)", "#2563eb", export).pack(side="left", padx=20)
        load()

    # ---------------------------------------------------------- الخزينة
    def update_party_options(self):
        cat = self.party_category_cb.get()
        pairs = []
        with db() as conn:
            if cat == CAT_PARTNER:
                pairs = [(f"{c} - {n}", (c, n)) for c, n in
                         conn.execute("SELECT partner_code, name FROM business_partners ORDER BY name")]
            elif cat == CAT_EXPENSE:
                # حسابات المصروفات الفرعية فقط، بدون تكلفة المبيعات والمردودات والإهلاك (يحسبها النظام)
                pairs = [(f"{c} - {n}", (c, n)) for c, n in conn.execute(
                    "SELECT account_code, account_name FROM chart_of_accounts WHERE account_type=? "
                    "AND COALESCE(is_group,0)=0 ORDER BY account_code", (TYPE_EXP,))
                    if not c.startswith(NON_CASH_EXPENSE_PREFIXES)]
                if not pairs:
                    pairs = [(n, (None, n)) for n in DEFAULT_EXPENSES]
            elif cat == CAT_EMPLOYEE:
                pairs = [(f"{c} - {n}", (c, n)) for c, n in employee_accounts(conn)]
            else:
                pairs = [(f"{c} - {n}", (None, n)) for c, n in
                         conn.execute("SELECT asset_code, asset_name FROM fixed_assets ORDER BY asset_code")]
                pairs.append(("أصول ثابتة - أخرى", (None, "أصول ثابتة - أخرى")))
        self.party_name_cb.set_items(pairs)
        self.party_name_cb.set('')
        self.refresh_open_invoices()

    def refresh_open_invoices(self):
        none_label = "بدون فاتورة محددة"
        pairs = [(none_label, None)]
        party = self.party_name_cb.selected()
        if self.party_category_cb.get() == CAT_PARTNER and party and party[0]:
            with db() as conn:
                if self.c_type.get().startswith("قبض"):
                    for i, d, tot, paid, no in conn.execute(
                            "SELECT invoice_id, invoice_date, total_amount - COALESCE((SELECT SUM(total_amount) FROM post_discounts "
                            "WHERE kind='sales' AND invoice_id=sales_invoices.invoice_id),0), paid_amount, "
                            "COALESCE(invoice_no, invoice_id) FROM sales_invoices "
                            "WHERE customer_code=? AND status='مفتوحة' ORDER BY invoice_id", (party[0],)):
                        pairs.append((f"مبيعات {no} ({d}) - متبقي: {tot - paid:,.2f}", ("sales", i)))
                else:
                    for i, d, tot, paid, no in conn.execute(
                            "SELECT purchase_id, purchase_date, total_amount - COALESCE((SELECT SUM(total_amount) FROM post_discounts "
                            "WHERE kind='purchase' AND invoice_id=purchase_invoices.purchase_id),0), paid_amount, "
                            "COALESCE(invoice_no, purchase_id) FROM purchase_invoices "
                            "WHERE supplier_code=? AND status='مفتوحة' ORDER BY purchase_id", (party[0],)):
                        pairs.append((f"مشتريات {no} ({d}) - متبقي: {tot - paid:,.2f}", ("purchase", i)))
        self.open_inv_cb.set_items(pairs)
        self.open_inv_cb.set(pairs[0][0])
        self.update_cash_rep_label()

    def update_cash_rep_label(self):
        """عرض المندوب الذي سيُورَّث للسند (من الفاتورة المختارة أو من بطاقة الطرف)."""
        lbl = getattr(self, "cash_rep_lbl", None)
        if lbl is None:
            return
        party = self.party_name_cb.selected()
        if self.party_category_cb.get() != CAT_PARTNER or not party or not party[0]:
            lbl.configure(text="👤 المندوب: —")
            return
        alloc = self.open_inv_cb.selected() or (None, None)
        try:
            with db() as conn:
                rep, src = voucher_rep_info(conn, party[0], parse_date(self.c_date.get()) or today(), *alloc)
        except Exception:  # noqa: BLE001  (عرض فقط)
            rep, src = None, ""
        lbl.configure(text=f"👤 المندوب: {rep} ({src} — تلقائي)" if rep else "👤 المندوب: بدون مندوب (الطرف بلا مندوب مسجل)")

    # ---------------------------------------------------------- الأصول الثابتة
    def setup_assets_tab(self, frame):
        f = ttk.LabelFrame(frame, text=" تسجيل أصل ثابت ", padding=10)
        f.pack(fill="x", padx=10, pady=8)
        self.ast_name_e = ttk.Entry(f, width=28)
        self.ast_date_e = date_entry(f, 14)
        self.ast_cost_e = ttk.Entry(f, width=14)
        self.ast_salvage_e = ttk.Entry(f, width=14)
        self.ast_salvage_e.insert(0, "0")
        self.ast_life_e = ttk.Entry(f, width=8)
        self.ast_life_e.insert(0, "5")
        rtl_row(f, 0, [("اسم الأصل:", self.ast_name_e), ("تاريخ الشراء:", self.ast_date_e), ("التكلفة:", self.ast_cost_e)], slots=3)
        rtl_row(f, 1, [("قيمة الخردة:", self.ast_salvage_e), ("العمر (سنوات):", self.ast_life_e)], slots=3)
        money_button(f, "حفظ الأصل وحساب الإهلاك (يبدأ من الشهر التالي للشراء) 🏢", "#16a34a", self.save_fixed_asset
                     ).grid(row=2, column=0, columnspan=6, pady=8)

        box = ttk.LabelFrame(frame, text=" جدول الأصول الثابتة ", padding=10)
        box.pack(fill="both", expand=True, padx=10, pady=5)
        self.ast_tree = make_tree(box, [
            ("code", "الكود", 80, "center"), ("name", "اسم الأصل", 220, "e"), ("date", "تاريخ الشراء", 95, "center"),
            ("cost", "التكلفة", 100, "center"), ("salv", "الخردة", 80, "center"), ("life", "العمر", 60, "center"),
            ("monthly", "إهلاك شهري", 100, "center"), ("acc", "مجمع الإهلاك", 110, "center"),
            ("book", "القيمة الدفترية", 110, "center")])
        self.load_assets()

    def save_fixed_asset(self):
        try:
            name = self.ast_name_e.get().strip()
            p_date = parse_date(self.ast_date_e.get())
            cost = parse_float(self.ast_cost_e.get())
            salvage = parse_float(self.ast_salvage_e.get() or "0")
            life_f = parse_float(self.ast_life_e.get())
            if not name:
                raise UserError("أدخل اسم الأصل.")
            if not p_date:
                raise UserError("صيغة التاريخ غير صحيحة، استخدم YYYY-MM-DD")
            if cost is None or cost <= 0 or salvage is None or salvage < 0 or salvage >= cost:
                raise UserError("التكلفة يجب أن تكون أكبر من الصفر وقيمة الخردة أقل من التكلفة.")
            if life_f is None or life_f < 1 or life_f != int(life_f):
                raise UserError("العمر الإنتاجي يجب أن يكون عدداً صحيحاً من السنوات (1 على الأقل).")
            life = int(life_f)
            monthly = (cost - salvage) / (life * 12)
            with db() as conn:
                nxt = conn.execute("SELECT COALESCE(MAX(asset_id),0)+1 FROM fixed_assets").fetchone()[0]
                conn.execute("INSERT INTO fixed_assets (asset_code, asset_name, purchase_date, purchase_cost, salvage_value, "
                             "useful_life_years, monthly_depreciation) VALUES (?,?,?,?,?,?,?)",
                             (f"AST-{nxt:03d}", name, p_date, cost, salvage, life, monthly))
            messagebox.showinfo("نجاح", f"تم تسجيل الأصل، الإهلاك الشهري: {fmt(monthly)}")
            self.ast_name_e.delete(0, tk.END)
            self.ast_cost_e.delete(0, tk.END)
            self.load_assets()
        except Exception as e:
            handle_error(e)

    def load_assets(self):
        clear_tree(self.ast_tree)
        now = date.today()
        with db() as conn:
            for code, name, p_date, cost, salv, life, monthly in conn.execute(
                    "SELECT asset_code, asset_name, purchase_date, purchase_cost, salvage_value, useful_life_years, "
                    "monthly_depreciation FROM fixed_assets ORDER BY asset_id"):
                acc = accumulated_depreciation(cost, salv, life, p_date, now)
                self.ast_tree.insert("", "end", iid=code, values=(code, name, p_date, fmt(cost), fmt(salv), life,
                                                                  fmt(monthly), fmt(acc), fmt(cost - acc)))

    # ---------------------------------------------------------- كشف الحساب
    def setup_statement_tab(self, frame):
        f = ttk.LabelFrame(frame, text=" كشف الحساب الموحد: اختر طرفاً لعرض حركاته، أو اتركه فارغاً لعرض ملخص أرصدة الفئة ",
                           padding=10)
        f.pack(fill="x", padx=10, pady=8)
        self.stmt_partner_cb = SearchableCombobox(f, width=38)
        self.stmt_scope_cb = ttk.Combobox(f, values=list(SUMMARY_SCOPES.values()), width=24, state="readonly")
        self.stmt_scope_cb.set(SUMMARY_SCOPES["both"])
        self.stmt_scope_cb.bind("<<ComboboxSelected>>", lambda e: self.show_statement())
        clear = money_button(f, "✖ إلغاء اختيار الطرف (ملخص)", "#475569", self.clear_statement_partner, 9)
        rtl_row(f, 0, [("العميل / المورد (اختياري):", self.stmt_partner_cb), (None, clear), ("الفئة:", self.stmt_scope_cb)],
                slots=3)
        self.stmt_from_e = DateEntry(f, width=12)
        self.stmt_to_e = DateEntry(f, width=12)
        for e_ in (self.stmt_from_e, self.stmt_to_e):
            e_.bind("<<DateSelected>>", lambda e: self.show_statement(), add="+")
            e_.bind("<Return>", lambda e: self.show_statement(), add="+")
        self.stmt_q_e = ttk.Entry(f, width=22)
        self.stmt_q_e.bind("<Return>", lambda e: self.show_statement())
        self.stmt_hide_zero = tk.BooleanVar(value=False)
        hz = ttk.Checkbutton(f, text="إخفاء الأرصدة الصفرية", variable=self.stmt_hide_zero, command=self.show_statement)
        rtl_row(f, 1, [("من تاريخ:", self.stmt_from_e), ("إلى تاريخ:", self.stmt_to_e), ("بحث في الملخص:", self.stmt_q_e),
                       (None, hz)], slots=4)
        btns = tk.Frame(f)
        btns.grid(row=2, column=0, columnspan=8, sticky="e", pady=(4, 0))
        money_button(btns, "عرض 🔍", "#0284c7", self.show_statement, 10).pack(side="right", padx=3)
        money_button(btns, "📄 تصدير PDF", "#dc2626", self.export_statement_pdf, 10).pack(side="right", padx=3)
        money_button(btns, "📊 تصدير Excel", UI["teal"], self.export_statement_excel, 10).pack(side="right", padx=3)
        money_button(btns, "↺ كل الفترات (رصيد لحظي)", "#64748b", self.clear_statement_dates, 9).pack(side="right", padx=3)
        tk.Label(btns, text="التواريخ فارغة = رصيد تراكمي من بداية العمل حتى الآن", fg=UI["muted"],
                 font=(UI_FONT, 8)).pack(side="right", padx=8)
        self.stmt_links_lbl = tk.Label(f, text="", fg=UI["teal"], font=(UI_FONT, 9, "bold"))
        self.stmt_links_lbl.grid(row=3, column=0, columnspan=8, sticky="e", padx=6)

        self.stmt_box = box = ttk.LabelFrame(frame, text=" التفاصيل (الرصيد = مدين - دائن، والمدين يعني أن الطرف مدين لنا) ",
                                             padding=10)
        box.pack(fill="both", expand=True, padx=10, pady=5)
        self.stmt_header_lbl = ttk.Label(box, text="", font=("Arial", 10, "bold"))
        self.stmt_header_lbl.pack(fill="x", pady=2)
        self.stmt_summary_lbl = tk.Label(box, text="", font=("Arial", 11, "bold"), bg="#e2e8f0", pady=8)
        self.stmt_summary_lbl.pack(fill="x", pady=5)
        # جدول الحركات التفصيلية (طرف محدد)
        self.stmt_detail_holder = tk.Frame(box)
        self.stmt_tree = make_tree(self.stmt_detail_holder, [
            ("ref", "المرجع", 100, "center"), ("date", "التاريخ", 90, "center"), ("type", "النوع", 110, "center"),
            ("desc", "البيان", 260, "e"), ("debit", "مدين", 100, "center"), ("credit", "دائن", 100, "center"),
            ("bal", "الرصيد", 130, "center")])
        # جدول الملخص (سطر واحد لكل طرف)
        self.stmt_sum_holder = tk.Frame(box)
        self.stmt_sum_tree = make_tree(self.stmt_sum_holder, [
            ("code", "الكود", 110, "center"), ("name", "الاسم", 240, "e"), ("role", "الصفة", 120, "center"),
            ("prior", "رصيد سابق", 120, "center"), ("debit", "مدين الفترة", 110, "center"),
            ("credit", "دائن الفترة", 110, "center"), ("bal", "الرصيد", 140, "center"), ("n", "عدد الحركات", 80, "center"),
            ("last", "آخر حركة", 90, "center")])
        dc = self.stmt_sum_tree["displaycolumns"]
        self._stmt_sum_cols = list(self.stmt_sum_tree["columns"] if dc in ("#all", ("#all",)) else dc)
        self.stmt_sum_tree.tag_configure("dr", foreground="#1d4ed8")
        self.stmt_sum_tree.tag_configure("cr", foreground="#b45309")
        self.stmt_sum_tree.bind("<Double-1>", self._open_summary_row)
        self.stmt_sum_tree.bind("<Return>", self._open_summary_row)
        self.stmt_keys = {}
        self.stmt_detail_win = None
        self.stmt_hint_lbl = tk.Label(box, text="", fg="#0369a1", font=("Arial", 9, "bold"))
        self.stmt_hint_lbl.pack(side="bottom", anchor="e", pady=(4, 0))
        self.stmt_tree.bind("<ButtonRelease-1>", self._on_statement_click)
        self.stmt_tree.bind("<Return>", lambda e: self._open_statement_row(self.stmt_tree.focus()))
        self.stmt_partner_cb.bind("<<ComboboxSelected>>", lambda e: self.show_statement(), add="+")
        self.stmt_shown_code = None
        self.stmt_mode = None
        self._stmt_set_mode("summary")

    # ------------ أدوات شاشة الكشف
    def _stmt_scope(self):
        return {v: k for k, v in SUMMARY_SCOPES.items()}.get(self.stmt_scope_cb.get(), "both")

    def _stmt_dates(self):
        out = []
        for e_, lbl in ((self.stmt_from_e, "من تاريخ"), (self.stmt_to_e, "إلى تاريخ")):
            v = e_.get().strip()
            if v:
                d = parse_date(v)
                if not d:
                    raise UserError(f"«{lbl}» بصيغة YYYY-MM-DD أو اتركه فارغاً.")
                out.append(d)
            else:
                out.append(None)
        if out[0] and out[1] and out[0] > out[1]:
            raise UserError("تاريخ البداية بعد تاريخ النهاية.")
        return out

    def _stmt_set_mode(self, mode):
        if mode == self.stmt_mode:
            return
        self.stmt_mode = mode
        self.stmt_detail_holder.pack_forget()
        self.stmt_sum_holder.pack_forget()
        (self.stmt_sum_holder if mode == "summary" else self.stmt_detail_holder).pack(fill="both", expand=True)
        self.stmt_box.configure(text=(" ملخص الأرصدة: سطر واحد لكل طرف (اضغط مرتين على الطرف لعرض كشفه التفصيلي) "
                                      if mode == "summary" else
                                      " التفاصيل (الرصيد = مدين - دائن، والمدين يعني أن الطرف مدين لنا) "))
        self.stmt_hint_lbl.config(text=("💡 اضغط مرتين على أي طرف لعرض حركاته الكاملة." if mode == "summary" else
                                        "💡 انقر على أي فاتورة أو حركة في الكشف لعرض تفاصيلها الكاملة (الأصناف والكميات والأسعار)."))

    def clear_statement_partner(self):
        self.stmt_partner_cb.set("")
        self.stmt_shown_code = None
        self.show_statement()

    def clear_statement_dates(self):
        self.stmt_from_e.delete(0, tk.END)
        self.stmt_to_e.delete(0, tk.END)
        self.show_statement()

    def _stmt_allowed_scope(self, scope):
        cust_ok, supp_ok = self.can_fin("customer"), self.can_fin("supplier")
        if scope == "customer" and not cust_ok or scope == "supplier" and not supp_ok:
            raise UserError("لا توجد صلاحية لعرض البيانات المالية لهذه الفئة.")
        if scope == "both" and not (cust_ok and supp_ok):
            return "customer" if cust_ok else "supplier"
        return scope

    def show_statement(self, quiet=False):
        """طرف محدد ← كشف تفصيلي بالحركات. بدون طرف ← ملخص الأرصدة (سطر لكل طرف) للفئة المختارة."""
        code = self.stmt_partner_cb.selected() if self.stmt_partner_cb.get().strip() else None
        if code:
            self.show_partner_statement(code, quiet=quiet)
        else:
            self.show_balances_summary(quiet=quiet)

    def show_balances_summary(self, quiet=False):
        try:
            d_from, d_to = self._stmt_dates()
            scope = self._stmt_allowed_scope(self._stmt_scope())
            res = partners_balance_summary(scope, d_from, d_to, self.stmt_q_e.get(), self.stmt_hide_zero.get())
            self._stmt_set_mode("summary")
            self.stmt_shown_code = None
            self.stmt_links_lbl.config(text="")
            clear_tree(self.stmt_sum_tree)
            # عمود «رصيد سابق» يظهر فقط عند تحديد «من تاريخ»
            self.stmt_sum_tree["displaycolumns"] = [c for c in self._stmt_sum_cols if c != "prior" or d_from]
            for r, vals in zip(res["rows"], summary_table_rows(res, bool(d_from))):
                tag = "dr" if r["balance"] > 0.005 else ("cr" if r["balance"] < -0.005 else "")
                self.stmt_sum_tree.insert("", "end", iid=r["code"], values=vals, tags=(tag,) if tag else ())
            t = res["totals"]
            self.stmt_header_lbl.config(text=f"ملخص أرصدة {SUMMARY_SCOPES[scope]}  |  {period_label(d_from, d_to)}  |  "
                                             f"عدد الأطراف: {len(res['rows'])}")
            self.stmt_summary_lbl.config(text=f"مدين (لنا): {fmt(t['receivable'])}  |  دائن (علينا): {fmt(t['payable'])}  |  "
                                              f"الصافي: {balance_text(t['net'])}")
        except Exception as e:
            if not quiet:
                handle_error(e)

    def _open_summary_row(self, event=None):
        sel = self.stmt_sum_tree.selection()
        if not sel:
            return
        code = sel[0]
        label = next((lbl for lbl, c in self.stmt_partner_cb._map.items() if c == code), None)
        if label is None:
            self.update_statement_partners_dropdown()
            label = next((lbl for lbl, c in self.stmt_partner_cb._map.items() if c == code), None)
        if label:
            self.stmt_partner_cb.set(label)
        self.show_partner_statement(code)

    def export_statement_excel(self):
        try:
            d_from, d_to = self._stmt_dates()
            code = self.stmt_partner_cb.selected() if self.stmt_partner_cb.get().strip() else None
            scope = self._stmt_scope()
            if code:
                _info, rows, *_ = fetch_partner_statement(code, scope=scope, d_from=d_from, d_to=d_to)
                df = pd.DataFrame(rows, columns=["المرجع", "التاريخ", "النوع", "البيان", "مدين", "دائن", "الرصيد"])
                name = f"كشف_حساب_{code}.xlsx"
            else:
                scope = self._stmt_allowed_scope(scope)
                res = partners_balance_summary(scope, d_from, d_to, self.stmt_q_e.get(), self.stmt_hide_zero.get())
                df = pd.DataFrame(summary_table_rows(res, bool(d_from)), columns=SUMMARY_HEADERS)
                name = f"ملخص_أرصدة_{scope}.xlsx"
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")], initialfile=name)
            if not path:
                return
            df.to_excel(path, index=False)
            messagebox.showinfo("نجاح", "تم التصدير إلى Excel.")
        except Exception as e:
            handle_error(e)

    def _refresh_open_statement(self, partner_code=None):
        """يعيد تحميل الكشف المعروض (بعد حفظ فاتورة أو العودة للصفحة) حتى تفتح النقرات الفاتورة الصحيحة دائماً."""
        code = getattr(self, "stmt_shown_code", None)
        if code and (partner_code is None or partner_code == code):
            self.show_partner_statement(code, quiet=True)
        elif getattr(self, "stmt_mode", None) == "summary" and partner_code is None:
            self.show_balances_summary(quiet=True)

    def _on_statement_click(self, event):
        if self.stmt_tree.identify_region(event.x, event.y) not in ("cell", "tree"):
            return  # النقر على رأس العمود (للترتيب) لا يفتح شيئاً
        iid = self.stmt_tree.identify_row(event.y)
        if iid:
            self.stmt_tree.selection_set(iid)
            self.stmt_tree.focus(iid)
        self._open_statement_row(iid)

    def _open_statement_row(self, iid):
        key = self.stmt_keys.get(iid) if iid else None
        if key:
            try:
                self.show_transaction_details(key)
            except Exception as e:
                handle_error(e)
                self._refresh_open_statement()  # الحركة حُذفت أو تغيّرت: تحديث الكشف

    def show_transaction_details(self, key):
        """نافذة تفاصيل حركة كشف الحساب (فاتورة مبيعات/مشتريات، مرتجع، سند، رصيد افتتاحي).
        تُستخدم نافذة واحدة يُعاد ملؤها مع كل نقرة بدلاً من فتح نوافذ كثيرة."""
        kind, ref_id = key
        header, columns, lines, totals = [], [], [], []
        with db() as conn:
            if kind == "purchase":
                h = conn.execute(
                    "SELECT COALESCE(invoice_no, purchase_id), purchase_date, supplier_code, supplier_name, subtotal, vat_amount, "
                    "total_amount, paid_amount, status, COALESCE(notes,''), COALESCE(source_type,'مباشر'), "
                    "COALESCE(supplier_ref,'') FROM purchase_invoices WHERE purchase_id=?", (ref_id,)).fetchone()
                if not h:
                    raise UserError("الفاتورة غير موجودة (ربما حُذفت). حدّث الكشف.")
                title = f"فاتورة مشتريات رقم {h[0]}" + (" (شراء مباشر)" if h[10] == "مباشر" else f" ({h[10]})")
                header = [("رقم الفاتورة", h[0]), ("التاريخ", h[1]), ("المورد", f"{h[2]} - {h[3]}"),
                          ("مرجع المورد", h[11] or "-"), ("المصدر", h[10]), ("الحالة", h[8]), ("ملاحظات", h[9] or "-")]
                columns = PURCHASE_LINE_COLUMNS
                for r in conn.execute("SELECT item_code, item_name, quantity, unit, unit_cost, subtotal, vat_amount, total_amount, "
                                      "production_date, expiry_date FROM purchase_invoice_lines WHERE purchase_id=? "
                                      "ORDER BY line_id", (ref_id,)):
                    lines.append((r[0], r[1], f"{r[2]:g}", r[3] or "-", fmt(r[4]), fmt(r[5]), fmt(r[6]), fmt(r[7]),
                                  r[8] or "-", r[9] or "-"))
                totals = [("قبل الضريبة", h[4]), ("الضريبة", h[5]), ("الإجمالي", h[6]), ("المسدد", h[7]),
                          ("المتبقي", max(h[6] - h[7], 0.0))]
            elif kind == "sales":
                h = conn.execute(
                    "SELECT COALESCE(invoice_no, invoice_id), invoice_date, customer_code, customer_name, subtotal, vat_amount, "
                    "total_amount, paid_amount, status, COALESCE(notes,''), COALESCE(source_type,'مباشر'), "
                    "COALESCE(sales_rep,'') FROM sales_invoices WHERE invoice_id=?", (ref_id,)).fetchone()
                if not h:
                    raise UserError("الفاتورة غير موجودة (ربما حُذفت). حدّث الكشف.")
                title = f"فاتورة مبيعات رقم {h[0]}"
                header = [("رقم الفاتورة", h[0]), ("التاريخ", h[1]), ("العميل", f"{h[2]} - {h[3]}"),
                          ("المندوب", h[11] or "-"), ("المصدر", h[10]), ("الحالة", h[8]), ("ملاحظات", h[9] or "-")]
                columns = [("code", "الكود", 75, "center"), ("name", "الصنف", 210, "e"), ("qty", "الكمية", 65, "center"),
                           ("unit", "الوحدة", 55, "center"), ("price", "سعر الوحدة", 85, "center"),
                           ("disc", "الخصم الخاص", 90, "center"), ("subtotal", "الخاضع للضريبة", 115, "center"),
                           ("vat", "الضريبة 15%", 85, "center"), ("total", "الإجمالي", 100, "center")]
                for r in conn.execute("SELECT item_code, item_name, quantity, unit, unit_price, COALESCE(discount_amount,0), "
                                      "subtotal, vat_amount, total_amount FROM sales_invoice_lines WHERE invoice_id=? "
                                      "ORDER BY line_id", (ref_id,)):
                    lines.append((r[0], r[1], f"{r[2]:g}", r[3] or "-", fmt(r[4]), fmt(r[5]) if r[5] > 0.005 else "-",
                                  fmt(r[6]), fmt(r[7]), fmt(r[8])))
                totals = [("قبل الضريبة", h[4]), ("الضريبة", h[5]), ("الإجمالي", h[6]), ("المسدد", h[7]),
                          ("المتبقي", max(h[6] - h[7], 0.0))]
            elif kind == "return":
                r = conn.execute("SELECT return_id, return_date, return_type, party_code, party_name, item_code, item_name, "
                                 "quantity, unit_price, subtotal, vat_amount, total_amount, COALESCE(notes,'') "
                                 "FROM returns_invoices WHERE return_id=?", (ref_id,)).fetchone()
                if not r:
                    raise UserError("المرتجع غير موجود (ربما حُذف). حدّث الكشف.")
                title = f"{r[2]} رقم {r[0]}"
                header = [("رقم المرتجع", r[0]), ("التاريخ", r[1]), ("النوع", r[2]), ("الطرف", f"{r[3]} - {r[4]}"),
                          ("ملاحظات", r[12] or "-")]
                columns = [("code", "الكود", 75, "center"), ("name", "الصنف", 230, "e"), ("qty", "الكمية", 70, "center"),
                           ("price", "سعر الوحدة", 95, "center"), ("subtotal", "قبل الضريبة", 110, "center"),
                           ("vat", "الضريبة 15%", 90, "center"), ("total", "الإجمالي", 110, "center")]
                lines = [(r[5], r[6], f"{r[7]:g}", fmt(r[8]), fmt(r[9]), fmt(r[10]), fmt(r[11]))]
                totals = [("قبل الضريبة", r[9]), ("الضريبة", r[10]), ("الإجمالي", r[11])]
            elif kind == "cash":
                r = conn.execute("SELECT trans_id, trans_type, trans_date, account_bank, party_code, party_name, amount, "
                                 "COALESCE(alloc_type,''), alloc_id, COALESCE(description,'') FROM cash_transactions "
                                 "WHERE trans_id=?", (ref_id,)).fetchone()
                if not r:
                    raise UserError("السند غير موجود (ربما حُذف). حدّث الكشف.")
                title = f"سند {r[1]} رقم {r[0]}"
                alloc = "-"
                if r[8]:
                    alloc_kind = ("sales" if r[7] in ("sales", "مبيعات") else
                                  "purchase" if r[7] in ("purchase", "مشتريات") else None)
                    alloc = f"فاتورة {invoice_number(alloc_kind, r[8])}" if alloc_kind else f"{r[7]} #{r[8]}"
                header = [("رقم السند", r[0]), ("النوع", f"سند {r[1]}"), ("التاريخ", r[2]), ("الخزينة / البنك", r[3]),
                          ("الطرف", f"{r[4] or '-'} - {r[5]}"), ("مخصص لـ", alloc), ("البيان", r[9] or "-")]
                totals = [("المبلغ", r[6])]
            elif kind == "discount":
                d, dl = load_post_discount(ref_id)
                title = f"{DISC_KIND_SHORT[d['kind']]} رقم {d['discount_no']}"
                header = [("رقم الإشعار", d["discount_no"]), ("التاريخ", d["discount_date"]),
                          ("الطرف", f"{d['party_code']} - {d['party_name']}"), ("الفاتورة", d["invoice_no"] or "-"),
                          ("المعالجة", DISC_METHOD_LABEL[d["method"]] if d["kind"] == DISC_PURCHASE else "خصم مسموح به"),
                          ("السبب", d["reason"] or "-"), ("ملاحظات", d["notes"] or "-"),
                          ("القيد الآلي", d["journal_entry_id"] or "غير مرحّل (راجع حسابات القيد)"),
                          ("بواسطة", d["created_by"] or "-")]
                if dl:
                    columns = [("code", "الكود", 90, "center"), ("name", "الصنف", 260, "e"),
                               ("amount", "الخصم الموزع", 120, "center"), ("applied", "خُفّض من تكلفة المخزون", 150, "center"),
                               ("rest", "خصم مكتسب (بضاعة مباعة)", 160, "center")]
                    lines = [(c_, n_, fmt(a_), fmt(ap_), fmt(a_ - ap_)) for c_, n_, a_, ap_ in dl]
                totals = [("قبل الضريبة", d["subtotal"]), ("الضريبة", d["vat_amount"]), ("الإجمالي", d["total_amount"])]
            elif kind == "opening":
                r = conn.execute("SELECT partner_code, name, COALESCE(opening_debit,0), COALESCE(opening_credit,0), "
                                 "COALESCE(opening_date,'') FROM business_partners WHERE partner_code=?", (ref_id,)).fetchone()
                if not r:
                    raise UserError("الشريك غير موجود.")
                title = f"الرصيد الافتتاحي - {r[1]}"
                header = [("الطرف", f"{r[0]} - {r[1]}"), ("تاريخ الرصيد", r[4] or "-")]
                totals = [("مدين (ما عليه)", r[2]), ("دائن (ما له)", r[3])]
            else:
                return

        if kind in ("sales", "purchase"):
            wins = self.__dict__.setdefault("_invoice_wins", {})
            win = wins.get(key)
            if win is None or not win.winfo_exists():
                win = wins[key] = tk.Toplevel(self.root)
                win.geometry("1060x460")
                win.transient(self.root)
            self.trail_push("invoice", key, f"🧾 {title}")
        else:
            win = self.stmt_detail_win
            if win is None or not win.winfo_exists():
                win = self.stmt_detail_win = tk.Toplevel(self.root)
                win.geometry("1060x460")
                win.transient(self.root)
        for child in win.winfo_children():
            child.destroy()
        win.title(title)
        tk.Label(win, text=title, font=("Arial", 13, "bold"), bg="#0f172a", fg="white", pady=8).pack(fill="x")
        info = tk.Frame(win, bg="#f1f5f9", padx=10, pady=6)
        info.pack(fill="x")
        for i, (lab, val) in enumerate(header):
            r_, c_ = divmod(i, 3)
            cell = tk.Frame(info, bg="#f1f5f9")
            cell.grid(row=r_, column=2 - c_, sticky="e", padx=10, pady=2)
            tk.Label(cell, text=str(val), bg="#f1f5f9", fg="#0f172a", font=("Arial", 10)).pack(side="right")
            tk.Label(cell, text=f"{lab}:", bg="#f1f5f9", fg="#475569", font=("Arial", 10, "bold")).pack(side="right")
        for c in range(3):
            info.grid_columnconfigure(c, weight=1)
        tot = tk.Frame(win, bg="#e2e8f0", pady=6)
        tot.pack(fill="x", side="bottom")
        tk.Label(tot, text="  |  ".join(f"{lab}: {fmt(val)}" for lab, val in totals), bg="#e2e8f0",
                 font=("Arial", 11, "bold")).pack()
        if columns:
            box = tk.Frame(win)
            box.pack(fill="both", expand=True, padx=10, pady=8)
            tree = make_tree(box, columns, rtl=True)
            for v in lines:
                tree.insert("", "end", values=v)
            if not lines:
                tk.Label(box, text="لا توجد بنود مسجلة لهذه الفاتورة.", fg="#dc2626").pack()
        self._modal_extras(win, kind, ref_id)
        win.deiconify()
        win.lift()
        win.focus_force()

    def update_statement_partners_dropdown(self):
        with db() as conn:
            rows = conn.execute("SELECT partner_code, name, is_customer, is_supplier FROM business_partners ORDER BY name").fetchall()
        cust_ok, supp_ok = self.can_fin("customer"), self.can_fin("supplier")
        pairs = [(f"{c} - {n}", c) for c, n, is_c, is_s in rows
                 if (is_c and cust_ok) or (is_s and supp_ok) or (not is_c and not is_s)]
        self.stmt_partner_cb.set_items(pairs)
        self._refresh_open_statement()  # العودة للصفحة: الكشف المعروض يُحدَّث بآخر الفواتير

    def show_partner_statement(self, code=None, quiet=False):
        try:
            code = code or self.stmt_partner_cb.selected()
            if not code:
                raise UserError("اختر الشريك من القائمة.")
            d_from, d_to = self._stmt_dates()
            scope = self._stmt_scope()
            info, rows, tot_d, tot_c, bal, keys = fetch_partner_statement(code, with_keys=True, scope=scope,
                                                                          d_from=d_from, d_to=d_to)
            parties = statement_parties(code)
            self.stmt_links_lbl.config(text=("🔗 كيان مزدوج مرتبط تلقائياً (مقاصة): " +
                                             "  +  ".join(f"{c} {n} ({r})" for c, n, r in parties)) if len(parties) > 1 else "")
            self._stmt_set_mode("detail")
            clear_tree(self.stmt_tree)
            self.stmt_keys = {}
            self.stmt_shown_code = code
            for r, key in zip(rows, keys):
                iid = self.stmt_tree.insert("", "end", values=r)
                self.stmt_keys[iid] = key
            self.stmt_header_lbl.config(text=f"الطرف: {info[0]}  |  هاتف: {info[1] or '-'}  |  الرقم الضريبي: {info[3] or '-'}"
                                             f"  |  {period_label(d_from, d_to)}")
            side = "مدين" if bal > 0.005 else ("دائن" if bal < -0.005 else "")
            self.stmt_summary_lbl.config(text=f"إجمالي المدين: {fmt(tot_d)}  |  إجمالي الدائن: {fmt(tot_c)}  |  "
                                              f"الرصيد النهائي: {abs(bal):,.2f} {side}")
        except Exception as e:
            if quiet:
                self.stmt_shown_code = None
            else:
                handle_error(e)

    def export_statement_pdf(self):
        try:
            d_from, d_to = self._stmt_dates()
            code = self.stmt_partner_cb.selected() if self.stmt_partner_cb.get().strip() else None
            scope = self._stmt_scope()
            name = f"كشف_حساب_{code}.pdf" if code else f"ملخص_أرصدة_{scope}.pdf"
            path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                                title="حفظ كشف الحساب", initialfile=name)
            if not path:
                return
            if code:
                build_statement_pdf(path, code, scope, d_from, d_to)
            else:
                build_balances_summary_pdf(path, self._stmt_allowed_scope(scope), d_from, d_to, self.stmt_q_e.get(),
                                           self.stmt_hide_zero.get())
            messagebox.showinfo("نجاح", "تم تصدير PDF بنجاح!")
        except Exception as e:
            handle_error(e)

    # ---------------------------------------------------------- قائمة الدخل
    def setup_income_tab(self, frame):
        f = ttk.Frame(frame, padding=12)
        f.pack(fill="both", expand=True)
        bar = tk.Frame(f)
        bar.pack(fill="x")
        self.inc_from_e = date_entry(bar, 14, f"{date.today().year}-01-01")
        self.inc_to_e = date_entry(bar, 14)
        rtl_row(bar, 0, [("من تاريخ:", self.inc_from_e), ("إلى تاريخ:", self.inc_to_e),
                         (None, money_button(bar, "🔄 تحديث القائمة", "#0284c7", self.generate_income_statement, 10))], slots=3)
        self.income_text = tk.Text(f, font=("Courier New", 11), bg="#f8fafc", fg="#0f172a", wrap="none", padx=12, pady=12)
        self.income_text.pack(expand=True, fill="both", pady=8)

    def generate_income_statement(self):
        try:
            d1, d2 = parse_date(self.inc_from_e.get()), parse_date(self.inc_to_e.get())
            if not d1 or not d2:
                raise UserError("صيغة التاريخ YYYY-MM-DD")
            if d1 > d2:
                raise UserError("تاريخ البداية بعد تاريخ النهاية.")
            self.income_text.delete("1.0", tk.END)
            self.income_text.insert(tk.END, build_income_report(d1, d2))
        except Exception as e:
            handle_error(e)

    # ---------------------------------------------------------- بيانات الشركة
    def setup_company_tab(self, frame):
        """صفحة الإعدادات داخل إطار قابل للتمرير (بعرض كامل + عجلة الفأرة)."""
        sf = ScrollFrame(frame, height=600)
        sf.pack(fill="both", expand=True)
        sf.canvas.configure(bg=UI["bg"])
        item = sf.canvas.find_all()[0]
        sf.canvas.bind("<Configure>", lambda e: sf.canvas.itemconfigure(item, width=e.width))
        wheel = lambda ev: sf.canvas.yview_scroll(int(-ev.delta / 120), "units")
        sf.canvas.bind("<Enter>", lambda e: sf.canvas.bind_all("<MouseWheel>", wheel))
        sf.canvas.bind("<Leave>", lambda e: sf.canvas.unbind_all("<MouseWheel>"))
        self._setup_company_content(sf.inner)

    def _setup_company_content(self, frame):
        f = ttk.LabelFrame(frame, text=" المنشأة المُصدِرة للفواتير والإشعارات (البيانات القانونية على المستندات PDF) ", padding=12)
        f.pack(fill="x", padx=10, pady=10)
        self.comp_entries = {}
        fields = [("company_name", "اسم المنشأة (عربي):"), ("company_name_en", "Company Name (EN):"),
                  ("vat_number", "الرقم الضريبي:"), ("cr_number", "السجل التجاري:"),
                  ("address", "العنوان (عربي):"), ("address_en", "Address (EN):"),
                  ("bank_name", "البنك:"), ("iban", "رقم الآيبان IBAN:"), ("phone", "الهاتف:"),
                  ("rep_name", "اسم المندوب الافتراضي:")]
        for i in range(0, len(fields), 2):
            pair = []
            for key, label in fields[i:i + 2]:
                e = (ttk.Combobox(f, width=40, values=[""] + rep_names(), state="readonly") if key == "rep_name"
                     else ttk.Entry(f, width=42))
                self.comp_entries[key] = e
                pair.append((label, e))
            rtl_row(f, i // 2, pair, slots=2, pady=4)
        tk.Label(f, text="هوية البرنامج «نظام الحاجري المالي» منفصلة؛ هذه البيانات تظهر فقط على الفواتير والإشعارات المطبوعة.",
                 fg=UI["muted"], font=(UI_FONT, 8)).grid(row=len(fields) // 2, column=0, columnspan=4, sticky="e", pady=(4, 0))
        money_button(f, "حفظ بيانات المنشأة 💾", "#16a34a", self.save_company).grid(row=len(fields) // 2 + 1, column=0,
                                                                                 columnspan=4, pady=10)
        if self.user["is_admin"]:
            hf = ttk.LabelFrame(frame, text=" 🩺 فحص سلامة البيانات (مراجعة محاسبية ومخزنية شاملة - قراءة فقط) ", padding=10)
            hf.pack(fill="x", padx=10, pady=5)
            tk.Label(hf, text="يفحص: توازن القيود وميزان المراجعة، تطابق الفواتير مع بنودها والضريبة 15%، المسدد مع السندات، "
                              "أرصدة المخزون مع الحركات، الأرقام المكررة، والمراجع المفقودة.",
                     fg=UI["muted"], font=(UI_FONT, 9), wraplength=900, justify="right").pack(anchor="e")
            money_button(hf, "🩺 تشغيل الفحص الآن", UI["primary"], self.open_integrity_window, 10).pack(pady=6)
        pf = ttk.LabelFrame(frame, text=" الأمان ", padding=10)
        pf.pack(fill="x", padx=10, pady=5)
        self.admin_pwd_e = ttk.Entry(pf, width=24, show="*")
        rtl_row(pf, 0, [("كلمة مرور المدير (لفتح الأيام المغلقة):", self.admin_pwd_e)], slots=1)
        money_button(pf, "حفظ كلمة المرور", "#16a34a", self.save_admin_password, 9).grid(row=1, column=0, columnspan=2, pady=6)
        if self.user["is_admin"]:
            sf = ttk.LabelFrame(frame, text=" الترقيم التسلسلي المستقل (للمدير) ", padding=10)
            sf.pack(fill="x", padx=10, pady=5)
            for col, h in enumerate(("مثال الرقم التالي", "عدد الأرقام", "الرقم التالي", "البادئة", "التسلسل")):
                ttk.Label(sf, text=h, font=("Arial", 9, "bold")).grid(row=0, column=col, padx=10, pady=3)
            self.seq_entries = {}
            for i, (name, label, *_) in enumerate(SEQUENCE_DEFS, start=1):
                ttk.Label(sf, text=label).grid(row=i, column=4, sticky="e", padx=10)
                pre, nxt, wid = ttk.Entry(sf, width=10), ttk.Entry(sf, width=12), ttk.Entry(sf, width=6)
                pre.grid(row=i, column=3, padx=10, pady=2)
                nxt.grid(row=i, column=2, padx=10, pady=2)
                wid.grid(row=i, column=1, padx=10, pady=2)
                prev = ttk.Label(sf, text="", foreground="#0369a1")
                prev.grid(row=i, column=0, padx=10)
                self.seq_entries[name] = (pre, nxt, wid, prev)
            money_button(sf, "حفظ إعدادات الترقيم 💾", "#16a34a", self.save_sequences, 9).grid(
                row=len(SEQUENCE_DEFS) + 1, column=0, columnspan=5, pady=6)
            self.load_sequences_form()
        if not HAS_AR_LIBS:
            tk.Label(frame, fg="#b91c1c", justify="right", text=(
                "تنبيه: مكتبتا arabic-reshaper و python-bidi غير مثبتتين، وبدونهما سيظهر العربي في ملفات PDF "
                "مفصول الحروف ومعكوساً.\nللتثبيت:  pip install arabic-reshaper python-bidi")).pack(padx=10, pady=5)
        self.load_company()

    def open_integrity_window(self):
        win = tk.Toplevel(self.root)
        win.title("فحص سلامة البيانات")
        win.geometry("1150x560")
        win.transient(self.root)
        head = tk.Label(win, text="", font=(UI_FONT, 12, "bold"))
        head.pack(anchor="e", padx=10, pady=8)
        holder = tk.Frame(win)
        holder.pack(fill="both", expand=True, padx=10)
        t = make_tree(holder, [("lvl", "النتيجة", 70, "center"), ("area", "المجال", 130, "center"),
                               ("desc", "الفحص", 430, "e"), ("n", "العدد", 60, "center"), ("ex", "أمثلة", 380, "e")], rtl=True)
        t.tag_configure("err", background="#fee2e2")
        t.tag_configure("warn", background="#ffedd5")
        t.tag_configure("ok", foreground="#15803d")

        def run():
            clear_tree(t)
            try:
                res = integrity_check()
            except Exception as e:
                handle_error(e)
                return
            for lvl, area, desc, n, ex in res:
                tag = {"خطأ": "err", "تنبيه": "warn"}.get(lvl, "ok")
                t.insert("", "end", values=("✖ " + lvl if lvl == "خطأ" else ("⚠ " + lvl if lvl == "تنبيه" else "✔ سليم"),
                                            area, desc, n or "", ex or ""), tags=(tag,))
            errs = sum(1 for r in res if r[0] == "خطأ")
            warns = sum(1 for r in res if r[0] == "تنبيه")
            head.config(text=f"نتيجة الفحص: {len(res)} فحصاً  |  أخطاء: {errs}  |  تنبيهات: {warns}  |  سليم: {len(res) - errs - warns}",
                        fg="#b91c1c" if errs else ("#b45309" if warns else "#15803d"))

        def repair(fn, label):
            if not messagebox.askyesno("إصلاح", f"{label}\nتُحفظ نسخة احتياطية تلقائياً قبل الإصلاح. متابعة؟", parent=win):
                return
            try:
                fn()
                self.refresh_all()
                run()
                messagebox.showinfo("تم", "تم الإصلاح وإعادة الفحص.", parent=win)
            except Exception as e:
                handle_error(e)

        bar = tk.Frame(win)
        bar.pack(fill="x", padx=10, pady=8)
        money_button(bar, "🔄 إعادة الفحص", UI["primary"], run, 9).pack(side="right", padx=3)
        money_button(bar, "🛠️ إعادة احتساب أرصدة وتكلفة المخزون", UI["teal"],
                     lambda: repair(repair_inventory, "سيُعاد بناء رصيد ومتوسط تكلفة كل صنف من حركاته."), 9).pack(side="right", padx=3)
        money_button(bar, "🛠️ إعادة احتساب المسدد وحالة الفواتير", "#475569",
                     lambda: repair(repair_invoice_paid, "سيُعاد احتساب المسدد لكل فاتورة من السندات المربوطة."), 9
                     ).pack(side="right", padx=3)
        run()

    def load_company(self):
        info = company_info()
        for key, e in self.comp_entries.items():
            if isinstance(e, ttk.Combobox):
                e.configure(values=[""] + rep_names(include=info.get(key) or None))
                e.set(info.get(key) or "")
                continue
            e.delete(0, tk.END)
            e.insert(0, info.get(key) or "")
        if hasattr(self, "seq_entries"):
            self.load_sequences_form()
        if hasattr(self, "admin_pwd_e"):
            self.admin_pwd_e.delete(0, tk.END)
            self.admin_pwd_e.insert(0, get_setting("admin_password", "2026"))

    def load_sequences_form(self):
        with db() as conn:
            for name, (pre, nxt, wid, prev) in self.seq_entries.items():
                prefix, n, width = conn.execute("SELECT prefix, next_value, width FROM sequences WHERE name=?", (name,)).fetchone()
                for e, v in ((pre, prefix), (nxt, str(n)), (wid, str(width))):
                    e.delete(0, tk.END)
                    e.insert(0, v)
                prev.config(text=peek_serial(conn, name)[0])

    def save_sequences(self):
        try:
            for name, (pre, nxt, wid, _) in self.seq_entries.items():
                n, w = parse_float(nxt.get()), parse_float(wid.get())
                if n is None or w is None:
                    raise UserError("الرقم التالي وعدد الأرقام يجب أن تكون أرقاماً.")
                update_sequence(name, pre.get(), int(n), int(w))
            self.load_sequences_form()
            messagebox.showinfo("نجاح", "تم حفظ إعدادات الترقيم التسلسلي.")
        except Exception as e:
            self.load_sequences_form() if not isinstance(e, UserError) else None
            handle_error(e)

    def save_company(self):
        try:
            save_company_info({k: e.get() for k, e in self.comp_entries.items()})
            self.load_company()
            messagebox.showinfo("نجاح", "تم حفظ بيانات المنشأة المصدرة للفواتير.")
        except Exception as e:
            handle_error(e)


    # ============================================================
    # دورة المبيعات والمشتريات: أمر ← تسليم/استلام ← فاتورة
    # ============================================================
    # ---------------------------------------------------------- الخصم اللاحق (إشعار خصم مورد / إشعار دائن عميل)
    def build_discount_tab(self, parent, kind):
        """تبويب متماثل للمبيعات والمشتريات: خصم يصل بعد اكتمال الفاتورة ودخول/تسليم البضاعة."""
        st = self.flow[kind]
        sales = kind == DISC_SALES
        st["disc_editing"] = None
        who = "العميل" if sales else "المورد"
        intro = ("إشعار دائن للعميل: خصم أو تسوية لاحقة بعد إصدار الفاتورة وتسليم البضاعة. يخفض رصيد العميل (ما عليه) "
                 "وصافي المبيعات وضريبة المخرجات، ولا يغير المخزون أو تكلفة المبيعات."
                 if sales else
                 "إشعار خصم من المورد: خصم مكتسب يصل بعد دخول البضاعة المخزن. يخفض رصيد المورد (ما له علينا) وضريبة المدخلات. "
                 "«تخفيض تكلفة المخزون» يوزّع الخصم على أصناف الفاتورة ويخفض متوسط تكلفتها (وما يخص بضاعة سبق بيعها يظهر "
                 "خصماً مكتسباً)، أما «خصم مباشر» فيُسجَّل خصماً مكتسباً بقائمة الدخل.")
        tk.Label(parent, text=intro, wraplength=980, justify="right", anchor="e", bg="#e0f2fe", fg="#0c4a6e",
                 font=("Arial", 9), padx=10, pady=6).pack(fill="x", padx=8, pady=(6, 2))
        f = ttk.LabelFrame(parent, text=f" {DISC_KIND_LABEL[kind]} ", padding=10)
        f.pack(fill="x", padx=8, pady=4)
        st["disc_party_cb"] = SearchableCombobox(f, width=32)
        st["disc_party_cb"].bind("<<ComboboxSelected>>", lambda e, k=kind: self._disc_party_changed(k), add="+")
        st["disc_party_cb"].bind("<FocusOut>", lambda e, k=kind: self._disc_party_changed(k), add="+")
        st["disc_date_e"] = date_entry(f)
        st["disc_inv_cb"] = SearchableCombobox(f, width=36)
        rtl_row(f, 0, [(f"{who}:", st["disc_party_cb"]), ("التاريخ:", st["disc_date_e"]),
                       ("الفاتورة (اختياري):", st["disc_inv_cb"])], slots=3)
        st["disc_amount_e"] = ttk.Entry(f, width=14)
        st["disc_vat_var"] = tk.BooleanVar(value=True)
        vat_chk = ttk.Checkbutton(f, text="تسوية ضريبة القيمة المضافة 15%", variable=st["disc_vat_var"])
        st["disc_reason_e"] = ttk.Entry(f, width=36)
        rtl_row(f, 1, [("الخصم قبل الضريبة:", st["disc_amount_e"]), ("", vat_chk),
                       ("سبب الخصم:", st["disc_reason_e"])], slots=3)
        row = 2
        if not sales:
            st["disc_method_cb"] = ttk.Combobox(f, values=list(DISC_METHOD_LABEL.values()), width=32, state="readonly")
            st["disc_method_cb"].set(DISC_METHOD_LABEL[DISC_STOCK])
            st["disc_item_cb"] = SearchableCombobox(f, width=26)
            rtl_row(f, row, [("المعالجة:", st["disc_method_cb"]), ("صنف محدد:", st["disc_item_cb"])], slots=3)
            row += 1
        st["disc_notes_e"] = ttk.Entry(f, width=32)
        rtl_row(f, row, [("ملاحظات:", st["disc_notes_e"])], slots=3)
        bar = tk.Frame(f)
        bar.grid(row=row + 1, column=0, columnspan=6, pady=8)
        st["disc_save_btn"] = money_button(bar, "💾 حفظ الإشعار وترحيله", "#16a34a", lambda k=kind: self.save_discount(k), 11)
        st["disc_save_btn"].pack(side="right", padx=4)
        money_button(bar, "✖ مسح / إلغاء التعديل", "#64748b", lambda k=kind: self.clear_discount_form(k), 9).pack(side="right", padx=4)
        if self.user["is_admin"]:
            money_button(bar, "⚙️ حسابات القيد الآلي", "#475569", self.open_disc_accounts_dialog, 9).pack(side="right", padx=4)

        box = ttk.LabelFrame(parent, text=" سجل الإشعارات ", padding=8)
        box.pack(fill="both", expand=True, padx=8, pady=4)
        adm = tk.Frame(box)
        adm.pack(fill="x", pady=(0, 4))
        money_button(adm, "🖨️ طباعة الإشعار PDF", UI["primary"], lambda k=kind: self.print_discount_selected(k), 9
                     ).pack(side="right", padx=3)
        if self.can_modify():
            money_button(adm, "✏️ تعديل الإشعار المحدد", "#d97706", lambda k=kind: self.edit_discount(k), 9).pack(side="right", padx=3)
            money_button(adm, "🗑️ حذف الإشعار المحدد", "#dc2626", lambda k=kind: self.delete_discount(k), 9).pack(side="right", padx=3)
        holder = tk.Frame(box)
        holder.pack(fill="both", expand=True)
        cols = [("no", "رقم الإشعار", 95, "center"), ("date", "التاريخ", 90, "center"), ("party", who, 200, "e"),
                ("inv", "الفاتورة", 95, "center")]
        if not sales:
            cols.append(("method", "المعالجة", 150, "center"))
        cols += [("sub", "قبل الضريبة", 95, "center"), ("vat", "الضريبة", 80, "center"), ("total", "الإجمالي", 95, "center")]
        if not sales:
            cols.append(("applied", "خُفّض من المخزون", 105, "center"))
        cols.append(("reason", "السبب", 170, "e"))
        st["disc_tree"] = make_tree(holder, cols, rtl=True)
        st["disc_tree"].bind("<Double-1>", lambda e, k=kind: self.show_transaction_details(
            ("discount", int(st["disc_tree"].selection()[0]))) if st["disc_tree"].selection() else None)
        st["disc_sum_lbl"] = tk.Label(box, text="", font=("Arial", 10, "bold"), fg="#0f172a")
        st["disc_sum_lbl"].pack(fill="x", pady=(4, 0))

    def refresh_discount_tab(self, kind):
        st = self.flow.get(kind) or {}
        if "disc_tree" not in st:
            return
        st["disc_party_cb"].set_items(self.partner_pairs("customer" if kind == DISC_SALES else "supplier"))
        if "disc_item_cb" in st:
            st["disc_item_cb"].set_items(self.item_pairs())
        self._disc_party_changed(kind)
        self.load_discounts(kind)

    def _disc_party_changed(self, kind):
        st = self.flow[kind]
        party = st["disc_party_cb"].selected()
        none_label = "بدون فاتورة محددة (على الحساب)"
        pairs = [(none_label, None)]
        if party:
            for iid, no, d, sub, tot, prev in discount_invoice_choices(kind, party[0]):
                pairs.append((f"{no} ({d}) - قبل الضريبة {fmt(sub)} - خصومات سابقة {fmt(prev)}", iid))
        current = st["disc_inv_cb"].get()
        st["disc_inv_cb"].set_items(pairs)
        if current not in dict(pairs):
            st["disc_inv_cb"].set(none_label)

    def clear_discount_form(self, kind):
        st = self.flow[kind]
        st["disc_editing"] = None
        for key in ("disc_amount_e", "disc_reason_e", "disc_notes_e"):
            st[key].delete(0, tk.END)
        st["disc_vat_var"].set(True)
        if "disc_item_cb" in st:
            st["disc_item_cb"].set("")
            st["disc_method_cb"].set(DISC_METHOD_LABEL[DISC_STOCK])
        st["disc_save_btn"].config(text="💾 حفظ الإشعار وترحيله")

    def save_discount(self, kind):
        st = self.flow[kind]
        try:
            party = st["disc_party_cb"].selected()
            if not party:
                raise UserError("اختر " + ("العميل" if kind == DISC_SALES else "المورد") + " من القائمة.")
            d = parse_date(st["disc_date_e"].get())
            if not d:
                raise UserError("صيغة التاريخ غير صحيحة، استخدم YYYY-MM-DD")
            amount = parse_float(st["disc_amount_e"].get())
            if amount is None or amount <= 0:
                raise UserError("أدخل قيمة خصم صحيحة أكبر من الصفر.")
            inv_id = st["disc_inv_cb"].selected()
            method, item = DISC_PL, None
            if "disc_method_cb" in st:
                method = DISC_STOCK if st["disc_method_cb"].get() == DISC_METHOD_LABEL[DISC_STOCK] else DISC_PL
                if st["disc_item_cb"].get().strip():
                    item = st["disc_item_cb"].selected()
                    if not item:
                        raise UserError("الصنف المكتوب غير واضح. اختره من القائمة أو اترك الحقل فارغاً.")
            self.check_period_open(d)
            editing = st["disc_editing"]
            _, no, total = save_post_discount(
                kind, d, party, amount, apply_vat=st["disc_vat_var"].get(), invoice_id=inv_id, method=method, item=item,
                reason=st["disc_reason_e"].get().strip(), notes=st["disc_notes_e"].get().strip(), discount_id=editing,
                user=self.user["username"], allow_closed=True)
            effect = "خُفّض رصيد العميل (ما عليه)" if kind == DISC_SALES else "خُفّض رصيد المورد (ما له علينا)"
            messagebox.showinfo("تم", f"تم {'تعديل' if editing else 'تسجيل'} الإشعار رقم {no} بإجمالي {fmt(total)} ر.س.\n"
                                      f"{effect} وتم ترحيل القيد اليومي الآلي.")
            self.clear_discount_form(kind)
            self.refresh_discount_tab(kind)
            self.load_inventory()
            if kind == DISC_SALES:
                self.load_sales_invoices()
            else:
                self.load_purchase_invoices()
            self._refresh_open_statement(party[0])
        except Exception as e:
            handle_error(e)

    def load_discounts(self, kind):
        st = self.flow[kind]
        tree = st["disc_tree"]
        clear_tree(tree)
        tot = sub_tot = 0.0
        rows = list_post_discounts(kind)
        for r in rows:
            vals = [r[1], r[3], r[5], r[6] or "-"]
            if kind == DISC_PURCHASE:
                vals.append("تخفيض التكلفة" if r[7] == DISC_STOCK else "خصم مكتسب مباشر")
            vals += [fmt(r[8]), fmt(r[9]), fmt(r[10])]
            if kind == DISC_PURCHASE:
                vals.append(fmt(r[11]) if r[7] == DISC_STOCK else "-")
            vals.append(r[12] or "-")
            tree.insert("", "end", iid=str(r[0]), values=vals)
            tot += r[10]
            sub_tot += r[8]
        st["disc_sum_lbl"].config(text=f"عدد الإشعارات: {len(rows)}   |   الإجمالي قبل الضريبة: {fmt(sub_tot)}   |   "
                                       f"الإجمالي شامل الضريبة: {fmt(tot)} ر.س")

    def _selected_discount(self, kind):
        sel = self.flow[kind]["disc_tree"].selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر إشعاراً من السجل.")
            return None
        return int(sel[0])

    def edit_discount(self, kind):
        if not self.require_modify():
            return
        did = self._selected_discount(kind)
        if did is None:
            return
        try:
            d, lines = load_post_discount(did)
        except Exception as e:
            handle_error(e)
            return
        st = self.flow[kind]
        st["disc_party_cb"].set(f"{d['party_code']} - {d['party_name']}")
        self._disc_party_changed(kind)
        for label, iid in st["disc_inv_cb"]._map.items():
            if iid == d["invoice_id"]:
                st["disc_inv_cb"].set(label)
                break
        for key, val in (("disc_date_e", d["discount_date"]), ("disc_amount_e", f"{d['subtotal']:g}"),
                         ("disc_reason_e", d["reason"]), ("disc_notes_e", d["notes"])):
            st[key].delete(0, tk.END)
            st[key].insert(0, val)
        st["disc_vat_var"].set(d["vat_amount"] > 0.005)
        if "disc_method_cb" in st:
            st["disc_method_cb"].set(DISC_METHOD_LABEL[d["method"]])
            st["disc_item_cb"].set("")
            if d["method"] == DISC_STOCK and not d["invoice_id"] and lines:
                st["disc_item_cb"].set(f"{lines[0][0]} - {lines[0][1]}")
        st["disc_editing"] = did
        st["disc_save_btn"].config(text=f"💾 حفظ تعديل الإشعار {d['discount_no']}")

    def delete_discount(self, kind):
        if not self.require_modify():
            return
        did = self._selected_discount(kind)
        if did is None:
            return
        if not messagebox.askyesno("حذف إشعار", "حذف الإشعار المحدد؟ سيُعاد رصيد الطرف وتكلفة الأصناف ويُحذف قيده الآلي."):
            return
        try:
            delete_post_discount(did, allow_closed=self.user["is_admin"])
            self.clear_discount_form(kind)
            self.refresh_discount_tab(kind)
            self.load_inventory()
        except Exception as e:
            handle_error(e)

    def open_disc_accounts_dialog(self):
        if not self.user["is_admin"]:
            return
        win = tk.Toplevel(self.root)
        win.title("حسابات القيد الآلي للخصم اللاحق")
        win.transient(self.root)
        f = ttk.Frame(win, padding=14)
        f.pack(fill="both", expand=True)
        with db() as conn:
            current = disc_accounts(conn)
            accounts = [(f"{c} - {n}", c) for c, n in conn.execute(
                "SELECT account_code, account_name FROM chart_of_accounts WHERE COALESCE(is_group,0)=0 ORDER BY account_code")]
        names = {c: label for label, c in accounts}
        boxes = {}
        for i, (key, (_, label)) in enumerate(DISC_ACCOUNT_DEFAULTS.items()):
            ttk.Label(f, text=f"{label}:").grid(row=i, column=1, sticky="e", padx=6, pady=4)
            cb = SearchableCombobox(f, width=46)
            cb.set_items(accounts)
            cb.set(names.get(current[key], current[key]))
            cb.grid(row=i, column=0, sticky="w", padx=6, pady=4)
            boxes[key] = cb

        def save():
            try:
                mapping = {}
                for key, cb in boxes.items():
                    code = cb.selected()
                    if not code:
                        raise UserError(f"اختر حساب «{DISC_ACCOUNT_DEFAULTS[key][1]}» من القائمة.")
                    mapping[key] = code
                set_disc_accounts(mapping)
                messagebox.showinfo("تم", "تم حفظ الحسابات وإعادة بناء قيود الخصم اللاحق.", parent=win)
                win.destroy()
                self.load_journal_register()
            except Exception as e:
                handle_error(e)

        money_button(f, "💾 حفظ وإعادة ترحيل القيود", "#16a34a", save, 10).grid(
            row=len(DISC_ACCOUNT_DEFAULTS), column=0, columnspan=2, pady=10)

    # ---------------------------------------------------------- مركز استيراد البيانات
    IMPORT_ORDER = ("items", "customers", "suppliers", "cash")
    IMPORT_ICONS = {"items": "📦", "customers": "👥", "suppliers": "🏭", "cash": "🏦"}

    def setup_import_tab(self, frame):
        self.imp = {}
        hero = tk.Frame(frame, bg=UI["slate2"])
        hero.pack(fill="x", padx=10, pady=(8, 4))
        tk.Label(hero, text="📥 مركز استيراد البيانات والترحيل", bg=UI["slate2"], fg="#ffffff",
                 font=(UI_FONT, 15, "bold")).pack(anchor="e", padx=18, pady=(12, 0))
        tk.Label(hero, text="١) نزّل النموذج واملأه   ٢) اختر الملف لمعاينة النتيجة بدون حفظ   ٣) نفّذ الاستيراد "
                            "(تُحفظ نسخة احتياطية تلقائياً)\nالأكواد الموحدة: الأصناف ELH-0001  •  العملاء C-0001  •  الموردون S-0001",
                 bg=UI["slate2"], fg="#a5f3fc", font=(UI_FONT, 9), wraplength=1100, justify="right").pack(
            anchor="e", padx=18, pady=(2, 12))
        nb = ttk.Notebook(frame)
        nb.pack(fill="both", expand=True, padx=10, pady=4)
        tabs = {k: ttk.Frame(nb) for k in self.IMPORT_ORDER}
        for k in reversed(self.IMPORT_ORDER):
            nb.add(tabs[k], text=f"  {self.IMPORT_ICONS[k]} {IMPORT_KINDS[k]['title']}  ")
        nb.select(tabs["items"])
        for k in self.IMPORT_ORDER:
            self._build_import_kind(tabs[k], k)

    def _build_import_kind(self, parent, kind):
        spec = IMPORT_KINDS[kind]
        st = self.imp[kind] = {"df": None}
        if kind == "cash":
            self._build_cash_opening_form(parent)
        card = ttk.LabelFrame(parent, text=f" {spec['title']} من ملف ", padding=10)
        card.pack(fill="x", padx=8, pady=6)
        tk.Label(card, text=spec["help"] + "\nأعمدة النموذج: " + " ، ".join(spec["columns"]), wraplength=1080,
                 justify="right", anchor="e", fg=UI["muted"], font=(UI_FONT, 9)).pack(fill="x")
        opts = tk.Frame(card)
        opts.pack(fill="x", pady=6)
        if kind == "items":
            st["upd"], st["recode"] = tk.BooleanVar(value=True), tk.BooleanVar(value=False)
            ttk.Checkbutton(opts, text="تحديث بيانات الأصناف المطابقة (الوحدة، سعر البيع، الباركود، الرصيد الافتتاحي)",
                            variable=st["upd"]).pack(side="right", padx=6)
            ttk.Checkbutton(opts, text="توحيد أكواد الأصناف المطابقة مع حفظ الكود القديم كمرجع: ELH-0001",
                            variable=st["recode"]).pack(side="right", padx=6)
        else:
            tk.Label(opts, text="تاريخ الرصيد الافتتاحي:").pack(side="right", padx=4)
            st["asof"] = date_entry(opts)
            st["asof"].pack(side="right")
        bar = tk.Frame(card)
        bar.pack(fill="x", pady=4)
        money_button(bar, "⬇️ نموذج Excel", "#475569", lambda k=kind: self.save_import_template(k, ".xlsx"), 9).pack(side="right", padx=3)
        money_button(bar, "⬇️ نموذج CSV", "#475569", lambda k=kind: self.save_import_template(k, ".csv"), 9).pack(side="right", padx=3)
        money_button(bar, "📂 اختيار ملف ومعاينة", UI["primary"], lambda k=kind: self.preview_import(k), 10).pack(side="right", padx=10)
        st["run_btn"] = money_button(bar, "✅ تنفيذ الاستيراد والحفظ", UI["teal"], lambda k=kind: self.execute_import(k), 10)
        st["run_btn"].pack(side="right", padx=3)
        st["run_btn"].configure(state="disabled")
        st["file_lbl"] = tk.Label(bar, text="لم يُختر ملف بعد", fg=UI["muted"])
        st["file_lbl"].pack(side="right", padx=10)
        res = ttk.LabelFrame(parent, text=" النتيجة ", padding=8)
        res.pack(fill="both", expand=True, padx=8, pady=4)
        st["stats_lbl"] = tk.Label(res, text="", font=(UI_FONT, 10, "bold"), fg=UI["primary_dk"], anchor="e",
                                   justify="right", wraplength=1080)
        st["stats_lbl"].pack(fill="x")
        st["prob_lbl"] = tk.Label(res, text="", fg="#dc2626", anchor="e", justify="right", wraplength=1080)
        st["prob_lbl"].pack(fill="x")
        holder = tk.Frame(res)
        holder.pack(fill="both", expand=True)
        cols = [(f"c{i}", title or "-", 70 if i == 0 else (240 if i == 3 else 130), "e" if i == 3 else "center")
                for i, title in enumerate(import_preview_columns(kind))]
        st["tree"] = make_tree(holder, cols, height=9, rtl=True)

    def _build_cash_opening_form(self, parent):
        card = ttk.LabelFrame(parent, text=" ⚡ إعداد سريع: أرصدة الخزينة والبنوك أول المدة ", padding=10)
        card.pack(fill="x", padx=8, pady=6)
        top = tk.Frame(card)
        top.pack(fill="x")
        tk.Label(top, text="تاريخ الأرصدة:").pack(side="right", padx=4)
        self.imp_cash_date = date_entry(top)
        self.imp_cash_date.pack(side="right")
        money_button(top, "💾 حفظ أرصدة الخزينة والبنوك", UI["teal"], self.save_cash_opening_inline, 10).pack(side="left", padx=4)
        self.imp_cash_grid = tk.Frame(card)
        self.imp_cash_grid.pack(fill="x", pady=6)
        self.imp_cash_entries = {}

    def refresh_import_page(self):
        if not hasattr(self, "imp_cash_grid"):
            return
        for w in self.imp_cash_grid.winfo_children():
            w.destroy()
        self.imp_cash_entries = {}
        with db() as conn:
            names = cash_account_names(conn)
            codes = cash_accounts(conn)
        as_of, amounts = get_cash_opening()
        if as_of:
            self.imp_cash_date.delete(0, tk.END)
            self.imp_cash_date.insert(0, as_of)
        for i, name in enumerate(names):
            r, c = divmod(i, 3)
            cell = tk.Frame(self.imp_cash_grid, bg=UI["card"], highlightbackground=UI["border"], highlightthickness=1)
            cell.grid(row=r, column=2 - c, sticky="ew", padx=4, pady=3)
            tk.Label(cell, text=f"{name} ({codes.get(name, '-')})", bg=UI["card"], fg=UI["slate3"],
                     font=(UI_FONT, 9, "bold"), anchor="e").pack(side="right", padx=6, pady=4)
            e = ttk.Entry(cell, width=14, justify="center")
            e.insert(0, f"{amounts.get(name, 0.0):g}")
            e.pack(side="left", padx=6, pady=4)
            self.imp_cash_entries[name] = e
        for c in range(3):
            self.imp_cash_grid.grid_columnconfigure(c, weight=1)

    def save_cash_opening_inline(self):
        try:
            as_of = parse_date(self.imp_cash_date.get())
            if not as_of:
                raise UserError("صيغة التاريخ غير صحيحة، استخدم YYYY-MM-DD")
            amounts = {}
            for name, e in self.imp_cash_entries.items():
                v = parse_float(e.get() or "0")
                if v is None:
                    raise UserError(f"قيمة غير صحيحة لحساب {name}.")
                amounts[name] = v
            if not messagebox.askyesno("تأكيد", f"حفظ الأرصدة الافتتاحية للخزينة والبنوك بتاريخ {as_of}؟\n"
                                                 f"الإجمالي: {fmt(sum(amounts.values()))}\n(يستبدل قيد الأرصدة الافتتاحية السابق)"):
                return
            set_cash_opening(as_of, amounts)
            messagebox.showinfo("تم", "تم حفظ الأرصدة الافتتاحية للخزينة والبنوك.")
            self.refresh_import_page()
            self.refresh_cash_accounts()
        except Exception as e:
            handle_error(e)

    def save_import_template(self, kind, ext):
        path = filedialog.asksaveasfilename(defaultextension=ext, initialfile=f"نموذج_{IMPORT_KINDS[kind]['title']}{ext}",
                                            filetypes=[("Excel", "*.xlsx")] if ext == ".xlsx" else [("CSV", "*.csv")])
        if not path:
            return
        try:
            write_import_template(path, kind)
            messagebox.showinfo("تم", f"تم حفظ النموذج:\n{path}")
        except Exception as e:
            handle_error(e)

    def _import_opts(self, kind):
        st = self.imp[kind]
        if kind == "items":
            return dict(update_existing=st["upd"].get(), recode_matched=st["recode"].get())
        as_of = parse_date(st["asof"].get())
        if not as_of:
            raise UserError("تاريخ الرصيد الافتتاحي غير صحيح (YYYY-MM-DD).")
        return dict(as_of=as_of)

    def _show_import_result(self, kind, res, done=False):
        st = self.imp[kind]
        head = "✅ تم الاستيراد: " if done else "🔎 معاينة (لم يُحفظ شيء بعد): "
        st["stats_lbl"].config(text=head + import_stats_text(res["stats"]), fg=UI["teal"] if done else UI["primary_dk"])
        probs = res["problems"]
        st["prob_lbl"].config(text=("⚠️ " + "\n⚠️ ".join(probs[:6]) + (f"\n... و{len(probs) - 6} ملاحظة أخرى" if len(probs) > 6 else ""))
                              if probs else "")
        clear_tree(st["tree"])
        for r in res["preview"][:1000]:
            st["tree"].insert("", "end", values=[fmt(v) if isinstance(v, float) else v for v in r])

    def preview_import(self, kind):
        path = filedialog.askopenfilename(title=f"اختر ملف {IMPORT_KINDS[kind]['title']}",
                                          filetypes=[("Excel / CSV", "*.xlsx *.xls *.csv"), ("كل الملفات", "*.*")])
        if not path:
            return
        st = self.imp[kind]
        try:
            df = read_table_file(path)
            res = run_import(kind, df, dry_run=True, **self._import_opts(kind))
        except Exception as e:
            handle_error(e)
            return
        st.update(df=df, path=path)
        st["file_lbl"].config(text=f"📄 {os.path.basename(path)} ({len(df)} صف)")
        self._show_import_result(kind, res)
        st["run_btn"].configure(state="normal")

    def execute_import(self, kind):
        st = self.imp[kind]
        if st.get("df") is None:
            messagebox.showinfo("تنبيه", "اختر ملفاً واعرض المعاينة أولاً.")
            return
        if not messagebox.askyesno("تأكيد الاستيراد", f"تنفيذ {IMPORT_KINDS[kind]['title']} من الملف:\n"
                                                      f"{os.path.basename(st['path'])}؟\nستُحفظ نسخة احتياطية قبل التنفيذ."):
            return
        try:
            backup_database("before_import")
            res = run_import(kind, st["df"], dry_run=False, **self._import_opts(kind))
        except Exception as e:
            handle_error(e)
            return
        self._show_import_result(kind, res, done=True)
        st["run_btn"].configure(state="disabled")
        st["df"] = None
        messagebox.showinfo("تم", "تم الاستيراد بنجاح.\n" + import_stats_text(res["stats"]).replace("   |   ", "\n"))
        for fn in (self.load_inventory, self.load_entities, self.refresh_import_page, self.refresh_cash_accounts):
            try:
                fn()
            except Exception:
                pass

    # ---------------------------------------------------------- الحسابات: سجل الخصومات اللاحقة وأثرها
    def setup_discounts_tab(self, frame):
        hero = tk.Frame(frame, bg=UI["slate2"])
        hero.pack(fill="x", padx=10, pady=(8, 4))
        tk.Label(hero, text="🏷️ سجل الخصومات اللاحقة والإشعارات الدائنة", bg=UI["slate2"], fg="#ffffff",
                 font=(UI_FONT, 15, "bold")).pack(anchor="e", padx=18, pady=(12, 0))
        tk.Label(hero, text="كل خصم لاحق من الموردين وإشعار دائن للعملاء مع أثره على رصيد الطرف والمخزون والأرباح والضريبة.",
                 bg=UI["slate2"], fg="#a5f3fc", font=(UI_FONT, 9)).pack(anchor="e", padx=18, pady=(2, 12))
        bar = tk.Frame(frame)
        bar.pack(fill="x", padx=10, pady=4)
        self.dr_type = ttk.Combobox(bar, values=["الكل", "خصومات الموردين", "إشعارات العملاء"], width=16, state="readonly")
        self.dr_type.set("الكل")
        self.dr_from = date_entry(bar)
        self.dr_from.delete(0, tk.END)
        self.dr_from.insert(0, f"{date.today().year}-01-01")
        self.dr_to = date_entry(bar)
        for w, lbl in ((self.dr_type, "النوع:"), (self.dr_from, "من:"), (self.dr_to, "إلى:")):
            tk.Label(bar, text=lbl).pack(side="right", padx=(8, 2))
            w.pack(side="right")
        money_button(bar, "🔍 عرض", UI["primary"], self.load_discounts_page, 9).pack(side="right", padx=8)
        if self.can("purchase.discount"):
            money_button(bar, "➕ خصم لاحق من مورد", UI["teal"], lambda: self.open_discount_for_invoice("purchase", None),
                         9).pack(side="left", padx=3)
        if self.can("sales.discount"):
            money_button(bar, "➕ إشعار دائن لعميل", UI["teal"], lambda: self.open_discount_for_invoice("sales", None),
                         9).pack(side="left", padx=3)
        money_button(bar, "🖨️ طباعة المحدد PDF", "#475569", self.print_selected_register_discount, 9).pack(side="left", padx=3)
        money_button(bar, "📊 تصدير Excel", "#475569", self.export_discount_register, 9).pack(side="left", padx=3)

        cards = tk.Frame(frame)
        cards.pack(fill="x", padx=10, pady=6)
        self.dr_cards = {}
        for i, (key, title, color) in enumerate((
                ("count", "عدد الإشعارات", UI["primary"]), ("sup", "خصومات الموردين (شامل الضريبة)", UI["teal"]),
                ("cus", "إشعارات العملاء (شامل الضريبة)", UI["cyan"]), ("stock", "خفض تكلفة المخزون", "#7c3aed"),
                ("pl", "صافي الأثر على الأرباح", "#0891b2"), ("vat", "صافي الأثر على الضريبة المستحقة", "#d97706"))):
            card = tk.Frame(cards, bg=UI["card"], highlightbackground=UI["border"], highlightthickness=1)
            card.grid(row=0, column=5 - i, sticky="nsew", padx=4)
            tk.Frame(card, bg=color, height=4).pack(fill="x")
            tk.Label(card, text=title, bg=UI["card"], fg=UI["muted"], font=(UI_FONT, 9, "bold")).pack(anchor="e", padx=10, pady=(6, 0))
            v = tk.Label(card, text="-", bg=UI["card"], fg=UI["ink"], font=(UI_FONT, 14, "bold"))
            v.pack(anchor="e", padx=10, pady=(0, 8))
            self.dr_cards[key] = v
            cards.grid_columnconfigure(5 - i, weight=1)

        box = ttk.LabelFrame(frame, text=" الإشعارات (اضغط مرتين لعرض التفاصيل وتوزيع الخصم على الأصناف) ", padding=6)
        box.pack(fill="both", expand=True, padx=10, pady=4)
        holder = tk.Frame(box)
        holder.pack(fill="both", expand=True)
        self.dr_tree = make_tree(holder, [
            ("no", "رقم الإشعار", 88, "center"), ("date", "التاريخ", 82, "center"), ("type", "النوع", 50, "center"),
            ("party", "الطرف", 165, "e"), ("inv", "الفاتورة", 82, "center"), ("method", "المعالجة", 140, "center"),
            ("sub", "قبل الضريبة", 82, "center"), ("vat", "الضريبة", 68, "center"), ("total", "الإجمالي", 82, "center"),
            ("stock", "خفض المخزون", 88, "center"), ("pl", "أثر الأرباح", 80, "center"), ("je", "القيد", 45, "center")],
            rtl=True)
        self.dr_tree.bind("<Double-1>", lambda e: self.show_transaction_details(("discount", int(self.dr_tree.selection()[0])))
                          if self.dr_tree.selection() else None)
        tk.Label(frame, text="أثر الأرباح: خصم المورد المباشر (+ خصم مكتسب) وما يخص بضاعة مباعة، والإشعار الدائن للعميل (- يخفض صافي "
                             "المبيعات). خفض المخزون: الجزء الذي خفّض متوسط تكلفة الرصيد المتبقي (يظهر في تكلفة المبيعات لاحقاً).",
                 fg=UI["muted"], font=(UI_FONT, 8), wraplength=1100, justify="right").pack(anchor="e", padx=12, pady=(0, 6))

    def _discount_register_filter(self):
        d1, d2 = parse_date(self.dr_from.get()), parse_date(self.dr_to.get())
        if not d1 or not d2:
            raise UserError("صيغة التاريخ YYYY-MM-DD")
        kind = {"خصومات الموردين": DISC_PURCHASE, "إشعارات العملاء": DISC_SALES}.get(self.dr_type.get())
        return kind, d1, d2

    def load_discounts_page(self):
        if not hasattr(self, "dr_tree"):
            return
        try:
            kind, d1, d2 = self._discount_register_filter()
            data = post_discount_register(kind, d1, d2)
        except Exception as e:
            handle_error(e)
            return
        clear_tree(self.dr_tree)
        for r in data:
            self.dr_tree.insert("", "end", iid=str(r["id"]), values=(
                r["no"], r["date"], "مورد" if r["kind"] == DISC_PURCHASE else "عميل", f"{r['party_code']} - {r['party']}",
                r["invoice"] or "-", r["method"], fmt(r["subtotal"]), fmt(r["vat"]), fmt(r["total"]),
                fmt(r["stock"]) if r["stock"] else "-", fmt(r["pl"]), r["journal"] or "-"))
        sup = [r for r in data if r["kind"] == DISC_PURCHASE]
        cus = [r for r in data if r["kind"] == DISC_SALES]
        vat_effect = sum(r["vat"] for r in sup) - sum(r["vat"] for r in cus)  # المدخلات تقل (+مستحق) والمخرجات تقل (-مستحق)
        for key, val in (("count", f"{len(data)}"), ("sup", fmt(sum(r["total"] for r in sup))),
                         ("cus", fmt(sum(r["total"] for r in cus))), ("stock", fmt(sum(r["stock"] for r in data))),
                         ("pl", fmt(sum(r["pl"] for r in data))), ("vat", fmt(vat_effect))):
            self.dr_cards[key].config(text=val)

    def print_selected_register_discount(self):
        sel = self.dr_tree.selection() if hasattr(self, "dr_tree") else ()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر إشعاراً من السجل.")
            return
        self.export_document_pdf(build_post_discount_pdf, f"إشعار_{self.dr_tree.item(sel[0])['values'][0]}", int(sel[0]))

    def export_discount_register(self):
        try:
            kind, d1, d2 = self._discount_register_filter()
            rep = run_report("pdisc", dict(d_from=d1, d_to=d2))
            if kind:
                key = "مورد" if kind == DISC_PURCHASE else "عميل"
                rep.rows = [r for r in rep.rows if r[2] == key]
            path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel", "*.xlsx")],
                                                initialfile=f"الخصومات_اللاحقة_{d1}_{d2}.xlsx")
            if path:
                export_report_xlsx(path, rep)
                messagebox.showinfo("تم", f"تم التصدير:\n{path}")
        except Exception as e:
            handle_error(e)

    # ---------------------------------------------------------- طباعة المستندات والانتقال للخصم اللاحق
    def export_document_pdf(self, builder, name, *args):
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF", "*.pdf")], title="حفظ المستند",
                                            initialfile=f"{name}.pdf")
        if not path:
            return
        try:
            builder(path, *args)
            if messagebox.askyesno("تم", "تم إنشاء المستند بنجاح. هل تريد فتحه الآن؟"):
                try:
                    os.startfile(path)  # ويندوز
                except (AttributeError, OSError):
                    pass
        except Exception as e:
            handle_error(e)

    def print_invoice_selected(self, kind):
        inv_id = self._selected_invoice(kind)
        if inv_id is not None:
            self.export_document_pdf(build_invoice_pdf, f"فاتورة_{invoice_number(kind, inv_id)}", kind, inv_id)

    def print_discount_selected(self, kind):
        did = self._selected_discount(kind)
        if did is not None:
            no = self.flow[kind]["disc_tree"].item(str(did))["values"][0]
            self.export_document_pdf(build_post_discount_pdf, f"إشعار_{no}", did)

    def open_discount_for_invoice(self, kind, inv_id="selected"):
        """فتح تبويب «خصم لاحق / إشعار دائن» داخل المبيعات/المشتريات، مع اختيار الطرف والفاتورة المحددة تلقائياً."""
        if not self.can(f"{kind}.discount"):
            messagebox.showwarning("صلاحيات", "ليس لديك صلاحية تسجيل الخصم اللاحق في هذه الصفحة.")
            return
        if inv_id == "selected":
            inv_id = self._selected_invoice(kind)
            if inv_id is None:
                return
        party = None
        if inv_id:
            c = FLOW[kind]
            with db() as conn:
                party = conn.execute(f"SELECT {c['pcode']}, {c['pname']} FROM {c['inv']} WHERE {c['iid']}=?",
                                     (inv_id,)).fetchone()
        self.show_page(kind)
        self.flow_select(kind, "x")
        self.refresh_discount_tab(kind)
        self.clear_discount_form(kind)
        st = self.flow[kind]
        if party:
            label = next((lbl for lbl, v in st["disc_party_cb"]._map.items() if v and v[0] == party[0]),
                         f"{party[0]} - {party[1]}")
            st["disc_party_cb"].set(label)
            self._disc_party_changed(kind)
            for label, iid in st["disc_inv_cb"]._map.items():
                if iid == inv_id:
                    st["disc_inv_cb"].set(label)
                    break
        st["disc_amount_e"].focus_set()

    # ---------------------------------------------------------- العملاء / الموردون (صفحتان مستقلتان)
    def setup_customers_tab(self, frame):
        self._setup_partner_page(frame, "customer")

    def setup_suppliers_tab(self, frame):
        self._setup_partner_page(frame, "supplier")

    # ---------------------------------------------------------- سجل المناديب
    def setup_reps_tab(self, frame):
        st = self.rp = {"editing": None}
        f = ttk.LabelFrame(frame, text=" 🧑‍💼 بيانات المندوب ",
                           padding=10)
        f.pack(fill="x", padx=10, pady=8)
        st["code_lbl"] = tk.Label(f, text="", fg=UI["primary_dk"], font=(UI_FONT, 10, "bold"), wraplength=700, justify="right")
        st["code_lbl"].grid(row=0, column=0, columnspan=6, sticky="e", pady=(0, 4))
        st["name"] = ttk.Entry(f, width=26)
        st["phone"] = ttk.Entry(f, width=16)
        st["email"] = ttk.Entry(f, width=22)
        st["national_id"] = ttk.Entry(f, width=16)
        st["region"] = ttk.Entry(f, width=22)
        st["job_title"] = ttk.Combobox(f, width=20, values=["مندوب مبيعات", "مندوب مشتريات", "مندوب مبيعات ومشتريات",
                                                             "مشرف مبيعات", "مسؤول مشتريات"])
        st["hire_date"] = DateEntry(f, width=12)
        st["commission_pct"] = ttk.Entry(f, width=8)
        st["notes"] = ttk.Entry(f, width=22)
        st["active"] = tk.BooleanVar(value=True)
        rtl_row(f, 1, [("اسم المندوب *:", st["name"]), ("رقم الهاتف:", st["phone"]), ("البريد الإلكتروني:", st["email"])],
                slots=3)
        rtl_row(f, 2, [("رقم الهوية / الإقامة:", st["national_id"]), ("المنطقة / خط السير:", st["region"]),
                       ("المسمى الوظيفي:", st["job_title"])], slots=3)
        rtl_row(f, 3, [("تاريخ التعيين:", st["hire_date"]), ("نسبة العمولة %:", st["commission_pct"]),
                       ("ملاحظات:", st["notes"])], slots=3)
        ttk.Checkbutton(f, text="نشط (يظهر في القوائم المنسدلة)", variable=st["active"]).grid(row=4, column=5, sticky="e", pady=4)
        bar = tk.Frame(f)
        bar.grid(row=5, column=0, columnspan=8, pady=(8, 0))
        st["save_btn"] = money_button(bar, "حفظ وإضافة المندوب ➕", "#16a34a", self.save_rep_form, 10)
        st["save_btn"].pack(side="right", padx=3)
        money_button(bar, "🆕 جديد", "#475569", self.clear_rep_form, 9).pack(side="right", padx=3)
        money_button(bar, "🗑️ حذف / إيقاف", "#dc2626", self.delete_rep_selected, 9).pack(side="right", padx=3)
        money_button(bar, "🔗 ربط العملاء/الموردين بمناديب فواتيرهم", UI["teal"], self.fill_partner_reps_from_invoices, 9
                     ).pack(side="left", padx=3)
        money_button(bar, "👥 عملاء وموردو المندوب", UI["teal"], self.show_rep_partners, 9).pack(side="right", padx=3)

        lst = ttk.LabelFrame(frame, text=" سجل المناديب (اضغط مرتين للتعديل) ", padding=8)
        lst.pack(fill="both", expand=True, padx=10, pady=5)
        sf = tk.Frame(lst)
        sf.pack(fill="x", pady=(0, 4))
        st["q"] = ttk.Entry(sf, width=30)
        st["q"].bind("<KeyRelease>", lambda e: self.load_reps())
        ttk.Label(sf, text="بحث (الاسم / الهاتف / المنطقة / الكود):").pack(side="right", padx=5)
        st["q"].pack(side="right")
        st["count_lbl"] = tk.Label(sf, text="", fg=UI["muted"])
        st["count_lbl"].pack(side="left", padx=6)
        holder = tk.Frame(lst)
        holder.pack(fill="both", expand=True)
        st["tree"] = make_tree(holder, [
            ("code", "الكود", 60, "center"), ("name", "اسم المندوب", 170, "e"), ("phone", "الهاتف", 100, "center"),
            ("email", "البريد", 140, "center"), ("region", "المنطقة / خط السير", 140, "e"), ("job", "المسمى", 110, "center"),
            ("comm", "العمولة %", 65, "center"), ("n", "عدد الأطراف", 75, "center"),
            ("hire", "تاريخ التعيين", 85, "center"), ("active", "الحالة", 60, "center")], rtl=True)
        st["tree"].tag_configure("off", foreground="#94a3b8")
        st["tree"].bind("<Double-1>", lambda e: self.edit_rep_selected())
        self.clear_rep_form()

    def _rep_set(self, w, v):
        if isinstance(w, ttk.Combobox):
            w.set(v or "")
        else:
            w.delete(0, tk.END)
            w.insert(0, v or "")

    def clear_rep_form(self):
        st = self.rp
        st["editing"] = None
        for k in ("name", "phone", "email", "national_id", "region", "job_title", "hire_date", "commission_pct", "notes"):
            self._rep_set(st[k], "")
        st["active"].set(True)
        st["save_btn"].config(text="حفظ وإضافة المندوب ➕")
        st["code_lbl"].config(text="➕ مندوب جديد (الكود تلقائي) — النشطون يظهرون في قوائم المبيعات والمشتريات والعملاء والموردين")

    def load_reps(self):
        st = getattr(self, "rp", None)
        if not st:
            return
        clear_tree(st["tree"])
        rows = list_reps(q=st["q"].get())
        for r in rows:
            st["tree"].insert("", "end", iid=str(r[0]), tags=() if r[10] else ("off",), values=(
                r[1], r[2], r[3] or "-", r[4] or "-", r[6] or "-", r[7] or "-", f"{r[9]:g}" if r[9] else "-", r[12],
                r[8] or "-", "نشط" if r[10] else "موقوف"))
        st["count_lbl"].config(text=f"عدد المناديب: {len(rows)}  |  النشطون: {sum(1 for r in rows if r[10])}")

    def edit_rep_selected(self):
        st = self.rp
        sel = st["tree"].selection()
        if not sel:
            return
        r = next((x for x in list_reps() if x[0] == int(sel[0])), None)
        if not r:
            return
        self.clear_rep_form()
        st["editing"] = r[0]
        for k, v in zip(("name", "phone", "email", "national_id", "region", "job_title", "hire_date"), r[2:9]):
            self._rep_set(st[k], v)
        self._rep_set(st["commission_pct"], f"{r[9]:g}" if r[9] else "")
        self._rep_set(st["notes"], r[11])
        st["active"].set(bool(r[10]))
        st["save_btn"].config(text="💾 حفظ تعديلات المندوب")
        st["code_lbl"].config(text=f"✏️ تعديل المندوب {r[1]} — تغيير الاسم ينتقل تلقائياً للفواتير والأوامر والأطراف المرتبطة")

    def save_rep_form(self):
        st = self.rp
        try:
            data = {k: st[k].get() for k in ("name", "phone", "email", "national_id", "region", "job_title", "hire_date",
                                             "commission_pct", "notes")}
            data["active"] = st["active"].get()
            editing = st["editing"]
            rid = save_rep(editing, data, self.user["username"])
            self.clear_rep_form()
            self.load_reps()
            self.refresh_rep_dropdowns()
            if st["tree"].exists(str(rid)):
                st["tree"].selection_set(str(rid))
                st["tree"].see(str(rid))
            messagebox.showinfo("تم", "تم حفظ بيانات المندوب." if editing else "تمت إضافة المندوب إلى السجل.")
        except Exception as e:
            handle_error(e)

    def delete_rep_selected(self):
        st = self.rp
        sel = st["tree"].selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر مندوباً من السجل.")
            return
        name = st["tree"].item(sel[0])["values"][1]
        if not messagebox.askyesno("حذف مندوب", f"حذف المندوب «{name}»؟\nإذا كان مرتبطاً بفواتير أو عملاء سيتم إيقافه فقط "
                                                "(يختفي من القوائم ويبقى في التقارير والسجلات القديمة)."):
            return
        try:
            res = delete_rep(int(sel[0]))
            self.clear_rep_form()
            self.load_reps()
            self.refresh_rep_dropdowns()
            messagebox.showinfo("تم", "تم حذف المندوب." if res == "deleted" else
                                "المندوب مرتبط بمستندات سابقة، لذلك تم إيقافه بدلاً من حذفه.")
        except Exception as e:
            handle_error(e)

    def show_rep_partners(self):
        st = self.rp
        sel = st["tree"].selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر مندوباً من السجل.")
            return
        name = st["tree"].item(sel[0])["values"][1]
        win = tk.Toplevel(self.root)
        win.title(f"عملاء وموردو المندوب: {name}")
        win.geometry("760x420")
        win.transient(self.root)
        tk.Label(win, text=f"🧑‍💼 {name} — الأطراف المسندة إليه حالياً", font=(UI_FONT, 11, "bold"),
                 fg=UI["primary_dk"]).pack(anchor="e", padx=10, pady=8)
        holder = tk.Frame(win)
        holder.pack(fill="both", expand=True, padx=10, pady=(0, 10))
        t = make_tree(holder, [("code", "الكود", 80, "center"), ("name", "الاسم", 260, "e"), ("role", "الصفة", 100, "center"),
                               ("phone", "الهاتف", 120, "center"), ("since", "منذ", 100, "center")], rtl=True)
        with db() as conn:
            for code, pname, is_c, is_s, phone in conn.execute(
                    "SELECT partner_code, name, is_customer, is_supplier, COALESCE(phone,'') FROM business_partners "
                    "WHERE rep_name=? ORDER BY name", (name,)):
                since = conn.execute("SELECT start_date FROM partner_rep_history WHERE partner_code=? AND rep_name=? "
                                     "AND end_date IS NULL ORDER BY start_date DESC LIMIT 1", (code, name)).fetchone()
                role = " و".join(x for x, f_ in (("عميل", is_c), ("مورد", is_s)) if f_) or "-"
                t.insert("", "end", values=(code, pname, role, phone or "-", since[0] if since else "-"))

    def refresh_rep_dropdowns(self):
        """تحديث كل القوائم المنسدلة للمناديب بعد أي تعديل في السجل."""
        names = rep_names()
        for role in ("customer", "supplier"):
            st = getattr(self, "pp", {}).get(role)
            if st and "rep" in st:
                cur = st["rep"].get()
                st["rep"].configure(values=[""] + rep_names(include=cur or None))
        for kind in ("sales", "purchase"):
            st = getattr(self, "flow", {}).get(kind)
            if st and "o_rep" in st:
                cur = st["o_rep"].get()
                st["o_rep"].configure(values=[""] + (names if not cur or cur in names else names + [cur]))
        e = getattr(self, "comp_entries", {}).get("rep_name")
        if isinstance(e, ttk.Combobox):
            e.configure(values=[""] + names)

    def _setup_partner_page(self, frame, role):
        if not hasattr(self, "pp"):
            self.pp = {}
        st = self.pp[role] = {"editing": None}
        who = "العميل" if role == "customer" else "المورد"
        f = ttk.LabelFrame(frame, text=f" إضافة / تعديل {who} (الكود يُولَّد تلقائياً) ", padding=10)
        f.pack(fill="x", padx=10, pady=6)
        for key, w in (("name", 30), ("phone", 16), ("email", 24), ("address", 30), ("cr", 16), ("vat", 18)):
            st[key] = ttk.Entry(f, width=w)
        rtl_row(f, 0, [(f"اسم {who}:", st["name"]), ("الهاتف:", st["phone"]), ("البريد:", st["email"])], slots=3)
        rtl_row(f, 1, [("العنوان:", st["address"]), ("السجل التجاري:", st["cr"]), ("الرقم الضريبي:", st["vat"])], slots=3)
        st["rep"] = ttk.Combobox(f, width=22, values=[""] + rep_names(), state="readonly")
        st["rep_date"] = date_entry(f)
        st["rep_hist_btn"] = money_button(f, "🕘 سجل المناديب", "#475569", lambda r=role: self.show_rep_history(r), 9)
        rtl_row(f, 2, [("المندوب (افتراضي):", st["rep"]), ("تاريخ سريان المندوب:", st["rep_date"]),
                       (None, st["rep_hist_btn"])], slots=3)
        st["od"], st["oc"], st["odate"] = ttk.Entry(f, width=14), ttk.Entry(f, width=14), DateEntry(f, width=12)
        rtl_row(f, 3, [("رصيد افتتاحي مدين (عليه):", st["od"]), ("رصيد افتتاحي دائن (له):", st["oc"]),
                       ("تاريخ الرصيد:", st["odate"])], slots=3)
        if not self.can_modify():
            for k in ("od", "oc", "odate"):
                st[k].state(["disabled"])
        # الحد الائتماني وشروط السداد (مدة السداد تُستخدم في توقع التدفق النقدي للعملاء والموردين)
        st["cr_days"] = ttk.Entry(f, width=8)
        if role == "customer":
            st["cr_limit"] = ttk.Entry(f, width=14)
            st["cr_policy"] = ttk.Combobox(f, width=30, state="readonly",
                                           values=["(الافتراضية)"] + list(_credit.CREDIT_POLICIES.values()))
            rtl_row(f, 4, [("الحد الائتماني (0 = بدون):", st["cr_limit"]), ("مدة السداد (يوم، فارغ = الافتراضية):", st["cr_days"]),
                           ("سياسة التجاوز:", st["cr_policy"])], slots=3)
        else:
            rtl_row(f, 4, [("مدة السداد للمورد (يوم، فارغ = الافتراضية):", st["cr_days"])], slots=3)
        st["code_lbl"] = ttk.Label(f, text="", foreground=UI["primary_dk"], font=(UI_FONT, 10, "bold"))
        st["code_lbl"].grid(row=5, column=0, columnspan=6, sticky="e", padx=10)
        btns = tk.Frame(f)
        btns.grid(row=6, column=0, columnspan=6, pady=(6, 2))
        st["save_btn"] = money_button(btns, f"حفظ وإضافة {who} ➕", "#16a34a", lambda r=role: self._pp_save(r))
        st["save_btn"].pack(side="right", padx=5)
        money_button(btns, "مسح الحقول ✖", "#64748b", lambda r=role: self._pp_clear(r)).pack(side="right", padx=5)
        money_button(btns, "📜 الكشف الموحد للمحدد", UI["primary"], lambda r=role: self._pp_statement(r), 9).pack(side="right", padx=5)
        if role == "supplier":
            money_button(btns, "📎 مرفقات المورد المحدد", UI["teal"], self.open_selected_supplier_attachments, 9
                         ).pack(side="right", padx=5)
        if self.can("import"):
            money_button(btns, "📥 استيراد من Excel/CSV", UI["cyan"], lambda: self.show_page("import"), 9).pack(side="right", padx=5)
        if self.can_modify():
            money_button(btns, f"🗑️ حذف {who} المحدد", "#dc2626", lambda r=role: self._pp_delete(r), 9).pack(side="left", padx=5)
        sf = tk.Frame(f)
        sf.grid(row=7, column=0, columnspan=6, pady=2)
        ttk.Label(sf, text="بحث (الاسم / الكود / الهاتف / الرقم الضريبي / المندوب):").pack(side="right", padx=5)
        st["search"] = ttk.Entry(sf, width=32)
        st["search"].pack(side="right", padx=5)
        st["search"].bind("<KeyRelease>", lambda e, r=role: self.load_partner_page(r))

        box = ttk.LabelFrame(frame, text=f" قائمة {'العملاء' if role == 'customer' else 'الموردين'} (اضغط مرتين للتعديل) - "
                                          "«مرتبط بـ» = نفس الكيان في الشجرة الأخرى (مقاصة بالكشف الموحد) ", padding=8)
        box.pack(fill="both", expand=True, padx=10, pady=5)
        st["tree"] = make_tree(box, [
            ("code", "الكود", 80, "center"), ("name", "الاسم", 250, "e"), ("phone", "الهاتف", 110, "center"),
            ("vat", "الرقم الضريبي", 125, "center"), ("rep", "المندوب", 110, "center"), ("opening", "الرصيد الافتتاحي", 130, "center"),
            ("linked", "مرتبط بـ", 110, "center"), ("legacy", "الكود القديم", 90, "center")])
        st["tree"].bind("<Double-1>", lambda e, r=role: self._pp_edit(r))
        self._pp_clear(role)

    def _pp_set(self, entry, value):
        disabled = entry.instate(["disabled"])
        entry.state(["!disabled"])
        entry.delete(0, tk.END)
        entry.insert(0, value or "")
        if disabled:
            entry.state(["disabled"])

    def _pp_clear(self, role):
        st = self.pp[role]
        st["editing"] = None
        for k in ("name", "phone", "email", "address", "cr", "vat", "od", "oc", "odate", "cr_days", "cr_limit"):
            if k in st:
                self._pp_set(st[k], "")
        if "cr_policy" in st:
            st["cr_policy"].set("(الافتراضية)")
        st["rep"].set("")
        st["rep"].configure(values=[""] + rep_names())
        self._pp_set(st["rep_date"], today())
        who = "العميل" if role == "customer" else "المورد"
        st["save_btn"].config(text=f"حفظ وإضافة {who} ➕")
        with db() as conn:
            st["code_lbl"].config(text=f"الكود التلقائي لل{who[2:]} الجديد: {peek_serial(conn, role)[0]}")

    def _pp_edit(self, role):
        st = self.pp[role]
        sel = st["tree"].selection()
        if not sel:
            return
        with db() as conn:
            r = conn.execute("SELECT partner_code, name, phone, email, address, cr_number, vat_number, COALESCE(rep_name,''), "
                             "COALESCE(opening_debit,0), COALESCE(opening_credit,0), COALESCE(opening_date,''), "
                             "COALESCE(credit_limit,0), credit_days, COALESCE(credit_policy,'') "
                             "FROM business_partners WHERE partner_code=?", (sel[0],)).fetchone()
        if not r:
            return
        self._pp_clear(role)
        st["editing"] = r[0]
        for k, v in zip(("name", "phone", "email", "address", "cr", "vat"), r[1:7]):
            self._pp_set(st[k], v)
        st["rep"].configure(values=[""] + rep_names(include=r[7] or None))
        st["rep"].set(r[7])
        self._pp_set(st["od"], f"{r[8]:g}" if r[8] else "")
        self._pp_set(st["oc"], f"{r[9]:g}" if r[9] else "")
        self._pp_set(st["odate"], r[10])
        self._pp_set(st["cr_days"], "" if r[12] is None else str(r[12]))
        if "cr_limit" in st:
            self._pp_set(st["cr_limit"], f"{r[11]:g}" if r[11] else "")
            st["cr_policy"].set(_credit.CREDIT_POLICIES.get(r[13], "(الافتراضية)"))
        st["save_btn"].config(text=f"حفظ التعديلات على {r[0]} 💾")
        st["code_lbl"].config(text=f"تعديل: {r[0]}   |   تغيير المندوب يُسجَّل اعتباراً من «تاريخ سريان المندوب» ولا يغيّر حركات الفترات السابقة")

    def _pp_save(self, role):
        st = self.pp[role]
        try:
            vat = normalize_digits(st["vat"].get()).strip()
            if vat and not (vat.isdigit() and len(vat) == 15):
                if not messagebox.askyesno("تنبيه", "الرقم الضريبي السعودي يتكون من 15 رقماً.\nهل تريد الحفظ كما هو؟"):
                    return
            opening = None
            if self.can_modify():
                od, oc = parse_float(st["od"].get() or "0"), parse_float(st["oc"].get() or "0")
                if od is None or oc is None or od < 0 or oc < 0:
                    raise UserError("الرصيد الافتتاحي (مدين/دائن) أرقام صفر أو أكثر.")
                opening = (od, oc, st["odate"].get().strip())
            rep_date = parse_date(st["rep_date"].get())
            if st["rep"].get().strip() and not rep_date:
                raise UserError("تاريخ سريان المندوب غير صحيح (YYYY-MM-DD).")
            editing = st["editing"]
            days_t = normalize_digits(st["cr_days"].get()).strip()
            if days_t and not days_t.isdigit():
                raise UserError("مدة السداد عدد أيام صحيح أو اتركها فارغة للافتراضية.")
            credit = (None, int(days_t) if days_t else "", None)
            if "cr_limit" in st:
                lim = parse_float(st["cr_limit"].get() or "0")
                if lim is None or lim < 0:
                    raise UserError("الحد الائتماني رقم صفر أو أكثر (0 = بدون حد).")
                pol = next((k for k, v in _credit.CREDIT_POLICIES.items() if v == st["cr_policy"].get()), "")
                credit = (lim, credit[1], pol)
            code = save_partner_record(role, editing, st["name"].get(), st["phone"].get(), st["email"].get(),
                                       st["address"].get(), st["cr"].get(), vat, st["rep"].get(), rep_date,
                                       opening, self.user["username"], credit=credit)
            links = [c for c in linked_partner_codes(code) if c != code]
            messagebox.showinfo("نجاح", f"تم {'تعديل' if editing else 'إضافة'} {ROLE_LABEL[role]} بالكود {code}"
                                       + (f"\nمرتبط تلقائياً بنفس الكيان: {' ، '.join(links)} (مقاصة في الكشف الموحد)" if links else ""))
            self._pp_clear(role)
            self.load_partner_page(role)
        except Exception as e:
            handle_error(e)

    def _pp_delete(self, role):
        if not self.require_modify():
            return
        sel = self.pp[role]["tree"].selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر من القائمة أولاً.")
            return
        if not messagebox.askyesno("حذف", f"حذف {sel[0]} نهائياً؟"):
            return
        try:
            delete_partner(sel[0])
            self._pp_clear(role)
            self.load_partner_page(role)
        except Exception as e:
            handle_error(e)

    def _pp_statement(self, role):
        sel = self.pp[role]["tree"].selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر من القائمة أولاً.")
            return
        self.show_page("statement")
        self.update_statement_partners_dropdown()
        label = next((lbl for lbl, c in self.stmt_partner_cb._map.items() if c == sel[0]), None)
        if label:
            self.stmt_partner_cb.set(label)
        self.stmt_scope_cb.set(SUMMARY_SCOPES["both"])
        self.show_partner_statement(sel[0])

    def load_partner_page(self, role):
        st = getattr(self, "pp", {}).get(role)
        if not st:
            return
        clear_tree(st["tree"])
        fin = self.can_fin(role)
        for code, name, phone, vat, rep, ob, legacy, *_rest, links in list_partners(role, st["search"].get().strip()):
            opening = "-" if abs(ob) < 0.005 else f"{abs(ob):,.2f} {'مدين' if ob > 0 else 'دائن'}"
            st["tree"].insert("", "end", iid=code, values=(code, name, phone or "-", vat or "-", rep or "-",
                                                           opening if fin else "🔒", links or "-", legacy or "-"))

    def load_entities(self):
        for role in ("customer", "supplier"):
            self.load_partner_page(role)

    def show_rep_history(self, role):
        st = self.pp[role]
        code = st["editing"] or (st["tree"].selection() or [None])[0]
        if not code:
            messagebox.showinfo("تنبيه", "اختر الطرف من القائمة أولاً.")
            return
        win = tk.Toplevel(self.root)
        win.title(f"سجل المناديب - {code}")
        win.geometry("640x320")
        win.transient(self.root)
        tk.Label(win, text=f"سجل تغيير المناديب للطرف {code} (الحركات تُنسب للمندوب حسب فترة تواجده الفعلية)",
                 font=(UI_FONT, 10, "bold"), fg=UI["primary_dk"]).pack(anchor="e", padx=10, pady=8)
        holder = tk.Frame(win)
        holder.pack(fill="both", expand=True, padx=10, pady=4)
        t = make_tree(holder, [("rep", "المندوب", 160, "e"), ("from", "من", 100, "center"), ("to", "إلى", 100, "center"),
                               ("by", "بواسطة", 90, "center"), ("at", "وقت التسجيل", 140, "center")], rtl=True)
        for rep, s_, e_, by, at in partner_rep_history(code):
            t.insert("", "end", values=(rep, s_, e_ or "حتى الآن", by or "-", at or "-"))

    # ---------------------------------------------------------- شجرة الحسابات: فروع + كشف تحليلي
    def _selected_account(self):
        sel = self.coa_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر حساباً من الشجرة أولاً.")
            return None
        return sel[0]

    def add_sub_account_selected(self):
        parent = self._selected_account()
        if not parent:
            return
        win = tk.Toplevel(self.root)
        win.title("إضافة حساب فرعي")
        win.transient(self.root)
        win.grab_set()
        f = ttk.Frame(win, padding=14)
        f.pack(fill="both", expand=True)
        with db() as conn:
            pname = conn.execute("SELECT account_name FROM chart_of_accounts WHERE account_code=?", (parent,)).fetchone()[0]
            nxt = next_child_account_code(conn, parent)
        ttk.Label(f, text=f"تحت الحساب: {parent} - {pname}", font=(UI_FONT, 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="e")
        ttk.Label(f, text=f"الكود التلقائي: {nxt}", foreground=UI["primary_dk"]).grid(row=1, column=0, columnspan=2, sticky="e", pady=4)
        name_e = ttk.Entry(f, width=36)
        ttk.Label(f, text="اسم الحساب الفرعي:").grid(row=2, column=1, sticky="e", padx=4, pady=6)
        name_e.grid(row=2, column=0, sticky="w", pady=6)
        grp_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(f, text="حساب تجميعي (ستُضاف تحته حسابات فرعية أخرى)", variable=grp_var).grid(row=3, column=0, columnspan=2, sticky="e")
        ttk.Label(f, text="أمثلة: «سلفة الموظف محمد» تحت 11410، «خزينة فرع جدة» تحت 11110، «حساب جاري - الراجحي» تحت 11120.",
                  foreground=UI["muted"], font=(UI_FONT, 8)).grid(row=4, column=0, columnspan=2, sticky="e", pady=4)

        def save():
            # نقرأ كل القيم قبل إغلاق النافذة: أي قراءة من عناصرها بعد destroy تسبب TclError: invalid command name
            name, is_group = name_e.get().strip(), bool(grp_var.get())
            try:
                code = create_sub_account(parent, name, None, is_group)
            except Exception as e:
                handle_error(e)
                return
            try:
                win.grab_release()
            except tk.TclError:
                pass
            win.destroy()
            self.root.after_idle(lambda: self._after_sub_account_saved(code, name))

        win.bind("<Return>", lambda e: save())

        money_button(f, "💾 حفظ الحساب الفرعي", "#16a34a", save, 10).grid(row=5, column=0, columnspan=2, pady=10)
        name_e.focus_set()

    def _after_sub_account_saved(self, code, name):
        """تحديث الشجرة بعد إغلاق النافذة بأمان (خارج دورة أحداث النافذة المغلقة)."""
        try:
            self.load_coa_data()
            tree = getattr(self, "coa_tree", None)
            if tree is not None and tree.winfo_exists() and tree.exists(code):
                parent = tree.parent(code)
                while parent:
                    tree.item(parent, open=True)
                    parent = tree.parent(parent)
                tree.selection_set(code)
                tree.see(code)
        except tk.TclError:
            pass
        for fn in ("refresh_cash_accounts", "load_journal_accounts"):
            try:
                getattr(self, fn)()
            except Exception:
                pass
        messagebox.showinfo("تم", f"تم إنشاء الحساب {code} - {name}")

    def open_account_statement(self, code=None):
        code = code or self._selected_account()
        if not code:
            return
        win = tk.Toplevel(self.root)
        win.title(f"كشف حساب تحليلي - {code}")
        win.geometry("1100x620")
        win.transient(self.root)
        bar = tk.Frame(win)
        bar.pack(fill="x", padx=10, pady=8)
        d1, d2 = date_entry(bar), date_entry(bar)
        d1.delete(0, tk.END)
        d1.insert(0, f"{date.today().year}-01-01")
        head = tk.Label(win, text="", font=(UI_FONT, 12, "bold"), fg=UI["primary_dk"])
        cards = tk.Frame(win)
        holder = tk.Frame(win)
        state = {}

        def card(title, color):
            c = tk.Frame(cards, bg=UI["card"], highlightbackground=UI["border"], highlightthickness=1)
            c.pack(side="right", fill="x", expand=True, padx=4)
            tk.Frame(c, bg=color, height=4).pack(fill="x")
            tk.Label(c, text=title, bg=UI["card"], fg=UI["muted"], font=(UI_FONT, 9, "bold")).pack(anchor="e", padx=8)
            v = tk.Label(c, text="-", bg=UI["card"], font=(UI_FONT, 13, "bold"))
            v.pack(anchor="e", padx=8, pady=(0, 6))
            return v

        v_open = card("رصيد أول المدة", "#64748b")
        v_dr = card("المسحوب / مدين", UI["primary"])
        v_cr = card("المسدد / دائن", UI["teal"])
        v_close = card("المتبقي (الرصيد)", "#d97706")
        tree = make_tree(holder, [("date", "التاريخ", 90, "center"), ("ref", "المرجع", 85, "center"),
                                  ("acc", "الحساب", 230, "e"), ("memo", "البيان", 330, "e"), ("dr", "مدين", 95, "center"),
                                  ("cr", "دائن", 95, "center"), ("bal", "الرصيد", 110, "center")], rtl=True)

        def load():
            try:
                a, b = parse_date(d1.get()), parse_date(d2.get())
                if not a or not b:
                    raise UserError("صيغة التاريخ YYYY-MM-DD")
                st_ = account_statement(code, a, b)
            except Exception as e:
                handle_error(e)
                return
            state["st"], state["range"] = st_, (a, b)
            head.config(text=f"{st_['code']} - {st_['name']} ({st_['type']})" +
                             (f"   |   يشمل {st_['accounts'] - 1} حساباً فرعياً" if st_["group"] else ""))
            clear_tree(tree)
            for d, ref, acc, memo, dr, cr, bal in st_["rows"]:
                tree.insert("", "end", values=(d, ref, acc, memo, fmt(dr) if dr else "-", fmt(cr) if cr else "-", fmt(bal)))
            for lbl, v in ((v_open, st_["opening"]), (v_dr, st_["debit"]), (v_cr, st_["credit"]), (v_close, st_["closing"])):
                lbl.config(text=fmt(v))

        def export(kind):
            st_ = state.get("st")
            if not st_:
                return
            rep = Report(f"كشف حساب تحليلي: {st_['code']} - {st_['name']}",
                         [("التاريخ", "t"), ("المرجع", "t"), ("الحساب", "t"), ("البيان", "t"), ("مدين", "m"), ("دائن", "m"),
                          ("الرصيد", "m")], st_["rows"], ("", "", "", "الإجمالي", st_["debit"], st_["credit"], st_["closing"]),
                         [f"رصيد أول المدة: {fmt(st_['opening'])}   |   المتبقي: {fmt(st_['closing'])}"],
                         f"من {state['range'][0]} إلى {state['range'][1]}")
            ext = ".xlsx" if kind == "x" else ".pdf"
            path = filedialog.asksaveasfilename(defaultextension=ext, initialfile=f"كشف_حساب_{code}{ext}", parent=win)
            if path:
                try:
                    (export_report_xlsx if kind == "x" else build_report_pdf)(path, rep)
                    messagebox.showinfo("تم", f"تم الحفظ:\n{path}", parent=win)
                except Exception as e:
                    handle_error(e)

        for w, lbl in ((d1, "من:"), (d2, "إلى:")):
            tk.Label(bar, text=lbl).pack(side="right", padx=(8, 2))
            w.pack(side="right")
        money_button(bar, "🔍 عرض", UI["primary"], load, 9).pack(side="right", padx=8)
        money_button(bar, "📊 Excel", "#475569", lambda: export("x"), 9).pack(side="left", padx=3)
        money_button(bar, "📄 PDF", "#475569", lambda: export("p"), 9).pack(side="left", padx=3)
        head.pack(anchor="e", padx=12)
        cards.pack(fill="x", padx=10, pady=6)
        holder.pack(fill="both", expand=True, padx=10, pady=4)
        load()

    # ---------------------------------------------------------- غير المفوتر (مبيعات + مشتريات في جدول واحد)
    UNBILLED_TYPES = {"الكل": None, "المبيعات": "sales", "المشتريات": "purchase"}

    def load_unbilled(self, kind):
        st = self.flow.get(kind) or {}
        if "u_tree" not in st:
            return
        tree = st["u_tree"]
        clear_tree(tree)
        rows = all_unbilled(self.UNBILLED_TYPES.get(st["u_type"].get()), st["u_search"].get().strip())
        tot = 0.0
        for r in rows:
            tag = ("old",) if r["age"] > 30 else (("aging",) if r["age"] > 7 else ())
            tree.insert("", "end", iid=f"{r['kind']}:{r['id']}", tags=tag, values=(
                "مبيعات (تسليم)" if r["kind"] == "sales" else "مشتريات (استلام)", r["id"], r["date"], r["age"],
                r["order"], f"{r['party_code']} - {r['party']}", r["lines"], fmt(r["subtotal"]), fmt(r["vat"]), fmt(r["total"])))
            tot += r["total"]
        st["u_count_lbl"].config(text=f"عدد الأذون غير المفوترة: {len(rows)}   |   الإجمالي شامل الضريبة: {fmt(tot)} ر.س"
                                      "   |   البرتقالي أقدم من 7 أيام، والأحمر أقدم من 30 يوماً")

    def make_flow_invoice(self, kind):
        st = self.flow[kind]
        try:
            sel = st["u_tree"].selection()
            if not sel:
                raise UserError("حدد إذناً واحداً على الأقل من الجدول (Ctrl للتحديد المتعدد).")
            kinds = {s.split(":")[0] for s in sel}
            if len(kinds) > 1:
                raise UserError("حددت أذون مبيعات ومشتريات معاً. الفاتورة الواحدة تكون من نوع واحد ولطرف واحد.")
            k = kinds.pop()
            ids = [int(s.split(":")[1]) for s in sel]
            inv_date = parse_date(st["i_date"].get())
            if not inv_date:
                raise UserError("صيغة تاريخ الفاتورة غير صحيحة.")
            self.check_period_open(inv_date)
            ref = st["i_ref"].get().strip() if ("i_ref" in st and k == "purchase") else ""
            iid, total = invoice_unbilled_selection(k, ids, inv_date, st["i_notes"].get().strip(), ref,
                                                    allow_closed=True)  # check_period_open تحقق بالفعل
            st["i_notes"].delete(0, tk.END)
            if "i_ref" in st:
                st["i_ref"].delete(0, tk.END)
            for kk in ("sales", "purchase"):
                self.load_unbilled(kk)
            self.load_orders(k)
            self.load_fulfilments(k)
            if k == "sales":
                self.load_sales_invoices()
                if messagebox.askyesno("نجاح", f"تم إصدار الفاتورة رقم {invoice_number(k, iid)} بإجمالي {fmt(total)} ر.س\n\n"
                                              "هل تريد طباعتها PDF الآن؟"):
                    self.export_sales_invoice(iid)
            else:
                self.load_purchase_invoices()
                if messagebox.askyesno("نجاح", f"تم تسجيل فاتورة المورد رقم {invoice_number(k, iid)} بإجمالي {fmt(total)} ر.س\n\n"
                                              "هل تريد إرفاق صورة/ملف فاتورة المورد الأصلية الآن؟"):
                    self.open_attachments_window("purchase", iid, upload_now=True)
        except Exception as e:
            handle_error(e)

    # ---------------------------------------------------------- النوافذ المنبثقة (Double-Click)
    def open_invoice_modal(self, kind, event=None):
        tree = self.sal_tree if kind == "sales" else self.pur_tree
        sel = tree.selection()
        if sel:
            try:
                self.show_transaction_details((kind, int(sel[0])))
            except Exception as e:
                handle_error(e)

    def _modal_extras(self, win, kind, ref_id):
        """أزرار مستقلة داخل نافذة التفاصيل: تصدير PDF، ومرفقات فاتورة المورد."""
        bar = tk.Frame(win)
        kids = win.winfo_children()
        anchor = kids[1] if len(kids) > 2 else None  # بعد شريط البيانات مباشرة (قبل جدول البنود المتمدد)
        if anchor is not None:
            bar.pack(fill="x", padx=10, pady=(4, 2), after=anchor)
        else:
            bar.pack(fill="x", padx=10, pady=(4, 2))
        if kind in ("sales", "purchase"):
            money_button(bar, "🖨️ تصدير PDF", UI["primary"], lambda: self.export_document_pdf(
                build_invoice_pdf, f"فاتورة_{invoice_number(kind, ref_id)}", kind, ref_id), 9).pack(side="right", padx=3)
            try:
                pre = invoice_payment_prefill(kind, ref_id)
            except UserError:
                pre = None
            perm = "cash.receipt" if kind == "sales" else "cash.payment"
            if pre and pre["remaining"] > 0.005 and self.can(perm):
                money_button(bar, f"💳 سداد هذه الفاتورة ({fmt(pre['remaining'])})", "#16a34a",
                             lambda: (win.destroy(), self.prefill_payment(kind, ref_id)), 9).pack(side="right", padx=3)
            if pre and pre["remaining"] > 0.005 and self.can(perm):
                avail = party_advance_available(pre["party"][0], kind)
                if avail > 0.005:
                    money_button(bar, f"🔗 تسوية من الدفعات المقدمة (متاح {fmt(avail)})", UI["teal"],
                                 lambda: self.settle_invoice_advances(kind, ref_id, win), 9).pack(side="right", padx=3)
        if kind == "discount":
            money_button(bar, "🖨️ تصدير PDF", UI["primary"], lambda: self.export_document_pdf(
                build_post_discount_pdf, f"إشعار_{ref_id}", ref_id), 9).pack(side="right", padx=3)
        if kind == "purchase":
            n = len(list_attachments("purchase", ref_id))
            money_button(bar, f"📎 مرفقات فاتورة المورد ({n})", UI["teal"],
                         lambda: self.open_attachments_window("purchase", ref_id, parent=win), 9).pack(side="right", padx=3)

    def show_item_modal(self, event=None):
        sel = self.inv_tree.selection()
        if not sel:
            return
        try:
            d = item_details(sel[0])
        except Exception as e:
            handle_error(e)
            return
        win = tk.Toplevel(self.root)
        win.title(f"الصنف {d['code']} - {d['name']}")
        win.geometry("1080x560")
        win.transient(self.root)
        tk.Label(win, text=f"{d['name']}", font=(UI_FONT, 14, "bold"), bg=UI["slate2"], fg="white", pady=8).pack(fill="x")
        info = tk.Frame(win, bg=UI["card"], padx=10, pady=6)
        info.pack(fill="x", padx=10, pady=6)
        fields = [("الكود", d["code"]), ("الوحدة", d["unit"]), ("الرصيد الحالي", f"{d['qty']:,.3f}"),
                  ("متوسط التكلفة", fmt(d["cost"])), ("قيمة المخزون", fmt(d["value"])),
                  ("سعر البيع", fmt(d["price"]) if d["price"] else "-"),
                  ("هامش الربح", f"{d['margin']:.1f}%" if d["margin"] is not None else "-"),
                  ("آخر تكلفة شراء", fmt(d["last_cost"]) if d["last_cost"] is not None else "-"),
                  ("الباركود", d["barcode"] or "-"), ("الكود القديم", d["legacy"] or "-"),
                  ("إجمالي المبيعات (كمية)", f"{d['sold_qty']:,.3f}"), ("إجمالي المشتريات (كمية)", f"{d['bought_qty']:,.3f}")]
        for i, (lab, val) in enumerate(fields):
            r_, c_ = divmod(i, 4)
            cell = tk.Frame(info, bg=UI["card"])
            cell.grid(row=r_, column=3 - c_, sticky="e", padx=10, pady=2)
            tk.Label(cell, text=str(val), bg=UI["card"], font=(UI_FONT, 10, "bold")).pack(side="right")
            tk.Label(cell, text=f"{lab}:", bg=UI["card"], fg=UI["muted"], font=(UI_FONT, 9)).pack(side="right")
        for c in range(4):
            info.grid_columnconfigure(c, weight=1)
        bar = tk.Frame(win)
        bar.pack(fill="x", padx=10)
        tk.Label(bar, text="آخر الحركات", font=(UI_FONT, 10, "bold"), fg=UI["primary_dk"]).pack(side="right")
        money_button(bar, "🗂️ كارت الصنف الكامل / PDF", UI["primary"], lambda: (win.destroy(), self.open_item_card()), 9).pack(side="left", padx=3)
        if self.can_modify():
            money_button(bar, "🛠️ تعديل الصنف", "#d97706", lambda: (win.destroy(), self.edit_item_dialog()), 9).pack(side="left", padx=3)
        holder = tk.Frame(win)
        holder.pack(fill="both", expand=True, padx=10, pady=6)
        t = make_tree(holder, [("date", "التاريخ", 90, "center"), ("type", "الحركة", 200, "e"), ("party", "الطرف", 220, "e"),
                               ("qty", "الكمية", 80, "center"), ("bal", "الرصيد", 90, "center"), ("price", "السعر", 85, "center"),
                               ("value", "القيمة", 95, "center"), ("notes", "ملاحظات", 160, "e")], rtl=True)
        for r in reversed(d["moves"]):
            t.insert("", "end", values=r)

    # ---------------------------------------------------------- المرفقات
    def open_attachments_window(self, doc_type, doc_id, parent=None, upload_now=False):
        win = tk.Toplevel(parent or self.root)
        label = ATTACH_DOCS[doc_type][2]
        num = invoice_number("purchase", doc_id) if doc_type == "purchase" else doc_id
        win.title(f"مرفقات {label} {num}")
        win.geometry("980x440")
        win.transient(parent or self.root)
        tk.Label(win, text=f"📎 مرفقات {label} رقم {num}  —  اضغط مرتين على الملف لفتحه", font=(UI_FONT, 11, "bold"),
                 fg=UI["primary_dk"]).pack(anchor="e", padx=10, pady=8)
        body = tk.Frame(win)
        body.pack(fill="both", expand=True, padx=10)
        # معاينة فورية للصور داخل النظام (يسار) + قائمة الملفات (يمين)
        prev = tk.Frame(body, bg=UI["card"], highlightbackground=UI["border"], highlightthickness=1, width=300)
        prev.pack(side="left", fill="y", padx=(0, 8))
        prev.pack_propagate(False)
        prev_lbl = tk.Label(prev, text="المعاينة\n\nحدد ملفاً لعرضه هنا", bg=UI["card"], fg=UI["muted"], font=(UI_FONT, 9))
        prev_lbl.pack(fill="both", expand=True)
        holder = tk.Frame(body)
        holder.pack(side="right", fill="both", expand=True)
        t = make_tree(holder, [("name", "اسم الملف", 300, "e"), ("size", "الحجم", 90, "center"), ("by", "بواسطة", 90, "center"),
                               ("at", "وقت الرفع", 150, "center")], rtl=True)

        def show_preview(event=None):
            sel = t.selection()
            if not sel:
                return
            try:
                name, _mime, data = get_attachment(int(sel[0]))
                ext = os.path.splitext(name)[1].lower()
                if ext in (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".tif", ".tiff"):
                    from PIL import ImageTk
                    im = Image.open(io.BytesIO(data))
                    im.thumbnail((290, 400))
                    photo = ImageTk.PhotoImage(im)
                    prev_lbl.configure(image=photo, text="")
                    prev_lbl.image = photo
                else:
                    prev_lbl.configure(image="", text=f"📄 {name}\n\n{len(data) / 1024:,.1f} KB\n\nاضغط مرتين أو «📂 فتح»\nلعرض الملف")
                    prev_lbl.image = None
            except Exception:
                prev_lbl.configure(image="", text="تعذرت المعاينة\nاستخدم «📂 فتح»")
                prev_lbl.image = None

        t.bind("<<TreeviewSelect>>", show_preview, add="+")

        def load():
            clear_tree(t)
            rows = list_attachments(doc_type, doc_id)
            for aid, name, size, _mime, by, at in rows:
                t.insert("", "end", iid=str(aid), values=(name, f"{size / 1024:,.1f} KB", by or "-", at or "-"))
            if rows:
                t.selection_set(str(rows[0][0]))
            else:
                prev_lbl.configure(image="", text="لا توجد مرفقات بعد\n\nاضغط «➕ رفع ملفات»")
                prev_lbl.image = None
            # تحديث عدد المرفقات في سجل الفواتير/القيود
            for fn in (("load_purchase_invoices",) if doc_type == "purchase" else ()):
                try:
                    getattr(self, fn)()
                except Exception:
                    pass
            if doc_type == "receipt":
                try:
                    self.load_fulfilments("purchase")
                except Exception:
                    pass

        def add():
            paths = filedialog.askopenfilenames(parent=win, title="اختر الملفات (صور أو PDF)",
                                                filetypes=[("صور و PDF", "*.pdf *.png *.jpg *.jpeg *.gif *.bmp *.webp *.tif *.tiff"),
                                                           ("كل الملفات المدعومة", " ".join(f"*{e}" for e in ATTACH_EXTS))])
            for path in paths or ():
                try:
                    with open(path, "rb") as fh:
                        add_attachment(doc_type, doc_id, os.path.basename(path), fh.read(), self.user["username"])
                except Exception as e:
                    handle_error(e)
            load()

        def open_selected(event=None):
            sel = t.selection()
            if not sel:
                return
            try:
                name, _mime, data = get_attachment(int(sel[0]))
                folder = os.path.join(tempfile.gettempdir(), "elhagry_attachments")
                os.makedirs(folder, exist_ok=True)
                path = os.path.join(folder, f"{sel[0]}_{name}")
                with open(path, "wb") as fh:
                    fh.write(data)
                try:
                    os.startfile(path)
                except (AttributeError, OSError):
                    messagebox.showinfo("المرفق", f"تم حفظ الملف مؤقتاً:\n{path}", parent=win)
            except Exception as e:
                handle_error(e)

        def save_as():
            sel = t.selection()
            if not sel:
                return
            name, _mime, data = get_attachment(int(sel[0]))
            path = filedialog.asksaveasfilename(parent=win, initialfile=name)
            if path:
                with open(path, "wb") as fh:
                    fh.write(data)

        def remove():
            sel = t.selection()
            if sel and self.require_modify() and messagebox.askyesno("حذف مرفق", "حذف المرفق المحدد؟", parent=win):
                delete_attachment(int(sel[0]))
                load()

        bar = tk.Frame(win)
        bar.pack(fill="x", padx=10, pady=8)
        money_button(bar, "➕ رفع ملفات", UI["teal"], add, 9).pack(side="right", padx=3)
        money_button(bar, "📂 فتح", UI["primary"], open_selected, 9).pack(side="right", padx=3)
        money_button(bar, "💾 حفظ باسم", "#475569", save_as, 9).pack(side="right", padx=3)
        if self.can_modify():
            money_button(bar, "🗑️ حذف", "#dc2626", remove, 9).pack(side="left", padx=3)
        t.bind("<Double-1>", open_selected)
        load()
        if upload_now:
            win.after(200, add)

    def open_purchase_attachments_selected(self):
        inv_id = self._selected_invoice("purchase")
        if inv_id is not None:
            self.open_attachments_window("purchase", inv_id)

    def open_journal_attachments_selected(self):
        sel = self.jv_reg_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر قيداً من السجل.")
            return
        self.open_attachments_window("journal", int(sel[0]))

    FLOW_TEXT = {
        "sales": dict(
            new="المرحلة 1: إنشاء أمر بيع (لا يؤثر على المخزن)", orders="أوامر البيع", party="العميل", rep_label="المندوب",
            price="السعر", save_btn="حفظ أمر البيع 📝", go_btn="🚚 ترحيل إلى إذن تسليم",
            ful_title="المرحلة 2: تسليم البضاعة (يخصم من المخزن، ويدعم التسليم على دفعات)", ful_btn="🚚 حفظ إذن التسليم وخصم المخزن",
            ful_noun="إذن التسليم", ful_reg="سجل أذونات التسليم", inv_title="المرحلة 3: إصدار فاتورة ضريبية من تسليمات العميل",
            tab_o="📝 1) أوامر البيع", tab_f="🚚 2) تسليم البضاعة", tab_i="🧾 3) الفواتير الضريبية", tab_d="⚡ فاتورة مباشرة (بدون أمر)",
            tab_x="🏷️ تسجيل خصم لاحق / إشعار دائن",
            done_col="% المسلَّم", unbilled="التسليمات غير المفوترة"),
        "purchase": dict(
            new="المرحلة 1: إنشاء أمر شراء (لا يؤثر على المخزن)", orders="أوامر الشراء", party="المورد", rep_label="المسؤول",
            price="التكلفة", save_btn="حفظ أمر الشراء 📝", go_btn="📥 ترحيل إلى إذن استلام",
            ful_title="المرحلة 2: استلام البضاعة بالمخزن (يضيف للمخزن ويحدّث التكلفة، ويدعم الاستلام على دفعات)",
            ful_btn="📥 حفظ إذن الاستلام وإضافة للمخزن", ful_noun="إذن الاستلام", ful_reg="سجل أذونات الاستلام",
            inv_title="المرحلة 3: تسجيل فاتورة المورد من الاستلامات", tab_o="📝 1) أوامر الشراء", tab_f="📥 2) استلام المخزن",
            tab_i="🧾 3) فواتير الموردين", tab_d="⚡ 4) فاتورة شراء مباشر (بدون أمر)",
            tab_x="🏷️ تسجيل خصم لاحق / إشعار دائن", done_col="% المستلَم", unbilled="الاستلامات غير المفوترة"),
    }

    def check_period_open(self, d):
        """يوم مغلق بالإغلاق اليومي يحتاج كلمة مرور المدير (المدير نفسه له صلاحية كاملة)."""
        if self.user["is_admin"]:
            return
        if period_is_closed(d):
            pwd = simpledialog.askstring("تأكيد إداري", f"تاريخ ({d}) تم إغلاقه!\nأدخل كلمة مرور المدير للمتابعة:", show="*")
            if pwd != get_setting("admin_password", "2026"):
                raise UserError("كلمة المرور غير صحيحة. تم إلغاء العملية.")

    def setup_sales_tab(self, frame):
        self.build_flow_page(frame, "sales")

    def setup_purchase_tab(self, frame):
        self.build_flow_page(frame, "purchase")

    def build_flow_page(self, frame, kind):
        T = self.FLOW_TEXT[kind]
        nb = ttk.Notebook(frame)
        nb.pack(fill="both", expand=True, padx=4, pady=4)
        tabs = {}
        tab_defs = (("o", T["tab_o"]), ("f", T["tab_f"]), ("i", T["tab_i"]), ("x", T["tab_x"]))
        for key, text in tab_defs:
            tabs[key] = ttk.Frame(nb)
        for key, text in reversed(tab_defs):
            nb.add(tabs[key], text=f"  {text}  ")  # ترتيب عربي: المرحلة 1 في أقصى اليمين
        self.flow[kind] = {"nb": nb, "tabs": tabs}
        self.build_orders_tab(tabs["o"], kind)
        self.build_fulfil_tab(tabs["f"], kind)
        self.build_flow_invoices_tab(tabs["i"], kind)
        self.build_discount_tab(tabs["x"], kind)
        nb.bind("<<NotebookTabChanged>>", lambda e, k=kind: self.refresh_flow(k))
        visible = []
        for key in ("o", "f", "i", "x"):
            if self.can(self.FLOW_TAB_PERMS[kind][key]):
                visible.append(key)
            else:
                nb.hide(tabs[key])
        if visible:
            nb.select(tabs[visible[0]])

    FLOW_TAB_PERMS = {"sales": {"o": "sales.orders", "f": "sales.delivery", "i": "sales.invoices", "x": "sales.discount"},
                      "purchase": {"o": "purchase.orders", "f": "purchase.receipt", "i": "purchase.invoices",
                                   "x": "purchase.discount"}}

    def flow_select(self, kind, key):
        """فتح تبويب داخل المبيعات/المشتريات فقط لو كان مسموحاً (اختيار تبويب مخفي في ttk يعيد إظهاره)."""
        if not self.can(self.FLOW_TAB_PERMS[kind][key]):
            return False
        st = self.flow[kind]
        st["nb"].select(st["tabs"][key])
        return True

    def refresh_flow(self, kind):
        st = self.flow.get(kind)
        if not st or "o_tree" not in st or "u_tree" not in st:
            return
        try:
            pairs = self.partner_pairs("customer" if kind == "sales" else "supplier")
            st["party_cb"].set_items(pairs)
            cur = st["o_rep"].get()
            st["o_rep"].configure(values=[""] + rep_names(include=cur or None))
            st["cart"].set_items(self.item_pairs())
            self.load_unbilled(kind)
            self.load_orders(kind)
            self.refresh_order_combo(kind)
            self.load_fulfilments(kind)
            if kind == "sales":
                self.load_sales_invoices()
            else:
                self.load_purchase_invoices()
            self.refresh_discount_tab(kind)
        except Exception as e:
            handle_error(e)

    # ---------------------------------------------------------- المرحلة 1: الأوامر
    def build_orders_tab(self, parent, kind):
        T, st = self.FLOW_TEXT[kind], self.flow[kind]
        top = ttk.LabelFrame(parent, text=f" {T['new']} ", padding=8)
        top.pack(fill="x", padx=8, pady=(8, 4))
        hdr = tk.Frame(top)
        hdr.pack(fill="x")
        st["o_date"] = date_entry(hdr)
        st["party_cb"] = SearchableCombobox(hdr, width=34)
        st["o_rep"] = ttk.Combobox(hdr, width=18, values=[""] + rep_names(), state="readonly")
        with db() as conn:
            default_rep = conn.execute("SELECT rep_name FROM company_profile WHERE id=1").fetchone()
        st["o_rep"].set(default_rep[0] if default_rep and default_rep[0] else "")
        st["party_cb"].bind("<<ComboboxSelected>>", lambda e, k=kind: self._order_default_rep(k), add="+")
        st["o_exp"] = DateEntry(hdr, width=13)
        st["o_notes"] = ttk.Entry(hdr, width=44)
        rtl_row(hdr, 0, [("التاريخ:", st["o_date"]), (T["party"] + ":", st["party_cb"]), (T["rep_label"] + ":", st["o_rep"]),
                         ("موعد التنفيذ المتوقع:", st["o_exp"])])
        rtl_row(hdr, 1, [("ملاحظات:", st["o_notes"])])
        st["cart"] = CartPanel(top, price_label=T["price"], warn_below_cost=(kind == "sales"),
                               allow_discount=(kind == "sales"))
        if self.can(f"{kind}.orders_create"):
            money_button(top, T["save_btn"], "#16a34a", lambda: self.save_order(kind), 10).pack(pady=6)
        else:
            ttk.Label(top, text="👁 عرض فقط: إنشاء الأوامر غير متاح لصلاحياتك.", foreground="#b45309").pack(pady=6)

        reg = ttk.LabelFrame(parent, text=f" سجل {T['orders']} (اضغط مرتين لعرض البنود وما نُفّذ منها) ", padding=6)
        reg.pack(fill="both", expand=True, padx=8, pady=4)
        bar = tk.Frame(reg)
        bar.pack(fill="x", pady=(0, 4))
        money_button(bar, T["go_btn"], "#d97706", lambda: self.order_go_fulfil(kind), 9).pack(side="right", padx=3)
        money_button(bar, "🖨️ طباعة الأمر PDF", "#0284c7", lambda: self.print_order(kind), 9).pack(side="right", padx=3)
        money_button(bar, "⛔ إغلاق (إنهاء المتبقي)", "#64748b", lambda: self.order_close_selected(kind), 9).pack(side="right", padx=3)
        money_button(bar, "❌ إلغاء الأمر", "#dc2626", lambda: self.order_cancel_selected(kind), 9).pack(side="right", padx=3)
        st["o_filter"] = ttk.Combobox(bar, values=["المفتوحة فقط", "الكل"], state="readonly", width=14)
        st["o_filter"].set("المفتوحة فقط")
        st["o_filter"].bind("<<ComboboxSelected>>", lambda e: self.load_orders(kind))
        st["o_filter"].pack(side="left", padx=3)
        holder = tk.Frame(reg)
        holder.pack(fill="both", expand=True)
        st["o_tree"] = make_tree(holder, [
            ("id", "رقم الأمر", 70, "center"), ("date", "التاريخ", 90, "center"), ("party", T["party"], 200, "e"),
            ("rep", T["rep_label"], 100, "center"), ("n", "عدد البنود", 70, "center"), ("val", "القيمة شاملة الضريبة", 120, "center"),
            ("done", T["done_col"], 80, "center"), ("inv", "% المفوتر", 80, "center"), ("status", "الحالة", 110, "center")])
        st["o_tree"].bind("<Double-1>", lambda e: self.show_order_lines(kind))

    def _order_default_rep(self, kind):
        """اختيار الطرف في الأمر يملأ مندوبه الافتراضي من بطاقته (ويمكن تغييره من القائمة)."""
        st = self.flow[kind]
        party = st["party_cb"].selected()   # (الكود، الاسم) من partner_pairs
        if not party:
            return
        code = party[0] if isinstance(party, (tuple, list)) else party
        with db() as conn:
            r = conn.execute("SELECT COALESCE(rep_name,'') FROM business_partners WHERE partner_code=?", (code,)).fetchone()
        if r and r[0]:
            names = rep_names(include=r[0])
            st["o_rep"].configure(values=[""] + names)
            st["o_rep"].set(r[0])

    def save_order(self, kind):
        st = self.flow[kind]
        try:
            o_date = parse_date(st["o_date"].get())
            if not o_date:
                raise UserError("صيغة التاريخ غير صحيحة، استخدم YYYY-MM-DD")
            party = st["party_cb"].selected()
            if not party:
                raise UserError(f"اختر {self.FLOW_TEXT[kind]['party']} من القائمة.")
            exp_raw = st["o_exp"].get().strip()
            expected = parse_date(exp_raw) if exp_raw else None
            if exp_raw and not expected:
                raise UserError("موعد التنفيذ المتوقع بصيغة YYYY-MM-DD أو اتركه فارغاً.")
            lines = list(st["cart"].lines)
            if kind == "sales" and lines:
                amount = round(sum(l["total"] for l in lines), 2)
                if not self.credit_precheck(party[0], amount, "order"):
                    return
            oid = self.with_credit(create_order, kind, o_date, party, st["o_notes"].get().strip(), lines, expected,
                                   st["o_rep"].get().strip())
            st["cart"].clear()
            st["o_notes"].delete(0, tk.END)
            st["o_exp"].delete(0, tk.END)
            self.load_orders(kind)
            self.refresh_order_combo(kind)
            noun = "إذن تسليم" if kind == "sales" else "إذن استلام"
            if self.can(self.FLOW_TAB_PERMS[kind]["f"]) and messagebox.askyesno(
                    "نجاح", f"تم حفظ {self.FLOW_TEXT[kind]['orders'][:-1]} رقم {oid}.\nالمخزن لم يتأثر بعد.\n\n"
                            f"هل تريد ترحيله الآن إلى {noun}؟ (تُنقل كل الأصناف والكميات تلقائياً)"):
                self.order_go_fulfil(kind, oid)
        except Exception as e:
            handle_error(e)

    def load_orders(self, kind):
        st, c = self.flow[kind], FLOW[kind]
        clear_tree(st["o_tree"])
        where = "WHERE o.status NOT IN ('مفوتر','مغلق','ملغي')" if st["o_filter"].get().startswith("المفتوحة") else ""
        sql = (f"SELECT o.order_id, o.order_date, o.{c['pname']}, COALESCE(o.{c['rep']},''), COUNT(l.line_id), o.total_amount, "
               f"COALESCE(SUM(l.{c['done']}),0), COALESCE(SUM(l.invoiced_qty),0), COALESCE(SUM(l.quantity),0), o.status "
               f"FROM {c['order']} o LEFT JOIN {c['olines']} l ON l.order_id=o.order_id {where} "
               f"GROUP BY o.order_id ORDER BY o.order_id DESC")
        with db() as conn:
            for oid, d, party, rep, n, total, done, inv, qty, status in conn.execute(sql):
                st["o_tree"].insert("", "end", iid=str(oid), values=(
                    oid, d, party, rep or "-", n, fmt(total), f"{_pct(done, qty):.0f}%", f"{_pct(inv, qty):.0f}%", status))

    def order_selected(self, kind):
        sel = self.flow[kind]["o_tree"].selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر أمراً من السجل أولاً.")
            return None
        return int(sel[0])

    def refresh_order_combo(self, kind):
        st, c = self.flow[kind], FLOW[kind]
        with db() as conn:
            pairs = [(f"أمر {oid} - {party} ({d}) - {status}", oid) for oid, party, d, status in conn.execute(
                f"SELECT order_id, {c['pname']}, order_date, status FROM {c['order']} "
                f"WHERE status NOT IN ('مفوتر','مغلق','ملغي') ORDER BY order_id DESC")]
        st["f_order_pairs"] = pairs
        st["f_order_cb"].set_items(pairs)

    def order_go_fulfil(self, kind, oid=None):
        """ترحيل الأمر إلى شاشة الإذن: يُختار الأمر وتُحمَّل كل بنوده بالكمية المتبقية جاهزة للحفظ."""
        oid = oid or self.order_selected(kind)
        if oid is None:
            return
        st = self.flow[kind]
        self.refresh_order_combo(kind)
        match = next((disp for disp, payload in st["f_order_pairs"] if payload == oid), None)
        if match is None:
            messagebox.showinfo("تنبيه", "هذا الأمر غير قابل للتنفيذ (مفوتر / مغلق / ملغي).")
            return
        st["f_order_cb"].set(match)
        if not self.flow_select(kind, "f"):
            messagebox.showinfo("صلاحيات", "تم حفظ الأمر. تنفيذ الأذون غير متاح لصلاحياتك.")
            return
        self.load_fulfil_lines(kind)

    def order_close_selected(self, kind):
        oid = self.order_selected(kind)
        if oid is None:
            return
        try:
            if messagebox.askyesno("إغلاق الأمر", "سيُغلق الأمر ولن يمكن تنفيذ الكمية المتبقية منه.\nما نُفّذ سابقاً يبقى كما هو.\n\nمتابعة؟"):
                close_order(kind, oid)
                self.load_orders(kind)
                self.refresh_order_combo(kind)
        except Exception as e:
            handle_error(e)

    def order_cancel_selected(self, kind):
        oid = self.order_selected(kind)
        if oid is None:
            return
        try:
            if messagebox.askyesno("إلغاء الأمر", f"إلغاء الأمر رقم {oid}؟ (متاح فقط لأمر لم يُنفَّذ منه شيء)"):
                cancel_order(kind, oid)
                self.load_orders(kind)
                self.refresh_order_combo(kind)
        except Exception as e:
            handle_error(e)

    def print_order(self, kind):
        oid = self.order_selected(kind)
        if oid is None:
            return
        name = "sales_order" if kind == "sales" else "purchase_order"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                            title="حفظ الأمر", initialfile=f"{name}_{oid}.pdf")
        if not path:
            return
        try:
            build_order_pdf(path, kind, oid)
            messagebox.showinfo("نجاح", "تم إنشاء ملف الأمر PDF.")
        except Exception as e:
            handle_error(e)

    def show_order_lines(self, kind):
        oid = self.order_selected(kind)
        if oid is None:
            return
        c = FLOW[kind]
        win = tk.Toplevel(self.root)
        win.title(f"تفاصيل الأمر رقم {oid}")
        win.geometry("900x340")
        box = tk.Frame(win)
        box.pack(fill="both", expand=True, padx=10, pady=10)
        tree = make_tree(box, [("name", "الصنف", 250, "e"), ("qty", "المطلوب", 80, "center"), ("done", "المنفّذ", 80, "center"),
                               ("rem", "المتبقي", 80, "center"), ("inv", "المفوتر", 80, "center"),
                               ("price", "السعر", 90, "center"), ("total", "الإجمالي", 110, "center")])
        with db() as conn:
            for name, q, done, inv, price, total in conn.execute(
                    f"SELECT item_name, quantity, {c['done']}, invoiced_qty, {c['oprice']}, total_amount FROM {c['olines']} "
                    f"WHERE order_id=? ORDER BY line_id", (oid,)):
                tree.insert("", "end", values=(name, f"{q:g}", f"{done:g}", f"{q - done:g}", f"{inv:g}", fmt(price), fmt(total)))

    # ---------------------------------------------------------- المرحلة 2: التسليم / الاستلام
    @staticmethod
    def _grid_cells(parent, row, widgets):
        n = len(widgets)
        for j, w in enumerate(widgets):
            w.grid(row=row, column=n - 1 - j, padx=6, pady=2, sticky="e")

    def build_fulfil_tab(self, parent, kind):
        T, st = self.FLOW_TEXT[kind], self.flow[kind]
        top = ttk.LabelFrame(parent, text=f" {T['ful_title']} ", padding=8)
        top.pack(fill="x", padx=8, pady=(8, 4))
        h1 = tk.Frame(top)
        h1.pack(fill="x")
        st["f_order_cb"] = SearchableCombobox(h1, width=64)
        st["f_order_cb"].bind("<<ComboboxSelected>>", lambda e: self.load_fulfil_lines(kind))
        rtl_row(h1, 0, [("اختر الأمر:", st["f_order_cb"]),
                        (None, money_button(h1, "📋 تحميل بنود الأمر", "#0284c7", lambda: self.load_fulfil_lines(kind), 9))], slots=2)
        st["f_grid"] = ScrollFrame(top, height=170)
        st["f_grid"].pack(fill="x", pady=4)
        st["f_rows"] = []
        h2 = tk.Frame(top)
        h2.pack(fill="x")
        st["f_date"] = date_entry(h2)
        st["f_notes"] = ttk.Entry(h2, width=34)
        items = [("التاريخ:", st["f_date"])]
        if kind == "purchase":
            st["f_ref"] = ttk.Entry(h2, width=20)
            items.append(("رقم إذن/مستند المورد:", st["f_ref"]))
        items.append(("ملاحظات:", st["f_notes"]))
        rtl_row(h2, 0, items, slots=3)
        create_key = f"{kind}.{'delivery' if kind == 'sales' else 'receipt'}_create"
        st["f_files"] = []
        if kind == "purchase" and self.can(create_key):
            att = tk.Frame(top)
            att.pack(fill="x", pady=(4, 0))
            money_button(att, "📎 إرفاق فاتورة المورد (صورة / PDF)", UI["teal"],
                         lambda: self.pick_receipt_files(), 9).pack(side="right", padx=4)
            st["f_files_lbl"] = ttk.Label(att, text="لم تُرفق فاتورة المورد بعد (اختياري: تُحفظ مع إذن الاستلام).",
                                          foreground=UI["muted"])
            st["f_files_lbl"].pack(side="right", padx=6)
        if self.can(create_key):
            money_button(top, T["ful_btn"], "#d97706", lambda: self.save_fulfilment(kind), 10).pack(pady=6)
        else:
            ttk.Label(top, text="👁 عرض فقط: التنفيذ والاعتماد غير متاحين لصلاحياتك.", foreground="#b45309").pack(pady=6)

        reg = ttk.LabelFrame(parent, text=f" {T['ful_reg']} ", padding=6)
        reg.pack(fill="both", expand=True, padx=8, pady=4)
        bar = tk.Frame(reg)
        bar.pack(fill="x", pady=(0, 4))
        money_button(bar, "🖨️ طباعة الإذن PDF", "#0284c7", lambda: self.print_fulfilment_selected(kind), 9).pack(side="right", padx=3)
        money_button(bar, "👁 عرض البنود", "#64748b", lambda: self.show_fulfilment_lines(kind), 9).pack(side="right", padx=3)
        if kind == "purchase":
            money_button(bar, "📎 مرفقات فاتورة المورد", UI["teal"], lambda: self.open_receipt_attachments(), 9
                         ).pack(side="right", padx=3)
        money_button(bar, "🧾 إصدار فاتورة من الإذن المحدد", "#16a34a",
                     lambda: self.goto_invoice_for(kind, self._selected_fulfilment(kind)), 9).pack(side="right", padx=3)
        holder = tk.Frame(reg)
        holder.pack(fill="both", expand=True)
        st["f_tree"] = make_tree(holder, [
            ("id", "رقم الإذن", 70, "center"), ("date", "التاريخ", 90, "center"), ("order", "رقم الأمر", 70, "center"),
            ("party", T["party"], 220, "e"), ("n", "عدد البنود", 70, "center"), ("val", "القيمة (قبل الضريبة)", 130, "center"),
            ("inv", "الفوترة", 110, "center")] + ([("att", "📎 فاتورة المورد", 100, "center")] if kind == "purchase" else []))
        if kind == "purchase":
            st["f_tree"].bind("<Double-1>", lambda e: self.open_receipt_attachments())

    def load_fulfil_lines(self, kind):
        st, c = self.flow[kind], FLOW[kind]
        inner = st["f_grid"].inner
        for w in inner.winfo_children():
            w.destroy()
        st["f_rows"] = []
        oid = st["f_order_cb"].selected()
        if not oid:
            return
        with db() as conn:
            rows = conn.execute(
                f"SELECT l.line_id, l.item_code, l.item_name, l.unit, l.quantity, l.{c['oprice']}, l.{c['done']}, "
                f"COALESCE(i.quantity,0) FROM {c['olines']} l LEFT JOIN inventory i ON i.item_code=l.item_code "
                f"WHERE l.order_id=? ORDER BY l.line_id", (oid,)).fetchall()
        sales = kind == "sales"
        heads = (["الصنف", "المطلوب", "المنفّذ", "المتبقي", "المتاح بالمخزن", "الكمية الآن"] if sales else
                 ["الصنف", "المطلوب", "المنفّذ", "المتبقي", "السعر 🔒 (من الأمر)", "تاريخ الإنتاج", "تاريخ الانتهاء",
                 "الكمية الآن"])
        self._grid_cells(inner, 0, [ttk.Label(inner, text=h, font=("Arial", 9, "bold")) for h in heads])
        for r_i, (lid, code, name, unit, q, price, done, stock) in enumerate(rows, start=1):
            rem = q - done
            qty_e = ttk.Entry(inner, width=10)
            price_e = prod_e = exp_e = None
            cells = [ttk.Label(inner, text=f"{name} ({unit or '-'})", width=34 if not sales else 38, anchor="e"),
                     ttk.Label(inner, text=f"{q:g}"), ttk.Label(inner, text=f"{done:g}"), ttk.Label(inner, text=f"{rem:g}")]
            if sales:
                cells.append(ttk.Label(inner, text=f"{stock:g}"))
                default = min(rem, max(stock, 0))
            else:
                # السعر مسحوب من أمر الشراء المعتمد ومقفول؛ الكمية فقط قابلة للتعديل (عجز/زيادة)
                sty = ttk.Style()
                sty.configure("Locked.TEntry", foreground="#334155")
                sty.map("Locked.TEntry", fieldbackground=[("readonly", "#e2e8f0")])
                price_e = ttk.Entry(inner, width=10, justify="center", style="Locked.TEntry")
                price_e.insert(0, f"{price:g}")
                price_e.configure(state="readonly")
                prod_e = DateEntry(inner, width=11)
                exp_e = DateEntry(inner, width=11)
                cells += [price_e, prod_e, exp_e]
                default = rem
            cells.append(qty_e)
            if rem <= EPS:
                qty_e.insert(0, "0")
                qty_e.configure(state="disabled")
            else:
                qty_e.insert(0, f"{default:g}")
            self._grid_cells(inner, r_i, cells)
            st["f_rows"].append(dict(line_id=lid, name=name, qty_e=qty_e, price_e=price_e, prod_e=prod_e, exp_e=exp_e))

    def save_fulfilment(self, kind):
        T, st = self.FLOW_TEXT[kind], self.flow[kind]
        try:
            oid = st["f_order_cb"].selected()
            if not oid:
                raise UserError("اختر الأمر أولاً.")
            f_date = parse_date(st["f_date"].get())
            if not f_date:
                raise UserError("صيغة التاريخ غير صحيحة، استخدم YYYY-MM-DD")
            entries = []
            for r in st["f_rows"]:
                q = parse_float(r["qty_e"].get() or "0")
                if q is None or q < 0:
                    raise UserError(f"كمية غير صحيحة للصنف {r['name']}")
                pr = None  # السعر يُسحب دائماً من أمر الشراء المعتمد
                prod = exp = None
                if r.get("prod_e") is not None and q > EPS:
                    prod, exp = self._read_batch_dates(r["prod_e"], r["exp_e"], r["name"])
                entries.append((r["line_id"], q, pr, prod, exp))
            self.check_period_open(f_date)
            ref = st["f_ref"].get().strip() if "f_ref" in st else ""
            fid = self.with_credit(fulfil_order, kind, oid, f_date, st["f_notes"].get().strip(), entries, ref)
            n_att = self.attach_receipt_files(fid) if kind == "purchase" else 0
            if kind == "sales":
                with db() as conn:
                    code = conn.execute("SELECT customer_code FROM sales_orders WHERE order_id=?", (oid,)).fetchone()[0]
                cst = _credit.credit_status(code, 0.0, "delivery")
                if cst["exceeded"]:
                    messagebox.showwarning("تنبيه الحد الائتماني", cst["message"])
            st["f_notes"].delete(0, tk.END)
            if "f_ref" in st:
                st["f_ref"].delete(0, tk.END)
            self.refresh_order_combo(kind)
            self.load_fulfil_lines(kind)
            self.load_fulfilments(kind)
            self.load_orders(kind)
            self.load_inventory()
            if self.can(self.FLOW_TAB_PERMS[kind]["i"]) and messagebox.askyesno(
                    "نجاح", f"تم حفظ {T['ful_noun']} رقم {fid} وتحديث المخزن."
                            + (f"\n📎 تم إرفاق {n_att} ملف لفاتورة المورد." if n_att else "") + "\n\n"
                            "هل تريد إصدار الفاتورة الآن من هذا الإذن؟ (تُعرض جاهزة للمراجعة والاعتماد)"):
                self.goto_invoice_for(kind, fid)
            elif messagebox.askyesno("طباعة", f"هل تريد طباعة {T['ful_noun']} رقم {fid} PDF الآن؟"):
                self.print_fulfilment(kind, fid)
        except Exception as e:
            handle_error(e)

    def load_fulfilments(self, kind):
        st, c = self.flow[kind], FLOW[kind]
        clear_tree(st["f_tree"])
        att = attachment_counts("receipt") if kind == "purchase" else {}
        with db() as conn:
            for fid, d, oid, party, n, val, inv in conn.execute(
                    f"SELECT f.{c['fid']}, f.{c['fdate']}, f.order_id, f.{c['pname']}, COUNT(l.line_id), "
                    f"COALESCE(SUM({c['fval']}),0), f.invoice_id FROM {c['ful']} f "
                    f"LEFT JOIN {c['flines']} l ON l.{c['fid']}=f.{c['fid']} GROUP BY f.{c['fid']} ORDER BY f.{c['fid']} DESC"):
                st["f_tree"].insert("", "end", iid=str(fid), values=(
                    fid, d, oid, party, n, fmt(val), f"فاتورة {inv}" if inv else "غير مفوتر")
                    + ((f"📎 {att[fid]}" if att.get(fid) else "—",) if kind == "purchase" else ()))

    def show_fulfilment_lines(self, kind):
        sel = self.flow[kind]["f_tree"].selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر إذناً من السجل.")
            return
        c = FLOW[kind]
        win = tk.Toplevel(self.root)
        win.title(f"بنود الإذن رقم {sel[0]}")
        win.geometry("760x300")
        box = tk.Frame(win)
        box.pack(fill="both", expand=True, padx=10, pady=10)
        tree = make_tree(box, [("name", "الصنف", 240, "e"), ("qty", "الكمية", 80, "center"), ("unit", "الوحدة", 70, "center"),
                               ("price", "سعر الوحدة", 95, "center"), ("gross", "الإجمالي قبل الخصم", 120, "center"),
                               ("disc", "الخصم الخاص", 95, "center"), ("val", "القيمة بعد الخصم", 120, "center")], rtl=True)
        with db() as conn:
            for name, unit, q, price, disc in conn.execute(
                    f"SELECT item_name, unit, quantity, {c['fprice']}, {c['fdisc']} FROM {c['flines']} WHERE {c['fid']}=? "
                    f"ORDER BY line_id", (int(sel[0]),)):
                tree.insert("", "end", values=(name, f"{q:g}", unit or "-", fmt(price), fmt(q * price),
                                               fmt(disc) if disc > 0.005 else "-", fmt(q * price - disc)))

    def print_fulfilment_selected(self, kind):
        sel = self.flow[kind]["f_tree"].selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر إذناً من السجل.")
            return
        self.print_fulfilment(kind, int(sel[0]))

    def print_fulfilment(self, kind, fid):
        name = "delivery_note" if kind == "sales" else "receipt_note"
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                            title="حفظ الإذن", initialfile=f"{name}_{fid}.pdf")
        if not path:
            return
        try:
            build_fulfilment_pdf(path, kind, fid)
            messagebox.showinfo("نجاح", "تم إنشاء الإذن PDF.")
        except Exception as e:
            handle_error(e)

    # ---------------------------------------------------------- المرحلة 3: الفواتير
    def build_flow_invoices_tab(self, parent, kind):
        st = self.flow[kind]
        top = ttk.LabelFrame(parent, text=" 📋 غير المفوتر: كل أذون التسليم والاستلام المفتوحة (مبيعات + مشتريات) - الأقدم أولاً ",
                             padding=8)
        top.pack(fill="x", padx=8, pady=(8, 4))
        hdr = tk.Frame(top)
        hdr.pack(fill="x")
        st["u_search"] = ttk.Entry(hdr, width=28)
        st["u_search"].bind("<KeyRelease>", lambda e: self.load_unbilled(kind))
        st["u_type"] = ttk.Combobox(hdr, values=list(self.UNBILLED_TYPES), width=12, state="readonly")
        st["u_type"].set("الكل")
        st["u_type"].bind("<<ComboboxSelected>>", lambda e: self.load_unbilled(kind))
        st["i_date"] = date_entry(hdr)
        rtl_row(hdr, 0, [("بحث (طرف / رقم إذن / أمر):", st["u_search"]), ("النوع:", st["u_type"]),
                         ("تاريخ الفاتورة:", st["i_date"]),
                         (None, money_button(hdr, "🔄 تحديث", "#0284c7", lambda: self.load_unbilled(kind), 9))], slots=4)
        st["i_notes"] = ttk.Entry(hdr, width=40)
        items = [("ملاحظات الفاتورة:", st["i_notes"])]
        if kind == "purchase":
            st["i_ref"] = ttk.Entry(hdr, width=22)
            items.insert(0, ("رقم فاتورة المورد:", st["i_ref"]))
        rtl_row(hdr, 1, items, slots=4)
        holder = tk.Frame(top)
        holder.pack(fill="x")
        st["u_tree"] = make_tree(holder, [
            ("kind", "النوع", 110, "center"), ("id", "رقم الإذن", 70, "center"), ("date", "التاريخ", 88, "center"),
            ("age", "منذ (يوم)", 70, "center"), ("order", "رقم الأمر", 70, "center"), ("party", "الطرف", 230, "e"),
            ("n", "البنود", 55, "center"), ("sub", "قبل الضريبة", 100, "center"), ("vat", "الضريبة", 85, "center"),
            ("total", "الإجمالي", 100, "center")], height=6, rtl=True)
        st["u_tree"].tag_configure("aging", background="#ffedd5")
        st["u_tree"].tag_configure("old", background="#fee2e2")
        st["u_count_lbl"] = tk.Label(top, text="", fg=UI["muted"], font=(UI_FONT, 9))
        st["u_count_lbl"].pack(anchor="e")
        tk.Label(top, text="حدد إذناً أو أكثر لنفس الطرف ونفس النوع (Ctrl للتحديد المتعدد) ثم أصدر الفاتورة الضريبية.",
                 fg="#78716c").pack(anchor="e")
        money_button(top, "🧾 إصدار الفاتورة من الأذون المحددة", "#16a34a", lambda: self.make_flow_invoice(kind), 10).pack(pady=6)

        reg = ttk.LabelFrame(parent, text=" سجل الفواتير (اضغط مرتين لعرض الفاتورة وبنودها في نافذة) ", padding=6)
        reg.pack(fill="both", expand=True, padx=8, pady=4)
        act = tk.Frame(reg)
        act.pack(fill="x", pady=(0, 4))
        if self.can(f"{kind}.discount"):
            money_button(act, "🏷️ تسجيل خصم لاحق / إشعار دائن للفاتورة المحددة", UI["teal"],
                         lambda: self.open_discount_for_invoice(kind), 10).pack(side="right", padx=3)
        money_button(act, "🖨️ طباعة الفاتورة PDF", UI["primary"], lambda: self.print_invoice_selected(kind), 10
                     ).pack(side="right", padx=3)
        money_button(act, "🔎 تفاصيل الفاتورة", "#475569", lambda: self.open_invoice_modal(kind), 10).pack(side="right", padx=3)
        if kind == "purchase":
            money_button(act, "📎 مرفقات فاتورة المورد", UI["teal"], self.open_purchase_attachments_selected, 10
                         ).pack(side="right", padx=3)
        if self.can_modify():
            adm = tk.Frame(reg)
            adm.pack(fill="x", pady=(0, 4))
            money_button(adm, "✏️ تعديل الفاتورة المحددة", "#d97706", lambda: self.open_invoice_editor(kind), 9
                         ).pack(side="right", padx=3)
            money_button(adm, "🗑️ حذف الفاتورة المحددة", "#dc2626", lambda: self.delete_invoice_selected(kind), 9
                         ).pack(side="right", padx=3)
            tk.Label(adm, text="صلاحية المدير: التعديل والحذف يعيدان احتساب المخزون والتكلفة تلقائياً.",
                     fg="#78716c").pack(side="left", padx=6)
        holder2 = tk.Frame(reg)
        holder2.pack(fill="both", expand=True)
        if kind == "sales":
            self.sal_tree = make_tree(holder2, [
                ("id", "رقم الفاتورة", 95, "center"), ("date", "التاريخ", 90, "center"), ("cust", "العميل", 210, "e"),
                ("src", "المصدر", 70, "center"), ("rep", "المندوب", 90, "center"), ("n", "عدد الأصناف", 70, "center"),
                ("disc", "الخصم الخاص", 90, "center"), ("subtotal", "المبلغ الأساسي", 105, "center"),
                ("vat", "الضريبة", 85, "center"), ("total", "الإجمالي", 105, "center"), ("paid", "المسدد", 95, "center"),
                ("status", "الحالة", 75, "center")])
            self.sal_tree.bind("<Double-1>", lambda e: self.open_invoice_modal("sales"))
        else:
            self.pur_tree = make_tree(holder2, [
                ("id", "رقم الفاتورة", 90, "center"), ("date", "التاريخ", 86, "center"), ("supp", "المورد", 165, "e"),
                ("src", "المصدر", 70, "center"), ("ref", "مرجع المورد", 100, "center"), ("n", "عدد الأصناف", 70, "center"),
                ("subtotal", "المبلغ الأساسي", 105, "center"), ("vat", "الضريبة", 85, "center"),
                ("total", "الإجمالي", 105, "center"), ("paid", "المسدد", 95, "center"), ("status", "الحالة", 75, "center"),
                ("att", "📎 المرفقات", 80, "center")])
            self.pur_tree.bind("<Double-1>", self._pur_tree_double)
            tk.Label(reg, text="اضغط مرتين على الفاتورة لعرضها، أو على عمود «📎 المرفقات» لفتح ملف فاتورة المورد الأصلية.",
                     fg=UI["muted"], font=(UI_FONT, 8)).pack(anchor="e")

    def _pur_tree_double(self, event=None):
        """الضغط المزدوج على عمود المرفقات يفتحها مباشرة، وعلى باقي الأعمدة يفتح تفاصيل الفاتورة."""
        tree = self.pur_tree
        if event is not None:
            row, col = tree.identify_row(event.y), tree.identify_column(event.x)
            if row and col:
                cols = list(tree["displaycolumns"]) if tree["displaycolumns"] not in ("#all", ("#all",)) else list(tree["columns"])
                try:
                    name = cols[int(col.replace("#", "")) - 1]
                except (ValueError, IndexError):
                    name = ""
                if name == "att":
                    tree.selection_set(row)
                    self.open_attachments_window("purchase", int(row))
                    return "break"
        self.open_invoice_modal("purchase")

    def _selected_invoice(self, kind):
        tree = self.sal_tree if kind == "sales" else self.pur_tree
        sel = tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر فاتورة من السجل أولاً.")
            return None
        return int(sel[0])

    def _after_invoice_change(self, kind):
        self.load_sales_invoices() if kind == "sales" else self.load_purchase_invoices()
        self.load_orders(kind)
        self.load_fulfilments(kind)
        self.load_inventory()

    def delete_invoice_selected(self, kind):
        if not self.require_modify():
            return
        inv_id = self._selected_invoice(kind)
        if inv_id is None:
            return
        no = invoice_number(kind, inv_id)
        if not messagebox.askyesno("حذف فاتورة", f"حذف الفاتورة {no} نهائياً؟\n\n"
                                                  "• البنود المباشرة تُعكس من المخزون مع إعادة احتساب التكلفة\n"
                                                  "• أذون التسليم/الاستلام المرتبطة تعود «غير مفوترة»\n"
                                                  "• السندات المربوطة تبقى دفعات على حساب الطرف", icon="warning"):
            return
        try:
            self.with_stock_override(delete_invoice, kind, inv_id)
            self._after_invoice_change(kind)
            messagebox.showinfo("تم", f"تم حذف الفاتورة {no}.")
        except Exception as e:
            handle_error(e)

    def open_invoice_editor(self, kind):
        """محرر الفاتورة للمدير: التاريخ والطرف والملاحظات والبنود (إضافة/تعديل/حذف) والأسعار والخصم والتواريخ."""
        if not self.require_modify():
            return
        inv_id = self._selected_invoice(kind)
        if inv_id is None:
            return
        try:
            header, lines = load_invoice(kind, inv_id)
        except Exception as e:
            handle_error(e)
            return
        sales = kind == "sales"
        linked_any = any(l["linked"] for l in lines)
        state = {"lines": [dict(l) for l in lines], "editing": None}
        win = tk.Toplevel(self.root)
        win.title(f"تعديل الفاتورة {header['invoice_no']}")
        win.geometry("1180x640")
        win.grab_set()

        top = ttk.LabelFrame(win, text=f" بيانات الفاتورة {header['invoice_no']} ", padding=8)
        top.pack(fill="x", padx=10, pady=(10, 4))
        date_e = date_entry(top, 13, header["date"])
        party_cb = SearchableCombobox(top, width=34)
        party_cb.set_items(self.partner_pairs("customer" if sales else "supplier"))
        party_disp = f"{header['party'][0]} - {header['party'][1]}"
        party_cb.set(party_disp)
        notes_e = ttk.Entry(top, width=30)
        notes_e.insert(0, header["notes"])
        items = [("التاريخ:", date_e), ("العميل:" if sales else "المورد:", party_cb), ("ملاحظات:", notes_e)]
        ref_e = None
        if not sales:
            ref_e = ttk.Entry(top, width=16)
            ref_e.insert(0, header["supplier_ref"])
            items.append(("مرجع المورد:", ref_e))
        rtl_row(top, 0, items)
        if linked_any:
            party_cb.configure(state="disabled")
            ttk.Label(top, text="فاتورة صادرة من أذون تسليم/استلام: يمكن تعديل الأسعار والخصم والتواريخ، "
                                "أما الطرف والأصناف والكميات فمرتبطة بالأذون (لتغييرها احذف الفاتورة وأعد إصدارها).",
                      foreground="#b45309").grid(row=1, column=0, columnspan=8, sticky="e")

        box = ttk.LabelFrame(win, text=" البنود ", padding=6)
        box.pack(fill="both", expand=True, padx=10, pady=4)
        cols = [("code", "الكود", 75, "center"), ("name", "الصنف", 200, "e"), ("qty", "الكمية", 65, "center"),
                ("unit", "الوحدة", 55, "center"), ("price", "سعر الوحدة", 80, "center")]
        cols += ([("gross", "قبل الخصم", 90, "center"), ("discount", "الخصم الخاص", 85, "center")] if sales else [])
        cols += [("subtotal", "الخاضع للضريبة", 105, "center"), ("vat", "الضريبة", 80, "center"), ("total", "الإجمالي", 95, "center")]
        cols += ([] if sales else [("production", "تاريخ الإنتاج", 90, "center"), ("expiry", "تاريخ الانتهاء", 90, "center")])
        cols += [("src", "المصدر", 65, "center")]
        tree = make_tree(box, cols, height=9, rtl=True)
        total_lbl = tk.Label(win, text="", font=("Arial", 10, "bold"), bg="#e2e8f0", pady=5)
        total_lbl.pack(fill="x", padx=10)

        def refresh():
            clear_tree(tree)
            sub = vat = 0.0
            for i, l in enumerate(state["lines"]):
                L = make_line(l["code"], l["name"], l.get("unit", ""), l["qty"], l["price"], l.get("discount", 0.0) if sales else 0.0)
                sub += L["subtotal"]
                vat += L["vat"]
                vals = [l["code"], l["name"], f"{L['qty']:g}", l.get("unit") or "-", fmt(L["price"])]
                vals += ([fmt(L["gross"]), fmt(L["discount"]) if L["discount"] else "-"] if sales else [])
                vals += [fmt(L["subtotal"]), fmt(L["vat"]), fmt(L["total"])]
                vals += ([] if sales else [l.get("production") or "-", l.get("expiry") or "-"])
                vals += ["إذن" if l.get("linked") else "مباشر"]
                tree.insert("", "end", iid=str(i), values=vals)
            total_lbl.config(text=f"الأساسي: {fmt(sub)}   |   الضريبة: {fmt(vat)}   |   الإجمالي الجديد: {fmt(sub + vat)} ر.س")

        ed = ttk.LabelFrame(win, text=" تعديل / إضافة بند ", padding=8)
        ed.pack(fill="x", padx=10, pady=4)
        item_cb = SearchableCombobox(ed, width=34)
        item_cb.set_items(self.item_pairs())
        qty_e, price_e = ttk.Entry(ed, width=10), ttk.Entry(ed, width=12)
        disc_e, prod_e, exp_e = ttk.Entry(ed, width=10), DateEntry(ed, width=12), DateEntry(ed, width=12)
        row_items = [("الصنف:", item_cb), ("الكمية:", qty_e), ("السعر:", price_e)]
        if sales:
            row_items.append(("الخصم الخاص (مبلغ أو %):", disc_e))
        rtl_row(ed, 0, row_items)
        if not sales:
            rtl_row(ed, 1, [("تاريخ الإنتاج:", prod_e), ("تاريخ الانتهاء:", exp_e)])

        def fill(l):
            item_cb.configure(state="normal")
            qty_e.configure(state="normal")
            item_cb.set(f"{l['code']} - {l['name']}")
            for e, v in ((qty_e, f"{l['qty']:g}"), (price_e, f"{l['price']:g}"), (disc_e, f"{l.get('discount') or 0:g}"),
                         (prod_e, l.get("production") or ""), (exp_e, l.get("expiry") or "")):
                e.delete(0, tk.END)
                e.insert(0, v)
            if l.get("linked"):
                item_cb.configure(state="disabled")
                qty_e.configure(state="disabled")

        def read(existing=None):
            item = item_cb.selected()
            if not item and existing and item_cb.get().strip() == f"{existing['code']} - {existing['name']}":
                item = (existing["code"], existing["name"], existing.get("unit", ""))
            if not item:
                raise UserError("اختر الصنف من القائمة.")
            qty, price = parse_float(qty_e.get()), parse_float(price_e.get())
            if qty is None or qty <= 0 or price is None or price < 0:
                raise UserError("أدخل كمية أكبر من الصفر وسعراً صحيحاً.")
            out = dict(code=item[0], name=item[1], unit=item[2] or "", qty=qty, price=price)
            if sales:
                out["discount"] = parse_discount(disc_e.get(), round(qty * price, 2))
                make_line(item[0], item[1], item[2], qty, price, out["discount"])  # تحقق الخصم
            else:
                out["production"], out["expiry"] = self._read_batch_dates(prod_e, exp_e)
            return out

        def load_sel():
            sel = tree.selection()
            if not sel:
                messagebox.showinfo("تنبيه", "اختر بنداً من الجدول.", parent=win)
                return
            state["editing"] = int(sel[0])
            fill(state["lines"][state["editing"]])

        def update_line():
            try:
                idx = state["editing"]
                if idx is None:
                    raise UserError("حمّل البند المحدد أولاً للتعديل.")
                l = state["lines"][idx]
                new = read(l)
                if l.get("linked"):
                    new.update(code=l["code"], name=l["name"], unit=l.get("unit", ""), qty=l["qty"])
                l.update(new)
                refresh()
            except Exception as ex:
                handle_error(ex)

        def add_line():
            try:
                new = read()
                new.update(line_id=None, linked=False)
                state["lines"].append(new)
                refresh()
            except Exception as ex:
                handle_error(ex)

        def remove_line():
            sel = tree.selection()
            if not sel:
                return
            idx = int(sel[0])
            if state["lines"][idx].get("linked"):
                messagebox.showwarning("تنبيه", "بنود الأذون لا تُحذف منفردة.", parent=win)
                return
            del state["lines"][idx]
            state["editing"] = None
            refresh()

        def save():
            try:
                d = parse_date(date_e.get())
                if not d:
                    raise UserError("تاريخ الفاتورة غير صحيح (YYYY-MM-DD).")
                party = party_cb.selected()
                if not party and party_cb.get().strip() == party_disp:
                    party = header["party"]
                if not party:
                    raise UserError("اختر الطرف من القائمة.")
                hdr = dict(date=d, party=party, notes=notes_e.get().strip(),
                           supplier_ref=ref_e.get().strip() if ref_e else "")
                total = self.with_stock_override(update_invoice, kind, inv_id, hdr, state["lines"])
                win.destroy()
                self._after_invoice_change(kind)
                messagebox.showinfo("تم", f"تم حفظ تعديلات الفاتورة {header['invoice_no']}. الإجمالي الجديد: {fmt(total)} ر.س")
            except Exception as ex:
                handle_error(ex)

        bar = tk.Frame(ed)
        bar.grid(row=2, column=0, columnspan=8, pady=(6, 0))
        for text, color, cmd in (("⬇️ تحميل البند المحدد", "#0284c7", load_sel), ("✔ تحديث البند", "#d97706", update_line),
                                 ("➕ إضافة كبند جديد", "#16a34a", add_line), ("➖ حذف البند", "#dc2626", remove_line)):
            money_button(bar, text, color, cmd, 9).pack(side="right", padx=3)
        money_button(win, "💾 حفظ تعديلات الفاتورة", "#16a34a", save, 11).pack(pady=8)
        tree.bind("<Double-1>", lambda e: load_sel())
        refresh()

    def load_sales_invoices(self):
        clear_tree(self.sal_tree)
        with db() as conn:
            for r in conn.execute(
                    "SELECT h.invoice_id, COALESCE(h.invoice_no, h.invoice_id), h.invoice_date, h.customer_name, "
                    "COALESCE(h.source_type,'مباشر'), COALESCE(h.sales_rep,''), "
                    "(SELECT COUNT(*) FROM sales_invoice_lines l WHERE l.invoice_id=h.invoice_id), "
                    "(SELECT COALESCE(SUM(l.discount_amount),0) FROM sales_invoice_lines l WHERE l.invoice_id=h.invoice_id), "
                    "h.subtotal, h.vat_amount, h.total_amount, h.paid_amount, h.status FROM sales_invoices h "
                    "ORDER BY h.invoice_id DESC"):
                self.sal_tree.insert("", "end", iid=str(r[0]), values=(
                    r[1], r[2], r[3], r[4], r[5] or "-", r[6], fmt(r[7]) if r[7] > 0.005 else "-", fmt(r[8]), fmt(r[9]),
                    fmt(r[10]), fmt(r[11]), r[12]))

    def load_purchase_invoices(self):
        clear_tree(self.pur_tree)
        att = attachment_counts("purchase")
        with db() as conn:
            for r in conn.execute(
                    "SELECT h.purchase_id, COALESCE(h.invoice_no, h.purchase_id), h.purchase_date, h.supplier_name, "
                    "COALESCE(h.source_type,'مباشر'), COALESCE(h.supplier_ref,''), "
                    "(SELECT COUNT(*) FROM purchase_invoice_lines l WHERE l.purchase_id=h.purchase_id), h.subtotal, h.vat_amount, "
                    "h.total_amount, h.paid_amount, h.status FROM purchase_invoices h ORDER BY h.purchase_id DESC"):
                n_att = att.get(r[0], 0)
                self.pur_tree.insert("", "end", iid=str(r[0]), values=(
                    r[1], r[2], r[3], r[4], r[5] or "-", r[6], fmt(r[7]), fmt(r[8]), fmt(r[9]), fmt(r[10]), r[11],
                    f"📎 {n_att}" if n_att else "-"))

    # ============================================================
    # الاستيراد من الإكسيل + الرصيد الافتتاحي
    # ============================================================
    FILE_TYPES = [("Excel / CSV", "*.xlsx *.xls *.csv"), ("Excel", "*.xlsx *.xls"), ("CSV", "*.csv")]

    def import_partners_excel(self, role):
        label = "العملاء" if role == "customer" else "الموردين"
        path = filedialog.askopenfilename(title=f"اختر ملف {label}", filetypes=self.FILE_TYPES)
        if not path:
            return
        try:
            df = read_table_file(path)
            pre = import_partners_from_df(df, role, True, dry_run=True)
            merge = True
            if pre["merged"]:
                other = "موردين" if role == "customer" else "عملاء"
                merge = messagebox.askyesno(
                    "أسماء موجودة في القائمة الأخرى",
                    f"{pre['merged']} اسم في هذا الملف موجود بالفعل في قائمة ال{other}.\n\n"
                    f"نعم = اعتبارهم {label} أيضاً (شريك واحد: عميل ومورد معاً).\n"
                    f"لا = تجاهلهم، وإضافة الأسماء الجديدة فقط.\n\n"
                    f"تنبيه: لو الملف نسخة من نفس قائمة الأسماء بدون تمييز حقيقي بين العميل والمورد فاختر «لا».")
            stats = import_partners_from_df(df, role, merge)
            self.load_entities()
            messagebox.showinfo("تم الاستيراد", (
                f"أسماء جديدة أُضيفت: {stats['added']}\n"
                f"دُمجت مع شركاء موجودين (أُضيفت صفة {label[:-1] if label.endswith('ء') else label}): {stats['merged']}\n"
                f"موجودون مسبقاً بنفس الصفة: {stats['existing_same_role']}\n"
                f"تم تجاهلهم: {stats['skipped_other_role']}\n"
                f"أسماء مكررة داخل الملف: {stats['duplicates']}\n"
                f"صفوف فارغة: {stats['empty']}\n\n"
                f"الأكواد تؤخذ من عمود الكود (No) إن وُجد، وإلا تُرقَّم تلقائياً."))
        except Exception as e:
            handle_error(e)

    def import_partner_openings(self, role):
        """رفع الأرصدة الافتتاحية للعملاء أو الموردين (ما عليه / ما له) من Excel أو CSV."""
        if not self.require_modify():
            return
        label = "العملاء" if role == "customer" else "الموردين"
        path = filedialog.askopenfilename(title=f"ملف الأرصدة الافتتاحية - {label}", filetypes=self.FILE_TYPES)
        if not path:
            return
        try:
            as_of = parse_date(simpledialog.askstring("تاريخ الرصيد الافتتاحي", "تاريخ الأرصدة (YYYY-MM-DD):",
                                                      initialvalue=f"{date.today().year}-01-01") or "")
            if not as_of:
                raise UserError("تاريخ غير صحيح. تم الإلغاء.")
            df = read_table_file(path)
            pre, problems = import_partner_openings_from_df(df, role, as_of, dry_run=True)
            msg = (f"معاينة رفع أرصدة {label} بتاريخ {as_of}:\n\n"
                   f"شركاء موجودون سيُحدَّث رصيدهم: {pre['updated']}\nشركاء جدد سيُضافون: {pre['created']}\n"
                   f"إجمالي المدين (عليه): {fmt(pre['total_debit'])}\nإجمالي الدائن (له): {fmt(pre['total_credit'])}\n"
                   f"صفوف بها أخطاء (ستُتجاهل): {pre['bad']}\n\n"
                   f"الرصيد الجديد يحل محل أي رصيد افتتاحي سابق للشريك. متابعة؟")
            if problems:
                msg += "\n\n" + "\n".join(problems[:8])
            if not messagebox.askyesno("تأكيد الرفع", msg):
                return
            stats, _ = import_partner_openings_from_df(df, role, as_of)
            self.load_entities()
            messagebox.showinfo("تم", f"تم تحديث {stats['updated']} وإضافة {stats['created']} شريك بأرصدتهم الافتتاحية.")
        except Exception as e:
            handle_error(e)

    def download_template(self, kind):
        names = {"customer_openings": "نموذج_أرصدة_افتتاحية", "supplier_openings": "نموذج_أرصدة_موردين",
                 "items": "نموذج_رفع_الأصناف"}
        path = filedialog.asksaveasfilename(title="حفظ النموذج", defaultextension=".xlsx", initialfile=names[kind],
                                            filetypes=[("Excel", "*.xlsx"), ("CSV", "*.csv")])
        if not path:
            return
        try:
            write_upload_template(path, kind)
            messagebox.showinfo("تم", f"تم حفظ النموذج:\n{path}")
        except Exception as e:
            handle_error(e)

    def import_items_excel(self):
        path = filedialog.askopenfilename(title="اختر ملف الأصناف / أرصدة المخزون", filetypes=self.FILE_TYPES)
        if not path:
            return
        try:
            df = read_table_file(path)
            update_opening = False
            pre = import_items_from_df(df, dry_run=True, update_opening=True)
            if pre["opening_updated"]:
                if self.can_modify():
                    update_opening = messagebox.askyesno(
                        "أرصدة افتتاحية للمخزون",
                        f"الملف يحتوي على كميات/تكاليف لـ {pre['opening_updated']} صنف موجود بالفعل.\n\n"
                        f"نعم = تحديث الرصيد والتكلفة الافتتاحية لهذه الأصناف (مع إعادة احتساب أرصدتها).\n"
                        f"لا = إضافة الأصناف الجديدة فقط.")
                else:
                    messagebox.showinfo("تنبيه", "تحديث الأرصدة الافتتاحية للأصناف الموجودة يحتاج صلاحية المدير؛ "
                                                 "سيتم إضافة الأصناف الجديدة فقط.")
            stats = import_items_from_df(df, update_opening=update_opening)
            self.load_inventory()
            messagebox.showinfo("تم الاستيراد", (
                f"أصناف جديدة أُضيفت: {stats['added']} (منها {stats['auto_codes']} بترقيم تلقائي)\n"
                f"أرصدة افتتاحية حُدِّثت: {stats['opening_updated']}\n"
                f"موجودة مسبقاً (لم تتغير): {stats['existing']}\n"
                f"أكواد مكررة داخل الملف: {stats['duplicates']}\nصفوف ناقصة: {stats['empty']}\n\n"
                f"الأعمدة المدعومة: كود، صنف، الوحدة، الكمية، التكلفة (Excel أو CSV)."))
        except Exception as e:
            handle_error(e)

    def edit_opening_balance(self):
        sel = self.inv_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر صنفاً من الجدول أولاً.")
            return
        code = sel[0]
        try:
            with db() as conn:
                r = conn.execute("SELECT item_name, unit, opening_qty, opening_cost FROM inventory WHERE item_code=?",
                                 (code,)).fetchone()
                moves = item_has_movements(conn, code)
            if moves and not self.can_modify():
                raise UserError("لا يمكن تعديل الرصيد الافتتاحي لصنف عليه حركات (فواتير/تسليم/استلام/مرتجعات).\n"
                                "هذا التعديل متاح للمدير فقط (مع إعادة احتساب الأرصدة).")
            qty = simpledialog.askfloat("رصيد افتتاحي", f"الكمية الافتتاحية للصنف:\n{r[0]}", minvalue=0.0,
                                        initialvalue=r[2] or 0.0)
            if qty is None:
                return
            cost = simpledialog.askfloat("تكلفة الوحدة", f"تكلفة الوحدة الافتتاحية للصنف:\n{r[0]}", minvalue=0.0,
                                         initialvalue=r[3] or 0.0)
            if cost is None:
                return
            self.with_stock_override(update_item, code, code, r[0], r[1], qty, cost)
            self.load_inventory()
        except Exception as e:
            handle_error(e)

    def open_cash_opening_dialog(self):
        """الأرصدة الافتتاحية للخزينة والبنوك (قيد واحد مقابل ح/ وسيط افتتاحي)."""
        if not self.require_modify():
            return
        as_of, current = get_cash_opening()
        with db() as conn:
            names = cash_account_names(conn)
            codes = cash_accounts(conn)
        win = tk.Toplevel(self.root)
        win.title("الأرصدة الافتتاحية للخزينة والبنوك")
        win.geometry(f"560x{220 + 34 * len(names)}")
        win.grab_set()
        f = ttk.Frame(win, padding=16)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text="تاريخ الأرصدة الافتتاحية:").grid(row=0, column=1, sticky="e", pady=6)
        date_e = date_entry(f, 14, as_of or f"{date.today().year}-01-01")
        date_e.grid(row=0, column=0, sticky="w", pady=6)
        entries = {}
        for i, n in enumerate(names, start=1):
            ttk.Label(f, text=f"{n} ({codes.get(n, '-')}):").grid(row=i, column=1, sticky="e", pady=3)
            e = ttk.Entry(f, width=18)
            if current.get(n):
                e.insert(0, f"{current[n]:g}")
            e.grid(row=i, column=0, sticky="w", pady=3)
            entries[n] = e
        ttk.Label(f, text="يُسجَّل كقيد يومية واحد: من ح/ الخزينة والبنوك إلى ح/ وسيط افتتاحي (99999).\n"
                          "إعادة الحفظ تستبدل القيد السابق. الرصيد السالب = سحب على المكشوف.\n"
                          "تنبيه: لو سجّلت رأس المال أو الأرصدة بقيد يدوي سابق فلا تُدخلها هنا مرة أخرى.",
                  foreground="#78716c", justify="right").grid(row=len(names) + 1, column=0, columnspan=2, sticky="e", pady=8)

        def save():
            try:
                d = parse_date(date_e.get())
                if not d:
                    raise UserError("تاريخ غير صحيح (YYYY-MM-DD).")
                amounts = {}
                for n, e in entries.items():
                    v = parse_float(e.get() or "0")
                    if v is None:
                        raise UserError(f"قيمة غير صحيحة للحساب {n}.")
                    amounts[n] = v
                eid = set_cash_opening(d, amounts)
                win.destroy()
                self.load_cash_transactions()
                self.load_journal_register()
                messagebox.showinfo("تم", "تم حفظ الأرصدة الافتتاحية للخزينة والبنوك" +
                                    (f" (قيد رقم {eid})." if eid else " (تم حذف القيد لأن كل الأرصدة صفر)."))
            except Exception as ex:
                handle_error(ex)

        money_button(f, "حفظ الأرصدة الافتتاحية 💾", "#16a34a", save, 10).grid(row=len(names) + 2, column=0, columnspan=2, pady=6)

    # ============================================================
    # الأمان
    # ============================================================
    def save_admin_password(self):
        pwd = self.admin_pwd_e.get().strip()
        if len(pwd) < 4:
            messagebox.showwarning("تنبيه", "كلمة المرور 4 أحرف على الأقل.")
            return
        set_setting("admin_password", pwd)
        messagebox.showinfo("نجاح", "تم حفظ كلمة مرور المدير.")

    # ============================================================
    # الصلاحية والأصناف الراكدة
    # ============================================================
    def setup_expiry_tab(self, frame):
        bg = "#f1f5f9"
        page = tk.Frame(frame, bg=bg)
        page.pack(fill="both", expand=True)
        top = tk.Frame(page, bg=bg)
        top.pack(fill="x", padx=16, pady=(14, 6))
        tk.Label(top, text="⏳ الصلاحية والأصناف الراكدة", font=("Arial", 16, "bold"), bg=bg, fg=CHART_INK).pack(side="right")
        self.exp_alert_e = ttk.Entry(top, width=6)
        self.exp_alert_e.insert(0, "30")
        self.exp_stale_e = ttk.Entry(top, width=6)
        self.exp_stale_e.insert(0, str(STALE_DAYS_DEFAULT))
        tk.Button(top, text="🔄 تحديث", font=("Arial", 9, "bold"), bg=CHART_INK, fg="white", bd=0, padx=14, pady=6,
                  cursor="hand2", command=self.load_expiry_page).pack(side="left")
        for label, e in (("تنبيه قبل الانتهاء (يوم):", self.exp_alert_e), ("الركود بعد (يوم بلا حركة):", self.exp_stale_e)):
            e.pack(side="left", padx=(12, 2))
            tk.Label(top, text=label, bg=bg, fg="#334155").pack(side="left")

        cards = tk.Frame(page, bg=bg)
        cards.pack(fill="x", padx=10)
        self.exp_cards = {}
        for i, (key, title, color) in enumerate((("expired", "⛔ منتهية الصلاحية", "#dc2626"), ("near", "⚠️ قريبة الانتهاء", "#ea580c"),
                                                 ("mid", "🕒 تنتهي خلال 90 يوماً", "#f59e0b"), ("stagnant", "🐢 أصناف راكدة", "#7c3aed"))):
            card = tk.Frame(cards, bg="white", highlightbackground="#e2e8f0", highlightthickness=1)
            card.grid(row=0, column=3 - i, padx=6, pady=6, sticky="nsew")
            tk.Frame(card, bg=color, width=5).pack(side="right", fill="y")
            body = tk.Frame(card, bg="white")
            body.pack(side="right", fill="both", expand=True, padx=14, pady=10)
            tk.Label(body, text=title, bg="white", fg=CHART_TEXT, font=("Arial", 9, "bold"), anchor="e").pack(fill="x")
            lbl = tk.Label(body, text="-", bg="white", fg=color, font=("Arial", 14, "bold"), anchor="e", justify="right")
            lbl.pack(fill="x", pady=(4, 0))
            self.exp_cards[key] = lbl
            cards.grid_columnconfigure(i, weight=1, uniform="exp_cards")

        b1 = ttk.LabelFrame(page, text=" دفعات المخزون بتاريخ صلاحية (الأقرب انتهاءً أولاً - الكمية المتبقية بافتراض صرف الأقدم أولاً) ",
                            padding=6)
        b1.pack(fill="both", expand=True, padx=16, pady=4)
        self.exp_tree = make_tree(b1, [
            ("code", "الكود", 80, "center"), ("name", "الصنف", 220, "e"), ("ref", "الدفعة", 130, "center"),
            ("party", "المورد", 150, "e"), ("prod", "تاريخ الإنتاج", 95, "center"), ("exp", "تاريخ الانتهاء", 95, "center"),
            ("days", "الأيام المتبقية", 95, "center"), ("qty", "الكمية المتبقية", 95, "center"),
            ("value", "القيمة بالتكلفة", 110, "center"), ("status", "الحالة", 110, "center")], height=8)
        for tag, color in (("expired", "#fee2e2"), ("near", "#ffedd5"), ("mid", "#fef9c3")):
            self.exp_tree.tag_configure(tag, background=color)
        b2 = ttk.LabelFrame(page, text=" الأصناف الراكدة (لها رصيد وبلا أي حركة شراء أو بيع أو مرتجع لأكثر من المدة) ", padding=6)
        b2.pack(fill="both", expand=True, padx=16, pady=(4, 12))
        self.stag_tree = make_tree(b2, [
            ("code", "الكود", 80, "center"), ("name", "الصنف", 260, "e"), ("unit", "الوحدة", 60, "center"),
            ("qty", "الرصيد", 90, "center"), ("cost", "متوسط التكلفة", 100, "center"), ("value", "القيمة المحجوزة", 120, "center"),
            ("last", "آخر حركة", 150, "center"), ("days", "أيام بلا حركة", 100, "center")], height=7)

    def load_expiry_page(self):
        try:
            alert = int(parse_float(self.exp_alert_e.get() or "30") or 30)
            stale = int(parse_float(self.exp_stale_e.get() or str(STALE_DAYS_DEFAULT)) or STALE_DAYS_DEFAULT)
            s = fetch_expiry_summary(alert, stale)
        except Exception as e:
            handle_error(e)
            return
        self.exp_cards["expired"].config(text=f"{len(s['expired'])} دفعة • {fmt(s['expired_value'])} ر.س")
        self.exp_cards["near"].config(text=f"{len(s['near'])} دفعة • {fmt(s['near_value'])} ر.س")
        self.exp_cards["mid"].config(text=f"{len(s['mid'])} دفعة • {fmt(s['mid_value'])} ر.س")
        self.exp_cards["stagnant"].config(text=f"{len(s['stagnant'])} صنف • {fmt(s['stagnant_value'])} ر.س")
        clear_tree(self.exp_tree)
        for i, r in enumerate(s["rows"]):
            tag = "expired" if r["days"] < 0 else ("near" if r["days"] <= alert else ("mid" if r["days"] <= 90 else ""))
            days = f"منتهي منذ {-r['days']} يوم" if r["days"] < 0 else f"{r['days']}"
            self.exp_tree.insert("", "end", iid=str(i), tags=(tag,) if tag else (), values=(
                r["code"], r["name"], r["ref"], r["party"], r["production"] or "-", r["expiry"], days,
                f"{r['qty']:g}", fmt(r["value"]), r["status"]))
        clear_tree(self.stag_tree)
        for code, name, unit, q, cost, val, last, since in s["stagnant"]:
            self.stag_tree.insert("", "end", values=(code, name, unit, f"{q:g}", fmt(cost), fmt(val), last,
                                                     since if since is not None else "-"))

    # ============================================================
    # لوحة المؤشرات
    # ============================================================
    # مجموعات اللوحة بترتيب منطقي: المبيعات ← المشتريات والمخزون ← الخزينة والبنوك ← المصروفات والنتائج
    # (مفتاح المؤشر، اللون، أيقونة)
    DASH_GROUPS = [
        ("g_sales", "#2563eb", [("sales", "#2563eb", "🛒"), ("recv", "#d97706", "📥"),
                                ("best_cust", "#0284c7", "⭐"), ("best_item", "#db2777", "🔥")]),
        ("g_purch", "#7c3aed", [("purchases", "#7c3aed", "📦"), ("payb", "#db2777", "📤"),
                                ("best_supp", "#0d9488", "⭐"), ("stock", "#6366f1", "🏭")]),
        ("g_cash", "#0d9488", [("cash_box", "#0d9488", "💵"), ("banks", "#2563eb", "🏦"),
                               ("cash_total", "#0891b2", "💰"), ("net_position", "#16a34a", "⚖️")]),
        ("g_exp", "#dc2626", [("expenses", "#dc2626", "💸"), ("depreciation", "#ea580c", "📉"),
                              ("gross", "#0284c7", "📊"), ("profit", "#16a34a", "📈")]),
    ]
    DASH_CHARTS = ["flow", "cash", "exp", "expiry"]
    DASH_TEXT = {
        "ar": dict(
            title="لوحة المؤشرات والأداء العام", subtitle="من الفواتير والسندات والقيود الفعلية • المبالغ قبل الضريبة",
            refresh="🔄 تحديث الآن", auto="تحديث تلقائي كل 30 ثانية", updated="آخر تحديث: {}",
            dark="🌙 الوضع الداكن", light="☀️ الوضع الفاتح", lang="English", cur="ر.س",
            g_sales="المبيعات والعملاء", g_purch="المشتريات والمخزون", g_cash="الخزينة والبنوك والنقدية",
            g_exp="المصروفات ونتائج الأعمال",
            sales="صافي المبيعات (قبل الضريبة)", recv="الذمم المدينة / أرصدة العملاء", best_cust="أفضل عميل",
            best_item="الصنف الأكثر مبيعاً", purchases="صافي المشتريات (قبل الضريبة)",
            payb="الذمم الدائنة / أرصدة الموردين", best_supp="أكبر مورد", stock="قيمة المخزون (بالتكلفة)",
            cash_box="رصيد الصناديق (النقدية)", banks="أرصدة البنوك", cash_total="إجمالي النقدية وما في حكمها",
            net_position="صافي المركز (نقدية + مدينون − دائنون)",
            expenses="المصروفات التشغيلية", depreciation="مصروف الإهلاك", gross="مجمل الربح", profit="صافي الربح",
            none="لا يوجد", unit="وحدة",
            c_flow="📊 المبيعات مقابل المشتريات (آخر 6 أشهر)", c_cash="🏦 أرصدة الخزينة والبنوك",
            c_exp="💸 توزيع المصروفات التشغيلية (هذه السنة)", c_expiry="⏳ الصلاحية والركود (بالقيمة)",
            s_sales="المبيعات", s_purch="المشتريات", other="بنود أخرى",
            x_expired="منتهية الصلاحية", x_near="تنتهي خلال {} يوماً", x_mid="تنتهي خلال 90 يوماً",
            x_stag="راكدة (بلا حركة +{} يوماً)", x_clear="✓ لا توجد أصناف منتهية أو راكدة",
            a_warn="⚠️ تنبيه", a_expired="دفعات منتهية الصلاحية: {} ({} ر.س)",
            a_near="تنتهي خلال {} يوماً: {} دفعة ({} ر.س)", a_stag="أصناف راكدة بلا حركة أكثر من {} يوماً: {}",
            a_more="(اضغط للتفاصيل)", a_ok="✓ لا توجد أصناف منتهية أو قريبة الانتهاء أو راكدة"),
        "en": dict(
            title="Performance Dashboard", subtitle="From actual invoices, vouchers and journals • Amounts before VAT",
            refresh="🔄 Refresh", auto="Auto-refresh every 30 s", updated="Updated: {}",
            dark="🌙 Dark mode", light="☀️ Light mode", lang="العربية", cur="SAR",
            g_sales="Sales & Customers", g_purch="Purchases & Inventory", g_cash="Treasury, Banks & Cash",
            g_exp="Expenses & Results",
            sales="Net Sales (excl. VAT)", recv="Accounts Receivable", best_cust="Top Customer",
            best_item="Best-Selling Item", purchases="Net Purchases (excl. VAT)", payb="Accounts Payable",
            best_supp="Top Supplier", stock="Inventory Value (at cost)", cash_box="Cash on Hand",
            banks="Bank Balances", cash_total="Total Cash & Equivalents",
            net_position="Net Position (Cash + AR − AP)",
            expenses="Operating Expenses (OpEx)", depreciation="Depreciation Expense", gross="Gross Profit",
            profit="Net Profit", none="None", unit="units",
            c_flow="📊 Sales vs Purchases (last 6 months)", c_cash="🏦 Treasury & Bank Balances",
            c_exp="💸 Operating Expenses Breakdown (YTD)", c_expiry="⏳ Expiry & Stagnant Stock (value)",
            s_sales="Sales", s_purch="Purchases", other="Other items",
            x_expired="Expired", x_near="Expiring ≤ {} days", x_mid="Expiring ≤ 90 days",
            x_stag="Stagnant (no movement > {} days)", x_clear="✓ No expired or stagnant items",
            a_warn="⚠️ Alert", a_expired="Expired batches: {} ({} SAR)",
            a_near="Expiring within {} days: {} batches ({} SAR)", a_stag="Stagnant items (no movement > {} days): {}",
            a_more="(click for details)", a_ok="✓ No expired, near-expiry or stagnant items"),
    }
    DASH_THEMES = {
        "light": dict(bg="#f1f5f9", card="#ffffff", border="#e2e8f0", ink="#0f172a", text="#64748b", label="#334155",
                      grid="#e2e8f0", axis="#94a3b8", head="#0f172a", btn="#0f172a", btn_fg="#ffffff",
                      btn_active="#1e293b", toggle="#e2e8f0", toggle_fg="#0f172a", pos="#16a34a", neg="#dc2626",
                      tip_bg="#0f172a", tip_fg="#ffffff",
                      sales="#2563eb", purch="#f59e0b", exp="#e11d48", cash="#0d9488",
                      ok=("#dcfce7", "#166534"), warn=("#fef3c7", "#92400e"), bad=("#fee2e2", "#991b1b")),
        "dark": dict(bg="#0b1220", card="#131c2e", border="#23314d", ink="#e5e7eb", text="#94a3b8", label="#cbd5e1",
                     grid="#22304a", axis="#475569", head="#f8fafc", btn="#2563eb", btn_fg="#ffffff",
                     btn_active="#1d4ed8", toggle="#1e293b", toggle_fg="#e5e7eb", pos="#4ade80", neg="#f87171",
                     tip_bg="#e5e7eb", tip_fg="#0f172a",
                     sales="#60a5fa", purch="#fbbf24", exp="#fb7185", cash="#2dd4bf",
                     ok=("#052e1a", "#86efac"), warn=("#3a2a06", "#fcd34d"), bad=("#3b0d0d", "#fca5a5")),
    }
    DASH_BG = DASH_THEMES["light"]["bg"]

    # ---------------- تفضيلات العرض (لكل مستخدم)
    def _dash_pref(self, name, default, allowed):
        try:
            v = get_setting(f"dash_{name}_{self.user.get('username', '')}", default)
        except Exception:
            v = default
        return v if v in allowed else default

    def _dt(self, key, *args):
        text = self.DASH_TEXT[self.dash_lang].get(key, key)
        return text.format(*args) if args else text

    @property
    def _dp(self):
        return self.DASH_THEMES[self.dash_theme]

    def toggle_dashboard_theme(self):
        self.dash_theme = "light" if self.dash_theme == "dark" else "dark"
        self._save_dash_pref("theme", self.dash_theme)
        self._build_dashboard()

    def toggle_dashboard_lang(self):
        self.dash_lang = "en" if self.dash_lang == "ar" else "ar"
        self._save_dash_pref("lang", self.dash_lang)
        self._build_dashboard()

    def _save_dash_pref(self, name, value):
        try:
            set_setting(f"dash_{name}_{self.user.get('username', '')}", value)
        except Exception as e:
            print(f"تعذر حفظ تفضيل اللوحة: {e}")

    # ---------------- البناء
    def setup_dashboard_tab(self, frame):
        self.dash_theme = self._dash_pref("theme", "light", self.DASH_THEMES)
        self.dash_lang = self._dash_pref("lang", "ar", self.DASH_TEXT)
        self.dash_auto_var = tk.IntVar(value=1)
        self.dash_metrics, self.dash_chart_data, self.dash_expiry = None, {}, None
        self.dash_host = tk.Frame(frame)
        self.dash_host.pack(fill="both", expand=True)
        self._build_dashboard(reload=True)

    def _build_dashboard(self, reload=False):
        """يبني اللوحة كاملة بالوضع واللغة الحاليين (يُعاد البناء فوراً عند التبديل دون إعادة القراءة من القاعدة)."""
        P, rtl = self._dp, self.dash_lang == "ar"
        R, L = ("right", "left") if rtl else ("left", "right")
        anchor = "e" if rtl else "w"
        host = self.dash_host
        for w in host.winfo_children():
            w.destroy()
        host.configure(bg=P["bg"])

        # منطقة قابلة للتمرير حتى تظهر كل المجموعات على الشاشات الصغيرة
        canvas = tk.Canvas(host, bg=P["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(host, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side=L, fill="y")
        canvas.pack(side=R, fill="both", expand=True)
        page = tk.Frame(canvas, bg=P["bg"])
        win = canvas.create_window((0, 0), window=page, anchor="nw")
        page.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(win, width=e.width))

        def wheel(e):
            if page.winfo_height() > canvas.winfo_height():
                canvas.yview_scroll(int(-e.delta / 120) or (-1 if e.delta > 0 else 1), "units")
        canvas.bind("<Enter>", lambda e: canvas.bind_all("<MouseWheel>", wheel))
        canvas.bind("<Leave>", lambda e: canvas.unbind_all("<MouseWheel>"))

        # ---- الشريط العلوي: العنوان + التبديل بين الوضعين واللغتين + التحديث
        top = tk.Frame(page, bg=P["bg"])
        top.pack(fill="x", padx=16, pady=(14, 6))
        title_box = tk.Frame(top, bg=P["bg"])
        title_box.pack(side=R)
        tk.Label(title_box, text=self._dt("title"), font=("Arial", 16, "bold"), bg=P["bg"], fg=P["head"],
                 anchor=anchor).pack(fill="x")
        tk.Label(title_box, text=self._dt("subtitle"), font=("Arial", 9), bg=P["bg"], fg=P["text"],
                 anchor=anchor).pack(fill="x")

        def tbtn(text, cmd, primary=False):
            return tk.Button(top, text=text, command=cmd, font=("Arial", 9, "bold"), bd=0, padx=12, pady=6,
                             cursor="hand2", relief="flat",
                             bg=P["btn"] if primary else P["toggle"], fg=P["btn_fg"] if primary else P["toggle_fg"],
                             activebackground=P["btn_active"] if primary else P["border"],
                             activeforeground=P["btn_fg"] if primary else P["toggle_fg"])
        tbtn(self._dt("refresh"), self.load_dashboard_data, primary=True).pack(side=L)
        tbtn(self._dt("light") if self.dash_theme == "dark" else self._dt("dark"),
             self.toggle_dashboard_theme).pack(side=L, padx=(8, 0))
        tbtn("🌐 " + self._dt("lang"), self.toggle_dashboard_lang).pack(side=L, padx=(8, 0))
        tk.Checkbutton(top, text=self._dt("auto"), variable=self.dash_auto_var, bg=P["bg"], fg=P["label"],
                       activebackground=P["bg"], activeforeground=P["ink"], selectcolor=P["card"],
                       font=("Arial", 9)).pack(side=L, padx=12)
        self.dash_updated_lbl = tk.Label(top, text="", bg=P["bg"], fg=P["text"], font=("Arial", 9))
        self.dash_updated_lbl.pack(side=L)
        self.dash_alert = tk.Frame(page, bg=P["bg"], cursor="hand2")
        self.dash_alert.pack(fill="x", padx=16, pady=(0, 4))

        # ---- مجموعات المؤشرات (2 × 2): كل مجموعة لوحة بعنوان وأربع بطاقات
        groups = tk.Frame(page, bg=P["bg"])
        groups.pack(fill="x", padx=10)
        self.dash_labels = {}
        for gi, (gkey, gcolor, metrics) in enumerate(self.DASH_GROUPS):
            gr, gc = divmod(gi, 2)
            panel = tk.Frame(groups, bg=P["card"], highlightbackground=P["border"], highlightthickness=1)
            panel.grid(row=gr, column=(1 - gc) if rtl else gc, padx=6, pady=6, sticky="nsew")
            head = tk.Frame(panel, bg=P["card"])
            head.pack(fill="x", padx=12, pady=(10, 2))
            tk.Frame(head, bg=gcolor, width=4, height=18).pack(side=R)
            tk.Label(head, text=self._dt(gkey), bg=P["card"], fg=P["head"], font=("Arial", 11, "bold")).pack(
                side=R, padx=8)
            body = tk.Frame(panel, bg=P["card"])
            body.pack(fill="both", expand=True, padx=6, pady=(0, 8))
            for mi, (key, color, icon) in enumerate(metrics):
                r, c = divmod(mi, 2)
                card = tk.Frame(body, bg=P["bg"], highlightthickness=0)
                card.grid(row=r, column=(1 - c) if rtl else c, padx=5, pady=5, sticky="nsew")
                tk.Frame(card, bg=color, width=4).pack(side=R, fill="y")
                inner = tk.Frame(card, bg=P["bg"])
                inner.pack(side=R, fill="both", expand=True, padx=10, pady=7)
                tk.Label(inner, text=f"{icon}  {self._dt(key)}" if rtl else f"{icon}  {self._dt(key)}", bg=P["bg"],
                         fg=P["text"], font=("Arial", 9, "bold"), anchor=anchor).pack(fill="x")
                lbl = tk.Label(inner, text="-", bg=P["bg"], fg=P["ink"], font=("Arial", 14, "bold"), anchor=anchor,
                               justify="right" if rtl else "left", wraplength=250)
                lbl.pack(fill="x", pady=(3, 0))
                self.dash_labels[key] = lbl
            for c in range(2):
                body.grid_columnconfigure(c, weight=1, uniform=f"dash_{gkey}")
        for c in range(2):
            groups.grid_columnconfigure(c, weight=1, uniform="dash_groups")

        # ---- الرسوم بنفس ترتيب المجموعات
        charts = tk.Frame(page, bg=P["bg"])
        charts.pack(fill="both", expand=True, padx=10, pady=(4, 12))
        self.dash_charts = {}
        for i, key in enumerate(self.DASH_CHARTS):
            r, c = divmod(i, 2)
            card = tk.Frame(charts, bg=P["card"], highlightbackground=P["border"], highlightthickness=1)
            card.grid(row=r, column=(1 - c) if rtl else c, padx=6, pady=6, sticky="nsew")
            tk.Label(card, text=self._dt(f"c_{key}"), bg=P["card"], fg=P["head"], font=("Arial", 10, "bold"),
                     anchor=anchor).pack(fill="x", padx=14, pady=(10, 0))
            cv = tk.Canvas(card, bg=P["card"], highlightthickness=0, height=230)
            cv.pack(fill="both", expand=True, padx=8, pady=8)
            cv.bind("<Configure>", lambda e, k=key: self.draw_dashboard_chart(k))
            self.dash_charts[key] = cv
        for c in range(2):
            charts.grid_columnconfigure(c, weight=1, uniform="dash_charts")

        if reload or self.dash_metrics is None:
            self.load_dashboard_data()
        else:
            self._render_dashboard()

    # ---------------- البيانات والعرض
    def load_dashboard_data(self, silent=False):
        try:
            self.dash_metrics = fetch_dashboard_metrics()
            self.dash_chart_data = fetch_dashboard_charts()
            self.dash_expiry = fetch_expiry_summary(stale_days=STALE_DAYS_DEFAULT)
        except Exception as e:
            if silent:
                print(f"تعذر تحديث لوحة المؤشرات: {e}")
            else:
                handle_error(e)
            return
        self._render_dashboard()
        self.dash_updated_lbl.config(text=self._dt("updated", f"{datetime.now():%H:%M:%S}"))

    def _render_dashboard(self):
        P, m, ex = self._dp, self.dash_metrics, self.dash_expiry
        if m is None:
            return
        cur = self._dt("cur")
        for key, lbl in self.dash_labels.items():
            v = m.get(key, 0)
            if isinstance(v, tuple) or v is None:  # (الاسم، القيمة) لأفضل عميل/مورد/صنف
                if not v:
                    lbl.config(text=self._dt("none"), fg=P["text"])
                elif key == "best_item":
                    lbl.config(text=f"{v[0]} ({v[1]:,.0f} {self._dt('unit')})", fg=P["ink"])
                else:
                    lbl.config(text=f"{v[0]} ({fmt(v[1])} {cur})", fg=P["ink"])
                lbl.config(font=("Arial", 11, "bold"))
                continue
            lbl.config(text=f"{v:,.2f} {cur}")
            if key in ("profit", "gross", "net_position"):
                lbl.config(fg=P["pos"] if v >= 0 else P["neg"])
            else:
                lbl.config(fg=P["neg"] if v < -0.005 else P["ink"])
        if ex is not None:
            self.dash_chart_data["expiry"] = [
                (self._dt("x_expired"), ex["expired_value"], P["neg"]),
                (self._dt("x_near", ex["alert_days"]), ex["near_value"], "#ea580c"),
                (self._dt("x_mid"), ex["mid_value"], "#f59e0b"),
                (self._dt("x_stag", ex["stale_days"]), ex["stagnant_value"], "#8b5cf6")]
        for key in self.dash_charts:
            self.draw_dashboard_chart(key)
        if ex is None:
            return
        parts = []
        if ex["expired"]:
            parts.append(self._dt("a_expired", len(ex["expired"]), fmt(ex["expired_value"])))
        if ex["near"]:
            parts.append(self._dt("a_near", ex["alert_days"], len(ex["near"]), fmt(ex["near_value"])))
        if ex["stagnant"]:
            parts.append(self._dt("a_stag", ex["stale_days"], len(ex["stagnant"])))
        if parts:
            bg, fg = P["bad"] if (ex["expired"] or ex["near"]) else P["warn"]
            self._set_alert_banner([self._dt("a_warn")] + parts + [self._dt("a_more")], bg, fg)
        else:
            self._set_alert_banner([self._dt("a_ok")], *P["ok"])

    def _set_alert_banner(self, parts, bg, fg):
        """شريط التنبيهات: كل جزء في عنوان مستقل مرصوص باتجاه اللغة (يمين→يسار بالعربي)."""
        bar = self.dash_alert
        side = "right" if self.dash_lang == "ar" else "left"
        for w in bar.winfo_children():
            w.destroy()
        bar.config(bg=bg)
        go = lambda e: self.show_page("expiry") if self.can("expiry") else None
        for i, text in enumerate(parts):
            if 1 < i < len(parts) - 1:
                tk.Label(bar, text="•", bg=bg, fg=fg, font=("Arial", 10, "bold")).pack(side=side, padx=4, pady=7)
            tk.Label(bar, text=text, bg=bg, fg=fg, font=("Arial", 10, "bold"), cursor="hand2").pack(
                side=side, padx=(6, 6) if i else (14, 6), pady=7)
        for w in (bar, *bar.winfo_children()):
            w.bind("<Button-1>", go)

    def draw_dashboard_chart(self, key):
        cv = self.dash_charts.get(key) if hasattr(self, "dash_charts") else None
        data = self.dash_chart_data.get(key)
        if cv is None or data is None or not cv.winfo_exists():
            return
        P = self._dp
        CHART_PAL.update(grid=P["grid"], text=P["text"], ink=P["ink"], label=P["label"], axis=P["axis"],
                         tip_bg=P["tip_bg"], tip_fg=P["tip_fg"], neg=P["neg"], rtl=self.dash_lang == "ar",
                         lang=self.dash_lang)
        try:
            if key == "flow":
                keys, sales, purch = data
                months = AR_MONTHS if self.dash_lang == "ar" else EN_MONTHS
                labels = [months[int(k[5:7]) - 1] for k in keys]
                draw_grouped_bars(cv, labels, [(self._dt("s_sales"), P["sales"], sales),
                                               (self._dt("s_purch"), P["purch"], purch)])
            elif key == "exp":
                draw_hbars(cv, [(n if n is not None else self._dt("other"), v) for n, v in data], P["exp"])
            elif key == "expiry":
                if not any(v > 0.005 for _, v, _ in data):
                    cv.delete("all")
                    _chart_empty(cv, self._dt("x_clear"))
                else:
                    draw_hbars(cv, [(l, v) for l, v, _ in data], P["exp"], colors=[c for _, _, c in data])
            else:
                draw_hbars(cv, data, P["cash"])
        finally:
            CHART_PAL.clear()
            CHART_PAL.update(CHART_PAL_DEFAULT)

    def _schedule_dashboard_refresh(self):
        self._dash_job = self.root.after(DASH_REFRESH_MS, self._auto_refresh_dashboard)

    def _auto_refresh_dashboard(self):
        """تحديث حي: كل 30 ثانية طالما لوحة المؤشرات معروضة والتحديث التلقائي مفعّل."""
        self._dash_job = None
        try:
            if self.dash_auto_var.get() and self.notebook.select() == str(self.frames["dashboard"]):
                self.load_dashboard_data(silent=True)
        except tk.TclError:
            return  # النافذة أُغلقت
        self._schedule_dashboard_refresh()

    # ============================================================
    # الخزينة: بحث وتصفية وطباعة السندات وملخص اليوم
    # ============================================================
    def setup_cash_tab(self, frame):
        f = ttk.LabelFrame(frame, text=" سند قبض / صرف وربط الفواتير ", padding=10)
        f.pack(fill="x", padx=10, pady=(8, 4))

        types = [t for t, k in (("قبض (وارد)", "cash.receipt"), ("صرف (منصرف)", "cash.payment")) if self.can(k)]
        self.c_type = ttk.Combobox(f, values=types or ["قبض (وارد)"], width=14, state="readonly")
        self.c_type.set((types or ["قبض (وارد)"])[0])
        self.c_type.bind("<<ComboboxSelected>>", lambda e: self.refresh_open_invoices())
        self.c_bank = ttk.Combobox(f, values=BANK_ACCOUNTS, width=26, state="readonly")
        self.c_bank.set(BANK_ACCOUNTS[0])
        self.refresh_cash_accounts()
        self.c_date = date_entry(f, 12)
        rtl_row(f, 0, [("نوع السند:", self.c_type), ("الحساب:", self.c_bank), ("التاريخ:", self.c_date)])

        self.party_category_cb = ttk.Combobox(f, values=[CAT_PARTNER, CAT_EXPENSE, CAT_ASSET, CAT_EMPLOYEE], width=18,
                                              state="readonly")
        self.party_category_cb.set(CAT_PARTNER)
        self.party_category_cb.bind("<<ComboboxSelected>>", lambda e: self.update_party_options())
        self.party_name_cb = SearchableCombobox(f, width=32)
        self.party_name_cb.bind("<<ComboboxSelected>>", lambda e: self.refresh_open_invoices())
        self.open_inv_cb = SearchableCombobox(f, width=40)
        rtl_row(f, 1, [("التصنيف:", self.party_category_cb), ("الشريك / الحساب:", self.party_name_cb),
                       ("الفاتورة للسداد:", self.open_inv_cb)])

        self.open_inv_cb.bind("<<ComboboxSelected>>", lambda e: self.update_cash_rep_label(), add="+")
        self.c_amount = ttk.Entry(f, width=15)
        self.c_desc = ttk.Entry(f, width=60)
        # المندوب لا يُختار يدوياً: يُورَّث من الفاتورة المسددة أو من بطاقة الطرف ويُحفظ مع السند تلقائياً
        self.cash_rep_lbl = ttk.Label(f, text="👤 المندوب: —", foreground=UI["primary_dk"], font=(UI_FONT, 9, "bold"))
        rtl_row(f, 2, [("المبلغ:", self.c_amount), ("البيان:", self.c_desc), (None, self.cash_rep_lbl)])
        cbar = tk.Frame(f)
        cbar.grid(row=3, column=0, columnspan=8, pady=6)
        money_button(cbar, "حفظ السند وتسوية الفاتورة 💾", "#16a34a", self.save_cash_transaction, 11).pack(side="right", padx=6)
        money_button(cbar, "🔗 تسوية الدفعات المقدمة للطرف", UI["teal"], self.settle_party_advances, 9).pack(side="right", padx=6)
        self.cash_pay_lbl = tk.Label(f, text="", fg="#15803d", font=(UI_FONT, 10, "bold"), justify="right", anchor="e")
        self.cash_pay_lbl.grid(row=4, column=0, columnspan=8, sticky="e")
        if not types:
            f.pack_forget()
        show_register = self.can("cash.register")

        flt = ttk.LabelFrame(frame, text=" بحث وتصفية السندات ", padding=6)
        if show_register:
            flt.pack(fill="x", padx=10, pady=2)
        self.cash_from_e = date_entry(flt, 12, f"{date.today().year}-01-01")
        self.cash_to_e = date_entry(flt, 12)
        self.cash_q_e = ttk.Entry(flt, width=26)
        self.cash_q_e.bind("<Return>", lambda e: self.load_cash_transactions())
        rtl_row(flt, 0, [("من:", self.cash_from_e), ("إلى:", self.cash_to_e), ("بحث (طرف/بيان/رقم):", self.cash_q_e),
                         (None, money_button(flt, "🔍 تطبيق", "#0284c7", self.load_cash_transactions, 9))], slots=4)
        bf = tk.Frame(flt)
        bf.grid(row=0, column=8, padx=8)
        if self.can("cash.print"):
            money_button(bf, "🖨️ ملخص يوم PDF", "#2563eb", self.print_daily_cash_summary, 9).pack()
            if not show_register:
                pf = tk.Frame(frame)
                pf.pack(fill="x", padx=10, pady=2)
                money_button(pf, "🖨️ ملخص يوم PDF", "#2563eb", self.print_daily_cash_summary, 9).pack(side="left")
        if self.can_modify():
            adm = tk.Frame(frame)
            adm.pack(fill="x", padx=10, pady=(2, 0))
            money_button(adm, "🏦 الأرصدة الافتتاحية للخزينة والبنوك", "#b45309", self.open_cash_opening_dialog, 9
                         ).pack(side="right", padx=3)
            money_button(adm, "✏️ تعديل السند المحدد", "#d97706", self.edit_voucher_dialog, 9).pack(side="left", padx=3)
            money_button(adm, "🗑️ حذف السند المحدد", "#dc2626", self.delete_voucher_selected, 9).pack(side="left", padx=3)

        box = ttk.LabelFrame(frame, text=" تقرير الخزينة (اضغط مرتين على سند لعرضه وطباعته) ", padding=8)
        if show_register:
            box.pack(fill="both", expand=True, padx=10, pady=4)
        self.summary_lbl = tk.Label(box, text="", font=("Arial", 11, "bold"), bg="#e2e8f0", pady=6)
        self.summary_lbl.pack(fill="x", pady=(0, 2))
        self.accounts_lbl = tk.Label(box, text="", font=("Arial", 9), bg="#f1f5f9", pady=4)
        self.accounts_lbl.pack(fill="x", pady=(0, 5))
        holder = tk.Frame(box)
        holder.pack(fill="both", expand=True)
        self.cash_tree = make_tree(holder, [
            ("id", "رقم السند", 70, "center"), ("type", "النوع", 60, "center"), ("date", "التاريخ", 90, "center"),
            ("bank", "الحساب", 130, "center"), ("cat", "التصنيف", 110, "center"), ("party", "الطرف / البند", 200, "e"),
            ("amount", "المبلغ", 100, "center"), ("inv", "الفاتورة المسددة", 130, "center"), ("desc", "البيان", 200, "center")])
        self.cash_tree.bind("<Double-1>", self.on_cash_row_double_click)
        self.load_cash_transactions()
        self.update_party_options()

    def save_cash_transaction(self):
        try:
            t_type = "قبض" if self.c_type.get().startswith("قبض") else "صرف"
            if not self.can("cash.receipt" if t_type == "قبض" else "cash.payment"):
                raise UserError(f"ليس لديك صلاحية تسجيل سندات ال{t_type}.")
            category = self.party_category_cb.get()
            t_date = parse_date(self.c_date.get())
            if not t_date:
                raise UserError("صيغة التاريخ غير صحيحة، استخدم YYYY-MM-DD")
            if t_date > today():
                raise UserError("لا يمكن تسجيل سند بتاريخ مستقبلي.")
            party = self.party_name_cb.selected()
            if not party:
                raise UserError("اختر الشريك / الحساب من القائمة.")
            amount = parse_float(self.c_amount.get())
            alloc = self.open_inv_cb.selected() if category == CAT_PARTNER else None
            self.check_period_open(t_date)
            vid = save_voucher(t_type, t_date, self.c_bank.get(), category, party, amount, alloc, self.c_desc.get())

            self.c_amount.delete(0, tk.END)
            self.c_desc.delete(0, tk.END)
            if getattr(self, "cash_pay_lbl", None) is not None:
                self.cash_pay_lbl.configure(text="✅ تم الحفظ. ارجع لكشف الحساب أو الفاتورة من «مسار العمل» أعلى الشاشة.")
            self.load_cash_transactions()
            self.refresh_open_invoices()
            messagebox.showinfo("نجاح", f"تم حفظ السند رقم {vid} وتسوية الحساب.")
        except Exception as e:
            handle_error(e)

    def load_cash_transactions(self):
        clear_tree(self.cash_tree)
        d1 = parse_date(self.cash_from_e.get()) or "0000-01-01"
        d2 = parse_date(self.cash_to_e.get()) or "9999-12-31"
        q = self.cash_q_e.get().strip()
        total_in = total_out = 0.0
        sql = ("SELECT trans_id, trans_type, trans_date, account_bank, category, party_name, amount, alloc_type, alloc_id, description "
               "FROM cash_transactions WHERE trans_date BETWEEN ? AND ?")
        params = [d1, d2]
        if q:
            sql += " AND (party_name LIKE ? OR description LIKE ? OR CAST(trans_id AS TEXT) = ?)"
            params += [f"%{q}%", f"%{q}%", q]
        with db() as conn:
            nos = {("sales", i): n for i, n in conn.execute("SELECT invoice_id, invoice_no FROM sales_invoices")}
            nos.update({("purchase", i): n for i, n in conn.execute("SELECT purchase_id, invoice_no FROM purchase_invoices")})
            for r in conn.execute(sql + " ORDER BY trans_date DESC, trans_id DESC", params):
                inv = (f"{'مبيعات' if r[7] == 'sales' else 'مشتريات'} {nos.get((r[7], r[8]), r[8])}"
                       if r[7] in ("sales", "purchase") else "-")
                self.cash_tree.insert("", "end", iid=str(r[0]), values=(r[0], r[1], r[2], r[3], r[4], r[5], fmt(r[6]), inv, r[9] or "-"))
                if r[1] == "قبض":
                    total_in += r[6]
                else:
                    total_out += r[6]
            balances = [(b, v) for b, v in sorted(cash_balances(conn).items()) if abs(v) > 0.005]
        self.summary_lbl.config(text=f"(الفترة المعروضة) المقبوضات: {fmt(total_in)}  |  المدفوعات: {fmt(total_out)}  |  "
                                     f"الصافي: {fmt(total_in - total_out)}")
        self.accounts_lbl.config(text="أرصدة الحسابات (سندات + قيود، كل الفترات):  " +
                                      ("   |   ".join(f"{b}: {fmt(v)}" for b, v in balances) or "لا توجد حركات"))

    def _selected_voucher(self):
        sel = self.cash_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر سنداً من الجدول أولاً.")
            return None
        return int(sel[0])

    def delete_voucher_selected(self):
        if not self.require_modify():
            return
        tid = self._selected_voucher()
        if tid is None:
            return
        if not messagebox.askyesno("حذف سند", f"حذف السند رقم {tid} نهائياً؟\nلو كان مربوطاً بفاتورة يُعاد احتساب المسدد منها."):
            return
        try:
            delete_voucher(tid)
            self.load_cash_transactions()
            self.refresh_open_invoices()
        except Exception as e:
            handle_error(e)

    def edit_voucher_dialog(self):
        if not self.require_modify():
            return
        tid = self._selected_voucher()
        if tid is None:
            return
        with db() as conn:
            r = conn.execute("SELECT trans_type, trans_date, account_bank, category, party_name, amount, COALESCE(description,'') "
                             "FROM cash_transactions WHERE trans_id=?", (tid,)).fetchone()
            names = cash_account_names(conn)
        if not r:
            return
        win = tk.Toplevel(self.root)
        win.title(f"تعديل السند رقم {tid}")
        win.geometry("520x300")
        win.grab_set()
        f = ttk.Frame(win, padding=16)
        f.pack(fill="both", expand=True)
        ttk.Label(f, text=f"سند {r[0]} • {r[3]} • {r[4]}", font=("Arial", 10, "bold")).grid(row=0, column=0, columnspan=2, sticky="e")
        date_e = date_entry(f, 14, r[1])
        acc_cb = ttk.Combobox(f, values=names, width=28, state="readonly")
        acc_cb.set(r[2])
        amount_e = ttk.Entry(f, width=16)
        amount_e.insert(0, f"{r[5]:g}")
        desc_e = ttk.Entry(f, width=40)
        desc_e.insert(0, r[6])
        for i, (label, w) in enumerate((("التاريخ:", date_e), ("الحساب:", acc_cb), ("المبلغ:", amount_e), ("البيان:", desc_e)), start=1):
            ttk.Label(f, text=label).grid(row=i, column=1, sticky="e", pady=5)
            w.grid(row=i, column=0, sticky="w", pady=5)
        ttk.Label(f, text="لتغيير الطرف أو التصنيف: احذف السند وسجّل سنداً جديداً.", foreground="#78716c").grid(
            row=5, column=0, columnspan=2, sticky="e")

        def save():
            try:
                update_voucher(tid, parse_date(date_e.get()), acc_cb.get(), parse_float(amount_e.get()), desc_e.get().strip())
                win.destroy()
                self.load_cash_transactions()
                self.refresh_open_invoices()
            except Exception as ex:
                handle_error(ex)

        money_button(f, "حفظ تعديل السند 💾", "#16a34a", save, 10).grid(row=6, column=0, columnspan=2, pady=10)

    def refresh_cash_accounts(self):
        """قائمة حسابات الخزينة من دليل الحسابات (فرع النقدية والبنوك)."""
        with db() as conn:
            names = cash_account_names(conn)
        current = self.c_bank.get()
        self.c_bank.configure(values=names)
        if current not in names:
            self.c_bank.set(names[0])

    def on_cash_row_double_click(self, event=None):
        sel = self.cash_tree.selection()
        if not sel:
            return
        with db() as conn:
            r = conn.execute("SELECT trans_id, trans_type, trans_date, account_bank, category, party_name, amount, description "
                             "FROM cash_transactions WHERE trans_id=?", (int(sel[0]),)).fetchone()
        if not r:
            return
        detail = (f"سند {r[1]} رقم {r[0]}\nالتاريخ: {r[2]}\nالحساب: {r[3]}\nالتصنيف: {r[4]}\nالطرف / البند: {r[5]}\n"
                  f"المبلغ: {fmt(r[6])} ر.س\nالبيان: {r[7] or '-'}")
        if messagebox.askyesno("تفاصيل السند", detail + "\n\nهل تريد طباعة السند PDF؟"):
            path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                                title="حفظ السند", initialfile=f"voucher_{r[0]}.pdf")
            if not path:
                return
            try:
                rows = [("نوع السند", r[1]), ("رقم السند", str(r[0])), ("التاريخ", r[2]), ("الحساب", r[3]), ("التصنيف", r[4]),
                        ("الطرف / البند", r[5]), ("المبلغ", f"{fmt(r[6])} ر.س"), ("البيان", r[7] or "-")]
                sign = "المستلم: ........................     أمين الصندوق: ........................     التوقيع: ................"
                _simple_report_pdf(path, [f"سند {r[1]} رقم {r[0]}"], ["البند", "التفاصيل"], rows, [150, 380], [sign])
                messagebox.showinfo("نجاح", "تم إنشاء السند PDF.")
            except Exception as e:
                handle_error(e)

    def print_daily_cash_summary(self):
        d = simpledialog.askstring("ملخص يوم", "التاريخ (YYYY-MM-DD):", initialvalue=today())
        d = parse_date(d or "")
        if not d:
            return
        path = filedialog.asksaveasfilename(defaultextension=".pdf", filetypes=[("PDF files", "*.pdf")],
                                            title="حفظ ملخص اليوم", initialfile=f"cash_{d}.pdf")
        if not path:
            return
        try:
            with db() as conn:
                data = conn.execute("SELECT trans_id, trans_type, account_bank, party_name, amount, description "
                                    "FROM cash_transactions WHERE trans_date=? ORDER BY trans_id", (d,)).fetchall()
            rows = [(str(r[0]), r[1], r[2], r[3], fmt(r[4]), r[5] or "-") for r in data]
            tin = sum(r[4] for r in data if r[1] == "قبض")
            tout = sum(r[4] for r in data if r[1] == "صرف")
            _simple_report_pdf(path, [f"ملخص حركة الخزينة ليوم {d}"], ["رقم", "النوع", "الحساب", "الطرف", "المبلغ", "البيان"], rows,
                               [40, 45, 100, 150, 80, 125],
                               [f"إجمالي المقبوضات: {fmt(tin)}   |   إجمالي المدفوعات: {fmt(tout)}   |   الصافي: {fmt(tin - tout)} ر.س"])
            messagebox.showinfo("نجاح", "تم إنشاء ملخص اليوم PDF.")
        except Exception as e:
            handle_error(e)

    # ============================================================
    # صفحة التقارير والتحليل
    # ============================================================
    PERIODS = ["هذا الشهر", "الشهر الماضي", "هذا الربع", "هذه السنة", "السنة الماضية", "آخر 30 يوم", "آخر 90 يوم", "كل الفترات"]
    SERIES_REPORTS = ("s_period", "pnl", "coll", "cash_flow", "vat")

    def setup_reports_tab(self, frame):
        main = tk.Frame(frame)
        main.pack(fill="both", expand=True)
        cat = ttk.LabelFrame(main, text=" قائمة التقارير ", padding=6)
        cat.pack(side="right", fill="y", padx=(8, 4), pady=4)
        self.rep_catalog = ttk.Treeview(cat, show="tree", selectmode="browse", height=26)
        self.rep_catalog.column("#0", width=270)
        self.rep_catalog.pack(fill="y", expand=True)
        for group, items in REPORT_CATALOG:
            allowed = [(key, title) for key, title, _, _ in items if self.can_report(key)]
            if not allowed:
                continue
            gid = self.rep_catalog.insert("", "end", text=group, open=True)
            for key, title in allowed:
                self.rep_catalog.insert(gid, "end", iid=key, text=title)
        self.rep_catalog.bind("<<TreeviewSelect>>", self.on_report_selected)

        area = tk.Frame(main)
        area.pack(side="left", fill="both", expand=True, pady=4)
        flt = ttk.LabelFrame(area, text=" المرشحات ", padding=8)
        flt.pack(fill="x")
        self.rep_period_cb = ttk.Combobox(flt, values=self.PERIODS, state="readonly", width=16)
        self.rep_period_cb.set("هذه السنة")
        self.rep_period_cb.bind("<<ComboboxSelected>>", self.apply_quick_period)
        self.rep_from_e = DateEntry(flt, width=13)
        self.rep_to_e = DateEntry(flt, width=13)
        self.rep_group_cb = ttk.Combobox(flt, values=["شهري", "يومي"], state="readonly", width=8)
        self.rep_group_cb.set("شهري")
        self.rep_stale_e = ttk.Entry(flt, width=6)
        self.rep_stale_e.insert(0, str(STALE_DAYS_DEFAULT))
        self.rep_late_e = ttk.Entry(flt, width=6)
        self.rep_late_e.insert(0, "7")
        rtl_row(flt, 0, [("فترة سريعة:", self.rep_period_cb), ("من تاريخ:", self.rep_from_e), ("إلى تاريخ:", self.rep_to_e),
                         ("التجميع:", self.rep_group_cb)])
        self.rep_expiry_e = ttk.Entry(flt, width=6)
        self.rep_expiry_e.insert(0, "30")
        self.rep_rep_cb = ttk.Combobox(flt, width=18, state="readonly")
        self.rep_cust_cb = SearchableCombobox(flt, width=26)
        rtl_row(flt, 1, [("أيام الركود/الانقطاع:", self.rep_stale_e), ("مهلة تأخر الأوامر (يوم):", self.rep_late_e),
                         ("تنبيه الصلاحية (يوم):", self.rep_expiry_e)])
        rtl_row(flt, 3, [("المندوب (تقرير مدير المبيعات):", self.rep_rep_cb), ("العميل:", self.rep_cust_cb)], slots=4)
        bf = tk.Frame(flt)
        bf.grid(row=2, column=0, columnspan=8, pady=4)
        money_button(bf, "🔍 عرض التقرير", "#0284c7", self.run_current_report, 10).pack(side="left", padx=4)
        money_button(bf, "📗 تصدير Excel", "#16a34a", lambda: self.export_report("xlsx"), 10).pack(side="left", padx=4)
        money_button(bf, "📄 تصدير PDF", "#dc2626", lambda: self.export_report("pdf"), 10).pack(side="left", padx=4)

        self.rep_desc_lbl = tk.Label(area, text="اختر تقريراً من القائمة على اليمين", anchor="e", justify="right",
                                     fg="#78716c", wraplength=900)
        self.rep_desc_lbl.pack(fill="x", padx=6, pady=(4, 0))
        self.rep_note_lbl = tk.Label(area, text="", anchor="e", justify="right", bg="#e2e8f0", wraplength=900,
                                     font=("Arial", 9, "bold"))
        self.rep_note_lbl.pack(fill="x", padx=6, pady=3)
        tree_box = tk.Frame(area)
        tree_box.pack(fill="both", expand=True, padx=6)
        self.rep_tree = SortableTree(tree_box, columns=("c0",), show="headings")
        self.rep_tree.setup([("c0", "-", 120, "center")])
        self.rep_tree.bind("<Double-1>", self.drill_report_row)
        vsb = ttk.Scrollbar(tree_box, orient="vertical", command=self.rep_tree.yview)
        hsb = ttk.Scrollbar(tree_box, orient="horizontal", command=self.rep_tree.xview)
        self.rep_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.rep_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tree_box.grid_rowconfigure(0, weight=1)
        tree_box.grid_columnconfigure(0, weight=1)
        self.rep_tree.tag_configure("total", background="#fde68a", font=("Arial", 9, "bold"))
        self.rep_chart = tk.Canvas(area, height=150, bg="white", highlightthickness=1, highlightbackground="#cbd5e1")
        self.rep_chart.pack(fill="x", padx=6, pady=4)
        self.rep_current = None
        self.rep_last = None
        self.apply_quick_period()

    def apply_quick_period(self, event=None):
        t = date.today()
        choice = self.rep_period_cb.get()
        if choice == "هذا الشهر":
            a, b = t.replace(day=1), t
        elif choice == "الشهر الماضي":
            b = t.replace(day=1) - timedelta(days=1)
            a = b.replace(day=1)
        elif choice == "هذا الربع":
            a, b = date(t.year, ((t.month - 1) // 3) * 3 + 1, 1), t
        elif choice == "السنة الماضية":
            a, b = date(t.year - 1, 1, 1), date(t.year - 1, 12, 31)
        elif choice == "آخر 30 يوم":
            a, b = t - timedelta(days=30), t
        elif choice == "آخر 90 يوم":
            a, b = t - timedelta(days=90), t
        elif choice == "كل الفترات":
            a, b = date(2000, 1, 1), t
        else:
            a, b = date(t.year, 1, 1), t
        self.rep_from_e.delete(0, tk.END)
        self.rep_from_e.insert(0, a.strftime("%Y-%m-%d"))
        self.rep_to_e.delete(0, tk.END)
        self.rep_to_e.insert(0, b.strftime("%Y-%m-%d"))
        if self.rep_current:
            self.run_current_report()

    def on_reports_shown(self):
        try:   # قوائم فلاتر تقرير مدير المبيعات
            self.rep_rep_cb.configure(values=["كل المناديب"] + rep_names() + [NO_REP])
            if not self.rep_rep_cb.get():
                self.rep_rep_cb.set("كل المناديب")
            self.rep_cust_cb.set_items([(lbl, v[0]) for lbl, v in self.partner_pairs("customer")])
        except Exception:  # noqa: BLE001
            pass
        if not self.rep_current:
            first = next((k for _, items in REPORT_CATALOG for k, *_ in items if self.can_report(k)), None)
            if first:
                self.rep_catalog.selection_set(first)

    def on_report_selected(self, event=None):
        sel = self.rep_catalog.selection()
        if not sel or sel[0] not in REPORT_INDEX or not self.can_report(sel[0]):
            return
        self.rep_current = sel[0]
        self.rep_desc_lbl.config(text=REPORT_INDEX[sel[0]][2])
        self.run_current_report()

    def report_params(self):
        d1, d2 = parse_date(self.rep_from_e.get()), parse_date(self.rep_to_e.get())
        if not d1 or not d2:
            raise UserError("صيغة التاريخ YYYY-MM-DD")
        if d1 > d2:
            raise UserError("تاريخ البداية بعد تاريخ النهاية.")
        stale, late = parse_float(self.rep_stale_e.get() or str(STALE_DAYS_DEFAULT)), parse_float(self.rep_late_e.get() or "7")
        if stale is None or late is None or stale < 0 or late < 0:
            raise UserError("أيام الركود ومهلة التأخر يجب أن تكون أرقاماً.")
        expd = parse_float(self.rep_expiry_e.get() or "30")
        if expd is None or expd < 0:
            raise UserError("أيام تنبيه الصلاحية يجب أن تكون رقماً.")
        rep_name = self.rep_rep_cb.get()
        cust = self.rep_cust_cb.selected() if self.rep_cust_cb.get().strip() else None
        if self.rep_cust_cb.get().strip() and not cust:
            raise UserError("اختر العميل من القائمة أو اترك الحقل فارغاً لكل العملاء.")
        return dict(d_from=d1, d_to=d2, group="month" if self.rep_group_cb.get() == "شهري" else "day",
                    stale_days=int(stale), late_days=int(late), expiry_days=int(expd),
                    rep="" if rep_name in ("", "كل المناديب") else rep_name, customer=cust or "")

    def run_current_report(self):
        if not self.rep_current:
            messagebox.showinfo("تنبيه", "اختر تقريراً من القائمة أولاً.")
            return
        try:
            rep = run_report(self.rep_current, self.report_params())
            self.rep_last = rep
            configure_tree(self.rep_tree, rep.columns, getattr(rep, "highlight", ()))
            for row in rep.rows:
                self.rep_tree.insert("", "end", values=[fmt_cell(v, k) for v, (_, k) in zip(row, rep.columns)])
            if rep.total:
                self.rep_tree.insert("", "end", values=[fmt_cell(v, k) for v, (_, k) in zip(rep.total, rep.columns)], tags=("total",))
            self.rep_note_lbl.config(text="\n".join([f"{rep.title}  |  {rep.subtitle}"] + rep.notes))
            self.root.update_idletasks()
            if rep.chart and rep.rows:
                li, vi = rep.chart
                rows = rep.rows[-15:] if self.rep_current in self.SERIES_REPORTS else rep.rows
                draw_bar_chart(self.rep_chart, [str(r[li]) for r in rows], [float(r[vi] or 0) for r in rows], rep.title)
            else:
                self.rep_chart.delete("all")
        except Exception as e:
            handle_error(e)

    def export_report(self, kind):
        if not self.rep_last:
            messagebox.showinfo("تنبيه", "اعرض تقريراً أولاً.")
            return
        ext, types = (".xlsx", [("Excel", "*.xlsx")]) if kind == "xlsx" else (".pdf", [("PDF files", "*.pdf")])
        path = filedialog.asksaveasfilename(defaultextension=ext, filetypes=types, title="حفظ التقرير",
                                            initialfile=f"{self.rep_current}_{today()}{ext}")
        if not path:
            return
        try:
            if kind == "xlsx":
                export_report_xlsx(path, self.rep_last)
            else:
                build_report_pdf(path, self.rep_last)
            messagebox.showinfo("نجاح", "تم تصدير التقرير.")
        except Exception as e:
            handle_error(e)

    # ============================================================
    # القيود اليومية العامة ودليل الحسابات (من النسخة الأحدث)
    # ============================================================
    def setup_journal_tab(self, frame):
        # عنوان التبويب
        tk.Label(frame, text="تسجيل قيد يومية عام (مركب / بسيط)", font=("Arial", 14, "bold"), bg="#f4f4f4").pack(pady=5)
        
        # 1. إطار البيانات الأساسية (التاريخ والبيان العام)
        top_frame = tk.Frame(frame, bg="#f4f4f4")
        top_frame.pack(fill="x", padx=20, pady=5)
        
        tk.Label(top_frame, text="التاريخ:", bg="#f4f4f4").grid(row=0, column=5, padx=5)
        self.jv_date_e = tk.Entry(top_frame, width=15)
        self.jv_date_e.insert(0, datetime.now().strftime("%Y-%m-%d"))
        self.jv_date_e.grid(row=0, column=4, padx=5)
        
        tk.Label(top_frame, text="البيان / الشرح العام:", bg="#f4f4f4").grid(row=0, column=3, padx=5)
        self.jv_desc_e = tk.Entry(top_frame, width=40)
        self.jv_desc_e.grid(row=0, column=2, padx=5)

        # 2. إطار إدخال أسطر القيد (الحساب والمدين والدائن)
        line_frame = ttk.LabelFrame(frame, text=" إدخال أطراف القيد ", padding=10)
        line_frame.pack(fill="x", padx=20, pady=5)

        self.jv_acc_cb = SearchableCombobox(line_frame, width=48)
        self.jv_debit_e = ttk.Entry(line_frame, width=12)
        self.jv_debit_e.insert(0, "0")
        self.jv_credit_e = ttk.Entry(line_frame, width=12)
        self.jv_credit_e.insert(0, "0")
        self.jv_memo_e = ttk.Entry(line_frame, width=25)

        btn_add_line = money_button(line_frame, "➕ إضافة السطر", "#0284c7", self.add_journal_line, 9)

        rtl_row(line_frame, 0, [
            ("الحساب:", self.jv_acc_cb),
            ("مدين:", self.jv_debit_e),
            ("دائن:", self.jv_credit_e),
            ("ملاحظات:", self.jv_memo_e),
            (None, btn_add_line)
        ], slots=5)

        # 3. جدول أطراف القيد
        columns = ("Account", "Debit", "Credit", "Memo")
        self.jv_tree = ttk.Treeview(frame, columns=columns, show="headings", height=8)
        self.jv_tree.heading("Account", text="الحساب")
        self.jv_tree.heading("Debit", text="مدين")
        self.jv_tree.heading("Credit", text="دائن")
        self.jv_tree.heading("Memo", text="ملاحظات")
        self.jv_tree.configure(height=6)
        self.jv_tree.pack(fill="x", padx=20, pady=6)

        # 4. أزرار التحكم والعمليات
        btn_bar = tk.Frame(frame, bg="#f4f4f4")
        btn_bar.pack(pady=10)

        self.jv_editing_id = None
        self.jv_save_btn = money_button(btn_bar, "💾 حفظ وترحيل القيد اليومي", "#27ae60", self.save_journal_entry, 10)
        self.jv_save_btn.pack(side="left", padx=5)
        money_button(btn_bar, "➖ حذف السطر المحدد", "#dc2626", self.remove_journal_line, 9).pack(side="left", padx=5)
        money_button(btn_bar, "✖ قيد جديد / إلغاء التعديل", "#64748b", self.new_journal_entry, 9).pack(side="left", padx=5)

        reg = ttk.LabelFrame(frame, text=" سجل القيود اليومية (اضغط مرتين على قيد لتحميله للتعديل - صلاحية المدير) ", padding=6)
        reg.pack(fill="both", expand=True, padx=20, pady=(0, 8))
        if self.can_modify():
            adm = tk.Frame(reg)
            adm.pack(fill="x", pady=(0, 4))
            money_button(adm, "✏️ تحميل القيد المحدد للتعديل", "#d97706", self.load_journal_for_edit, 9).pack(side="right", padx=3)
            money_button(adm, "🗑️ حذف القيد المحدد", "#dc2626", self.delete_journal_selected, 9).pack(side="right", padx=3)
        att = tk.Frame(reg)
        att.pack(fill="x", pady=(0, 4))
        money_button(att, "📎 مرفقات القيد المحدد (رفع / فتح)", UI["teal"], self.open_journal_attachments_selected, 9
                     ).pack(side="right", padx=3)
        holder = tk.Frame(reg)
        holder.pack(fill="both", expand=True)
        self.jv_reg_tree = make_tree(holder, [("id", "رقم القيد", 70, "center"), ("date", "التاريخ", 90, "center"),
                                              ("desc", "البيان", 320, "e"), ("n", "عدد الأسطر", 80, "center"),
                                              ("total", "الإجمالي", 120, "center"), ("accs", "الحسابات", 380, "e")], height=7)
        self.jv_reg_tree.bind("<Double-1>", lambda e: self.load_journal_for_edit() if self.can_modify() else None)
        self.load_journal_accounts()
        self.load_journal_register()

    def load_journal_accounts(self):
        # الترحيل على الحسابات الفرعية فقط (الحسابات الرئيسية تجميعية)
        with db() as conn:
            pairs = [(f"{c} - {n}", (c, n)) for c, n in conn.execute(
                "SELECT account_code, account_name FROM chart_of_accounts WHERE COALESCE(is_group,0)=0 ORDER BY account_code")]
        self.jv_acc_cb.set_items(pairs)

    def add_journal_line(self):
        try:
            acc = self.jv_acc_cb.selected()
            if not acc:
                raise UserError("يرجى اختيار الحساب أولاً.")

            debit = parse_float(self.jv_debit_e.get() or "0")
            credit = parse_float(self.jv_credit_e.get() or "0")
            memo = self.jv_memo_e.get().strip()

            if debit is None or credit is None:
                raise UserError("أدخل أرقاماً صحيحة في المدين والدائن.")
            debit, credit = round(debit, 2), round(credit, 2)
            if debit < 0 or credit < 0:
                raise UserError("لا تُقبل مبالغ سالبة. لعكس الأثر استخدم الجانب الآخر (مدين/دائن).")
            if debit == 0 and credit == 0:
                raise UserError("أدخل قيمة في المدين أو الدائن.")
            if debit > 0 and credit > 0:
                raise UserError("لا يمكن إدخال قيمة في المدين والدائن معاً بنفس السطر.")

            acc_display = f"{acc[0]} - {acc[1]}"
            self.jv_tree.insert("", "end", values=(acc_display, fmt(debit), fmt(credit), memo))

            # إعادة ضبط حقول الإدخال
            self.jv_debit_e.delete(0, tk.END)
            self.jv_debit_e.insert(0, "0")
            self.jv_credit_e.delete(0, tk.END)
            self.jv_credit_e.insert(0, "0")
            self.jv_memo_e.delete(0, tk.END)
        except Exception as e:
            handle_error(e)

    def remove_journal_line(self):
        sel = self.jv_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "حدد السطر المراد حذفه من الجدول.")
            return
        self.jv_tree.delete(sel[0])

    def save_journal_entry(self):
        try:
            children = self.jv_tree.get_children()
            if not children:
                raise UserError("القيد فارغ! أضف أسطر القيد أولاً.")
            e_date = parse_date(self.jv_date_e.get())
            if not e_date:
                raise UserError("تاريخ القيد غير صحيح.")
            desc = self.jv_desc_e.get().strip() or "قيد يومية عام"
            lines_data = []
            for child in children:
                vals = self.jv_tree.item(child)["values"]
                parts = str(vals[0]).split(" - ", 1)
                lines_data.append((parts[0].strip(), parts[1].strip() if len(parts) > 1 else "",
                                   parse_float(str(vals[1])) or 0.0, parse_float(str(vals[2])) or 0.0,
                                   str(vals[3]) if len(vals) > 3 else ""))
            if self.jv_editing_id is not None and not self.require_modify():
                return
            self.check_period_open(e_date)
            editing = self.jv_editing_id
            entry_id, total = save_journal(e_date, desc, lines_data, entry_id=editing)
            messagebox.showinfo("نجاح", f"تم {'تعديل' if editing else 'حفظ وترحيل'} القيد رقم {entry_id}!\n"
                                        f"إجمالي القيد: {fmt(total)} ر.س")
            self.new_journal_entry()
            self.load_journal_register()
        except Exception as e:
            handle_error(e)

    def new_journal_entry(self):
        self.jv_editing_id = None
        clear_tree(self.jv_tree)
        self.jv_desc_e.delete(0, tk.END)
        self.jv_date_e.delete(0, tk.END)
        self.jv_date_e.insert(0, today())
        self.jv_save_btn.config(text="💾 حفظ وترحيل القيد اليومي")

    def load_journal_register(self):
        if not hasattr(self, "jv_reg_tree"):
            return
        clear_tree(self.jv_reg_tree)
        with db() as conn:
            for eid, d, desc, n, tot, accs in conn.execute(
                    "SELECT e.entry_id, e.entry_date, COALESCE(e.description,''), COUNT(l.line_id), COALESCE(SUM(l.debit),0), "
                    "GROUP_CONCAT(l.account_name, '، ') FROM journal_entries e LEFT JOIN journal_lines l ON l.entry_id=e.entry_id "
                    "GROUP BY e.entry_id ORDER BY e.entry_date DESC, e.entry_id DESC"):
                self.jv_reg_tree.insert("", "end", iid=str(eid), values=(eid, d, desc, n, fmt(tot), accs or "-"))

    def load_journal_for_edit(self):
        if not self.require_modify():
            return
        sel = self.jv_reg_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر قيداً من السجل.")
            return
        eid = int(sel[0])
        try:
            with db() as conn:
                _block_system_entry(conn, eid)
        except Exception as e:
            handle_error(e)
            return
        with db() as conn:
            head = conn.execute("SELECT entry_date, COALESCE(description,'') FROM journal_entries WHERE entry_id=?", (eid,)).fetchone()
            lines = conn.execute("SELECT account_code, account_name, debit, credit, COALESCE(memo,'') FROM journal_lines "
                                 "WHERE entry_id=? ORDER BY line_id", (eid,)).fetchall()
        if not head:
            return
        self.new_journal_entry()
        self.jv_date_e.delete(0, tk.END)
        self.jv_date_e.insert(0, head[0])
        self.jv_desc_e.insert(0, head[1])
        for code, name, d, c, memo in lines:
            self.jv_tree.insert("", "end", values=(f"{code} - {name}", fmt(d), fmt(c), memo))
        self.jv_editing_id = eid
        self.jv_save_btn.config(text=f"💾 حفظ تعديلات القيد رقم {eid}")

    def delete_journal_selected(self):
        if not self.require_modify():
            return
        sel = self.jv_reg_tree.selection()
        if not sel:
            messagebox.showinfo("تنبيه", "اختر قيداً من السجل.")
            return
        if not messagebox.askyesno("حذف قيد", f"حذف القيد رقم {sel[0]} نهائياً؟"):
            return
        try:
            delete_journal_entry(int(sel[0]))
            if self.jv_editing_id == int(sel[0]):
                self.new_journal_entry()
            self.load_journal_register()
        except Exception as e:
            handle_error(e)

    def open_add_account_dialog(self):
        """نافذة منبثقة لإضافة حساب رئيسي أو فرعي جديد لشجرة الحسابات"""
        win = tk.Toplevel(self.root)
        win.title("إضافة حساب جديد لشجرة الحسابات")
        win.geometry("540x470")
        win.grab_set()
        
        f = ttk.LabelFrame(win, text=" بيانات الحساب الجديد ", padding=15)
        f.pack(fill="both", expand=True, padx=15, pady=15)

        # مستوى الحساب (رئيسي أم فرعي)
        ttk.Label(f, text="مستوى الحساب:").grid(row=0, column=1, sticky="e", pady=8)
        level_cb = ttk.Combobox(f, values=["حساب رئيسي", "حساب فرعي"], width=22, state="readonly")
        level_cb.set("حساب فرعي")
        level_cb.grid(row=0, column=0, sticky="w", pady=8)

        # الحساب الرئيسي (يظهر عند اختيار فرعي)
        ttk.Label(f, text="الحساب الرئيسي:").grid(row=1, column=1, sticky="e", pady=8)
        with db() as conn:
            parents = conn.execute("SELECT account_code, account_name FROM chart_of_accounts ORDER BY account_code").fetchall()
        parent_cb = SearchableCombobox(f, width=28)
        parent_cb.set_items([(f"{c} - {n}", (c, n)) for c, n in parents])
        parent_cb.grid(row=1, column=0, sticky="w", pady=8)

        # كود الحساب
        ttk.Label(f, text="كود الحساب:").grid(row=2, column=1, sticky="e", pady=8)
        code_entry = ttk.Entry(f, width=25)
        code_entry.grid(row=2, column=0, sticky="w", pady=8)

        # اسم الحساب
        ttk.Label(f, text="اسم الحساب:").grid(row=3, column=1, sticky="e", pady=8)
        name_entry = ttk.Entry(f, width=30)
        name_entry.grid(row=3, column=0, sticky="w", pady=8)

        # نوع الحساب (أنواع قياسية فقط، الافتراضي = نوع الحساب الرئيسي)
        ttk.Label(f, text="نوع الحساب:").grid(row=4, column=1, sticky="e", pady=8)
        type_cb = ttk.Combobox(f, values=ACCOUNT_TYPES, width=22, state="readonly")
        type_cb.set(TYPE_ASSET)
        type_cb.grid(row=4, column=0, sticky="w", pady=8)
        ttk.Label(f, text="مثال: لعامل جديد اختر الرئيسي 11410 ذمم العاملين أو 11420 عهد العاملين.",
                  foreground="#78716c").grid(row=6, column=0, columnspan=2, sticky="e")

        def on_level_changed(e=None):
            parent_cb.configure(state="normal" if level_cb.get() == "حساب فرعي" else "disabled")

        level_cb.bind("<<ComboboxSelected>>", on_level_changed)

        # اقتراح كود فريد ونوع الحساب بناءً على الحساب الرئيسي
        def on_parent_selected(e=None):
            sel = parent_cb.selected()
            if not sel:
                return
            p_code = sel[0]
            with db() as conn:
                p_type = conn.execute("SELECT account_type FROM chart_of_accounts WHERE account_code=?", (p_code,)).fetchone()
                existing = {r[0] for r in conn.execute("SELECT account_code FROM chart_of_accounts")}
                children = [r[0] for r in conn.execute("SELECT account_code FROM chart_of_accounts WHERE parent_code=?",
                                                       (p_code,)) if r[0].isdigit()]
            next_code = str(max(int(c) for c in children) + 1) if children else f"{p_code}01"
            while next_code in existing and next_code.isdigit():
                next_code = str(int(next_code) + 1)
            code_entry.delete(0, tk.END)
            code_entry.insert(0, next_code)
            if p_type and p_type[0] in ACCOUNT_TYPES:
                type_cb.set(p_type[0])

        parent_cb.bind("<<ComboboxSelected>>", on_parent_selected)

        def save_new_account():
            try:
                acc_code = normalize_digits(code_entry.get()).strip()
                acc_name = name_entry.get().strip()
                acc_type = type_cb.get().strip()
                parent = parent_cb.selected() if level_cb.get() == "حساب فرعي" else None
                parent_code = parent[0] if parent else None

                if not acc_code or not acc_name:
                    raise UserError("يرجى إدخال كود واسم الحساب بشكل صحيح.")
                if not acc_code.isdigit():
                    raise UserError("كود الحساب يجب أن يكون أرقاماً فقط.")
                if acc_type not in ACCOUNT_TYPES:
                    raise UserError("اختر نوع الحساب.")
                if level_cb.get() == "حساب فرعي" and not parent_code:
                    raise UserError("اختر الحساب الرئيسي من القائمة.")

                with db() as conn:
                    exists = conn.execute("SELECT 1 FROM chart_of_accounts WHERE account_code=?", (acc_code,)).fetchone()
                    if exists:
                        raise UserError(f"كود الحساب ({acc_code}) موجود مسبقاً، يرجى اختيار كود آخر.")
                    if parent_code:
                        p_has_moves = conn.execute("SELECT COUNT(*) FROM journal_lines WHERE account_code=?",
                                                   (parent_code,)).fetchone()[0]
                        p_has_moves += conn.execute("SELECT COUNT(*) FROM cash_transactions WHERE party_code=? AND category!=?",
                                                    (parent_code, CAT_PARTNER)).fetchone()[0]
                        if p_has_moves and not messagebox.askyesno(
                                "تنبيه", f"الحساب الرئيسي {parent_code} عليه حركات مرحّلة.\nبعد إضافة فرع له سيصبح حساباً "
                                         f"تجميعياً ولن يمكن الترحيل عليه مباشرة (حركاته السابقة تبقى كما هي).\n\nمتابعة؟",
                                parent=win):
                            return
                    conn.execute(
                        "INSERT INTO chart_of_accounts (account_code, account_name, account_type, parent_code, is_group) "
                        "VALUES (?, ?, ?, ?, ?)", (acc_code, acc_name, acc_type, parent_code, 0 if parent_code else 1))
                    if parent_code:
                        conn.execute("UPDATE chart_of_accounts SET is_group=1 WHERE account_code=?", (parent_code,))

                messagebox.showinfo("نجاح", f"تم إضافة الحساب ({acc_name}) بنجاح إلى شجرة الحسابات!", parent=win)
                win.destroy()
                self.load_coa_data()
                self.load_journal_accounts()
                self.update_party_options()
            except Exception as ex:
                handle_error(ex)

        money_button(f, "حفظ الحساب الجديد 💾", "#16a34a", save_new_account, 12).grid(row=5, column=0, columnspan=2, pady=15)



def main():
    if not HAS_TK:
        sys.exit("واجهة سطح المكتب تحتاج Tkinter غير المتوفر في هذه البيئة.\n"
                 "لتشغيل واجهة الويب استخدم:  streamlit run streamlit_app.py")
    init_db()
    auto_backup_daily()
    root = tk.Tk()
    root.withdraw()  # الواجهة الرئيسية لا تظهر قبل تسجيل الدخول
    state = {"note_shown": False}

    def open_app(user):
        root.deiconify()
        ElhagryERP(root, user, on_logout=start_login)
        if _engine.MIGRATION_NOTE and not state["note_shown"]:
            state["note_shown"] = True
            root.after(400, lambda: messagebox.showinfo("تحديث قاعدة البيانات", _engine.MIGRATION_NOTE))

    def start_login():
        LoginWindow(root, on_success=open_app)

    start_login()
    root.mainloop()

