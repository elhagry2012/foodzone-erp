# -*- coding: utf-8 -*-
"""
شاشات المرحلة الثانية في واجهة سطح المكتب (Mixin على ElhagryERP):
    📚 الدفتر العام والقوائم المالية (حالة الترحيل الآلي، إعادة الترحيل، حسابات الترحيل، ميزان شامل، ميزانية، مطابقة)
    🧭 التخطيط المالي (توقع التدفق النقدي، الموازنة التقديرية، عمولات المناديب، حدود الائتمان)
    🛡️ سجل التدقيق والرقابة الداخلية
    + فحص الحد الائتماني عند حفظ أوامر البيع وأذون التسليم.
المنطق كله في elhagry_core؛ هنا العرض والإدخال فقط.
"""
import elhagry_core as C


def _A():
    from elhagry_desktop import app
    return app


class Phase2Mixin:
    # ---------------------------------------------------------- عارض تقارير مضمّن قابل لإعادة الاستخدام
    def _report_box(self, parent, height=16):
        A = _A()
        box = A.ttk.Frame(parent)
        box.pack(fill="both", expand=True, padx=6, pady=4)
        bar = A.tk.Frame(box)
        bar.pack(fill="x")
        holder = {"rep": None}
        note = A.ttk.Label(box, text="", foreground=A.UI["muted"], justify="right", anchor="e", wraplength=1100)
        tree = A.make_tree(box, [("c0", "-", 120, "center")], height=height)
        note.pack(fill="x", pady=(4, 0))

        def export(kind):
            rep = holder["rep"]
            if not rep:
                A.messagebox.showinfo("تنبيه", "اعرض التقرير أولاً.")
                return
            ext, types = (".xlsx", [("Excel", "*.xlsx")]) if kind == "xlsx" else (".pdf", [("PDF files", "*.pdf")])
            path = A.filedialog.asksaveasfilename(defaultextension=ext, filetypes=types, initialfile=f"report_{A.today()}{ext}")
            if not path:
                return
            try:
                (A.export_report_xlsx if kind == "xlsx" else A.build_report_pdf)(path, rep)
                A.messagebox.showinfo("نجاح", "تم التصدير.")
            except Exception as e:
                A.handle_error(e)

        A.money_button(bar, "📗 Excel", "#0d9488", lambda: export("xlsx"), 9).pack(side="left", padx=3)
        A.money_button(bar, "📄 PDF", "#475569", lambda: export("pdf"), 9).pack(side="left", padx=3)
        return dict(tree=tree, note=note, holder=holder, bar=bar)

    def _show_report(self, box, rep):
        A = _A()
        box["holder"]["rep"] = rep
        A.configure_tree(box["tree"], rep.columns)
        for row in rep.rows:
            box["tree"].insert("", "end", values=[A.fmt_cell(v, k) for v, (_, k) in zip(row, rep.columns)])
        if rep.total:
            box["tree"].insert("", "end", values=[A.fmt_cell(v, k) for v, (_, k) in zip(rep.total, rep.columns)],
                               tags=("total",))
        box["note"].config(text="\n".join([f"{rep.title}  |  {rep.subtitle}"] + list(rep.notes)))

    def _dates_row(self, parent, row, d_from=None, d_to=None):
        A = _A()
        e1, e2 = A.DateEntry(parent, width=12), A.DateEntry(parent, width=12)
        e1.insert(0, d_from or f"{A.today()[:4]}-01-01")
        e2.insert(0, d_to or A.today())
        return e1, e2

    def _read_dates(self, e1, e2):
        A = _A()
        a, b = A.parse_date(e1.get()), A.parse_date(e2.get())
        if not a or not b:
            raise A.UserError("صيغة التاريخ YYYY-MM-DD")
        if a > b:
            raise A.UserError("تاريخ البداية بعد تاريخ النهاية.")
        return a, b

    # ---------------------------------------------------------- الحد الائتماني
    def _credit_override(self, err):
        """سياسة «منع»: المدير يوافق صراحة، وغيره يحتاج كلمة مرور المدير. يرجع True للمتابعة بتجاوز الحد."""
        A = _A()
        if self.user["is_admin"]:
            return A.messagebox.askyesno("تجاوز الحد الائتماني", f"{err}\n\nهل توافق (كمدير للنظام) على تجاوز الحد لهذه العملية؟",
                                         icon="warning")
        pwd = A.simpledialog.askstring("تجاوز الحد الائتماني", f"{err}\n\nللمتابعة أدخل كلمة مرور المدير:", show="*")
        if pwd is None:
            return False
        if not A._engine.verify_modify_password(pwd):
            raise A.UserError("كلمة المرور غير صحيحة. تم إلغاء العملية.")
        return True

    def credit_precheck(self, code, amount, stage):
        """تنبيه فوري قبل الحفظ لسياسة «تنبيه». يرجع False لو اختار المستخدم الإلغاء."""
        A = _A()
        st = C.credit_status(code, amount, stage)
        if st["exceeded"] and st["policy"] == C.POLICY_WARN:
            return A.messagebox.askyesno("تنبيه الحد الائتماني", st["message"] + "\n\nهل تريد المتابعة رغم التجاوز؟",
                                         icon="warning")
        return True

    def with_credit(self, fn, *args, **kwargs):
        """تنفيذ عملية بيع آجل؛ عند رفض المحرك (سياسة منع) يُطلب تجاوز إداري ثم تُعاد بإذن صريح."""
        try:
            return fn(*args, **kwargs)
        except C.CreditLimitError as e:
            if not self._credit_override(str(e)):
                raise _A().UserError("تم إلغاء العملية: تجاوز الحد الائتماني للعميل.")
            return fn(*args, allow_over_credit=True, **kwargs)

    # ---------------------------------------------------------- 📚 الدفتر العام
    def setup_ledger_tab(self, frame):
        A = _A()
        top = A.ttk.LabelFrame(frame, text=" حالة الترحيل الآلي إلى دفتر الأستاذ العام ", padding=10)
        top.pack(fill="x", padx=10, pady=6)
        self.gl_status_lbl = A.ttk.Label(top, text="", justify="right", anchor="e", font=(A.UI_FONT, 10))
        self.gl_status_lbl.pack(fill="x")
        btns = A.tk.Frame(top)
        btns.pack(fill="x", pady=(6, 0))
        A.money_button(btns, "🔄 تحديث الحالة", A.UI["primary"], self.load_ledger_page, 9).pack(side="right", padx=4)
        if self.user["is_admin"]:
            A.money_button(btns, "♻️ إعادة ترحيل كل المستندات", "#0d9488", self.ledger_rebuild, 9).pack(side="right", padx=4)
            A.money_button(btns, "⚙️ حسابات الترحيل الآلي", "#475569", self.open_gl_accounts_dialog, 9).pack(side="right", padx=4)
        self.gl_src_tree = A.make_tree(top, [("src", "المصدر", 220, "e"), ("n", "عدد القيود", 110, "center"),
                                             ("v", "إجمالي المدين", 150, "center")], height=6)

        rep = A.ttk.LabelFrame(frame, text=" القوائم المالية من الدفتر العام ", padding=8)
        rep.pack(fill="both", expand=True, padx=10, pady=4)
        f = A.tk.Frame(rep)
        f.pack(fill="x")
        self.gl_from_e, self.gl_to_e = self._dates_row(f, 0)
        self.gl_rep_cb = A.ttk.Combobox(f, state="readonly", width=40,
                                        values=[f"{k} | {C.REPORT_INDEX[k][0]}" for k in ("gl_tb", "bs", "gl_pl", "gl_recon",
                                                                                          "gl_journal")])
        self.gl_rep_cb.current(0)
        A.rtl_row(f, 0, [("التقرير:", self.gl_rep_cb), ("من:", self.gl_from_e), ("إلى:", self.gl_to_e),
                         (None, A.money_button(f, "عرض ▶", A.UI["primary"], self.run_ledger_report, 9))], slots=4)
        self.gl_box = self._report_box(rep, height=14)

    def load_ledger_page(self):
        A = _A()
        try:
            C.sync_ledger()
            st = C.ledger_status()
        except Exception as e:
            A.handle_error(e)
            return
        lines = [f"قيود الترحيل الآلي: {st['auto_entries']:,}   |   القيود اليدوية والآلية بالقيود اليومية: {st['manual_entries']:,}"
                 f"   |   بانتظار الترحيل: {st['pending']}",
                 f"إجمالي المدين = {st['debit']:,.2f}   |   إجمالي الدائن = {st['credit']:,.2f}   |   "
                 + ("الدفتر متوازن ✓" if st["balanced"] else "⚠ الدفتر غير متوازن"),
                 f"آخر إعادة ترحيل شاملة: {st['rebuilt_at'] or '-'}   |   الحساب المعلق {st['suspense_code']}: "
                 f"{st['suspense_balance']:,.2f}"]
        if st["last_error"]:
            lines.append(f"⚠ آخر خطأ ترحيل: {st['last_error']}")
        self.gl_status_lbl.config(text="\n".join(lines))
        A.clear_tree(self.gl_src_tree)
        for s, n, v in st["by_source"]:
            self.gl_src_tree.insert("", "end", values=(s, n, A.fmt(v)))

    def ledger_rebuild(self):
        A = _A()
        if not A.messagebox.askyesno("إعادة الترحيل", "سيُعاد بناء كل قيود الترحيل الآلي من المستندات (آمن، لا يمس المستندات). متابعة؟"):
            return
        try:
            C.rebuild_ledger()
            with C.db() as conn:
                C.audit_event(conn, C.ACT_EVENT, "الدفتر العام", "إعادة ترحيل شاملة")
            self.load_ledger_page()
            A.messagebox.showinfo("تم", "تمت إعادة ترحيل كل المستندات.")
        except Exception as e:
            A.handle_error(e)

    def run_ledger_report(self):
        A = _A()
        try:
            a, b = self._read_dates(self.gl_from_e, self.gl_to_e)
            key = self.gl_rep_cb.get().split(" | ")[0]
            self._show_report(self.gl_box, C.run_report(key, dict(d_from=a, d_to=b)))
        except Exception as e:
            A.handle_error(e)

    def open_gl_accounts_dialog(self):
        A = _A()
        win = A.tk.Toplevel(self.root)
        win.title("حسابات الترحيل الآلي")
        win.transient(self.root)
        with C.db() as conn:
            cur = C.gl_accounts(conn)
            accs = conn.execute("SELECT account_code, account_name FROM chart_of_accounts WHERE COALESCE(is_group,0)=0 "
                                "ORDER BY account_code").fetchall()
        values = [f"{c} - {n}" for c, n in accs]
        names = dict(accs)
        combos = {}
        for i, (key, (_d, label, _t, _p)) in enumerate(C.GL_ACCOUNT_DEFAULTS.items()):
            A.ttk.Label(win, text=label + ":").grid(row=i, column=1, sticky="e", padx=8, pady=3)
            cb = A.ttk.Combobox(win, values=values, width=48, state="readonly")
            code = cur.get(key)
            if code in names:
                cb.set(f"{code} - {names[code]}")
            cb.grid(row=i, column=0, padx=8, pady=3)
            combos[key] = cb

        def save():
            try:
                C.set_gl_accounts({k: cb.get().split(" - ")[0] for k, cb in combos.items() if cb.get()})
                win.destroy()
                self.load_ledger_page()
                A.messagebox.showinfo("تم", "تم حفظ الحسابات وإعادة ترحيل كل المستندات.")
            except Exception as e:
                A.handle_error(e)

        A.money_button(win, "💾 حفظ وإعادة الترحيل", "#16a34a", save, 10).grid(row=len(combos), column=0, columnspan=2, pady=10)

    # ---------------------------------------------------------- 🧭 التخطيط المالي
    def setup_planning_tab(self, frame):
        A = _A()
        nb = A.ttk.Notebook(frame)
        nb.pack(fill="both", expand=True, padx=8, pady=6)
        # توقع التدفق النقدي
        t1 = A.ttk.Frame(nb)
        nb.add(t1, text="💧 توقع التدفق النقدي")
        f = A.tk.Frame(t1)
        f.pack(fill="x", padx=6, pady=6)
        self.cf_to_e = A.DateEntry(f, width=12)
        self.cf_to_e.insert(0, (A.date.today() + A.timedelta(days=90)).isoformat())
        self.cf_gran_cb = A.ttk.Combobox(f, values=["شهري", "أسبوعي"], state="readonly", width=10)
        self.cf_gran_cb.set("شهري")
        self.cf_kind_cb = A.ttk.Combobox(f, values=["التوقع المجمع", "الفواتير المفتوحة بالاستحقاق"], state="readonly", width=26)
        self.cf_kind_cb.set("التوقع المجمع")
        A.rtl_row(f, 0, [("حتى تاريخ:", self.cf_to_e), ("التجميع:", self.cf_gran_cb), ("العرض:", self.cf_kind_cb),
                         (None, A.money_button(f, "عرض ▶", A.UI["primary"], self.run_cash_forecast, 9))], slots=4)
        self.cf_box = self._report_box(t1)
        # الموازنة التقديرية
        t2 = A.ttk.Frame(nb)
        nb.add(t2, text="🎯 الموازنة التقديرية مقابل الفعلي")
        ent = A.ttk.LabelFrame(t2, text=" إدخال الموازنة الشهرية لسنة ", padding=8)
        ent.pack(fill="x", padx=6, pady=4)
        self.bud_year_e = A.ttk.Entry(ent, width=8)
        self.bud_year_e.insert(0, A.today()[:4])
        self.bud_line_cb = A.SearchableCombobox(ent, width=60)
        A.rtl_row(ent, 0, [("السنة:", self.bud_year_e), ("البند:", self.bud_line_cb),
                           (None, A.money_button(ent, "تحميل", "#475569", self.load_budget_line, 9))], slots=3)
        months = A.tk.Frame(ent)
        months.grid(row=1, column=0, columnspan=8, pady=4)
        self.bud_month_e = []
        for i, m in enumerate(A.AR_MONTHS):
            A.ttk.Label(months, text=m).grid(row=(i // 6) * 2, column=5 - i % 6, padx=4)
            e = A.ttk.Entry(months, width=11)
            e.grid(row=(i // 6) * 2 + 1, column=5 - i % 6, padx=4, pady=2)
            self.bud_month_e.append(e)
        A.money_button(ent, "💾 حفظ موازنة البند", "#16a34a", self.save_budget_line, 9).grid(row=2, column=0, columnspan=8, pady=4)
        f = A.tk.Frame(t2)
        f.pack(fill="x", padx=6)
        self.bva_from_e, self.bva_to_e = self._dates_row(f, 0)
        A.rtl_row(f, 0, [("من:", self.bva_from_e), ("إلى:", self.bva_to_e),
                         (None, A.money_button(f, "عرض التقديري مقابل الفعلي ▶", A.UI["primary"], self.run_bva, 9))], slots=3)
        self.bva_box = self._report_box(t2, height=10)
        # العمولات
        t3 = A.ttk.Frame(nb)
        nb.add(t3, text="💼 عمولات المناديب")
        f = A.tk.Frame(t3)
        f.pack(fill="x", padx=6, pady=6)
        self.com_from_e, self.com_to_e = self._dates_row(f, 0, d_from=A.today()[:8] + "01")
        self.com_base_cb = A.ttk.Combobox(f, values=list(C.COMMISSION_BASES.values()), state="readonly", width=24)
        self.com_base_cb.set(C.COMMISSION_BASES[C.commission_base()])
        self.com_base_cb.bind("<<ComboboxSelected>>", lambda e: self.run_commissions())
        self.com_uniform_e = A.ttk.Entry(f, width=8)
        self.com_uniform_e.bind("<KeyRelease>", lambda e: self._com_schedule())
        self._com_rates, self._com_job = {}, None
        A.rtl_row(f, 0, [("من:", self.com_from_e), ("إلى:", self.com_to_e), ("أساس الاحتساب:", self.com_base_cb),
                         (None, A.money_button(f, "احتساب ▶", A.UI["primary"], self.run_commissions, 9))], slots=4)
        A.rtl_row(f, 1, [("نسبة موحدة للكل % (فارغ = نسبة كل مندوب):", self.com_uniform_e),
                         (None, A.money_button(f, "↺ إلغاء التعديلات", "#64748b", self.reset_commission_rates, 9)),
                         (None, A.money_button(f, "💾 حفظ النسب في سجل المناديب", "#16a34a", self.save_commission_rates, 9))],
                  slots=4)
        A.ttk.Label(t3, text="✏️ اضغط مرتين على أي مندوب لتعديل نسبته مباشرة؛ العمولة والإجماليات تُحتسب فوراً.",
                    foreground=A.UI["muted"]).pack(anchor="e", padx=10)
        self.com_total_lbl = A.ttk.Label(t3, text="", font=(A.UI_FONT, 11, "bold"), foreground=A.UI["primary_dk"])
        self.com_total_lbl.pack(anchor="e", padx=10)
        self.com_box = self._report_box(t3)
        self.com_box["tree"].bind("<Double-1>", self.edit_commission_rate)
        # حدود الائتمان
        t4 = A.ttk.Frame(nb)
        nb.add(t4, text="🛡️ حدود الائتمان")
        f = A.ttk.LabelFrame(t4, text=" الإعدادات الافتراضية (تُطبق على العملاء بدون سياسة/مدة خاصة) ", padding=8)
        f.pack(fill="x", padx=6, pady=6)
        self.cr_policy_cb = A.ttk.Combobox(f, values=list(C.CREDIT_POLICIES.values()), state="readonly", width=30)
        self.cr_policy_cb.set(C.CREDIT_POLICIES[C.default_policy()])
        self.cr_days_e = A.ttk.Entry(f, width=8)
        self.cr_days_e.insert(0, str(C.default_credit_days()))
        A.rtl_row(f, 0, [("السياسة الافتراضية:", self.cr_policy_cb), ("مدة السداد الافتراضية (يوم):", self.cr_days_e),
                         (None, A.money_button(f, "💾 حفظ", "#16a34a", self.save_credit_defaults, 9)),
                         (None, A.money_button(f, "عرض تقرير الحدود ▶", A.UI["primary"], self.run_credit_report, 9))], slots=4)
        if not self.user["is_admin"]:
            self.cr_policy_cb.state(["disabled"])
            self.cr_days_e.state(["disabled"])
        self.cr_box = self._report_box(t4)

    def load_planning_page(self):
        A = _A()
        try:
            choices = C.budget_line_choices()
            self.bud_line_cb.set_items([(lbl, k) for k, lbl in choices])
            if not self.bud_line_cb.get() and choices:
                self.bud_line_cb.set(choices[0][1])
        except Exception as e:
            A.handle_error(e)

    def _bud_key(self):
        key = self.bud_line_cb.selected()
        if not key:
            raise _A().UserError("اختر بند الموازنة من القائمة.")
        return key

    def load_budget_line(self):
        A = _A()
        try:
            year = int(self.bud_year_e.get())
            vals = C.budget_year_matrix(year).get(self._bud_key(), [None] * 12)
            for e, v in zip(self.bud_month_e, vals):
                e.delete(0, A.tk.END)
                if v is not None:
                    e.insert(0, f"{v:g}")
        except Exception as e:
            A.handle_error(e)

    def save_budget_line(self):
        A = _A()
        try:
            year = int(self.bud_year_e.get())
            amounts = []
            for e in self.bud_month_e:
                t = e.get().strip()
                v = A.parse_float(t) if t else None
                if t and v is None:
                    raise A.UserError("مبالغ الموازنة أرقام فقط.")
                amounts.append(v)
            C.save_budget_year(year, self._bud_key(), amounts, self.user["username"])
            A.messagebox.showinfo("تم", "تم حفظ الموازنة.")
        except Exception as e:
            A.handle_error(e)

    def run_bva(self):
        A = _A()
        try:
            a, b = self._read_dates(self.bva_from_e, self.bva_to_e)
            self._show_report(self.bva_box, C.run_report("bva", dict(d_from=a, d_to=b)))
        except Exception as e:
            A.handle_error(e)

    def run_cash_forecast(self):
        A = _A()
        try:
            to = A.parse_date(self.cf_to_e.get())
            if not to:
                raise A.UserError("صيغة التاريخ YYYY-MM-DD")
            key = "cf_forecast" if self.cf_kind_cb.get() == "التوقع المجمع" else "due_inv"
            p = dict(d_to=to, group="month" if self.cf_gran_cb.get() == "شهري" else "day")
            self._show_report(self.cf_box, C.run_report(key, p))
        except Exception as e:
            A.handle_error(e)

    def _com_uniform(self):
        A = _A()
        t = A.normalize_digits(self.com_uniform_e.get()).replace("%", "").strip()
        if not t:
            return None
        v = A.parse_float(t)
        if v is None or v < 0 or v > 100:
            raise A.UserError("النسبة الموحدة رقم بين 0 و 100.")
        return v

    def _com_schedule(self):
        if self._com_job:
            self.root.after_cancel(self._com_job)
        self._com_job = self.root.after(450, self.run_commissions)

    def run_commissions(self):
        """احتساب لحظي: الفترة + أساس الاحتساب + النسبة الموحدة أو النسب المعدلة لكل مندوب."""
        A = _A()
        self._com_job = None
        try:
            a, b = self._read_dates(self.com_from_e, self.com_to_e)
            base = next(k for k, v in C.COMMISSION_BASES.items() if v == self.com_base_cb.get())
            if base != C.commission_base() and self.user["is_admin"]:
                C.set_commission_base(base)
            p = dict(d_from=a, d_to=b, base=base, uniform=self._com_uniform(), rates=dict(self._com_rates))
            rep = C.run_report("commissions", p)
            self._show_report(self.com_box, rep)
            for iid, row in zip(self.com_box["tree"].get_children(), rep.rows):
                if row[1] in self._com_rates:
                    self.com_box["tree"].item(iid, tags=("edited",))
            self.com_box["tree"].tag_configure("edited", background="#ecfeff")
            self.com_total_lbl.config(text=f"إجمالي العمولات المستحقة: {A.fmt(rep.total[6])} ر.س   |   "
                                           f"أساس الاحتساب: {A.fmt(rep.total[5])}")
        except Exception as e:
            A.handle_error(e)

    def edit_commission_rate(self, event=None):
        A = _A()
        tree = self.com_box["tree"]
        iid = tree.identify_row(event.y) if event else tree.focus()
        rep = self.com_box["holder"]["rep"]
        if not iid or not rep:
            return
        idx = tree.index(iid)
        if idx >= len(rep.rows):
            return   # سطر الإجمالي
        name, cur = rep.rows[idx][1], rep.rows[idx][2]
        v = A.simpledialog.askfloat("نسبة العمولة", f"نسبة عمولة المندوب «{name}» (%):", initialvalue=cur,
                                    minvalue=0.0, maxvalue=100.0, parent=self.root)
        if v is None:
            return
        self._com_rates[name] = v
        self.run_commissions()

    def reset_commission_rates(self):
        self._com_rates = {}
        self.com_uniform_e.delete(0, _A().tk.END)
        self.run_commissions()

    def save_commission_rates(self):
        A = _A()
        try:
            rep = self.com_box["holder"]["rep"]
            if not rep:
                raise A.UserError("احتسب العمولات أولاً.")
            if not (self.user["is_admin"] or self.can("reps")):
                raise A.UserError("حفظ النسب في سجل المناديب يحتاج صلاحية «سجل المناديب».")
            rates = {r[1]: r[2] for r in rep.rows}
            n = C.save_commission_rates(rates)
            self._com_rates = {}
            self.com_uniform_e.delete(0, A.tk.END)
            A.messagebox.showinfo("تم", f"تم حفظ نسب {n} مندوب في سجل المناديب.")
            self.run_commissions()
        except Exception as e:
            A.handle_error(e)

    # ---------------------------------------------------------- التقارير التفاعلية والسداد السريع
    DRILL_STATEMENT = {"rec": "customer", "age_c": "customer", "credit": "customer", "pay": "supplier", "age_s": "supplier"}

    def drill_report_row(self, event=None):
        """الضغط المزدوج على طرف في تقارير المديونيات/الأعمار يفتح كشف حسابه، وعلى فاتورة يفتح تفاصيلها."""
        A = _A()
        key = getattr(self, "rep_current", None)
        tree = self.rep_tree
        iid = tree.identify_row(event.y) if event else tree.focus()
        if not iid or "total" in tree.item(iid, "tags"):
            return
        vals = tree.item(iid, "values")
        try:
            if key in A.PARTY_REPORTS and vals:
                # تقارير العملاء/الموردين/المقاصة: الكود في أول عمود (أو عمود المورد في المقاصة)؛ سطور الحركات بلا كود
                code = next((str(v).split(" ، ")[0] for v in vals[:2 if key == "pt_net" else 1]
                             if str(v).strip() and str(v) != "-"), None)
                if code:
                    self._trail_push_report()
                    self.open_statement_for(code, {"pt_cust": "customer", "pt_supp": "supplier"}.get(key, "both"))
            elif key in self.DRILL_STATEMENT and vals and vals[0]:
                self._trail_push_report()
                self.open_statement_for(vals[0], self.DRILL_STATEMENT[key])
            elif key in ("open_inv_c", "open_inv_s") and vals:
                ref = C.open_invoice_drill(key, vals)
                if ref:
                    self._trail_push_report()
                    self.show_transaction_details(ref)
            elif key == "due_inv" and len(vals) > 1:
                kind = "sales" if vals[0].startswith("تحصيل") else "purchase"
                inv_id = C.find_invoice_id(kind, vals[1])
                if inv_id:
                    self._trail_push_report()
                    self.show_transaction_details((kind, inv_id))
        except Exception as e:
            A.handle_error(e)

    def open_statement_for(self, code, scope="both"):
        A = _A()
        if not self.can("statement"):
            raise A.UserError("ليس لديك صلاحية «كشوف الحساب».")
        self.show_page("statement")
        self.update_statement_partners_dropdown()
        label = next((lbl for lbl, c in self.stmt_partner_cb._map.items() if c == code), None)
        if label:
            self.stmt_partner_cb.set(label)
        allowed = self._stmt_allowed_scope(scope)
        self.stmt_scope_cb.set(A.SUMMARY_SCOPES[allowed])
        self.show_partner_statement(code)
        name = (label or code).partition(" - ")[2] or code
        self.trail_push("statement", (code, allowed), f"📜 كشف {name}")

    def prefill_payment(self, kind, inv_id):
        """«سداد هذه الفاتورة»: فتح الخزينة بكل البيانات جاهزة (النوع، الطرف، الفاتورة، المتبقي، البيان)."""
        A = _A()
        try:
            pre = C.invoice_payment_prefill(kind, inv_id)
            if pre["remaining"] <= 0.005:
                raise A.UserError("الفاتورة مسددة بالكامل.")
            opt = next((v for v in self.c_type["values"] if str(v).startswith(pre["t_type"])), None)
            if opt is None:
                raise A.UserError(f"ليس لديك صلاحية سندات ال{pre['t_type']}.")
            self.trail_push("pay", (kind, inv_id), f"💳 سداد {pre['no']}")
            self.show_page("cash")          # يحمّل الصفحة الآن (متزامناً) قبل التعبئة
            self.c_type.set(opt)
            self.party_category_cb.set(C.CAT_PARTNER)
            self.update_party_options()
            label = next((lbl for lbl, v in self.party_name_cb._map.items() if v and v[0] == pre["party"][0]), None)
            if label is None:
                raise A.UserError(f"الطرف {pre['party'][0]} غير موجود في قائمة الشركاء.")
            self.party_name_cb.set(label)
            self.refresh_open_invoices()
            inv_label = next((lbl for lbl, v in self.open_inv_cb._map.items() if v == pre["alloc"]), None)
            if inv_label is None:
                raise A.UserError(f"الفاتورة {pre['no']} ليست مفتوحة للسداد.")
            self.open_inv_cb.set(inv_label)
            for w, v in ((self.c_amount, f"{pre['remaining']:.2f}"), (self.c_desc, pre["description"])):
                w.delete(0, A.tk.END)
                w.insert(0, v)
            self.c_date.delete(0, A.tk.END)
            self.c_date.insert(0, A.today())
            self.c_amount.focus_set()
            self.update_cash_rep_label()
            if getattr(self, "cash_pay_lbl", None) is not None:
                self.cash_pay_lbl.configure(
                    text=f"💳 سند {pre['t_type']} جاهز لسداد الفاتورة {pre['no']} — {label} — المبلغ المتبقي "
                         f"{A.fmt(pre['remaining'])} ر.س. راجع الحساب ثم اضغط «حفظ».")
        except Exception as e:
            A.handle_error(e)

    # ---------------------------------------------------------- 📎 مرفقات فاتورة المورد على استلام المخزن
    def pick_receipt_files(self):
        A = _A()
        paths = A.filedialog.askopenfilenames(
            title="اختر صورة / PDF فاتورة المورد",
            filetypes=[("صور و PDF", "*.pdf *.png *.jpg *.jpeg *.gif *.bmp *.webp *.tif *.tiff *.heic"),
                       ("كل الملفات المدعومة", " ".join(f"*{e}" for e in A.ATTACH_EXTS))])
        if not paths:
            return
        st = self.flow["purchase"]
        st["f_files"] = list(dict.fromkeys(list(st.get("f_files") or []) + list(paths)))
        names = "، ".join(A.os.path.basename(p) for p in st["f_files"])
        st["f_files_lbl"].configure(text=f"📎 {len(st['f_files'])} ملف جاهز للإرفاق: {names[:90]}",
                                    foreground=A.UI["primary_dk"])

    def attach_receipt_files(self, fid):
        """إرفاق الملفات المختارة بإذن الاستلام بعد حفظه (فشل ملف لا يلغي الاستلام)."""
        A = _A()
        st = self.flow["purchase"]
        n = 0
        for path in st.get("f_files") or []:
            try:
                with open(path, "rb") as fh:
                    C.add_receipt_attachment(fid, A.os.path.basename(path), fh.read())
                n += 1
            except Exception as e:
                A.handle_error(e)
        st["f_files"] = []
        if st.get("f_files_lbl") is not None:
            st["f_files_lbl"].configure(text="لم تُرفق فاتورة المورد بعد (اختياري: تُحفظ مع إذن الاستلام).",
                                        foreground=A.UI["muted"])
        return n

    def open_receipt_attachments(self):
        fid = self._selected_fulfilment("purchase")
        if fid:
            self.open_attachments_window("receipt", fid)

    def open_selected_supplier_attachments(self):
        A = _A()
        sel = self.pp["supplier"]["tree"].selection()
        if not sel:
            A.messagebox.showinfo("تنبيه", "اختر مورداً من القائمة أولاً.")
            return
        self.open_supplier_attachments(sel[0])

    def open_supplier_attachments(self, code):
        """كل مرفقات المورد (فواتيره الأصلية المرفوعة عند الاستلام وعلى فواتير المشتريات) في نافذة واحدة."""
        A = _A()
        rows = C.supplier_attachments(code)
        win = A.tk.Toplevel(self.root)
        win.title(f"مرفقات المورد {code}")
        win.geometry("1050x440")
        win.transient(self.root)
        A.tk.Label(win, text=f"📎 مرفقات المورد {code} — {len(rows)} ملف (اضغط مرتين لفتح الملف)",
                   font=(A.UI_FONT, 11, "bold"), fg=A.UI["primary_dk"]).pack(anchor="e", padx=10, pady=8)
        holder = A.tk.Frame(win)
        holder.pack(fill="both", expand=True, padx=10)
        t = A.make_tree(holder, [("doc", "المستند", 110, "center"), ("rc", "رقم الإذن", 80, "center"),
                                 ("inv", "رقم الفاتورة", 110, "center"), ("date", "التاريخ", 90, "center"),
                                 ("name", "اسم الملف", 300, "e"), ("size", "الحجم", 80, "center"),
                                 ("by", "بواسطة", 90, "center"), ("at", "وقت الرفع", 140, "center")], rtl=True)
        for aid, doc, rc, inv, d, name, size, by, at in rows:
            t.insert("", "end", iid=str(aid), values=(doc, rc, inv, d or "-", name, f"{size / 1024:,.0f} KB", by or "-", at))

        def open_selected(event=None):
            sel = t.selection()
            if not sel:
                return
            try:
                name, _mime, data = C.get_attachment(int(sel[0]))
                folder = A.os.path.join(A.tempfile.gettempdir(), "elhagry_attachments")
                A.os.makedirs(folder, exist_ok=True)
                path = A.os.path.join(folder, f"{sel[0]}_{name}")
                with open(path, "wb") as fh:
                    fh.write(data)
                try:
                    A.os.startfile(path)
                except (AttributeError, OSError):
                    A.messagebox.showinfo("المرفق", f"تم حفظ الملف مؤقتاً:\n{path}", parent=win)
            except Exception as e:
                A.handle_error(e)

        t.bind("<Double-1>", open_selected)
        bar = A.tk.Frame(win)
        bar.pack(fill="x", padx=10, pady=8)
        A.money_button(bar, "📂 فتح", A.UI["primary"], open_selected, 9).pack(side="right", padx=3)
        return win, t

    # ---------------------------------------------------------- 🧑‍💼 الأدوار الوظيفية
    def users_load_roles(self, role_id):
        roles = C.list_roles()
        self._roles = {f"{n} ({cnt} مستخدم)": (rid, perms) for rid, n, _d, perms, cnt in roles}
        none = "بدون دور (صلاحيات فردية)"
        self.u_role_cb.configure(values=[none] + list(self._roles))
        self.u_role_cb.set(next((lbl for lbl, (rid, _) in self._roles.items() if rid == role_id), none))

    def users_selected_role_id(self):
        v = getattr(self, "_roles", {}).get(self.u_role_cb.get())
        return v[0] if v else None

    def users_apply_role(self):
        """اختيار دور يملأ مربعات الصلاحيات بصلاحياته (ويمكن التعديل الفردي قبل الحفظ)."""
        v = getattr(self, "_roles", {}).get(self.u_role_cb.get())
        if not v or self.u_admin_var.get():
            return
        perms = set(v[1])
        for k, var in self.u_perm_vars.items():
            var.set(1 if k in perms else 0)
        for parent in C.SUB_PERMISSIONS:
            self.users_parent_toggled(parent, auto_fill=False)

    def save_role_from_form(self):
        A = _A()
        try:
            perms = [k for k in C.ALL_PERMISSION_KEYS if k in self.u_perm_vars and self.u_perm_vars[k].get()]
            cur = getattr(self, "_roles", {}).get(self.u_role_cb.get())
            rid, name = None, None
            if cur and A.messagebox.askyesno("الدور", f"تحديث الدور «{self.u_role_cb.get().rsplit(' (', 1)[0]}» "
                                             "بالصلاحيات المحددة؟\n(لا = إنشاء دور جديد)\n\n"
                                             "التحديث يُطبَّق فوراً على كل مستخدمي هذا الدور."):
                rid, name = cur[0], self.u_role_cb.get().rsplit(" (", 1)[0]
            if name is None:
                from tkinter import simpledialog
                name = simpledialog.askstring("دور جديد", "اسم الدور الوظيفي (مثال: أمين المخزن):", parent=self.root)
                if not name:
                    return
            rid, n = C.save_role(rid, name, perms)
            self.users_load_roles(rid)
            self.load_users()
            A.messagebox.showinfo("نجاح", f"تم حفظ الدور «{name}»." + (f"\nتم تحديث صلاحيات {n} مستخدم." if n else ""))
        except Exception as e:
            A.handle_error(e)

    def delete_selected_role(self):
        A = _A()
        rid = self.users_selected_role_id()
        if not rid:
            A.messagebox.showinfo("تنبيه", "اختر دوراً من قائمة «الدور الوظيفي» أولاً.")
            return
        if A.messagebox.askyesno("حذف دور", "حذف الدور؟ مستخدموه يحتفظون بصلاحياتهم الحالية كصلاحيات فردية."):
            C.delete_role(rid)
            self.users_load_roles(None)
            self.load_users()

    # ---------------------------------------------------------- 👤 ربط الأطراف بمناديبهم (للتوريث التلقائي في السندات)
    def fill_partner_reps_from_invoices(self):
        A = _A()
        try:
            sug = C.suggest_partner_reps()
            if not sug:
                A.messagebox.showinfo("المناديب", "لا يوجد عملاء/موردون بلا مندوب لهم فواتير مسجلة بمندوب.")
                return None
            lines = "\n".join(f"{code} - {name}  ←  {rep}  ({n} فاتورة، من {first})" for code, name, rep, n, first in sug[:25])
            more = f"\n... و{len(sug) - 25} آخرين" if len(sug) > 25 else ""
            if not A.messagebox.askyesno("ربط الأطراف بمناديبهم",
                                         f"بطاقات هذه الأطراف بلا مندوب، بينما فواتيرهم مسجلة بمندوب:\n\n{lines}{more}\n\n"
                                         "تسجيل المندوب في بطاقة كل طرف اعتباراً من أول فاتورة له؟\n"
                                         "(تُورَّث بعدها سنداته السابقة واللاحقة تلقائياً لنفس المندوب)"):
                return None
            n = C.apply_partner_reps(sug)
            for fn in ("load_reps", "update_cash_rep_label"):
                try:
                    getattr(self, fn)()
                except Exception:
                    pass
            A.messagebox.showinfo("تم", f"تم ربط {n} طرف بمندوبه وتوريث المندوب لسنداتهم.")
            return n
        except Exception as e:
            A.handle_error(e)
            return None

    # ---------------------------------------------------------- 🔗 تسوية الدفعات المقدمة
    def settle_invoice_advances(self, kind, inv_id, win=None):
        A = _A()
        try:
            with C.db() as conn:
                c = C.FLOW[kind]
                code = conn.execute(f"SELECT {c['pcode']} FROM {c['inv']} WHERE {c['iid']}=?", (inv_id,)).fetchone()[0]
            avail = C.party_advance_available(code, kind)
            rem = C.invoice_payment_prefill(kind, inv_id)["remaining"]
            if not A.messagebox.askyesno("تسوية من الدفعات المقدمة",
                                         f"المتاح من دفعات الطرف المقدمة: {A.fmt(avail)}\nالمتبقي على الفاتورة: {A.fmt(rem)}\n\n"
                                         f"تسوية {A.fmt(min(avail, rem))} على الفاتورة الآن؟ (لا يُنشأ قيد جديد)"):
                return None
            done = C.settle_invoice_from_advances(kind, inv_id)
            if win is not None:
                try:
                    win.destroy()
                except Exception:
                    pass
            self.show_transaction_details((kind, inv_id))
            self._after_settlement()
            A.messagebox.showinfo("تمت التسوية", "تمت تسوية " + A.fmt(sum(a for _, a in done)) + " من "
                                  + "، ".join(f"سند {t}" for t, _ in done) + ".")
            return done
        except Exception as e:
            A.handle_error(e)
            return None

    def settle_party_advances(self):
        """من شاشة الخزينة: تسوية كل دفعات الطرف المحدد على فواتيره المفتوحة (الأقدم أولاً)."""
        A = _A()
        try:
            party = self.party_name_cb.selected()
            if self.party_category_cb.get() != C.CAT_PARTNER or not party or not party[0]:
                raise A.UserError("اختر التصنيف «شريك» ثم الطرف أولاً.")
            kind = "sales" if self.c_type.get().startswith("قبض") else "purchase"
            avail = C.party_advance_available(party[0], kind)
            if avail <= 0.005:
                raise A.UserError("لا توجد دفعات مقدمة/تحت الحساب متاحة لهذا الطرف " +
                                  ("(سندات قبض غير مربوطة)." if kind == "sales" else "(سندات صرف غير مربوطة)."))
            if not A.messagebox.askyesno("تسوية الدفعات المقدمة",
                                         f"المتاح للتسوية: {A.fmt(avail)}\nسيتم توزيعه على أقدم الفواتير المفتوحة أولاً. متابعة؟"):
                return None
            done = C.auto_settle_party(party[0], kind)
            self.refresh_open_invoices()
            self._after_settlement()
            A.messagebox.showinfo("تمت التسوية", "\n".join(f"فاتورة {no}: {A.fmt(a)} (من سند {t})" for t, no, a in done))
            return done
        except Exception as e:
            A.handle_error(e)
            return None

    def _after_settlement(self):
        for fn in ("load_sales_invoices", "load_purchase_invoices", "_refresh_open_statement", "load_cash_transactions"):
            try:
                getattr(self, fn)()
            except Exception:
                pass

    # ---------------------------------------------------------- 🧭 مسار العمل (Navigation stack)
    # تقرير ← كشف حساب ← فاتورة ← سداد: كل خطوة زر في شريط أعلى الشاشة يعيدها بسياقها بنقرة واحدة.
    TRAIL_MAX = 8

    def build_trail_bar(self, parent):
        A = _A()
        self.trail = []
        self.trail_frame = A.tk.Frame(parent, bg=A.UI["bg"])
        self.trail_frame.pack(fill="x", side="top", padx=15, pady=(8, 0))
        self._render_trail()

    def trail_push(self, kind, key, label):
        """إضافة خطوة (أو تحديثها لو موجودة) دون حذف ما قبلها."""
        trail = getattr(self, "trail", None)
        if trail is None:
            return
        if not trail and kind != "report":
            ctx = self._trail_context()
            if ctx and ctx[:2] != (kind, key):
                trail.append(ctx)
        for i, s in enumerate(trail):
            if s[:2] == (kind, key):
                trail[i] = (kind, key, label)
                break
        else:
            trail.append((kind, key, label))
        del trail[:-self.TRAIL_MAX]
        self._trail_active = (kind, key)
        self._render_trail()

    def _trail_context(self):
        """الشاشة التي بدأ منها المستخدم (كشف معروض) تُضاف كخطوة أولى."""
        if getattr(self, "_active_page", None) == "statement" and getattr(self, "stmt_shown_code", None):
            code = self.stmt_shown_code
            return ("statement", (code, self._stmt_scope()), f"📜 كشف {code}")
        return None

    def _trail_push_report(self):
        key = getattr(self, "rep_current", None)
        if key and key in C.REPORT_INDEX:
            self.trail_push("report", key, "📊 " + C.REPORT_INDEX[key][0])

    def trail_open(self, i):
        A = _A()
        if not 0 <= i < len(self.trail):
            return
        kind, key, _ = self.trail[i]
        try:
            if kind == "report":
                self.show_page("reports")
                if self.rep_current != key:
                    self.rep_current = key
                    self.run_current_report()
                self._trail_active = (kind, key)
                self._render_trail()
            elif kind == "statement":
                self.open_statement_for(*key)
            elif kind == "invoice":
                self.show_transaction_details(key)
            elif kind == "pay":
                self.prefill_payment(*key)
        except Exception as e:
            A.handle_error(e)

    def trail_clear(self):
        self.trail.clear()
        self._render_trail()

    def _render_trail(self):
        A = _A()
        f = self.trail_frame
        for w in f.winfo_children():
            w.destroy()
        if len(self.trail) < 2:
            f.configure(height=1)
            return
        A.tk.Label(f, text="🧭 مسار العمل:", bg=A.UI["bg"], fg=A.UI["muted"], font=(A.UI_FONT, 9, "bold")
                   ).pack(side="right", padx=(0, 6))
        active = getattr(self, "_trail_active", None)
        for i, (kind, key, label) in enumerate(self.trail):
            if i:
                A.tk.Label(f, text="◀", bg=A.UI["bg"], fg=A.UI["muted"]).pack(side="right")
            on = (kind, key) == active
            A.tk.Button(f, text=label[:40], relief="flat", cursor="hand2", padx=10, pady=3,
                        bg=A.UI["primary"] if on else "#e2e8f0", fg="#ffffff" if on else "#0f172a",
                        font=(A.UI_FONT, 9, "bold"), command=lambda n=i: self.trail_open(n)).pack(side="right", padx=2)
        A.tk.Button(f, text="✖ مسح المسار", relief="flat", cursor="hand2", bg=A.UI["bg"], fg="#b91c1c",
                    font=(A.UI_FONT, 9), command=self.trail_clear).pack(side="left")

    # ---------------------------------------------------------- من الإذن إلى الفاتورة
    def _selected_fulfilment(self, kind):
        sel = self.flow[kind]["f_tree"].selection()
        if not sel:
            _A().messagebox.showinfo("تنبيه", "اختر إذناً من السجل.")
            return None
        return int(sel[0])

    def goto_invoice_for(self, kind, fid):
        """ينقل الإذن إلى شاشة الفاتورة ويعرض مسودة الفاتورة (البنود والضريبة والإجمالي) للمراجعة والاعتماد."""
        A = _A()
        if not fid:
            return
        c = A.FLOW[kind]
        try:
            with C.db() as conn:
                h = conn.execute(f"SELECT {c['pcode']}, {c['pname']}, invoice_id FROM {c['ful']} WHERE {c['fid']}=?",
                                 (fid,)).fetchone()
                if h is None:
                    raise A.UserError("الإذن غير موجود.")
                if h[2]:
                    raise A.UserError(f"الإذن رقم {fid} مفوتر بالفعل (فاتورة {A.invoice_number(kind, h[2])}).")
                if kind == "sales":
                    raw = conn.execute("SELECT item_code, item_name, unit, quantity, unit_price, COALESCE(discount_amount,0) "
                                       "FROM sales_delivery_lines WHERE delivery_id=? ORDER BY line_id", (fid,)).fetchall()
                else:
                    raw = conn.execute("SELECT item_code, item_name, unit, quantity, unit_cost, 0 FROM purchase_receipt_lines "
                                       "WHERE receipt_id=? ORDER BY line_id", (fid,)).fetchall()
            lines = [C.make_line(*r) for r in raw]
            if not self.flow_select(kind, "i"):
                raise A.UserError("إصدار الفواتير غير متاح لصلاحياتك.")
            st = self.flow[kind]
            st["u_type"].set(next(k for k, v in self.UNBILLED_TYPES.items() if v == kind))
            st["u_search"].delete(0, A.tk.END)
            self.load_unbilled(kind)
            iid = f"{kind}:{fid}"
            if st["u_tree"].exists(iid):
                st["u_tree"].selection_set(iid)
                st["u_tree"].see(iid)
            self._invoice_review(kind, fid, h, lines)
        except Exception as e:
            A.handle_error(e)

    def _invoice_review(self, kind, fid, head, lines):
        A = _A()
        st = self.flow[kind]
        win = A.tk.Toplevel(self.root)
        win.title("مراجعة الفاتورة قبل الاعتماد")
        win.transient(self.root)
        win.geometry("900x520")
        title = "فاتورة ضريبية" if kind == "sales" else "فاتورة مورد"
        A.ttk.Label(win, text=f"{title} من {'إذن التسليم' if kind == 'sales' else 'إذن الاستلام'} رقم {fid}  —  "
                              f"{head[0]} - {head[1]}", font=(A.UI_FONT, 12, "bold")).pack(anchor="e", padx=12, pady=8)
        box = A.tk.Frame(win)
        box.pack(fill="both", expand=True, padx=10)
        tree = A.make_tree(box, [("code", "الكود", 90, "center"), ("name", "الصنف", 240, "e"), ("qty", "الكمية", 80, "center"),
                                 ("price", "السعر", 90, "center"), ("disc", "الخصم", 80, "center"),
                                 ("sub", "قبل الضريبة", 110, "center"), ("vat", "الضريبة 15%", 100, "center"),
                                 ("tot", "الإجمالي", 110, "center")], height=9, rtl=True)
        for L in lines:
            tree.insert("", "end", values=(L["code"], L["name"], f"{L['qty']:g}", A.fmt(L["price"]), A.fmt(L["discount"]),
                                           A.fmt(L["subtotal"]), A.fmt(L["vat"]), A.fmt(L["total"])))
        sub, vat = sum(L["subtotal"] for L in lines), sum(L["vat"] for L in lines)
        A.ttk.Label(win, text=f"قبل الضريبة: {A.fmt(sub)}   |   الضريبة: {A.fmt(vat)}   |   الإجمالي: {A.fmt(sub + vat)} ر.س",
                    font=(A.UI_FONT, 11, "bold"), foreground=A.UI["primary_dk"]).pack(anchor="e", padx=12, pady=6)
        f = A.tk.Frame(win)
        f.pack(fill="x", padx=10)
        d_e = A.DateEntry(f, width=12)
        d_e.insert(0, A.today())
        n_e = A.ttk.Entry(f, width=36)
        items = [("تاريخ الفاتورة:", d_e), ("ملاحظات:", n_e)]
        r_e = None
        if kind == "purchase":
            r_e = A.ttk.Entry(f, width=20)
            items.append(("رقم فاتورة المورد:", r_e))
        A.rtl_row(f, 0, items, slots=3)

        def approve():
            for w, v in ((st["i_date"], d_e.get()), (st["i_notes"], n_e.get())):
                w.delete(0, A.tk.END)
                w.insert(0, v)
            if r_e is not None and "i_ref" in st:
                st["i_ref"].delete(0, A.tk.END)
                st["i_ref"].insert(0, r_e.get())
            st["u_tree"].selection_set(f"{kind}:{fid}")
            win.destroy()
            self.make_flow_invoice(kind)

        A.money_button(win, "✅ اعتماد وإصدار الفاتورة", "#16a34a", approve, 11).pack(pady=10)

    def save_credit_defaults(self):
        A = _A()
        try:
            pol = next(k for k, v in C.CREDIT_POLICIES.items() if v == self.cr_policy_cb.get())
            days = A.parse_float(self.cr_days_e.get())
            if days is None:
                raise A.UserError("مدة السداد رقم بالأيام.")
            C.set_credit_defaults(pol, int(days))
            A.messagebox.showinfo("تم", "تم حفظ إعدادات الائتمان الافتراضية.")
        except Exception as e:
            A.handle_error(e)

    def run_credit_report(self):
        A = _A()
        try:
            self._show_report(self.cr_box, C.run_report("credit", dict(d_from=A.today(), d_to=A.today())))
        except Exception as e:
            A.handle_error(e)

    # ---------------------------------------------------------- 🛡️ سجل التدقيق
    def setup_audit_tab(self, frame):
        A = _A()
        f = A.ttk.LabelFrame(frame, text=" سجل التدقيق والرقابة الداخلية — غير قابل للتعديل أو الحذف ", padding=8)
        f.pack(fill="x", padx=10, pady=6)
        self.au_from_e, self.au_to_e = self._dates_row(f, 0, d_from=(A.date.today() - A.timedelta(days=30)).isoformat())
        self.au_user_cb = A.ttk.Combobox(f, state="readonly", width=16)
        self.au_type_cb = A.ttk.Combobox(f, state="readonly", width=22)
        self.au_act_cb = A.ttk.Combobox(f, state="readonly", width=12, values=["", C.ACT_UPDATE, C.ACT_DELETE, C.ACT_EVENT])
        self.au_q_e = A.ttk.Entry(f, width=18)
        A.rtl_row(f, 0, [("من:", self.au_from_e), ("إلى:", self.au_to_e), ("المستخدم:", self.au_user_cb),
                         ("نوع المستند:", self.au_type_cb)], slots=4)
        A.rtl_row(f, 1, [("العملية:", self.au_act_cb), ("بحث (رقم/قيمة):", self.au_q_e),
                         (None, A.money_button(f, "عرض ▶", A.UI["primary"], self.load_audit_page, 9))], slots=4)
        self.au_chain_lbl = A.ttk.Label(f, text="", font=(A.UI_FONT, 10, "bold"))
        self.au_chain_lbl.grid(row=2, column=0, columnspan=8, sticky="e", padx=8)
        self.au_box = self._report_box(frame, height=18)

    def load_audit_page(self):
        A = _A()
        try:
            users, types = C.audit_filter_options()
            self.au_user_cb.configure(values=[""] + users)
            self.au_type_cb.configure(values=[""] + types)
            a, b = self._read_dates(self.au_from_e, self.au_to_e)
            rep = C.audit_report(d_from=a, d_to=b, user=self.au_user_cb.get(), doc_type=self.au_type_cb.get(),
                                 action=self.au_act_cb.get(), q=self.au_q_e.get().strip())
            self._show_report(self.au_box, rep)
            ok, n, bad = C.verify_audit_chain()
            self.au_chain_lbl.config(text=(f"سلسلة البصمات سليمة ✓ ({n:,} سجل)" if ok else
                                           f"⚠ تنبيه: تلاعب محتمل في السجل عند السجل رقم {bad}"),
                                     foreground="#0d9488" if ok else "#dc2626")
        except Exception as e:
            A.handle_error(e)
