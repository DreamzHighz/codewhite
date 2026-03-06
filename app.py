import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os

# ============================================================
#  ระบบดึงข้อมูลรายการ  (PostgreSQL + Excel)
# ============================================================

DB_DEFAULTS = {
    "host":     "172.16.16.1",
    "port":     "5434",
    "database": "imed_mfu",
    "username": "codewhite",
    "password": "MCH41509chaokuy",
}

COLUMNS = ("ชื่อรายการ", "วิธีการใช้", "ราคา")

COL_WIDTHS = {
    "ชื่อรายการ": 350,
    "วิธีการใช้": 300,
    "ราคา":      120,
}

# ─── helpers ────────────────────────────────────────────────

def try_import(name):
    try:
        __import__(name)
        return True
    except ImportError:
        return False


def make_scrollable_treeview(parent):
    frame = tk.Frame(parent, bg="#f0f4f8")
    frame.pack(fill="both", expand=True, padx=10, pady=6)

    style = ttk.Style()
    style.configure(
        "Custom.Treeview",
        rowheight=28,
        font=("Tahoma", 10),
        background="#ffffff",
        fieldbackground="#ffffff",
    )
    style.configure(
        "Custom.Treeview.Heading",
        font=("Tahoma", 10, "bold"),
        background="#2c6fad",
        foreground="white",
    )
    style.map("Custom.Treeview", background=[("selected", "#aed6f1")])

    tree = ttk.Treeview(
        frame,
        columns=COLUMNS,
        show="headings",
        style="Custom.Treeview",
        selectmode="browse",
    )
    for col in COLUMNS:
        tree.heading(col, text=col, anchor="w")
        tree.column(col, width=COL_WIDTHS[col], anchor="w", minwidth=80)

    vsb = ttk.Scrollbar(frame, orient="vertical",   command=tree.yview)
    hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
    tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    tree.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    frame.rowconfigure(0, weight=1)
    frame.columnconfigure(0, weight=1)

    return tree


def populate_tree(tree, rows, status_var, count_var):
    for item in tree.get_children():
        tree.delete(item)
    for i, row in enumerate(rows):
        tag = "even" if i % 2 == 0 else "odd"
        tree.insert("", "end", values=row, tags=(tag,))
    tree.tag_configure("even", background="#ffffff")
    tree.tag_configure("odd",  background="#eaf4fb")
    count_var.set(f"พบข้อมูล {len(rows)} รายการ")
    status_var.set("สำเร็จ ✔")


# ============================================================
#  Main Application
# ============================================================

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ระบบดึงข้อมูลรายการยา / Item Viewer")
        self.geometry("1100x720")
        self.configure(bg="#e8f0fe")
        self.resizable(True, True)

        self._build_ui()

    # ── UI shell ─────────────────────────────────────────────

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self, bg="#1a5276", height=56)
        hdr.pack(fill="x")
        tk.Label(
            hdr,
            text="  ระบบดึงข้อมูลรายการ  |  Item Data Viewer",
            font=("Tahoma", 14, "bold"),
            bg="#1a5276",
            fg="white",
        ).pack(side="left", padx=14, pady=10)

        # Notebook (tabs)
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure(
            "TNotebook.Tab",
            font=("Tahoma", 11, "bold"),
            padding=[14, 6],
            background="#cdd7e0",
            foreground="#1a3c5e",
        )
        style.map(
            "TNotebook.Tab",
            background=[("selected", "#2c6fad")],
            foreground=[("selected", "white")],
        )

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=8)

        self._build_pg_tab(nb)
        self._build_excel_tab(nb)

    # ── PostgreSQL Tab ────────────────────────────────────────

    def _build_pg_tab(self, nb):
        tab = tk.Frame(nb, bg="#f0f4f8")
        nb.add(tab, text="  🗄  PostgreSQL  ")

        # Connection frame
        cf = tk.LabelFrame(
            tab, text="  การเชื่อมต่อฐานข้อมูล",
            font=("Tahoma", 10, "bold"), bg="#f0f4f8",
            fg="#1a5276", bd=2, relief="groove",
        )
        cf.pack(fill="x", padx=12, pady=(10, 4))

        labels  = ["Host / Server", "Port", "Database", "Username", "Password"]
        keys    = ["host", "port", "database", "username", "password"]
        self._pg_vars = {}

        for col, (lbl, key) in enumerate(zip(labels, keys)):
            tk.Label(cf, text=lbl + ":", font=("Tahoma", 9, "bold"),
                     bg="#f0f4f8", fg="#2c3e50").grid(
                row=0, column=col * 2, padx=(10, 2), pady=8, sticky="e")
            var = tk.StringVar(value=DB_DEFAULTS[key])
            self._pg_vars[key] = var
            show = "*" if key == "password" else ""
            w = 10 if key in ("port",) else 18
            ttk.Entry(cf, textvariable=var, show=show, width=w).grid(
                row=0, column=col * 2 + 1, padx=(0, 8), pady=8, sticky="w")

        # Query frame
        qf = tk.LabelFrame(
            tab, text="  SQL Query",
            font=("Tahoma", 10, "bold"), bg="#f0f4f8",
            fg="#1a5276", bd=2, relief="groove",
        )
        qf.pack(fill="x", padx=12, pady=4)

        self._pg_query = tk.Text(
            qf, height=4, font=("Courier New", 10),
            relief="flat", bd=4, bg="#1e2a38", fg="#a9d1f7",
            insertbackground="white", wrap="word",
        )
        self._pg_query.pack(fill="x", padx=8, pady=6)
        self._pg_query.insert(
            "1.0",
            "SELECT item_name AS \"ชื่อรายการ\",\n"
            "       usage_method AS \"วิธีการใช้\",\n"
            "       price AS \"ราคา\"\n"
            "FROM   items\n"
            "ORDER  BY item_name\n"
            "LIMIT  500;",
        )

        # Button row
        btn_row = tk.Frame(tab, bg="#f0f4f8")
        btn_row.pack(fill="x", padx=12, pady=4)

        self._pg_status = tk.StringVar(value="")
        self._pg_count  = tk.StringVar(value="")

        tk.Button(
            btn_row, text="  ▶  ดึงข้อมูล (Execute)",
            font=("Tahoma", 10, "bold"),
            bg="#2c6fad", fg="white", activebackground="#1a5276",
            relief="flat", padx=12, pady=5, cursor="hand2",
            command=self._fetch_pg_thread,
        ).pack(side="left")

        tk.Button(
            btn_row, text="  📋  แสดงตาราง (Show Tables)",
            font=("Tahoma", 10),
            bg="#27ae60", fg="white", activebackground="#1e8449",
            relief="flat", padx=12, pady=5, cursor="hand2",
            command=self._show_tables,
        ).pack(side="left", padx=8)

        tk.Label(btn_row, textvariable=self._pg_status, font=("Tahoma", 10),
                 bg="#f0f4f8", fg="#27ae60").pack(side="left", padx=10)
        tk.Label(btn_row, textvariable=self._pg_count, font=("Tahoma", 10),
                 bg="#f0f4f8", fg="#7f8c8d").pack(side="right", padx=10)

        # Results
        self._pg_tree = make_scrollable_treeview(tab)

    # ── Excel Tab ─────────────────────────────────────────────

    def _build_excel_tab(self, nb):
        tab = tk.Frame(nb, bg="#f0f4f8")
        nb.add(tab, text="  📊  Excel File  ")

        # File picker
        ff = tk.LabelFrame(
            tab, text="  เลือกไฟล์ Excel",
            font=("Tahoma", 10, "bold"), bg="#f0f4f8",
            fg="#1a5276", bd=2, relief="groove",
        )
        ff.pack(fill="x", padx=12, pady=(10, 4))

        self._xls_path = tk.StringVar()
        path_entry = ttk.Entry(ff, textvariable=self._xls_path, width=70)
        path_entry.grid(row=0, column=0, padx=(10, 4), pady=8, sticky="ew")
        ff.columnconfigure(0, weight=1)

        tk.Button(
            ff, text="  📂  Browse",
            font=("Tahoma", 10, "bold"), bg="#e67e22", fg="white",
            activebackground="#d35400", relief="flat",
            padx=10, pady=4, cursor="hand2",
            command=self._browse_excel,
        ).grid(row=0, column=1, padx=4, pady=8)

        # Column mapping
        mf = tk.LabelFrame(
            tab, text="  จับคู่คอลัมน์ (Column Mapping)",
            font=("Tahoma", 10, "bold"), bg="#f0f4f8",
            fg="#1a5276", bd=2, relief="groove",
        )
        mf.pack(fill="x", padx=12, pady=4)

        mapping_labels = ["ชื่อรายการ", "วิธีการใช้", "ราคา"]
        self._col_vars  = {}
        self._col_combos = {}

        for i, lbl in enumerate(mapping_labels):
            tk.Label(mf, text=lbl + " →", font=("Tahoma", 9, "bold"),
                     bg="#f0f4f8", fg="#2c3e50").grid(
                row=0, column=i * 2, padx=(14, 4), pady=8, sticky="e")
            var = tk.StringVar()
            self._col_vars[lbl] = var
            cb = ttk.Combobox(mf, textvariable=var, width=22, state="readonly")
            cb.grid(row=0, column=i * 2 + 1, padx=(0, 12), pady=8, sticky="w")
            self._col_combos[lbl] = cb

        tk.Label(mf, text="(จะอัปเดตอัตโนมัติเมื่อเลือกไฟล์)",
                 font=("Tahoma", 8), bg="#f0f4f8", fg="#95a5a6").grid(
            row=1, column=0, columnspan=6, padx=14, pady=(0, 6), sticky="w")

        # Button row
        btn_row = tk.Frame(tab, bg="#f0f4f8")
        btn_row.pack(fill="x", padx=12, pady=4)

        self._xls_status = tk.StringVar(value="")
        self._xls_count  = tk.StringVar(value="")

        tk.Button(
            btn_row, text="  ▶  โหลดข้อมูล (Load)",
            font=("Tahoma", 10, "bold"),
            bg="#2c6fad", fg="white", activebackground="#1a5276",
            relief="flat", padx=12, pady=5, cursor="hand2",
            command=self._fetch_excel_thread,
        ).pack(side="left")

        tk.Label(btn_row, textvariable=self._xls_status, font=("Tahoma", 10),
                 bg="#f0f4f8", fg="#27ae60").pack(side="left", padx=10)
        tk.Label(btn_row, textvariable=self._xls_count, font=("Tahoma", 10),
                 bg="#f0f4f8", fg="#7f8c8d").pack(side="right", padx=10)

        # Results
        self._xls_tree = make_scrollable_treeview(tab)

    # ── PostgreSQL logic ──────────────────────────────────────

    def _fetch_pg_thread(self):
        self._pg_status.set("กำลังเชื่อมต่อ…")
        self._pg_count.set("")
        t = threading.Thread(target=self._fetch_pg, daemon=True)
        t.start()

    def _fetch_pg(self):
        if not try_import("psycopg2"):
            messagebox.showerror(
                "ไม่พบ Library",
                "ไม่พบ psycopg2\nโปรดติดตั้งด้วยคำสั่ง:\n  pip install psycopg2-binary",
            )
            self._pg_status.set("ผิดพลาด ✘")
            return
        import psycopg2

        try:
            conn = psycopg2.connect(
                host=self._pg_vars["host"].get(),
                port=int(self._pg_vars["port"].get()),
                dbname=self._pg_vars["database"].get(),
                user=self._pg_vars["username"].get(),
                password=self._pg_vars["password"].get(),
                connect_timeout=10,
            )
            cur = conn.cursor()
            query = self._pg_query.get("1.0", "end").strip()
            cur.execute(query)
            col_names = [desc[0] for desc in cur.description]
            raw_rows  = cur.fetchall()
            conn.close()

            # map first 3 columns to display
            rows = []
            for row in raw_rows:
                r = list(row) + ["", "", ""]
                rows.append((str(r[0] or ""), str(r[1] or ""), str(r[2] or "")))

            self.after(0, lambda: populate_tree(
                self._pg_tree, rows, self._pg_status, self._pg_count))

        except Exception as exc:
            self._pg_status.set("ผิดพลาด ✘")
            self.after(0, lambda: messagebox.showerror("ข้อผิดพลาด", str(exc)))

    def _show_tables(self):
        self._pg_status.set("กำลังดึงรายชื่อตาราง…")
        t = threading.Thread(target=self._do_show_tables, daemon=True)
        t.start()

    def _do_show_tables(self):
        if not try_import("psycopg2"):
            messagebox.showerror("ไม่พบ Library", "ไม่พบ psycopg2\nโปรดติดตั้ง: pip install psycopg2-binary")
            return
        import psycopg2
        try:
            conn = psycopg2.connect(
                host=self._pg_vars["host"].get(),
                port=int(self._pg_vars["port"].get()),
                dbname=self._pg_vars["database"].get(),
                user=self._pg_vars["username"].get(),
                password=self._pg_vars["password"].get(),
                connect_timeout=10,
            )
            cur = conn.cursor()
            cur.execute(
                "SELECT table_schema, table_name "
                "FROM information_schema.tables "
                "WHERE table_type='BASE TABLE' "
                "  AND table_schema NOT IN ('pg_catalog','information_schema') "
                "ORDER BY table_schema, table_name;"
            )
            tables = cur.fetchall()
            conn.close()
            table_list = "\n".join(f"  {s}.{t}" for s, t in tables) or "(ไม่พบตาราง)"
            self._pg_status.set("สำเร็จ ✔")
            self.after(0, lambda: messagebox.showinfo(
                "รายชื่อตาราง", f"พบ {len(tables)} ตาราง:\n\n{table_list}"))
        except Exception as exc:
            self._pg_status.set("ผิดพลาด ✘")
            self.after(0, lambda: messagebox.showerror("ข้อผิดพลาด", str(exc)))

    # ── Excel logic ───────────────────────────────────────────

    def _browse_excel(self):
        path = filedialog.askopenfilename(
            title="เลือกไฟล์ Excel",
            filetypes=[("Excel Files", "*.xlsx *.xls *.xlsb *.xlsm"), ("All Files", "*.*")],
        )
        if not path:
            return
        self._xls_path.set(path)
        self._load_excel_columns(path)

    def _load_excel_columns(self, path):
        if not try_import("pandas"):
            messagebox.showerror("ไม่พบ Library", "ไม่พบ pandas\nโปรดติดตั้ง: pip install pandas openpyxl")
            return
        import pandas as pd
        try:
            df = pd.read_excel(path, nrows=0)
            cols = list(df.columns)
            for cb in self._col_combos.values():
                cb["values"] = cols
                cb.set("")

            # auto-guess columns by keyword
            guess_map = {
                "ชื่อรายการ": ["name", "item", "ชื่อ", "รายการ", "drug", "ยา", "สินค้า"],
                "วิธีการใช้":  ["usage", "use", "วิธี", "route", "direction"],
                "ราคา":       ["price", "cost", "ราคา", "amount", "บาท"],
            }
            for field, keywords in guess_map.items():
                for col in cols:
                    col_lower = str(col).lower()
                    if any(kw in col_lower for kw in keywords):
                        self._col_combos[field].set(col)
                        break

            self._xls_status.set(f"โหลด Header สำเร็จ ({len(cols)} คอลัมน์)")
        except Exception as exc:
            messagebox.showerror("ข้อผิดพลาด", str(exc))

    def _fetch_excel_thread(self):
        self._xls_status.set("กำลังโหลด…")
        self._xls_count.set("")
        t = threading.Thread(target=self._fetch_excel, daemon=True)
        t.start()

    def _fetch_excel(self):
        if not try_import("pandas"):
            messagebox.showerror("ไม่พบ Library", "ไม่พบ pandas\nโปรดติดตั้ง: pip install pandas openpyxl")
            self._xls_status.set("ผิดพลาด ✘")
            return
        import pandas as pd

        path = self._xls_path.get()
        if not path or not os.path.isfile(path):
            self.after(0, lambda: messagebox.showwarning("คำเตือน", "กรุณาเลือกไฟล์ Excel ก่อน"))
            self._xls_status.set("")
            return

        name_col  = self._col_vars["ชื่อรายการ"].get()
        usage_col = self._col_vars["วิธีการใช้"].get()
        price_col = self._col_vars["ราคา"].get()

        if not any([name_col, usage_col, price_col]):
            self.after(0, lambda: messagebox.showwarning("คำเตือน", "กรุณาเลือกคอลัมน์ที่ต้องการแสดง"))
            self._xls_status.set("")
            return

        try:
            df = pd.read_excel(path, dtype=str)
            df.fillna("", inplace=True)

            rows = []
            for _, row in df.iterrows():
                name  = row.get(name_col,  "") if name_col  else ""
                usage = row.get(usage_col, "") if usage_col else ""
                price = row.get(price_col, "") if price_col else ""
                rows.append((name, usage, price))

            self.after(0, lambda: populate_tree(
                self._xls_tree, rows, self._xls_status, self._xls_count))

        except Exception as exc:
            self._xls_status.set("ผิดพลาด ✘")
            self.after(0, lambda: messagebox.showerror("ข้อผิดพลาด", str(exc)))


# ── entry point ──────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
