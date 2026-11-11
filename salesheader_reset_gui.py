#! python3.11
# SalesHeader Interface Reset Tool — tkinter GUI
# Programmer: Rykan Collins
# Look up errors for an ExternalUID, optionally correct fields, then reset to NEW.

import configparser
import os
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import pyodbc

_cfg_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "db.ini")
_cfg = configparser.ConfigParser()
_cfg.read(_cfg_path)
_sec = _cfg["sqlserver"]
CONN_STR = (
    "Driver=ODBC Driver 17 for SQL Server;"
    f"Server={_sec['server']};"
    f"Database={_sec['database']};"
    f"UID={_sec['uid']};"
    f"PWD={_sec['pwd']};"
)

# ── SQL ────────────────────────────────────────────────────────────────────────

SQL_HEADER = """
SELECT
    shi.IDSalesHeaderInterface,
    shi.ExternalUID,
    shi.InterfaceStatus,
    ISNULL(shi.InterfaceErrorText, '')  AS InterfaceErrorText,
    ISNULL(shi.orderid, '')             AS OrderID,
    ISNULL(she.ErrorText, '')           AS StatusErrorText,
    CONVERT(varchar(20), she.ErrorDate, 120) AS ErrorDate
FROM dbo.SalesHeader_Interface shi
LEFT JOIN dbo.SalesHeader_InterfaceError she
    ON she.IDSalesHeaderInterface = shi.IDSalesHeaderInterface
WHERE shi.ExternalUID = ?
ORDER BY she.ErrorDate DESC
"""

SQL_DETAIL = """
SELECT
    sdi.IDSalesDetailInterface,
    sdi.InterfaceStatus,
    ISNULL(sdi.InterfaceErrorText, '')  AS InterfaceErrorText,
    ISNULL(sdi.orderid, '')             AS OrderID,
    ISNULL(sde.ErrorText, '')           AS StatusErrorText,
    CONVERT(varchar(20), sde.ErrorDate, 120) AS ErrorDate
FROM dbo.SalesDetail_Interface sdi
LEFT JOIN dbo.SalesDetail_InterfaceError sde
    ON sde.IDSalesDetailInterface = sdi.IDSalesDetailInterface
WHERE sdi.IDSalesHeaderInterface IN (
    SELECT IDSalesHeaderInterface FROM dbo.SalesHeader_Interface WHERE ExternalUID = ?
)
ORDER BY sde.ErrorDate DESC
"""


# ── App ────────────────────────────────────────────────────────────────────────

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("SalesHeader Interface Reset Tool")
        self.geometry("1050x720")
        self.resizable(True, True)
        self.configure(bg="#f0f0f0")
        self._conn = None
        self._header_rows = []
        self._detail_rows = []
        self._build_ui()

    # ── connection ─────────────────────────────────────────────────────────────

    def _get_conn(self):
        if self._conn is None:
            self._conn = pyodbc.connect(CONN_STR)
        return self._conn

    # ── UI construction ────────────────────────────────────────────────────────

    def _build_ui(self):
        # ── lookup bar ──
        top = tk.Frame(self, bg="#f0f0f0", pady=8, padx=10)
        top.pack(fill="x")

        tk.Label(top, text="ExternalUID:", bg="#f0f0f0", font=("Arial", 11, "bold")).pack(side="left")
        self.uid_var = tk.StringVar()
        uid_entry = tk.Entry(top, textvariable=self.uid_var, width=20, font=("Arial", 11))
        uid_entry.pack(side="left", padx=(6, 10))
        uid_entry.bind("<Return>", lambda _: self._lookup())

        tk.Button(top, text="Look Up", width=10, command=self._lookup,
                  bg="#2d7dd2", fg="white", font=("Arial", 10, "bold")).pack(side="left")
        tk.Button(top, text="Clear", width=8, command=self._clear,
                  bg="#888", fg="white", font=("Arial", 10)).pack(side="left", padx=4)

        # ── notebook (Header / Detail tabs) ──
        nb_frame = tk.Frame(self, bg="#f0f0f0", padx=10)
        nb_frame.pack(fill="both", expand=True)

        self.nb = ttk.Notebook(nb_frame)
        self.nb.pack(fill="both", expand=True)

        # Header tab
        h_tab = tk.Frame(self.nb, bg="white")
        self.nb.add(h_tab, text="Header Errors")
        self.h_tree = self._make_tree(h_tab, [
            ("IDSalesHeaderInterface", 160),
            ("ExternalUID",            90),
            ("InterfaceStatus",        110),
            ("InterfaceErrorText",     220),
            ("OrderID",                80),
            ("StatusErrorText",        230),
            ("ErrorDate",              130),
        ])
        self.h_tree.bind("<<TreeviewSelect>>", self._on_header_select)

        # Detail tab
        d_tab = tk.Frame(self.nb, bg="white")
        self.nb.add(d_tab, text="Detail Errors")
        self.d_tree = self._make_tree(d_tab, [
            ("IDSalesDetailInterface", 160),
            ("InterfaceStatus",        110),
            ("InterfaceErrorText",     220),
            ("OrderID",                80),
            ("StatusErrorText",        230),
            ("ErrorDate",              130),
        ])
        self.d_tree.bind("<<TreeviewSelect>>", self._on_detail_select)

        # ── edit panel ──
        edit = tk.LabelFrame(self, text=" Correct Selected Record ",
                             bg="#f0f0f0", font=("Arial", 10, "bold"),
                             padx=8, pady=6)
        edit.pack(fill="x", padx=10, pady=(4, 0))

        # Row 1 — record id + status
        r1 = tk.Frame(edit, bg="#f0f0f0")
        r1.pack(fill="x", pady=2)
        tk.Label(r1, text="Record ID:", bg="#f0f0f0", width=14, anchor="w").pack(side="left")
        self.rec_id_var = tk.StringVar()
        tk.Entry(r1, textvariable=self.rec_id_var, state="readonly",
                 width=20, bg="#e8e8e8").pack(side="left", padx=(0, 20))
        tk.Label(r1, text="Interface Status:", bg="#f0f0f0", width=16, anchor="w").pack(side="left")
        self.status_var = tk.StringVar()
        status_cb = ttk.Combobox(r1, textvariable=self.status_var, width=12,
                                  values=["NEW", "ERROR", "PROCESSING", "COMPLETE"])
        status_cb.pack(side="left")

        # Row 2 — error text
        r2 = tk.Frame(edit, bg="#f0f0f0")
        r2.pack(fill="x", pady=2)
        tk.Label(r2, text="Error Text:", bg="#f0f0f0", width=14, anchor="w").pack(side="left")
        self.err_text_var = tk.StringVar()
        tk.Entry(r2, textvariable=self.err_text_var, width=80).pack(side="left")

        # Row 3 — edit buttons
        r3 = tk.Frame(edit, bg="#f0f0f0")
        r3.pack(fill="x", pady=(4, 0))
        self.rec_type_var = tk.StringVar(value="")  # "header" or "detail"
        tk.Button(r3, text="Save Correction", width=16, command=self._save_correction,
                  bg="#f0a500", fg="white", font=("Arial", 10, "bold")).pack(side="left", padx=(0, 8))
        tk.Label(r3, text="(Updates the selected record only — does not reset to NEW)",
                 bg="#f0f0f0", fg="#555", font=("Arial", 9)).pack(side="left")

        # ── action buttons ──
        act = tk.Frame(self, bg="#f0f0f0", pady=6, padx=10)
        act.pack(fill="x")
        tk.Button(act, text="Reset ALL to NEW", width=18, command=self._reset_all,
                  bg="#c0392b", fg="white", font=("Arial", 11, "bold")).pack(side="left", padx=(0, 10))
        tk.Label(act, text="Deletes all status/error rows and sets InterfaceStatus = NEW for every record under this UID",
                 bg="#f0f0f0", fg="#555", font=("Arial", 9)).pack(side="left")

        # ── log ──
        log_frame = tk.LabelFrame(self, text=" Activity Log ",
                                  bg="#f0f0f0", font=("Arial", 10, "bold"),
                                  padx=6, pady=4)
        log_frame.pack(fill="x", padx=10, pady=(4, 8))
        self.log = scrolledtext.ScrolledText(log_frame, height=5, state="disabled",
                                              font=("Courier New", 9), bg="#1e1e1e", fg="#d4d4d4")
        self.log.pack(fill="x")

    def _make_tree(self, parent, columns):
        frame = tk.Frame(parent)
        frame.pack(fill="both", expand=True)
        cols = [c[0] for c in columns]
        tree = ttk.Treeview(frame, columns=cols, show="headings", selectmode="browse")
        vsb = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        hsb = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        for col, width in columns:
            tree.heading(col, text=col)
            tree.column(col, width=width, minwidth=60, anchor="w")
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        tree.pack(fill="both", expand=True)
        return tree

    # ── logging ────────────────────────────────────────────────────────────────

    def _log(self, msg, tag="info"):
        colors = {"info": "#d4d4d4", "ok": "#4ec9b0", "warn": "#dcdcaa", "err": "#f44747"}
        self.log.configure(state="normal")
        self.log.tag_config(tag, foreground=colors.get(tag, "#d4d4d4"))
        self.log.insert("end", msg + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    # ── lookup ─────────────────────────────────────────────────────────────────

    def _lookup(self):
        uid = self.uid_var.get().strip()
        if not uid:
            messagebox.showwarning("Input required", "Please enter an ExternalUID.")
            return
        self._clear_tables()
        try:
            conn = self._get_conn()
            cur = conn.cursor()

            cur.execute(SQL_HEADER, uid)
            self._header_rows = cur.fetchall()
            for row in self._header_rows:
                self.h_tree.insert("", "end", values=list(row))

            cur.execute(SQL_DETAIL, uid)
            self._detail_rows = cur.fetchall()
            for row in self._detail_rows:
                self.d_tree.insert("", "end", values=list(row))

            if not self._header_rows:
                self._log(f"No records found for ExternalUID '{uid}'.", "warn")
            else:
                self._log(f"Loaded {len(self._header_rows)} header row(s) and "
                          f"{len(self._detail_rows)} detail row(s) for UID '{uid}'.", "ok")
                # Auto-select first header row so edit panel pre-fills
                first = self.h_tree.get_children()
                if first:
                    self.h_tree.selection_set(first[0])
                    self.nb.select(0)

        except pyodbc.Error as e:
            self._log(f"DB error: {e}", "err")
            messagebox.showerror("Database Error", str(e))

    # ── tree selection → edit panel ────────────────────────────────────────────

    def _on_header_select(self, _event):
        sel = self.h_tree.selection()
        if not sel:
            return
        vals = self.h_tree.item(sel[0], "values")
        # cols: IDSalesHeaderInterface, ExternalUID, InterfaceStatus,
        #       InterfaceErrorText, OrderID, StatusErrorText, ErrorDate
        self.rec_id_var.set(vals[0])
        self.status_var.set(vals[2])
        self.err_text_var.set(vals[3])
        self.rec_type_var.set("header")

    def _on_detail_select(self, _event):
        sel = self.d_tree.selection()
        if not sel:
            return
        vals = self.d_tree.item(sel[0], "values")
        # cols: IDSalesDetailInterface, InterfaceStatus, InterfaceErrorText,
        #       OrderID, StatusErrorText, ErrorDate
        self.rec_id_var.set(vals[0])
        self.status_var.set(vals[1])
        self.err_text_var.set(vals[2])
        self.rec_type_var.set("detail")

    # ── save correction (single record) ───────────────────────────────────────

    def _save_correction(self):
        rec_id = self.rec_id_var.get().strip()
        rec_type = self.rec_type_var.get()
        if not rec_id or not rec_type:
            messagebox.showwarning("No record selected", "Select a row from the Header or Detail tab first.")
            return

        new_status = self.status_var.get().strip()
        new_err = self.err_text_var.get().strip()

        if rec_type == "header":
            sql = """UPDATE DBO.SalesHeader_Interface
                     SET InterfaceStatus = ?, InterfaceErrorText = ?
                     WHERE IDSalesHeaderInterface = ?"""
            label = "SalesHeader_Interface"
        else:
            sql = """UPDATE DBO.SalesDetail_Interface
                     SET InterfaceStatus = ?, InterfaceErrorText = ?
                     WHERE IDSalesDetailInterface = ?"""
            label = "SalesDetail_Interface"

        if not messagebox.askyesno("Confirm Correction",
                                   f"Update {label} record {rec_id}?\n\n"
                                   f"Status → {new_status}\nError Text → {new_err or '(empty)'}"):
            return
        try:
            conn = self._get_conn()
            cur = conn.cursor()
            cur.execute(sql, new_status, new_err, rec_id)
            conn.commit()
            self._log(f"Corrected {label} ID {rec_id}: status='{new_status}', errorText='{new_err}'", "ok")
            self._lookup()   # refresh tables
        except pyodbc.Error as e:
            self._log(f"DB error: {e}", "err")
            messagebox.showerror("Database Error", str(e))

    # ── full reset ─────────────────────────────────────────────────────────────

    def _reset_all(self):
        uid = self.uid_var.get().strip()
        if not uid:
            messagebox.showwarning("No UID", "Enter and look up an ExternalUID first.")
            return
        if not self._header_rows:
            messagebox.showwarning("No data", "Look up the ExternalUID first.")
            return
        if not messagebox.askyesno("Confirm Full Reset",
                                   f"Reset ALL interface records for ExternalUID '{uid}' to NEW?\n\n"
                                   "This will:\n"
                                   "  • Delete all SalesDetail_InterfaceStatus rows\n"
                                   "  • Delete all SalesDetail_InterfaceError rows\n"
                                   "  • Set SalesDetail_Interface → InterfaceStatus = NEW\n"
                                   "  • Delete all SalesHeader_InterfaceStatus rows\n"
                                   "  • Delete all SalesHeader_InterfaceError rows\n"
                                   "  • Set SalesHeader_Interface → InterfaceStatus = NEW\n\n"
                                   "This cannot be undone."):
            return
        try:
            conn = self._get_conn()
            cur = conn.cursor()

            cur.execute("""DELETE FROM DBO.SalesDetail_InterfaceStatus
                           WHERE IDSalesDetailInterface IN (
                               SELECT IDSalesDetailInterface FROM DBO.SalesDetail_Interface
                               WHERE IDSalesHeaderInterface IN (
                                   SELECT IDSalesHeaderInterface FROM dbo.SalesHeader_Interface
                                   WHERE ExternalUID = ?))""", uid)
            self._log(f"  Deleted {cur.rowcount} SalesDetail_InterfaceStatus row(s).", "info")

            cur.execute("""DELETE FROM DBO.SalesDetail_InterfaceError
                           WHERE IDSalesDetailInterface IN (
                               SELECT IDSalesDetailInterface FROM DBO.SalesDetail_Interface
                               WHERE IDSalesHeaderInterface IN (
                                   SELECT IDSalesHeaderInterface FROM dbo.SalesHeader_Interface
                                   WHERE ExternalUID = ?))""", uid)
            self._log(f"  Deleted {cur.rowcount} SalesDetail_InterfaceError row(s).", "info")

            cur.execute("""UPDATE DBO.SalesDetail_Interface
                           SET InterfaceStatus = 'NEW', InterfaceErrorText = '', orderid = ''
                           WHERE IDSalesHeaderInterface IN (
                               SELECT IDSalesHeaderInterface FROM dbo.SalesHeader_Interface
                               WHERE ExternalUID = ?)""", uid)
            self._log(f"  Updated {cur.rowcount} SalesDetail_Interface row(s) → NEW.", "info")

            cur.execute("""DELETE FROM DBO.SalesHeader_InterfaceStatus
                           WHERE IDSalesHeaderInterface IN (
                               SELECT IDSalesHeaderInterface FROM dbo.SalesHeader_Interface
                               WHERE ExternalUID = ?)""", uid)
            self._log(f"  Deleted {cur.rowcount} SalesHeader_InterfaceStatus row(s).", "info")

            cur.execute("""DELETE FROM dbo.SalesHeader_InterfaceError
                           WHERE IDSalesHeaderInterface IN (
                               SELECT IDSalesHeaderInterface FROM dbo.SalesHeader_Interface
                               WHERE ExternalUID = ?)""", uid)
            self._log(f"  Deleted {cur.rowcount} SalesHeader_InterfaceError row(s).", "info")

            cur.execute("""UPDATE DBO.SalesHeader_Interface
                           SET InterfaceStatus = 'NEW', InterfaceErrorText = '', orderid = ''
                           WHERE ExternalUID = ?""", uid)
            self._log(f"  Updated {cur.rowcount} SalesHeader_Interface row(s) → NEW.", "info")

            conn.commit()
            self._log(f"COMMIT — UID '{uid}' fully reset to NEW.", "ok")
            self._lookup()   # refresh tables to confirm

        except pyodbc.Error as e:
            try:
                conn.rollback()
            except Exception:
                pass
            self._log(f"ERROR — rolled back: {e}", "err")
            messagebox.showerror("Database Error", str(e))

    # ── helpers ────────────────────────────────────────────────────────────────

    def _clear_tables(self):
        for item in self.h_tree.get_children():
            self.h_tree.delete(item)
        for item in self.d_tree.get_children():
            self.d_tree.delete(item)
        self._header_rows = []
        self._detail_rows = []
        self.rec_id_var.set("")
        self.status_var.set("")
        self.err_text_var.set("")
        self.rec_type_var.set("")

    def _clear(self):
        self.uid_var.set("")
        self._clear_tables()
        self._log("Cleared.", "info")

    def destroy(self):
        if self._conn:
            try:
                self._conn.close()
            except Exception:
                pass
        super().destroy()


# ── entry point ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
