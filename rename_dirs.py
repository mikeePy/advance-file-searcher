"""
rename_dirs_gui.py
==================
A desktop GUI for renaming directories using an Excel spreadsheet.

Excel format:
  Column A : Full path of the existing directory
  Column B : New directory name (name only, no path)

Run:
  python rename_dirs_gui.py

Dependencies:
  pip install pandas openpyxl          # (add xlrd if you need .xls files)
"""

import re
from datetime import datetime
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd

# ---------------------------------------------------------------------------
# Configuration & status constants
# ---------------------------------------------------------------------------

INVALID_RE = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
HEADER_HINTS = {
    "path", "original directory", "original path",
    "directory", "old path", "source",
}
LOG_DIR = Path("rename_logs")

STATUS_READY      = "✅ Ready"
STATUS_SKIP_EMPTY = "⏭️ Skipped: empty new name"
STATUS_SKIP_SAME  = "⏭️ Already named correctly"
STATUS_INVALID    = "❌ New name is empty after sanitising"
STATUS_NO_SOURCE  = "❌ Source path does not exist"
STATUS_NOT_DIR    = "❌ Source is not a directory"
STATUS_DUP        = "⚠️ Duplicate target in this batch"
STATUS_EXISTS     = "⚠️ Target already exists"

# ---------------------------------------------------------------------------
# Core logic (same as the Streamlit version)
# ---------------------------------------------------------------------------


def clean_cell(value) -> str:
    """Normalize a cell: NaN -> '', strip whitespace and surrounding quotes."""
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if len(s) >= 2 and s[0] == s[-1] and s[0] in ("'", '"'):
        s = s[1:-1].strip()
    return s


def sanitize_name(name: str) -> str:
    """Make a string safe as a folder name on Windows and Linux."""
    name = INVALID_RE.sub("_", name).strip()
    name = name.rstrip(". ")
    return name


def build_plan(df: pd.DataFrame, has_header: bool) -> list[dict]:
    plan: list[dict] = []
    proposed: set[str] = set()

    for idx, row in df.iterrows():
        path = clean_cell(row.iloc[0])
        new_name = clean_cell(row.iloc[1])

        # Skip header row if present
        if idx == 0 and (has_header or path.lower() in HEADER_HINTS):
            continue
        if not path:
            continue

        entry = {
            "row": int(idx) + 1,
            "original_path": path,
            "new_name": new_name,
            "new_name_clean": "",
            "new_path": "",
            "status": "",
        }

        if not new_name:
            entry["status"] = STATUS_SKIP_EMPTY
            plan.append(entry)
            continue

        clean = sanitize_name(new_name)
        entry["new_name_clean"] = clean
        if not clean:
            entry["status"] = STATUS_INVALID
            plan.append(entry)
            continue

        src = Path(path)
        if not src.exists():
            entry["status"] = STATUS_NO_SOURCE
            plan.append(entry)
            continue
        if not src.is_dir():
            entry["status"] = STATUS_NOT_DIR
            plan.append(entry)
            continue

        dst = src.parent / clean
        entry["new_path"] = str(dst)

        key = str(dst).lower()
        src_key = str(src).lower()

        if key == src_key:
            entry["status"] = STATUS_SKIP_SAME
        elif key in proposed:
            entry["status"] = STATUS_DUP
        elif dst.exists():
            entry["status"] = STATUS_EXISTS
        else:
            entry["status"] = STATUS_READY
            proposed.add(key)

        plan.append(entry)
    return plan


def execute_plan(plan: list[dict]) -> list[dict]:
    results: list[dict] = []
    for row in plan:
        if row["status"] != STATUS_READY:
            results.append({**row, "result": "skipped"})
            continue
        try:
            Path(row["original_path"]).rename(row["new_path"])
            results.append({**row, "result": "renamed"})
        except Exception as exc:
            results.append({**row, "result": f"error: {exc}"})
    return results


def save_log(results: list[dict]) -> Path:
    LOG_DIR.mkdir(exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    log_file = LOG_DIR / f"rename_log_{ts}.csv"
    pd.DataFrame(results).to_csv(log_file, index=False, encoding="utf-8-sig")
    return log_file


# ---------------------------------------------------------------------------
# GUI
# ---------------------------------------------------------------------------


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Directory Renamer")
        self.geometry("1200x720")
        self.minsize(900, 500)

        self.plan: list[dict] = []
        self.results: list[dict] = []
        self._current_df: pd.DataFrame | None = None

        self.has_header = tk.BooleanVar(value=False)
        self.confirm_var = tk.BooleanVar(value=False)

        self._build_ui()

    # -- layout ----------------------------------------------------------
    def _build_ui(self):
        top = ttk.Frame(self, padding=10)
        top.pack(fill=tk.X)

        ttk.Button(top, text="📂 Open Excel file…",
                   command=self.load_excel).pack(side=tk.LEFT)
        ttk.Button(top, text="⬇️ Save template",
                   command=self.save_template).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Checkbutton(top, text="First row is a header",
                        variable=self.has_header,
                        command=self._rebuild).pack(side=tk.LEFT, padx=(20, 0))

        self.file_label = ttk.Label(top, text="No file loaded",
                                    foreground="gray")
        self.file_label.pack(side=tk.LEFT, padx=(20, 0))

        mid = ttk.Frame(self, padding=(10, 0, 10, 0))
        mid.pack(fill=tk.BOTH, expand=True)

        cols = ("row", "orig", "new", "full", "status")
        self.tree = ttk.Treeview(mid, columns=cols, show="headings",
                                 selectmode="browse")
        self.tree.heading("row", text="#")
        self.tree.heading("orig", text="Original Path")
        self.tree.heading("new", text="New Name")
        self.tree.heading("full", text="New Full Path")
        self.tree.heading("status", text="Status")
        self.tree.column("row", width=50, anchor="e", stretch=False)
        self.tree.column("orig", width=380)
        self.tree.column("new", width=200)
        self.tree.column("full", width=380)
        self.tree.column("status", width=230)

        self.tree.tag_configure("ready", background="#e6f4ea")
        self.tree.tag_configure("skip", background="#f5f5f5")
        self.tree.tag_configure("error", background="#fdecea")
        self.tree.tag_configure("warn", background="#fff4e5")

        vsb = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(mid, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        mid.rowconfigure(0, weight=1)
        mid.columnconfigure(0, weight=1)

        bottom = ttk.Frame(self, padding=10)
        bottom.pack(fill=tk.X)

        self.summary = ttk.Label(bottom, text="Load an Excel file to begin.")
        self.summary.pack(anchor="w")

        actions = ttk.Frame(bottom)
        actions.pack(fill=tk.X, pady=(8, 0))

        ttk.Checkbutton(
            actions,
            text="I understand this will rename folders on disk",
            variable=self.confirm_var,
        ).pack(side=tk.LEFT)

        self.execute_btn = ttk.Button(
            actions, text="🚀 Rename directories",
            command=self.execute_rename, state=tk.DISABLED,
        )
        self.execute_btn.pack(side=tk.RIGHT)

        self.log_btn = ttk.Button(
            actions, text="⬇️ Save results CSV",
            command=self.save_results_csv, state=tk.DISABLED,
        )
        self.log_btn.pack(side=tk.RIGHT, padx=(0, 8))

    # -- actions ---------------------------------------------------------
    def load_excel(self):
        path = filedialog.askopenfilename(
            title="Select Excel file",
            filetypes=[("Excel files", "*.xlsx *.xls"),
                       ("All files", "*.*")],
        )
        if not path:
            return
        try:
            df = pd.read_excel(path, header=None)
        except Exception as exc:
            messagebox.showerror("Read error",
                                 f"Could not read file:\n\n{exc}")
            return
        if df.shape[1] < 2:
            messagebox.showerror(
                "Invalid format",
                "The file needs at least two columns: A=path, B=new name.",
            )
            return
        self._current_df = df.iloc[:, :2].reset_index(drop=True)
        self.file_label.config(text=Path(path).name, foreground="black")
        self._rebuild()

    def _rebuild(self):
        if self._current_df is None:
            return
        self.plan = build_plan(self._current_df, self.has_header.get())
        self._refresh_tree()
        self._refresh_summary()

    def _refresh_tree(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        for r in self.plan:
            status = r["status"]
            if status == STATUS_READY:
                tag = "ready"
            elif status.startswith("❌"):
                tag = "error"
            elif status.startswith("⚠️"):
                tag = "warn"
            else:
                tag = "skip"
            self.tree.insert(
                "", "end",
                values=(r["row"], r["original_path"], r["new_name"],
                        r["new_path"], status),
                tags=(tag,),
            )

    def _refresh_summary(self):
        ready     = sum(1 for r in self.plan if r["status"] == STATUS_READY)
        already   = sum(1 for r in self.plan if r["status"] == STATUS_SKIP_SAME)
        missing   = sum(1 for r in self.plan if r["status"].startswith("❌"))
        conflicts = sum(1 for r in self.plan if r["status"].startswith("⚠️"))
        total     = len(self.plan)
        self.summary.config(
            text=(f"Total: {total}   ·   ✅ Ready: {ready}"
                  f"   ·   ⏭️ Already named: {already}"
                  f"   ·   ⚠️ Conflicts: {conflicts}"
                  f"   ·   ❌ Missing/invalid: {missing}")
        )
        self.execute_btn.config(state=tk.NORMAL if ready else tk.DISABLED)
        self.confirm_var.set(False)

    def execute_rename(self):
        if not self.confirm_var.get():
            messagebox.showwarning("Confirm",
                                   "Please tick the confirmation checkbox first.")
            return
        ready = [r for r in self.plan if r["status"] == STATUS_READY]
        if not ready:
            messagebox.showinfo("Nothing to do",
                                "No rows are ready to rename.")
            return
        if not messagebox.askyesno(
            "Confirm rename", f"Rename {len(ready)} directories now?"
        ):
            return

        self.configure(cursor="watch")
        self.update_idletasks()
        try:
            results = execute_plan(self.plan)
        finally:
            self.configure(cursor="")

        self.results = results
        log_path = save_log(results)
        ok   = sum(1 for r in results if r["result"] == "renamed")
        fail = sum(1 for r in results if str(r["result"]).startswith("error"))

        messagebox.showinfo(
            "Done",
            f"Renamed: {ok}\nFailed: {fail}\n\nLog saved to:\n{log_path}",
        )

        # Update the workbook dataframe so the next preview reflects reality
        df = self._current_df.copy()
        for r in results:
            if r["result"] == "renamed":
                df.iloc[r["row"] - 1, 0] = r["new_path"]
        self._current_df = df

        self._rebuild()
        self.log_btn.config(state=tk.NORMAL)

    def save_results_csv(self):
        if not self.results:
            return
        path = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv")],
            initialfile=f"rename_results_{datetime.now():%Y%m%d_%H%M%S}.csv",
        )
        if not path:
            return
        pd.DataFrame(self.results).to_csv(
            path, index=False, encoding="utf-8-sig"
        )
        messagebox.showinfo("Saved", f"Saved to:\n{path}")

    def save_template(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")],
            initialfile="rename_template.xlsx",
        )
        if not path:
            return
        df = pd.DataFrame(
            [
                [r"D:\Movie\Hotel Iris 2021 720p BluRay AAC-[YTS MX]",
                 "Hotel Iris 2021"],
                [r"D:\Movie\Honey Lemon Soda (2021)",
                 "Honey Lemon Soda 2021"],
            ],
            columns=["Original Path", "New Name"],
        )
        df.to_excel(path, index=False)
        messagebox.showinfo("Saved", f"Template saved to:\n{path}")


def main():
    App().mainloop()


if __name__ == "__main__":
    main()