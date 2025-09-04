import os
import re
import json
import time
import html
import queue
import threading
import unicodedata
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, List, Dict

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter import font as tkfont

# Outlook COM
import pythoncom
import win32com.client as win32

APP_CFG = Path.home() / ".mailmerge_gui_config.json"

DEFAULT_SUBJECT = "{ten}_{msnv} luong"
DEFAULT_BODY = (
    "Kính gửi: {ten} - {msnv}\n"
    "Phòng Nhân sự (HR) gửi chi tiết Phiếu lương T07/2025 như file đính kèm.\n"
    "- Vui lòng kiểm tra và KHÔNG phản hồi qua email này.\n"
    "- Hãy nhập 4 số cuối CMND/CCCD để mở file đính kèm.\n"
    "Trân trọng."
)

PREFERRED_EXTS = [".pdf", ".PDF", ".docx", ".xlsx", ".doc", ".xls"]

# ---------------- Utils ----------------
def load_cfg():
    if APP_CFG.exists():
        try:
            return json.loads(APP_CFG.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}

def save_cfg(d: dict):
    try:
        APP_CFG.write_text(json.dumps(d, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass

def slugify(s: str) -> str:
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s

def norm_key(col: str) -> str:
    s = unicodedata.normalize("NFKD", col)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", "_", s.strip())
    s = re.sub(r"[^A-Za-z0-9_]", "_", s)
    while "__" in s:
        s = s.replace("__", "_")
    return s.strip("_").lower()

def css_font_stack(name: str) -> str:
    n = (name or "").strip().lower()
    if n == "times new roman":
        return "'Times New Roman', Times, serif"
    if n == "arial":
        return "Arial, 'Helvetica Neue', Helvetica, sans-serif"
    if n == "segoe ui":
        return "'Segoe UI', Tahoma, Arial, 'Helvetica Neue', Helvetica, sans-serif"
    if n == "tahoma":
        return "Tahoma, 'Segoe UI', Arial, sans-serif"
    if n == "calibri":
        return "Calibri, 'Segoe UI', Arial, sans-serif"
    if n == "verdana":
        return "Verdana, Arial, sans-serif"
    if n == "georgia":
        return "Georgia, 'Times New Roman', serif"
    if n == "courier new":
        return "'Courier New', Courier, monospace"
    # fallback
    return f"'{name}', Arial, sans-serif"

def smart_find_attachment(base_dir: Path, base_name: str) -> Optional[Path]:
    if not base_name:
        return None
    base_name = base_name.strip().strip('"').strip("'")
    p = Path(base_name)
    if p.is_file():
        return p
    p2 = base_dir / base_name
    if p2.is_file():
        return p2
    stem, ext = os.path.splitext(base_name)
    if ext:
        target = (stem + ext).lower()
        for cand in base_dir.rglob("*"):
            if cand.is_file() and cand.name.lower() == target:
                return cand
    else:
        for e in PREFERRED_EXTS:
            cand = base_dir / f"{base_name}{e}"
            if cand.is_file():
                return cand
    candidates: List[Path] = []
    if not ext and stem:
        for e in PREFERRED_EXTS:
            candidates += list(base_dir.rglob(f"{stem}{e}"))
    if not candidates:
        tgt = slugify(stem if stem else base_name)
        for cand in base_dir.rglob("*"):
            if cand.is_file() and slugify(cand.stem) == tgt:
                candidates.append(cand)
    if not candidates:
        return None
    def score(path: Path):
        try:
            return PREFERRED_EXTS.index(path.suffix)
        except ValueError:
            return len(PREFERRED_EXTS)
    candidates.sort(key=score)
    return candidates[0]

def send_with_outlook(to_addr: str, subject: str, body_text: Optional[str],
                      attachment: Optional[Path], sender_account: str,
                      html_body: Optional[str] = None):
    pythoncom.CoInitialize()
    outlook = win32.Dispatch("Outlook.Application")
    mail = outlook.CreateItem(0)
    try:
        for acc in outlook.Session.Accounts:
            if acc.SmtpAddress.lower() == sender_account.lower():
                mail._oleobj_.Invoke(*(64209, 0, 8, 0, acc))
                break
    except Exception:
        pass
    mail.To = to_addr
    mail.Subject = subject
    if html_body is not None:
        mail.BodyFormat = 2
        mail.HTMLBody = html_body
    else:
        mail.Body = body_text or ""
    if attachment and attachment.exists():
        mail.Attachments.Add(str(attachment))
    mail.Send()

# ---- Convert Text (tags) -> HTML with bullets & numbers ----
def textwidget_to_html(t: tk.Text, css_family: str, font_px: int) -> str:
    def char_html_at(index: str) -> str:
        ch = t.get(index)
        tags = set(t.tag_names(index))
        s = html.escape(ch)
        if "bold" in tags: s = f"<b>{s}</b>"
        if "italic" in tags: s = f"<i>{s}</i>"
        if "underline" in tags: s = f"<u>{s}</u>"
        return s

    text = t.get("1.0", "end-1c")
    lines = text.split("\n")
    html_lines: List[str] = []
    ul_open = False
    ol_open = False
    def close_lists():
        nonlocal ul_open, ol_open
        if ul_open: html_lines.append("</ul>"); ul_open = False
        if ol_open: html_lines.append("</ol>"); ol_open = False

    for i, _line in enumerate(lines, start=1):
        i0 = f"{i}.0"; i1 = f"{i}.{len(_line)}"
        lstr = _line.lstrip()
        is_bullet = lstr.startswith("• ") or lstr.startswith("- ")
        is_number = bool(re.match(r"^\d+[\.\)]\s", lstr))
        if is_bullet or is_number:
            # xác định vị trí sau marker
            if is_bullet:
                m = re.search(r"(• |- )", _line)
                start_col = (m.end() if m else 0)
                if not ul_open:
                    close_lists(); html_lines.append("<ul>"); ul_open = True
            else:
                m = re.search(r"(\d+[\.\)]\s)", _line)
                start_col = (m.end() if m else 0)
                if not ol_open:
                    close_lists(); html_lines.append("<ol>"); ol_open = True
            a = f"{i}.{start_col}"
            b = i1
            frag = []
            idx = a
            while True:
                cur = t.index(idx)
                if t.compare(cur, ">=", b): break
                frag.append(char_html_at(cur)); idx = t.index(f"{cur}+1c")
            html_lines.append(f"<li>{''.join(frag).strip()}</li>")
        else:
            close_lists()
            frag = []
            idx = i0
            while True:
                cur = t.index(idx)
                if t.compare(cur, ">=", i1): break
                frag.append(char_html_at(cur)); idx = t.index(f"{cur}+1c")
            html_lines.append(f"<p style='margin:0 0 8px 0'>{''.join(frag) or '&nbsp;'}</p>")

    close_lists()
    style = f"font-family:{css_family}; font-size:{font_px}px; line-height:1.5; color:#202124;"
    return "<!doctype html><html><head><meta charset='utf-8'></head><body>" \
           f"<div style='{style}'>" + "\n".join(html_lines) + "</div></body></html>"

# ---------------- Data model ----------------
@dataclass
class RowJob:
    idx: int
    row_map: Dict[str, str]
    ten: str
    msnv: str
    tenfile: str
    email: str

# ---------------- GUI ----------------
class MailMergeApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Auto send mail - Design by SMV-IT")
        self.geometry("1180x830")

        self.cfg = load_cfg()

        self.excel_path = tk.StringVar(value=self.cfg.get("excel_path", ""))
        self.data_dir = tk.StringVar(value=self.cfg.get("data_dir", "D:\\Data"))
        self.sender = tk.StringVar(value=self.cfg.get("sender", ""))
        self.subject_tpl = tk.StringVar(value=self.cfg.get("subject_tpl", DEFAULT_SUBJECT))
        self.delay = tk.DoubleVar(value=float(self.cfg.get("delay", 1.0)))
        self.dry_run = tk.BooleanVar(value=bool(self.cfg.get("dry_run", False)))
        self.require_attach = tk.BooleanVar(value=bool(self.cfg.get("require_attach", True)))
        self.filter_complete = tk.BooleanVar(value=bool(self.cfg.get("filter_complete", True)))
        self.send_selected_only = tk.BooleanVar(value=True)

        # font setup
        self.font_family = tk.StringVar(value=self.cfg.get("font_family", "Times New Roman"))
        self.font_size = tk.IntVar(value=int(self.cfg.get("font_size", 13)))

        self.df_raw: Optional[pd.DataFrame] = None
        self.columns: List[str] = []
        self.col_name = tk.StringVar()
        self.col_msnv = tk.StringVar()
        self.col_email = tk.StringVar()
        self.col_file = tk.StringVar()

        self.rows: List[RowJob] = []
        self.selected: set[int] = set()
        self.worker_thread = None
        self.stop_flag = False
        self.log_q = queue.Queue()

        self._build_ui()
        self.after(200, self._drain_logs)

    # ---------- UI ----------
    def _build_ui(self):
        main = ttk.Frame(self, padding=10)
        main.pack(fill="both", expand=True)

        r = 0
        ttk.Label(main, text="File Excel (xlsx):").grid(row=r, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.excel_path, width=60).grid(row=r, column=1, sticky="we", padx=6)
        ttk.Button(main, text="Chọn file…", command=self._pick_excel).grid(row=r, column=2, padx=4)
        ttk.Button(main, text="Nạp cột", command=self.read_columns).grid(row=r, column=3, padx=4)
        r += 1

        ttk.Label(main, text="Thư mục Data (tìm file đính kèm):").grid(row=r, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.data_dir, width=60).grid(row=r, column=1, sticky="we", padx=6)
        ttk.Button(main, text="Chọn thư mục…", command=self._pick_dir).grid(row=r, column=2, padx=4)
        r += 1

        ttk.Label(main, text="Tài khoản Outlook gửi từ (SMTP):").grid(row=r, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.sender, width=40).grid(row=r, column=1, sticky="w", padx=6)
        r += 1

        ttk.Label(main, text="Subject template:").grid(row=r, column=0, sticky="w")
        ttk.Entry(main, textvariable=self.subject_tpl, width=60).grid(row=r, column=1, sticky="we", padx=6)
        r += 1

        # ===== Rich text editor =====
        ttk.Label(main, text="Body (soạn thảo định dạng):").grid(row=r, column=0, sticky="w")
        r += 1

        tb = ttk.Frame(main)
        tb.grid(row=r, column=0, columnspan=4, sticky="we")

        # Font list: ưu tiên Times New Roman, Arial trước
        all_fonts = sorted(set(tkfont.families()))
        pref = []
        for want in ["Times New Roman", "Arial", "Segoe UI", "Tahoma", "Calibri", "Verdana", "Georgia", "Courier New"]:
            if want in all_fonts:
                pref.append(want)
                all_fonts.remove(want)
        font_values = pref + all_fonts

        ttk.Label(tb, text="Font:").pack(side="left")
        self.cb_font = ttk.Combobox(tb, values=font_values, textvariable=self.font_family, width=34)
        self.cb_font.pack(side="left", padx=(4,8))
        self.cb_font.bind("<<ComboboxSelected>>", lambda e: self._apply_preview_font())

        ttk.Label(tb, text="Size:").pack(side="left")
        self.spin_size = ttk.Spinbox(tb, from_=10, to=28, textvariable=self.font_size, width=4,
                                     command=self._apply_preview_font)
        self.spin_size.pack(side="left", padx=(4,12))

        ttk.Button(tb, text="B", width=3, command=lambda: self._toggle_tag("bold")).pack(side="left")
        ttk.Button(tb, text="I", width=3, command=lambda: self._toggle_tag("italic")).pack(side="left", padx=2)
        ttk.Button(tb, text="U", width=3, command=lambda: self._toggle_tag("underline")).pack(side="left")

        ttk.Button(tb, text="• Bullet", command=lambda: self._toggle_list(kind="ul")).pack(side="left", padx=8)
        ttk.Button(tb, text="1. Numbered", command=lambda: self._toggle_list(kind="ol")).pack(side="left")

        r += 1
        self.body = tk.Text(main, height=8, wrap="word")
        self.body.grid(row=r, column=0, columnspan=4, sticky="we")
        self._setup_text_tags()
        self.body.insert("1.0", load_cfg().get("body_tpl", DEFAULT_BODY))
        self._apply_preview_font()
        r += 1

        # Mapping
        mapf = ttk.LabelFrame(main, text="Ánh xạ cột từ Excel (chọn đúng các cột cần dùng)")
        mapf.grid(row=r, column=0, columnspan=4, sticky="we", pady=(8, 4))
        ttk.Label(mapf, text="Cột TÊN:").grid(row=0, column=0, sticky="w")
        self.cb_ten = ttk.Combobox(mapf, textvariable=self.col_name, width=28, state="readonly")
        self.cb_ten.grid(row=0, column=1, sticky="w", padx=6)
        ttk.Label(mapf, text="Cột MSNV:").grid(row=0, column=2, sticky="w")
        self.cb_msnv = ttk.Combobox(mapf, textvariable=self.col_msnv, width=28, state="readonly")
        self.cb_msnv.grid(row=0, column=3, sticky="w", padx=6)
        ttk.Label(mapf, text="Cột EMAIL:").grid(row=1, column=0, sticky="w")
        self.cb_email = ttk.Combobox(mapf, textvariable=self.col_email, width=28, state="readonly")
        self.cb_email.grid(row=1, column=1, sticky="w", padx=6)
        ttk.Label(mapf, text="Cột TÊN FILE:").grid(row=1, column=2, sticky="w")
        self.cb_file = ttk.Combobox(mapf, textvariable=self.col_file, width=28, state="readonly")
        self.cb_file.grid(row=1, column=3, sticky="w", padx=6)
        ttk.Checkbutton(mapf, text="Chỉ nạp các hàng đủ 4 cột trên",
                        variable=self.filter_complete).grid(row=2, column=0, columnspan=4, sticky="w", pady=(6,2))
        ttk.Button(mapf, text="📂 Nạp danh sách theo ánh xạ", command=self.apply_mapping).grid(row=3, column=0, pady=(6,2))
        mapf.columnconfigure(1, weight=1); mapf.columnconfigure(3, weight=1)
        r += 1

        opts = ttk.Frame(main)
        opts.grid(row=r, column=0, columnspan=4, sticky="we", pady=(4, 4))
        ttk.Checkbutton(opts, text="Dry run (không gửi thật)", variable=self.dry_run).pack(side="left", padx=6)
        ttk.Checkbutton(opts, text="Bắt buộc có file đính kèm", variable=self.require_attach).pack(side="left", padx=6)
        ttk.Checkbutton(opts, text="Chỉ gửi các dòng đã chọn", variable=self.send_selected_only).pack(side="left", padx=6)
        ttk.Label(opts, text="Delay (giây):").pack(side="left", padx=(12,4))
        ttk.Entry(opts, textvariable=self.delay, width=6).pack(side="left")

        bar = ttk.Frame(main)
        bar.grid(row=r+1, column=0, columnspan=4, sticky="we", pady=(2,8))
        ttk.Button(bar, text="Chọn tất cả", command=self.select_all).pack(side="left")
        ttk.Button(bar, text="Bỏ chọn tất cả", command=self.clear_selection).pack(side="left", padx=4)
        ttk.Button(bar, text="▶ Gửi", command=self.start_send).pack(side="left", padx=12)
        ttk.Button(bar, text="⏹ Dừng", command=self.stop_send).pack(side="left")
        ttk.Button(bar, text="💾 Lưu cấu hình", command=self.save_settings).pack(side="right")

        self.cols = ("sel","idx","ten","msnv","email","tenfile","attach","status")
        self.col_index = {c:i for i,c in enumerate(self.cols)}
        self.tree = ttk.Treeview(main, columns=self.cols, show="headings", height=16)
        heads = {"sel":"✓","idx":"#","ten":"tên","msnv":"msnv","email":"mail","tenfile":"tên file","attach":"đính kèm","status":"trạng thái"}
        for c in self.cols:
            self.tree.heading(c, text=heads[c])
            if c == "sel":
                self.tree.column(c, width=40, anchor="center")
            elif c in ("idx","status"):
                self.tree.column(c, width=90, anchor="w")
            else:
                self.tree.column(c, width=160, anchor="w")
        self.tree.grid(row=r+2, column=0, columnspan=4, sticky="nsew")
        self.tree.bind("<Button-1>", self.on_tree_click)

        ttk.Label(main, text="Log:").grid(row=r+3, column=0, sticky="w")
        self.log = tk.Text(main, height=8)
        self.log.grid(row=r+3, column=1, columnspan=3, sticky="nsew")

        main.columnconfigure(1, weight=1); main.columnconfigure(3, weight=1)
        main.rowconfigure(r+2, weight=1)

    # ----- rich text helpers -----
    def _setup_text_tags(self):
        base = tkfont.Font(family="Times New Roman", size=13)
        self.body.configure(font=base)
        self.font_bold = tkfont.Font(self.body, self.body.cget("font")); self.font_bold.configure(weight="bold")
        self.font_italic = tkfont.Font(self.body, self.body.cget("font")); self.font_italic.configure(slant="italic")
        self.font_underline = tkfont.Font(self.body, self.body.cget("font")); self.font_underline.configure(underline=1)
        self.body.tag_configure("bold", font=self.font_bold)
        self.body.tag_configure("italic", font=self.font_italic)
        self.body.tag_configure("underline", font=self.font_underline)

    def _apply_preview_font(self, *args):
        fam = self.font_family.get() or "Times New Roman"
        size = int(self.font_size.get() or 13)
        self.body.configure(font=(fam, size))
        # cập nhật tag fonts theo base mới
        self.font_bold.configure(family=fam, size=size, weight="bold")
        self.font_italic.configure(family=fam, size=size, slant="italic")
        self.font_underline.configure(family=fam, size=size, underline=1)

    def _toggle_tag(self, tag: str):
        try:
            a, b = self.body.index("sel.first"), self.body.index("sel.last")
        except tk.TclError:
            return
        if tag in self.body.tag_names("sel.first"):
            self.body.tag_remove(tag, a, b)
        else:
            self.body.tag_add(tag, a, b)

    def _selected_line_range(self):
        try:
            a = self.body.index("sel.first linestart")
            b = self.body.index("sel.last lineend")
        except tk.TclError:
            a = self.body.index("insert linestart")
            b = self.body.index("insert lineend")
        return a, b

    def _toggle_list(self, kind: str = "ul"):
        a, b = self._selected_line_range()
        # lấy tất cả lines giữa a..b
        cur = a
        lines = []
        while self.body.compare(cur, "<=", b):
            lines.append(cur)
            cur = self.body.index(f"{cur}+1line")
            if self.body.compare(cur, ">", b):
                break
        # kiểm tra trạng thái hiện tại
        def is_bullet_line(text: str) -> bool:
            t = text.lstrip()
            return t.startswith("• ") or t.startswith("- ")
        def is_number_line(text: str) -> bool:
            return bool(re.match(r"^\s*\d+[\.\)]\s", text))

        all_bulleted = True
        all_numbered = True
        for ln in lines:
            s = self.body.get(ln, f"{ln} lineend")
            if not is_bullet_line(s): all_bulleted = False
            if not is_number_line(s): all_numbered = False

        if kind == "ul":
            if all_bulleted:
                # remove bullet
                for ln in lines:
                    s = self.body.get(ln, f"{ln} lineend")
                    if is_bullet_line(s):
                        pos = self.body.search(r"(• |- )", ln, regexp=True, stopindex=f"{ln} lineend")
                        if pos:
                            self.body.delete(pos, f"{pos}+2c")
            else:
                for ln in lines:
                    s = self.body.get(ln, f"{ln} lineend")
                    if not is_bullet_line(s) and not is_number_line(s):
                        self.body.insert(ln, "• ")
        else:  # ordered list
            if all_numbered:
                for ln in lines:
                    s = self.body.get(ln, f"{ln} lineend")
                    m = re.search(r"\s*\d+[\.\)]\s", s)
                    if m:
                        start = self.body.index(f"{ln}+{m.start()}c")
                        end = self.body.index(f"{ln}+{m.end()}c")
                        self.body.delete(start, end)
            else:
                # thêm số 1..N
                num = 1
                for ln in lines:
                    s = self.body.get(ln, f"{ln} lineend")
                    if not is_number_line(s) and not is_bullet_line(s):
                        self.body.insert(ln, f"{num}. ")
                        num += 1

    # ----- pickers -----
    def _pick_excel(self):
        p = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx *.xls")])
        if p: self.excel_path.set(p)

    def _pick_dir(self):
        d = filedialog.askdirectory()
        if d: self.data_dir.set(d)

    # ----- logging -----
    def _log(self, msg: str):
        self.log_q.put(msg)

    def _drain_logs(self):
        try:
            while True:
                msg = self.log_q.get_nowait()
                self.log.insert("end", msg + "\n")
                self.log.see("end")
        except queue.Empty:
            pass
        self.after(200, self._drain_logs)

    # ----- settings -----
    def save_settings(self):
        d = {
            "excel_path": self.excel_path.get().strip(),
            "data_dir": self.data_dir.get().strip(),
            "sender": self.sender.get().strip(),
            "subject_tpl": self.subject_tpl.get().strip(),
            "body_tpl": self.body.get("1.0","end").strip(),
            "delay": float(self.delay.get() or 1.0),
            "dry_run": bool(self.dry_run.get()),
            "require_attach": bool(self.require_attach.get()),
            "filter_complete": bool(self.filter_complete.get()),
            "send_selected_only": bool(self.send_selected_only.get()),
            "font_family": self.font_family.get().strip(),
            "font_size": int(self.font_size.get() or 13),
        }
        save_cfg(d)
        messagebox.showinfo("OK", f"Đã lưu cấu hình: {APP_CFG}")

    # ----- read columns -----
    def read_columns(self):
        path = self.excel_path.get().strip()
        if not path or not Path(path).exists():
            messagebox.showerror("Lỗi", "Hãy chọn file Excel hợp lệ."); return
        try:
            self.df_raw = pd.read_excel(path, dtype=str).fillna("")
            self.df_raw.columns = self.df_raw.columns.map(lambda c: str(c).strip())
        except Exception as e:
            messagebox.showerror("Lỗi đọc Excel", str(e)); return

        self.columns = list(self.df_raw.columns)
        values = self.columns.copy()
        self.cb_ten["values"] = values
        self.cb_msnv["values"] = values
        self.cb_email["values"] = values
        self.cb_file["values"] = values

        def find_like(substrs):
            for c in self.columns:
                lc = c.lower()
                if any(s in lc for s in substrs):
                    return c
            return ""

        self.col_name.set(find_like(["tên","ten","name"]) or (values[0] if values else ""))
        self.col_msnv.set(find_like(["msnv","mã","ma nv","mnv"]) or (values[0] if values else ""))
        self.col_email.set(find_like(["mail","email","e-mail"]) or (values[0] if values else ""))
        self.col_file.set(find_like(["file","tệp","ten file","tên file"]) or (values[0] if values else ""))

        self._log(f"Đã nạp {len(self.columns)} cột. Chọn ánh xạ rồi bấm '📂 Nạp danh sách theo ánh xạ'.")

    # ----- apply mapping + preview attach -----
    def apply_mapping(self):
        if self.df_raw is None:
            self.read_columns()
            if self.df_raw is None:
                return
        c_ten, c_msnv, c_email, c_file = (self.col_name.get().strip(),
                                          self.col_msnv.get().strip(),
                                          self.col_email.get().strip(),
                                          self.col_file.get().strip())
        for c in (c_ten, c_msnv, c_email, c_file):
            if c not in self.df_raw.columns:
                messagebox.showerror("Lỗi ánh xạ", f"Cột '{c}' không tồn tại trong Excel."); return

        must_have_all = bool(self.filter_complete.get())
        data_dir = Path(self.data_dir.get().strip() or ".")

        self.rows, self.selected = [], set()
        for i, row in self.df_raw.iterrows():
            m: Dict[str, str] = {}
            for col in self.df_raw.columns:
                cell = row.get(col, "")
                val = "" if pd.isna(cell) else str(cell).strip()
                m[col] = val; m[norm_key(col)] = val
            ten, email, msnv, tenfile = m.get(c_ten,""), m.get(c_email,""), m.get(c_msnv,""), m.get(c_file,"")
            if must_have_all and (not ten or not email or not msnv or not tenfile):
                continue
            self.rows.append(RowJob(idx=i+1, row_map=m, ten=ten, msnv=msnv, tenfile=tenfile, email=email))

        for x in self.tree.get_children():
            self.tree.delete(x)
        for r in self.rows:
            att = smart_find_attachment(data_dir, r.tenfile)
            att_name = att.name if att else "KHÔNG THẤY"
            self.tree.insert("", "end", iid=str(r.idx),
                             values=("☐", r.idx, r.ten, r.msnv, r.email, r.tenfile, att_name, "Chưa gửi"))
        self._log(f"Nạp {len(self.rows)} hàng (lọc thiếu dữ liệu={must_have_all}).")

    # ----- table checkbox -----
    def on_tree_click(self, event):
        if self.tree.identify("region", event.x, event.y) != "cell":
            return
        if self.tree.identify_column(event.x) != "#1":
            return
        row_id = self.tree.identify_row(event.y)
        if not row_id: return
        vals = list(self.tree.item(row_id, "values"))
        if vals[0] == "☑":
            vals[0] = "☐"; self.selected.discard(int(row_id))
        else:
            vals[0] = "☑"; self.selected.add(int(row_id))
        self.tree.item(row_id, values=vals)

    def select_all(self):
        for iid in self.tree.get_children():
            vals = list(self.tree.item(iid, "values"))
            if vals[0] != "☑":
                vals[0] = "☑"; self.tree.item(iid, values=vals)
            self.selected.add(int(iid))

    def clear_selection(self):
        for iid in self.tree.get_children():
            vals = list(self.tree.item(iid, "values"))
            if vals[0] != "☐":
                vals[0] = "☐"; self.tree.item(iid, values=vals)
        self.selected.clear()

    # ----- send -----
    def start_send(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Đang chạy", "Đang gửi, vui lòng đợi hoặc bấm Dừng."); return
        if not self.rows:
            messagebox.showwarning("Chưa có danh sách", "Hãy nạp danh sách theo ánh xạ trước."); return
        sender = self.sender.get().strip()
        if not sender:
            messagebox.showwarning("Thiếu người gửi", "Nhập tài khoản Outlook gửi đi."); return
        data_dir = Path(self.data_dir.get().strip() or ".")
        if not data_dir.exists():
            messagebox.showerror("Lỗi", "Thư mục Data không tồn tại."); return

        subject_tpl = self.subject_tpl.get().strip() or DEFAULT_SUBJECT
        delay = float(self.delay.get() or 1.0)
        dry = bool(self.dry_run.get())
        must_have_attach = bool(self.require_attach.get())
        only_selected = bool(self.send_selected_only.get())

        if only_selected:
            rows_to_send = [r for r in self.rows if r.idx in self.selected]
            if not rows_to_send:
                messagebox.showwarning("Chưa chọn dòng", "Bật 'Chỉ gửi các dòng đã chọn' nhưng chưa chọn dòng nào."); 
                return
        else:
            rows_to_send = self.rows

        css_family = css_font_stack(self.font_family.get() or "Times New Roman")
        font_px = int(self.font_size.get() or 13)

        self.stop_flag = False

        def worker():
            sent = 0; skipped = 0
            class SafeDict(dict):
                def __missing__(self, k): return "{"+k+"}"

            for r in rows_to_send:
                if self.stop_flag:
                    self._log("Đã dừng theo yêu cầu."); break

                ctx = dict(r.row_map); ctx.update({"ten":r.ten,"msnv":r.msnv,"tenfile":r.tenfile,"email":r.email})
                try:
                    subject = subject_tpl.format_map(SafeDict(**ctx))
                except Exception as e:
                    self._update_tree(r.idx, status=f"Lỗi template subject: {e}"); skipped += 1; continue

                # render HTML từ text widget theo font/size
                html_body = textwidget_to_html(self.body, css_family, font_px)
                try:
                    html_body = html_body.format_map(SafeDict(**ctx))
                except Exception as e:
                    self._update_tree(r.idx, status=f"Lỗi template body: {e}"); skipped += 1; continue

                attach = smart_find_attachment(data_dir, r.tenfile)
                att_name = attach.name if attach else "KHÔNG THẤY"
                self._update_tree(r.idx, attach=att_name)

                if must_have_attach and attach is None:
                    self._update_tree(r.idx, status="Bỏ qua: không có file"); skipped += 1; continue

                if dry:
                    self._log(f"[DRY] {r.email} | subj='{subject}' | attach='{att_name}'")
                    self._update_tree(r.idx, status="(DRY) OK"); sent += 1; continue

                try:
                    send_with_outlook(r.email, subject, None, attach, sender, html_body=html_body)
                    self._log(f"Đã gửi: {r.email} | {att_name}")
                    self._update_tree(r.idx, status="ĐÃ GỬI"); sent += 1
                    time.sleep(delay)
                except Exception as e:
                    self._log(f"[ERR] {r.email}: {e}")
                    self._update_tree(r.idx, status=f"Lỗi: {e}"); skipped += 1

            self._log(f"[SUMMARY] Sent={sent}, Skipped={skipped}, Total={len(rows_to_send)}")

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def stop_send(self):
        self.stop_flag = True

    def _update_tree(self, idx: int, attach: Optional[str] = None, status: Optional[str] = None):
        iid = str(idx)
        try:
            vals = list(self.tree.item(iid, "values"))
            if attach is not None:
                vals[self.col_index["attach"]] = attach
            if status is not None:
                vals[self.col_index["status"]] = status
            self.tree.item(iid, values=vals)
            self.update_idletasks()
        except Exception:
            pass

# ---------- main ----------
if __name__ == "__main__":
    app = MailMergeApp()
    app.mainloop()
