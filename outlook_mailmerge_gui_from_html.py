# outlook_mailmerge_gui_html_only.py

import os, re, json, time, queue, threading, unicodedata
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, List, Dict
from string import Template

import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# Outlook COM
import pythoncom
import win32com.client as win32
try:
    import winreg as wr  # lấy chữ ký Outlook
except Exception:
    wr = None

APP_CFG = Path.home() / ".mailmerge_gui_config.json"

DEFAULT_SUBJECT = "{ten}_{msnv} luong"
DEFAULT_HTML = """<p><strong>Kính gửi:</strong></p>
<p><strong>Anh/chị: {ten} - {msnv}</strong></p>
<ul>
  <li>Phòng Nhân sự (HR) gửi chi tiết Phiếu lương T07/2025 như file đính kèm.</li>
  <li>Vui lòng kiểm tra và <strong>KHÔNG</strong> phản hồi qua email này.</li>
  <li>Hãy nhập 4 số cuối CMND/CCCD để mở file đính kèm.</li>
</ul>
<p>Trân trọng.</p>
"""

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
    import unicodedata, re
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = s.lower()
    s = re.sub(r"[^a-z0-9]+", "-", s).strip("-")
    return s

def norm_key(col: str) -> str:
    import unicodedata, re
    s = unicodedata.normalize("NFKD", col)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    s = re.sub(r"\s+", "_", s.strip())
    s = re.sub(r"[^A-Za-z0-9_]", "_", s)
    while "__" in s: s = s.replace("__", "_")
    return s.strip("_").lower()

def smart_find_attachment(base_dir: Path, base_name: str) -> Optional[Path]:
    if not base_name:
        return None
    base_name = str(base_name).strip().strip('"').strip("'")
    stem, ext = os.path.splitext(base_name)

    # Tạo các biến thể: giữ nguyên + đảo vị trí nếu phát hiện mã NV (chữ+ số)
    def make_variants(s: str) -> List[str]:
        import re
        parts = re.split(r'[\s_\-]+', s.strip())
        code = next((p for p in parts if re.match(r'[A-Za-z]{2,}\d{3,}', p)), None)
        if code:
            name = " ".join([p for p in parts if p != code]).strip()
            var = [f"{name}_{code}", f"{code}_{name}"]
        else:
            var = [s]
        # thêm bản không dấu/chuẩn hoá cho so khớp
        return var

    variants = make_variants(stem if stem else base_name)

    # 1) Thử khớp tuyệt đối nếu người dùng đưa đủ đường dẫn
    p = Path(base_name)
    if p.is_file():
        return p
    p2 = base_dir / base_name
    if p2.is_file():
        return p2

    # 2) Nếu có đuôi, tìm theo exact name trong toàn thư mục
    if ext:
        target = (stem + ext).lower()
        for cand in base_dir.rglob("*"):
            if cand.is_file() and cand.name.lower() == target:
                return cand

    # 3) Thử từng biến thể với danh sách đuôi ưu tiên
    cands = []
    for v in variants:
        if ext:
            cands += list(base_dir.rglob(f"{v}{ext}"))
        else:
            for e in PREFERRED_EXTS:
                cands += list(base_dir.rglob(f"{v}{e}"))

    # 4) Nếu vẫn chưa thấy: so khớp “slug” (bỏ dấu, bỏ ký tự đặc biệt) giữa các biến thể và file trong thư mục
    if not cands:
        want_slugs = {slugify(v) for v in variants}
        for cand in base_dir.rglob("*"):
            if cand.is_file() and slugify(cand.stem) in want_slugs:
                cands.append(cand)

    if not cands:
        return None

    def score(path: Path):
        try:
            return PREFERRED_EXTS.index(path.suffix)
        except ValueError:
            return len(PREFERRED_EXTS)
    cands.sort(key=score)
    return cands[0]


def extract_body_fragment(full_html: str) -> str:
    import re
    m = re.search(r"<body[^>]*>(.*)</body>", full_html, flags=re.I | re.S)
    return m.group(1) if m else full_html

def get_default_signature_html() -> Optional[str]:
    if wr is None: return None
    for ver in ("16.0", "15.0", "14.0", "12.0"):
        try:
            k = wr.OpenKey(wr.HKEY_CURRENT_USER, rf"Software\Microsoft\Office\{ver}\Common\MailSettings")
            sig_name, _ = wr.QueryValueEx(k, "NewSignature"); wr.CloseKey(k)
            if not sig_name: continue
            sig_dir = Path(os.environ.get("APPDATA","")) / "Microsoft" / "Signatures"
            html_path = sig_dir / f"{sig_name}.htm"
            if html_path.exists():
                try: return html_path.read_text(encoding="utf-8", errors="ignore")
                except Exception: return html_path.read_text(encoding="cp1252", errors="ignore")
        except Exception:
            continue
    return None

def apply_variables(html_text: str, ctx: Dict[str, str]) -> str:
    # Hỗ trợ $ten và {ten}
    html_text = Template(html_text).safe_substitute(**ctx)
    import re
    def repl(m): 
        k = m.group(1)
        return str(ctx.get(k, m.group(0)))
    return re.sub(r"\{([A-Za-z_][A-Za-z0-9_]*)\}", repl, html_text)


# ---------------- Model ----------------
@dataclass
class RowJob:
    idx: int
    row_map: Dict[str, str]
    ten: str
    msnv: str
    tenfile: str
    email: str


# ---------------- GUI ----------------
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Outlook Mail Merge (HTML only) — SMV-IT")
        self.geometry("1180x820")

        self.cfg = load_cfg()

        self.excel_path = tk.StringVar(value=self.cfg.get("excel_path", ""))
        self.data_dir   = tk.StringVar(value=self.cfg.get("data_dir", "D:\\Data"))
        self.sender     = tk.StringVar(value=self.cfg.get("sender", ""))
        self.subject_tpl= tk.StringVar(value=self.cfg.get("subject_tpl", DEFAULT_SUBJECT))
        self.delay      = tk.DoubleVar(value=float(self.cfg.get("delay", 1.0)))
        self.dry_run    = tk.BooleanVar(value=bool(self.cfg.get("dry_run", False)))
        self.require_attach = tk.BooleanVar(value=bool(self.cfg.get("require_attach", True)))
        self.send_selected_only = tk.BooleanVar(value=True)
        self.append_signature   = tk.BooleanVar(value=bool(self.cfg.get("append_signature", True)))

        # HTML template
        self.html_src_txt = self.cfg.get("html_src", DEFAULT_HTML)

        # global attachments
        self.extra_files: List[Path] = [Path(p) for p in self.cfg.get("extra_files", []) if p]

        # mapping
        self.df_raw: Optional[pd.DataFrame] = None
        self.columns: List[str] = []
        self.col_name = tk.StringVar()
        self.col_msnv = tk.StringVar()
        self.col_email= tk.StringVar()
        self.col_file = tk.StringVar()
        self.filter_complete = tk.BooleanVar(value=bool(self.cfg.get("filter_complete", True)))

        # runtime
        self.rows: List[RowJob] = []; self.selected:set[int] = set()
        self.worker_thread=None; self.stop_flag=False; self.log_q=queue.Queue()

        self._build_ui()
        self.after(200, self._drain_logs)

    def _build_ui(self):
        main = ttk.Frame(self, padding=10); main.pack(fill="both", expand=True)

        r=0
        ttk.Label(main,text="File Excel (xlsx):").grid(row=r,column=0,sticky="w")
        ttk.Entry(main,textvariable=self.excel_path,width=60).grid(row=r,column=1,sticky="we",padx=6)
        ttk.Button(main,text="Chọn file…",command=self._pick_excel).grid(row=r,column=2,padx=4)
        ttk.Button(main,text="Nạp cột",command=self.read_columns).grid(row=r,column=3,padx=4); r+=1

        ttk.Label(main,text="Thư mục Data (tìm file đính kèm):").grid(row=r,column=0,sticky="w")
        ttk.Entry(main,textvariable=self.data_dir,width=60).grid(row=r,column=1,sticky="we",padx=6)
        ttk.Button(main,text="Chọn thư mục…",command=self._pick_dir).grid(row=r,column=2,padx=4); r+=1

        ttk.Label(main,text="Tài khoản Outlook gửi từ (SMTP):").grid(row=r,column=0,sticky="w")
        ttk.Entry(main,textvariable=self.sender,width=40).grid(row=r,column=1,sticky="w",padx=6); r+=1

        ttk.Label(main,text="Subject template:").grid(row=r,column=0,sticky="w")
        ttk.Entry(main,textvariable=self.subject_tpl,width=60).grid(row=r,column=1,sticky="we",padx=6); r+=1

        # HTML template (thu gọn, chiếm ít chỗ; có thể bung ra chỉnh khi cần)
        htmlf = ttk.LabelFrame(main, text="HTML template (sử dụng trực tiếp) – hỗ trợ {ten}/{msnv} hoặc $ten/$msnv")
        htmlf.grid(row=r,column=0,columnspan=4,sticky="we",pady=(6,6))
        self.html_src = tk.Text(htmlf, height=6, wrap="word", font=("Consolas", 11))
        self.html_src.grid(row=0, column=0, columnspan=3, sticky="we", padx=6, pady=6)
        self.html_src.insert("1.0", self.html_src_txt)
        ttk.Button(htmlf, text="Xem trước (dòng đang chọn)", command=self.preview_current_html).grid(row=0,column=3,sticky="n", padx=6)
        htmlf.columnconfigure(0, weight=1); htmlf.columnconfigure(1, weight=1); htmlf.columnconfigure(2, weight=1)
        r+=1

        # Mapping
        mapf = ttk.LabelFrame(main,text="Ánh xạ cột từ Excel")
        mapf.grid(row=r,column=0,columnspan=4,sticky="we",pady=(6,4))
        ttk.Label(mapf,text="Cột TÊN:").grid(row=0,column=0,sticky="w")
        self.cb_ten=ttk.Combobox(mapf,textvariable=self.col_name,width=28,state="readonly"); self.cb_ten.grid(row=0,column=1,sticky="w",padx=6)
        ttk.Label(mapf,text="Cột MSNV:").grid(row=0,column=2,sticky="w")
        self.cb_msnv=ttk.Combobox(mapf,textvariable=self.col_msnv,width=28,state="readonly"); self.cb_msnv.grid(row=0,column=3,sticky="w",padx=6)
        ttk.Label(mapf,text="Cột EMAIL:").grid(row=1,column=0,sticky="w")
        self.cb_email=ttk.Combobox(mapf,textvariable=self.col_email,width=28,state="readonly"); self.cb_email.grid(row=1,column=1,sticky="w",padx=6)
        ttk.Label(mapf,text="Cột TÊN FILE:").grid(row=1,column=2,sticky="w")
        self.cb_file=ttk.Combobox(mapf,textvariable=self.col_file,width=28,state="readonly"); self.cb_file.grid(row=1,column=3,sticky="w",padx=6)
        ttk.Checkbutton(mapf,text="Chỉ nạp các hàng đủ 4 cột trên",variable=self.filter_complete).grid(row=2,column=0,columnspan=4,sticky="w",pady=(6,2))
        ttk.Button(mapf,text="📂 Nạp danh sách theo ánh xạ",command=self.apply_mapping).grid(row=3,column=0,pady=(6,2))
        mapf.columnconfigure(1,weight=1); mapf.columnconfigure(3,weight=1)
        r+=1

        # Global attachments
        filesf = ttk.LabelFrame(main, text="Đính kèm chung (áp dụng cho TẤT CẢ email)")
        filesf.grid(row=r, column=0, columnspan=4, sticky="we", pady=(2,6))
        self.lb_extra = tk.Listbox(filesf, height=3)
        self.lb_extra.pack(side="left", fill="x", expand=True, padx=6, pady=6)
        fb = ttk.Frame(filesf); fb.pack(side="right", padx=6, pady=6)
        ttk.Button(fb, text="Thêm tệp…", command=self._add_extra_files).pack(fill="x")
        ttk.Button(fb, text="Xoá mục chọn", command=self._remove_selected_extra).pack(fill="x", pady=4)
        ttk.Button(fb, text="Xoá tất cả", command=self._clear_extra).pack(fill="x")
        self._refresh_extra_list()
        r+=1

        # Options & actions
        opts=ttk.Frame(main)
        opts.grid(row=r,column=0,columnspan=4,sticky="we",pady=(0,4))
        ttk.Checkbutton(opts,text="Thêm chữ ký Outlook (mặc định)",variable=self.append_signature).pack(side="left",padx=6)
        ttk.Checkbutton(opts,text="Dry run (không gửi thật)",variable=self.dry_run).pack(side="left",padx=6)
        ttk.Checkbutton(opts,text="Bắt buộc có file đính kèm cá nhân",variable=self.require_attach).pack(side="left",padx=6)
        ttk.Checkbutton(opts,text="Chỉ gửi các dòng đã chọn",variable=self.send_selected_only).pack(side="left",padx=6)
        ttk.Label(opts,text="Delay (giây):").pack(side="left",padx=(12,4))
        ttk.Entry(opts,textvariable=self.delay,width=6).pack(side="left")

        bar=ttk.Frame(main)
        bar.grid(row=r+1,column=0,columnspan=4,sticky="we",pady=(2,8))
        ttk.Button(bar,text="Chọn tất cả",command=self.select_all).pack(side="left")
        ttk.Button(bar,text="Bỏ chọn tất cả",command=self.clear_selection).pack(side="left",padx=4)
        ttk.Button(bar,text="▶ Gửi",command=self.start_send).pack(side="left",padx=12)
        ttk.Button(bar,text="⏹ Dừng",command=self.stop_send).pack(side="left")
        ttk.Button(bar,text="💾 Lưu cấu hình",command=self.save_settings).pack(side="right")

        # Bảng lớn hơn cho Excel
        self.cols=("sel","idx","ten","msnv","email","tenfile","attach","status")
        self.col_index={c:i for i,c in enumerate(self.cols)}
        self.tree=ttk.Treeview(main,columns=self.cols,show="headings",height=22)
        heads={"sel":"✓","idx":"#","ten":"tên","msnv":"msnv","email":"mail","tenfile":"tên file","attach":"đính kèm","status":"trạng thái"}
        for c in self.cols:
            self.tree.heading(c,text=heads[c])
            if c=="sel": self.tree.column(c,width=40,anchor="center")
            elif c in ("idx","status"): self.tree.column(c,width=100,anchor="w")
            else: self.tree.column(c,width=170,anchor="w")
        self.tree.grid(row=r+2,column=0,columnspan=4,sticky="nsew")
        self.tree.bind("<Button-1>",self.on_tree_click)

        ttk.Label(main,text="Log:").grid(row=r+3,column=0,sticky="w")
        self.log=tk.Text(main,height=7); self.log.grid(row=r+3,column=1,columnspan=3,sticky="nsew")

        main.columnconfigure(1,weight=1); main.columnconfigure(3,weight=1)
        main.rowconfigure(r+2,weight=1)

    # ----- global attachments helpers -----
    def _refresh_extra_list(self):
        self.lb_extra.delete(0,"end")
        for p in self.extra_files: self.lb_extra.insert("end",str(p))

    def _add_extra_files(self):
        paths=filedialog.askopenfilenames(title="Chọn tệp đính kèm chung")
        for p in paths:
            if p and Path(p) not in self.extra_files: self.extra_files.append(Path(p))
        self._refresh_extra_list()

    def _remove_selected_extra(self):
        for i in list(self.lb_extra.curselection())[::-1]:
            try: del self.extra_files[i]
            except Exception: pass
        self._refresh_extra_list()

    def _clear_extra(self):
        self.extra_files.clear(); self._refresh_extra_list()

    # ----- pickers -----
    def _pick_excel(self):
        p=filedialog.askopenfilename(filetypes=[("Excel","*.xlsx *.xls")])
        if p: self.excel_path.set(p)

    def _pick_dir(self):
        d=filedialog.askdirectory()
        if d: self.data_dir.set(d)

    # ----- logging -----
    def _log(self,msg:str): self.log_q.put(msg)
    def _drain_logs(self):
        try:
            while True:
                msg=self.log_q.get_nowait()
                self.log.insert("end",msg+"\n"); self.log.see("end")
        except queue.Empty:
            pass
        self.after(200,self._drain_logs)

    # ----- settings -----
    def save_settings(self):
        d={
            "excel_path":self.excel_path.get().strip(),
            "data_dir":self.data_dir.get().strip(),
            "sender":self.sender.get().strip(),
            "subject_tpl":self.subject_tpl.get().strip(),
            "html_src":self.html_src.get("1.0","end-1c"),
            "delay":float(self.delay.get() or 1.0),
            "dry_run":bool(self.dry_run.get()),
            "require_attach":bool(self.require_attach.get()),
            "send_selected_only":bool(self.send_selected_only.get()),
            "append_signature":bool(self.append_signature.get()),
            "filter_complete":bool(self.filter_complete.get()),
            "extra_files":[str(p) for p in self.extra_files],
        }
        save_cfg(d); messagebox.showinfo("OK", f"Đã lưu cấu hình: {APP_CFG}")

    # ----- read columns -----
    def read_columns(self):
        path=self.excel_path.get().strip()
        if not path or not Path(path).exists():
            messagebox.showerror("Lỗi","Hãy chọn file Excel hợp lệ."); return
        try:
            self.df_raw=pd.read_excel(path,dtype=str).fillna("")
            self.df_raw.columns=self.df_raw.columns.map(lambda c:str(c).strip())
        except Exception as e:
            messagebox.showerror("Lỗi đọc Excel",str(e)); return
        self.columns=list(self.df_raw.columns); values=self.columns.copy()
        self.cb_ten["values"]=values; self.cb_msnv["values"]=values
        self.cb_email["values"]=values; self.cb_file["values"]=values

        def find_like(keys):
            for c in self.columns:
                lc=c.lower()
                if any(k in lc for k in keys): return c
            return ""
        self.col_name.set(self.col_name.get() or find_like(["tên","ten","name"]) or (values[0] if values else ""))
        self.col_msnv.set(self.col_msnv.get() or find_like(["msnv","mã","ma nv","mnv","code"]) or (values[0] if values else ""))
        self.col_email.set(self.col_email.get() or find_like(["mail","email","e-mail"]) or (values[0] if values else ""))
        self.col_file.set(self.col_file.get() or find_like(["file","tệp","ten file","tên file"]) or (values[0] if values else ""))

        self._log(f"Đã nạp {len(self.columns)} cột. Chọn ánh xạ rồi bấm '📂 Nạp danh sách theo ánh xạ'.")

    # ----- apply mapping -----
    def apply_mapping(self):
        if self.df_raw is None:
            self.read_columns()
            if self.df_raw is None: return
        c_ten,c_msnv,c_email,c_file=(self.col_name.get().strip(),
                                     self.col_msnv.get().strip(),
                                     self.col_email.get().strip(),
                                     self.col_file.get().strip())
        for c in (c_ten,c_msnv,c_email,c_file):
            if c not in self.df_raw.columns:
                messagebox.showerror("Lỗi ánh xạ",f"Cột '{c}' không tồn tại trong Excel."); return

        must=bool(self.filter_complete.get())
        data_dir=Path(self.data_dir.get().strip() or ".")

        self.rows=[]; self.selected=set()
        for i,row in self.df_raw.iterrows():
            m:Dict[str,str]={}
            for col in self.df_raw.columns:
                cell=row.get(col,""); val="" if pd.isna(cell) else str(cell).strip()
                m[col]=val; m[norm_key(col)]=val
            ten,email,msnv,tenfile=m.get(c_ten,""),m.get(c_email,""),m.get(c_msnv,""),m.get(c_file,"")
            if must and (not ten or not email or not msnv or not tenfile): continue
            self.rows.append(RowJob(idx=i+1,row_map=m,ten=ten,msnv=msnv,tenfile=tenfile,email=email))

        for x in self.tree.get_children(): self.tree.delete(x)
        for r in self.rows:
            att=smart_find_attachment(data_dir,r.tenfile)
            att_name=att.name if att else "KHÔNG THẤY"
            extra = f" (+{len(self.extra_files)} chung)" if self.extra_files else ""
            self.tree.insert("", "end", iid=str(r.idx),
                             values=("☐", r.idx, r.ten, r.msnv, r.email, r.tenfile, att_name+extra, "Chưa gửi"))
        self._log(f"Nạp {len(self.rows)} hàng (lọc thiếu dữ liệu={must}).")

    # ----- table checkbox -----
    def on_tree_click(self,event):
        if self.tree.identify("region",event.x,event.y)!="cell": return
        if self.tree.identify_column(event.x)!="#1": return
        row_id=self.tree.identify_row(event.y)
        if not row_id: return
        vals=list(self.tree.item(row_id,"values"))
        if vals[0]=="☑":
            vals[0]="☐"; self.selected.discard(int(row_id))
        else:
            vals[0]="☑"; self.selected.add(int(row_id))
        self.tree.item(row_id,values=vals)

    def select_all(self):
        for iid in self.tree.get_children():
            v=list(self.tree.item(iid,"values"))
            if v[0]!="☑": v[0]="☑"; self.tree.item(iid,values=v)
            self.selected.add(int(iid))

    def clear_selection(self):
        for iid in self.tree.get_children():
            v=list(self.tree.item(iid,"values"))
            if v[0]!="☐": v[0]="☐"; self.tree.item(iid,values=v)
        self.selected.clear()

    # ----- preview -----
    def preview_current_html(self):
        if not self.rows:
            messagebox.showwarning("Chưa có dữ liệu","Hãy nạp danh sách trước."); return
        row=None
        if self.selected:
            first=sorted(self.selected)[0]
            row=next((r for r in self.rows if r.idx==first),None)
        if row is None: row=self.rows[0]
        ctx=dict(row.row_map); ctx.update({"ten":row.ten,"msnv":row.msnv,"tenfile":row.tenfile,"email":row.email})
        html_body=apply_variables(self.html_src.get("1.0","end-1c"), ctx)
        tmp=Path(os.environ.get("TEMP","."))/"mail_preview.html"
        tmp.write_text(html_body,encoding="utf-8")
        try: os.startfile(str(tmp))
        except Exception: messagebox.showinfo("Preview",f"Đã lưu {tmp}")

    # ----- send -----
    def start_send(self):
        if self.worker_thread and self.worker_thread.is_alive():
            messagebox.showwarning("Đang chạy","Đang gửi, vui lòng đợi hoặc bấm Dừng."); return
        if not self.rows:
            messagebox.showwarning("Chưa có danh sách","Hãy nạp danh sách trước."); return
        sender=self.sender.get().strip()
        if not sender:
            messagebox.showwarning("Thiếu người gửi","Nhập tài khoản Outlook gửi đi."); return
        data_dir=Path(self.data_dir.get().strip() or ".")
        if not data_dir.exists():
            messagebox.showerror("Lỗi","Thư mục Data không tồn tại."); return

        subject_tpl=self.subject_tpl.get().strip() or DEFAULT_SUBJECT
        delay=float(self.delay.get() or 1.0)
        dry=bool(self.dry_run.get())
        must=bool(self.require_attach.get())
        only_sel=bool(self.send_selected_only.get())
        add_sig=bool(self.append_signature.get())

        if only_sel:
            rows=[r for r in self.rows if r.idx in self.selected]
            if not rows:
                messagebox.showwarning("Chưa chọn dòng","Bạn bật 'Chỉ gửi dòng đã chọn' nhưng chưa chọn dòng nào."); return
        else:
            rows=self.rows

        self.stop_flag=False
        sig_cached=get_default_signature_html()

        def worker():
            sent=0; skipped=0
            class SafeDict(dict):
                def __missing__(self,k): return "{"+k+"}"

            for r in rows:
                if self.stop_flag: self._log("Đã dừng theo yêu cầu."); break

                ctx=dict(r.row_map); ctx.update({"ten":r.ten,"msnv":r.msnv,"tenfile":r.tenfile,"email":r.email})
                try: subject=subject_tpl.format_map(SafeDict(**ctx))
                except Exception as e:
                    self._update_tree(r.idx,status=f"Lỗi subject: {e}"); skipped+=1; continue

                html_body=apply_variables(self.html_src.get("1.0","end-1c"), ctx)

                per=smart_find_attachment(data_dir,r.tenfile)
                att_list=[]
                if per: att_list.append(per)
                att_list.extend([p for p in self.extra_files if p.exists()])
                att_note=(per.name if per else "KHÔNG THẤY") + (f" (+{len(self.extra_files)} chung)" if self.extra_files else "")
                self._update_tree(r.idx, attach=att_note)

                if must and per is None:
                    self._update_tree(r.idx,status="Bỏ qua: thiếu file cá nhân"); skipped+=1; continue

                if dry:
                    self._log(f"[DRY] {r.email} | subj='{subject}' | attach='{att_note}'")
                    self._update_tree(r.idx,status="(DRY) OK"); sent+=1; continue

                try:
                    send_with_outlook(r.email, subject, None, att_list, sender, html_body, add_sig, sig_cached)
                    self._log(f"Đã gửi: {r.email} | {att_note}")
                    self._update_tree(r.idx,status="ĐÃ GỬI"); sent+=1
                    time.sleep(delay)
                except Exception as e:
                    self._log(f"[ERR] {r.email}: {e}")
                    self._update_tree(r.idx,status=f"Lỗi: {e}"); skipped+=1
            self._log(f"[SUMMARY] Sent={sent}, Skipped={skipped}, Total={len(rows)}")

        self.worker_thread=threading.Thread(target=worker,daemon=True); self.worker_thread.start()

    def stop_send(self): self.stop_flag=True

    def _update_tree(self, idx:int, attach:Optional[str]=None, status:Optional[str]=None):
        iid=str(idx)
        try:
            v=list(self.tree.item(iid,"values"))
            if attach is not None: v[self.col_index["attach"]]=attach
            if status is not None: v[self.col_index["status"]]=status
            self.tree.item(iid,values=v); self.update_idletasks()
        except Exception: pass


# ----- Outlook send (kèm chữ ký, nhiều file đính kèm) -----
def send_with_outlook(to_addr: str, subject: str, body_text: Optional[str],
                      attachments: List[Path], sender_account: str,
                      html_body: Optional[str] = None,
                      append_signature: bool = True,
                      sig_html_cached: Optional[str] = None):
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

    if html_body is None:
        mail.Body = body_text or ""
    else:
        user_frag = extract_body_fragment(html_body)
        if append_signature:
            sig_html = sig_html_cached or get_default_signature_html()
            if sig_html:
                sig_frag = extract_body_fragment(sig_html)
                mail.BodyFormat = 2
                mail.HTMLBody = user_frag + "<br><br>" + sig_frag
            else:
                mail.Display(False)          # để Outlook tự chèn chữ ký
                current = mail.HTMLBody
                mail.HTMLBody = user_frag + "<br><br>" + current
        else:
            mail.BodyFormat = 2
            mail.HTMLBody = user_frag

    for p in attachments:
        try:
            if p and Path(p).exists(): mail.Attachments.Add(str(p))
        except Exception:
            pass

    mail.Send()


if __name__ == "__main__":
    app = App()
    app.mainloop()
