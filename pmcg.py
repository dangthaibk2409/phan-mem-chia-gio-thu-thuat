import customtkinter as ctk; from tkinter import ttk, messagebox, filedialog; import tkinter as tk
import sqlite3, pandas as pd, os, ctypes, random, json, traceback, sys, time, threading, urllib.request; from datetime import datetime

CURRENT_VERSION = "1.0"
URL_VERSION = "https://raw.githubusercontent.com/bacsi/phongkham/main/version.txt"
URL_FILE = "https://raw.githubusercontent.com/bacsi/phongkham/main/pmcg.py" # Hoặc pmcg.exe

DB_NAME = 'phongkham_v115.db'
def khoi_tao_db():
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    tbls = ["DanhSachMay (id INTEGER PRIMARY KEY AUTOINCREMENT, ten_loai TEXT, ma_may TEXT, trang_thai TEXT DEFAULT 'Sẵn sàng')", "ThuThuat (id INTEGER PRIMARY KEY AUTOINCREMENT, ten TEXT, loai_may TEXT, thoi_gian_may INTEGER, thoi_gian_nguoi INTEGER, loai_chuyen_mon TEXT DEFAULT 'YHCT', can_rut_may INTEGER DEFAULT 1, can_nguoi_phu INTEGER DEFAULT 0, phan_loai TEXT DEFAULT 'Chưa phân loại', ds_nguoi_phu TEXT DEFAULT '')", "PhongThuThuat (id INTEGER PRIMARY KEY AUTOINCREMENT, ten_phong TEXT, bac_si TEXT, ktv TEXT, danh_sach_may TEXT DEFAULT '', so_giuong INTEGER DEFAULT 15, danh_sach_giuong TEXT DEFAULT '')", "BenhNhan (id INTEGER PRIMARY KEY AUTOINCREMENT, ten TEXT, hsba TEXT, nam_sinh TEXT, gio_vao TEXT, gio_ban TEXT, gio_ra TEXT, phong TEXT, bac_si TEXT, thu_thuat TEXT, loai_bn TEXT DEFAULT 'Nội trú', ngay_vao TEXT DEFAULT '')", "NhanSu (id INTEGER PRIMARY KEY AUTOINCREMENT, ten TEXT, vai_tro TEXT, thoi_gian_lam TEXT, ky_nang TEXT, gio_ban TEXT, nguoi_thay_the TEXT DEFAULT '', trang_thai TEXT DEFAULT 'Đi làm')", "LichTrinh (id INTEGER PRIMARY KEY AUTOINCREMENT, id_bn INTEGER, ten_bn TEXT, ten_tt TEXT, ma_may TEXT, ten_ns TEXT, phong TEXT, gio_bat_dau INTEGER, gio_ket_nguoi INTEGER, gio_ket_may INTEGER)", "AppState (id INTEGER PRIMARY KEY, key TEXT UNIQUE, value TEXT)"]
    for t in tbls: cur.execute(f"CREATE TABLE IF NOT EXISTS {t}")
    
    def add_col(tbl, col, dtype):
        try:
            cols = [r[1] for r in cur.execute(f"PRAGMA table_info({tbl})").fetchall()]
            if col not in cols: cur.execute(f"ALTER TABLE {tbl} ADD COLUMN {col} {dtype}")
        except: pass

    add_col("ThuThuat", "loai_chuyen_mon", "TEXT DEFAULT 'YHCT'")
    add_col("ThuThuat", "can_rut_may", "INTEGER DEFAULT 1")
    add_col("ThuThuat", "can_nguoi_phu", "INTEGER DEFAULT 0")
    add_col("ThuThuat", "phan_loai", "TEXT DEFAULT 'Chưa phân loại'")
    add_col("ThuThuat", "ds_nguoi_phu", "TEXT DEFAULT ''")
    add_col("PhongThuThuat", "danh_sach_may", "TEXT DEFAULT ''")
    add_col("PhongThuThuat", "so_giuong", "INTEGER DEFAULT 15")
    add_col("PhongThuThuat", "danh_sach_giuong", "TEXT DEFAULT ''")
    add_col("BenhNhan", "loai_bn", "TEXT DEFAULT 'Nội trú'")
    add_col("BenhNhan", "ngay_vao", "TEXT DEFAULT ''")
    add_col("NhanSu", "nguoi_thay_the", "TEXT DEFAULT ''")
    add_col("NhanSu", "trang_thai", "TEXT DEFAULT 'Đi làm'")

    try: cur.execute("UPDATE NhanSu SET trang_thai='Nghỉ cả ngày' WHERE trang_thai='Nghỉ'")
    except: pass
    if cur.execute("SELECT COUNT(*) FROM DanhSachMay").fetchone()[0] == 0:
        md = [('điện châm', f'đc{i}', 'Sẵn sàng') for i in range(1, 29)] + [('điện xung', f'đx{i}', 'Sẵn sàng') for i in range(1, 5)] + [('siêu âm', 'sa', 'Sẵn sàng'), ('sóng ngắn', 'sn1', 'Sẵn sàng'), ('kéo giãn', 'kg', 'Sẵn sàng')] + [('thủy châm', f'tc{i}', 'Sẵn sàng') for i in range(1, 3)] + [('cấy chỉ', f'cchỉ{i}', 'Sẵn sàng') for i in range(1, 3)]
        cur.executemany("INSERT INTO DanhSachMay (ten_loai, ma_may, trang_thai) VALUES (?,?,?)", md)
    conn.commit(); conn.close()
    try: ctypes.windll.kernel32.SetFileAttributesW(DB_NAME, 2)
    except: pass

khoi_tao_db()
sel_1 = sel_2 = sel_3 = sel_4 = sel_5 = None; full_bn_list = []; sel_ns_slot = None 
g_sched = []; g_rot = []; g_staff = {}; g_proc = {}; g_req = {}; g_tl = {}; g_ca = {}; sat_cache = {}

def t2m(t_str):
    if not t_str or ":" not in t_str: return 0
    try: h, m = map(int, t_str.split(':')); return h * 60 + m
    except: return 0
def m2t(mins): return f"{mins//60:02d}:{mins%60:02d}"
def is_overlap(s1, e1, s2, e2): return max(s1, s2) < min(e1, e2)
def auto_format_time(event, entry):
    if event.keysym in ['BackSpace', 'Delete', 'Left', 'Right', 'Tab', 'Return', 'Up', 'Down']: return
    t = "".join(c for c in entry.get() if c.isdigit()); t = t[:4]; fmt = t[:2] + ":" + t[2:] if len(t) >= 2 else t
    if entry.get() != fmt: entry.delete(0, 'end'); entry.insert(0, fmt)
def auto_format_date(event, entry):
    if event.keysym in ['BackSpace', 'Delete', 'Left', 'Right', 'Tab', 'Return', 'Up', 'Down']: return
    t = "".join(c for c in entry.get() if c.isdigit()); t = t[:8]; fmt = t[:2] + "/" + t[2:4] + "/" + t[4:] if len(t) > 4 else (t[:2] + "/" + t[2:] if len(t) > 2 else t)
    if entry.get() != fmt: entry.delete(0, 'end'); entry.insert(0, fmt)
def gan_phim_enter(widgets, func):
    for w in widgets: w.bind("<Return>", func)
def sort_busy_slots(busy_str):
    if not busy_str: return ""
    slots = [s.strip() for s in busy_str.split(",") if s.strip()]
    def get_sort_key(slot):
        try: return t2m((slot.split(")")[-1].strip() if ")" in slot else slot).split("-")[0].strip())
        except: return 0
    slots.sort(key=get_sort_key); return ", ".join(slots)
def get_col_letter(col_idx):
    letter = ''
    while col_idx > 0: col_idx, remainder = divmod(col_idx - 1, 26); letter = chr(65 + remainder) + letter
    return letter
def auto_format_excel(path, dict_dfs):
    try:
        with pd.ExcelWriter(path, engine='openpyxl') as writer:
            for sheet_name, df in dict_dfs.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False); ws = writer.sheets[sheet_name]
                for i, col in enumerate(df.columns, 1): ws.column_dimensions[get_col_letter(i)].width = min(max(df[col].astype(str).map(len).max() if not df.empty else 0, len(str(col))) + 2, 80)
    except:
        with pd.ExcelWriter(path) as writer:
            for sheet_name, df in dict_dfs.items(): df.to_excel(writer, sheet_name=sheet_name, index=False)
def tao_file_mau(cols, f_name):
    p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f"Mau_{f_name}"); 
    if p: auto_format_excel(p, {"Mau": pd.DataFrame(columns=cols)}); messagebox.showinfo("Xong", f"Đã tạo mẫu:\n{p}")
def xuat_excel_db(tbl, cols, f_name):
    p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f_name); 
    if p: 
        try:
            conn = sqlite3.connect(DB_NAME); df = pd.read_sql_query(f"SELECT {', '.join(cols)} FROM {tbl}", conn); conn.close()
            auto_format_excel(p, {tbl: df}); messagebox.showinfo("Thành công", "Đã xuất dữ liệu!")
        except Exception as e: messagebox.showerror("Lỗi", str(e))
def nhap_excel_chung(tbl, cols, refresh_func):
    p = filedialog.askopenfilename(filetypes=[("Excel/CSV", "*.xlsx *.csv")]); 
    if p:
        try:
            df = pd.read_csv(p).fillna("") if p.endswith('.csv') else pd.read_excel(p).fillna(""); df_clean = df[cols]; conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
            if cur.execute(f"SELECT COUNT(*) FROM {tbl}").fetchone()[0] > 0:
                c = messagebox.askyesnocancel("Dữ liệu", "YES: Ghi đè. NO: Thêm tiếp. CANCEL: Hủy.")
                if c is None: conn.close(); return
                elif c is True: cur.execute(f"DELETE FROM {tbl}"); cur.execute(f"DELETE FROM sqlite_sequence WHERE name='{tbl}'")
            df_clean.to_sql(tbl, conn, if_exists='append', index=False); conn.commit(); conn.close(); refresh_func(); messagebox.showinfo("Thành công", "Đã nạp!")
        except Exception as e: messagebox.showerror("Lỗi", str(e))
def xuat_excel_lich_trinh(tree, f_name):
    if not tree.get_children() and not g_rot: return messagebox.showwarning("Cảnh báo", "Chưa có dữ liệu!")
    p = filedialog.asksaveasfilename(defaultextension=".xlsx", initialfile=f_name)
    if p:
        try:
            df_lich = pd.DataFrame([tree.item(i)['values'] for i in tree.get_children()], columns=[tree.heading(c)['text'] for c in tree['columns']])
            df_rot = pd.DataFrame(g_rot)
            if not df_rot.empty:
                df_rot.rename(columns={'bn': 'Tên Bệnh Nhân', 'ns': 'Năm Sinh', 'tt': 'Thủ Thuật', 'room': 'Phòng', 'staff': 'Nhân Sự', 'reason': 'Lý do'}, inplace=True); df_rot.insert(0, 'STT', range(1, 1 + len(df_rot)))
            else: df_rot = pd.DataFrame(columns=["STT", "Tên Bệnh Nhân", "Năm Sinh", "Thủ Thuật", "Phòng", "Nhân Sự", "Lý do"])
            auto_format_excel(p, {"Lịch Y Lệnh Đã Xếp": df_lich, "Ca Chưa Xếp Được": df_rot}); messagebox.showinfo("Thành công", "Đã xuất file!")
        except Exception as e: messagebox.showerror("Lỗi", str(e))
def luu_trang_thai():
    try:
        conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
        state = {'g_sched': g_sched, 'g_rot': g_rot, 'g_staff': g_staff, 'g_proc': g_proc, 'g_req': g_req, 'g_tl': g_tl, 'g_ca': g_ca}
        cur.execute("INSERT OR REPLACE INTO AppState (id, key, value) VALUES (1, 'main_state', ?)", (json.dumps(state),)); conn.commit(); conn.close()
    except Exception: pass
def tai_trang_thai():
    global g_sched, g_rot, g_staff, g_proc, g_req, g_tl, g_ca
    try:
        conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("SELECT value FROM AppState WHERE key='main_state'"); row = cur.fetchone(); conn.close()
        if row:
            st = json.loads(row[0]); g_sched = st.get('g_sched', []); g_rot = st.get('g_rot', []); g_staff = st.get('g_staff', {}); g_proc = st.get('g_proc', {}); g_req = st.get('g_req', {}); g_tl = st.get('g_tl', {}); g_ca = st.get('g_ca', {})
            if g_sched or g_rot: hien_thi_lich(); tai_thong_ke()
    except Exception: pass
def tree_sort(tv, col, reverse):
    l = [(tv.set(k, col), k) for k in tv.get_children('')]
    try: l.sort(key=lambda t: float(t[0]), reverse=reverse)
    except: l.sort(key=lambda t: str(t[0]).lower(), reverse=reverse)
    for index, (val, k) in enumerate(l): tv.move(k, '', index)
    for index, k in enumerate(tv.get_children(''), 1): vals = list(tv.item(k, 'values')); vals[0] = index; tv.item(k, values=vals)
    tv.heading(col, command=lambda: tree_sort(tv, col, not reverse))
def create_tree(parent, cols, widths):
    tree = ttk.Treeview(parent, columns=cols, show="headings", height=15)
    for c, w in zip(cols, widths):
        tree.heading(c, text=c, command=lambda _c=c: tree_sort(tree, _c, False))
        tree.column(c, width=w, anchor="center" if c in ["STT", "Giờ Vào", "Ngày Vào", "Năm Sinh", "Hệ", "Phân Loại", "Rút Máy", "Người Phụ"] else "w", stretch=False)
    tree.pack(side="left" if cols[0]=="STT" and len(cols)<5 else "right", fill="both", expand=True, padx=5, pady=5); return tree

def auto_scroll(event, sf):
    try:
        w = event.widget; y = w.winfo_rooty() - sf._parent_frame.winfo_rooty()
        th = sf._parent_frame.winfo_reqheight(); vh = sf._parent_canvas.winfo_height()
        if th <= vh or th == 0: return
        y0, y1 = sf._parent_canvas.yview(); wt = y / th; wb = (y + w.winfo_reqheight()) / th
        if wb > y1: sf._parent_canvas.yview_moveto(wb - vh/th + 0.05)
        elif wt < y0: sf._parent_canvas.yview_moveto(wt - 0.05)
    except: pass

def make_row(p, l, wc, **kw):
    f = ctk.CTkFrame(p, fg_color="transparent"); f.pack(fill="x", pady=4, padx=10)
    ctk.CTkLabel(f, text=l, width=80, anchor="w", font=("Arial", 12, "bold"), text_color="#2c3e50").pack(side="left")
    w = wc(f, width=210, **kw); w.pack(side="right"); return w
def make_compact_row(p, l, wc, **kw):
    f = ctk.CTkFrame(p, fg_color="transparent"); f.pack(fill="x", pady=2, padx=5)
    ctk.CTkLabel(f, text=l, width=120, anchor="w", font=("Arial", 11, "bold"), text_color="#2c3e50").pack(side="left")
    w = wc(f, width=80, **kw); w.pack(side="right"); return w
def make_double_row(p, l, wc, p1, p2, **kw):
    f = ctk.CTkFrame(p, fg_color="transparent"); f.pack(fill="x", pady=4, padx=10)
    ctk.CTkLabel(f, text=l, width=80, anchor="w", font=("Arial", 12, "bold"), text_color="#2c3e50").pack(side="left")
    fr = ctk.CTkFrame(f, fg_color="transparent"); fr.pack(side="right")
    w1 = wc(fr, width=100, placeholder_text=p1, **kw); w1.pack(side="left", padx=(0,10))
    w2 = wc(fr, width=100, placeholder_text=p2, **kw); w2.pack(side="left")
    return w1, w2
def tao_btn(parent, c_add, c_edit, c_del, c_clear, cols, f_mau, tbl, tree, func, f_xuat):
    f1 = ctk.CTkFrame(parent, fg_color="transparent"); f1.pack(pady=5)
    ctk.CTkButton(f1, text="Thêm", width=65, command=c_add).grid(row=0, column=0, padx=2); ctk.CTkButton(f1, text="Sửa", width=65, fg_color="#f39c12", command=c_edit).grid(row=0, column=1, padx=2)
    ctk.CTkButton(f1, text="Xóa", width=50, fg_color="#e74c3c", command=c_del).grid(row=0, column=2, padx=2); ctk.CTkButton(f1, text="Mới", width=50, fg_color="gray", command=c_clear).grid(row=0, column=3, padx=2)
    f2 = ctk.CTkFrame(parent, fg_color="transparent"); f2.pack(pady=2)
    ctk.CTkButton(f2, text="Tải Mẫu", width=60, fg_color="gray", command=lambda: tao_file_mau(cols, f_mau)).grid(row=0, column=0, padx=2)
    ctk.CTkButton(f2, text="Nhập", width=80, fg_color="#34495e", command=lambda: nhap_excel_chung(tbl, cols, func)).grid(row=0, column=1, padx=2)
    ctk.CTkButton(f2, text="Xuất", width=80, fg_color="#16a085", command=lambda: xuat_excel_db(tbl, cols, f_xuat)).grid(row=0, column=2, padx=2)

ctk.set_appearance_mode("Light"); app = ctk.CTk()
app.title(f"PHẦN MỀM CHIA GIỜ THỦ THUẬT made by DPT - V{CURRENT_VERSION} (BẢN CHÍNH THỨC)")
app.geometry("1300x750"); app.withdraw() 

# === GIAO DIỆN TỰ ĐỘNG CẬP NHẬT (OTA UPDATER) ===
class UpdaterUI(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Cập nhật hệ thống")
        self.geometry("420x160"); self.resizable(False, False)
        self.attributes('-topmost', True)
        self.protocol("WM_DELETE_WINDOW", self.disable_close)
        
        self.update_idletasks()
        ws = self.winfo_screenwidth(); hs = self.winfo_screenheight()
        self.geometry('%dx%d+%d+%d' % (420, 160, (ws/2)-(420/2), (hs/2)-(160/2)))

        self.lbl_st = ctk.CTkLabel(self, text="Đang kiểm tra phiên bản mới...", font=("Arial", 15, "bold"), text_color="#2980b9")
        self.lbl_st.pack(pady=(30, 15))
        self.prog = ctk.CTkProgressBar(self, width=320, progress_color="#27ae60")
        self.prog.set(0); self.prog.pack()
        self.lbl_v = ctk.CTkLabel(self, text=f"Phiên bản hiện tại: V{CURRENT_VERSION}", font=("Arial", 11), text_color="gray")
        self.lbl_v.pack(side="bottom", pady=8)
        
        threading.Thread(target=self.run_update, daemon=True).start()

    def disable_close(self): pass

    def run_update(self):
        time.sleep(0.8); self.prog.set(0.3)
        try:
            # KIỂM TRA PHIÊN BẢN TRÊN GITHUB (Mở comment 3 dòng dưới khi có link thật)
            # req = urllib.request.Request(URL_VERSION, headers={'Cache-Control': 'no-cache'})
            # with urllib.request.urlopen(req, timeout=4) as resp:
            #     latest_ver = resp.read().decode('utf-8').strip()
            latest_ver = CURRENT_VERSION  # Giả lập: Hiện tại đang là bản mới nhất
            
            if latest_ver != CURRENT_VERSION:
                self.lbl_st.configure(text=f"Đã có bản V{latest_ver}. Đang tải dữ liệu...")
                self.prog.set(0.6)
                
                # TẢI FILE MỚI VỀ VÀ GHI ĐÈ
                # with urllib.request.urlopen(URL_FILE, timeout=10) as resp: data = resp.read()
                # current_file = sys.argv[0]
                # if current_file.endswith('.exe'): # Vượt quyền Windows khi update file .exe đang chạy
                #     old_file = current_file + ".old"
                #     if os.path.exists(old_file): os.remove(old_file)
                #     os.rename(current_file, old_file)
                # with open(current_file, 'wb') as f: f.write(data)
                
                self.prog.set(1.0)
                self.lbl_st.configure(text="Cập nhật thành công! Đang khởi động lại...")
                time.sleep(1.5)
                os.execl(sys.executable, sys.executable, *sys.argv)
            else:
                self.lbl_st.configure(text="Phần mềm đã ở phiên bản mới nhất!")
                self.prog.set(1.0)
                time.sleep(0.8)
                self.finish()
        except Exception:
            self.lbl_st.configure(text="Bỏ qua cập nhật (Không có kết nối máy chủ).")
            self.prog.set(1.0)
            time.sleep(0.8)
            self.finish()

    def finish(self):
        self.destroy()
        app.deiconify() # Hiện phần mềm chính
        app.after(0, lambda: app.state('zoomed') if os.name == 'nt' else app.attributes('-zoomed', True))

UpdaterUI(app)

tabview = ctk.CTkTabview(app, segmented_button_selected_color="#3498db"); tabview.pack(fill="both", expand=True, padx=10, pady=5)
for name in ["1. Máy móc", "2. Thủ thuật", "3. Nhân sự", "4. Phòng", "5. Bệnh nhân", "6. AUTO XẾP LỊCH", "7. Thống kê", "8. TRỰC THỨ 7"]: tabview.add(name)

# --- TAB 1 ---
t1 = tabview.tab("1. Máy móc")
f1l = ctk.CTkFrame(t1, width=350); f1l.pack(side="left", fill="y", padx=10, pady=10); f1l.pack_propagate(False)
f1_bot = ctk.CTkFrame(f1l, fg_color="transparent"); f1_bot.pack(side="bottom", fill="x", pady=5)
f1_top = ctk.CTkFrame(f1l, fg_color="transparent"); f1_top.pack(side="top", fill="x", expand=True)
e1_ten = make_row(f1_top, "Tên loại:", ctk.CTkEntry, placeholder_text="Tên loại máy")
e1_ma = make_row(f1_top, "Ký hiệu:", ctk.CTkEntry, placeholder_text="Ký hiệu (Mã)")
e1_sl = make_row(f1_top, "Số lượng:", ctk.CTkEntry, placeholder_text="Số lượng")
cb1_tt = make_row(f1_top, "Trạng thái:", ctk.CTkComboBox, values=["Sẵn sàng", "Bảo trì"]); cb1_tt.set("Sẵn sàng")
f1r = ctk.CTkFrame(t1, fg_color="transparent"); f1r.pack(side="right", fill="both", expand=True, padx=10, pady=10)
tree1 = create_tree(f1r, ("STT", "Loại Máy", "Mã Máy", "Trạng Thái"), [50, 200, 150, 150])
def clear_1(): global sel_1; sel_1 = None; e1_ten.delete(0, 'end'); e1_ma.delete(0, 'end'); e1_sl.delete(0, 'end'); cb1_tt.set("Sẵn sàng")
def tai_ds1():
    for r in tree1.get_children(): tree1.delete(r)
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    for stt, row in enumerate(cur.execute("SELECT * FROM DanhSachMay"), 1): tree1.insert("", "end", iid=row[0], values=(stt, *row[1:]))
    conn.close(); tai_ds2()
def click_1(e):
    global sel_1; sel = tree1.selection()
    if sel: clear_1(); sel_1 = sel[0]; i = tree1.item(sel_1, 'values'); e1_ten.insert(0, i[1]); e1_ma.insert(0, i[2]); cb1_tt.set(i[3])
def add_1():
    try:
        conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); p, num = e1_ma.get(), int(e1_sl.get() if e1_sl.get() else 1)
        ex = [int(r[0][len(p):]) for r in cur.execute("SELECT ma_may FROM DanhSachMay WHERE ma_may LIKE ?", (p+'%',)).fetchall() if r[0][len(p):].isdigit()]; st = max(ex) if ex else 0
        for i in range(1, num + 1): cur.execute("INSERT INTO DanhSachMay (ten_loai, ma_may, trang_thai) VALUES (?,?,?)", (e1_ten.get(), f"{p}{st+i}", cb1_tt.get()))
        conn.commit(); conn.close(); tai_ds1(); clear_1(); e1_ten.focus() 
    except: pass
def edit_1():
    if sel_1: conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("UPDATE DanhSachMay SET ten_loai=?, ma_may=?, trang_thai=? WHERE id=?", (e1_ten.get(), e1_ma.get(), cb1_tt.get(), sel_1)); conn.commit(); conn.close(); tai_ds1(); clear_1(); e1_ten.focus()
def del_1(*args):
    if sel_1: conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("DELETE FROM DanhSachMay WHERE id=?", (sel_1,)); conn.commit(); conn.close(); tai_ds1(); clear_1()
tree1.bind("<ButtonRelease-1>", click_1); tree1.bind("<Delete>", del_1); tree1.bind("<BackSpace>", del_1); tree1.bind("<Return>", lambda e: edit_1() if sel_1 else add_1()); gan_phim_enter([e1_ten, e1_ma, e1_sl], lambda e: edit_1() if sel_1 else add_1())
tao_btn(f1_bot, add_1, edit_1, del_1, clear_1, ["ten_loai", "ma_may", "trang_thai"], "MayMoc", "DanhSachMay", tree1, tai_ds1, "DS_May")

# --- TAB 2 ---
t2 = tabview.tab("2. Thủ thuật")
f2l = ctk.CTkFrame(t2, width=350); f2l.pack(side="left", fill="y", padx=10, pady=10); f2l.pack_propagate(False)
f2_bot = ctk.CTkFrame(f2l, fg_color="transparent"); f2_bot.pack(side="bottom", fill="x", pady=5)
f2_top = ctk.CTkFrame(f2l, fg_color="transparent"); f2_top.pack(side="top", fill="x")
e2_ten = make_row(f2_top, "Tên TT:", ctk.CTkEntry, placeholder_text="Tên thủ thuật")
cb2_loai = make_row(f2_top, "Hệ:", ctk.CTkComboBox, values=["YHCT", "PHCN"]); cb2_loai.set("YHCT")
cb2_phanloai = make_row(f2_top, "Phân loại:", ctk.CTkComboBox, values=["Loại 1", "Loại 2", "Loại 3", "Chưa phân loại"]); cb2_phanloai.set("Chưa phân loại")
cb2_may = make_row(f2_top, "Loại Máy:", ctk.CTkComboBox, values=["Thủ công"]); cb2_may.set("Thủ công")
e2_tgm = make_row(f2_top, "TG Máy:", ctk.CTkEntry, placeholder_text="Phút")
e2_tgn = make_row(f2_top, "TG Người:", ctk.CTkEntry, placeholder_text="Phút")
v2_rut = ctk.IntVar(value=1); ctk.CTkCheckBox(f2_top, text="Có Điều dưỡng rút/tháo máy", variable=v2_rut, text_color="#2980b9", font=("Arial", 11, "bold")).pack(pady=4, anchor="w", padx=25)
v2_phu = ctk.IntVar(value=0); ctk.CTkCheckBox(f2_top, text="Yêu cầu kíp (1 Chính + 1 Phụ)", variable=v2_phu, text_color="#8e44ad", font=("Arial", 11, "bold")).pack(pady=4, anchor="w", padx=25)
fr2_phu = ctk.CTkScrollableFrame(f2l, label_text="Cấp phép Điều dưỡng phụ/rút"); fr2_phu.pack(side="top", fill="both", expand=True, pady=5); v2_nv_phu = {}
f2r = ctk.CTkFrame(t2, fg_color="transparent"); f2r.pack(side="right", fill="both", expand=True, padx=10, pady=10)
tree2 = create_tree(f2r, ("STT", "Tên TT", "Hệ", "Phân Loại", "Máy", "TG Máy", "TG NV", "Rút Máy", "Người Phụ", "DS Cấp Phép"), [30, 140, 50, 90, 80, 60, 60, 60, 70, 150])

def clear_2(): 
    global sel_2; sel_2 = None; e2_ten.delete(0, 'end'); e2_tgm.delete(0, 'end'); e2_tgn.delete(0, 'end'); cb2_loai.set("YHCT"); cb2_phanloai.set("Chưa phân loại"); v2_rut.set(1); v2_phu.set(0)
    for v in v2_nv_phu.values(): v.set("off")
def tai_ds2():
    for r in tree2.get_children(): tree2.delete(r)
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    for stt, r in enumerate(cur.execute("SELECT id, ten, loai_chuyen_mon, phan_loai, loai_may, thoi_gian_may, thoi_gian_nguoi, can_rut_may, can_nguoi_phu, ds_nguoi_phu FROM ThuThuat"), 1):
        tree2.insert("", "end", iid=r[0], values=(stt, r[1], r[2], r[3], r[4], r[5], r[6], "Có" if r[7]==1 else "Không", "Có" if r[8]==1 else "Không", r[9]))
    cur.execute("SELECT DISTINCT ten_loai FROM DanhSachMay"); cb2_may.configure(values=["Thủ công"] + [r[0] for r in cur.fetchall()])
    for w in fr2_phu.winfo_children(): w.destroy()
    v2_nv_phu.clear()
    for r in cur.execute("SELECT ten FROM NhanSu WHERE vai_tro='Điều dưỡng'"):
        v = ctk.StringVar(value="off"); v2_nv_phu[r[0]] = v; cb = ctk.CTkCheckBox(fr2_phu, text=r[0], variable=v, onvalue=r[0], offvalue="off"); cb.pack(anchor="w", padx=10, pady=2)
        cb.bind("<Return>", lambda e: edit_2() if sel_2 else add_2()); cb.bind("<FocusIn>", lambda ev, sf=fr2_phu: auto_scroll(ev, sf))
    conn.close(); tai_ds3()
def click_2(e):
    global sel_2; sel = tree2.selection(); 
    if sel: 
        clear_2(); sel_2 = sel[0]; i = tree2.item(sel_2, 'values'); e2_ten.insert(0, i[1]); cb2_loai.set(i[2]); cb2_phanloai.set(i[3]); cb2_may.set(i[4]); e2_tgm.insert(0, i[5]); e2_tgn.insert(0, i[6]); v2_rut.set(1 if i[7]=="Có" else 0); v2_phu.set(1 if i[8]=="Có" else 0)
        for nv in str(i[9]).split(", "):
            if nv in v2_nv_phu: v2_nv_phu[nv].set(nv)
def get_ds_phu(): return ", ".join([v.get() for v in v2_nv_phu.values() if v.get() not in ["off", "0", ""]])
def add_2():
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("INSERT INTO ThuThuat (ten, loai_chuyen_mon, phan_loai, loai_may, thoi_gian_may, thoi_gian_nguoi, can_rut_may, can_nguoi_phu, ds_nguoi_phu) VALUES (?,?,?,?,?,?,?,?,?)", (e2_ten.get(), cb2_loai.get(), cb2_phanloai.get(), cb2_may.get(), e2_tgm.get(), e2_tgn.get(), v2_rut.get(), v2_phu.get(), get_ds_phu())); conn.commit(); conn.close(); tai_ds2(); clear_2(); e2_ten.focus()
def edit_2():
    if sel_2: conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("UPDATE ThuThuat SET ten=?, loai_chuyen_mon=?, phan_loai=?, loai_may=?, thoi_gian_may=?, thoi_gian_nguoi=?, can_rut_may=?, can_nguoi_phu=?, ds_nguoi_phu=? WHERE id=?", (e2_ten.get(), cb2_loai.get(), cb2_phanloai.get(), cb2_may.get(), e2_tgm.get(), e2_tgn.get(), v2_rut.get(), v2_phu.get(), get_ds_phu(), sel_2)); conn.commit(); conn.close(); tai_ds2(); clear_2(); e2_ten.focus()
def del_2(*args):
    if sel_2: conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("DELETE FROM ThuThuat WHERE id=?", (sel_2,)); conn.commit(); conn.close(); tai_ds2(); clear_2()
tree2.bind("<ButtonRelease-1>", click_2); tree2.bind("<Delete>", del_2); tree2.bind("<BackSpace>", del_2); tree2.bind("<Return>", lambda e: edit_2() if sel_2 else add_2()); gan_phim_enter([e2_ten, e2_tgm, e2_tgn], lambda e: edit_2() if sel_2 else add_2())
tao_btn(f2_bot, add_2, edit_2, del_2, clear_2, ["ten", "loai_chuyen_mon", "phan_loai", "loai_may", "thoi_gian_may", "thoi_gian_nguoi", "can_rut_may", "can_nguoi_phu", "ds_nguoi_phu"], "ThuThuat", "ThuThuat", tree2, tai_ds2, "DS_ThuThuat")

def tai_ds_pri():
    for r in tree_pri.get_children(): tree_pri.delete(r)
    for r in tree_ns_ban.get_children(): tree_ns_ban.delete(r)
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    for r in cur.execute("SELECT id, ten, nam_sinh, gio_ra FROM BenhNhan WHERE gio_ra != '' AND gio_ra IS NOT NULL"): tree_pri.insert("", "end", iid=r[0], values=(r[0], f"{r[1]} ({r[2]})", r[3]))
    stt_ns = 1
    for r in cur.execute("SELECT ten, gio_ban FROM NhanSu WHERE gio_ban != '' AND gio_ban IS NOT NULL AND trang_thai != 'Nghỉ cả ngày'"):
        for slot in str(r[1]).split(","):
            if slot.strip(): tree_ns_ban.insert("", "end", values=(stt_ns, r[0], slot.strip())); stt_ns += 1
    conn.close()
def tai_cb_t6():
    global full_bn_list; conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    full_bn_list = [f"{r[0]} - {r[1]} ({r[2]})" for r in cur.execute("SELECT id, ten, nam_sinh FROM BenhNhan WHERE gio_ra = '' OR gio_ra IS NULL")]
    nss = [r[0] for r in cur.execute("SELECT ten FROM NhanSu WHERE trang_thai != 'Nghỉ cả ngày'")]
    cb6_ns.configure(values=nss if nss else ["Trống"]); conn.close(); tai_ds_pri()

# --- TAB 3 ---
t3 = tabview.tab("3. Nhân sự")
f3l = ctk.CTkFrame(t3, width=350); f3l.pack(side="left", fill="y", padx=10, pady=10); f3l.pack_propagate(False)
f3_bot = ctk.CTkFrame(f3l, fg_color="transparent"); f3_bot.pack(side="bottom", fill="x", pady=5)
f3_top = ctk.CTkFrame(f3l, fg_color="transparent"); f3_top.pack(side="top", fill="x")
cb3_trangthai = make_row(f3_top, "Trạng thái:", ctk.CTkComboBox, values=["Đi làm", "Nghỉ sáng", "Nghỉ chiều", "Nghỉ cả ngày"]); cb3_trangthai.set("Đi làm")
e3_n = make_row(f3_top, "Tên NV:", ctk.CTkEntry, placeholder_text="Tên nhân viên")
cb3_v = make_row(f3_top, "Vai trò:", ctk.CTkComboBox, values=["Bác sĩ", "Kỹ thuật viên", "Điều dưỡng"]); cb3_v.set("Bác sĩ")
e3_s1, e3_s2 = make_double_row(f3_top, "Ca Sáng:", ctk.CTkEntry, "Từ", "Đến")
e3_c1, e3_c2 = make_double_row(f3_top, "Ca Chiều:", ctk.CTkEntry, "Từ", "Đến")
for e in [e3_s1, e3_s2, e3_c1, e3_c2]: e.bind("<KeyRelease>", lambda ev, w=e: auto_format_time(ev, w))
cb3_thaythe = make_row(f3_top, "Thay thế:", ctk.CTkComboBox, values=["Không"])
e3_ban = make_row(f3_top, "Giờ bận:", ctk.CTkEntry, placeholder_text="VD: 08:00-09:00")
fr3_kn = ctk.CTkScrollableFrame(f3l, label_text="Kỹ năng"); fr3_kn.pack(side="top", fill="both", expand=True, pady=5); kn_vars = {}
f3r = ctk.CTkFrame(t3, fg_color="transparent"); f3r.pack(side="right", fill="both", expand=True, padx=10, pady=10)
tree3 = create_tree(f3r, ("STT", "Tên", "Vai Trò", "Trạng Thái", "Ca Làm", "Thay Thế Bởi", "Giờ Bận", "Kỹ Năng"), [30, 140, 100, 90, 100, 120, 180, 300])

def clear_3():
    global sel_3; sel_3 = None; e3_n.delete(0, 'end'); e3_ban.delete(0, 'end'); cb3_thaythe.set("Không"); cb3_trangthai.set("Đi làm")
    for e in [e3_s1, e3_s2, e3_c1, e3_c2]: e.delete(0, 'end')
    for v in kn_vars.values(): v.set("off")
def tai_ds3():
    for r in tree3.get_children(): tree3.delete(r)
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); all_staff = []
    for stt, r in enumerate(cur.execute("SELECT id, ten, vai_tro, trang_thai, thoi_gian_lam, nguoi_thay_the, gio_ban, ky_nang FROM NhanSu"), 1):
        tree3.insert("", "end", iid=r[0], values=(stt, *r[1:])); all_staff.append(r[1])
    cb3_thaythe.configure(values=["Không"] + all_staff)
    for w in fr3_kn.winfo_children(): w.destroy()
    kn_vars.clear()
    for r in cur.execute("SELECT ten FROM ThuThuat"):
        v = ctk.StringVar(value="off"); kn_vars[r[0]] = v; cb = ctk.CTkCheckBox(fr3_kn, text=r[0], variable=v, onvalue=r[0], offvalue="off"); cb.pack(anchor="w", padx=10, pady=2)
        cb.bind("<Return>", lambda e: edit_3() if sel_3 else add_3()); cb.bind("<FocusIn>", lambda ev, sf=fr3_kn: auto_scroll(ev, sf))
    
    for w in fr2_phu.winfo_children(): w.destroy()
    v2_nv_phu.clear()
    for r in cur.execute("SELECT ten FROM NhanSu WHERE vai_tro='Điều dưỡng'"):
        v = ctk.StringVar(value="off"); v2_nv_phu[r[0]] = v; cb = ctk.CTkCheckBox(fr2_phu, text=r[0], variable=v, onvalue=r[0], offvalue="off"); cb.pack(anchor="w", padx=10, pady=2)
        cb.bind("<Return>", lambda e: edit_2() if sel_2 else add_2()); cb.bind("<FocusIn>", lambda ev, sf=fr2_phu: auto_scroll(ev, sf))

    conn.close(); tai_ds4()

def click_3(e):
    global sel_3; sel = tree3.selection(); 
    if sel:
        clear_3(); sel_3 = sel[0]; i = tree3.item(sel_3, 'values'); e3_n.insert(0, i[1]); cb3_v.set(i[2]); cb3_trangthai.set(i[3])
        for c in str(i[4]).split(", "):
            p = c.split("-")
            if len(p) == 2:
                if int(p[0][:2]) < 12: e3_s1.insert(0, p[0]); e3_s2.insert(0, p[1])
                else: e3_c1.insert(0, p[0]); e3_c2.insert(0, p[1])
        cb3_thaythe.set(i[5] if i[5] else "Không")
        if i[6] and i[6] != "None": e3_ban.insert(0, i[6]) 
        for kn in str(i[7]).split(", "):
            if kn in kn_vars: kn_vars[kn].set(kn)
def get_ca3(): return ", ".join([f"{s}-{e}" for s, e in [(e3_s1.get(), e3_s2.get()), (e3_c1.get(), e3_c2.get())] if s and e])
def format_ban3(): return f"({datetime.now().strftime('%d/%m/%Y')}) {e3_ban.get().strip()}" if e3_ban.get().strip() and "(" not in e3_ban.get().strip() else e3_ban.get().strip()
def add_3():
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); kn = ", ".join([v.get() for v in kn_vars.values() if v.get() not in ["off", "0", ""]])
    cur.execute("INSERT INTO NhanSu (ten, vai_tro, thoi_gian_lam, ky_nang, gio_ban, nguoi_thay_the, trang_thai) VALUES (?,?,?,?,?,?,?)", (e3_n.get(), cb3_v.get(), get_ca3(), kn, sort_busy_slots(format_ban3()), cb3_thaythe.get() if cb3_thaythe.get() != "Không" else "", cb3_trangthai.get()))
    conn.commit(); conn.close(); tai_ds3(); clear_3(); e3_n.focus()
def edit_3():
    if sel_3:
        conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); kn = ", ".join([v.get() for v in kn_vars.values() if v.get() not in ["off", "0", ""]])
        cur.execute("UPDATE NhanSu SET ten=?, vai_tro=?, thoi_gian_lam=?, ky_nang=?, gio_ban=?, nguoi_thay_the=?, trang_thai=? WHERE id=?", (e3_n.get(), cb3_v.get(), get_ca3(), kn, sort_busy_slots(format_ban3()), cb3_thaythe.get() if cb3_thaythe.get() != "Không" else "", cb3_trangthai.get(), sel_3))
        conn.commit(); conn.close(); tai_ds3(); clear_3(); e3_n.focus()
def del_3(*args):
    if sel_3: conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("DELETE FROM NhanSu WHERE id=?", (sel_3,)); conn.commit(); conn.close(); tai_ds3(); clear_3()
tree3.bind("<ButtonRelease-1>", click_3); tree3.bind("<Delete>", del_3); tree3.bind("<BackSpace>", del_3); tree3.bind("<Return>", lambda e: edit_3() if sel_3 else add_3()); gan_phim_enter([e3_n, e3_s1, e3_s2, e3_c1, e3_c2, e3_ban], lambda e: edit_3() if sel_3 else add_3())
tao_btn(f3_bot, add_3, edit_3, del_3, clear_3, ["ten", "vai_tro", "trang_thai", "thoi_gian_lam", "ky_nang", "gio_ban", "nguoi_thay_the"], "NhanSu", "NhanSu", tree3, tai_ds3, "DS_NhanSu")

# --- TAB 4 ---
t4 = tabview.tab("4. Phòng")
f4l = ctk.CTkFrame(t4, width=350); f4l.pack(side="left", fill="y", padx=10, pady=10); f4l.pack_propagate(False)
f4_bot = ctk.CTkFrame(f4l, fg_color="transparent"); f4_bot.pack(side="bottom", fill="x", pady=5)
f4_top = ctk.CTkFrame(f4l, fg_color="transparent"); f4_top.pack(side="top", fill="x")
e4_t = make_row(f4_top, "Tên Phòng:", ctk.CTkEntry, placeholder_text="Tên Phòng")
e4_g = make_row(f4_top, "Số giường:", ctk.CTkEntry, placeholder_text="Mặc định: 15")
e4_dg = make_row(f4_top, "Các giường:", ctk.CTkEntry, placeholder_text="Để trống tự sinh (G1, G2..)")
f4_mid = ctk.CTkScrollableFrame(f4l); f4_mid.pack(side="top", fill="both", expand=True)
fr4_bs = ctk.CTkScrollableFrame(f4_mid, label_text="👩‍⚕️ Bác sĩ trong phòng:", height=80); fr4_bs.pack(fill="x", pady=2); v4bs = {}
fr4_ktv = ctk.CTkScrollableFrame(f4_mid, label_text="👨‍⚕️ KTV/ĐD trong phòng:", height=80); fr4_ktv.pack(fill="x", pady=2); v4kt = {}
fr4_may = ctk.CTkScrollableFrame(f4_mid, label_text="⚙️ Số Máy (Tự sinh mã):", height=120); fr4_may.pack(fill="both", expand=True, pady=2); v4may_entries = {}
f4r = ctk.CTkFrame(t4, fg_color="transparent"); f4r.pack(side="right", fill="both", expand=True, padx=10, pady=10)
tree4 = create_tree(f4r, ("STT", "Tên Phòng", "Bác sĩ", "KTV/ĐD", "Các Máy", "Số Giường", "Các Giường"), [30, 120, 160, 160, 200, 80, 250])

def update_machine_inputs(room_id=None):
    for w in fr4_may.winfo_children(): w.destroy()
    v4may_entries.clear(); conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    loai_list = [r[0] for r in cur.execute("SELECT DISTINCT ten_loai FROM DanhSachMay").fetchall()]
    current_counts = {l: 0 for l in loai_list}
    if room_id:
        r_may = cur.execute("SELECT danh_sach_may FROM PhongThuThuat WHERE id=?", (room_id,)).fetchone()
        if r_may and r_may[0] and str(r_may[0]) != 'None':
            assigned = [m.strip() for m in str(r_may[0]).split(',') if m.strip()]
            for m_code in assigned:
                r_loai = cur.execute("SELECT ten_loai FROM DanhSachMay WHERE ma_may=?", (m_code,)).fetchone()
                if r_loai: current_counts[r_loai[0]] += 1
    for loai in loai_list:
        e = make_compact_row(fr4_may, f"{loai}:", ctk.CTkEntry, placeholder_text="0")
        if current_counts[loai] > 0: e.insert(0, str(current_counts[loai]))
        e.bind("<FocusIn>", lambda ev, sf=fr4_may: auto_scroll(ev, sf))
        v4may_entries[loai] = e
    conn.close()

def tu_dong_chia_may(id_bo_qua=None):
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); all_machines = {}
    for r in cur.execute("SELECT ten_loai, ma_may FROM DanhSachMay").fetchall():
        if r[0] not in all_machines: all_machines[r[0]] = []
        all_machines[r[0]].append(r[1])
    q = "SELECT danh_sach_may FROM PhongThuThuat"; p = ()
    if id_bo_qua: q += " WHERE id != ?"; p = (id_bo_qua,)
    assigned_machines = set()
    for r in cur.execute(q, p).fetchall():
        if r[0] and r[0] != 'None':
            for m in str(r[0]).split(','): assigned_machines.add(m.strip())
    free_machines = {}
    for loai, m_list in all_machines.items():
        free_machines[loai] = [m for m in m_list if m not in assigned_machines]
        free_machines[loai].sort(key=lambda x: int("".join(filter(str.isdigit, x)) or 0))
    new_m = []; warnings = []
    for loai, e in v4may_entries.items():
        val = e.get().strip()
        if val.isdigit() and int(val) > 0:
            num_req = int(val); available = free_machines.get(loai, [])
            if num_req > len(available):
                warnings.append(f"- {loai}: Yêu cầu {num_req}, Kho chỉ còn {len(available)}")
                num_req = len(available)
            new_m.extend(available[:num_req])
    conn.close()
    if warnings: messagebox.showwarning("Thiếu máy", "Kho máy khai báo ở Tab 1 không đủ dư:\n" + "\n".join(warnings))
    return ", ".join(new_m)

def clear_4():
    global sel_4; sel_4 = None; e4_t.delete(0, 'end'); e4_g.delete(0, 'end'); e4_g.insert(0, '15'); e4_dg.delete(0, 'end')
    for v in list(v4bs.values()) + list(v4kt.values()): v.set("off")
    update_machine_inputs(None) 
def tai_ds4():
    for r in tree4.get_children(): tree4.delete(r)
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    for stt, row in enumerate(cur.execute("SELECT id, ten_phong, bac_si, ktv, danh_sach_may, so_giuong, danh_sach_giuong FROM PhongThuThuat"), 1): tree4.insert("", "end", iid=row[0], values=(stt, *row[1:]))
    for w in fr4_bs.winfo_children(): w.destroy()
    for w in fr4_ktv.winfo_children(): w.destroy()
    v4bs.clear(); v4kt.clear()
    for r in cur.execute("SELECT ten FROM NhanSu WHERE vai_tro='Bác sĩ'"):
        v = ctk.StringVar(value="off"); v4bs[r[0]] = v; cb = ctk.CTkCheckBox(fr4_bs, text=r[0], variable=v, onvalue=r[0], offvalue="off"); cb.pack(anchor="w", padx=10, pady=2)
        cb.bind("<Return>", lambda e: edit_4() if sel_4 else add_4()); cb.bind("<FocusIn>", lambda ev, sf=fr4_bs: auto_scroll(ev, sf))
    for r in cur.execute("SELECT ten FROM NhanSu WHERE vai_tro IN ('Kỹ thuật viên', 'Điều dưỡng')"):
        v = ctk.StringVar(value="off"); v4kt[r[0]] = v; cb = ctk.CTkCheckBox(fr4_ktv, text=r[0], variable=v, onvalue=r[0], offvalue="off"); cb.pack(anchor="w", padx=10, pady=2)
        cb.bind("<Return>", lambda e: edit_4() if sel_4 else add_4()); cb.bind("<FocusIn>", lambda ev, sf=fr4_ktv: auto_scroll(ev, sf))
    conn.close(); update_machine_inputs(sel_4); tai_ds5()
def click_4(e):
    global sel_4; sel = tree4.selection(); 
    if sel:
        for v in list(v4bs.values()) + list(v4kt.values()): v.set("off")
        sel_4 = sel[0]; i = tree4.item(sel_4, 'values'); e4_t.delete(0, 'end'); e4_t.insert(0, i[1]); e4_g.delete(0, 'end'); e4_g.insert(0, i[5] if i[5] else '15'); e4_dg.delete(0, 'end'); e4_dg.insert(0, i[6] if str(i[6]) != 'None' else '')
        for bs in str(i[2]).split(", "):
            if bs in v4bs: v4bs[bs].set(bs)
        for kt in str(i[3]).split(", "):
            if kt in v4kt: v4kt[kt].set(kt)
        update_machine_inputs(sel_4)

def tu_dong_chia_giuong(so_giuong, id_bo_qua=None):
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); q = "SELECT danh_sach_giuong FROM PhongThuThuat"; p = ()
    if id_bo_qua: q += " WHERE id != ?"; p = (id_bo_qua,)
    max_g = 0
    for r in cur.execute(q, p).fetchall():
        if r[0] and r[0] != 'None':
            for b in str(r[0]).split(','):
                num = "".join(filter(str.isdigit, b))
                if num: max_g = max(max_g, int(num))
    conn.close(); return ", ".join([f"G{i}" for i in range(max_g + 1, max_g + 1 + so_giuong)])
def add_4():
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    bs = ", ".join([v.get() for v in v4bs.values() if v.get() not in ["off", "0", ""]]); kt = ", ".join([v.get() for v in v4kt.values() if v.get() not in ["off", "0", ""]])
    may = tu_dong_chia_may()
    so_g = int(e4_g.get()) if e4_g.get().isdigit() else 15
    ds_g = e4_dg.get().strip() if e4_dg.get().strip() else tu_dong_chia_giuong(so_g)
    cur.execute("INSERT INTO PhongThuThuat (ten_phong, bac_si, ktv, danh_sach_may, so_giuong, danh_sach_giuong) VALUES (?,?,?,?,?,?)", (e4_t.get(), bs, kt, may, so_g, ds_g)); conn.commit(); conn.close(); tai_ds4(); clear_4(); e4_t.focus()
def edit_4():
    if sel_4:
        conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
        bs = ", ".join([v.get() for v in v4bs.values() if v.get() not in ["off", "0", ""]]); kt = ", ".join([v.get() for v in v4kt.values() if v.get() not in ["off", "0", ""]])
        may = tu_dong_chia_may(sel_4)
        so_g = int(e4_g.get()) if e4_g.get().isdigit() else 15
        ds_g = e4_dg.get().strip() if e4_dg.get().strip() else tu_dong_chia_giuong(so_g, sel_4)
        cur.execute("UPDATE PhongThuThuat SET ten_phong=?, bac_si=?, ktv=?, danh_sach_may=?, so_giuong=?, danh_sach_giuong=? WHERE id=?", (e4_t.get(), bs, kt, may, so_g, ds_g, sel_4)); conn.commit(); conn.close(); tai_ds4(); clear_4(); e4_t.focus()
def del_4(*args):
    if sel_4: conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("DELETE FROM PhongThuThuat WHERE id=?", (sel_4,)); conn.commit(); conn.close(); tai_ds4(); clear_4()
tree4.bind("<ButtonRelease-1>", click_4); tree4.bind("<Delete>", del_4); tree4.bind("<BackSpace>", del_4); tree4.bind("<Return>", lambda e: edit_4() if sel_4 else add_4()); gan_phim_enter([e4_t, e4_g, e4_dg], lambda e: edit_4() if sel_4 else add_4())
tao_btn(f4_bot, add_4, edit_4, del_4, clear_4, ["ten_phong", "bac_si", "ktv", "danh_sach_may", "so_giuong", "danh_sach_giuong"], "Phong", "PhongThuThuat", tree4, tai_ds4, "DS_Phong")

# --- TAB 5 ---
t5 = tabview.tab("5. Bệnh nhân")
f5l = ctk.CTkFrame(t5, width=350); f5l.pack(side="left", fill="y", padx=10, pady=10); f5l.pack_propagate(False)
f5_bot = ctk.CTkFrame(f5l, fg_color="transparent"); f5_bot.pack(side="bottom", fill="x", pady=5)
f5_top = ctk.CTkFrame(f5l, fg_color="transparent"); f5_top.pack(side="top", fill="x")

e5_t = make_row(f5_top, "Tên BN:", ctk.CTkEntry, placeholder_text="Tên bệnh nhân")
e5_n = make_row(f5_top, "Năm sinh:", ctk.CTkEntry, placeholder_text="Năm sinh")
cb5_loai = make_row(f5_top, "Phân loại:", ctk.CTkComboBox, values=["Nội trú", "Nội trú ban ngày"]); cb5_loai.set("Nội trú")
e5_ngay = make_row(f5_top, "Ngày vào:", ctk.CTkEntry, placeholder_text="DD/MM/YYYY")
e5_ngay.insert(0, datetime.now().strftime("%d/%m/%Y")); e5_ngay.bind("<KeyRelease>", lambda ev: auto_format_date(ev, e5_ngay))
e5_v = make_row(f5_top, "Giờ vào:", ctk.CTkEntry, placeholder_text="VD: 07:59")
e5_v.bind("<KeyRelease>", lambda ev: auto_format_time(ev, e5_v))
e5_ban = make_row(f5_top, "Giờ bận:", ctk.CTkEntry, placeholder_text="VD: 08:00-09:00, 10:00-11:00")
e5_r = make_row(f5_top, "Giờ ra viện:", ctk.CTkEntry, placeholder_text="Cập nhật 2 chiều", text_color="#d35400")
e5_r.bind("<KeyRelease>", lambda ev: auto_format_time(ev, e5_r))
cb5_p = make_row(f5_top, "Phòng:", ctk.CTkComboBox)

fr5_tt = ctk.CTkScrollableFrame(f5l, label_text="Thủ thuật"); fr5_tt.pack(side="top", fill="both", expand=True, pady=5); v5tt = {}
f5r = ctk.CTkFrame(t5, fg_color="transparent"); f5r.pack(side="right", fill="both", expand=True, padx=10, pady=10)
f5_search = ctk.CTkFrame(f5r, fg_color="transparent"); f5_search.pack(fill="x", pady=(0, 5))
ctk.CTkLabel(f5_search, text="🔍 Tìm kiếm:", font=("Arial", 12, "bold")).pack(side="left", padx=5); e5_search = ctk.CTkEntry(f5_search, placeholder_text="Tìm theo bất kỳ (Tên, Năm, Phòng, Tên thủ thuật...)", width=350); e5_search.pack(side="left", padx=5)

tree5 = create_tree(f5r, ("STT", "Tên BN", "Năm Sinh", "Loại BN", "Ngày Vào", "Giờ Vào", "Giờ Bận", "Giờ Ra", "Phòng", "Thủ Thuật"), [30, 150, 60, 90, 80, 60, 120, 60, 80, 250])

def clear_5():
    global sel_5; sel_5 = None; e5_t.delete(0, 'end'); e5_n.delete(0, 'end'); e5_v.delete(0, 'end'); e5_ban.delete(0, 'end'); e5_r.delete(0, 'end')
    e5_ngay.delete(0, 'end'); e5_ngay.insert(0, datetime.now().strftime("%d/%m/%Y")); cb5_loai.set("Nội trú")
    for v in v5tt.values(): v.set("off")
def tai_ds5():
    tukhoa = e5_search.get().lower(); stt = 1
    for r in tree5.get_children(): tree5.delete(r)
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    for row in cur.execute("SELECT id, ten, nam_sinh, loai_bn, ngay_vao, gio_vao, gio_ban, gio_ra, phong, thu_thuat FROM BenhNhan"):
        chuoi_check = f"{row[1]} {row[2]} {row[3]} {row[4]} {row[5]} {row[6]} {row[7]} {row[8]} {row[9]}".lower()
        if tukhoa in chuoi_check: tree5.insert("", "end", iid=row[0], values=(stt, *row[1:])); stt += 1
    phs = [r[0] for r in cur.execute("SELECT ten_phong FROM PhongThuThuat")]
    if phs: cb5_p.configure(values=["Tự động chọn phòng"] + phs); cb5_p.set("Tự động chọn phòng")
    for w in fr5_tt.winfo_children(): w.destroy()
    v5tt.clear()
    for r in cur.execute("SELECT ten FROM ThuThuat"):
        v = ctk.StringVar(value="off"); v5tt[r[0]] = v; cb = ctk.CTkCheckBox(fr5_tt, text=r[0], variable=v, onvalue=r[0], offvalue="off"); cb.pack(anchor="w", padx=10, pady=2)
        cb.bind("<Return>", lambda e: edit_5() if sel_5 else add_5()); cb.bind("<FocusIn>", lambda ev, sf=fr5_tt: auto_scroll(ev, sf))
    conn.close(); tai_cb_t6()
e5_search.bind("<KeyRelease>", lambda e: tai_ds5())
def click_5(e):
    global sel_5; sel = tree5.selection(); 
    if sel:
        clear_5(); sel_5 = sel[0]; i = tree5.item(sel_5, 'values')
        e5_t.delete(0, 'end'); e5_t.insert(0, i[1]); e5_n.delete(0, 'end'); e5_n.insert(0, i[2]); cb5_loai.set(i[3])
        e5_ngay.delete(0, 'end'); e5_ngay.insert(0, i[4] if i[4] and i[4] != "None" else "")
        e5_v.delete(0, 'end'); e5_v.insert(0, i[5]); e5_ban.delete(0, 'end'); e5_ban.insert(0, i[6] if i[6] and i[6] != "None" else "")
        e5_r.delete(0, 'end'); e5_r.insert(0, i[7] if i[7] and i[7] != "None" else "")
        cb5_p.set(i[8])
        for tt in str(i[9]).split(", "):
            if tt in v5tt: v5tt[tt].set(tt)
def get_b5(): return sort_busy_slots(e5_ban.get())
def get_auto_room():
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    vr = [r[0] for r in cur.execute("SELECT ten_phong FROM PhongThuThuat").fetchall() if "sóng ngắn" not in str(r[0]).lower() and "cấy chỉ" not in str(r[0]).lower()]
    if not vr: vr = ["Phòng 1"]
    rc = {rm: 0 for rm in vr}
    for r in cur.execute("SELECT phong FROM BenhNhan WHERE gio_ra = '' OR gio_ra IS NULL").fetchall():
        if r[0] in rc: rc[r[0]] += 1
    conn.close(); return min(rc, key=rc.get) if rc else "Phòng 1"
def add_5():
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    tt = ", ".join([v.get() for v in v5tt.values() if v.get() not in ["off", "0", ""]]); gv = e5_v.get() if e5_v.get() else "07:59"
    p = cb5_p.get() if cb5_p.get() != "Tự động chọn phòng" else get_auto_room()
    cur.execute("INSERT INTO BenhNhan (ten, hsba, nam_sinh, ngay_vao, gio_vao, gio_ban, phong, bac_si, thu_thuat, gio_ra, loai_bn) VALUES (?,?,?,?,?,?,?,?,?,?,?)", (e5_t.get(), "", e5_n.get(), e5_ngay.get(), gv, get_b5(), p, "", tt, e5_r.get(), cb5_loai.get()))
    conn.commit(); conn.close(); e5_search.delete(0, 'end'); tai_ds5(); clear_5(); e5_t.focus() 
def edit_5():
    if sel_5:
        conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
        tt = ", ".join([v.get() for v in v5tt.values() if v.get() not in ["off", "0", ""]]); gv = e5_v.get() if e5_v.get() else "07:59"
        p = cb5_p.get() if cb5_p.get() != "Tự động chọn phòng" else get_auto_room()
        cur.execute("UPDATE BenhNhan SET ten=?, nam_sinh=?, ngay_vao=?, gio_vao=?, gio_ban=?, phong=?, thu_thuat=?, gio_ra=?, loai_bn=? WHERE id=?", (e5_t.get(), e5_n.get(), e5_ngay.get(), gv, get_b5(), p, tt, e5_r.get(), cb5_loai.get(), sel_5))
        conn.commit(); conn.close(); e5_search.delete(0, 'end'); tai_ds5(); clear_5(); e5_t.focus()
def del_5(*args):
    if sel_5: conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("DELETE FROM BenhNhan WHERE id=?", (sel_5,)); conn.commit(); conn.close(); tai_ds5(); clear_5()
tree5.bind("<ButtonRelease-1>", click_5); tree5.bind("<Delete>", del_5); tree5.bind("<BackSpace>", del_5); tree5.bind("<Return>", lambda e: edit_5() if sel_5 else add_5())
gan_phim_enter([e5_t, e5_n, e5_ngay, e5_v, e5_ban, e5_r], lambda e: edit_5() if sel_5 else add_5())
tao_btn(f5_bot, add_5, edit_5, del_5, clear_5, ["ten", "nam_sinh", "loai_bn", "ngay_vao", "gio_vao", "gio_ban", "gio_ra", "phong", "thu_thuat"], "BN", "BenhNhan", tree5, tai_ds5, "DS_BN")
# --- TAB 6 ---
t6 = tabview.tab("6. AUTO XẾP LỊCH")
f6t = ctk.CTkFrame(t6, fg_color="transparent"); f6t.pack(fill="x", padx=10, pady=5)
f6_l = ctk.CTkFrame(f6t); f6_l.pack(side="left", fill="both", expand=True, padx=(0,5))
ctk.CTkLabel(f6_l, text="THIẾT LẬP ƯU TIÊN RA VIỆN", font=("Arial", 12, "bold")).pack(pady=(5,0))
f6_in = ctk.CTkFrame(f6_l, fg_color="transparent"); f6_in.pack(pady=2)
f6_search_wrapper = ctk.CTkFrame(f6_in, fg_color="transparent"); f6_search_wrapper.pack(side="left", padx=2)
e6_search_bn = ctk.CTkEntry(f6_search_wrapper, placeholder_text="🔍 Gõ tên (Tìm nhanh)...", width=250); e6_search_bn.pack(side="left")
lb_suggestions = tk.Listbox(t6, font=("Arial", 11), selectbackground="#3498db", selectforeground="white", bd=1, relief="solid")

def update_suggestions(event):
    tukhoa = e6_search_bn.get().lower(); lb_suggestions.delete(0, 'end')
    if not tukhoa: lb_suggestions.place_forget(); return
    kq = [bn for bn in full_bn_list if tukhoa in bn.lower()]
    if kq:
        for bn in kq: lb_suggestions.insert('end', bn)
        lb_suggestions.configure(height=min(len(kq), 5)); lb_suggestions.place(x=e6_search_bn.winfo_rootx()-t6.winfo_rootx(), y=e6_search_bn.winfo_rooty()-t6.winfo_rooty()+e6_search_bn.winfo_height(), width=e6_search_bn.winfo_width()); lb_suggestions.lift()
    else: lb_suggestions.place_forget()
def on_suggestion_select(event=None):
    if lb_suggestions.curselection(): e6_search_bn.delete(0, 'end'); e6_search_bn.insert(0, lb_suggestions.get(lb_suggestions.curselection())); lb_suggestions.place_forget(); e6_r.focus()
e6_search_bn.bind("<KeyRelease>", update_suggestions); lb_suggestions.bind("<ButtonRelease-1>", on_suggestion_select); lb_suggestions.bind("<Return>", on_suggestion_select)          
e6_r = ctk.CTkEntry(f6_in, placeholder_text="Mặc định 14:00", width=120); e6_r.pack(side="left", padx=2)
def set_ra(event=None):
    try:
        chuoi = e6_search_bn.get()
        if " - " in chuoi:
            bid = chuoi.split(" - ")[0]; gio_ra = e6_r.get().strip() if e6_r.get().strip() else "14:00"
            conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("UPDATE BenhNhan SET gio_ra=? WHERE id=?", (gio_ra, bid)); conn.commit(); conn.close(); tai_ds5(); e6_search_bn.delete(0, 'end'); e6_r.delete(0, 'end'); lb_suggestions.place_forget()
    except: pass
e6_r.bind("<KeyRelease>", lambda ev: auto_format_time(ev, e6_r)); e6_r.bind("<Return>", set_ra); ctk.CTkButton(f6_in, text="Ghim/Sửa", width=70, command=set_ra).pack(side="left", padx=2)

ctk.CTkLabel(f6_l, text="BÁO BẬN NHÂN SỰ ĐỘT XUẤT", font=("Arial", 12, "bold"), text_color="#d35400").pack(pady=(10,0))
f6_ns = ctk.CTkFrame(f6_l, fg_color="transparent"); f6_ns.pack(pady=2)
cb6_ns = ctk.CTkComboBox(f6_ns, width=130); cb6_ns.pack(side="left", padx=2)
e6_ns_b1 = ctk.CTkEntry(f6_ns, width=55, placeholder_text="Từ"); e6_ns_b1.pack(side="left", padx=2); e6_ns_b2 = ctk.CTkEntry(f6_ns, width=55, placeholder_text="Đến"); e6_ns_b2.pack(side="left", padx=2)
for e in [e6_ns_b1, e6_ns_b2]: e.bind("<KeyRelease>", lambda ev, w=e: auto_format_time(ev, w))
def check_conflict_and_warn(staff_name, b1, b2):
    global g_sched
    if not g_sched: return True 
    busy_start = t2m(b1); busy_end = t2m(b2); conflicts = []
    for row in g_sched:
        staff_str = str(row.get('NV CHÍNH', '')) + " " + str(row.get('NV PHỤ', ''))
        if staff_name in staff_str:
            if is_overlap(busy_start, busy_end, t2m(row['GIODIENRA']), t2m(row['GIOKETTHUC'])): conflicts.append(f"• {row['GIODIENRA']} - {row['GIOKETTHUC']}: {row['HOTEN']} ({row['DICHVU']})")
    if conflicts:
        msg = f"⚠️ CẢNH BÁO TRÙNG LỊCH!\n\nNhân sự '{staff_name}' đang bị vướng {len(conflicts)} ca đã xếp trong khung giờ {b1}-{b2}:\n" + "\n".join(conflicts[:10]) + (f"\n... và {len(conflicts)-10} ca khác." if len(conflicts)>10 else "") + "\n\nBạn có muốn LƯU giờ bận này không?\n(Sau khi LƯU, hãy bấm 'CHẠY XẾP LỊCH TỔNG' lại để AI điều phối lại!)"
        return messagebox.askyesno("Phát hiện trùng lịch", msg)
    return True
def add_ns_ban():
    ns = cb6_ns.get(); b1 = e6_ns_b1.get(); b2 = e6_ns_b2.get(); ngay = e_ngay_xep.get()
    if ns and b1 and b2 and check_conflict_and_warn(ns, b1, b2):
        chuoi = f"({ngay}) {b1}-{b2}"; conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("SELECT gio_ban FROM NhanSu WHERE ten=?", (ns,)); r = cur.fetchone()
        if r: cur.execute("UPDATE NhanSu SET gio_ban=? WHERE ten=?", (sort_busy_slots(r[0] + ", " + chuoi if r[0] else chuoi), ns)); conn.commit()
        conn.close(); tai_ds3(); tai_ds_pri(); e6_ns_b1.delete(0, 'end'); e6_ns_b2.delete(0, 'end'); e6_ns_b1.focus()
def edit_ns_ban():
    global sel_ns_slot; ns = cb6_ns.get(); b1 = e6_ns_b1.get(); b2 = e6_ns_b2.get(); ngay = e_ngay_xep.get()
    if ns and b1 and b2 and sel_ns_slot and check_conflict_and_warn(ns, b1, b2):
        chuoi_moi = f"({ngay}) {b1}-{b2}"; conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("SELECT gio_ban FROM NhanSu WHERE ten=?", (ns,)); r = cur.fetchone()
        if r and r[0]: cur.execute("UPDATE NhanSu SET gio_ban=? WHERE ten=?", (sort_busy_slots(r[0].replace(sel_ns_slot, chuoi_moi)), ns)); conn.commit()
        conn.close(); tai_ds3(); tai_ds_pri(); e6_ns_b1.delete(0, 'end'); e6_ns_b2.delete(0, 'end'); sel_ns_slot = None; e6_ns_b1.focus()
def del_1_ns_ban():
    global sel_ns_slot; ns = cb6_ns.get()
    if ns and sel_ns_slot:
        conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("SELECT gio_ban FROM NhanSu WHERE ten=?", (ns,)); r = cur.fetchone()
        if r and r[0]: cur.execute("UPDATE NhanSu SET gio_ban=? WHERE ten=?", (sort_busy_slots(r[0].replace(sel_ns_slot, "").replace(", ,", ",").strip(", ")), ns)); conn.commit()
        conn.close(); tai_ds3(); tai_ds_pri(); e6_ns_b1.delete(0, 'end'); e6_ns_b2.delete(0, 'end'); sel_ns_slot = None
e6_ns_b2.bind("<Return>", lambda e: edit_ns_ban() if sel_ns_slot else add_ns_ban())
ctk.CTkButton(f6_ns, text="Lưu", width=45, fg_color="#d35400", command=lambda: edit_ns_ban() if sel_ns_slot else add_ns_ban()).pack(side="left", padx=2)
ctk.CTkButton(f6_ns, text="Xóa", width=45, fg_color="#e74c3c", command=del_1_ns_ban).pack(side="left", padx=2)

f6_r = ctk.CTkFrame(f6t); f6_r.pack(side="right", fill="both", expand=True, padx=(5,0))
f6_r_top = ctk.CTkFrame(f6_r, fg_color="transparent"); f6_r_top.pack(fill="x", expand=True)
tree_pri = ttk.Treeview(f6_r_top, columns=("Mã", "Tên BN", "Giờ Ra"), show="headings", height=8)
tree_pri.heading("Mã", text="Mã"); tree_pri.column("Mã", width=40, anchor="center", stretch=False)
tree_pri.heading("Tên BN", text="Danh sách Ưu tiên ra viện"); tree_pri.column("Tên BN", width=250, stretch=False)
tree_pri.heading("Giờ Ra", text="Giờ Ra"); tree_pri.column("Giờ Ra", width=80, anchor="center", stretch=False)
tree_pri.pack(side="left", fill="both", expand=True, padx=2, pady=2)
def click_pri(e):
    sel = tree_pri.selection()
    if sel:
        item = tree_pri.item(sel[0], 'values'); conn = sqlite3.connect(DB_NAME); cur = conn.cursor(); cur.execute("SELECT id, ten, nam_sinh FROM BenhNhan WHERE id=?", (item[0],)); r = cur.fetchone()
        if r: e6_search_bn.delete(0, 'end'); e6_search_bn.insert(0, f"{r[0]} - {r[1]} ({r[2]})")
        conn.close(); e6_r.delete(0, 'end'); e6_r.insert(0, item[2])
tree_pri.bind("<ButtonRelease-1>", click_pri); tree_pri.bind("<Delete>", lambda e: del_ra())
tree_ns_ban = ttk.Treeview(f6_r_top, columns=("STT", "Tên NV", "Giờ Bận"), show="headings", height=8)
tree_ns_ban.heading("STT", text="STT"); tree_ns_ban.column("STT", width=40, anchor="center", stretch=False)
tree_ns_ban.heading("Tên NV", text="Nhân sự bận"); tree_ns_ban.column("Tên NV", width=160, stretch=False)
tree_ns_ban.heading("Giờ Bận", text="Khung giờ bận"); tree_ns_ban.column("Giờ Bận", width=220, anchor="center", stretch=False)
tree_ns_ban.pack(side="right", fill="both", expand=True, padx=2, pady=2)
def click_ns_ban(e):
    global sel_ns_slot; sel = tree_ns_ban.selection()
    if sel:
        item = tree_ns_ban.item(sel[0], 'values'); cb6_ns.set(item[1]); sel_ns_slot = item[2] 
        time_part = sel_ns_slot.split(")")[-1].strip() if ")" in sel_ns_slot else sel_ns_slot; ban = time_part.split("-")
        if len(ban) == 2: e6_ns_b1.delete(0, 'end'); e6_ns_b1.insert(0, ban[0]); e6_ns_b2.delete(0, 'end'); e6_ns_b2.insert(0, ban[1])
tree_ns_ban.bind("<ButtonRelease-1>", click_ns_ban); tree_ns_ban.bind("<Delete>", lambda e: del_1_ns_ban()) 

f6_bot = ctk.CTkFrame(t6); f6_bot.pack(fill="both", expand=True, padx=10, pady=5)
def tim_gio_chi_dinh():
    global g_tl, g_ca
    for r in tree_cd.get_children(): tree_cd.delete(r) 
    if not g_tl: return messagebox.showwarning("Cảnh báo", "Vui lòng bấm 'Chạy xếp lịch tổng' trước!")
    vao_str = e6_v_cd.get()
    if not vao_str: return messagebox.showwarning("Cảnh báo", "Vui lòng nhập 'Giờ vào viện' (VD: 14:04)!")
    t_vao = t2m(vao_str); conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    docs = [r[0] for r in cur.execute("SELECT ten FROM NhanSu WHERE vai_tro='Bác sĩ' AND trang_thai != 'Nghỉ cả ngày'")]
    conn.close(); found = False
    for doc in docs:
        if doc not in g_tl: continue
        ca = g_ca.get(doc, []); timeline = g_tl.get(doc, [])
        busy = timeline + [(720, 780)]; busy = sorted(busy, key=lambda x: x[0]); merged_busy = []
        for b in busy:
            if not merged_busy: merged_busy.append(b)
            else:
                last = merged_busy[-1]
                if b[0] <= last[1]: merged_busy[-1] = (last[0], max(last[1], b[1]))
                else: merged_busy.append(b)
        free_blocks = []
        for s_start, s_end in ca:
            curr = s_start
            for b_start, b_end in merged_busy:
                if b_start >= s_end: break
                if curr < b_start: free_blocks.append((curr, b_start))
                curr = max(curr, b_end)
            if curr < s_end: free_blocks.append((curr, s_end))
        for g_start, g_end in free_blocks:
            valid_start = max(g_start, t_vao + 1)
            if valid_start < g_end:
                mins = g_end - valid_start; tree_cd.insert("", "end", values=(f"👨‍⚕️ {doc.upper()}", f"{m2t(valid_start)} - {m2t(g_end)}", f"{mins}")); found = True; break 
    if not found: tree_cd.insert("", "end", values=("Không có Bác sĩ rảnh", "-", "-"))

f6_cd_frame = ctk.CTkFrame(f6_bot, fg_color="transparent"); f6_cd_frame.pack(fill="x", pady=(2, 0))
f6_cd_left = ctk.CTkFrame(f6_cd_frame, fg_color="transparent"); f6_cd_left.pack(side="left", padx=5)
ctk.CTkLabel(f6_cd_left, text="⏱ TÌM GIỜ CHỈ ĐỊNH:", font=("Arial", 12, "bold"), text_color="#27ae60").pack(side="left", padx=5)
e6_v_cd = ctk.CTkEntry(f6_cd_left, placeholder_text="Giờ vào (VD: 14:04)", width=150); e6_v_cd.pack(side="left", padx=5)
e6_v_cd.bind("<KeyRelease>", lambda ev: auto_format_time(ev, e6_v_cd)); e6_v_cd.bind("<Return>", lambda ev: tim_gio_chi_dinh())
ctk.CTkButton(f6_cd_left, text="Tìm Bác Sĩ Rảnh", width=120, fg_color="#27ae60", command=tim_gio_chi_dinh).pack(side="left", padx=5)
tree_cd = ttk.Treeview(f6_cd_frame, columns=("Bác sĩ", "Khung giờ", "Thời lượng"), show="headings", height=7)
tree_cd.heading("Bác sĩ", text="Bác sĩ rảnh"); tree_cd.column("Bác sĩ", width=160, anchor="w", stretch=False)
tree_cd.heading("Khung giờ", text="Khung giờ trống"); tree_cd.column("Khung giờ", width=120, anchor="center", stretch=False)
tree_cd.heading("Thời lượng", text="Tổng (Phút)"); tree_cd.column("Thời lượng", width=80, anchor="center", stretch=False)
tree_cd.pack(side="left", fill="both", expand=True, padx=20)
f6_filter = ctk.CTkFrame(f6_bot, fg_color="transparent"); f6_filter.pack(fill="x", pady=2)
ctk.CTkLabel(f6_filter, text="🔍 Tìm kiếm trong lịch:", font=("Arial", 12, "bold")).pack(side="left", padx=5)
e6_search_lich = ctk.CTkEntry(f6_filter, placeholder_text="Gõ tên BN, năm sinh, KTV...", width=300); e6_search_lich.pack(side="left", padx=5)
tree6 = ttk.Treeview(f6_bot, columns=("STT", "NGAY", "HOTEN", "NAMSINH", "PHONG", "DICHVU", "GIODIENRA", "GIOKETTHUC", "NV CHÍNH", "NV PHỤ", "MAY", "GIUONG"), show="headings")
for col, w in zip(tree6["columns"], [40, 90, 160, 60, 100, 130, 80, 80, 120, 120, 90, 90]):
    tree6.heading(col, text=col, command=lambda c=col: tree_sort(tree6, c, False))
    tree6.column(col, width=w, anchor="center" if col not in ["HOTEN", "DICHVU", "NV CHÍNH", "NV PHỤ"] else "w", stretch=False)
tree6.pack(fill="both", expand=True)

def hien_thi_lich(event=None):
    global g_sched
    tukhoa = e6_search_lich.get().lower(); stt = 1
    for r in tree6.get_children(): tree6.delete(r)
    for row in g_sched:
        chuoi_check = f"{row.get('NGAY','')} {row.get('HOTEN','')} {row.get('NAMSINH','')} {row.get('PHONG','')} {row.get('DICHVU','')} {row.get('GIODIENRA','')} {row.get('GIOKETTHUC','')} {row.get('NV CHÍNH','')} {row.get('NV PHỤ','')} {row.get('MAY','')} {row.get('GIUONG','')}".lower()
        if tukhoa in chuoi_check:
            tree6.insert("", "end", values=(stt, row['NGAY'], row['HOTEN'], row['NAMSINH'], row['PHONG'], row['DICHVU'], row['GIODIENRA'], row['GIOKETTHUC'], row.get('NV CHÍNH',''), row.get('NV PHỤ',''), row['MAY'], row['GIUONG'])); stt += 1
e6_search_lich.bind("<KeyRelease>", hien_thi_lich)

f6_mid = ctk.CTkFrame(t6, fg_color="transparent"); f6_mid.pack(pady=5)
ctk.CTkLabel(f6_mid, text="Ngày xếp lịch:").pack(side="left", padx=5)
e_ngay_xep = ctk.CTkEntry(f6_mid, width=100); e_ngay_xep.insert(0, datetime.now().strftime("%d/%m/%Y")); e_ngay_xep.pack(side="left", padx=5)

def run_auto():
    global g_staff, g_proc, g_req, g_rot, g_sched, g_tl, g_ca
    try:
        for r in tree6.get_children(): tree6.delete(r)
        conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
        m_types = {}; m_timeline = {"Thủ công": []} 
        for r in cur.execute("SELECT ten_loai, ma_may FROM DanhSachMay WHERE trang_thai='Sẵn sàng'"):
            if r[0] not in m_types: m_types[r[0]] = []
            m_types[r[0]].append(r[1]); m_timeline[r[1]] = []
        replacement_map = {r[0]: r[1] for r in cur.execute("SELECT ten, nguoi_thay_the FROM NhanSu") if r[1] and r[1] != "Không"}
        
        g_staff.clear(); g_proc.clear(); tt_i = {}; g_rot.clear(); results = []; g_req.clear(); s_timeline = {}; s_k = {}; s_ca = {}; s_ban = {}; s_role = {}
        for r in cur.execute("SELECT ten, vai_tro, ky_nang, thoi_gian_lam, gio_ban, trang_thai FROM NhanSu"):
            s = r[0]; status = r[5]
            if status == 'Nghỉ cả ngày': continue 
            s_timeline[s] = []; s_role[s] = r[1]; s_k[s] = [x.strip() for x in str(r[2]).split(", ")] if r[2] else []
            raw_shifts = [(t2m(p.split("-")[0]), t2m(p.split("-")[1])) for p in str(r[3]).split(", ") if "-" in p] if r[3] else [(480, 720), (780, 990)]
            if status == 'Nghỉ sáng': s_ca[s] = [(cs, ce) for cs, ce in raw_shifts if cs >= 720]
            elif status == 'Nghỉ chiều': s_ca[s] = [(cs, ce) for cs, ce in raw_shifts if ce <= 780]
            else: s_ca[s] = raw_shifts
            if status in ['Nghỉ sáng', 'Nghỉ chiều'] and not s_ca[s]: s_ca[s] = [(780, 990)] if status == 'Nghỉ sáng' else [(480, 720)]
            s_ban[s] = []
            if r[4]:
                for p in str(r[4]).split(","):
                    p_clean = p.strip(); time_part = p_clean.split(")")[-1].strip() if ")" in p_clean else p_clean
                    if "-" in time_part: s_ban[s].append((t2m(time_part.split("-")[0]), t2m(time_part.split("-")[1]))); s_timeline[s].append((t2m(time_part.split("-")[0]), t2m(time_part.split("-")[1])))
            g_staff[s] = {'shift_mins': sum([ce - cs for cs, ce in s_ca[s]]) if s_ca[s] else 540, 'busy_mins': sum([be - bs for bs, be in s_ban[s]]), 'used_mins': 0, 'procs_done': {}, 'skills': s_k[s]}
            
        for r in cur.execute("SELECT ten, loai_may, thoi_gian_may, thoi_gian_nguoi, loai_chuyen_mon, can_rut_may, can_nguoi_phu, ds_nguoi_phu FROM ThuThuat"):
            dm = int(r[2]) if r[2] else 15; dn = int(r[3]) if r[3] else 5; ds_phu = [x.strip() for x in str(r[7]).split(',')] if r[7] else []
            tt_i[r[0]] = (r[1], max(1, dm), max(1, dn), r[4], int(r[5]) if r[5] else 1, int(r[6]) if r[6] else 0, ds_phu)
            g_proc[r[0]] = 0; g_req[r[0]] = 0
            
        room_staff_raw = {}; room_beds_tracker = {}; room_beds_count = {}
        for r in cur.execute("SELECT ten_phong, bac_si, ktv, danh_sach_may, so_giuong, danh_sach_giuong FROM PhongThuThuat"):
            room_name = r[0]; num_beds = int(r[4]) if r[4] else 15; room_beds_count[room_name] = num_beds
            bed_str = str(r[5]).strip() if r[5] else ""
            bed_list = [x.strip() for x in bed_str.split(",") if x.strip()] if bed_str and bed_str != 'None' else [f"Giường {i}" for i in range(1, num_beds + 1)]
            room_beds_tracker[room_name] = {b: [] for b in bed_list}
            bs_list = [x.strip() for x in str(r[1]).split(",")] if r[1] else []; ktv_list = [x.strip() for x in str(r[2]).split(",")] if r[2] else []
            exp_bs = [replacement_map.get(bs, bs) for bs in bs_list]
            exp_ktv = [replacement_map.get(ktv, ktv) for ktv in ktv_list]
            room_staff_raw[room_name] = {'bs': [x for x in exp_bs if x], 'ktv': [x for x in exp_ktv if x]}
            
        room_staff = {room: list(set(data['bs'] + data['ktv'])) for room, data in room_staff_raw.items()}
        all_assigned_staff = set(); [all_assigned_staff.update(sl) for sl in room_staff.values()]
        floaters = {s for s in s_k.keys() if s not in all_assigned_staff}
        
        patients = []; raw_patients = []
        for r in cur.execute("SELECT id, ten, nam_sinh, gio_vao, gio_ban, gio_ra, phong, bac_si, thu_thuat, loai_bn FROM BenhNhan ORDER BY phong ASC, id ASC"):
            if r[8]: raw_patients.append(r)
            
        room_load = {r: 0 for r in room_beds_count.keys() if "sóng ngắn" not in r.lower() and "cấy chỉ" not in r.lower()}
        if not room_load: room_load = {"Phòng 1": 0}
        for r in raw_patients:
            pr = r[6]
            if pr and pr != "Tự động chọn phòng" and pr in room_load: room_load[pr] += 1
                
        for r in raw_patients:
            pending_list = [x.strip() for x in r[8].split(",") if x.strip()]
            for ttn in pending_list: g_req[ttn] = g_req.get(ttn, 0) + 1
            
            pending_list.sort(key=lambda ttn: (
                1 if "cấy chỉ" in ttn.lower() else 0,                                 
                0 if tt_i.get(ttn, ("", 0, 0, "PHCN", 1, 0, []))[3] == "YHCT" else 1,  
                -tt_i.get(ttn, ("", 0, 999, "PHCN", 1, 0, []))[2]                       
            ))
            
            arrive_time = t2m(r[3]) if r[3] else 479; arrive_time = 479 if arrive_time == 0 else arrive_time
            free_time = 780 if 660 <= arrive_time < 780 else arrive_time
            patient_room = r[6] if r[6] else "Tự động chọn phòng"
            if patient_room == "Tự động chọn phòng": patient_room = min(room_load, key=room_load.get) if room_load else "Phòng 1"; room_load[patient_room] += 1
            patients.append({'name': f"{r[1].upper()}", 'nam_sinh': r[2], 'loai_bn': r[9] if r[9] else 'Nội trú ban ngày', 'primary_room': patient_room, 'arrive': arrive_time, 'leave': t2m(r[5]) if r[5] else 9999, 'busy': [(t2m(p.split("-")[0]), t2m(p.split("-")[1])) for p in str(r[4]).split(",") if "-" in p] if r[4] else [], 'pending': pending_list, 'session_locked': None, 'total_procs': len(pending_list), 'procs_done_morning': 0, 'free_at': free_time, 'failed': False, 'skip_phase1': False, 'rand_seed': random.random()})
            
        ngay_xep_lich = e_ngay_xep.get()
        for phase_idx in [1, 2]:
            for t_now in range(420, 1200):
                if 720 <= t_now < 780: continue 
                el_pats = [p for p in patients if not p['failed'] and p['pending'] and p['free_at'] <= t_now]
                el_pats.sort(key=lambda x: (x['leave'], x['rand_seed'])) 
                for p in el_pats:
                    if phase_idx == 1 and p['skip_phase1']: continue
                    in_cls = False
                    for bs, be in p['busy']:
                        if bs <= t_now < be: p['free_at'] = be; in_cls = True; break
                    if in_cls: continue
                    scheduled_something = False
                    for i, ttn in enumerate(p['pending']):
                        if ttn not in tt_i: p['pending'].pop(i); scheduled_something = True; break
                        lm, dm, dn, loai_cm, can_rut, can_phu, ds_phu = tt_i[ttn]
                        if lm == "Thủ công": dn = dm
                        if t_now < 720 and (t_now + dm) > 720: continue 
                        time_limit = 9999 if p['loai_bn'] == 'Nội trú ban ngày' else 540
                        is_time_breached = (t_now + dm) > (p['arrive'] + time_limit)
                        is_leave_breached = (p['leave'] != 9999 and t_now + dm > p['leave'])
                        if is_time_breached or is_leave_breached:
                            p['skip_phase1'] = True
                            if phase_idx == 2:
                                g_rot.append({'bn': p['name'], 'ns': p['nam_sinh'], 'tt': ttn, 'room': p['primary_room'], 'staff': ", ".join(room_staff.get(p['primary_room'], [])), 'reason': f"Vượt giới hạn ({m2t(p['leave'])})" if is_leave_breached else "Vượt giới hạn sinh học"})
                                p['pending'].pop(i) 
                            scheduled_something = True; break 
                        
                        setup_start = t_now; setup_end = t_now + dn + 1; m_end = t_now + dm + 1
                        
                        # FIX LỖI NONETYPE: Chỉ khởi tạo tear_start nếu có yêu cầu rút máy
                        if dm > dn and can_rut == 1: tear_start = t_now + dm - 1; tear_end = t_now + dm + 1
                        else: tear_start = None; tear_end = None
                        
                        possible_machines = ["Thủ công"] if lm == "Thủ công" else m_types.get(lm, [])
                        if not possible_machines: continue
                        target_room = p['primary_room'] 
                        for m in possible_machines:
                            if m != "Thủ công":
                                overlap_m = False
                                for ms, me in m_timeline[m]:
                                    if is_overlap(t_now, m_end, ms, me): overlap_m = True; break
                                if overlap_m: continue 
                            assigned_bed = None
                            if target_room in room_beds_tracker:
                                for b_id, b_timeline in room_beds_tracker[target_room].items():
                                    overlap_b = False
                                    for bs, be in b_timeline:
                                        if is_overlap(t_now, m_end, bs, be): overlap_b = True; break
                                    if not overlap_b: assigned_bed = b_id; break
                            if not assigned_bed and target_room in room_beds_tracker and len(room_beds_tracker[target_room]) > 0: continue 
                            
                            main_cands = []; sub_cands = []; mon_cands = []
                            for s in s_k.keys():
                                if ttn not in s_k.get(s, []): continue
                                is_loc = (s in room_staff.get(p['primary_room'], [])); is_flt = (s in floaters)
                                
                                if not (is_loc or is_flt): continue 
                                
                                shift_ok = False
                                if not s_ca[s]: shift_ok = True
                                else:
                                    for cs, ce in s_ca[s]:
                                        if setup_start >= cs and m_end <= ce: shift_ok = True; break
                                if not shift_ok: continue
                                r_role = s_role.get(s, '')
                                f_set = True
                                for bs, be in s_timeline[s]:
                                    if is_overlap(setup_start, setup_end, bs, be): f_set = False; break
                                
                                if f_set and r_role != 'Điều dưỡng':
                                    if loai_cm == 'YHCT' and r_role == 'Bác sĩ': main_cands.append(s)
                                    elif loai_cm == 'PHCN':
                                        if r_role == 'Kỹ thuật viên': main_cands.append(s)
                                        elif r_role == 'Bác sĩ' and phase_idx == 2: main_cands.append(s)

                                if not ds_phu or s in ds_phu:
                                    if can_phu == 1 and f_set and r_role in ['Kỹ thuật viên', 'Điều dưỡng']: sub_cands.append(s)
                                    # CHỐT CHẶN NONETYPE TẠI ĐÂY
                                    if can_rut == 1 and r_role in ['Kỹ thuật viên', 'Điều dưỡng'] and tear_start is not None:
                                        f_tear = True
                                        for bs, be in s_timeline[s]:
                                            if is_overlap(tear_start, tear_end, bs, be): f_tear = False; break
                                        if f_tear:
                                            shift_ok_tear = False
                                            if not s_ca[s]: shift_ok_tear = True
                                            else:
                                                for cs, ce in s_ca[s]:
                                                    if tear_start >= cs and tear_end <= ce: shift_ok_tear = True; break
                                            if shift_ok_tear: mon_cands.append(s)
                                            
                            if not main_cands: continue

                            if can_phu == 1:
                                if not sub_cands: continue 
                                main_cands.sort(key=lambda x: (0 if x in room_staff.get(p['primary_room'], []) else 1, g_staff[x]['used_mins']))
                                fs_main = main_cands[0]
                                valid_subs = [x for x in sub_cands if x != fs_main]
                                if not valid_subs: continue
                                valid_subs.sort(key=lambda x: (0 if x in room_staff.get(p['primary_room'], []) else 1, g_staff[x]['used_mins']))
                                fs_sub = valid_subs[0]
                                fs_mon = None
                                if can_rut == 1 and tear_start is not None:
                                    valid_mons = [x for x in mon_cands if x not in [fs_main, fs_sub]]
                                    if valid_mons: 
                                        valid_mons.sort(key=lambda x: (0 if x in room_staff.get(p['primary_room'], []) else 1, g_staff[x]['used_mins']))
                                        fs_mon = valid_mons[0]
                                
                                bed_str = f"{assigned_bed}" if assigned_bed else ""
                                results.append({'NGAY': ngay_xep_lich, 'HOTEN': p['name'], 'NAMSINH': p['nam_sinh'], 'PHONG': target_room, 'DICHVU': ttn, 'GIODIENRA': m2t(t_now), 'GIOKETTHUC': m2t(t_now+dm), 'NV CHÍNH': fs_main, 'NV PHỤ': fs_sub, 'MAY': m, 'GIUONG': bed_str, 't_sort': t_now})
                                g_proc[ttn] += 1; u_time = setup_end - setup_start
                                g_staff[fs_main]['procs_done'][ttn] = g_staff[fs_main]['procs_done'].get(ttn, 0) + 1; g_staff[fs_main]['used_mins'] += u_time; s_timeline[fs_main].append((setup_start, setup_end))
                                g_staff[fs_sub]['procs_done'][ttn] = g_staff[fs_sub]['procs_done'].get(ttn, 0) + 1; g_staff[fs_sub]['used_mins'] += u_time; s_timeline[fs_sub].append((setup_start, setup_end))
                                if fs_mon: g_staff[fs_mon]['used_mins'] += (tear_end - tear_start); s_timeline[fs_mon].append((tear_start, tear_end))
                                if m != "Thủ công": m_timeline[m].append((t_now, m_end)) 
                                if target_room in room_beds_tracker and assigned_bed: room_beds_tracker[target_room][assigned_bed].append((t_now, m_end))
                                p['free_at'] = m_end; p['pending'].pop(i); scheduled_something = True; break 
                            else:
                                main_cands.sort(key=lambda x: (0 if loai_cm == 'PHCN' and s_role.get(x, '') != 'Bác sĩ' else 1, 0 if x in room_staff.get(p['primary_room'], []) else 1, g_staff[x]['used_mins']))
                                fs_main = main_cands[0]; fs_mon = None
                                if loai_cm == 'YHCT' and can_rut == 1 and tear_start is not None:
                                    valid_mons = [x for x in mon_cands if x != fs_main]
                                    if valid_mons:
                                        valid_mons.sort(key=lambda x: (0 if x in room_staff.get(p['primary_room'], []) else 1, g_staff[x]['used_mins']))
                                        fs_mon = valid_mons[0]
                                
                                bed_str = f"{assigned_bed}" if assigned_bed else ""
                                results.append({'NGAY': ngay_xep_lich, 'HOTEN': p['name'], 'NAMSINH': p['nam_sinh'], 'PHONG': target_room, 'DICHVU': ttn, 'GIODIENRA': m2t(t_now), 'GIOKETTHUC': m2t(t_now+dm), 'NV CHÍNH': fs_main, 'NV PHỤ': fs_mon if fs_mon else "", 'MAY': m, 'GIUONG': bed_str, 't_sort': t_now})
                                g_proc[ttn] += 1; g_staff[fs_main]['procs_done'][ttn] = g_staff[fs_main]['procs_done'].get(ttn, 0) + 1
                                used_time = setup_end - setup_start + (tear_end - tear_start if tear_start else 0)
                                g_staff[fs_main]['used_mins'] += used_time; s_timeline[fs_main].append((setup_start, setup_end)); s_timeline[fs_main].append((tear_start, tear_end)) if tear_start else None
                                if fs_mon: g_staff[fs_mon]['used_mins'] += (tear_end - tear_start); s_timeline[fs_mon].append((tear_start, tear_end))
                                if m != "Thủ công": m_timeline[m].append((t_now, m_end)) 
                                if target_room in room_beds_tracker and assigned_bed: room_beds_tracker[target_room][assigned_bed].append((t_now, m_end))
                                p['free_at'] = m_end; p['pending'].pop(i); scheduled_something = True; break 
                        if scheduled_something: break 
        for p in patients:
            for ttn in p['pending']: g_rot.append({'bn': p['name'], 'ns': p['nam_sinh'], 'tt': ttn, 'room': p['primary_room'], 'staff': "Không có người", 'reason': "Kẹt nhân sự/máy"})
        results.sort(key=lambda x: x['t_sort'])
        g_sched.clear(); g_sched.extend(results); g_tl.clear(); g_tl.update(s_timeline); g_ca.clear(); g_ca.update(s_ca)
        e6_search_lich.delete(0, 'end'); hien_thi_lich(); conn.close(); tai_thong_ke(); luu_trang_thai()
        messagebox.showinfo("Thành công", f"Đã xếp lịch xong\nĐã xếp {len(g_sched)} ca, bị rớt lịch {len(g_rot)} ca")
    except Exception as e:
        messagebox.showerror("Lỗi thuật toán", f"Chi tiết: {traceback.format_exc()}")

ctk.CTkButton(f6_mid, text="CHẠY XẾP LỊCH TỔNG", fg_color="green", height=40, font=("Arial", 14, "bold"), command=run_auto).pack(side="left", padx=10)
ctk.CTkButton(f6_mid, text="XUẤT LỊCH Y LỆNH", fg_color="#16a085", height=40, command=lambda: xuat_excel_lich_trinh(tree6, f"Lich_YLenh_{datetime.now().strftime('%Y%m%d')}")).pack(side="left", padx=10)

def chuyen_ngay_moi():
    global g_sched, g_rot, sat_cache, g_req, g_proc, g_staff, g_tl, g_ca
    if messagebox.askyesno("Chốt sổ", "Bệnh nhân ĐÃ RA VIỆN sẽ bị xóa khỏi danh sách.\nBệnh nhân ở lại sẽ được làm mới giờ bận.\n\nBạn có chắc chắn chốt sổ hôm nay?"):
        conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
        cur.execute("DELETE FROM BenhNhan WHERE gio_ra != '' AND gio_ra IS NOT NULL")
        cur.execute("UPDATE BenhNhan SET gio_ban='', gio_vao='07:59', gio_ra=''")
        cur.execute("UPDATE NhanSu SET gio_ban=''")
        conn.commit(); conn.close()
        e5_search.delete(0, 'end'); e6_search_bn.delete(0, 'end'); e6_r.delete(0, 'end'); e6_ns_b1.delete(0, 'end'); e6_ns_b2.delete(0, 'end'); e6_search_lich.delete(0, 'end')
        lb_suggestions.place_forget()
        for r in tree_cd.get_children(): tree_cd.delete(r)
        g_sched.clear(); g_rot.clear(); sat_cache.clear(); g_req.clear(); g_proc.clear(); g_staff.clear(); g_tl.clear(); g_ca.clear()
        hien_thi_lich(); tai_ds5(); tai_ds3(); tai_ds8()
        for r in tree7_tt.get_children(): tree7_tt.delete(r)
        for r in tree7_ns.get_children(): tree7_ns.delete(r)
        for r in tree7_fail.get_children(): tree7_fail.delete(r)
        luu_trang_thai(); messagebox.showinfo("Thành công", "Đã dọn dẹp hệ thống. Sẵn sàng cho ngày mới!")

ctk.CTkButton(f6_mid, text="CHỐT SỔ & SANG NGÀY MỚI", fg_color="#8e44ad", hover_color="#9b59b6", height=40, font=("Arial", 12, "bold"), command=chuyen_ngay_moi).pack(side="left", padx=30)

# --- TAB 7 ---
t7 = tabview.tab("7. Thống kê")
f7_top = ctk.CTkFrame(t7, fg_color="transparent"); f7_top.pack(fill="x", pady=5)
ctk.CTkLabel(f7_top, text="BÁO CÁO NĂNG SUẤT & QUẢN TRỊ CHẤT LƯỢNG", font=("Arial", 18, "bold")).pack(side="left", padx=20)
def xuat_excel_thong_ke():
    global g_staff, g_rot
    if not g_staff and not g_rot: return messagebox.showwarning("Cảnh báo", "Chưa có dữ liệu!")
    path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=f"BaoCao_ThongKe_{datetime.now().strftime('%Y%m%d')}")
    if not path: return
    try:
        df_tt = pd.DataFrame([tree7_tt.item(i)['values'] for i in tree7_tt.get_children()], columns=[tree7_tt.heading(c)['text'] for c in tree7_tt['columns']])
        df_ns = pd.DataFrame([tree7_ns.item(i)['values'] for i in tree7_ns.get_children()], columns=[tree7_ns.heading(c)['text'] for c in tree7_ns['columns']])
        auto_format_excel(path, {"Tổng Y Lệnh Toàn Viện": df_tt, "Thống kê Chi Tiết Nhân Sự": df_ns}); messagebox.showinfo("Thành công", "Đã xuất báo cáo Thống kê ra Excel!")
    except Exception as e: messagebox.showerror("Lỗi", str(e))
ctk.CTkButton(f7_top, text="📥 XUẤT EXCEL THỐNG KÊ", fg_color="#d35400", command=xuat_excel_thong_ke).pack(side="right", padx=20)

f7_mid = ctk.CTkFrame(t7, fg_color="transparent"); f7_mid.pack(fill="both", expand=True, padx=5, pady=5)
tree7_tt = create_tree(f7_mid, ("Phân Loại / Thủ Thuật", "Tổng Y Lệnh", "Đã Làm", "Rớt Lịch"), [300, 100, 100, 100])
tree7_ns = create_tree(f7_mid, ("Nhân Viên / Thủ Thuật", "Phân Loại", "Số Ca"), [300, 120, 100])

f7_bot = ctk.CTkFrame(t7, fg_color="transparent"); f7_bot.pack(fill="x", side="bottom", padx=5, pady=5)
ctk.CTkLabel(f7_bot, text="⚠️ DANH SÁCH BỆNH NHÂN RỚT LỊCH", font=("Arial", 14, "bold"), text_color="#c0392b").pack(anchor="w", padx=5, pady=(5,0))
tree7_fail = create_tree(f7_bot, ("STT", "Tên Bệnh Nhân", "Năm Sinh", "Thủ Thuật Bị Hủy", "Phòng", "Nguyên Nhân Cốt Lõi"), [40, 200, 80, 200, 150, 300])

def tai_thong_ke():
    global g_req, g_proc, g_staff, g_rot
    for r in tree7_tt.get_children(): tree7_tt.delete(r)
    for r in tree7_ns.get_children(): tree7_ns.delete(r)
    for r in tree7_fail.get_children(): tree7_fail.delete(r)
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    type_map = {r[0]: r[1] for r in cur.execute("SELECT ten, phan_loai FROM ThuThuat")}; conn.close()
    stats_by_type = {'Loại 1': {}, 'Loại 2': {}, 'Loại 3': {}, 'Chưa phân loại': {}}; grand_req = grand_done = grand_fail = 0
    all_procs = set(list(g_req.keys()) + list(g_proc.keys()))
    for proc in all_procs:
        req = g_req.get(proc, 0); done = g_proc.get(proc, 0); fail = req - done if req > done else 0
        if req > 0 or done > 0:
            p_type = type_map.get(proc, 'Chưa phân loại')
            if p_type not in stats_by_type: stats_by_type[p_type] = {}
            stats_by_type[p_type][proc] = {'req': req, 'done': done, 'fail': fail}
            grand_req += req; grand_done += done; grand_fail += fail
    tree7_tt.tag_configure('type_header', background='#dff9fb', font=('Arial', 10, 'bold'))
    tree7_tt.tag_configure('totalrow', background='#f1c40f', font=('Arial', 10, 'bold'))
    for p_type in ["Loại 1", "Loại 2", "Loại 3", "Chưa phân loại"]:
        if p_type in stats_by_type and stats_by_type[p_type]:
            type_req = sum([d['req'] for d in stats_by_type[p_type].values()]); type_done = sum([d['done'] for d in stats_by_type[p_type].values()]); type_fail = sum([d['fail'] for d in stats_by_type[p_type].values()])
            tree7_tt.insert("", "end", values=(f"▶ {p_type.upper()}", type_req, type_done, type_fail), tags=('type_header',))
            for proc, counts in stats_by_type[p_type].items(): tree7_tt.insert("", "end", values=(f"    • {proc}", counts['req'], counts['done'], counts['fail']))
    if grand_req > 0: tree7_tt.insert("", "end", values=("★ TỔNG CỘNG TOÀN VIỆN", grand_req, grand_done, grand_fail), tags=('totalrow',))
    tree7_ns.tag_configure('staff_total', background='#dff9fb', font=('Arial', 10, 'bold'))
    tree7_ns.tag_configure('type_total', background='#f5f6fa', font=('Arial', 10, 'italic'))
    for staff, data in g_staff.items():
        if not data['procs_done']: continue
        type_groups = {'Loại 1': {}, 'Loại 2': {}, 'Loại 3': {}, 'Chưa phân loại': {}}; total_staff = 0
        for proc, count in data['procs_done'].items():
            p_type = type_map.get(proc, 'Chưa phân loại')
            if p_type not in type_groups: type_groups[p_type] = {}
            type_groups[p_type][proc] = count; total_staff += count
        tree7_ns.insert("", "end", values=(f"BS/KTV: {staff.upper()}", "-", total_staff), tags=('staff_total',))
        for p_type in ["Loại 1", "Loại 2", "Loại 3", "Chưa phân loại"]:
            if p_type in type_groups and type_groups[p_type]:
                type_total = sum(type_groups[p_type].values())
                for proc, count in type_groups[p_type].items(): tree7_ns.insert("", "end", values=(f"  • {proc}", p_type, count))
                tree7_ns.insert("", "end", values=(f"  >> Tổng cộng {p_type}", "-", type_total), tags=('type_total',))
    stt = 1
    for fail in g_rot: tree7_fail.insert("", "end", values=(stt, fail['bn'], fail['ns'], fail['tt'], fail['room'], fail['reason'])); stt += 1

# --- TAB 8 ---
t8 = tabview.tab("8. TRỰC THỨ 7")
f8_top = ctk.CTkFrame(t8, fg_color="transparent"); f8_top.pack(fill="x", padx=10, pady=5)
ctk.CTkLabel(f8_top, text="ĐIỀU PHỐI CA TRỰC THỨ 7", font=("Arial", 18, "bold"), text_color="#8e44ad").pack(side="left")
f8_main = ctk.CTkFrame(t8, fg_color="transparent"); f8_main.pack(fill="both", expand=True, padx=10, pady=5)
f8_l = ctk.CTkFrame(f8_main, width=280); f8_l.pack(side="left", fill="y", padx=5); f8_l.pack_propagate(False)
ctk.CTkLabel(f8_l, text="👨‍⚕️ Ai đi trực hôm nay?", font=("Arial", 13, "bold")).pack(pady=10)
fr8_ns = ctk.CTkScrollableFrame(f8_l); fr8_ns.pack(fill="both", expand=True, padx=5, pady=5)
t8_ns_vars = {}; t8_ns_time_frames = {}; t8_ns_entries = {}
f8_r = ctk.CTkFrame(f8_main); f8_r.pack(side="left", fill="both", expand=True, padx=5)
f8_r_top1 = ctk.CTkFrame(f8_r, fg_color="transparent"); f8_r_top1.pack(fill="x", pady=5)
ctk.CTkLabel(f8_r_top1, text="📋 Danh sách BN (Theo Phòng):", font=("Arial", 13, "bold")).pack(side="left", padx=10)
e8_search_bn_t8 = ctk.CTkEntry(f8_r_top1, placeholder_text="🔍 Gõ tên hoặc phòng...", width=200); e8_search_bn_t8.pack(side="left", padx=10)
ctk.CTkButton(f8_r_top1, text="🔄 Tải/Làm mới BN", command=lambda: tai_ds8(), width=130, fg_color="#3498db").pack(side="right", padx=10)
f8_r_top2 = ctk.CTkFrame(f8_r, fg_color="transparent"); f8_r_top2.pack(fill="x", pady=5)

def xuat_ds_thu7():
    if not sat_cache: return messagebox.showwarning("Lỗi", "Chưa có danh sách bệnh nhân!")
    data = []
    for bid, d in sat_cache.items():
        chosen = [k for k, v in d['vars'].items() if v.get() not in ["off", "0", ""]]; r = d['info']
        data.append({'Mã BN': r[0], 'Tên Bệnh Nhân': r[1], 'Thủ Thuật Thứ 7': ", ".join(chosen)})
    path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")], initialfile=f"DS_ChonThuThuat_T7_{datetime.now().strftime('%Y%m%d')}")
    if path: auto_format_excel(path, {"DS_ChonThuThuat_T7": pd.DataFrame(data)}); messagebox.showinfo("Thành công", "Đã xuất danh sách chọn thủ thuật!")

def nhap_ds_thu7():
    if not sat_cache: return messagebox.showwarning("Lỗi", "Hãy 'Tải Bệnh Nhân' trước khi nhập file!")
    path = filedialog.askopenfilename(filetypes=[("Excel/CSV files", "*.xlsx *.xls *.csv")])
    if not path: return
    try:
        df = pd.read_csv(path).fillna("") if path.endswith('.csv') else pd.read_excel(path).fillna("")
        for d in sat_cache.values():
            for v in d['vars'].values(): v.set("off")
        count = 0
        for _, row in df.iterrows():
            bid = row.get('Mã BN'); tt_str = row.get('Thủ Thuật Thứ 7')
            if bid in sat_cache and tt_str:
                for tt_name in [x.strip() for x in str(tt_str).split(",") if x.strip()]:
                    if tt_name in sat_cache[bid]['vars']: sat_cache[bid]['vars'][tt_name].set(tt_name)
                count += 1
        messagebox.showinfo("Thành công", f"Đã nạp file và tự động tích chọn cho {count} bệnh nhân!")
    except Exception as e: messagebox.showerror("Lỗi", str(e))

ctk.CTkButton(f8_r_top2, text="📂 Nhập DS Đã Lưu", command=nhap_ds_thu7, width=130, fg_color="#8e44ad").pack(side="left", padx=10)
ctk.CTkButton(f8_r_top2, text="💾 Lưu/Xuất DS", command=xuat_ds_thu7, width=120, fg_color="#d35400").pack(side="left", padx=10)
ctk.CTkButton(f8_r_top2, text="🚀 CHẠY LỊCH THỨ 7", fg_color="green", font=("Arial", 12, "bold"), command=lambda: run_saturday_schedule()).pack(side="right", padx=10)
fr8_bn_tt = ctk.CTkScrollableFrame(f8_r); fr8_bn_tt.pack(fill="both", expand=True, padx=10, pady=5)

def loc_bn_t8(event=None):
    tukhoa = e8_search_bn_t8.get().lower()
    for widget in fr8_bn_tt.winfo_children(): widget.pack_forget()
    for bn_id, data in sat_cache.items():
        if tukhoa in str(data['info'][1]).lower() or tukhoa in str(data['info'][4]).lower(): data['frame'].pack(fill="x", pady=5, padx=5)
e8_search_bn_t8.bind("<KeyRelease>", loc_bn_t8)

def tai_ds8():
    for w in fr8_ns.winfo_children(): w.destroy()
    t8_ns_vars.clear(); t8_ns_time_frames.clear(); t8_ns_entries.clear()
    conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
    for r in cur.execute("SELECT ten FROM NhanSu"):
        name = r[0]; cont = ctk.CTkFrame(fr8_ns, fg_color="transparent"); cont.pack(fill="x", pady=2)
        v = ctk.StringVar(value="off"); t8_ns_vars[name] = v; tf = ctk.CTkFrame(cont, fg_color="transparent"); t8_ns_time_frames[name] = tf
        def toggle(n=name, var=v):
            if var.get() != "off": t8_ns_time_frames[n].pack(fill="x", padx=15, pady=2)
            else: t8_ns_time_frames[n].pack_forget()
        ctk.CTkCheckBox(cont, text=name, variable=v, onvalue=name, offvalue="off", command=toggle).pack(anchor="w", padx=5)
        r1 = ctk.CTkFrame(tf, fg_color="transparent"); r1.pack(fill="x")
        e_s1 = ctk.CTkEntry(r1, width=65, placeholder_text="08:00"); e_s1.insert(0,"08:00"); e_s1.pack(side="left", padx=1); e_s2 = ctk.CTkEntry(r1, width=65, placeholder_text="12:00"); e_s2.insert(0,"12:00"); e_s2.pack(side="left", padx=1)
        r2 = ctk.CTkFrame(tf, fg_color="transparent"); r2.pack(fill="x")
        e_c1 = ctk.CTkEntry(r2, width=65, placeholder_text="Chiều"); e_c1.pack(side="left", padx=1); e_c2 = ctk.CTkEntry(r2, width=65, placeholder_text="Đến"); e_c2.pack(side="left", padx=1)
        for e in [e_s1, e_s2, e_c1, e_c2]: e.bind("<KeyRelease>", lambda ev, w=e: auto_format_time(ev, w))
        t8_ns_entries[name] = (e_s1, e_s2, e_c1, e_c2)
    for w in fr8_bn_tt.winfo_children(): w.destroy()
    global sat_cache; sat_cache.clear()
    for r in cur.execute("SELECT id, ten, nam_sinh, thu_thuat, phong, gio_vao, gio_ra, gio_ban, loai_bn FROM BenhNhan WHERE thu_thuat != '' ORDER BY phong ASC, ten ASC"):
        bn_id, bn_name, y_lenh = r[0], r[1], r[3]
        bn_cont = ctk.CTkFrame(fr8_bn_tt, border_width=1); bn_cont.pack(fill="x", pady=5, padx=5)
        ctk.CTkLabel(bn_cont, text=f"👤 {bn_name.upper()} ({r[2]}) - {r[4]}", font=("Arial", 12, "bold")).pack(anchor="w", padx=10, pady=2)
        tt_frame = ctk.CTkFrame(bn_cont, fg_color="transparent"); tt_frame.pack(fill="x", padx=25)
        sat_cache[bn_id] = {'info': r, 'vars': {}, 'frame': bn_cont}
        for tt in [x.strip() for x in y_lenh.split(",") if x.strip()]:
            v = ctk.StringVar(value="off"); sat_cache[bn_id]['vars'][tt] = v; ctk.CTkCheckBox(tt_frame, text=tt, variable=v, onvalue=tt, offvalue="off").pack(side="left", padx=8)
    conn.close(); e8_search_bn_t8.delete(0, 'end'); loc_bn_t8()

def run_saturday_schedule():
    global g_staff, g_proc, g_req, g_rot, g_sched, g_tl, g_ca
    allowed_staff = [v.get() for v in t8_ns_vars.values() if v.get() not in ["off", "0", ""]]
    if not allowed_staff: return messagebox.showwarning("Cảnh báo", "Chọn Nhân sự trực!")
    final_pats = []
    for bid, data in sat_cache.items():
        chosen = [v.get() for v in data['vars'].values() if v.get() not in ["off", "0", ""]]
        if chosen:
            r = data['info']; final_pats.append({'id':r[0],'ten':r[1],'ns':r[2],'tt':", ".join(chosen),'phong':r[4],'vao':r[5],'ra':r[6],'ban':r[7],'loai':r[8], 'rand_seed': random.random()})
    if not final_pats: return messagebox.showwarning("Cảnh báo", "Chưa chọn thủ thuật nào cho bệnh nhân!")

    try:
        for r in tree6.get_children(): tree6.delete(r)
        conn = sqlite3.connect(DB_NAME); cur = conn.cursor()
        m_types = {}; m_timeline = {"Thủ công": []} 
        for r in cur.execute("SELECT ten_loai, ma_may FROM DanhSachMay WHERE trang_thai='Sẵn sàng'"):
            if r[0] not in m_types: m_types[r[0]] = []
            m_types[r[0]].append(r[1]); m_timeline[r[1]] = []
        s_timeline = {s: [] for s in allowed_staff}; s_k = {}; s_ca = {}; s_role = {}
        g_staff.clear(); g_proc.clear(); tt_i = {}; g_rot.clear(); results = []; g_req.clear()
        for r in cur.execute("SELECT ten, vai_tro, ky_nang FROM NhanSu"):
            s = r[0]
            if s not in allowed_staff: continue
            s_role[s] = r[1]; s_k[s] = [x.strip() for x in str(r[2]).split(", ")] if r[2] else []
            e_s1, e_s2, e_c1, e_c2 = t8_ns_entries[s]; staff_shifts = []
            if e_s1.get() and e_s2.get(): staff_shifts.append((t2m(e_s1.get()), t2m(e_s2.get())))
            if e_c1.get() and e_c2.get(): staff_shifts.append((t2m(e_c1.get()), t2m(e_c2.get())))
            s_ca[s] = staff_shifts
            g_staff[s] = {'shift_mins': sum([ce - cs for cs, ce in staff_shifts]) if staff_shifts else 0, 'busy_mins': 0, 'used_mins': 0, 'procs_done': {}, 'skills': s_k[s]}
            
        for r in cur.execute("SELECT ten, loai_may, thoi_gian_may, thoi_gian_nguoi, loai_chuyen_mon, can_rut_may, can_nguoi_phu, ds_nguoi_phu FROM ThuThuat"):
            dm = int(r[2]) if r[2] else 15; dn = int(r[3]) if r[3] else 5; ds_phu = [x.strip() for x in str(r[7]).split(',')] if r[7] else []
            tt_i[r[0]] = (r[1], max(1, dm), max(1, dn), r[4], int(r[5]) if r[5] else 1, int(r[6]) if r[6] else 0, ds_phu)
            g_proc[r[0]] = 0; g_req[r[0]] = 0
            
        room_beds_tracker = {}
        for r in cur.execute("SELECT ten_phong, so_giuong, danh_sach_giuong FROM PhongThuThuat"):
            room_name = r[0]; num_beds = int(r[1]) if r[1] else 15; bed_str = str(r[2]).strip() if r[2] else ""
            bed_list = [x.strip() for x in bed_str.split(",") if x.strip()] if bed_str and bed_str != 'None' else [f"Giường {i}" for i in range(1, num_beds + 1)]
            room_beds_tracker[room_name] = {b: [] for b in bed_list}
                
        patients = []
        for bn in final_pats:
            raw_pending = [x.strip() for x in bn['tt'].split(",") if x.strip()]; clean_pending = []
            for ttn in raw_pending: clean_pending.append(ttn); g_req[ttn] = g_req.get(ttn, 0) + 1
            clean_pending.sort(key=lambda ttn: (1 if "cấy chỉ" in ttn.lower() else 0, 0 if tt_i.get(ttn, ("", 0, 0, "PHCN", 1, 0, []))[3] == "YHCT" else 1, -tt_i.get(ttn, ("", 0, 999, "PHCN", 1, 0, []))[2]))
            arrive_time = t2m(bn['vao']) if bn['vao'] else 479; arrive_time = 479 if arrive_time == 0 else arrive_time
            patients.append({'name': f"{bn['ten'].upper()}", 'nam_sinh': bn['ns'], 'loai_bn': bn['loai'], 'primary_room': bn['phong'], 'arrive': arrive_time, 'leave': t2m(bn['ra']) if bn['ra'] else 9999, 'busy': [(t2m(p.split("-")[0]), t2m(p.split("-")[1])) for p in str(bn['ban']).split(",") if "-" in p] if bn['ban'] else [], 'pending': clean_pending, 'free_at': arrive_time, 'failed': False, 'skip_phase1': False, 'rand_seed': bn['rand_seed']})
            
        ngay_xep_lich = e_ngay_xep.get()
        for phase_idx in [1, 2]:
            for t_now in range(420, 1200):
                if 720 <= t_now < 780: continue 
                el_pats = [p for p in patients if not p['failed'] and p['pending'] and p['free_at'] <= t_now]
                el_pats.sort(key=lambda x: (x['leave'], x['rand_seed'])) 
                for p in el_pats:
                    if phase_idx == 1 and p['skip_phase1']: continue
                    in_cls = False
                    for bs, be in p['busy']:
                        if bs <= t_now < be: p['free_at'] = be; in_cls = True; break
                    if in_cls: continue
                    scheduled_something = False
                    for i, ttn in enumerate(p['pending']):
                        if ttn not in tt_i: p['pending'].pop(i); scheduled_something = True; break
                        lm, dm, dn, loai_cm, can_rut, can_phu, ds_phu = tt_i[ttn]
                        if lm == "Thủ công": dn = dm
                        if t_now < 720 and (t_now + dm) > 720: continue 
                        time_limit = 9999 if p['loai_bn'] == 'Nội trú ban ngày' else 540
                        is_time_breached = (t_now + dm) > (p['arrive'] + time_limit)
                        is_leave_breached = (p['leave'] != 9999 and t_now + dm > p['leave'])
                        if is_time_breached or is_leave_breached:
                            p['skip_phase1'] = True
                            if phase_idx == 2:
                                g_rot.append({'bn': p['name'], 'ns': p['nam_sinh'], 'tt': ttn, 'room': p['primary_room'], 'staff': "Trực Thứ 7", 'reason': f"Vượt giới hạn ({m2t(p['leave'])})" if is_leave_breached else "Vượt giới hạn sinh học"})
                                p['pending'].pop(i) 
                            scheduled_something = True; break 
                        
                        setup_start = t_now; setup_end = t_now + dn + 1; m_end = t_now + dm + 1
                        
                        # CHỐT CHẶN AN TOÀN NONETYPE
                        if dm > dn and can_rut == 1: tear_start = t_now + dm - 1; tear_end = t_now + dm + 1
                        else: tear_start = None; tear_end = None
                        
                        possible_machines = ["Thủ công"] if lm == "Thủ công" else m_types.get(lm, [])
                        if not possible_machines: continue
                        target_room = p['primary_room'] 
                        for m in possible_machines:
                            if m != "Thủ công":
                                overlap_m = False
                                for ms, me in m_timeline[m]:
                                    if is_overlap(t_now, m_end, ms, me): overlap_m = True; break
                                if overlap_m: continue 
                            assigned_bed = None
                            if target_room in room_beds_tracker:
                                for b_id, b_timeline in room_beds_tracker[target_room].items():
                                    overlap_b = False
                                    for bs, be in b_timeline:
                                        if is_overlap(t_now, m_end, bs, be): overlap_b = True; break
                                    if not overlap_b: assigned_bed = b_id; break
                            if not assigned_bed and target_room in room_beds_tracker and len(room_beds_tracker[target_room]) > 0: continue 
                            
                            main_cands = []; sub_cands = []; mon_cands = []
                            for s in s_k.keys():
                                if ttn not in s_k.get(s, []): continue
                                shift_ok = False
                                if not s_ca[s]: shift_ok = True
                                else:
                                    for cs, ce in s_ca[s]:
                                        if setup_start >= cs and m_end <= ce: shift_ok = True; break
                                if not shift_ok: continue
                                r_role = s_role.get(s, '')
                                f_set = True
                                for bs, be in s_timeline[s]:
                                    if is_overlap(setup_start, setup_end, bs, be): f_set = False; break
                                
                                if f_set and r_role != 'Điều dưỡng':
                                    if loai_cm == 'YHCT' and r_role == 'Bác sĩ': main_cands.append(s)
                                    elif loai_cm == 'PHCN' and r_role in ['Bác sĩ', 'Kỹ thuật viên']: main_cands.append(s)

                                if not ds_phu or s in ds_phu:
                                    if can_phu == 1 and f_set and r_role in ['Kỹ thuật viên', 'Điều dưỡng']: sub_cands.append(s)
                                    if can_rut == 1 and r_role in ['Kỹ thuật viên', 'Điều dưỡng'] and tear_start is not None:
                                        f_tear = True
                                        for bs, be in s_timeline[s]:
                                            if is_overlap(tear_start, tear_end, bs, be): f_tear = False; break
                                        if f_tear:
                                            shift_ok_tear = False
                                            if not s_ca[s]: shift_ok_tear = True
                                            else:
                                                for cs, ce in s_ca[s]:
                                                    if tear_start >= cs and tear_end <= ce: shift_ok_tear = True; break
                                            if shift_ok_tear: mon_cands.append(s)

                            if not main_cands: continue

                            if can_phu == 1:
                                if not sub_cands: continue 
                                main_cands.sort(key=lambda x: g_staff[x]['used_mins'])
                                fs_main = main_cands[0]
                                valid_subs = [x for x in sub_cands if x != fs_main]
                                if not valid_subs: continue
                                valid_subs.sort(key=lambda x: g_staff[x]['used_mins'])
                                fs_sub = valid_subs[0]
                                fs_mon = None
                                if can_rut == 1 and tear_start is not None:
                                    valid_mons = [x for x in mon_cands if x not in [fs_main, fs_sub]]
                                    if valid_mons: 
                                        valid_mons.sort(key=lambda x: g_staff[x]['used_mins'])
                                        fs_mon = valid_mons[0]
                                
                                bed_str = f"{assigned_bed}" if assigned_bed else ""
                                results.append({'NGAY': ngay_xep_lich, 'HOTEN': p['name'], 'NAMSINH': p['nam_sinh'], 'PHONG': target_room, 'DICHVU': ttn, 'GIODIENRA': m2t(t_now), 'GIOKETTHUC': m2t(t_now+dm), 'NV CHÍNH': fs_main, 'NV PHỤ': fs_sub, 'MAY': m, 'GIUONG': bed_str, 't_sort': t_now})
                                g_proc[ttn] += 1; u_time = setup_end - setup_start
                                g_staff[fs_main]['procs_done'][ttn] = g_staff[fs_main]['procs_done'].get(ttn, 0) + 1; g_staff[fs_main]['used_mins'] += u_time; s_timeline[fs_main].append((setup_start, setup_end))
                                g_staff[fs_sub]['procs_done'][ttn] = g_staff[fs_sub]['procs_done'].get(ttn, 0) + 1; g_staff[fs_sub]['used_mins'] += u_time; s_timeline[fs_sub].append((setup_start, setup_end))
                                if fs_mon: g_staff[fs_mon]['used_mins'] += (tear_end - tear_start); s_timeline[fs_mon].append((tear_start, tear_end))
                                if m != "Thủ công": m_timeline[m].append((t_now, m_end)) 
                                if target_room in room_beds_tracker and assigned_bed: room_beds_tracker[target_room][assigned_bed].append((t_now, m_end))
                                p['free_at'] = m_end; p['pending'].pop(i); scheduled_something = True; break 
                            else:
                                main_cands.sort(key=lambda x: (0 if loai_cm == 'PHCN' and s_role.get(x, '') != 'Bác sĩ' else 1, g_staff[x]['used_mins']))
                                fs_main = main_cands[0]; fs_mon = None
                                if loai_cm == 'YHCT' and can_rut == 1 and tear_start is not None:
                                    valid_mons = [x for x in mon_cands if x != fs_main]
                                    if valid_mons:
                                        valid_mons.sort(key=lambda x: g_staff[x]['used_mins'])
                                        fs_mon = valid_mons[0]
                                
                                bed_str = f"{assigned_bed}" if assigned_bed else ""
                                results.append({'NGAY': ngay_xep_lich, 'HOTEN': p['name'], 'NAMSINH': p['nam_sinh'], 'PHONG': target_room, 'DICHVU': ttn, 'GIODIENRA': m2t(t_now), 'GIOKETTHUC': m2t(t_now+dm), 'NV CHÍNH': fs_main, 'NV PHỤ': fs_mon if fs_mon else "", 'MAY': m, 'GIUONG': bed_str, 't_sort': t_now})
                                g_proc[ttn] += 1; g_staff[fs_main]['procs_done'][ttn] = g_staff[fs_main]['procs_done'].get(ttn, 0) + 1
                                used_time = setup_end - setup_start + (tear_end - tear_start if tear_start else 0)
                                g_staff[fs_main]['used_mins'] += used_time; s_timeline[fs_main].append((setup_start, setup_end)); s_timeline[fs_main].append((tear_start, tear_end)) if tear_start else None
                                if fs_mon: g_staff[fs_mon]['used_mins'] += (tear_end - tear_start); s_timeline[fs_mon].append((tear_start, tear_end))
                                if m != "Thủ công": m_timeline[m].append((t_now, m_end)) 
                                if target_room in room_beds_tracker and assigned_bed: room_beds_tracker[target_room][assigned_bed].append((t_now, m_end))
                                p['free_at'] = m_end; p['pending'].pop(i); scheduled_something = True; break 
                        if scheduled_something: break 
        for p in patients:
            for ttn in p['pending']: g_rot.append({'bn': p['name'], 'ns': p['nam_sinh'], 'tt': ttn, 'room': p['primary_room'], 'staff': "Không có người", 'reason': "Kẹt nhân sự/máy/giường"})
        results.sort(key=lambda x: x['t_sort'])
        g_sched.clear(); g_sched.extend(results); g_tl.clear(); g_tl.update(s_timeline); g_ca.clear(); g_ca.update(s_ca)
        e6_search_lich.delete(0, 'end'); hien_thi_lich(); conn.close(); tai_thong_ke(); luu_trang_thai()
        messagebox.showinfo("Thành công", f"Đã xếp lịch xong\nĐã xếp {len(g_sched)} ca, bị rớt lịch {len(g_rot)} ca")
    except Exception as e:
        messagebox.showerror("Lỗi thuật toán", f"Chi tiết: {traceback.format_exc()}")

tai_ds1(); tai_trang_thai(); app.mainloop()