import os
import customtkinter as ctk
from tkinter import ttk
from tkinter import messagebox
from pygame import mixer
import pandas as pd
from datetime import datetime
import gspread
import win32print
import win32ui
import win32con
import threading

# Thử import thư viện qrcode hỗ trợ tạo mã QR
try:
    import qrcode
    HAS_QRCODE = True
except ImportError:
    HAS_QRCODE = False
    print("=" * 70)
    print("❌ CHƯA CÀI THƯ VIỆN 'qrcode' -> TEM IN RA SẼ KHÔNG CÓ MÃ QR!")
    print("👉 Cách khắc phục: mở CMD/Terminal rồi chạy lệnh: pip install qrcode")
    print("=" * 70)

# --- CẤU HÌNH LOGIC PHÂN CHIA RỔ & GỘP PHÂN LOẠI THEO TỪNG THƯƠNG HIỆU ---
CATEGORY_MAPPING = {
    'CB Bypass': 'cb450',
    'CB Other': 'cb450',
    'SPX-SBS': 'SBS',
    'SPX-HV': 'HV',
    'SPX-loreal': 'loreal',
    'SPX-vinamilk': 'vinamilk'
}

def get_basket_name(category, is_spx=True):
    if pd.isna(category) or not category or str(category).strip().lower() in ['none', 'nan', '']:
        return "KO PHÂN LOẠI"
    cat_str = str(category).strip()
    return CATEGORY_MAPPING.get(cat_str, cat_str)

PREDEFINED_BASKETS = [
    'cb450',
    'CB High Risk',
    'SBS',
    'HV',
    'loreal',
    'vinamilk',
    'No',
    'lite',
    'resell',
    'KO PHÂN LOẠI'
]

# --- QUẢN LÝ TRẠNG THÁI RỔ TRONG BỘ NHỚ ---
baskets_data_spx = {b: [] for b in PREDEFINED_BASKETS} 
baskets_data_ghn = {b: [] for b in PREDEFINED_BASKETS}
baskets_ui_spx = {}
baskets_ui_ghn = {}

# CACHE ĐỄ QUẢN LÝ SỐ THỨ TỰ RỔ TỰ ĐỘNG TĂNG KHÔNG CẦN CHỜ SHEET
basket_seq_cache = {} 
_basket_seq_lock = threading.Lock()
_basket_seq_sheet_values = {}

# --- KẾT NỐI HỆ THỐNG GOOGLE SHEETS ---
gd_spx = gspread.service_account(filename=r'file2/phong-thanhnguyen.json')
gs_spx = gd_spx.open_by_key('1F-TqEnz9TwHAyxf28zywwwA1Z5L2Yc9VqINT4f1nKPo')
ws_spx = gs_spx.worksheet('SPX_v2')
return_list_spx = pd.DataFrame({'Return TN': ws_spx.col_values(1)[1:]})

gd_ghn = gspread.service_account(filename=r'file2/phong-thanhnguyen.json')
gs_ghn = gd_ghn.open_by_key('1F-TqEnz9TwHAyxf28zywwwA1Z5L2Yc9VqINT4f1nKPo')
ws_ghn = gs_ghn.worksheet('GHN_v2')
return_list_ghn = pd.DataFrame({'Return TN': ws_ghn.col_values(1)[1:]})

gd = gspread.service_account(filename=r'file2/phong-thanhnguyen.json')
gs = gd.open_by_key('1F-TqEnz9TwHAyxf28zywwwA1Z5L2Yc9VqINT4f1nKPo')
ws1 = gs.worksheet("source_data")
source = ws1.get("A2:G")
columns_source = ['Return TN','Order SN','Return_Status','Seller Type','Shop','CAT','Type']
source = [(row + [''] * len(columns_source))[:len(columns_source)] for row in source]
df_source = pd.DataFrame(source, columns=columns_source)

# ========================================================================================
# NẠP LẠI DỮ LIỆU PHÂN LOẠI (SHEET "source_data")
# ========================================================================================
def reload_source_data():
    def _worker():
        global df_source
        root.after(0, lambda: btn_reload_source.configure(state="disabled", text="⏳ Đang nạp..."))
        try:
            new_source = ws1.get("A2:G")
            new_source = [(row + [''] * len(columns_source))[:len(columns_source)] for row in new_source]
            df_source = pd.DataFrame(new_source, columns=columns_source)
            so_dong = len(df_source)
            gio_cap_nhat = datetime.now().strftime("%H:%M:%S")
            print(f"✅ Đã nạp lại source_data thành công: {so_dong} dòng vào lúc {gio_cap_nhat}.")
            root.after(0, lambda: label_source_status.configure(
                text=f"📄 Source Data: {so_dong} dòng | Cập nhật lúc {gio_cap_nhat}"))
        except Exception as e:
            print(f"⚠️ Lỗi khi nạp lại source_data: {e}")
            root.after(0, lambda err=e: messagebox.showerror(
                "Lỗi Nạp Data", f"Không thể nạp lại dữ liệu từ sheet 'source_data':\n{err}"))
        finally:
            root.after(0, lambda: btn_reload_source.configure(state="normal", text="🔄 Nạp Lại Source Data"))
    threading.Thread(target=_worker, daemon=True).start()

countspx = 0
counttotalspx = 0
countghn = 0
counttotalghn = 0
listscanspx = []
listscanghn = []

def voice(voice_name):
    try:
        mixer.init()  
        mixer.music.load(f"file2/{voice_name}.mp3")  
        mixer.music.play() 
    except Exception as e:
        print(f"Lỗi âm thanh: {e}")

# ========================================================================================
# LOGIC LƯU TRỮ OFFLINE CSV
# ========================================================================================
def log_scan_csv(is_spx, ma_vc, phan_loai, gio_nhan, ma_ro="", status="Chưa Submit"):
    filename = "backup_spx.csv" if is_spx else "backup_ghn.csv"
    new_row = pd.DataFrame([{
        "Mã Vận Chuyển": str(ma_vc).strip(),
        "Phân Loại": str(phan_loai).strip(),
        "Giờ Nhận": str(gio_nhan).strip(),
        "Mã Rổ": str(ma_ro).strip(),
        "Trạng Thái Submit": str(status).strip()
    }])
    if not os.path.exists(filename):
        new_row.to_csv(filename, index=False, encoding='utf-8-sig')
    else:
        new_row.to_csv(filename, mode='a', header=False, index=False, encoding='utf-8-sig')

def update_csv_after_submit(is_spx, ma_vc_list, ma_ro):
    filename = "backup_spx.csv" if is_spx else "backup_ghn.csv"
    if os.path.exists(filename):
        try:
            df = pd.read_csv(filename, encoding='utf-8-sig', dtype=str)
            ma_vc_set = set(str(x).strip() for x in ma_vc_list)
            
            mask = (
                df['Mã Vận Chuyển'].astype(str).str.strip().isin(ma_vc_set)
                & (df['Trạng Thái Submit'].astype(str).str.strip() == 'Chưa Submit')
            )
            df.loc[mask, 'Trạng Thái Submit'] = 'Đã Submit'
            df.loc[mask, 'Mã Rổ'] = str(ma_ro)
            
            df.to_csv(filename, index=False, encoding='utf-8-sig')
            print(f"✅ Đã cập nhật file CSV offline [{filename}] sang trạng thái 'Đã Submit'.")
        except Exception as e:
            print(f"Lỗi đồng bộ file CSV offline: {e}")

def load_offline_backup(is_spx):
    filename = "backup_spx.csv" if is_spx else "backup_ghn.csv"
    if not os.path.exists(filename):
        return
    try:
        df = pd.read_csv(filename, encoding='utf-8-sig', dtype=str).fillna('')
        df_pending = df[df['Trạng Thái Submit'].astype(str).str.strip() == 'Chưa Submit']
        
        tree_widget = tree if is_spx else tree2
        baskets_data = baskets_data_spx if is_spx else baskets_data_ghn
        baskets_ui = baskets_ui_spx if is_spx else baskets_ui_ghn
        global countspx, countghn
        
        for _, row in df_pending.iterrows():
            ma_vc = str(row['Mã Vận Chuyển']).strip()
            fn = str(row['Phân Loại']).strip()
            gio = str(row['Giờ Nhận']).strip()
            ma_ro = str(row['Mã Rổ']).strip()
            
            tree_widget.insert('', 'end', values=(ma_vc, fn, gio, ma_ro))
            
            if not ma_ro or "Lỗi mạng" in ma_ro:
                basket_name = get_basket_name(fn, is_spx)
                if basket_name not in baskets_data:
                    baskets_data[basket_name] = []
                baskets_data[basket_name].append(ma_vc)
                
                if is_spx: countspx += 1
                else: countghn += 1
                
        for b_name in baskets_data:
            if b_name in baskets_ui:
                baskets_ui[b_name].configure(text=str(len(baskets_data[b_name])))
                
        if is_spx: qtyspx.configure(text=str(countspx))
        else: qtyghn.configure(text=str(countghn))
        print(f"♻️ Đã khôi phục thành công các đơn tồn đọng từ file CSV offline của {'SPX' if is_spx else 'GHN'}!")
    except Exception as e:
        print(f"Lỗi trong quá trình đọc file khôi phục dữ liệu: {e}")

# ========================================================================================
# CHỨC NĂNG IN NHÃN MÁY NHIỆT KHỔ LỚN 10cm x 5cm
# ========================================================================================
def print_basket_label(basket_id, qty=0):
    try:
        printer_name = win32print.GetDefaultPrinter()
        hdc = win32ui.CreateDC()
        hdc.CreatePrinterDC(printer_name)
        
        hdc.StartDoc("Basket Label QR")
        hdc.StartPage()
        hdc.SetMapMode(win32con.MM_TEXT)
        
        dpi_x = hdc.GetDeviceCaps(win32con.LOGPIXELSX)
        dpi_y = hdc.GetDeviceCaps(win32con.LOGPIXELSY)
        
        width_px = int(3.937 * dpi_x)
        height_px = int(1.968 * dpi_y)
        
        margin = int(0.12 * dpi_x)
        pen = win32ui.CreatePen(win32con.PS_SOLID, int(dpi_x * 0.02), 0)
        old_pen = hdc.SelectObject(pen)
        hdc.Rectangle((margin, margin, width_px - margin, height_px - margin))
        hdc.SelectObject(old_pen)
        
        qr_size_px = int(1.45 * dpi_x)
        qr_x = width_px - margin - qr_size_px - int(0.08 * dpi_x)
        qr_y = (height_px - qr_size_px) // 2
        
        if HAS_QRCODE:
            try:
                qr = qrcode.QRCode(version=1, box_size=1, border=4, error_correction=qrcode.constants.ERROR_CORRECT_L)
                qr.add_data(basket_id)
                qr.make(fit=True)
                matrix = qr.get_matrix()
                
                num_blocks = len(matrix)
                block_size = qr_size_px // num_blocks
                actual_qr_size = block_size * num_blocks
                
                qr_x_aligned = qr_x + (qr_size_px - actual_qr_size) // 2
                qr_y_aligned = qr_y + (qr_size_px - actual_qr_size) // 2
                
                brush = win32ui.CreateBrush(win32con.BS_SOLID, 0, 0)
                old_brush = hdc.SelectObject(brush)
                pen_null = win32ui.CreatePen(win32con.PS_SOLID, 0, 0)
                old_pen2 = hdc.SelectObject(pen_null)
                
                for r in range(num_blocks):
                    for c in range(num_blocks):
                        if matrix[r][c]:
                            bx1 = qr_x_aligned + c * block_size
                            by1 = qr_y_aligned + r * block_size
                            bx2 = bx1 + block_size
                            by2 = by1 + block_size
                            hdc.Rectangle((bx1, by1, bx2, by2))
                            
                hdc.SelectObject(old_brush)
                hdc.SelectObject(old_pen2)
            except Exception as qr_err:
                print(f"⚠️ Lỗi khởi dựng mã QR: {qr_err}")
        
        font_title_h = int(dpi_y * 0.10)
        font_id_h_max = int(dpi_y * 0.45)
        font_id_h_min = int(dpi_y * 0.30)
        font_seq_h = int(dpi_y * 0.4)
        font_qty_h = int(dpi_y * 0.25)
        font_footer_h = int(dpi_y * 0.10)
        
        font_title = win32ui.CreateFont({"name": "Arial", "height": font_title_h, "weight": win32con.FW_BOLD})
        font_seq = win32ui.CreateFont({"name": "Arial", "height": font_seq_h, "weight": win32con.FW_BOLD})
        font_qty = win32ui.CreateFont({"name": "Arial", "height": font_qty_h, "weight": win32con.FW_BOLD})
        font_footer = win32ui.CreateFont({"name": "Arial", "height": font_footer_h, "weight": win32con.FW_NORMAL})
        
        text_margin = int(margin * 1.3)
        max_text_width = qr_x - text_margin - int(0.05 * dpi_x)
        
        hdc.SelectObject(font_title)
        hdc.TextOut(text_margin, int(margin * 1.2), "CHỐT RỔ / COMPLETED")
        
        if ' ' in basket_id:
            line_name, line_seq = basket_id.rsplit(' ', 1)
        else:
            line_name, line_seq = basket_id, ""
        
        fit_h = font_id_h_max
        while True:
            font_id = win32ui.CreateFont({"name": "Arial", "height": fit_h, "weight": win32con.FW_BOLD})
            hdc.SelectObject(font_id)
            text_w, _ = hdc.GetTextExtent(line_name)
            if text_w <= max_text_width or fit_h <= font_id_h_min:
                break
            fit_h -= 2
        
        line1_y = int(margin * 1.2 + font_title_h * 1.1)
        hdc.TextOut(text_margin, line1_y, line_name)
        
        hdc.SelectObject(font_seq)
        line2_y = int(line1_y + fit_h * 1.05)
        hdc.TextOut(text_margin, line2_y, line_seq)

        hdc.SelectObject(font_qty)
        line3_y = int(line2_y + font_seq_h * 1.1)
        hdc.TextOut(text_margin, line3_y, f"SL: {qty}")
        
        hdc.SelectObject(font_footer)
        now_str = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
        hdc.TextOut(text_margin, int(height_px - margin * 1.3 - font_footer_h * 2.0), f"Thời gian: {now_str}")
        hdc.TextOut(text_margin, int(height_px - margin * 1.3 - font_footer_h * 1.0), "Shopee VN Reverse Logistics")
        
        hdc.EndPage()
        hdc.EndDoc()
        print(f"🖨️ Đã in tem rổ: {basket_id} | SL: {qty} thành công!")
    except Exception as e:
        print(f"⚠️ Lỗi in tem: {e}")

# ========================================================================================
# LOGIC DÒ SỐ THỨ TỰ RỔ TỰ ĐỘNG TĂNG TRONG BỘ NHỚ
# ========================================================================================
def get_next_basket_seq_from_sheet(is_spx, basket_name):
    ws_sheet = ws_spx if is_spx else ws_ghn
    today_str = datetime.now().strftime("%d/%m")
    prefix = "" if is_spx else "GHN "
    target_basket_prefix = f"{prefix}{basket_name}".strip().lower()
    
    try:
        # Fetch column D once per sheet/day instead of once per basket type.
        sheet_cache_key = (id(ws_sheet), today_str)
        if sheet_cache_key not in _basket_seq_sheet_values:
            _basket_seq_sheet_values[sheet_cache_key] = ws_sheet.col_values(4)[1:]
        col_d_values = _basket_seq_sheet_values[sheet_cache_key]
    except Exception as e:
        print(f"⚠️ Không thể đọc Cột D từ Sheet: {e}. Mặc định bắt đầu từ STT 1.")
        return 1

    max_seq = 0
    for raw_val in col_d_values:
        val = str(raw_val).strip()
        if not val or ' ' not in val:
            continue
        
        name_part, dateseq_part = val.rsplit(' ', 1)
        if '-' not in dateseq_part:
            continue
            
        date_part, seq_part = dateseq_part.rsplit('-', 1)
        
        if date_part == today_str and seq_part.isdigit():
            if name_part.strip().lower() == target_basket_prefix:
                seq_num = int(seq_part)
                if seq_num > max_seq:
                    max_seq = seq_num

    return max_seq + 1

def get_next_basket_seq(is_spx, basket_name):
    prefix = "SPX" if is_spx else "GHN"
    today_str = datetime.now().strftime("%d/%m")
    cache_key = (prefix, basket_name, today_str)

    with _basket_seq_lock:
        if cache_key not in basket_seq_cache:
            initial_seq = get_next_basket_seq_from_sheet(is_spx, basket_name)
            basket_seq_cache[cache_key] = initial_seq
        else:
            basket_seq_cache[cache_key] += 1
        return basket_seq_cache[cache_key]

# ========================================================================================
# LOGIC CHỐT RỔ TỐI ƯU KHÔNG ĐƠ TOOL (BACKGROUND THREADING)
# ========================================================================================
def submit_basket(is_spx, basket_name):
    global countspx, counttotalspx, countghn, counttotalghn

    baskets_data = baskets_data_spx if is_spx else baskets_data_ghn
    baskets_ui = baskets_ui_spx if is_spx else baskets_ui_ghn
    tree_widget = tree if is_spx else tree2
    ws_sheet = ws_spx if is_spx else ws_ghn

    tns_in_basket = baskets_data.get(basket_name, [])

    if not tns_in_basket:
        messagebox.showwarning("Trống", f"Rổ '{basket_name}' hiện không có đơn nào mới cần chốt!")
        return

    # 1. TẠO MÃ RỔ THEO NGÀY VÀ GIỜ/PHÚT (VD: lite 10/10-1054)
    now = datetime.now()
    date_str = now.strftime("%d/%m")
    time_str = now.strftime("%H%M")
    prefix = "" if is_spx else "GHN "
    basket_id = f"{prefix}{basket_name} {date_str}-{time_str}"

    # 2. LẤY DỮ LIỆU ĐÓNG GÓI CHỐT
    rows_to_upload = []
    tns_set = set(tns_in_basket)

    for child in tree_widget.get_children():
        values = tree_widget.item(child, 'values')
        if values[0] in tns_set:
            rows_to_upload.append([values[0], values[1], values[2], basket_id])

    if not rows_to_upload:
        for tn in tns_in_basket:
            rows_to_upload.append([tn, basket_name, datetime.now().strftime("%Y-%m-%d %H:%M:%S"), basket_id])

    total_qty = len(rows_to_upload)

    # 3. GIẢI PHÓNG UI NGAY TỨC THÌ ĐỂ NGƯỜI QUÉT TIẾP TỤC THAO TÁC
    children_to_delete = []
    for child in tree_widget.get_children():
        values = tree_widget.item(child, 'values')
        if values[0] in tns_set:
            children_to_delete.append(child)
    for child in children_to_delete:
        tree_widget.delete(child)

    baskets_data[basket_name] = []
    baskets_ui[basket_name].configure(text="0")

    if is_spx:
        countspx = max(0, countspx - total_qty)
        counttotalspx += total_qty
        qtyspx.configure(text=str(countspx))
        totalqtyspx.configure(text=str(counttotalspx))
        entry_spx.focus()
    else:
        countghn = max(0, countghn - total_qty)
        counttotalghn += total_qty
        qtyghn.configure(text=str(countghn))
        totalqtyghn.configure(text=str(counttotalghn))
        entry_ghn.focus()

    # 4. CHUYỂN XỬ LÝ MẠNG VÀ IN TEM SANG THREAD NGẦM
    def _async_submit_worker(rows, tns, b_id, qty):
        print(f"🔄 [Thread Ngầm] Đang đẩy {qty} đơn thuộc rổ {b_id} lên Google Sheets...")

        def restore_failed_basket(error):
            global countspx, counttotalspx, countghn, counttotalghn
            # Keep failed orders available for an immediate retry in this run.
            for row in rows:
                tn, fn, received_at, _ = row
                if tn not in baskets_data.setdefault(basket_name, []):
                    baskets_data[basket_name].append(tn)
                    tree_widget.insert('', 'end', values=(tn, fn, received_at, ''))
                    if is_spx:
                        countspx += 1
                    else:
                        countghn += 1
            baskets_ui[basket_name].configure(text=str(len(baskets_data[basket_name])))
            if is_spx:
                counttotalspx = max(0, counttotalspx - qty)
                qtyspx.configure(text=str(countspx))
                totalqtyspx.configure(text=str(counttotalspx))
            else:
                counttotalghn = max(0, counttotalghn - qty)
                qtyghn.configure(text=str(countghn))
                totalqtyghn.configure(text=str(counttotalghn))
            messagebox.showerror(
                "Lỗi Kết Nối Mạng",
                f"Chưa xác nhận được Google Sheets đã nhận rổ [{b_id}]. "
                f"{qty} đơn vẫn nằm trong hàng đợi offline và đã được khôi phục để thử lại.\n\n{error}"
            )

        try:
            ws_sheet.append_rows(rows, value_input_option='USER_ENTERED')
            print(f"✅ [Thread Ngầm] Đã chốt thành công rổ [{b_id}] lên Google Sheets!")

        except Exception as net_err:
            print(f"🚨 [Thread Ngầm] Lỗi kết nối Google Sheets: {net_err}")
            root.after(0, lambda err=net_err: restore_failed_basket(err))
            return

        # Printer failures are separate from Google Sheets upload status.
        update_csv_after_submit(is_spx, tns, b_id)

    # Start both operations right away. Printing no longer waits for the
    # Google Sheets network request to finish.
    threading.Thread(target=print_basket_label, args=(basket_id,), kwargs={"qty": total_qty}, daemon=True).start()
    threading.Thread(target=_async_submit_worker, args=(rows_to_upload, tns_in_basket, basket_id, total_qty), daemon=True).start()

def retry_all_pending_baskets(is_spx):
    baskets_data = baskets_data_spx if is_spx else baskets_data_ghn
    active_baskets = [b for b, items in baskets_data.items() if len(items) > 0]
    if not active_baskets:
        messagebox.showinfo("Thông báo", "Không tìm thấy rổ hàng tồn đọng nào cần đồng bộ!")
        return
    for b_name in active_baskets:
        submit_basket(is_spx, b_name)

def create_basket_ui(is_spx, basket_name):
    parent_frame = baskets_frame_spx if is_spx else baskets_frame_ghn
    baskets_ui = baskets_ui_spx if is_spx else baskets_ui_ghn
    
    frame = ctk.CTkFrame(parent_frame, fg_color="#2b2b2b", corner_radius=8, border_width=1, border_color="#3d3d3d")
    frame.pack(fill="x", pady=5, padx=5)
    
    lbl_name = ctk.CTkLabel(frame, text=basket_name, font=('Arial', 13, 'bold'), text_color="#ff9f00", width=125, anchor="w")
    lbl_name.grid(row=0, column=0, padx=10, pady=6)
    
    lbl_count = ctk.CTkLabel(frame, text="0", font=('Arial', 15, 'bold'), text_color="#10b981", width=35)
    lbl_count.grid(row=0, column=1, padx=10, pady=6)
    baskets_ui[basket_name] = lbl_count 
    
    btn_submit = ctk.CTkButton(frame, text="Chốt & In 🖨️", font=('Arial', 11, 'bold'), width=85,
                               fg_color="#0284c7", hover_color="#0369a1",
                               command=lambda b=basket_name: submit_basket(is_spx, b))
    btn_submit.grid(row=0, column=2, padx=10, pady=6)

# --- XỬ LÝ LƯU HOẠT ĐỘNG QUÉT SPX ---
def save_spx(event):
    global countspx
    return_tn_spx = entry_spx.get().strip()
    if not return_tn_spx:
        return
    
    # Xóa nội dung ô nhập
    entry_spx.delete(0, 'end')
    
    # 1. KIỂM TRA TRÙNG TRÊN GOOGLE SHEETS
    try:
        count_dup_rc = return_list_spx['Return TN'].value_counts()[return_tn_spx]
        voice("duplicate_voice")
        labelloaihangspx.configure(text="Trùng Đơn (Sheets)", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return
    except KeyError:
        pass

    # 2. TRA CỨU DỮ LIỆU PHÂN LOẠI
    try:
        ft = df_source[df_source['Return TN'] == f'{return_tn_spx}']
        if ft.empty:
            voice("nodata_voice")
            labelloaihangspx.configure(text="No_data", font=("Arial", 22, 'bold'), text_color="#ef4444")
            return
            
        raw_fn = ft.iloc[0, 6]
        if pd.isna(raw_fn) or not raw_fn or str(raw_fn).strip().lower() in ['nan', 'none', '']:
            fn = "KO PHÂN LOẠI"
        else:
            fn = str(raw_fn).strip()
            
        now = datetime.now().replace(microsecond=0)
        
        # 3. KIỂM TRA TRÙNG TRONG PHIÊN LÀM VIỆC (ĐÃ QUÉT TRƯỚC ĐÓ)
        if return_tn_spx in listscanspx:
            if fn == "KO PHÂN LOẠI":
                voice("NR_voice")
            else:
                voice(fn)
            # Cảnh báo giao diện và THOÁT ra, KHÔNG lưu thêm
            labelloaihangspx.configure(text=f"{fn} (Đã quét)", font=("Arial", 22, 'bold'), text_color="#ef4444")
            return
            
        # 4. LƯU ĐƠN MỚI
        listscanspx.append(return_tn_spx)
        
        if fn == "KO PHÂN LOẠI":
            voice("NR_voice")
            labelloaihangspx.configure(text=fn, font=("Arial", 22, 'bold'), text_color="#ff9f00")
        else:
            voice(fn)
            labelloaihangspx.configure(text=fn, font=("Arial", 22, 'bold'), text_color="#10b981")
        
        basket_name = get_basket_name(fn, is_spx=True)
        if basket_name not in baskets_data_spx:
            baskets_data_spx[basket_name] = []
            create_basket_ui(is_spx=True, basket_name=basket_name)
            
        baskets_data_spx[basket_name].append(return_tn_spx)
        baskets_ui_spx[basket_name].configure(text=str(len(baskets_data_spx[basket_name])))
        
        countspx += 1
        qtyspx.configure(text=str(countspx))
        
        log_scan_csv(is_spx=True, ma_vc=return_tn_spx, phan_loai=fn, gio_nhan=str(now), ma_ro="", status="Chưa Submit")
        tree.insert('', 'end', values=(return_tn_spx, fn, str(now), ""))
        tree.see(tree.get_children()[-1])
        
    except Exception as ex:
        print(ex)
        voice("nodata_voice")
        labelloaihangspx.configure(text="No_data", font=("Arial", 22, 'bold'), text_color="#ef4444")


# --- XỬ LÝ LƯU HOẠT ĐỘNG QUÉT GHN ---
def save_ghn(event):
    global countghn
    return_tn_ghn = entry_ghn.get().strip()
    if not return_tn_ghn:
        return
    
    # Xóa nội dung ô nhập
    entry_ghn.delete(0, 'end')
    
    # 1. KIỂM TRA TRÙNG TRÊN GOOGLE SHEETS
    try:
        count_dup_rc = return_list_ghn['Return TN'].value_counts()[return_tn_ghn]
        voice("duplicate_voice")
        labelloaihangghn.configure(text="Trùng Đơn (Sheets)", font=("Arial", 22, 'bold'), text_color="#ef4444")
        return
    except KeyError:
        pass

    # 2. TRA CỨU DỮ LIỆU PHÂN LOẠI
    try:
        ft = df_source[df_source['Return TN'] == f'{return_tn_ghn}']
        if ft.empty:
            voice("nodata_voice")
            labelloaihangghn.configure(text="No_data", font=("Arial", 22, 'bold'), text_color="#ef4444")
            return
            
        raw_fn = ft.iloc[0, 6]
        if pd.isna(raw_fn) or not raw_fn or str(raw_fn).strip().lower() in ['nan', 'none', '']:
            fn = "KO PHÂN LOẠI"
        else:
            fn = str(raw_fn).strip()
            
        now = datetime.now().replace(microsecond=0)
        
        # 3. KIỂM TRA TRÙNG TRONG PHIÊN LÀM VIỆC (ĐÃ QUÉT TRƯỚC ĐÓ)
        if return_tn_ghn in listscanghn:
            if fn == "KO PHÂN LOẠI":
                voice("NR_voice")
            else:
                voice(fn)
            # Cảnh báo giao diện và THOÁT ra, KHÔNG lưu thêm
            labelloaihangghn.configure(text=f"{fn} (Đã quét)", font=("Arial", 22, 'bold'), text_color="#ef4444")
            return
            
        # 4. LƯU ĐƠN MỚI
        listscanghn.append(return_tn_ghn)
        
        if fn == "KO PHÂN LOẠI":
            voice("NR_voice")
            labelloaihangghn.configure(text=fn, font=("Arial", 22, 'bold'), text_color="#ff9f00")
        else:
            voice(fn)
            labelloaihangghn.configure(text=fn, font=("Arial", 22, 'bold'), text_color="#10b981")
        
        basket_name = get_basket_name(fn, is_spx=False)
        if basket_name not in baskets_data_ghn:
            baskets_data_ghn[basket_name] = []
            create_basket_ui(is_spx=False, basket_name=basket_name)
            
        baskets_data_ghn[basket_name].append(return_tn_ghn)
        baskets_ui_ghn[basket_name].configure(text=str(len(baskets_data_ghn[basket_name])))
        
        countghn += 1
        qtyghn.configure(text=str(countghn))
        
        log_scan_csv(is_spx=False, ma_vc=return_tn_ghn, phan_loai=fn, gio_nhan=str(now), ma_ro="", status="Chưa Submit")
        tree2.insert('', 'end', values=(return_tn_ghn, fn, str(now), ""))
        tree2.see(tree2.get_children()[-1])
        
    except Exception as ex:
        print(ex)
        voice("nodata_voice")
        labelloaihangghn.configure(text="No_data", font=("Arial", 22, 'bold'), text_color="#ef4444")

# ========================================================================================
# GIAO DIỆN CHÍNH
# ========================================================================================
ctk.set_appearance_mode("dark")  
ctk.set_default_color_theme("blue")  
root = ctk.CTk()
root.title("Return Reverse Voice v3.3 - Dò STT Rổ Tự Động Tối Ưu Tốc Độ")
root.geometry("1350x760") 

style = ttk.Style()
style.theme_use('clam')  
style.configure("Treeview", background="#333333", foreground="white", fieldbackground="#333333", borderwidth=0, font=('Arial', 12)) 
style.configure("Treeview.Heading", background="#ff9f00", foreground="white", font=('Arial', 13, 'bold')) 
style.map('Treeview', background=[('selected', '#144870')], foreground=[('selected', 'white')])

# THANH TRÊN CÙNG
top_bar = ctk.CTkFrame(root, fg_color="#1f1f1f", corner_radius=8)
top_bar.pack(fill="x", padx=15, pady=(10, 0))

label_source_status = ctk.CTkLabel(
    top_bar,
    text=f"📄 Source Data: {len(df_source)} dòng | Cập nhật lúc {datetime.now().strftime('%H:%M:%S')}",
    font=('Arial', 12), text_color="#9ca3af")
label_source_status.pack(side="left", padx=15, pady=8)

btn_reload_source = ctk.CTkButton(
    top_bar, text="🔄 Nạp Lại Source Data", font=('Arial', 12, 'bold'),
    width=190, height=30, fg_color="#15803d", hover_color="#166534",
    command=reload_source_data)
btn_reload_source.pack(side="right", padx=15, pady=8)

tabview = ctk.CTkTabview(root)
tabview.pack(fill="both", expand=True, padx=15, pady=5)

tabview.add("SPX")
tabview.add("GHN")

# ================= TAB 1: SPX =================
tab1 = tabview.tab("SPX")
tab1.grid_columnconfigure(0, weight=6)
tab1.grid_columnconfigure(1, weight=4)
tab1.grid_rowconfigure(0, weight=1)

main_frame_spx = ctk.CTkFrame(tab1, fg_color="transparent")
main_frame_spx.grid(row=0, column=0, sticky="nsew", padx=(0,10), pady=5)

baskets_frame_spx = ctk.CTkScrollableFrame(tab1, label_text="HỆ THỐNG RỔ PHÂN LOẠI SPX", label_font=('Arial', 14, 'bold'), label_text_color="#ff9f00")
baskets_frame_spx.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

main_frame_spx.grid_columnconfigure(0, weight=1)
main_frame_spx.grid_columnconfigure(1, weight=2)
main_frame_spx.grid_rowconfigure(4, weight=1)

ctk.CTkLabel(main_frame_spx, text="Scan Return TN:", font=('Arial', 14, 'bold')).grid(row=0, column=0, padx=10, pady=10, sticky="e")
entry_spx = ctk.CTkEntry(main_frame_spx, width=280, font=('Arial', 14))
entry_spx.grid(row=0, column=1, padx=10, pady=10, sticky="w")
entry_spx.bind("<Return>", save_spx)

ctk.CTkLabel(main_frame_spx, text='Loại Hàng:', font=('Arial', 14, 'bold')).grid(row=1, column=0, padx=10, pady=5, sticky="e")
labelloaihangspx = ctk.CTkLabel(main_frame_spx, text='No_data', font=('Arial', 24, 'bold'), text_color="#10b981")
labelloaihangspx.grid(row=1, column=1, padx=10, pady=5, sticky="w")

ctk.CTkLabel(main_frame_spx, text='Số đơn chưa chốt:', font=('Arial', 13)).grid(row=2, column=0, padx=10, pady=5, sticky="e")
qtyspx = ctk.CTkLabel(main_frame_spx, text=countspx, font=('Arial', 14, 'bold'), text_color="#f59e0b")
qtyspx.grid(row=2, column=1, padx=10, pady=5, sticky="w")

ctk.CTkLabel(main_frame_spx, text='Số đơn đã chốt rổ:', font=('Arial', 13)).grid(row=3, column=0, padx=10, pady=5, sticky="e")
totalqtyspx = ctk.CTkLabel(main_frame_spx, text=counttotalspx, font=('Arial', 14, 'bold'), text_color="#3b82f6")
totalqtyspx.grid(row=3, column=1, padx=10, pady=5, sticky="w")

columns_tree = ("Mã Vận Chuyển", "Phân Loại", "Giờ Nhận", "Mã Rổ")
tree_container_spx = ctk.CTkFrame(main_frame_spx, fg_color="transparent")
tree_container_spx.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
tree_container_spx.grid_columnconfigure(0, weight=1)
tree_container_spx.grid_rowconfigure(0, weight=1)

tree = ttk.Treeview(tree_container_spx, columns=columns_tree, show="headings", style="Treeview")
for col in columns_tree:
    tree.heading(col, text=col)
tree.column("Mã Vận Chuyển", width=180, anchor="center")
tree.column("Phân Loại", width=110, anchor="center")
tree.column("Giờ Nhận", width=140, anchor="center")
tree.column("Mã Rổ", width=170, anchor="center")
tree.pack(side="left", fill="both", expand=True)

v_scrollbar_spx = ttk.Scrollbar(tree_container_spx, orient="vertical", command=tree.yview)
tree.configure(yscrollcommand=v_scrollbar_spx.set)
v_scrollbar_spx.pack(side="right", fill="y")

button_sync_spx = ctk.CTkButton(main_frame_spx, text="THỬ LẠI / ĐỒNG BỘ CÁC ĐƠN LỖI MẠNG LÊN GOOGLE SHEET", font=('Arial', 13, 'bold'),
                               height=38, fg_color="#1e3a8a", hover_color="#1e40af", command=lambda: retry_all_pending_baskets(is_spx=True))
button_sync_spx.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

# ================= TAB 2: GHN =================
tab2 = tabview.tab("GHN")
tab2.grid_columnconfigure(0, weight=6)
tab2.grid_columnconfigure(1, weight=4)
tab2.grid_rowconfigure(0, weight=1)

main_frame_ghn = ctk.CTkFrame(tab2, fg_color="transparent")
main_frame_ghn.grid(row=0, column=0, sticky="nsew", padx=(0,10), pady=5)

baskets_frame_ghn = ctk.CTkScrollableFrame(tab2, label_text="HỆ THỐNG RỔ PHÂN LOẠI GHN", label_font=('Arial', 14, 'bold'), label_text_color="#ff9f00")
baskets_frame_ghn.grid(row=0, column=1, sticky="nsew", padx=5, pady=5)

main_frame_ghn.grid_columnconfigure(0, weight=1)
main_frame_ghn.grid_columnconfigure(1, weight=2)
main_frame_ghn.grid_rowconfigure(4, weight=1)

ctk.CTkLabel(main_frame_ghn, text="Scan Return TN:", font=('Arial', 14, 'bold')).grid(row=0, column=0, padx=10, pady=10, sticky="e")
entry_ghn = ctk.CTkEntry(main_frame_ghn, width=280, font=('Arial', 14))
entry_ghn.grid(row=0, column=1, padx=10, pady=10, sticky="w")
entry_ghn.bind("<Return>", save_ghn)

ctk.CTkLabel(main_frame_ghn, text='Loại Hàng:', font=('Arial', 14, 'bold')).grid(row=1, column=0, padx=10, pady=5, sticky="e")
labelloaihangghn = ctk.CTkLabel(main_frame_ghn, text='No_data', font=("Arial", 22, 'bold'), text_color="#10b981")
labelloaihangghn.grid(row=1, column=1, padx=10, pady=5, sticky="w")

ctk.CTkLabel(main_frame_ghn, text='Số đơn chưa chốt:', font=('Arial', 13)).grid(row=2, column=0, padx=10, pady=5, sticky="e")
qtyghn = ctk.CTkLabel(main_frame_ghn, text=countghn, font=('Arial', 14, 'bold'), text_color="#f59e0b")
qtyghn.grid(row=2, column=1, padx=10, pady=5, sticky="w")

ctk.CTkLabel(main_frame_ghn, text='Số đơn đã chốt rổ:', font=('Arial', 13)).grid(row=3, column=0, padx=10, pady=5, sticky="e")
totalqtyghn = ctk.CTkLabel(main_frame_ghn, text=counttotalghn, font=('Arial', 14, 'bold'), text_color="#3b82f6")
totalqtyghn.grid(row=3, column=1, padx=10, pady=5, sticky="w")

tree_container_ghn = ctk.CTkFrame(main_frame_ghn, fg_color="transparent")
tree_container_ghn.grid(row=4, column=0, columnspan=2, padx=10, pady=10, sticky="nsew")
tree_container_ghn.grid_columnconfigure(0, weight=1)
tree_container_ghn.grid_rowconfigure(0, weight=1)

tree2 = ttk.Treeview(tree_container_ghn, columns=columns_tree, show="headings", style="Treeview")
for col in columns_tree:
    tree2.heading(col, text=col)
tree2.column("Mã Vận Chuyển", width=180, anchor="center")
tree2.column("Phân Loại", width=110, anchor="center")
tree2.column("Giờ Nhận", width=140, anchor="center")
tree2.column("Mã Rổ", width=170, anchor="center")
tree2.pack(side="left", fill="both", expand=True)

v_scrollbar_ghn = ttk.Scrollbar(tree_container_ghn, orient="vertical", command=tree2.yview)
tree2.configure(yscrollcommand=v_scrollbar_ghn.set)
v_scrollbar_ghn.pack(side="right", fill="y")

button_sync_ghn = ctk.CTkButton(main_frame_ghn, text="THỬ LẠI / ĐỒNG BỘ CÁC ĐƠN LỖI MẠNG LÊN GOOGLE SHEET", font=('Arial', 13, 'bold'),
                               height=38, fg_color="#1e3a8a", hover_color="#1e40af", command=lambda: retry_all_pending_baskets(is_spx=False))
button_sync_ghn.grid(row=5, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

# ================= KHỞI TẠO RỔ & KHÔI PHỤC DỮ LIỆU =================
for b_name in PREDEFINED_BASKETS:
    create_basket_ui(is_spx=True, basket_name=b_name)
    create_basket_ui(is_spx=False, basket_name=b_name)

load_offline_backup(is_spx=True)
load_offline_backup(is_spx=False)

status = ctk.CTkLabel(root, text="VNBD1", font=('Arial', 11))
status.pack(side=ctk.BOTTOM, fill=ctk.X)

# Tkinter widgets must only be touched from the UI thread. Set initial focus
# once instead of polling focus from a background thread.
root.after(250, entry_spx.focus_set)

if not HAS_QRCODE:
    root.after(500, lambda: messagebox.showwarning(
        "Thiếu thư viện mã QR",
        "Chưa cài thư viện 'qrcode' nên tem in ra sẽ KHÔNG có mã QR!\n\n"
        "Cách khắc phục: Mở CMD gõ: pip install qrcode"
    ))

root.mainloop()
