# ===================================================================================
#                       APLIKASI SETTING PROMO MASSAL MARKETPLACE
# ===================================================================================

import tkinter as tk
from tkinter import ttk, filedialog, scrolledtext, messagebox
import pandas as pd
import threading
import warnings
import os
import webbrowser
import sys
import queue
import ctypes

warnings.filterwarnings("ignore", category=UserWarning, module="openpyxl")

def resource_path(relative_path):
    """ Dapatkan path absolut ke resource, bekerja untuk dev dan untuk PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

class PromoAppFinal:
    def __init__(self, root):
        self.root = root
        self.root.title("Setting Promo Massal Marketplace")
        self.root.geometry("900x750")
        self.root.configure(bg="#f0f0f0")

        try:
            myappid = 'fahmi.myproduct.subproduct.version' 
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except Exception:
            pass

        try:
            icon_path = resource_path("logo.ico") 
            self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Peringatan: Tidak dapat memuat ikon - {e}")

        self.style = ttk.Style(self.root)
        self.style.theme_use('clam')
        self.configure_styles()

        self.MIN_PRICE_THRESHOLD = 1000
        self.MAX_DISCOUNT_PERCENTAGE = 0.90

        self.file_paths = {
            "promo_internal": tk.StringVar(), "db_master": tk.StringVar(),
            "db_shopee": None, "db_tiktok": None, "template_shopee": tk.StringVar(),
            "template_tiktok1": tk.StringVar(), "template_tiktok2": tk.StringVar(),
        }
        
        self.log_queue = queue.Queue()
        self.has_errors = False
        self.create_widgets()
        self.process_log_queue()

    def configure_styles(self):
        BG_COLOR, FRAME_COLOR, TEXT_COLOR = "#f0f0f0", "#ffffff", "#333333"
        ACCENT_COLOR, ACCENT_HOVER_COLOR, LABEL_HEADER_COLOR = "#0078D7", "#005a9e", "#0078D7"
        PLACEHOLDER_BG, SELECTED_BG = "#e9e9e9", "#dff0d8"

        self.style.configure('TFrame', background=BG_COLOR)
        self.style.configure('TLabelFrame', background=FRAME_COLOR, borderwidth=1, relief="solid", bordercolor="#d1d1d1")
        self.style.configure('TLabelFrame.Label', background=FRAME_COLOR, foreground=LABEL_HEADER_COLOR, font=('Segoe UI', 11, 'bold'))
        self.style.configure('TLabel', background=FRAME_COLOR, foreground=TEXT_COLOR, font=('Segoe UI', 9))
        self.style.configure('Header.TLabel', background=BG_COLOR, foreground=TEXT_COLOR, font=('Segoe UI', 16, 'bold'))
        self.style.configure('Link.TLabel', background=BG_COLOR, foreground=ACCENT_COLOR, font=('Segoe UI', 9, 'underline'))
        self.style.configure('Placeholder.TLabel', background=PLACEHOLDER_BG, foreground="grey", relief="sunken")
        self.style.configure('Selected.Placeholder.TLabel', background=SELECTED_BG, foreground="black", relief="sunken")
        self.style.configure('Accent.TButton', background=ACCENT_COLOR, foreground='white', font=('Segoe UI', 12, 'bold'), borderwidth=0)
        self.style.map('Accent.TButton', background=[('active', ACCENT_HOVER_COLOR)])
        self.style.configure('TEntry', fieldbackground='#fdfdfd', bordercolor="#cccccc", lightcolor="#cccccc", darkcolor="#cccccc")

    def create_widgets(self):
        # === PERBAIKAN: Mengatur ulang struktur layout ===
        
        # 1. Buat footer terlebih dahulu dan letakkan di bagian bawah
        footer_frame = ttk.Frame(self.root, style='TFrame', padding=(20, 10))
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        ttk.Separator(footer_frame).pack(fill=tk.X, pady=(0, 10))
        credit_label = ttk.Label(footer_frame, text="@khairudinfahmi", style='Link.TLabel', cursor="hand2")
        credit_label.pack()
        credit_label.bind("<Button-1>", lambda e: self.open_link("https://www.instagram.com/khairudinfahmi"))
        
        # 2. Buat main_frame untuk mengisi sisa ruang
        main_frame = ttk.Frame(self.root, padding=20, style='TFrame')
        main_frame.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Selamat Datang, Jangan Lupa Ngopi Dulu Lur!", style='Header.TLabel').pack(pady=(0, 20))

        file_frame = ttk.LabelFrame(main_frame, text="Langkah 1: Pilih Semua File yang Dibutuhkan", padding=(15, 10))
        file_frame.pack(fill=tk.X, pady=10, ipady=10)
        
        file_selection_rows = [
            ("File List Promo (Offline)", "promo_internal"), ("File Database Master (Online)", "db_master"),
            ("File Database Produk Shopee", "db_shopee"), ("File Database Produk TikTok", "db_tiktok"),
            ("File Template Promo Shopee", "template_shopee"), ("File Template Promo TikTok (Metode 1)", "template_tiktok1"),
            ("File Template Promo TikTok (Metode 2)", "template_tiktok2"),
        ]
        
        for i, (label_text, key) in enumerate(file_selection_rows):
            is_multiple = not isinstance(self.file_paths[key], tk.StringVar)
            command = self.select_multiple_files if is_multiple else self.select_file
            
            ttk.Label(file_frame, text=label_text).grid(row=i, column=0, sticky=tk.W, padx=5, pady=8)
            
            if not is_multiple:
                ttk.Entry(file_frame, textvariable=self.file_paths[key], width=60, state='readonly').grid(row=i, column=1, sticky=tk.EW, padx=5)
            else:
                label = ttk.Label(file_frame, text="Belum dipilih...", style='Placeholder.TLabel', anchor='w', padding=(2,2))
                label.grid(row=i, column=1, sticky=tk.EW, padx=5)
                self.file_paths[f"{key}_label"] = label
            
            ttk.Button(file_frame, text="Pilih File...", command=lambda k=key, c=command: c(k)).grid(row=i, column=2, padx=5)
        
        file_frame.columnconfigure(1, weight=1)

        action_frame = ttk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(20, 10))
        self.process_button = ttk.Button(action_frame, text="PROSES, VALIDASI & BUAT LAPORAN", style='Accent.TButton', command=self.start_processing)
        self.process_button.pack(fill=tk.X, ipady=8)

        log_frame = ttk.LabelFrame(main_frame, text="Log Proses", padding=10)
        log_frame.pack(fill=tk.BOTH, expand=True)
        self.log_text = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=15, state='disabled', font=("Consolas", 9), relief=tk.FLAT, bd=0, bg="#ffffff")
        self.log_text.pack(fill=tk.BOTH, expand=True)

    def process_log_queue(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self.log_text.config(state='normal')
                self.log_text.insert(tk.END, msg + "\n")
                self.log_text.see(tk.END)
                self.log_text.config(state='disabled')
        except queue.Empty:
            pass
        self.root.after(100, self.process_log_queue)

    def log(self, message):
        self.log_queue.put(message)

    def open_link(self, url):
        webbrowser.open_new(url)

    def select_file(self, key):
        path = filedialog.askopenfilename(title=f"Pilih file '{key}'", filetypes=[("Excel", "*.xlsx *.xls")])
        if path: self.file_paths[key].set(path)

    def select_multiple_files(self, key):
        paths = filedialog.askopenfilenames(title=f"Pilih file untuk '{key}'", filetypes=[("Excel", "*.xlsx *.xls")])
        if paths:
            self.file_paths[key] = paths
            self.file_paths[f"{key}_label"].config(text=f"{len(paths)} file dipilih", style='Selected.Placeholder.TLabel')

    def start_processing(self):
        is_ready = True
        for key, value in self.file_paths.items():
            if not isinstance(key, str) or key.endswith('_label'): continue
            if (isinstance(value, tk.StringVar) and not value.get()) or (not isinstance(value, tk.StringVar) and not value):
                is_ready = False; messagebox.showerror("Input Tidak Lengkap", f"Harap pilih file untuk:\n'{key.replace('_', ' ').title()}'"); return
        
        self.process_button.config(state='disabled')
        self.has_errors = False
        self.log_text.config(state='normal'); self.log_text.delete('1.0', tk.END); self.log_text.config(state='disabled')
        self.log("MEMULAI PROSES, VALIDASI & AUDIT...")
        threading.Thread(target=self.run_process_logic, daemon=True).start()

    def clean_value(self, value):
        if pd.isna(value): return ""
        return str(value).strip().upper().replace("O", "0").split('.')[0]
    
    def find_col_name(self, df, possible_names, file_type_for_error):
        df_cols_map = {str(c).replace('\xa0', ' ').strip().lower(): str(c) for c in df.columns}
        for name in possible_names:
            clean_name = name.lower().strip()
            if clean_name in df_cols_map: return df_cols_map[clean_name]
        raise ValueError(f"Di file '{file_type_for_error}', tidak bisa menemukan kolom '{possible_names[0]}'. Kolom yang ada: {list(df.columns)}")

    def clean_price_series(self, series):
        cleaned_series = series.astype(str).str.replace(r'\D', '', regex=True)
        return pd.to_numeric(cleaned_series, errors='coerce').fillna(0)

    def run_process_logic(self):
        try:
            self.log("\n[1] Membaca & Mempersiapkan Data Promo (Offline)...")
            promo_df_raw = pd.read_excel(self.file_paths["promo_internal"].get(), dtype=str)
            
            kode_barang_col = self.find_col_name(promo_df_raw, ['Kode Barang'], "Promo")
            harga_jual_col = self.find_col_name(promo_df_raw, ['Harga Jual'], "Promo")
            harga_promo_col = self.find_col_name(promo_df_raw, ['Harga Diskon', 'HARGA PROMO'], "Promo")
            promo_data = promo_df_raw[[kode_barang_col, harga_jual_col, harga_promo_col]].copy()
            promo_data.columns = ['sku_asli', 'harga_jual_offline', 'harga_promo_offline']
            promo_data['sku'] = promo_data['sku_asli'].apply(self.clean_value)
            total_promo_input = len(promo_data)

            self.log("[2] Membaca & Mempersiapkan Data Master (Online)...")
            master_df = pd.read_excel(self.file_paths["db_master"].get(), dtype=str)
            master_data = master_df[[self.find_col_name(master_df, ['KodeBarang'], "DB Master"), self.find_col_name(master_df, ['HargaJual'], "DB Master")]].copy()
            master_data.columns = ['sku', 'harga_jual_online']
            master_data['sku'] = master_data['sku'].apply(self.clean_value)
            
            self.log("\n[3] Validasi data, pembersihan harga, & kalkulasi harga final...")
            
            self.log("-> Memeriksa dan membersihkan SKU duplikat dari file input...")
            promo_count_before = len(promo_data)
            promo_data.drop_duplicates(subset=['sku'], keep='first', inplace=True)
            promo_duplicates_removed = promo_count_before - len(promo_data)
            if promo_duplicates_removed > 0: self.log(f"   - File Promo: Dihapus {promo_duplicates_removed} SKU duplikat.")

            master_count_before = len(master_data)
            master_data.drop_duplicates(subset=['sku'], keep='first', inplace=True)
            if master_count_before > len(master_data): self.log(f"   - File DB Master: Dihapus {master_count_before - len(master_data)} SKU duplikat.")

            promo_data['harga_jual_offline'] = self.clean_price_series(promo_data['harga_jual_offline'])
            promo_data['harga_promo_offline'] = self.clean_price_series(promo_data['harga_promo_offline'])
            master_data['harga_jual_online'] = self.clean_price_series(master_data['harga_jual_online'])

            final_df = pd.merge(promo_data, master_data, on='sku', how='inner')
            self.log(f"-> Ditemukan {len(final_df)} produk dengan SKU unik yang cocok antara promo list dan DB master.")

            if not final_df.empty:
                final_df['Potongan_Nominal'] = final_df['harga_jual_offline'] - final_df['harga_promo_offline']
                final_df['Harga_Diskon_Final'] = final_df['harga_jual_online'] - final_df['Potongan_Nominal']
                final_df['Persentase_Diskon'] = (final_df['Potongan_Nominal'] / final_df['harga_jual_online']).fillna(0)
                final_df.rename(columns={'sku': 'promo_sku_cleaned'}, inplace=True)

                self.log("-> Contoh hasil kalkulasi:")
                self.log(final_df[['promo_sku_cleaned', 'harga_jual_online', 'Harga_Diskon_Final']].head().to_string())

                summary_shopee = self.process_platform("Shopee", final_df.copy())
                summary_tiktok = self.process_platform("TikTok", final_df.copy())
                
                summary_data = {
                    'Metrik': [
                        'Total SKU di File Promo Awal', 'SKU Duplikat Dihapus', '',
                        '--- HASIL AUDIT SHOPEE ---', 'Produk Berhasil Diproses (Aman)', 'Produk Tidak Ditemukan di DB', 'Produk Perlu Tinjauan (Peringatan Harga)', '',
                        '--- HASIL AUDIT TIKTOK ---', 'Produk Berhasil Diproses (Aman)', 'Produk Tidak Ditemukan di DB', 'Produk Perlu Tinjauan (Peringatan Harga)'
                    ],
                    'Jumlah': [
                        total_promo_input, promo_duplicates_removed, '',
                        '', summary_shopee.get('safe', 0), summary_shopee.get('not_found', 0), summary_shopee.get('warning', 0), '',
                        '', summary_tiktok.get('safe', 0), summary_tiktok.get('not_found', 0), summary_tiktok.get('warning', 0)
                    ]
                }
                summary_df = pd.DataFrame(summary_data)
                
                with pd.ExcelWriter('RINGKASAN_PROSES_KESELURUHAN.xlsx', engine='openpyxl') as writer:
                    summary_df.to_excel(writer, sheet_name='Ringkasan Eksekutif', index=False)
                self.log("\n-> ✅ Laporan 'RINGKASAN_PROSES_KESELURUHAN.xlsx' telah dibuat.")

            else:
                self.log("\n[PERINGATAN] Tidak ada produk yang cocok ditemukan untuk diproses.")

            if not self.has_errors:
                self.log("\n====================\n✅ SEMUA PROSES, VALIDASI & AUDIT SELESAI ✅\n====================")
                messagebox.showinfo("Selesai", "Semua proses telah berhasil diselesaikan!\n\nLaporan Ringkasan utama ada di file:\n'RINGKASAN_PROSES_KESELURUHAN.xlsx'")
            else:
                self.log("\n====================\n⚠️ PROSES SELESAI DENGAN ERROR ⚠️\n====================")
                messagebox.showwarning("Selesai dengan Peringatan", "Proses selesai, namun ditemukan beberapa error. Silakan periksa log.")
        except Exception as e:
            self.log(f"❌ ERROR FATAL: {type(e).__name__}: {e}"); import traceback; self.log(traceback.format_exc()); messagebox.showerror("Error", f"Terjadi Error Fatal:\n{e}")
        finally:
            self.process_button.config(state='normal')

    def create_audit_report(self, filename, summary_df, safe_df, warning_df, not_found_df):
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                summary_df.to_excel(writer, sheet_name='Ringkasan Laporan', index=False)
                
                found_cols = {'promo_sku_cleaned': 'SKU (Sudah Dibersihkan)', 'nama_produk_platform': 'Nama Produk di Platform', 'Harga_Diskon_Final': 'Harga Promo Final'}
                report_found = safe_df[found_cols.keys()].copy()
                report_found.columns = found_cols.values()
                report_found.to_excel(writer, sheet_name='Produk Ditemukan (Aman)', index=False)
                
                if not warning_df.empty:
                    warning_cols = {'promo_sku_cleaned': 'SKU', 'nama_produk_platform': 'Nama Produk', 'harga_jual_online': 'Harga Asli', 'Harga_Diskon_Final': 'Harga Promo Final', 'Persentase_Diskon': 'Diskon', 'alasan_peringatan': 'Alasan Peringatan'}
                    report_warning = warning_df[warning_cols.keys()].copy()
                    report_warning.columns = warning_cols.values()
                    report_warning['Diskon'] = (report_warning['Diskon'] * 100).map('{:.2f}%'.format)
                    report_warning.to_excel(writer, sheet_name='Peringatan Harga (Perlu Tinjauan)', index=False)
                
                if not not_found_df.empty:
                    not_found_cols = {'promo_sku_cleaned': 'SKU Tidak Ditemukan (Dibersihkan)', 'sku_asli': 'SKU Asli (Dari File Promo)', 'harga_jual_offline': 'Harga Jual Asli (Offline)'}
                    report_not_found = not_found_df[not_found_cols.keys()].copy()
                    report_not_found.columns = not_found_cols.values()
                    report_not_found.to_excel(writer, sheet_name='Produk Tidak Ditemukan', index=False)
                    
            self.log(f"-> ✅ Laporan Audit Lengkap '{filename}' telah dibuat.")
        except Exception as e:
            self.log(f"-> ❌ Gagal membuat Laporan Audit '{filename}'. Error: {e}")
            self.has_errors = True

    def process_platform(self, platform_name, promo_data):
        self.log(f"\n[{platform_name.upper()}] Memulai proses & audit...")
        summary = {}
        try:
            db_paths = self.file_paths[f"db_{platform_name.lower()}"]
            if platform_name == 'Shopee':
                db_df = pd.concat([pd.read_excel(f, dtype=str).fillna('') for f in db_paths], ignore_index=True)
                db_cols = {'id_produk': self.find_col_name(db_df, ['et_title_product_id', 'ID Produk'], "DB Shopee"), 'id_variasi': self.find_col_name(db_df, ['et_title_variation_id', 'ID Variasi'], "DB Shopee"), 'sku_variasi': self.find_col_name(db_df, ['et_title_variation_sku', 'SKU'], "DB Shopee"), 'nama_produk': self.find_col_name(db_df, ['et_title_product_name'], "DB Shopee")}
                db_df['lookup_sku_cleaned'] = db_df[db_cols['sku_variasi']].apply(self.clean_value)
            else: # TikTok
                db_list = []
                for f in db_paths:
                    try: pd.read_excel(f, sheet_name="Template", header=None, dtype=str); sheet = "Template"
                    except: sheet = 0
                    peek_df = pd.read_excel(f, sheet_name=sheet, header=None, dtype=str)
                    h_idx = next((i for i, row in peek_df.head(10).iterrows() if 'product_id' in str(row.values).lower() or 'seller_sku' in str(row.values).lower()), 0)
                    db_list.append(pd.read_excel(f, sheet_name=sheet, header=h_idx, dtype=str).fillna(''))
                db_df = pd.concat(db_list, ignore_index=True)
                db_cols = {'id_produk': self.find_col_name(db_df, ['product_id'], "DB TikTok"), 'id_sku': self.find_col_name(db_df, ['sku_id'], "DB TikTok"), 'sku_penjual': self.find_col_name(db_df, ['seller_sku'], "DB TikTok"), 'nama_produk': self.find_col_name(db_df, ['product_name'], "DB TikTok")}
                db_df['lookup_sku_cleaned'] = db_df[db_cols['sku_penjual']].apply(self.clean_value)

            merged_df = pd.merge(promo_data, db_df, left_on='promo_sku_cleaned', right_on='lookup_sku_cleaned', how='left')
            merged_df.rename(columns={db_cols['nama_produk']: 'nama_produk_platform'}, inplace=True)
            
            found_mask = merged_df[db_cols['id_produk']].notna()
            found_df = merged_df[found_mask].copy()
            not_found_df = merged_df[~found_mask].copy()
            self.log(f"-> Ditemukan: {len(found_df)} produk. Tidak Ditemukan: {len(not_found_df)} produk.")

            price_too_low_mask = found_df['Harga_Diskon_Final'] < self.MIN_PRICE_THRESHOLD
            discount_too_high_mask = found_df['Persentase_Diskon'] > self.MAX_DISCOUNT_PERCENTAGE
            warning_mask = price_too_low_mask | discount_too_high_mask
            warning_df = found_df[warning_mask].copy()
            safe_df = found_df[~warning_mask].copy()
            
            def get_warning_reason(row):
                reasons = []
                if row['Harga_Diskon_Final'] < self.MIN_PRICE_THRESHOLD: reasons.append(f"Harga di bawah Rp {self.MIN_PRICE_THRESHOLD}")
                if row['Persentase_Diskon'] > self.MAX_DISCOUNT_PERCENTAGE: reasons.append(f"Diskon di atas {self.MAX_DISCOUNT_PERCENTAGE*100:.0f}%")
                return ', '.join(reasons)
            
            if not warning_df.empty: warning_df['alasan_peringatan'] = warning_df.apply(get_warning_reason, axis=1)
            self.log(f"-> Validasi Cerdas: {len(safe_df)} produk aman, {len(warning_df)} produk perlu tinjauan.")

            summary_data = {'Metrik': ['Produk Berhasil Diproses (Aman)', 'Produk Tidak Ditemukan di DB', 'Produk Perlu Tinjauan (Peringatan Harga)'], 'Jumlah': [len(safe_df), len(not_found_df), len(warning_df)]}
            self.create_audit_report(f'LAPORAN_AUDIT_{platform_name.upper()}.xlsx', pd.DataFrame(summary_data), safe_df, warning_df, not_found_df)
            
            summary = {'safe': len(safe_df), 'not_found': len(not_found_df), 'warning': len(warning_df)}

            if not safe_df.empty:
                if platform_name == 'Shopee':
                    template_df = pd.read_excel(self.file_paths["template_shopee"].get())
                    template_cols = {'id_produk': self.find_col_name(template_df, ['ID Produk', 'Kode Produk'], "Tmpl Shopee"), 'id_variasi': self.find_col_name(template_df, ['ID Variasi', 'Kode Variasi'], "Tmpl Shopee"), 'harga_diskon': self.find_col_name(template_df, ['Harga Diskon'], "Tmpl Shopee")}
                    output_df = pd.DataFrame({'id_produk': safe_df[db_cols['id_produk']], 'id_variasi': safe_df[db_cols['id_variasi']], 'harga_diskon': safe_df['Harga_Diskon_Final']})
                    output_df.rename(columns={'id_produk': template_cols['id_produk'], 'id_variasi': template_cols['id_variasi'], 'harga_diskon': template_cols['harga_diskon']}, inplace=True)
                    output_df[template_cols['harga_diskon']] = pd.to_numeric(output_df[template_cols['harga_diskon']], errors='coerce').round(0).astype('Int64')
                    for col in template_df.columns:
                        if col not in output_df.columns: output_df[col] = ''
                    output_df[template_df.columns].to_excel(f"HASIL_PROMO_{platform_name.upper()}.xlsx", index=False)
                    self.log(f"-> ✅ File 'HASIL_PROMO_{platform_name.upper()}.xlsx' telah dibuat.")
                else: # TikTok
                    template_m1_df = pd.read_excel(self.file_paths['template_tiktok1'].get())
                    m1_cols = {'id_produk': self.find_col_name(template_m1_df, ['Product_id (wajib) diisi', 'Product_id (wajib)'], "Tmpl TikTok M1"), 'id_sku': self.find_col_name(template_m1_df, ['SKU_id (wajib) diisi', 'SKU_id (wajib)'], "Tmpl TikTok M1"), 'harga_diskon': self.find_col_name(template_m1_df, ['Harga Penawaran (wajib) diisi', 'Harga Penawaran (wajib)'], "Tmpl TikTok M1")}
                    output_m1 = pd.DataFrame({'id_produk': safe_df[db_cols['id_produk']], 'id_sku': safe_df[db_cols['id_sku']], 'harga_diskon': safe_df['Harga_Diskon_Final']})
                    output_m1.rename(columns={'id_produk': m1_cols['id_produk'], 'id_sku': m1_cols['id_sku'], 'harga_diskon': m1_cols['harga_diskon']}, inplace=True)
                    output_m1[m1_cols['harga_diskon']] = pd.to_numeric(output_m1[m1_cols['harga_diskon']], errors='coerce').round(0).astype('Int64')
                    for col in template_m1_df.columns:
                        if col not in output_m1.columns: output_m1[col] = ''
                    output_m1[template_m1_df.columns].to_excel("HASIL_PROMO_TIKTOK_METODE1.xlsx", index=False)
                    self.log("   -> ✅ File 'HASIL_PROMO_TIKTOK_METODE1.xlsx' telah dibuat.")

                    unique_safe_df = safe_df.drop_duplicates(subset=[db_cols['id_produk']], keep='first').copy()
                    template_m2_df = pd.read_excel(self.file_paths['template_tiktok2'].get())
                    m2_cols = {'id_produk': self.find_col_name(template_m2_df, ['Product_id (wajib) diisi', 'Product_id (wajib)'], "Tmpl TikTok M2"), 'harga_diskon': self.find_col_name(template_m2_df, ['Harga Penawaran (wajib) diisi', 'Harga Penawaran (wajib)'], "Tmpl TikTok M2")}
                    output_m2 = pd.DataFrame({'id_produk': unique_safe_df[db_cols['id_produk']], 'harga_diskon': unique_safe_df['Harga_Diskon_Final']})
                    output_m2.rename(columns={'id_produk': m2_cols['id_produk'], 'harga_diskon': m2_cols['harga_diskon']}, inplace=True)
                    output_m2[m2_cols['harga_diskon']] = pd.to_numeric(output_m2[m2_cols['harga_diskon']], errors='coerce').round(0).astype('Int64')
                    for col in template_m2_df.columns:
                        if col not in output_m2.columns: output_m2[col] = ''
                    output_m2[template_m2_df.columns].to_excel("HASIL_PROMO_TIKTOK_METODE2.xlsx", index=False)
                    self.log("   -> ✅ File 'HASIL_PROMO_TIKTOK_METODE2.xlsx' telah dibuat.")
        except Exception as e:
            self.log(f"-> ❌ ERROR {platform_name}: {e}"); self.has_errors = True
        
        return summary

if __name__ == "__main__":
    root = tk.Tk()
    app = PromoAppFinal(root)
    root.mainloop()

