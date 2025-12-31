import streamlit as st
import pandas as pd
import numpy as np
import io
import traceback
from datetime import datetime
from typing import Tuple

# --- 1. Konfigurasi Halaman ---
st.set_page_config(page_title="E-Commerce Settlement Summary", layout="wide")
st.title("Marketplace Reconciliation")

# --- 2. Definisi Kolom Wajib Data Transaksi Valid dan Invalid ---
TIKTOK_REQUIRED = [
    'order_id', 'prev_order_id', 'type', 'order_settled_time', 'warehouse_name', 
    'shop_name', 'brand', 'product_quantity', 'product_total_amount', 'product_price_difference_amount', 
    'calculated_seller_discounts', 'calculated_fee', 'calculated_discount', 'calculated_shipping_fee', 
    'calculated_sfp_fee', 'calculated_affiliate_commission', 'calculated_refund', 'actual_settlement_amount', 
    'calculated_settlement_amount', 'selisih'
    ]

SHOPEE_REQUIRED = [
    'warehouse_name', 'shop_name', 'product', 'quantity', 'price',
    'discount_price_difference', 'voucher_seller', 'discount_seller',
    'total_discount', 'shipping_fee', 'affiliate_commission_fee',
    'commission_fee', 'service_fee', 'processing_fee',
    'total_comission_processing_and_service_fee', 'refund_amount',
    'calculated_payout_amount'
]

LAZADA_REQUIRED = [
    'jumlah_(termasuk_pajak)'
]

SHOPEE_INVALID_REQUIRED = [
    'order_payout_amount'
]

TIKTOK_INVALID_REQUIRED = [
    'customer_refund', 'voucher_xtra_service_fee'
]

LAZADA_INVALID_REQUIRED = [
    'jumlah_(termasuk_pajak)', 'reason'
]

LAINNYA_REQUIRED = [
    'reason_code', 'date', 'amount', 'reference', 
    'gl_account', 'cost_center', 'wbs'
]

SUMMARY_COLUMNS = [
    'transaction_date', 'settlement_id', 'invoice_id', 'valid',
    'warehouse_name', 'warehouse_code', 'sales', 'invoice',
    'total_discount', 'shipping_fee', 'marketing_fee', 'admin_fee','escrow_amount'
]

# --- 3. Helper Functions ---

def _only_digits(s):
    return ''.join(ch for ch in str(s) if ch.isdigit())

def _format_with_commas(digits):
    return f"{int(digits):,}" if digits else ""

def _format_input_callback(key):
    s = st.session_state.get(key, "")
    digits = _only_digits(s)
    if digits == "":
        st.session_state[key] = ""
        return
    formatted = _format_with_commas(digits)
    if s != formatted:
        st.session_state[key] = formatted

def _parse_int_from_state(key, fallback=0):
    digits = _only_digits(st.session_state.get(key, ""))
    return int(digits) if digits else fallback

def _warehouse_code_from_name(name):
    if pd.isna(name) or str(name).strip() == "":
        return ""
    parts = [p for p in str(name).split() if p]
    part_1 = parts[0][:2] if len(parts) > 0 else ""
    part_2 = parts[1][:1] if len(parts) > 1 else ""
    part_3 = parts[2][:3] if len(parts) > 2 else ""
    part_4 = parts[3][:2] if len(parts) > 3 else ""
    code = part_1 + part_2 + part_3 + part_4
    return code.upper()

# Generate prefixes untuk invoice_id dan settlement_id
def calculate_prefixes(platform: str, withdraw_date: datetime) -> Tuple[str, str]:
    platform_code = platform[:3].upper()
    date_str = withdraw_date.strftime('%d%m%y')
    inv_prefix = f"INV_{platform_code}{date_str}"
    se_prefix = f"SE_{platform_code}{date_str}"
    return inv_prefix, se_prefix

# Generate settlement_id
def create_settlement_id_generator(prefix: str):
    def generate_settlement_id(warehouse_code=None) -> str:
        return prefix
    return generate_settlement_id

# Generate invoice_id
def create_invoice_id_generator(prefix: str):
    def generate_invoice_id(warehouse_code, amount=0) -> str:
        # Jika nilai negatif, ganti INV_ menjadi RET_
        final_prefix = prefix
        if amount < 0:
            final_prefix = prefix.replace("INV_", "RET_")
        return final_prefix + str(warehouse_code)
    return generate_invoice_id

# General function to clean numeric columns
def clean_numeric_columns(df: pd.DataFrame, columns_to_clean: list) -> pd.DataFrame:
    """
    Membersihkan koma, konversi ke numerik, dan mengisi nilai kosong (NaN) dengan 0.
    """
    for col in columns_to_clean:
        if col in df.columns:
            # 1. Pastikan kolom menjadi string agar bisa replace
            # 2. Ubah koma menjadi titik
            # 3. Konversi ke numerik, paksa error menjadi NaN
            df[col] = pd.to_numeric(
                df[col].astype(str).str.replace(',', '.', regex=False), 
                errors='coerce'
            ).fillna(0).astype('float64') # 4. Isi NaN dengan 0 dan set tipe float
    return df

def standardize_summary(df: pd.DataFrame) -> pd.DataFrame:
    """
    Menyeragamkan kolom summary untuk semua platform.
    Menambahkan kolom yang hilang dengan nilai 0 dan mengurutkan sesuai SUMMARY_COLUMNS.
    """
    for col in SUMMARY_COLUMNS:
        if col not in df.columns:
            df[col] = 0  # Atau np.nan jika lebih suka kosong
    
    return df[SUMMARY_COLUMNS]

# --- 4. UI untuk Input Section (Layout 4 Kolom) ---
st.write("Silakan lengkapi konfigurasi di bawah ini sebelum melakukan upload file.")

# Membuat 4 kolom untuk Input Section
with st.container(border=True):
    col1, col2, col3, col4 = st.columns(4)
    
    # KOLOM 1: Konfigurasi Platform
    with col1:
        st.markdown("**1. Konfigurasi Platform**")
        platform = st.selectbox("Pilih Marketplace", ["Shopee", "Tiktok Shop", "Lazada"], index=0)
        
        if platform == "Shopee":
            withdraw_date = st.date_input("Tanggal Settlement")
        else:
            withdraw_date = datetime.now().date()
    
    # KOLOM 2: Saldo Awal
    with col2:
        st.markdown("**2. Saldo Awal**")
        # if 'initial_balance_str' not in st.session_state:
            # st.session_state['initial_balance_str'] = f"{500:,}"
        
        st.text_input(
            "Jumlah saldo saat penarikan", 
            placeholder="5,000,000",
            key='initial_balance_str', 
            on_change=lambda: _format_input_callback('initial_balance_str'),
            help="Masukkan saldo awal sebelum penarikan dilakukan"
        )

    # KOLOM 3: Dana Ditarik
    with col3:
        st.markdown("**3. Dana Ditarik**")
        # if 'extracted_amount_str' not in st.session_state:
            # st.session_state['extracted_amount_str'] = f"{500:,}"
            
        st.text_input(
            "Nominal dana ditarik (aktual)", 
            key='extracted_amount_str', 
            placeholder="5,000,000",
            on_change=lambda: _format_input_callback('extracted_amount_str'),
            help="Masukkan nominal dana yang sebenarnya masuk ke rekening bank"
        )

    # KOLOM 4: Assignment (BARU)
        with col4:
            st.markdown("**4. Assignment Code**")
            assignment_val = st.text_input(
                "Assignment Code untuk Template SAP",
                placeholder="PAYADV_ID21",
                help="Nilai ini akan mengisi kolom Assignment pada tabel Journal Upload"
            )

# Parsing nilai saldo
initial_balance = _parse_int_from_state('initial_balance_str', fallback=0)
extracted_amount = _parse_int_from_state('extracted_amount_str', fallback=0)

# # --- 5. UI untuk Upload Data ---
st.markdown(f"#### Upload Raw Data {platform}")

# Membuat 2 kolom untuk memisahkan upload data utama dan data invalid
col_valid, col_invalid, col_expense = st.columns(3)

with col_valid:
    # st.markdown("**🟢 Data Valid**")
    # st.caption("Upload file transaksi utama yang berstatus sukses/settled.")
    uploaded_file = st.file_uploader(
        "**🟢 Data Valid**", 
        type=["csv", "xlsx"],
        help="File ini wajib diisi untuk memproses summary."
    )

with col_invalid:
    # st.markdown("### 🔴 Data Invalid")
    # st.caption("Upload file transaksi yang dibatalkan atau ditolak (Opsional).")
    uploaded_file_invalid = st.file_uploader(
        "**🔴 Data Invalid**", 
        type=["csv", "xlsx"], 
        key="invalid",
        help="Gunakan ini jika terdapat data transaksi rejection/invalid terpisah."
    )

with col_expense:
    uploaded_file_lainnya = st.file_uploader(
    "**📁 Data Expense Lainnya (Opsional)**", 
    type=["csv", "xlsx"], 
    key="lainnya"
)


# --- 6. UI untuk Dropdown Template Expense Lainnya ---
with st.expander("Butuh template data Expense Lainnya?"):
    st.write("Pastikan Expense Lainnya memiliki kolom berikut:")
    st.code("REASON CODE, DATE, AMOUNT, REFERENCE, GL ACCOUNT, COST CENTER, WBS")
    # Membuat dummy data kosong dengan header yang benar
    dummy_template = pd.DataFrame(columns=[c.replace('_', ' ').upper() for c in LAINNYA_REQUIRED])
    # Membuat file Excel di memori (Buffer)
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='xlsxwriter') as writer:
        dummy_template.to_excel(writer, index=False, sheet_name='Template')
        # Mengatur lebar kolom agar rapi saat dibuka
        workbook = writer.book
        worksheet = writer.sheets['Template']
        header_format = workbook.add_format({'bold': True, 'valign': 'vcenter'})
        for i, col in enumerate(dummy_template.columns):
            # Set lebar kolom jadi 20 dan set header bold
            worksheet.set_column(i, i, 20)
            worksheet.write(0, i, col, header_format)
    # Menyiapkan buffer untuk didownload
    buffer.seek(0)
    # Tombol Download Excel
    st.download_button(
        label="Download Template Excel (.xlsx)",
        data=buffer,
        file_name="template_lainnya.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
st.markdown("---")

# --- 7. Logic Cek Kolom Wajib ---
def check_required_columns(df: pd.DataFrame, required_cols: list, platform: str) -> bool:
    """
    Check df columns against required_cols. Show notification if missing columns.
    Returns True if all required columns present, False otherwise.
    """
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.warning(f"Uploaded file columns do not match expected for {platform}. Missing columns: {', '.join(missing)}")
        return False
    return True

# --- 8. Memproses Data Shopee ---
def process_shopee(df, withdraw_date, initial_balance, extracted_amount):
    """
    Pemrposes data Shopee dan mengembalikan Data Summary dan Tabel Escrow.
    """
    if not check_required_columns(df, SHOPEE_REQUIRED, "Shopee"):
        return None, None
    # Tentukan kolom yang perlu dibersihkan
    cols_to_fix = [
        'shipping_fee', 'quantity', 'price', 'total_discount', 
        'affiliate_commission_fee', 'commission_fee', 'service_fee', 
        'processing_fee', 'calculated_payout_amount', 'refund_amount',
        'total_comission_processing_and_service_fee'
    ]
    df = clean_numeric_columns(df, cols_to_fix)

    df['sales'] = df['quantity'] * df['price']
    df['invoice'] = df['sales'] + df['total_discount']
    df['transaction_date'] = pd.to_datetime(withdraw_date, format='%Y-%m-%d')

    # ada data yang menggunakan kolom total_comission_and_service_fee dan ada yang total_comission_processing_and_service_fee    
    if 'total_comission_and_service_fee' not in df.columns:
        if 'total_comission_processing_and_service_fee' in df.columns:
            df.rename(columns={'total_comission_processing_and_service_fee': 'total_comission_and_service_fee'}, inplace=True)
        else:
            return None, None
    
    # 4. Mapping ke standar SUMMARY_COLUMNS
    df['marketing_fee'] = df['affiliate_commission_fee'] 
    df['escrow_amount'] = df['calculated_payout_amount']
    df['warehouse_code'] = df['warehouse_name'].apply(_warehouse_code_from_name)
    df['admin_fee'] = df['total_comission_and_service_fee']
    df['valid'] = 'VALID'

    print_df = df.groupby("warehouse_name")[['sales', 'invoice', 
                                            'total_discount',
                                            'shipping_fee', 
                                            'marketing_fee', 
                                            'admin_fee', 
                                            'escrow_amount',]].sum().reset_index()
    print_df['transaction_date'] = pd.to_datetime(withdraw_date)
    print_df['warehouse_code'] = print_df['warehouse_name'].apply(_warehouse_code_from_name)
    print_df['valid'] = 'VALID'

    # Generate invoice_id and settlement_id
    inv_prefix, se_prefix = calculate_prefixes(platform, withdraw_date)
    inv_generator = create_invoice_id_generator(inv_prefix)
    se_generator = create_settlement_id_generator(se_prefix)
    
    print_df['invoice_id'] = print_df.apply(
        lambda row: inv_generator(row['warehouse_code'], row['escrow_amount']), 
        axis=1)
    print_df['settlement_id'] = print_df['warehouse_code'].apply(se_generator)
    print_df = print_df.reindex(columns=SUMMARY_COLUMNS, fill_value=0)
    
    # Membuat Row Total
    numeric_cols = print_df.select_dtypes(include=[np.number]).columns
    total_row_modified = {col: print_df[col].sum() if col in numeric_cols else ' ' for col in print_df.columns}
    total_row_modified['settlement_id'] = 'TOTAL'
    total_row_modified['valid'] = ' '
    sum_df = pd.concat([print_df, pd.DataFrame([total_row_modified])], ignore_index=True)
    
    # Kalkulasi tabel Escrow
    total_escrow_amount = int(print_df['escrow_amount'].sum())
    escrow_amount = int(extracted_amount)
    tolerance = 500
    # Warnin jika selisih melebihi toleransi
    if abs(escrow_amount - total_escrow_amount) > tolerance:
        st.warning(f"Nominal dana ditarik (invoice) yakni Rp{total_escrow_amount:,} berbeda dengan nominal dana ditarik (aktual) yakni Rp{escrow_amount:,}, mohon periksa kembali\n\nbatas toleransi adalah ±Rp500")        
    remaining_balance = initial_balance - escrow_amount
    escrow_balance_table = pd.DataFrame({
        'Keterangan': ['Jumlah saldo saat penarikan', 'Nominal dana ditarik (invoice)', 'Nominal dana ditarik (aktual)', 'Sisa saldo setelah ditarik'],
        'Nominal': [initial_balance, total_escrow_amount, escrow_amount, remaining_balance]
    })

    return sum_df, escrow_balance_table

# --- 8. Logic Proses Tiktok Shop ---
def process_tiktok_shop(df, withdraw_date, initial_balance, extracted_amount):
    if not check_required_columns(df, TIKTOK_REQUIRED, "Tiktok Shop"):
        return None, None

    # Cleaning numeric columns
    cols_to_fix = [
        'product_total_amount', 'calculated_seller_discounts', 'calculated_fee', 
        'calculated_shipping_fee', 'calculated_affiliate_commission', 
        'calculated_refund', 'calculated_settlement_amount'
    ]
    df = clean_numeric_columns(df, cols_to_fix)
    df.columns = df.columns.str.strip().str.replace(r'[\t\r\n]+', '', regex=True)

    df['sales'] = df['product_total_amount']
    df['shipping_fee'] = df['shipping_cost']
    df['marketing_fee'] = df['affiliate_partner_shop_ads_commission'] + df['sfp_service_fee']
    df['admin_fee'] = df['total_fees']
    df['escrow_amount'] = df['total_settlement_amount']
    df['invoice'] = df['escrow_amount'] + df['admin_fee'] + df['marketing_fee'] + df['shipping_fee']
    df['total_discount'] = df['sales'] - df['invoice']

    if 'order_settled_time' in df.columns:
        df['transaction_date'] = pd.to_datetime(df['order_settled_time'], errors='coerce')
    
    # Create warehouse_code from warehouse_name
    df['warehouse_code'] = df['warehouse_name'].apply(_warehouse_code_from_name)    
    # Group by order_settled_time and warehouse_name
    groupby_cols = ['transaction_date', 'warehouse_name']
    print_df = df.groupby(groupby_cols)[['sales', 'invoice', 'total_discount', 'shipping_fee', 
                                         'marketing_fee', 'admin_fee', 'escrow_amount']].sum().reset_index()
    # Add warehouse_code back to print_df
    print_df['warehouse_code'] = print_df['warehouse_name'].apply(_warehouse_code_from_name)
    print_df['valid'] = 'VALID'
    # print_df.rename(columns={
    #     'order_settled_time': 'transaction_date',
    #     'settlement_amount': 'escrow_amount',
    #     'seller_dicounts': 'total_discount',
    #     'tiktok_admin_fee': 'admin_fee',
    #     'seller_shipping_fee': 'shipping_fee',
    #     'promotion_(SFP)': 'marketing_fee',
    #     'affiliate_fee': 'affiliate_commission_fee'
    # }, inplace=True)
    
    # print_df['invoice'] = print_df['sales'] + print_df['total_discount']
    
    # Use same ID generator as Shopee
    inv_prefix, se_prefix = calculate_prefixes("Tiktok Shop", withdraw_date)
    inv_generator = create_invoice_id_generator(inv_prefix)
    se_generator = create_settlement_id_generator(se_prefix)
    
    # print_df['invoice_id'] = print_df['warehouse_code'].apply(inv_generator)
    print_df['settlement_id'] = print_df['warehouse_code'].apply(se_generator)
    print_df['invoice_id'] = print_df.apply(
        lambda row: inv_generator(row['warehouse_code'], row['escrow_amount']), 
        axis=1
    )

    print_df = print_df.reindex(columns=SUMMARY_COLUMNS)

    # Logika RETUR: RET prefix when escrow_amount is negative
    # print_df['invoice_id'] = print_df.apply(
    #     lambda row: "RET" + row['invoice_id'][3:] if row['escrow_amount'] < 0 else row['invoice_id'], 
    #     axis=1
    # )

    # print_df = print_df[['transaction_date','settlement_id', 'invoice_id', 'valid', 'warehouse_name', 'warehouse_code', 'sales', 'invoice', 'total_discount', 
    #                      'shipping_fee', 'marketing_fee', 'admin_fee', 'affiliate_commission_fee', 'escrow_amount']]    

    numeric_cols = print_df.select_dtypes(include=[np.number]).columns
    total_row_modified = {col: print_df[col].sum() if col in numeric_cols else ' ' for col in print_df.columns}
    total_row_modified['settlement_id'] = 'TOTAL'
    sum_df = pd.concat([print_df, pd.DataFrame([total_row_modified])], ignore_index=True)
    
    # --- ESCROW CALCULATION & WARNING LOGIC (UPDATED) ---
    total_escrow_amount = int(print_df['escrow_amount'].sum())
    escrow_amount = int(extracted_amount) # Menggunakan input user apa adanya
    tolerance = 500
    
    # CEK SELISIH MUTLAK (Tanpa syarat > 0)
    if abs(escrow_amount - total_escrow_amount) > tolerance:
        st.warning(f"Nominal dana ditarik (invoice) yakni Rp{total_escrow_amount:,} berbeda dengan nominal dana ditarik (aktual) yakni Rp{escrow_amount:,}, mohon periksa kembali \n\n batas toleransi: Rp±500")

    remaining_balance = initial_balance - escrow_amount
    escrow_balance_table = pd.DataFrame({
        'Keterangan': ['Jumlah saldo saat penarikan', 'Nominal dana ditarik (invoice)', 'Nominal dana ditarik (aktual)', 'Sisa saldo setelah ditarik'],
        'Nominal': [initial_balance, total_escrow_amount, escrow_amount, remaining_balance]
    })
    
    return sum_df, escrow_balance_table


def process_lazada(df, withdraw_date, initial_balance, extracted_amount):
    # Lazada biasanya menggunakan kolom 'jumlah_termasuk_pajak'
    df = clean_numeric_columns(df, ['jumlah_termasuk_pajak'])

    # Clean up column names
    # df.columns = df.columns.str.lower().str.strip().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
    
    # Convert numeric columns
    df['jumlah_termasuk_pajak'] = pd.to_numeric(df['jumlah_termasuk_pajak'], errors='coerce').fillna(0)
    
    # Parse transaction date
    df['tanggal_transaksi'] = pd.to_datetime(df['tanggal_transaksi'], format='%d %b %Y', errors='coerce')
    
    # --- FILTER AND SUM BY NAMA_BIAYA AND GROUP BY TANGGAL_TRANSAKSI ---
    
    # Sales: Sum of "Omset Penjualan" grouped by tanggal_transaksi
    sales_df = df[df['nama_biaya'] == 'Omset Penjualan'].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    sales_df.rename(columns={'jumlah_termasuk_pajak': 'sales'}, inplace=True)
    
    # Affiliate Commission Fee: Sum of rows containing "Commission" grouped by tanggal_transaksi
    commission_df = df[df['nama_biaya'].str.contains('Afiliasi', case=False, na=False)].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    commission_df.rename(columns={'jumlah_termasuk_pajak': 'affiliate_commission_fee'}, inplace=True)
    
    # Total Discount: Sum of "Diskon LazKoin" grouped by tanggal_transaksi
    diskon_df = df[df['nama_biaya'] == 'Diskon LazKoin'].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    diskon_df.rename(columns={'jumlah_termasuk_pajak': 'total_discount'}, inplace=True)
    
    # Shipping Fee: Sum of "Biaya Free Shipping Max" grouped by tanggal_transaksi
    shipping_df = df[df['nama_biaya'] == 'Biaya Free Shipping Max'].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    shipping_df.rename(columns={'jumlah_termasuk_pajak': 'shipping_fee'}, inplace=True)
    
    # Marketing Fee Components - grouped by tanggal_transaksi
    promo_voucher_df = df[df['nama_biaya'] == 'Biaya Promosi Voucher'].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    promo_voucher_df.rename(columns={'jumlah_termasuk_pajak': 'promo_voucher'}, inplace=True)
    
    promo_flexi_df = df[df['nama_biaya'].str.contains('Flexi-Combo', case=False, na=False)].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    promo_flexi_df.rename(columns={'jumlah_termasuk_pajak': 'promo_flexi'}, inplace=True)
    
    promo_lazkoin_df = df[df['nama_biaya'] == 'Biaya Promosi Diskon LazKoin'].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    promo_lazkoin_df.rename(columns={'jumlah_termasuk_pajak': 'promo_lazkoin'}, inplace=True)
    
    # Admin Fee: Sum of "Reversal Order Processing Fee" grouped by tanggal_transaksi
    admin_df = df[df['nama_biaya'].str.contains('Reversal', case=False, na=False)].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    admin_df.rename(columns={'jumlah_termasuk_pajak': 'admin_fee'}, inplace=True)
    
    # Merge all dataframes on tanggal_transaksi
    print_df = sales_df.copy()
    print_df = print_df.merge(commission_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(diskon_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(shipping_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(promo_voucher_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(promo_flexi_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(promo_lazkoin_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(admin_df, on='tanggal_transaksi', how='left')
    
    # Fill NaN values with 0
    print_df = print_df.fillna(0)
    
    # Calculate marketing_fee from sum of promotional fees
    print_df['marketing_fee'] = print_df['promo_voucher'] + print_df['promo_flexi'] + print_df['promo_lazkoin']
    
    # Drop individual promo columns
    print_df = print_df.drop(columns=['promo_voucher', 'promo_flexi', 'promo_lazkoin'])
    
    # Calculate escrow_amount
    print_df['escrow_amount'] = print_df['sales'] - print_df['affiliate_commission_fee']
    
    # Add transaction_date and warehouse_code
    print_df['transaction_date'] = print_df['tanggal_transaksi']
    print_df['warehouse_code'] = 'LAZADA'
    
    # Generate IDs
    inv_prefix, se_prefix = calculate_prefixes("Lazada", withdraw_date)
    inv_generator = create_invoice_id_generator(inv_prefix)
    se_generator = create_settlement_id_generator(se_prefix)
    
    print_df['invoice_id'] = print_df['warehouse_code'].apply(inv_generator)
    print_df['settlement_id'] = print_df['warehouse_code'].apply(se_generator)
    
    # Reorder columns
    print_df = print_df[['transaction_date', 'settlement_id', 'invoice_id', 'warehouse_code', 'sales', 'total_discount', 
                         'shipping_fee', 'marketing_fee', 'admin_fee', 'affiliate_commission_fee', 'escrow_amount']]
    
    # Add total row
    numeric_cols = print_df.select_dtypes(include=[np.number]).columns
    total_row_modified = {col: print_df[col].sum() if col in numeric_cols else ' ' for col in print_df.columns}
    total_row_modified['settlement_id'] = 'TOTAL'
    sum_df = pd.concat([print_df, pd.DataFrame([total_row_modified])], ignore_index=True)
    
    # --- ESCROW CALCULATION ---
    total_escrow_amount = int(print_df['escrow_amount'].sum())
    escrow_amount_final = int(extracted_amount)
    tolerance = 500
    
    if abs(escrow_amount_final - total_escrow_amount) > tolerance:
        st.warning(f"Nominal dana ditarik (invoice) yakni {total_escrow_amount:,} berbeda dengan nominal dana ditarik (aktual) yakni {escrow_amount_final:,}, mohon periksa kembali\n\nbatas toleransi adalah ±500")
    
    remaining_balance = initial_balance - escrow_amount_final
    escrow_balance_table = pd.DataFrame({
        'Keterangan': ['Jumlah saldo saat penarikan', 'Nominal dana ditarik (invoice)', 'Nominal dana ditarik (aktual)', 'Sisa saldo setelah ditarik'],
        'Nominal': [initial_balance, total_escrow_amount, escrow_amount_final, remaining_balance]
    })
    
    return sum_df, escrow_balance_table

# --- ADD THESE FUNCTIONS AFTER process_lainnya FUNCTION ---

def process_shopee_invalid(df, withdraw_date, initial_balance, extracted_amount):
    """Process invalid/rejected Shopee transactions"""
    required_cols = SHOPEE_INVALID_REQUIRED # ['order_payout_amount']
    
    # Pastikan kolom dibersihkan dulu sebelum dicek
    df.columns = df.columns.str.lower().str.replace(' ', '_').str.rstrip('_')
    df = clean_numeric_columns(df, ['order_payout_amount'])
    
    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.warning("Missing columns in Invalid File: " + ", ".join(missing))
        return None # Konsisten mengembalikan None tunggal

    # Konversi ke numerik
    df['sales'] = df['order_seller_partial_refund_amount']
    df['transaction_date'] = pd.to_datetime(withdraw_date, format='%Y-%m-%d')
    df['marketing_fee'] = df['order_affiliate_commission_amount'] 
    df['escrow_amount'] = df['order_payout_amount']
    df['warehouse_code'] = df['warehouse_name'].apply(_warehouse_code_from_name)
    df['admin_fee'] = df['order_service_fee_amount'] + df['order_total_transaction_fee_amount']
    df['total_discount'] = df['order_total_seller_discount_amount'] + df['order_total_shopee_discount_amount']
    df['shipping_fee'] = df['order_final_shipping_amount']
    df['invoice'] = df['sales'] + df['total_discount']
    
    # Gunakan warehouse_name (pastikan kolom ini ada di CSV Anda)
    df['warehouse_code'] = df['warehouse_name'].apply(_warehouse_code_from_name)


    groupby_cols = ['transaction_date', 'warehouse_name']
    print_df = df.groupby(groupby_cols)[['sales', 'invoice', 'total_discount', 'shipping_fee', 
                                         'marketing_fee', 'admin_fee', 'escrow_amount']].sum().reset_index()
    print_df.rename(columns={'order_payout_amount': 'escrow_amount'}, inplace=True)
    print_df['transaction_date'] = pd.to_datetime(withdraw_date)

    inv_prefix, se_prefix = calculate_prefixes("Shopee", withdraw_date)
    inv_generator = create_invoice_id_generator(inv_prefix)
    se_generator = create_settlement_id_generator(se_prefix)
    
    print_df['warehouse_code'] = print_df['warehouse_name'].apply(_warehouse_code_from_name)
    print_df['invoice_id'] = print_df['warehouse_code'].apply(inv_generator)
    print_df['settlement_id'] = print_df['warehouse_code'].apply(se_generator)

    # If escrow negative, incove_id akan menjadi RET_ 
    print_df['invoice_id'] = print_df.apply(lambda r: ("RET_" + str(r['invoice_id']).split('_',1)[1]) if r['escrow_amount'] < 0 else r['invoice_id'], axis=1)

    # Mark validity
    print_df['valid'] = 'INVALID'

    # Pastikan urutan kolom seragam dengan summary utama (tanpa TOTAL row)
    cols_order = ['transaction_date', 'settlement_id', 'invoice_id', 'warehouse_name', 'warehouse_code',
                  'sales', 'invoice', 'total_discount', 'shipping_fee', 'marketing_fee', 'admin_fee', 'escrow_amount', 'valid']
    for c in cols_order:
        if c not in print_df.columns:
            print_df[c] = np.nan
    return print_df[cols_order]

def process_tiktok_invalid(df, withdraw_date, initial_balance, extracted_amount):
    """Process invalid/rejected Tiktok Shop transactions"""
    required_cols = TIKTOK_INVALID_REQUIRED # ['customer_refund', 'voucher_xtra_service_fee']

    df.columns = df.columns.str.strip().str.replace(r'[\t\r\n]+', '', regex=True)
    # Daftar kolom numerik Tiktok yang cukup banyak
    cols_to_fix = [
        'product_total_amount', 'calculated_seller_discounts', 'calculated_fee', 
        'calculated_shipping_fee', 'calculated_affiliate_commission', 
        'calculated_refund', 'calculated_settlement_amount'
    ]
    df = clean_numeric_columns(df, cols_to_fix)

    missing = [c for c in required_cols if c not in df.columns]
    if missing:
        st.warning("Missing columns in Invalid File: " + ", ".join(missing))
        return None
    
    df['sales'] = df['total_settlement_amount']
    df['shipping_fee'] = df['shipping_cost']
    df['marketing_fee'] = df['affiliate_partner_shop_ads_commission'] + df['sfp_service_fee']
    df['admin_fee'] = df['total_fees']
    df['escrow_amount'] = df['total_settlement_amount']
    df['invoice'] = df['escrow_amount'] + df['admin_fee'] + df['marketing_fee'] + df['shipping_fee']
    df['total_discount'] = df['sales'] - df['invoice']

    if 'order_settled_time' in df.columns:
        df['transaction_date'] = pd.to_datetime(df['order_settled_time'], errors='coerce')
    
    df['warehouse_code'] = df['warehouse_name'].apply(_warehouse_code_from_name)
    
    groupby_cols = ['transaction_date', 'warehouse_name']
    print_df = df.groupby(groupby_cols)[['sales', 'invoice', 'total_discount', 'shipping_fee', 
                                         'marketing_fee', 'admin_fee', 'escrow_amount']].sum().reset_index()
    
    print_df['warehouse_code'] = print_df['warehouse_name'].apply(_warehouse_code_from_name)
    print_df['valid'] = 'INVALID'
    
    inv_prefix, se_prefix = calculate_prefixes("Tiktok Shop", withdraw_date)
    inv_generator = create_invoice_id_generator(inv_prefix)
    se_generator = create_settlement_id_generator(se_prefix)
    
    # print_df['invoice_id'] = print_df['warehouse_code'].apply(inv_generator)
    print_df['settlement_id'] = print_df['warehouse_code'].apply(se_generator)
    print_df['invoice_id'] = print_df.apply(
        lambda row: inv_generator(row['warehouse_code'], row['escrow_amount']), 
        axis=1
    )
    print_df = print_df.reindex(columns=SUMMARY_COLUMNS)

    # Logika RETUR & Penanda INVALID (Sesuai gaya Shopee)
    # Jika escrow negatif pakai "RET", jika positif pakai "INV" + penanda _INVALID
    print_df['invoice_id'] = print_df.apply(
        lambda r: ("RET_" + str(r['invoice_id']).split('_', 1)[1]) 
        if r['escrow_amount'] < 0 
        else (str(r['invoice_id'])), 
        axis=1
    )
    
    return print_df

def process_lazada_invalid(df, withdraw_date, initial_balance, extracted_amount):
    """Process invalid/rejected Lazada transactions"""
    df.columns = df.columns.str.lower().str.strip().str.replace(' ', '_').str.replace('(', '').str.replace(')', '')
    df = clean_numeric_columns(df, ['jumlah_termasuk_pajak'])
    
    df['jumlah_termasuk_pajak'] = pd.to_numeric(df['jumlah_termasuk_pajak'], errors='coerce').fillna(0)
    
    if 'tanggal_transaksi' in df.columns:
        df['tanggal_transaksi'] = pd.to_datetime(df['tanggal_transaksi'], format='%d %b %Y', errors='coerce')
    else:
        df['tanggal_transaksi'] = pd.to_datetime(withdraw_date)
    
    sales_df = df[df['nama_biaya'] == 'Omset Penjualan'].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    sales_df.rename(columns={'jumlah_termasuk_pajak': 'sales'}, inplace=True)
    
    commission_df = df[df['nama_biaya'].str.contains('Afiliasi', case=False, na=False)].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    commission_df.rename(columns={'jumlah_termasuk_pajak': 'affiliate_commission_fee'}, inplace=True)
    
    diskon_df = df[df['nama_biaya'] == 'Diskon LazKoin'].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    diskon_df.rename(columns={'jumlah_termasuk_pajak': 'total_discount'}, inplace=True)
    
    shipping_df = df[df['nama_biaya'] == 'Biaya Free Shipping Max'].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    shipping_df.rename(columns={'jumlah_termasuk_pajak': 'shipping_fee'}, inplace=True)
    
    promo_voucher_df = df[df['nama_biaya'] == 'Biaya Promosi Voucher'].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    promo_voucher_df.rename(columns={'jumlah_termasuk_pajak': 'promo_voucher'}, inplace=True)
    
    promo_flexi_df = df[df['nama_biaya'].str.contains('Flexi-Combo', case=False, na=False)].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    promo_flexi_df.rename(columns={'jumlah_termasuk_pajak': 'promo_flexi'}, inplace=True)
    
    promo_lazkoin_df = df[df['nama_biaya'] == 'Biaya Promosi Diskon LazKoin'].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    promo_lazkoin_df.rename(columns={'jumlah_termasuk_pajak': 'promo_lazkoin'}, inplace=True)
    
    admin_df = df[df['nama_biaya'].str.contains('Reversal', case=False, na=False)].groupby('tanggal_transaksi')[['jumlah_termasuk_pajak']].sum().reset_index()
    admin_df.rename(columns={'jumlah_termasuk_pajak': 'admin_fee'}, inplace=True)
    
    print_df = sales_df.copy()
    print_df = print_df.merge(commission_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(diskon_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(shipping_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(promo_voucher_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(promo_flexi_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(promo_lazkoin_df, on='tanggal_transaksi', how='left')
    print_df = print_df.merge(admin_df, on='tanggal_transaksi', how='left')
    
    print_df = print_df.fillna(0)
    
    print_df['marketing_fee'] = print_df['promo_voucher'] + print_df['promo_flexi'] + print_df['promo_lazkoin']
    
    print_df = print_df.drop(columns=['promo_voucher', 'promo_flexi', 'promo_lazkoin'])
    
    print_df['escrow_amount'] = print_df['sales'] - print_df['affiliate_commission_fee']
    
    print_df['transaction_date'] = print_df['tanggal_transaksi']
    print_df['warehouse_code'] = 'LAZADA'
    
    inv_prefix, se_prefix = calculate_prefixes("Lazada", withdraw_date)
    inv_generator = create_invoice_id_generator(inv_prefix)
    se_generator = create_settlement_id_generator(se_prefix)
    
    print_df['invoice_id'] = print_df['warehouse_code'].apply(inv_generator)
    print_df['settlement_id'] = print_df['warehouse_code'].apply(se_generator)

    # Mark as INVALID
    print_df['invoice_id'] = "INV" + print_df['invoice_id'][3:]
    
    print_df = print_df[['transaction_date', 'settlement_id', 'invoice_id', 'warehouse_code', 'sales', 'total_discount', 
                         'shipping_fee', 'marketing_fee', 'admin_fee', 'affiliate_commission_fee', 'escrow_amount']]
    
    return print_df
    
def process_lainnya(df):
    """Memproses file LAINNYA: Validasi, Cleaning, dan Parsing Tanggal"""
    
    # 1. Bersihkan nama kolom
    df.columns = df.columns.str.lower().str.replace(' ', '_', regex=False)
    
    # 2. Cek kolom yang hilang
    required_std = [c.lower().replace(' ', '_') for c in LAINNYA_REQUIRED]
    missing = [c for c in required_std if c not in df.columns]
    
    if missing:
        missing_formatted = [m.replace('_', ' ').upper() for m in missing]
        return None, f"Kolom {', '.join(missing_formatted)} tidak terdapat pada data, mohon periksa kembali"

    # 3. Pastikan kolom AMOUNT numerik
    if 'amount' in df.columns:
        df['amount'] = df['amount'].astype(str).str.replace(',', '', regex=False)
        df['amount'] = pd.to_numeric(df['amount'], errors='coerce').fillna(0)

    # 4. (BARU) Pastikan kolom DATE berformat datetime agar cocok dengan Shopee
    if 'date' in df.columns:
        # Menggunakan dayfirst=True untuk asumsi format tanggal Indonesia (DD/MM/YYYY)
        df['date'] = pd.to_datetime(df['date'], dayfirst=True, errors='coerce')

    return df, None

def add_total_row(df):
    """Fungsi helper untuk menambahkan baris TOTAL di paling bawah"""
    if df is None or df.empty:
        return df
        
    # Identifikasi kolom numerik yg ingin dijumlahkan
    # Kita spesifik saja ke 'Amount in Functional Currency' agar 'Customer' (angka) tidak ikut ter-sum
    target_sum_col = 'Amount in Functional Currency'
    
    total_row = {col: ' ' for col in df.columns}
    
    if target_sum_col in df.columns:
        total_row[target_sum_col] = df[target_sum_col].sum()
    
    # Label "TOTAL" di kolom pertama (Reason Code)
    if len(df.columns) > 0:
        total_row[df.columns[0]] = 'TOTAL'
        
    return pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)

def merge_ecommerce_to_lainnya(df_lainnya_asli, df_ecommerce_summary):
    """Menggabungkan hasil E-Commerce ke dalam tabel LAINNYA dengan logika tanggal terpisah"""
    
    # 1. Ambil data E-Commerce (Filter baris TOTAL jika ada)
    if df_ecommerce_summary is not None and 'settlement_id' in df_ecommerce_summary.columns:
        ecom_data = df_ecommerce_summary[df_ecommerce_summary['settlement_id'] != 'TOTAL'].copy()
    else:
        ecom_data = pd.DataFrame()

    # 2. Buat DataFrame konversi untuk E-Commerce (Shopee/Lazada/Tiktok)
    df_converted = pd.DataFrame()
    
    if not ecom_data.empty:
        # A. LOGIKA TANGGAL SHOPEE/ECOMMERCE (INVOICE & RETUR)
        df_converted['date'] = ecom_data['transaction_date']
        
        # B. Mapping Kolom Lainnya
        df_converted['reference'] = ecom_data['invoice_id']
        df_converted['amount'] = ecom_data['escrow_amount']
        
        # Add validity column from ecom summary
        df_converted['valid'] = ecom_data.get('valid', 'VALID').values if 'valid' in ecom_data.columns else 'VALID'
        
        # Logic Reason Code
        def get_reason(ref):
            ref_str = str(ref)
            if ref_str.startswith("INV"):
                return "INVOICE"
            elif ref_str.startswith("RE"):
                return "RETUR"
            else:
                return np.nan 
        df_converted['reason_code'] = df_converted['reference'].apply(get_reason)
        
        # Set NaN untuk kolom lainnya
        for col in ['gl_account', 'cost_center', 'wbs']:
            df_converted[col] = np.nan

    # 3. Gabungkan dengan data LAINNYA (jika ada)
    target_cols = [c.lower().replace(' ', '_') for c in LAINNYA_REQUIRED]
    
    # Pastikan df_converted punya semua kolom target + valid
    for col in target_cols:
        if col not in df_converted.columns:
            df_converted[col] = np.nan
    if 'valid' not in df_converted.columns:
        df_converted['valid'] = np.nan
    
    # Ensure ordering includes valid (valid will pass through format_final_consolidation if needed)
    df_converted = df_converted[target_cols + ['valid']] if 'valid' in df_converted.columns else df_converted[target_cols]
    
    final_df = df_converted # Default jika tidak ada data lainnya

    if df_lainnya_asli is not None and not df_lainnya_asli.empty:
        df_lainnya_ready = df_lainnya_asli[target_cols]
        final_df = pd.concat([df_lainnya_ready, df_converted[target_cols + ['valid']]], ignore_index=True)

    return final_df

CUSTOMER_MAP = {
    "Shopee": "1000010001",
    "Lazada": "1000010002",
    "Tiktok Shop": "1000010003",
    "LAINNYA": "1000010004"
}

def format_final_consolidation(df, platform_selected):
    """
    Mengubah format kolom tabel gabungan sesuai standar accounting SAP/Finance.
    """
    if df is None or df.empty:
        return df

    # 1. Mapping Kolom yang sudah ada ke nama baru
    # TAMBAHAN: 'date' -> 'Clearing Date'
    rename_map = {
        'reason_code': 'Reason Code',
        'amount': 'Amount in Functional Currency',
        'reference': 'Billing No',
        'gl_account': 'GL Expense',
        'cost_center': 'Cost Center',
        'date': 'Clearing Date'  # <-- BARIS INI YANG DITAMBAHKAN
    }
    df = df.rename(columns=rename_map)

    # 2. Isi Kolom Fix / Logika Baru
    df['Company Code'] = 'ID03'
    
    # Customer Code berdasarkan platform
    cust_code = CUSTOMER_MAP.get(platform_selected, "1000099999")
    df['Customer'] = cust_code
    
    df['Business Area'] = '1201'
    df['House Bank'] = 'MDR01'
    df['Account ID'] = '97716'
    df['Currency'] = 'IDR'
    df['Assignment'] = assignment_val

    
    # 3. Definisikan Urutan Kolom Final
    final_columns = [
        "Reason Code", 
        "Company Code", 
        "Customer", 
        "Assignment", 
        "Clearing Date", 
        "Business Area", 
        "House Bank", 
        "Account ID", 
        "Currency", 
        "Amount in Functional Currency", 
        "Billing No", 
        "Sales Document", 
        "PO No", 
        "Faktur Pajak", 
        "Cheque/ Giro Bank Name", 
        "Cheque/ Giro Bank No", 
        "Cheque/ Giro due", 
        "GL Expense", 
        "Cost Center"
    ]
    
    # 4. Tambahkan kolom yang belum ada dengan nilai NaN
    for col in final_columns:
        if col not in df.columns:
            df[col] = np.nan
            
    # 5. Reorder kolom dan return
    return df[final_columns]

# --- EXECUTION LOGIC ---
# --- UPDATE EXECUTION LOGIC SECTION (Replace the entire "if uploaded_file:" block) ---
if uploaded_file:
    run = st.button("Run Processing")
    if run:
        with st.spinner('Sedang memproses data... Mohon tunggu ⏳'):
            # 1. BACA FILE UTAMA (VALID)
            if uploaded_file.name.endswith('.csv'):
                df = pd.read_csv(uploaded_file, sep=None, engine='python', encoding='windows-1252')
            else:
                df = pd.read_excel(uploaded_file)

            df.columns = df.columns.str.lower()
            df.dropna(how='all', inplace=True)
            df.fillna(0, inplace=True)
            df.columns = df.columns.str.replace(' ', '_', regex=False).str.rstrip('_')

            sum_df = None
            escrow_balance_table = None
            invalid_df = None

            # --- PROSES E-COMMERCE (VALID) ---
            if platform == "Tiktok Shop":
                if check_required_columns(df, TIKTOK_REQUIRED, platform):
                    sum_df, escrow_balance_table = process_tiktok_shop(df, withdraw_date, initial_balance, extracted_amount)
            elif platform == "Shopee":
                if check_required_columns(df, SHOPEE_REQUIRED, platform):
                    sum_df, escrow_balance_table = process_shopee(df, withdraw_date, initial_balance, extracted_amount)
            elif platform == "Lazada":
                if check_required_columns(df, LAZADA_REQUIRED, platform):
                    sum_df, escrow_balance_table = process_lazada(df, withdraw_date, initial_balance, extracted_amount)

            # --- PROSES FILE INVALID (OPTIONAL) ---
            if uploaded_file_invalid is not None:
                if uploaded_file_invalid.name.endswith('.csv'):
                    df_inv_raw = pd.read_csv(uploaded_file_invalid, sep=None, engine='python', encoding='windows-1252')
                else:
                    df_inv_raw = pd.read_excel(uploaded_file_invalid)

                df_inv_raw.columns = df_inv_raw.columns.str.lower()
                df_inv_raw.dropna(how='all', inplace=True)
                df_inv_raw.fillna(0, inplace=True)
                df_inv_raw.columns = df_inv_raw.columns.str.replace(' ', '_', regex=False).str.rstrip('_')

                if platform == "Tiktok Shop":
                    invalid_df = process_tiktok_invalid(df_inv_raw, withdraw_date, initial_balance, extracted_amount)
                elif platform == "Shopee":
                    invalid_df = process_shopee_invalid(df_inv_raw, withdraw_date, initial_balance, extracted_amount)
                elif platform == "Lazada":
                    invalid_df = process_lazada_invalid(df_inv_raw, withdraw_date, initial_balance, extracted_amount)

            # --- 2. GABUNGKAN INVALID KE SUMMARY (SINKRONISASI TOTAL) ---
            if sum_df is not None and invalid_df is not None:
                # Ambil data valid saja (buang baris TOTAL jika ada)
                base_valid = sum_df[sum_df['settlement_id'] != 'TOTAL'].copy()
                
                # Samakan kolom invalid agar bisa di-concat
                for c in base_valid.columns:
                    if c not in invalid_df.columns:
                        if c in base_valid.select_dtypes(include=[np.number]).columns:
                            invalid_df[c] = 0
                        else:
                            invalid_df[c] = np.nan
                
                # Gabungkan data asli
                combined_ecom = pd.concat([base_valid, invalid_df[base_valid.columns]], ignore_index=True)

                # Hitung ulang baris TOTAL untuk tampilan Summary
                numeric_cols = combined_ecom.select_dtypes(include=[np.number]).columns
                total_row_mod = {col: combined_ecom[col].sum() if col in numeric_cols else ' ' for col in combined_ecom.columns}
                if 'settlement_id' in combined_ecom.columns: 
                    total_row_mod['settlement_id'] = 'TOTAL'
                
                sum_df = pd.concat([combined_ecom, pd.DataFrame([total_row_mod])], ignore_index=True)

                # Update tabel Escrow Balance dengan total baru yang sudah digabung
                total_escrow_new = int(combined_ecom['escrow_amount'].sum())
                escrow_actual = int(extracted_amount)
                escrow_balance_table = pd.DataFrame({
                    'Keterangan': ['Jumlah saldo saat penarikan', 'Nominal dana ditarik (invoice)', 'Nominal dana ditarik (aktual)', 'Sisa saldo setelah ditarik'],
                    'Nominal': [initial_balance, total_escrow_new, escrow_actual, initial_balance - escrow_actual]
                })

            # --- 3. PROSES DATA LAINNYA & KONSOLIDASI JURNAL ---
            df_lainnya_clean = None
            if uploaded_file_lainnya is not None:
                if uploaded_file_lainnya.name.endswith('.csv'):
                    df_lain_raw = pd.read_csv(uploaded_file_lainnya, sep=None, engine='python', encoding='windows-1252')
                else:
                    df_lain_raw = pd.read_excel(uploaded_file_lainnya)
                
                df_lainnya_clean, err = process_lainnya(df_lain_raw)
                if err: 
                    st.error(err)

            # PEMBUATAN TABEL JURNAL (Hanya 1 jalur proses agar tidak double input)
            final_lainnya_df = None
            if sum_df is not None:
                # Gunakan sum_df yang sudah digabung dengan invalid di atas
                # Fungsi merge_ecommerce_to_lainnya sudah memfilter baris 'TOTAL' di dalamnya
                merged_for_journal = merge_ecommerce_to_lainnya(df_lainnya_clean, sum_df)
                
                # Terapkan format standar SAP
                formatted_journal = format_final_consolidation(merged_for_journal, platform)
                
                # Tambahkan baris TOTAL di paling bawah
                final_lainnya_df = add_total_row(formatted_journal)
            elif df_lainnya_clean is not None:
                # Jika hanya ada data expense lainnya tanpa e-commerce
                formatted_journal = format_final_consolidation(df_lainnya_clean, platform)
                final_lainnya_df = add_total_row(formatted_journal)
        # ---------------------------------------------------------
        # 4. DISPLAY
        # ---------------------------------------------------------
        def highlight_no_match(row):
            # Cek basis tema yang dikonfigurasi (light atau dark)
            theme_base = st.get_option("theme.base")
            if theme_base == "dark":
                # Merah gelap untuk tema dark
                color = "#4d0000"
            else:
                # Merah muda untuk tema light
                color = "#ffcccc"
            if 'warehouse_name' in row.index and 'No Match Found' in str(row['warehouse_name']):
                return [f'background-color: {color}'] * len(row)
            return [''] * len(row)


        if sum_df is not None:
            st.markdown(f"### Summary of {platform} Settlement on {withdraw_date}")
            
            # --- A. LOGIKA NOTICE (Terpusat) ---
            if 'warehouse_name' in sum_df.columns:
                has_no_match = sum_df['warehouse_name'].astype(str).str.contains('No Match Found', case=False).any()
                if has_no_match:
                    st.warning("⚠️ **Peringatan:** Terdeteksi data dengan **'No Match Found'** pada kolom Warehouse Name. Mohon periksa kembali atau hubungi pihak terkait.")

            # --- B. LOGIKA STYLING (Highlight & Format Angka) ---
            cols_to_format = ['sales', 'invoice', 'total_discount', 'shipping_fee', 'marketing_fee', 'admin_fee', 'affiliate_commission_fee', 'escrow_amount']
            valid_cols = [c for c in cols_to_format if c in sum_df.columns]
            
            # Menggabungkan highlight baris dan format ribuan
            styled_df = (sum_df.style
                         .apply(highlight_no_match, axis=1)
                         .format({col: "{:,.0f}" for col in valid_cols}))
            
            st.dataframe(styled_df,
                         use_container_width=True)

            if escrow_balance_table is not None:
                st.markdown("### Escrow Balance")
                st.dataframe(escrow_balance_table.style.format({"Nominal": "{:,.0f}"}))
        
        if final_lainnya_df is not None:
            st.divider()
            st.markdown("### Consolidated Data (SAP Format)")
            target_col = "Amount in Functional Currency"
            column_configuration = {
                "Billing No": st.column_config.TextColumn("Billing No", width="medium"),
                "Clearing Date": st.column_config.DateColumn("Clearing Date", width="medium"),
                "Reason Code": st.column_config.TextColumn("Reason Code", width="small"),
            }
            if target_col in final_lainnya_df.columns:
                st.dataframe(
                    final_lainnya_df.style.format({target_col: "{:,.0f}"}, na_rep="-"),
                    use_container_width=True,
                    column_config=column_configuration
                )
            else:
                st.dataframe(
                    final_lainnya_df,
                    use_container_width=True,
                    column_config=column_configuration
                )

        # ---------------------------------------------------------
        # 5. DOWNLOAD
        # ---------------------------------------------------------
        if sum_df is not None or final_lainnya_df is not None:
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
                if sum_df is not None:
                    sum_df.to_excel(writer, index=False, sheet_name='ProcessedData')
                if escrow_balance_table is not None:
                    escrow_balance_table.to_excel(writer, index=False, sheet_name='Escrow_Balance')
                if invalid_df is not None:
                    invalid_df.to_excel(writer, index=False, sheet_name='Invalid_Data')
                if final_lainnya_df is not None:
                    final_lainnya_df.to_excel(writer, index=False, sheet_name='Journal_Upload')

                workbook = writer.book
                int_fmt = workbook.add_format({'num_format': '#,##0'})
                
                if sum_df is not None:
                    ws = writer.sheets['ProcessedData']
                    for col in sum_df.select_dtypes(include=[np.number]).columns:
                        try: ws.set_column(sum_df.columns.get_loc(col), sum_df.columns.get_loc(col), 20, int_fmt)
                        except: pass
                
                if escrow_balance_table is not None:
                    ws = writer.sheets['Escrow_Balance']
                    try: ws.set_column(escrow_balance_table.columns.get_loc("Nominal"), escrow_balance_table.columns.get_loc("Nominal"), 20, int_fmt)
                    except: pass

                if invalid_df is not None:
                    ws = writer.sheets['Invalid_Data']
                    for col in invalid_df.select_dtypes(include=[np.number]).columns:
                        try: ws.set_column(invalid_df.columns.get_loc(col), invalid_df.columns.get_loc(col), 20, int_fmt)
                        except: pass

                if final_lainnya_df is not None:
                    ws = writer.sheets['Journal_Upload']
                    if "Amount in Functional Currency" in final_lainnya_df.columns:
                        try: 
                            idx = final_lainnya_df.columns.get_loc("Amount in Functional Currency")
                            ws.set_column(idx, idx, 20, int_fmt)
                        except: pass

            output.seek(0)
            file_name = f"Summary_{platform}_{withdraw_date.strftime('%d%m%y')}.xlsx"
            st.download_button(
                label="Download All Processed Data (Excel)",
                data=output,
                file_name=file_name,
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
            )
