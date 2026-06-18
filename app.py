import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
import json
import re
from io import BytesIO
from datetime import datetime

# ── Page Config ──────────────────────────────────────────────
st.set_page_config(
    page_title="SKEP Updater",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded",
)
st.write("App loaded")
try:
    st.write(st.secrets["gcp_service_account"]["client_email"])
except Exception as e:
    st.error(f"Secrets error: {e}")
# ── Styling ───────────────────────────────────────────────────
st.markdown("""
<style>
    .main { padding-top: 1rem; }
    .stAlert { border-radius: 8px; }
    .metric-card {
        background: #f0f4ff;
        border-left: 4px solid #4361ee;
        padding: 12px 16px;
        border-radius: 6px;
        margin: 4px 0;
    }
    .metric-card.green { background: #f0fff4; border-color: #2d9e5e; }
    .metric-card.red   { background: #fff0f0; border-color: #e53e3e; }
    .metric-card.gray  { background: #f7f7f7; border-color: #a0aec0; }
    .tag { display:inline-block; padding:2px 8px; border-radius:12px; font-size:12px; font-weight:600; }
    .tag-update { background:#d1fae5; color:#065f46; }
    .tag-nochange { background:#e2e8f0; color:#4a5568; }
    .tag-notfound { background:#fee2e2; color:#991b1b; }
</style>
""", unsafe_allow_html=True)

# ── Constants ─────────────────────────────────────────────────
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets",
    "https://www.googleapis.com/auth/drive.readonly",
]

# Mapping kolom SKEP (kondisi baru) → nama field internal
SKEP_BARU_COLS = {
    "Kode Bagan Baru":   "kode_bagan_baru",
    "Unit Kerja Baru":   "unit_kerja_baru",
    "Jabatan Baru":      "jabatan_baru",
    "Job Grade Baru":    "job_grade_baru",
    "Pangkat Baru":      "pangkat_baru",
    "Gaji Pokok Baru":   "gaji_pokok_baru",
}

# ── Helper: Connect to Google Sheets ─────────────────────────
@st.cache_resource(ttl=300)
def get_gspread_client():
    creds_dict = st.secrets["gcp_service_account"]
    creds = Credentials.from_service_account_info(creds_dict, scopes=SCOPES)
    return gspread.authorize(creds)

def get_sheet(spreadsheet_url: str, sheet_name: str):
    gc = get_gspread_client()
    sh = gc.open_by_url(spreadsheet_url)
    return sh.worksheet(sheet_name)

# ── Helper: Parse SKEP Excel ──────────────────────────────────
def parse_skep(file) -> tuple[pd.DataFrame, list[str]]:
    """
    Baca semua sheet dari file SKEP.
    Return: (DataFrame gabungan, list nama sheet)
    """
    xl = pd.ExcelFile(file)
    all_rows = []
    sheet_names = xl.sheet_names

    for sheet_name in sheet_names:
        raw = pd.read_excel(file, sheet_name=sheet_name, header=None)

        # Cari baris header (row 0 = grup, row 1 = sub-header)
        # Struktur: row0=NO.URUT|NAMA|NP|...|KONDISI LAMA|...|KONDISI BARU|...|KETERANGAN|Tanggal SKEP
        # row1=sub kolom
        # Deteksi jumlah kolom
        ncols = raw.shape[1]

        # Deteksi posisi blok "KONDISI BARU" dari row 0
        baru_start = None
        ket_col = None
        tgl_col = None

        for c in range(ncols):
            val = str(raw.iloc[0, c]).strip()
            if "KONDISI BARU" in val.upper():
                baru_start = c
            if "KETERANGAN" in val.upper():
                ket_col = c
            if "TANGGAL SKEP" in val.upper():
                tgl_col = c

        if baru_start is None:
            st.warning(f"Sheet '{sheet_name}': tidak bisa mendeteksi blok KONDISI BARU, dilewati.")
            continue

        # Sub-header di row 1: KODE BAGAN, UNIT KERJA, JABATAN, JOB GRADE, PANGKAT, GAJI POKOK
        # Blok lama: mulai col 3 s/d baru_start-1
        # Blok baru: mulai baru_start s/d ket_col-1

        for row_idx in range(2, len(raw)):
            row = raw.iloc[row_idx]

            no_urut   = row.iloc[0]
            nama      = row.iloc[1]
            np_val    = row.iloc[2]

            # Skip baris kosong
            if pd.isna(no_urut) or pd.isna(nama) or pd.isna(np_val):
                continue
            if str(nama).strip() == "" or str(np_val).strip() == "":
                continue

            # Kolom kondisi baru (baru_start = Kode Bagan baru)
            baru_kode   = row.iloc[baru_start]   if baru_start   < ncols else None
            baru_unit   = row.iloc[baru_start+1] if baru_start+1 < ncols else None
            baru_jabatan= row.iloc[baru_start+2] if baru_start+2 < ncols else None
            baru_grade  = row.iloc[baru_start+3] if baru_start+3 < ncols else None
            baru_pangkat= row.iloc[baru_start+4] if baru_start+4 < ncols else None
            baru_gaji   = row.iloc[baru_start+5] if baru_start+5 < ncols else None

            keterangan  = row.iloc[ket_col] if ket_col is not None and ket_col < ncols else None
            tgl_skep    = row.iloc[tgl_col] if tgl_col is not None and tgl_col < ncols else None

            all_rows.append({
                "sheet_skep":      sheet_name,
                "no_urut":         no_urut,
                "nama":            str(nama).strip(),
                "np":              str(np_val).strip(),
                "kode_bagan_baru": str(baru_kode).strip()   if pd.notna(baru_kode)    else "",
                "unit_kerja_baru": str(baru_unit).strip()   if pd.notna(baru_unit)    else "",
                "jabatan_baru":    str(baru_jabatan).strip() if pd.notna(baru_jabatan) else "",
                "job_grade_baru":  str(baru_grade).strip()  if pd.notna(baru_grade)   else "",
                "pangkat_baru":    str(baru_pangkat).strip() if pd.notna(baru_pangkat) else "",
                "gaji_pokok_baru": str(baru_gaji).strip()   if pd.notna(baru_gaji)    else "",
                "keterangan":      str(keterangan).strip()  if pd.notna(keterangan)   else "",
                "tanggal_skep":    str(tgl_skep).strip()    if pd.notna(tgl_skep)     else "",
            })

    if not all_rows:
        return pd.DataFrame(), sheet_names

    return pd.DataFrame(all_rows), sheet_names


# ── Helper: Update Google Sheet ───────────────────────────────
def update_master(worksheet, skep_df: pd.DataFrame, col_map: dict, np_col: str, preview_only=False):
    """
    col_map: { nama_kolom_master: field_skep, ... }
    np_col: nama kolom NP di master data
    """
    all_values = worksheet.get_all_values()
    if not all_values:
        return [], "Sheet kosong."

    headers = [h.strip() for h in all_values[0]]

    # Index NP di master
    try:
        np_idx = headers.index(np_col.strip())
    except ValueError:
        return [], f"Kolom '{np_col}' tidak ditemukan di master data."

    # Build NP → row_index map (1-based, row 0 = header)
    np_to_row = {}
    for i, row in enumerate(all_values[1:], start=1):
        np_val = row[np_idx].strip() if np_idx < len(row) else ""
        if np_val:
            np_to_row[np_val] = i

    log = []
    batch_updates = []  # [{range, values}]

    for _, emp in skep_df.iterrows():
        np_val = emp["np"]
        row_idx = np_to_row.get(np_val)

        if row_idx is None:
            log.append({
                "np": np_val, "nama": emp["nama"],
                "status": "tidak_ditemukan",
                "sheet_skep": emp["sheet_skep"],
                "changes": [],
            })
            continue

        master_row = all_values[row_idx]
        changes = []

        for master_col, skep_field in col_map.items():
            if not master_col or not skep_field:
                continue
            try:
                col_idx = headers.index(master_col.strip())
            except ValueError:
                continue

            new_val = emp.get(skep_field, "")
            old_val = master_row[col_idx].strip() if col_idx < len(master_row) else ""

            if new_val and new_val != old_val and new_val.lower() != "nan":
                changes.append({"kolom": master_col, "lama": old_val, "baru": new_val})
                if not preview_only:
                    # gspread cell notation: row_idx+1 (1-based, +1 for header)
                    cell = gspread.utils.rowcol_to_a1(row_idx + 1, col_idx + 1)
                    batch_updates.append({
                        "range": cell,
                        "values": [[new_val]],
                    })

        log.append({
            "np": np_val, "nama": emp["nama"],
            "status": "diupdate" if changes else "tidak_ada_perubahan",
            "sheet_skep": emp["sheet_skep"],
            "keterangan": emp.get("keterangan", ""),
            "changes": changes,
        })

    if not preview_only and batch_updates:
        worksheet.batch_update(batch_updates, value_input_option="USER_ENTERED")

    return log, None


# ════════════════════════════════════════════════════════════════
# ── UI ──────────────────────────────────────────────────────────
# ════════════════════════════════════════════════════════════════

st.title("📋 SKEP Updater")
st.caption("Update master data karyawan di Google Sheets secara otomatis dari file SKEP.")

# ── Sidebar: Konfigurasi ──────────────────────────────────────
with st.sidebar:
    st.header("⚙️ Konfigurasi")

    spreadsheet_url = st.text_input(
        "URL Google Sheets",
        placeholder="https://docs.google.com/spreadsheets/d/...",
        help="Pastikan Service Account sudah diberi akses Editor ke spreadsheet ini.",
    )

    sheet_name = st.text_input(
        "Nama Sheet (tab)",
        value="Master Data",
        help="Nama tab/sheet tempat data karyawan berada.",
    )

    np_col = st.text_input(
        "Nama Kolom NP di Master Data",
        value="NP",
        help="Kolom ini digunakan sebagai kunci untuk mencocokkan karyawan.",
    )

    st.divider()
    st.markdown("**Mapping Kolom**")
    st.caption("Pilih kolom di master data yang sesuai dengan data kondisi baru dari SKEP.")

    col_map_input = {}
    defaults = {
        "Kode Bagan Baru":  "Kode Bagan",
        "Unit Kerja Baru":  "Nama Unit Kerja",
        "Jabatan Baru":     "Jabatan",
        "Job Grade Baru":   "Job Grade",
        "Pangkat Baru":     "Pangkat",
        "Gaji Pokok Baru":  "Gaji Pokok",
    }
    skep_field_map = {
        "Kode Bagan Baru":  "kode_bagan_baru",
        "Unit Kerja Baru":  "unit_kerja_baru",
        "Jabatan Baru":     "jabatan_baru",
        "Job Grade Baru":   "job_grade_baru",
        "Pangkat Baru":     "pangkat_baru",
        "Gaji Pokok Baru":  "gaji_pokok_baru",
    }

    for label, default_master_col in defaults.items():
        val = st.text_input(label, value=default_master_col, key=f"map_{label}")
        if val.strip():
            col_map_input[val.strip()] = skep_field_map[label]

    st.divider()
    st.markdown("**Kolom Tambahan**")
    update_keterangan = st.checkbox("Update kolom Keterangan", value=True)
    ket_col_name = st.text_input("Nama kolom Keterangan", value="Keterangan", disabled=not update_keterangan)
    update_tgl = st.checkbox("Update kolom Tanggal SKEP", value=True)
    tgl_col_name = st.text_input("Nama kolom Tanggal SKEP", value="Tanggal SKEP", disabled=not update_tgl)

    if update_keterangan and ket_col_name:
        col_map_input[ket_col_name] = "keterangan"
    if update_tgl and tgl_col_name:
        col_map_input[tgl_col_name] = "tanggal_skep"


# ── Main Area ─────────────────────────────────────────────────
tab_upload, tab_panduan = st.tabs(["📤 Upload SKEP", "📖 Panduan Setup"])

# ────────────────────────────────────────────────────────────────
with tab_upload:
    uploaded_file = st.file_uploader(
        "Upload file SKEP (.xlsx)",
        type=["xlsx"],
        help="Bisa berisi satu atau banyak sheet SKEP sekaligus.",
    )

    if uploaded_file:
        with st.spinner("Membaca file SKEP..."):
            skep_df, sheet_names = parse_skep(uploaded_file)

        if skep_df.empty:
            st.error("Tidak ada data yang berhasil dibaca dari file SKEP. Periksa format file.")
        else:
            st.success(f"✅ Berhasil membaca **{len(skep_df)} karyawan** dari **{len(sheet_names)} sheet**: {', '.join(sheet_names)}")

            with st.expander("🔍 Preview data SKEP yang dibaca", expanded=False):
                st.dataframe(skep_df[[
                    "sheet_skep", "np", "nama",
                    "kode_bagan_baru", "unit_kerja_baru", "jabatan_baru",
                    "job_grade_baru", "pangkat_baru", "keterangan"
                ]].rename(columns={
                    "sheet_skep": "Sheet", "np": "NP", "nama": "Nama",
                    "kode_bagan_baru": "Kode Bagan Baru", "unit_kerja_baru": "Unit Kerja Baru",
                    "jabatan_baru": "Jabatan Baru", "job_grade_baru": "Job Grade Baru",
                    "pangkat_baru": "Pangkat Baru", "keterangan": "Keterangan",
                }), use_container_width=True)

            st.divider()

            if not spreadsheet_url:
                st.warning("⚠️ Masukkan URL Google Sheets di sidebar untuk melanjutkan.")
            else:
                col1, col2 = st.columns(2)

                with col1:
                    preview_btn = st.button("🔎 Preview Perubahan", use_container_width=True, type="secondary")
                with col2:
                    update_btn = st.button("🚀 Update ke Google Sheets", use_container_width=True, type="primary")

                if preview_btn or update_btn:
                    is_preview = preview_btn and not update_btn

                    try:
                        with st.spinner("Menghubungkan ke Google Sheets..."):
                            worksheet = get_sheet(spreadsheet_url, sheet_name)

                        action_label = "Preview" if is_preview else "Update"
                        with st.spinner(f"{action_label} data..."):
                            log, err = update_master(
                                worksheet, skep_df, col_map_input, np_col,
                                preview_only=is_preview
                            )

                        if err:
                            st.error(f"❌ Error: {err}")
                        else:
                            # Summary
                            n_update    = sum(1 for r in log if r["status"] == "diupdate")
                            n_nochange  = sum(1 for r in log if r["status"] == "tidak_ada_perubahan")
                            n_notfound  = sum(1 for r in log if r["status"] == "tidak_ditemukan")

                            if is_preview:
                                st.info("👁️ Mode Preview — belum ada perubahan yang disimpan ke Google Sheets.")
                            else:
                                st.success("✅ Update selesai!")

                            m1, m2, m3, m4 = st.columns(4)
                            m1.metric("Total Karyawan SKEP", len(log))
                            m2.metric("Akan/Sudah Diupdate", n_update, delta=None)
                            m3.metric("Tidak Ada Perubahan", n_nochange)
                            m4.metric("NP Tidak Ditemukan", n_notfound)

                            # Log detail
                            st.divider()
                            st.subheader("📋 Detail Perubahan")

                            filter_col1, filter_col2 = st.columns(2)
                            with filter_col1:
                                show_filter = st.selectbox(
                                    "Filter status",
                                    ["Semua", "Diupdate", "Tidak Ada Perubahan", "Tidak Ditemukan"],
                                )
                            with filter_col2:
                                search_np = st.text_input("Cari NP / Nama", placeholder="B881 atau Nagita...")

                            status_map = {
                                "Semua": None,
                                "Diupdate": "diupdate",
                                "Tidak Ada Perubahan": "tidak_ada_perubahan",
                                "Tidak Ditemukan": "tidak_ditemukan",
                            }
                            filter_status = status_map[show_filter]

                            for entry in log:
                                if filter_status and entry["status"] != filter_status:
                                    continue
                                if search_np and search_np.lower() not in entry["np"].lower() and search_np.lower() not in entry["nama"].lower():
                                    continue

                                status = entry["status"]
                                tag_class = {
                                    "diupdate": "tag-update",
                                    "tidak_ada_perubahan": "tag-nochange",
                                    "tidak_ditemukan": "tag-notfound",
                                }.get(status, "")
                                tag_label = {
                                    "diupdate": "✅ Diupdate",
                                    "tidak_ada_perubahan": "— Sama",
                                    "tidak_ditemukan": "❌ Tidak Ditemukan",
                                }.get(status, status)

                                with st.expander(
                                    f"**{entry['nama']}** ({entry['np']}) — "
                                    f"{entry.get('sheet_skep','')} — {tag_label}"
                                ):
                                    if status == "tidak_ditemukan":
                                        st.warning("NP karyawan ini tidak ada di master data.")
                                    elif not entry["changes"]:
                                        st.info("Semua kolom sudah sama, tidak ada yang perlu diubah.")
                                    else:
                                        change_df = pd.DataFrame(entry["changes"])
                                        change_df.columns = ["Kolom", "Nilai Lama", "Nilai Baru"]
                                        st.dataframe(change_df, use_container_width=True, hide_index=True)

                                    if entry.get("keterangan"):
                                        st.caption(f"Keterangan SKEP: {entry['keterangan']}")

                    except gspread.exceptions.SpreadsheetNotFound:
                        st.error("❌ Spreadsheet tidak ditemukan. Pastikan URL benar dan Service Account sudah diberi akses.")
                    except gspread.exceptions.WorksheetNotFound:
                        st.error(f"❌ Sheet tab '{sheet_name}' tidak ditemukan di spreadsheet.")
                    except Exception as ex:
                        st.error(f"❌ Error: {ex}")


# ────────────────────────────────────────────────────────────────
with tab_panduan:
    st.markdown("""
## Panduan Setup

### 1. Buat Google Service Account

1. Buka [Google Cloud Console](https://console.cloud.google.com/)
2. Buat project baru (atau pilih yang sudah ada)
3. Aktifkan dua API berikut:
   - **Google Sheets API**
   - **Google Drive API**
4. Pergi ke **IAM & Admin → Service Accounts → Create Service Account**
5. Beri nama (misal: `skep-updater`) lalu klik **Done**
6. Klik service account yang baru dibuat → tab **Keys → Add Key → JSON**
7. File JSON akan terdownload otomatis — **simpan baik-baik!**

---

### 2. Share Google Sheets ke Service Account

1. Buka Google Sheets master data Anda
2. Klik tombol **Share** (pojok kanan atas)
3. Masukkan **email service account** (ada di file JSON, field `client_email`)
4. Beri akses **Editor**
5. Klik Send

---

### 3. Setup Streamlit Secrets

Di Streamlit Community Cloud:

1. Masuk ke [share.streamlit.io](https://share.streamlit.io) → pilih app Anda
2. Klik **Settings → Secrets**
3. Paste konfigurasi berikut (isi dengan isi file JSON service account):

```toml
[gcp_service_account]
type = "service_account"
project_id = "your-project-id"
private_key_id = "your-key-id"
private_key = "-----BEGIN RSA PRIVATE KEY-----\\n...\\n-----END RSA PRIVATE KEY-----\\n"
client_email = "skep-updater@your-project.iam.gserviceaccount.com"
client_id = "123456789"
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "https://www.googleapis.com/robot/v1/metadata/x509/..."
```

> **Catatan:** Untuk `private_key`, salin persis dari file JSON termasuk `\\n` di dalamnya.

---

### 4. Cara Pakai App

1. Isi **URL Google Sheets** di sidebar
2. Isi **nama sheet/tab** yang berisi master data karyawan
3. Isi **nama kolom NP** (nomor pokok) di master data
4. Sesuaikan **mapping kolom** jika nama kolom di master data berbeda
5. Upload file SKEP (.xlsx) — bisa satu file berisi banyak sheet SKEP
6. Klik **Preview Perubahan** untuk melihat apa yang akan diubah
7. Jika sudah yakin, klik **Update ke Google Sheets**

---

### Struktur File SKEP yang Didukung

App ini otomatis mendeteksi struktur file SKEP dengan format:

| Kolom | Keterangan |
|-------|-----------|
| NO. URUT | Nomor urut |
| N A M A | Nama karyawan |
| NOMOR POKOK | NP karyawan (kunci pencocokan) |
| Kondisi Lama | Kode Bagan, Unit Kerja, Jabatan, Job Grade, Pangkat, Gaji |
| **Kondisi Baru** | Kode Bagan, Unit Kerja, Jabatan, Job Grade, Pangkat, Gaji ← yang diupdate |
| KETERANGAN | Jenis mutasi/promosi |
| Tanggal SKEP | Tanggal berlaku |

Satu file SKEP bisa berisi **banyak sheet** (misal: SKEP-56, SKEP-70) dan semuanya akan diproses sekaligus.
""")