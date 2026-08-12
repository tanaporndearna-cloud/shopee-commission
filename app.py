# -*- coding: utf-8 -*-
"""
Shopee Commission Calculator — Streamlit App
TRC Motorsport
"""
import streamlit as st
import tempfile, os, io, datetime, traceback
import requests

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="ค่าคอม SHOPEE | TRC Motorsport",
    page_icon="🛒",
    layout="wide",
)

st.title("🛒 คำนวณค่าคอมมิชชั่น SHOPEE")
st.caption("TRC Motorsport — powered by Claude")

# ── Constants ─────────────────────────────────────────────────────────────────
STOCK_SHEET_ID    = "1p4-DX-Sq-Mvo_FNAkYtGMkdEBvPDCuPM_sZ_b-_Tm-Y"
OUTPUT_FOLDER_ID  = "1ta-8NrKxcOlDO6MJwhw96Y8ifZ1RmgVs"

# ── Google Drive helper ────────────────────────────────────────────────────────
def get_drive_service():
    """สร้าง Google Drive service จาก service account ใน Streamlit secrets"""
    from google.oauth2.service_account import Credentials
    from googleapiclient.discovery import build

    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=["https://www.googleapis.com/auth/drive"],
    )
    return build("drive", "v3", credentials=creds, cache_discovery=False)


def download_stock_csv() -> str | None:
    """ดาวน์โหลด STOCK จาก Google Sheets → return path ของ CSV ชั่วคราว"""
    # ลอง export URL ก่อน (ถ้า sheet เปิดสาธารณะ)
    export_url = f"https://docs.google.com/spreadsheets/d/{STOCK_SHEET_ID}/export?format=csv"
    try:
        r = requests.get(export_url, timeout=30)
        if r.status_code == 200 and len(r.content) > 500:
            tmp = tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode='wb')
            tmp.write(r.content)
            tmp.close()
            return tmp.name
    except Exception:
        pass

    # fallback: ใช้ Drive API (ต้องมี secrets)
    try:
        from googleapiclient.http import MediaIoBaseDownload
        svc = get_drive_service()
        req = svc.files().export_media(fileId=STOCK_SHEET_ID, mimeType="text/csv")
        buf = io.BytesIO()
        dl = MediaIoBaseDownload(buf, req)
        done = False
        while not done:
            _, done = dl.next_chunk()
        tmp = tempfile.NamedTemporaryFile(suffix=".csv", delete=False, mode='wb')
        tmp.write(buf.getvalue())
        tmp.close()
        return tmp.name
    except Exception as e:
        st.warning(f"⚠️ ดาวน์โหลด STOCK ไม่ได้: {e}\n\nถ้าต้องการให้ดึง STOCK อัตโนมัติ กรุณาอัปโหลดไฟล์ STOCK.xlsx แทน")
        return None


def upload_to_drive(file_bytes: bytes, filename: str, folder_id: str) -> str:
    """อัปโหลดไฟล์ไปยัง Google Drive → return web link
    วิธี: อัปโหลดโดยไม่ระบุ parent ก่อน แล้ว move เข้า folder
    เพื่อหลีกเลี่ยง storageQuotaExceeded ของ service account
    """
    from googleapiclient.http import MediaIoBaseUpload
    svc = get_drive_service()
    media = MediaIoBaseUpload(
        io.BytesIO(file_bytes),
        mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        resumable=True,
    )
    # สร้างไฟล์ใน folder โดยตรง พร้อม supportsAllDrives
    meta = {"name": filename, "parents": [folder_id]}
    f = svc.files().create(
        body=meta,
        media_body=media,
        fields="id,webViewLink",
        supportsAllDrives=True,
    ).execute()
    file_id = f.get("id")
    # ทำให้ anyone with link สามารถดูได้
    try:
        svc.permissions().create(
            fileId=file_id,
            body={"type": "anyone", "role": "reader"},
            supportsAllDrives=True,
        ).execute()
    except Exception:
        pass
    return f.get("webViewLink", f"https://drive.google.com/file/d/{file_id}/view")


# ── Sidebar — เลือกเดือน ──────────────────────────────────────────────────────
with st.sidebar:
    st.header("⚙️ ตั้งค่า")
    now = datetime.date.today()
    month = st.selectbox("เดือน", list(range(1, 13)),
                         index=now.month - 1,
                         format_func=lambda m: f"{m:02d}")
    year_be = st.number_input("ปี (พ.ศ.)", value=now.year + 543, step=1, min_value=2560)
    year_ce = int(year_be) - 543
    folder_name = f"SP{str(year_ce)[2:]}{month:02d}"
    st.info(f"โฟลเดอร์: **{folder_name}**")

    st.divider()
    st.markdown("**ลำดับการทำงาน**")
    st.markdown("""
1. ดึง STOCK จาก Google Drive
2. รัน `build_combined.py`
3. รัน `build_analysis.py`
4. อัปโหลด → Google Drive
""")

# ── Main — อัปโหลดไฟล์ ────────────────────────────────────────────────────────
st.subheader("📁 อัปโหลดไฟล์")

col1, col2 = st.columns(2)
with col1:
    erp_file = st.file_uploader(
        "Data ERP.xlsx *",
        type=["xlsx"],
        help="ไฟล์ Data ERP จาก eFlowsys",
    )
    shopee_files = st.file_uploader(
        "Data SHOPEE (หลายไฟล์ได้)",
        type=["xlsx"],
        accept_multiple_files=True,
        help="ชื่อไฟล์ต้องขึ้นต้นด้วย SP-MAFIA / SP-TRC / SP-UTOPIA / SP-FREEROAD / SP-WORKFORCE",
    )

with col2:
    abbhsp_files = st.file_uploader(
        "ABBHSP (หลายไฟล์ได้)",
        type=["xlsx"],
        accept_multiple_files=True,
        help="เช่น ABBHSP04.xlsx, ABBHSP06.xlsx",
    )
    fin_files = st.file_uploader(
        "การเงิน SHOPEE (หลายไฟล์ได้)",
        type=["xlsx"],
        accept_multiple_files=True,
        help="ชื่อไฟล์ต้องขึ้นต้นด้วย การเงิน-MAFIA / การเงิน-TRC / การเงิน-UTOPIA / การเงิน-FREEROAD / การเงิน-WORKFORCE",
    )

# optional STOCK override
with st.expander("🗂️ อัปโหลด STOCK.xlsx (ถ้าไม่ต้องการดึงจาก Drive)"):
    stock_override = st.file_uploader("STOCK.xlsx (optional)", type=["xlsx"])

st.divider()

# ── Process button ────────────────────────────────────────────────────────────
if st.button("🚀 คำนวณค่าคอม", type="primary", use_container_width=True):
    if not erp_file:
        st.error("❌ กรุณาอัปโหลดไฟล์ Data ERP ก่อนค่ะ")
        st.stop()

    log_area = st.empty()
    logs = []

    def log(msg):
        logs.append(msg)
        log_area.code("\n".join(logs), language="")

    with st.spinner("กำลังประมวลผล..."):
        try:
            with tempfile.TemporaryDirectory() as tmpdir:
                base = os.path.join(tmpdir, folder_name)
                os.makedirs(base)

                # 1. STOCK
                stock_csv_path = None
                if stock_override:
                    stock_dir = os.path.join(base, "STOCK")
                    os.makedirs(stock_dir)
                    spath = os.path.join(stock_dir, "STOCK.xlsx")
                    with open(spath, "wb") as f:
                        f.write(stock_override.getvalue())
                    log("✅ STOCK: ใช้ไฟล์ที่อัปโหลด")
                else:
                    log("⏳ กำลังดึง STOCK จาก Google Drive...")
                    stock_csv_path = download_stock_csv()
                    if stock_csv_path:
                        import pandas as pd
                        n = len(pd.read_csv(stock_csv_path))
                        log(f"✅ STOCK: {n:,} rows")
                    else:
                        log("⚠️ ไม่มี STOCK — ดำเนินการต่อโดยไม่มีข้อมูล STOCK")

                # 2. Data ERP
                erp_dir = os.path.join(base, "Data ERP")
                os.makedirs(erp_dir)
                with open(os.path.join(erp_dir, "Data ERP.xlsx"), "wb") as f:
                    f.write(erp_file.getvalue())
                log(f"✅ Data ERP: {erp_file.name}")

                # 3. Data SHOPEE
                shopee_dir = os.path.join(base, "Data SHOPEE")
                os.makedirs(shopee_dir)
                for sf in shopee_files:
                    fpath = os.path.join(shopee_dir, sf.name)
                    with open(fpath, "wb") as f:
                        f.write(sf.getvalue())
                    log(f"  📄 {sf.name}")
                if shopee_files:
                    log(f"✅ Data SHOPEE: {len(shopee_files)} ไฟล์")

                # 4. ABBHSP
                abbhsp_dir = os.path.join(base, "ABBHSP")
                os.makedirs(abbhsp_dir)
                for af in abbhsp_files:
                    fpath = os.path.join(abbhsp_dir, af.name)
                    with open(fpath, "wb") as f:
                        f.write(af.getvalue())
                    log(f"  📄 {af.name}")
                if abbhsp_files:
                    log(f"✅ ABBHSP: {len(abbhsp_files)} ไฟล์")

                # 5. การเงิน SHOPEE
                fin_dir = os.path.join(base, "การเงิน SHOPEE")
                os.makedirs(fin_dir)
                for ff in fin_files:
                    fpath = os.path.join(fin_dir, ff.name)
                    with open(fpath, "wb") as f:
                        f.write(ff.getvalue())
                    log(f"  📄 {ff.name}")
                if fin_files:
                    log(f"✅ การเงิน SHOPEE: {len(fin_files)} ไฟล์")

                # 6. รัน build_combined
                out_path = os.path.join(base, "input_combined.xlsx")
                log("\n⏳ กำลังรัน build_combined...")
                import sys, io as _io
                old_stdout = sys.stdout
                sys.stdout = buf = _io.StringIO()
                try:
                    from build_combined import run as run_combined
                    run_combined(base, out_path, stock_csv_path)
                finally:
                    sys.stdout = old_stdout
                combined_log = buf.getvalue()
                for line in combined_log.splitlines():
                    log(f"  {line}")
                log("✅ build_combined เสร็จแล้ว")

                # 7. รัน build_analysis
                log("\n⏳ กำลังรัน build_analysis...")
                sys.stdout = buf = _io.StringIO()
                try:
                    from build_analysis import run as run_analysis
                    run_analysis(out_path, out_path)
                finally:
                    sys.stdout = old_stdout
                analysis_log = buf.getvalue()
                for line in analysis_log.splitlines():
                    log(f"  {line}")
                log("✅ build_analysis เสร็จแล้ว")

                # 8. อ่านไฟล์ผลลัพธ์
                with open(out_path, "rb") as f:
                    result_bytes = f.read()

                output_filename = f"input_combined_{folder_name}.xlsx"

                # 9. อัปโหลดไป Drive
                log(f"\n⏳ กำลังอัปโหลด {output_filename} ไปยัง Google Drive...")
                try:
                    drive_link = upload_to_drive(result_bytes, output_filename, OUTPUT_FOLDER_ID)
                    log(f"✅ อัปโหลดสำเร็จ!")
                    st.success(f"✅ อัปโหลดไปยัง Google Drive สำเร็จ!")
                    st.markdown(f"### [📄 เปิดไฟล์ใน Google Drive]({drive_link})")
                except Exception as e:
                    log(f"⚠️ อัปโหลด Drive ไม่ได้: {e}")
                    st.warning(f"อัปโหลดไป Drive ไม่ได้ค่ะ ({e}) — ดาวน์โหลดโดยตรงแทนได้เลย")

                # 10. Download button
                st.download_button(
                    label="⬇️ ดาวน์โหลดไฟล์",
                    data=result_bytes,
                    file_name=output_filename,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                # Cleanup temp stock CSV
                if stock_csv_path and os.path.exists(stock_csv_path):
                    os.unlink(stock_csv_path)

        except Exception as e:
            st.error(f"❌ เกิดข้อผิดพลาด: {e}")
            st.code(traceback.format_exc())

# ── Footer ─────────────────────────────────────────────────────────────────────
st.divider()
st.caption("TRC Motorsport © 2026 | shopee-commission v1.0")
