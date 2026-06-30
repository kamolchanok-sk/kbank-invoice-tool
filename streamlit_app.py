# -*- coding: utf-8 -*-
"""
streamlit_app.py — เว็บแอประบบจัดการใบกำกับภาษี KBank
รันด้วยคำสั่ง: streamlit run streamlit_app.py
หรือ deploy ฟรีบน https://share.streamlit.io
"""

import io
import tempfile
import shutil
import datetime
from pathlib import Path

import streamlit as st
import openpyxl
import pandas as pd

from core import process_pdfs, strip_pdf_password

# core.py ไม่มี DEFAULT_SHEET ระดับโมดูล จึงกำหนดในนี้
DEFAULT_SHEET = "KBANK"

st.set_page_config(
    page_title="ระบบจัดการใบกำกับภาษี KBank",
    page_icon="📄",
    layout="wide",
)

# ───────────────────────── CSS ปรับธีม ─────────────────────────
st.markdown("""
<style>
    .stApp { background-color: #F1F5F9; }
    div[data-testid="stMetricValue"] { font-size: 1.4rem; }
    .main-header {
        background: linear-gradient(135deg, #0B6E7A 0%, #028090 100%);
        padding: 1.2rem 1.5rem; border-radius: 10px; margin-bottom: 1.2rem;
    }
    .main-header h1 { color: white; font-size: 1.3rem; margin: 0; }
    .main-header p { color: #A7D9D4; font-size: 0.85rem; margin: 0.2rem 0 0 0; }
    .info-box {
        background: #E0F7FA; border: 1px solid #80DEEA; border-radius: 8px;
        padding: 0.9rem 1.1rem; margin-bottom: 1rem;
    }
    .success-box {
        background: #ECFDF5; border: 1px solid #6EE7B7; border-radius: 8px;
        padding: 0.9rem 1.1rem;
    }
    div.stButton > button {
        border-radius: 8px; font-weight: 600; padding: 0.5rem 1.2rem;
    }
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="main-header">
    <h1>📄 ระบบจัดการใบกำกับภาษี KBank</h1>
    <p>v3.0 (เว็ปแอพลิเคชั่น) — ระบบอ่าน PDF ใบกำกับภาษี กรอกลง Excel อัตโนมัติ และการปลดรหัสผ่าน PDF</p>
</div>
""", unsafe_allow_html=True)

tab1, tab2 = st.tabs(["📊  กรอกลง Excel", "🔓  ปลดรหัส PDF"])


# ════════════════════════════════════════════════════════════
#  TAB 1 — กรอกลง Excel
# ════════════════════════════════════════════════════════════
with tab1:
    st.markdown("##### ① เลือกไฟล์ PDF ใบกำกับภาษี")
    pdf_files = st.file_uploader(
        "ลากไฟล์ PDF มาวาง หรือกดเลือกไฟล์ (เลือกได้หลายไฟล์พร้อมกัน)",
        type=["pdf"], accept_multiple_files=True, key="excel_pdfs"
    )
    if pdf_files:
        st.caption(f"✓ เลือกแล้ว {len(pdf_files)} ไฟล์: " + ", ".join(f.name for f in pdf_files))

    st.markdown("##### ② ตั้งค่า")
    col1, col2 = st.columns(2)
    with col1:
        password = st.text_input("🔑 รหัสผ่าน PDF", type="password",
                                  help="ว่างเปล่าถ้าไฟล์ไม่ได้ล็อก", key="excel_pw")
        sheet_name = st.text_input("📋 ชื่อชีท Excel", value=DEFAULT_SHEET, key="excel_sheet")
    with col2:
        excel_file = st.file_uploader("📊 ไฟล์ Excel ปลายทาง (.xlsx)",
                                       type=["xlsx", "xlsm"], key="excel_target")

    st.markdown("##### ③ ประมวลผล")
    run = st.button("▶  เริ่มประมวลผล", type="primary", use_container_width=False, key="run_excel")

    if run:
        if not pdf_files:
            st.warning("กรุณาเลือกไฟล์ PDF ก่อน")
        elif not excel_file:
            st.warning("กรุณาอัปโหลดไฟล์ Excel ปลายทาง")
        else:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)

                # เขียน Excel ที่อัปโหลดลง temp
                excel_path = tmpdir / excel_file.name
                excel_path.write_bytes(excel_file.getvalue())

                # เขียน PDF ทุกไฟล์ลง temp
                pdf_paths = []
                for f in pdf_files:
                    p = tmpdir / f.name
                    p.write_bytes(f.getvalue())
                    pdf_paths.append(p)

                progress_bar = st.progress(0, text="กำลังเริ่มประมวลผล...")

                def progress_cb(cur, total, msg):
                    pct = cur / total if total else 0
                    progress_bar.progress(min(pct, 1.0), text=msg)

                try:
                    result = process_pdfs(
                        pdf_paths=pdf_paths,
                        excel_path=excel_path,
                        sheet_name=sheet_name.strip(),
                        password=password,
                        progress_cb=progress_cb,
                    )
                    progress_bar.progress(1.0, text="เสร็จสิ้น")

                    s, f = result["success"], result["failed"]
                    if f == 0:
                        st.success(f"✅ ประมวลผลครบ {s} ไฟล์ สำเร็จทั้งหมด")
                    else:
                        st.warning(f"สำเร็จ {s} ไฟล์  |  ❌ ล้มเหลว {f} ไฟล์ — ดูรายละเอียดด้านล่าง")

                    # ตารางผลลัพธ์
                    df = pd.DataFrame(result["results"])
                    rename_map = {
                        "file": "ไฟล์", "status": "สถานะ", "date": "วันที่",
                        "item": "จำนวน", "amount": "ยอดเงิน", "fee": "ค่าธรรมเนียม",
                        "vat": "VAT", "net": "สุทธิ", "detail": "หมายเหตุ",
                    }
                    cols_order = ["file", "status", "date", "item", "amount", "fee", "vat", "net", "detail"]
                    cols_order = [c for c in cols_order if c in df.columns]
                    df = df[cols_order].rename(columns=rename_map)
                    st.dataframe(df, use_container_width=True, hide_index=True)

                    # ปุ่มดาวน์โหลด Excel ที่อัปเดตแล้ว
                    updated_bytes = excel_path.read_bytes()
                    st.download_button(
                        label="⬇️ ดาวน์โหลดไฟล์ Excel ที่อัปเดตแล้ว",
                        data=updated_bytes,
                        file_name=excel_file.name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                    )

                    # ปุ่มดาวน์โหลด backup (เผื่ออยากเก็บ)
                    backup_bytes = result["backup_path"].read_bytes()
                    st.download_button(
                        label="🗄️ ดาวน์โหลดไฟล์ Excel ต้นฉบับ (ก่อนแก้ไข / backup)",
                        data=backup_bytes,
                        file_name=result["backup_path"].name,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )

                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")


# ════════════════════════════════════════════════════════════
#  TAB 2 — ปลดรหัส PDF
# ════════════════════════════════════════════════════════════
with tab2:
    st.markdown("""
    <div class="info-box">
        🔓 <b>ปลดรหัสผ่านออกจาก PDF</b> แล้วดาวน์โหลดไฟล์ใหม่ที่ไม่มีรหัสกลับมาทันที<br>
        <span style="color:#64748B; font-size:0.85rem;">
        ชื่อไฟล์ใหม่จะเป็น &nbsp;<code>ชื่อเดิม_unlocked.pdf</code>
        </span>
    </div>
    """, unsafe_allow_html=True)

    st.markdown("##### ① เลือกไฟล์ PDF ที่ต้องการปลดรหัส")
    unlock_files = st.file_uploader(
        "ลากไฟล์ PDF ที่ล็อกรหัสมาวาง หรือกดเลือกไฟล์ (เลือกได้หลายไฟล์พร้อมกัน)",
        type=["pdf"], accept_multiple_files=True, key="unlock_pdfs"
    )
    if unlock_files:
        st.caption(f"✓ เลือกแล้ว {len(unlock_files)} ไฟล์: " + ", ".join(f.name for f in unlock_files))

    st.markdown("##### ② ตั้งค่า")
    unlock_password = st.text_input("🔑 รหัสผ่าน PDF", type="password",
                                     help="รหัสผ่านที่ใช้ปลดล็อก PDF", key="unlock_pw")

    st.markdown("##### ③ ปลดรหัส")
    run_unlock = st.button("🔓  ปลดรหัสผ่านทุกไฟล์", type="primary", key="run_unlock")

    if run_unlock:
        if not unlock_files:
            st.warning("กรุณาเลือกไฟล์ PDF ก่อน")
        elif not unlock_password:
            st.warning("กรุณากรอกรหัสผ่าน PDF")
        else:
            with tempfile.TemporaryDirectory() as tmpdir:
                tmpdir = Path(tmpdir)
                in_dir = tmpdir / "in"
                out_dir = tmpdir / "out"
                in_dir.mkdir()

                pdf_paths = []
                for f in unlock_files:
                    p = in_dir / f.name
                    p.write_bytes(f.getvalue())
                    pdf_paths.append(p)

                progress_bar = st.progress(0, text="กำลังเริ่มปลดรหัส...")

                def progress_cb(cur, total, msg):
                    pct = cur / total if total else 0
                    progress_bar.progress(min(pct, 1.0), text=msg)

                try:
                    result = strip_pdf_password(
                        pdf_paths=pdf_paths,
                        password=unlock_password,
                        output_dir=out_dir,
                        progress_cb=progress_cb,
                    )
                    progress_bar.progress(1.0, text="เสร็จสิ้น")

                    s, f = result["success"], result["failed"]
                    if f == 0:
                        st.success(f"🔓 ปลดรหัสครบ {s} ไฟล์ สำเร็จทั้งหมด")
                    else:
                        st.warning(f"สำเร็จ {s} ไฟล์  |  ❌ ล้มเหลว {f} ไฟล์")

                    df = pd.DataFrame(result["results"])
                    rename_map = {"file": "ไฟล์ต้นฉบับ", "status": "สถานะ",
                                  "output": "ไฟล์ที่บันทึก", "detail": "หมายเหตุ"}
                    df = df.rename(columns=rename_map)
                    st.dataframe(df, use_container_width=True, hide_index=True)

                    # ถ้าสำเร็จมากกว่า 1 ไฟล์ → zip ให้โหลดทีเดียว
                    success_files = [r for r in result["results"] if r["status"].startswith("✅")]
                    if success_files:
                        if len(success_files) == 1:
                            only = success_files[0]
                            file_bytes = (out_dir / only["output"]).read_bytes()
                            st.download_button(
                                label=f"⬇️  ดาวน์โหลด {only['output']}",
                                data=file_bytes,
                                file_name=only["output"],
                                mime="application/pdf",
                                type="primary",
                            )
                        else:
                            import zipfile
                            zip_buf = io.BytesIO()
                            with zipfile.ZipFile(zip_buf, "w") as zf:
                                for r in success_files:
                                    zf.write(out_dir / r["output"], arcname=r["output"])
                            zip_buf.seek(0)
                            st.download_button(
                                label=f"⬇️  ดาวน์โหลดทั้งหมด ({len(success_files)} ไฟล์) เป็น .zip",
                                data=zip_buf.getvalue(),
                                file_name="unlocked_pdfs.zip",
                                mime="application/zip",
                                type="primary",
                            )

                except Exception as e:
                    st.error(f"เกิดข้อผิดพลาด: {e}")


# ───────────────────────── Footer ─────────────────────────
st.markdown("---")
st.caption("ระบบจัดการใบกำกับภาษี")
