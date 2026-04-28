import os
import re
import time
import io
import zipfile
import pandas as pd
import requests
from io import BytesIO
from PIL import Image
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry

import streamlit as st

# ==========================================================
# PAGE CONFIG
# ==========================================================
st.set_page_config(
    page_title="Image Batch Downloader",
    page_icon="🖼️",
    layout="wide",
)

st.title("🖼️ Product Image Batch Downloader")
st.markdown("Upload your Excel file with product items and image URLs. Images will be downloaded, converted to JPG, and packaged into a ZIP for download.")

# ==========================================================
# TEMPLATE DOWNLOAD
# ==========================================================
def build_template() -> bytes:
    template_data = {
        "Item": ["FSIAW7350-7D", "FSIAW7350-7-5D", "FSIAW7350-10-5EE"],
        "Image 1": ["https://example.com/img1.png", "https://example.com/img2.png", "https://example.com/img3.png"],
        "Image 2": ["https://example.com/img1b.png", "", ""],
        "Image 3": ["", "", ""],
    }
    df_template = pd.DataFrame(template_data)
    buf = BytesIO()
    df_template.to_excel(buf, index=False)
    return buf.getvalue()

with st.expander("📥 Need a template?"):
    st.markdown(
        "Download the template below. Fill in the **Item** column with your product SKUs "
        "and paste image URLs into the **Image 1**, **Image 2**, ... columns. Leave cells blank if there's no image for that position."
    )
    st.download_button(
        label="⬇️ Download Excel Template",
        data=build_template(),
        file_name="image_downloader_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

# ==========================================================
# SETTINGS SIDEBAR
# ==========================================================
with st.sidebar:
    st.header("⚙️ Settings")
    IMAGE_COL_PREFIX = st.text_input("Image column prefix", value="Image ")
    ITEM_COL = st.text_input("Item column name", value="Item")
    MAX_IMAGES = st.number_input("Max image columns to scan", min_value=1, max_value=20, value=10)
    POLITENESS_DELAY = st.slider("Delay between downloads (seconds)", 0.0, 2.0, 0.1, step=0.05)

# ==========================================================
# HELPERS
# ==========================================================
def get_base(item: str) -> str:
    return str(item).strip()

def is_blank(v) -> bool:
    if v is None:
        return True
    if isinstance(v, float) and pd.isna(v):
        return True
    s = str(v).strip()
    return s == "" or s.lower() in {"nan", "none", "null"}

def make_session() -> requests.Session:
    s = requests.Session()
    retry = Retry(
        total=6,
        backoff_factor=0.7,
        status_forcelist=[429, 500, 502, 503, 504],
        allowed_methods=["GET"]
    )
    adapter = HTTPAdapter(max_retries=retry, pool_connections=20, pool_maxsize=20)
    s.mount("http://", adapter)
    s.mount("https://", adapter)
    s.headers.update({"User-Agent": "Mozilla/5.0"})
    return s

def download_and_convert_to_jpg(session: requests.Session, url: str) -> bytes:
    """
    Downloads an image (png/jpg/etc) and returns it as JPG bytes.
    Handles transparency by compositing onto a white background.
    """
    r = session.get(url, timeout=30)
    r.raise_for_status()

    img = Image.open(BytesIO(r.content))

    # Normalize EXIF orientation
    try:
        exif = img.getexif()
        orientation = exif.get(274)
        if orientation == 3:
            img = img.rotate(180, expand=True)
        elif orientation == 6:
            img = img.rotate(270, expand=True)
        elif orientation == 8:
            img = img.rotate(90, expand=True)
    except Exception:
        pass

    # Flatten transparency onto white background
    has_alpha = (
        img.mode in ("RGBA", "LA") or
        (img.mode == "P" and "transparency" in img.info)
    )
    if has_alpha:
        img = img.convert("RGBA")
        white_bg = Image.new("RGBA", img.size, (255, 255, 255, 255))
        img = Image.alpha_composite(white_bg, img).convert("RGB")
    else:
        img = img.convert("RGB")

    buf = BytesIO()
    img.save(buf, format="JPEG", quality=95, subsampling=0, progressive=False, optimize=False)
    return buf.getvalue()

# ==========================================================
# FILE UPLOAD
# ==========================================================
uploaded_file = st.file_uploader("Upload your Excel file (.xlsx)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file)
        df.columns = [c.strip() for c in df.columns]
    except Exception as e:
        st.error(f"Failed to read Excel file: {e}")
        st.stop()

    if ITEM_COL not in df.columns:
        st.error(f"Column '{ITEM_COL}' not found. Available columns: {list(df.columns)}")
        st.stop()

    image_cols = [
        f"{IMAGE_COL_PREFIX}{i}"
        for i in range(1, int(MAX_IMAGES) + 1)
        if f"{IMAGE_COL_PREFIX}{i}" in df.columns
    ]

    if not image_cols:
        st.error(f"No image columns found. Expected headers like '{IMAGE_COL_PREFIX}1', '{IMAGE_COL_PREFIX}2', ...")
        st.stop()

    st.success(f"✅ Loaded {len(df)} rows | Item column: **{ITEM_COL}** | Image columns found: **{len(image_cols)}**")

    with st.expander("Preview data"):
        st.dataframe(df.head(10), use_container_width=True)

    # Clean URLs
    for c in image_cols:
        df[c] = df[c].astype(str).str.strip()
        df.loc[df[c].isin(["", "nan", "NaN", "None", "null", "NULL"]), c] = pd.NA

    df["Base"] = df[ITEM_COL].apply(get_base)

    # Build download plan
    plan = {}
    mismatches = []
    for _, row in df.iterrows():
        base = row["Base"]
        for pos, col in enumerate(image_cols, start=1):
            url = row[col]
            if is_blank(url):
                continue
            key = (base, pos)
            url_s = str(url).strip()
            if key not in plan:
                plan[key] = url_s
            else:
                if plan[key] != url_s:
                    mismatches.append((base, pos, plan[key], url_s))

    st.info(f"📋 Unique images to download: **{len(plan)}**")

    if mismatches:
        with st.expander(f"⚠️ {len(mismatches)} URL mismatches detected (keeping first URL per base+position)"):
            for base, pos, first_url, other_url in mismatches[:20]:
                st.markdown(f"- **{base}** Image {pos}: `{first_url}` vs `{other_url}`")

    # ==========================================================
    # RUN DOWNLOAD
    # ==========================================================
    if st.button("🚀 Start Download", type="primary"):

        session = make_session()
        downloaded = 0
        failed = 0
        failures = []

        zip_buffer = BytesIO()
        sorted_plan = sorted(plan.items(), key=lambda x: (x[0][0], x[0][1]))

        progress_bar = st.progress(0)
        status_text = st.empty()
        total = len(sorted_plan)

        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
            for i, ((base, pos), url) in enumerate(sorted_plan):
                filename = f"{base}__{pos}.jpg"
                status_text.text(f"Downloading {i+1}/{total}: {filename}")

                try:
                    jpg_bytes = download_and_convert_to_jpg(session, url)
                    zf.writestr(filename, jpg_bytes)
                    downloaded += 1
                    time.sleep(POLITENESS_DELAY)
                except Exception as e:
                    failed += 1
                    failures.append((filename, url, str(e)))

                progress_bar.progress((i + 1) / total)

        status_text.empty()
        progress_bar.empty()

        st.success(f"✅ Done! Downloaded: **{downloaded}** | Failed: **{failed}**")

        if failures:
            with st.expander(f"❌ {len(failures)} failed downloads"):
                for fname, url, err in failures:
                    st.markdown(f"- **{fname}**: {err}  \n  `{url}`")

        # ==========================================================
        # OUTPUT EXCEL
        # ==========================================================
        out = pd.DataFrame()
        out["Item"] = df[ITEM_COL].astype(str).str.strip()
        out["Base"] = df["Base"].astype(str).str.strip()
        for pos, col in enumerate(image_cols, start=1):
            def fname(r, col=col, pos=pos):
                if is_blank(r[col]):
                    return ""
                return f"{r['Base']}__{pos}.jpg"
            out[f"Image__{pos}"] = df.apply(fname, axis=1)

        excel_buffer = BytesIO()
        out.to_excel(excel_buffer, index=False)
        excel_buffer.seek(0)

        # Add output excel into zip too
        zip_buffer.seek(0)
        # Re-open zip to add excel (append mode)
        with zipfile.ZipFile(zip_buffer, "a") as zf:
            zf.writestr("output_mapping.xlsx", excel_buffer.read())

        zip_buffer.seek(0)

        col1, col2 = st.columns(2)
        with col1:
            st.download_button(
                label="⬇️ Download All Images + Mapping (ZIP)",
                data=zip_buffer,
                file_name="images_output.zip",
                mime="application/zip",
            )
        with col2:
            excel_buffer.seek(0)
            st.download_button(
                label="⬇️ Download Mapping Excel Only",
                data=excel_buffer,
                file_name="output_mapping.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
