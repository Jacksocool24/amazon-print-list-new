"""
Streamlit UI: left PDF HTML mirror + right MAT editor.
"""
import io

import pandas as pd
import streamlit as st
import streamlit.components.v1 as components

from logic_engine import (
    build_image_index,
    excel_text_str,
    export_shipping_list_excel,
    generate_rename_list,
    generate_shipping_list,
    process_data,
    save_to_excel_with_merge,
)
from pdf_generator import build_pdf_html_content, generate_pdf_with_images, pdf_to_images_zip

_INVALID_FILENAME_CHARS = r'\/:*?"<>|'


def sanitize_download_filename(title, default="final_print_list"):
    """Strip unsafe path characters for Windows/macOS download filenames."""
    name = (title or "").strip() or default
    for ch in _INVALID_FILENAME_CHARS:
        name = name.replace(ch, "_")
    return name.strip(" .") or default


def sync_data():
    """Incremental sync from st.session_state.main_editor into cached df + editor_data."""
    cache = st.session_state.get("processed_data_cache", {})
    main_ed = st.session_state.get("main_editor", {})
    if not cache or not isinstance(main_ed, dict):
        return

    df_work = cache.get("df")
    editor_data = cache.get("editor_data")
    edited_rows = main_ed.get("edited_rows", {})
    if df_work is None or editor_data is None or not edited_rows:
        return

    for row_idx, changes in edited_rows.items():
        if not isinstance(changes, dict) or "MAT" not in changes:
            continue
        ri = int(row_idx)
        if ri >= len(df_work):
            continue
        new_mat = excel_text_str(changes.get("MAT")) or "Canvas"
        df_work.at[df_work.index[ri], "材质"] = new_mat
        editor_data.at[ri, "MAT"] = new_mat

    cache["df"] = df_work
    cache["editor_data"] = editor_data
    st.session_state.processed_data_cache = cache
    st.session_state.editor_data = editor_data.copy()


st.set_page_config(page_title="Print List", layout="wide")

st.title("🖨️ 打印清单自动化指挥系统")
st.markdown("---")

if "processed_data_cache" not in st.session_state:
    st.session_state.processed_data_cache = {}
if "show_preview" not in st.session_state:
    st.session_state.show_preview = False
if "show_downloads" not in st.session_state:
    st.session_state.show_downloads = False
if "upload_signature" not in st.session_state:
    st.session_state.upload_signature = None

col1, col2 = st.columns(2)
with col1:
    source_file = st.file_uploader("📂 上传：亚马逊源文件 (XLSX)", type=["xlsx"])
with col2:
    ref_file = st.file_uploader("📋 上传：尺码参考表 (XLSX)", type=["xlsx"])

st.markdown("### 🖼️ 上传产品图片 (可选，支持多选)")
uploaded_images = st.file_uploader(
    "图片命名规范：purchase-date_运单号",
    type=["jpg", "jpeg", "png"],
    accept_multiple_files=True,
)
# 上传完成后的反馈与数量统计
if uploaded_images:
    st.success(
        f"✅ **上传完成！** 系统已成功接收并读取了 **{len(uploaded_images)}** 张产品图片。"
    )

pdf_title = st.text_input("📄 设定 PDF 标题", value="4.20_Order List_XM")
st.caption("⚠️ 提示：PDF导出仅支持全英文环境。如果在输入源中包含中文，PDF可能会出现乱码或墨块。")

images_signature = tuple(
    sorted((img.name, getattr(img, "size", 0)) for img in (uploaded_images or []))
)
current_upload_signature = (
    source_file.name if source_file else None,
    getattr(source_file, "size", 0) if source_file else 0,
    ref_file.name if ref_file else None,
    getattr(ref_file, "size", 0) if ref_file else 0,
    images_signature,
)

if "editor_data" not in st.session_state:
    st.session_state.editor_data = pd.DataFrame(columns=["ID2", "MAT"])

if current_upload_signature != st.session_state.upload_signature:
    st.session_state.upload_signature = current_upload_signature
    st.session_state.processed_data_cache = {}
    st.session_state.show_preview = False
    st.session_state.show_downloads = False
    for k in ("main_editor", "material_preview_editor", "mat_editor_key", "mat_editor"):
        if k in st.session_state:
            del st.session_state[k]


if source_file and ref_file:
    if st.button("🚀 第一步：一键生成并执行排序"):
        source_file_bytes = source_file.getvalue()
        source_df = pd.read_excel(io.BytesIO(source_file_bytes))
        ref_df = pd.read_excel(ref_file)

        with st.spinner("Processing..."):
            image_map = build_image_index(uploaded_images)
            processed_df = process_data(source_df, ref_df)
            processed_df["材质"] = "Canvas"
            processed_df["图片名称"] = (
                processed_df["序号E"].astype(str) + "-" + processed_df["材质"] + "-" + processed_df["尺寸"].astype(str)
            )
            preview_html = build_pdf_html_content(
                processed_df.drop(columns=["Original Row Index"]),
                image_map,
                pdf_title,
            )
            editor_data = pd.DataFrame(
                {
                    "ID2": processed_df["序号E"].astype(str),
                    "MAT": ["Canvas"] * len(processed_df),
                }
            )
            st.session_state.processed_data_cache = {
                "source_file_bytes": source_file_bytes,
                "image_map": image_map,
                "df": processed_df,
                "preview_html": preview_html,
                "editor_data": editor_data,
            }
            st.session_state.editor_data = editor_data.copy()
            for k in ("main_editor", "material_preview_editor", "mat_editor_key", "mat_editor"):
                if k in st.session_state:
                    del st.session_state[k]
            st.session_state.show_preview = True
            st.session_state.show_downloads = False


if st.session_state.get("show_preview"):
    cached = st.session_state.get("processed_data_cache", {})
    draft_df = cached.get("df")
    image_map = cached.get("image_map", {})
    if draft_df is not None and cached.get("editor_data") is not None:
        st.info("💡 战术面板：左侧为 PDF 真实镜像预览，右侧为材质 (MAT) 编辑区。编辑完成后请点击底部刷新按钮。")

        preview_left, preview_right = st.columns([3.5, 1.4], gap="large")
        with preview_left:
            components.html(cached.get("preview_html", ""), height=1000, scrolling=True)
        with preview_right:
            st.data_editor(
                st.session_state.editor_data,
                key="main_editor",
                on_change=sync_data,
                use_container_width=True,
                hide_index=True,
                column_config={
                    "ID2": st.column_config.TextColumn("ID2", disabled=True),
                    "MAT": st.column_config.TextColumn("MAT (在此列输入)", width="large"),
                },
                disabled=["ID2"],
            )
            st.caption("⚠️ 注意：请务必使用纯英文填写材质，输入中文将导致该单元格在 PDF 中变成墨块。")

            if st.button("🔄 确认修改并刷新预览"):
                sync_data()
                refreshed_df = st.session_state.processed_data_cache["df"].copy()
                refreshed_df["图片名称"] = (
                    refreshed_df["序号E"].astype(str) + "-" + refreshed_df["材质"] + "-" + refreshed_df["尺寸"].astype(str)
                )
                st.session_state.processed_data_cache["df"] = refreshed_df
                ed = st.session_state.processed_data_cache["editor_data"].copy()
                ed["MAT"] = refreshed_df["材质"].astype(str).values
                st.session_state.processed_data_cache["editor_data"] = ed
                st.session_state.editor_data = ed
                st.session_state.processed_data_cache["preview_html"] = build_pdf_html_content(
                    refreshed_df.drop(columns=["Original Row Index"]),
                    image_map,
                    pdf_title,
                )
                st.rerun()

            if st.button("✅ 确认数据无误，生成最终下载文件"):
                with st.spinner("Generating final files..."):
                    sync_data()
                    final_df = st.session_state.processed_data_cache["df"].copy()
                    final_df["图片名称"] = (
                        final_df["序号E"].astype(str) + "-" + final_df["材质"] + "-" + final_df["尺寸"].astype(str)
                    )
                    excel_data = save_to_excel_with_merge(final_df, cached["source_file_bytes"], image_map)
                    rename_list_data = generate_rename_list(final_df)
                    shipping_list_data = export_shipping_list_excel(
                        generate_shipping_list(final_df, pdf_title)
                    )
                    safe_pdf_title = sanitize_download_filename(pdf_title)
                    with st.spinner("Generating PDF..."):
                        pdf_data = generate_pdf_with_images(
                            final_df.drop(columns=["Original Row Index"]),
                            image_map,
                            pdf_title,
                        )
                    with st.spinner("Generating print list images..."):
                        images_zip_data = pdf_to_images_zip(pdf_data, safe_pdf_title)

                    st.session_state.processed_data_cache["excel_data"] = excel_data
                    st.session_state.processed_data_cache["rename_list_data"] = rename_list_data
                    st.session_state.processed_data_cache["shipping_list_data"] = shipping_list_data
                    st.session_state.processed_data_cache["pdf_data"] = pdf_data
                    st.session_state.processed_data_cache["images_zip_data"] = images_zip_data
                    st.session_state.show_downloads = True

            if st.session_state.get("show_downloads"):
                st.markdown("---")
                st.success("🎉 生成完毕！请点击下方按钮下载：")
                cached = st.session_state.get("processed_data_cache", {})
                safe_pdf_title = sanitize_download_filename(pdf_title)
                if cached.get("excel_data"):
                    st.download_button(
                        label="📥 下载最终打印清单 (Excel)",
                        data=cached["excel_data"],
                        file_name=f"{safe_pdf_title}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                if cached.get("pdf_data"):
                    st.download_button(
                        label="📥 下载最终打印清单 (PDF)",
                        data=cached["pdf_data"],
                        file_name=f"{safe_pdf_title}.pdf",
                        mime="application/pdf",
                        use_container_width=True,
                    )
                if cached.get("images_zip_data"):
                    st.download_button(
                        label="🖼️ 下载打印清单图片(ZIP)",
                        data=cached["images_zip_data"],
                        file_name=f"{safe_pdf_title}_images.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )
                if cached.get("rename_list_data"):
                    st.download_button(
                        label="📥 下载图片重命名清单(Execl)",
                        data=cached["rename_list_data"],
                        file_name="图片重命名清单.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                if cached.get("shipping_list_data"):
                    st.download_button(
                        label="📥 下载上传物流单号(Excel)",
                        data=cached["shipping_list_data"],
                        file_name="上传物流单号.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
