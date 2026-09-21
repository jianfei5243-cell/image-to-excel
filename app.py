"""Streamlit 网页版：多图上传 → 中文表格识别 → 汇总 Excel 下载。

部署到 Streamlit Community Cloud 后即可获得一个公网链接，无需自己维护服务器。
"""
from __future__ import annotations

import tempfile
from pathlib import Path

import streamlit as st

from extract import MIN_CONFIDENCE, SUPPORTED_EXTS, build_ocr, build_table_engine, images_to_excel_bytes

st.set_page_config(page_title="图片表格 → Excel", page_icon="📷", layout="wide")

# ---------- 引擎缓存：同一实例只加载一次，后续识别更快 ----------
@st.cache_resource(show_spinner="正在加载中文 OCR 引擎（首次约几秒）…")
def get_ocr():
    return build_ocr()


@st.cache_resource(show_spinner="正在加载表格结构引擎（首次可能需联网下载模型）…")
def get_table_engine():
    return build_table_engine()


# ---------- 顶部说明 ----------
st.title("📷 图片表格识别 → Excel")
st.caption("导入多张表格图片（支持中文 OCR），一键汇总成一个 Excel 文件。")

with st.expander("💡 使用说明", expanded=False):
    st.markdown(
        "- 支持 **PNG / JPG / BMP / TIFF / WEBP**，一次可上传多张。\n"
        "- 每张图片识别出的表格会写成单独的工作表，另附一个「汇总」工作表。\n"
        "- 建议图片清晰、表格边框完整，识别率更高。\n"
        "- 图片仅在本次会话中临时处理，处理完成后即删除，不会保存到服务器。"
    )

# ---------- 上传区 ----------
uploaded = st.file_uploader(
    "上传表格图片（可多选）",
    type=[ext.lstrip(".") for ext in sorted(SUPPORTED_EXTS)],
    accept_multiple_files=True,
)

if not uploaded:
    st.info("👆 请先上传一张或多张表格图片。")
    st.stop()

st.markdown(f"**已选择 {len(uploaded)} 张图片**")

# 可选：预览前几张，帮助确认图片是否正确
with st.expander("🖼 查看已上传的图片（最多显示 6 张）", expanded=False):
    cols = st.columns(3)
    for idx, uploaded_file in enumerate(uploaded[:6]):
        with cols[idx % 3]:
            st.image(uploaded_file, use_container_width=True, caption=uploaded_file.name)

# ---------- 识别选项 ----------
col1, col2, col3 = st.columns(3)
first_row_header = col1.checkbox("首行作为表头", value=False, help="适合单行表头的简单表格；多行表头建议不勾选。")
min_conf = col2.slider("OCR 最低置信度", 0, 99, MIN_CONFIDENCE, help="低于该置信度的文字会被忽略。")
start = col3.button("🚀 开始识别并生成 Excel", type="primary", use_container_width=True)

if not start:
    st.stop()

# ---------- 识别流程 ----------
with tempfile.TemporaryDirectory() as tmpdir:
    paths: list[Path] = []
    for uploaded_file in uploaded:
        path = Path(tmpdir) / uploaded_file.name
        path.write_bytes(uploaded_file.getbuffer())
        paths.append(path)

    progress_bar = st.progress(0.0, text="准备识别…")

    def progress(done: int, total: int, name: str) -> None:
        ratio = (done + 1) / total if total else 1.0
        progress_bar.progress(ratio, text=f"识别 {name}（{done + 1}/{total}）")

    try:
        data = images_to_excel_bytes(
            paths,
            min_confidence=int(min_conf),
            first_row_header=first_row_header,
            ocr=get_ocr(),
            table_engine=get_table_engine(),
            progress=progress,
        )
    except Exception as exc:  # 顶层兜底，避免页面直接报错
        st.session_state["result"] = None
        st.error(f"识别出错：{exc}")
    else:
        st.session_state["result"] = {"data": data, "count": len(uploaded)}
        progress_bar.progress(1.0, text="识别完成")

# ---------- 结果与下载 ----------
result = st.session_state.get("result")
if result:
    st.success(f"识别完成 ✅ 已处理 {result['count']} 张图片")
    st.download_button(
        "⬇️ 下载汇总 Excel",
        data=result["data"],
        file_name="汇总结果.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
    )
