# 图片表格识别 → Excel

导入多张表格图片，自动识别其中的表格内容（支持**中文 OCR**），汇总成一个 Excel 文件。

- ✅ 多图批量导入
- ✅ 中文识别 + 高精度表格结构还原（RapidOCR + RapidTable）
- ✅ 每个表格写为一个工作表，并生成「汇总」工作表
- ✅ 支持多行表头、合并单元格的复杂表格
- ✅ 网页版（在线使用，推荐）+ 桌面版（离线使用）+ 命令行

---

## 在线使用（网页版）

本仓库可直接部署到 [Streamlit Community Cloud](https://streamlit.io/cloud)（免费，公网可访问）。

### 一键部署

1. 把本仓库 fork 到自己的 GitHub 账号。
2. 打开 [streamlit.io/cloud](https://streamlit.io/cloud)，用 GitHub 登录。
3. 点击 **New app** → 选择本仓库 → Main file path 填 `app.py` → **Deploy**。
4. 等待部署完成，即可获得一个 `https://xxx.streamlit.app` 的公网链接，分享给任何人使用。

> 首次访问会加载识别引擎（可能需联网下载约 7MB 的表格模型），之后会自动缓存、速度更快。

### 本地运行网页版

```bash
python3 -m venv .venv
source .venv/bin/activate            # Windows: .venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

浏览器会自动打开 `http://localhost:8501`。

---

## 桌面版（离线使用）

**双击** `启动图片转Excel.command` 即可打开图形窗口（无需敲命令，数据不出本机）。

> 首次双击若提示无法打开，右键该文件 → 「打开」。
> 桌面版需要先按上文「本地运行」创建好 `.venv` 环境。

---

## 命令行

```bash
# 测试：用仓库自带的演示图
python extract.py demo/sample_table.png -o 汇总结果.xlsx

# 正式使用：传图片路径或目录，可混用
python extract.py 图片A.png 图片B.jpg ./截图目录 -o 汇总结果.xlsx
```

可选参数：

- `--min-confidence 30`：OCR 最低置信度（0–99）
- `--header`：把每张表的第一行作为表头（默认关闭）

---

## 技术栈

- `RapidOCR`：中文 OCR（ONNX，CPU 友好）
- `RapidTable`：表格结构识别（PaddleOCR SLANet-Plus 的 ONNX 版，中文表格精度高）
- `pandas` + `xlsxwriter`：汇总导出 Excel
- `Streamlit`：网页界面

环境要求：Python **3.10 ～ 3.14**（macOS / Windows / Linux 均可）。

---

## 输出说明

- 每张识别出的表格 → 一个工作表，命名 `图片名_序号`
- `汇总` 工作表 → 每张图片的识别状态、工作表名、行列数

## 已知局限

- 拍照变形、严重模糊、低清晰度图片的识别率会下降，建议图片尽量清晰。
- 合并单元格内的多行文字会在同一单元格内拼接显示，属正常现象。
- 对超宽、超密集的表格，个别单元格内容可能有轻微错位，可人工微调。

## 开源许可

本项目采用 [MIT License](LICENSE)。依赖库遵循各自协议（RapidOCR / RapidTable / onnxruntime：Apache-2.0）。
