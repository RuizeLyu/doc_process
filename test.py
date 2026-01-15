import os
import pandas as pd
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn
from pathlib import Path
from docx.enum.text import WD_ALIGN_PARAGRAPH

# ================== 配置 ==================
EXCEL_FILE = "人事管理.xlsx"
OUTPUT_FOLDER = Path(EXCEL_FILE).stem
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# ================== 创建单个 DOCX 文档 ==================
doc = Document()

# === 全局字体设置 ===
style = doc.styles['Normal']
font = style.font
font.name = 'Microsoft YaHei'
font.size = Pt(12)
style._element.rPr.rFonts.set(qn('w:eastAsia'), 'Microsoft YaHei')

# === 主标题：Excel 文件名（去掉扩展名）===
main_title = Path(EXCEL_FILE).stem
title = doc.add_heading(main_title, level=0)
title_run = title.runs[0]
title_run.font.size = Pt(30)
title_run.font.bold = True
title_run.font.color.rgb = RGBColor(0, 0, 0)
title.alignment = WD_ALIGN_PARAGRAPH.LEFT

# ================== 读取并处理每个工作表 ==================
with pd.ExcelFile(EXCEL_FILE) as xls:
    for sheet_name in xls.sheet_names:
        print(f"🔄 处理工作表: {sheet_name}")

        df_raw = pd.read_excel(xls, sheet_name=sheet_name, header=None, dtype=str)
        df_raw = df_raw.fillna("")

        if df_raw.empty:
            print(f"⚠️ 跳过空表: {sheet_name}")
            continue

        nrows, ncols = df_raw.shape
        col_map = {}
        data_start_row = -1

        # === 找表头 ===
        for i in range(nrows):
            row = df_raw.iloc[i]
            non_empty_cols = [(j, str(row[j]).strip()) for j in range(ncols) if str(row[j]).strip() != ""]
            if not non_empty_cols:
                continue
            values = [v for _, v in non_empty_cols]
            indices = [j for j, _ in non_empty_cols]
            target_fields = ["测试步骤", "功能路径", "输入数据/特殊信息", "预期结果"]
            if all(f in values for f in target_fields):
                for f in target_fields:
                    col_map[f] = indices[values.index(f)]
                data_start_row = i + 1
                break

        if not col_map or data_start_row == -1:
            print(f"⚠️ 未找到表头，跳过: {sheet_name}")
            continue

        # === 提取有效数据（严格模式）===
        data_rows = []
        for i in range(data_start_row, nrows):
            row = df_raw.iloc[i]
            try:
                step = str(row[col_map["测试步骤"]]).strip()
                path = str(row[col_map["功能路径"]]).strip()
                input_data = str(row[col_map["输入数据/特殊信息"]]).strip()
                expected = str(row[col_map["预期结果"]]).strip()
            except Exception:
                break
            if not step or not path or not input_data or not expected:
                break
            data_rows.append([step, path, input_data, expected])

        if not data_rows:
            print(f"⚠️ 无有效数据，跳过: {sheet_name}")
            continue

        # === 添加二级标题：Sheet 名称 ===
        sec_title = doc.add_heading(sheet_name, level=1)
        sec_title_run = sec_title.runs[0]
        sec_title_run.font.size = Pt(24)
        sec_title_run.font.bold = True
        sec_title_run.font.color.rgb = RGBColor(0, 0, 0)

        # === 添加每个用例 ===
        for idx, (step, path, input_data, expected) in enumerate(data_rows, 1):
            # 测试步骤
            h_step = doc.add_heading("测试步骤", level=2)
            h_step_run = h_step.runs[0]
            h_step_run.font.size = Pt(18)
            h_step_run.font.bold = True
            h_step_run.font.color.rgb = RGBColor(0, 0, 0)
            doc.add_paragraph(step)

            # 功能路径
            h_path = doc.add_heading("功能路径", level=2)
            h_path_run = h_path.runs[0]
            h_path_run.font.size = Pt(18)
            h_path_run.font.bold = True
            h_path_run.font.color.rgb = RGBColor(0, 0, 0)
            p_path = doc.add_paragraph(path.replace("->", " → "))
            p_path.runs[0].bold = True

            # 输入数据 / 特殊信息
            h_input = doc.add_heading("输入数据 / 特殊信息", level=2)
            h_input_run = h_input.runs[0]
            h_input_run.font.size = Pt(18)
            h_input_run.font.bold = True
            h_input_run.font.color.rgb = RGBColor(0, 0, 0)
            if input_data:
                lines = input_data.split('\n')
                for line in lines:
                    line = line.strip()
                    if line:
                        doc.add_paragraph(line)
            else:
                doc.add_paragraph("-")

            # 预期结果
            h_expected = doc.add_heading("预期结果", level=2)
            h_expected_run = h_expected.runs[0]
            h_expected_run.font.size = Pt(18)
            h_expected_run.font.bold = True
            h_expected_run.font.color.rgb = RGBColor(0, 0, 0)
            doc.add_paragraph(expected)

            # 用例之间加分页（可选，也可只在 Sheet 末尾分页）
            if idx < len(data_rows):
                doc.add_page_break()

        # 每个 Sheet 结束后加分页（避免混在一起）
        doc.add_page_break()

# ================== 保存最终文档 ==================
output_file = os.path.join(OUTPUT_FOLDER, f"{main_title}.docx")
doc.save(output_file)
print(f"\n✅ 已生成合并文档: {output_file}")
print(f"🎉 输出目录: {OUTPUT_FOLDER}")