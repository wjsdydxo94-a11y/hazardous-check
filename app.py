from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
from datetime import datetime, timedelta
import os
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
DATA_FILE = 'inspection_db.csv'

def init_db():
    required_cols = [
        "점검일시", "업체명", "점검자", 
        "건축물구조_상태", "건축물구조_비고", 
        "소방환기_상태", "소방환기_비고", 
        "표지저장_상태", "표지저장_비고", 
        "온습도누출_상태", "온습도누출_비고", 
        "온도", "습도"
    ]
    if not os.path.exists(DATA_FILE):
        df = pd.DataFrame(columns=required_cols)
        df.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
    else:
        df = pd.read_csv(DATA_FILE)
        if not all(col in df.columns for col in ["건축물구조_상태", "온도", "습도"]):
            df_new = pd.DataFrame(columns=required_cols)
            df_new.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')

def generate_monthly_excel(company_name, target_month=None):
    init_db()
    df_db = pd.read_csv(DATA_FILE)
    
    # 업체 필터 적용
    if company_name and company_name != 'all' and '업체명' in df_db.columns:
        df_db = df_db[df_db['업체명'] == company_name]
    
    # 월 필터 적용
    if target_month and target_month != 'all' and '점검일시' in df_db.columns:
        df_db = df_db[df_db['점검일시'].astype(str).str.startswith(target_month)]
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "월별점검대장"
    ws.views.sheetView[0].showGridLines = True

    font_title = Font(name="맑은 고딕", size=16, bold=True)
    font_header = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    font_body = Font(name="맑은 고딕", size=10)

    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_meta = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")

    align_center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    align_left = Alignment(horizontal="left", vertical="center", wrap_text=True)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    display_company = company_name if company_name and company_name != 'all' else "전체 업체"
    display_month = target_month if target_month and target_month != 'all' else "전체 기간"

    # 1. Title
    ws.merge_cells("A1:J1")
    ws["A1"] = f"위험물 저장소 일일점검 관리 대장 ({display_company} / {display_month})"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center
    ws.row_dimensions[1].height = 40

    ws.row_dimensions[2].height = 10

    # 2. Table Headers (대시보드와 동일한 컬럼 구조)
    headers = [
        "점검일시", "업체명", "점검자", 
        "1. 건축물/피뢰", "2. 소방/환기", 
        "3. 표지/저장취급", "4. 온습도/누출", 
        "온도(℃)", "습도(%)", "특이사항/조치내용"
    ]

    ws.row_dimensions[3].height = 28
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=3, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    # 3. Populate rows from filtered df_db
    if df_db.empty:
        ws.row_dimensions[4].height = 25
        ws.merge_cells("A4:J4")
        ws["A4"] = "조건에 해당하는 점검 기록이 없습니다."
        ws["A4"].alignment = align_center
        ws["A4"].font = font_body
    else:
        for idx, row in enumerate(df_db.itertuples(), start=4):
            ws.row_dimensions[idx].height = 22
            
            dt_str = getattr(row, '점검일시', '')
            comp_str = getattr(row, '업체명', '')
            insp_str = getattr(row, '점검자', '')
            
            s1 = getattr(row, '건축물구조_상태', '양호')
            b1 = getattr(row, '건축물구조_비고', '')
            v1_text = f"양호 ({b1})" if b1 and pd.notna(b1) and str(b1).strip() != '' and str(b1) != 'nan' else str(s1)
            
            s2 = getattr(row, '소방환기_상태', '양호')
            b2 = getattr(row, '소방환기_비고', '')
            v2_text = f"양호 ({b2})" if b2 and pd.notna(b2) and str(b2).strip() != '' and str(b2) != 'nan' else str(s2)
            
            s3 = getattr(row, '표지저장_상태', '양호')
            b3 = getattr(row, '표지저장_비고', '')
            v3_text = f"양호 ({b3})" if b3 and pd.notna(b3) and str(b3).strip() != '' and str(b3) != 'nan' else str(s3)
            
            s4 = getattr(row, '온습도누출_상태', '양호')
            b4 = getattr(row, '온습도누출_비고', '')
            v4_text = f"양호 ({b4})" if b4 and pd.notna(b4) and str(b4).strip() != '' and str(b4) != 'nan' else str(s4)
            
            temp_val = getattr(row, '온도', '')
            humid_val = getattr(row, '습도', '')
            
            remarks_list = []
            for item_name, b_val in [('1번', b1), ('2번', b2), ('3번', b3), ('4번', b4)]:
                if b_val and pd.notna(b_val) and str(b_val).strip() != '' and str(b_val) != 'nan':
                    remarks_list.append(f"{item_name}: {b_val}")
            remarks_str = " / ".join(remarks_list) if remarks_list else "-"

            row_vals = [
                dt_str, comp_str, insp_str, 
                v1_text, v2_text, v3_text, v4_text, 
                temp_val, humid_val, remarks_str
            ]

            for col_idx, val in enumerate(row_vals, 1):
                cell = ws.cell(row=idx, column=col_idx, value=val)
                cell.font = font_body
                cell.border = thin_border
                cell.alignment = align_center if col_idx <= 9 else align_left
                if idx % 2 == 1:
                    cell.fill = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

    col_widths = [20, 16, 14, 18, 18, 18, 18, 12, 12, 35]
    for idx, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    report_filename = f'위험물점검_관리대장_{display_company}_{display_month.replace(" ", "_")}.xlsx'
    wb.save(report_filename)
    return report_filename

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/submit', methods=['POST'])
def submit():
    init_db()
    company = request.form.get('company', '아로마솔루션')
    inspector = request.form.get('inspector')
    
    item1_status = request.form.get('item1_status', '양호')
    item1_remark = request.form.get('item1_remark', '')
    item2_status = request.form.get('item2_status', '양호')
    item2_remark = request.form.get('item2_remark', '')
    item3_status = request.form.get('item3_status', '양호')
    item3_remark = request.form.get('item3_remark', '')
    item4_status = request.form.get('item4_status', '양호')
    item4_remark = request.form.get('item4_remark', '')
    
    temp = request.form.get('temp')
    humidity = request.form.get('humidity')
    
    now = (datetime.utcnow() + timedelta(hours=9)).strftime('%Y-%m-%d %H:%M:%S')
    
    new_row = pd.DataFrame([{
        "점검일시": now,
        "업체명": company,
        "점검자": inspector,
        "건축물구조_상태": item1_status,
        "건축물구조_비고": item1_remark,
        "소방환기_상태": item2_status,
        "소방환기_비고": item2_remark,
        "표지저장_상태": item3_status,
        "표지저장_비고": item3_remark,
        "온습도누출_상태": item4_status,
        "온습도누출_비고": item4_remark,
        "온도": temp,
        "습도": humidity
    }])
    
    df = pd.read_csv(DATA_FILE)
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
    
    return render_template('success.html')

@app.route('/admin')
def admin():
    init_db()
    df = pd.read_csv(DATA_FILE)
    records = df.to_dict(orient='records')
    return render_template('admin.html', records=records)

@app.route('/download')
def download():
    company = request.args.get('company', 'all')
    month = request.args.get('month', 'all')
    report_filename = generate_monthly_excel(company, month)
    return send_file(report_filename, as_attachment=True)

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
