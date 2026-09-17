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
    
    if company_name and company_name != 'all' and '업체명' in df_db.columns:
        df_db = df_db[df_db['업체명'] == company_name]
    
    if target_month and '점검일시' in df_db.columns:
        df_db = df_db[df_db['점검일시'].astype(str).str.startswith(target_month)]
    
    latest_row = None
    if not df_db.empty:
        latest_row = df_db.iloc[-1]
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "위험물일일점검표"
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

    display_company = company_name if company_name and company_name != 'all' else "통합"
    display_month = target_month if target_month else datetime.now().strftime('%Y년 %m월')

    # 1. Title
    ws.merge_cells("A1:E1")
    ws["A1"] = f"위험물 저장소 일일점검 항목별 상세 절차 및 기준표 ({display_company} / {display_month})"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center
    ws.row_dimensions[1].height = 40

    # 2. Sub-note
    ws.merge_cells("A2:E2")
    ws["A2"] = "* 본 점검표는 초보자도 쉽게 이해할 수 있도록 4대 점검 항목의 세부 절차와 정상 기준을 명시하고 즉시 점검 결과를 기록하는 법정 서식입니다."
    ws["A2"].font = Font(name="맑은 고딕", size=9, italic=True, color="595959")
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 20

    # 3. Metadata Header Box
    metadata = [
        ("사업장명", f"주식회사 {display_company}", "점검년월", display_month),
        ("점검대상", "옥내저장소", "안전관리자", "")
    ]

    for r_idx, meta in enumerate(metadata, 3):
        ws.row_dimensions[r_idx].height = 24
        ws.merge_cells(start_row=r_idx, start_column=1, end_row=r_idx, end_column=1)
        ws.cell(row=r_idx, column=1, value=meta[0]).fill = fill_meta
        ws.cell(row=r_idx, column=1).font = font_body
        ws.cell(row=r_idx, column=1).alignment = align_center
        ws.cell(row=r_idx, column=1).border = thin_border

        ws.merge_cells(start_row=r_idx, start_column=2, end_row=r_idx, end_column=2)
        ws.cell(row=r_idx, column=2, value=meta[1]).alignment = align_center
        ws.cell(row=r_idx, column=2).font = font_body
        ws.cell(row=r_idx, column=2).border = thin_border

        ws.cell(row=r_idx, column=3, value="").fill = fill_meta
        ws.cell(row=r_idx, column=3).border = thin_border

        ws.cell(row=r_idx, column=4, value=meta[2]).fill = fill_meta
        ws.cell(row=r_idx, column=4).font = font_body
        ws.cell(row=r_idx, column=4).alignment = align_center
        ws.cell(row=r_idx, column=4).border = thin_border

        ws.merge_cells(start_row=r_idx, start_column=5, end_row=r_idx, end_column=5)
        ws.cell(row=r_idx, column=5, value=meta[3]).alignment = align_center
        ws.cell(row=r_idx, column=5).font = font_body
        ws.cell(row=r_idx, column=5).border = thin_border

    ws.row_dimensions[5].height = 10

    # 4. Table Headers
    headers = [
        "점검 항목", 
        "구체적인 점검 방법 (How to check)", 
        "정상(양호) 판정 기준 (Normal Criteria)", 
        "점검 결과\n(양호/정비요함)", 
        "특이사항 및 조치내용"
    ]

    ws.row_dimensions[6].height = 30
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=6, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = align_center
        cell.border = thin_border

    def get_status_mark(val):
        if latest_row is None or pd.isna(val):
            return "[   ]"
        val_str = str(val).strip()
        if any(kw in val_str for kw in ['양호', '정상', '이상없음', '적합', 'O', 'o']):
            return "양호 (O)"
        elif any(kw in val_str for kw in ['정비요함', '불량', 'X', 'x', '조치필요']):
            return "정비요함 (X)"
        return "[   ]"

    items_data = [
        (
            "1. 건축물 구조 및 피뢰설비",
            "• 저장소 외벽, 지붕, 바닥, 출입문에 균열이나 변형이 있는지 육안으로 살핍니다.\n• 건물 최상단에 설치된 피뢰침(피뢰도체)과 건물 외벽을 따라 내려오는 접지선이 끊어지거나 부식되지 않았는지 확인합니다.",
            "• 건물 구조체에 누수나 심한 균열이 없습니다.\n• 피뢰침이 건물의 가장 높은 곳에 단단히 고정되어 있고, 접지선이 땅속까지 끊김 없이 안전하게 연결되어 있습니다.",
            get_status_mark(latest_row.get('건축물구조_상태', '')) if latest_row is not None else "[   ]",
            str(latest_row.get('건축물구조_비고', '')) if latest_row is not None and pd.notna(latest_row.get('건축물구조_비고')) else ""
        ),
        (
            "2. 소방 및 환기·배출설비",
            "• 비치된 소화기의 압력계 바늘 위치를 확인합니다.\n• 환기팬(배기팬) 스위치를 켜서 정상적으로 회전하는지 확인합니다.",
            "• 소화기 압력계 바늘이 녹색(정상) 영역에 정확히 위치해 있습니다.\n• 환기팬 가동 시 이상 소음 없이 인화성 증기를 밖으로 원활하게 배출합니다.",
            get_status_mark(latest_row.get('소방환기_상태', '')) if latest_row is not None else "[   ]",
            str(latest_row.get('소방환기_비고', '')) if latest_row is not None and pd.notna(latest_row.get('소방환기_비고')) else ""
        ),
        (
            "3. 표지판 및 위험물 저장·취급",
            "• 출입구에 부착된 '위험물 옥내저장소', '화기엄금', '금연' 표지판이 잘 보이는지 확인합니다.\n• 저장소 내부 바닥과 통행로를 둘러봅니다.",
            "• 표지판이 훼손되거나 글씨가 지워지지 않고 선명하게 부착되어 있습니다.\n• 저장소 내부에 위험물 용기 외에 종이박스, 쓰레기 등 가연성 폐기물이 일절 없습니다.",
            get_status_mark(latest_row.get('표지저장_상태', '')) if latest_row is not None else "[   ]",
            str(latest_row.get('표지저장_비고', '')) if latest_row is not None and pd.notna(latest_row.get('표지저장_비고')) else ""
        ),
        (
            "4. 온습도 및 누출·비산 방지",
            "• 저장소 내부에 부착된 온·습도계 수치를 확인합니다.\n• 드럼 및 용기 하단부, 바닥 턱(방유제) 주변을 확인합니다.",
            "• 원료 변질을 유발하는 극단적인 고온·다습을 피해 적정 범위 내로 유지됩니다.\n• 바닥이나 용기 하단에 액체가 흘러내린 흔적(누유, 누수)이나 미세 누출이 전혀 없습니다.",
            get_status_mark(latest_row.get('온습도누출_상태', '')) if latest_row is not None else "[   ]",
            str(latest_row.get('온습도누출_비고', '')) if latest_row is not None and pd.notna(latest_row.get('온습도누출_비고')) else ""
        )
    ]

    latest_temp = ""
    latest_humid = ""
    if latest_row is not None:
        t_val = latest_row.get('온도', '')
        h_val = latest_row.get('습도', '')
        if pd.notna(t_val) and str(t_val).strip() != '':
            latest_temp = str(t_val)
        if pd.notna(h_val) and str(h_val).strip() != '':
            latest_humid = str(h_val)

    for idx, item in enumerate(items_data, 7):
        ws.row_dimensions[idx].height = 70
        
        c1 = ws.cell(row=idx, column=1, value=item[0])
        c2 = ws.cell(row=idx, column=2, value=item[1])
        c3 = ws.cell(row=idx, column=3, value=item[2])
        c4 = ws.cell(row=idx, column=4, value=item[3])
        c5 = ws.cell(row=idx, column=5, value=item[4])

        for cell in [c1, c2, c3, c4, c5]:
            cell.font = font_body
            cell.border = thin_border
            cell.alignment = align_center if cell != c2 and cell != c3 and cell != c5 else align_left
            if idx % 2 == 1:
                cell.fill = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

    r_env = 11
    ws.row_dimensions[r_env].height = 30
    ws.merge_cells(start_row=r_env, start_column=1, end_row=r_env, end_column=3)
    ws.cell(row=r_env, column=1, value="저장소 환경 측정값 (온도 / 습도)").font = Font(name="맑은 고딕", size=10, bold=True)
    ws.cell(row=r_env, column=1).alignment = align_center
    ws.cell(row=r_env, column=1).fill = fill_meta
    
    for c in range(1, 4):
        ws.cell(row=r_env, column=c).border = thin_border

    ws.cell(row=r_env, column=4, value=f"온도: {latest_temp} ℃" if latest_temp else "온도: [   ] ℃").alignment = align_center
    ws.cell(row=r_env, column=4).font = font_body
    ws.cell(row=r_env, column=4).border = thin_border

    ws.cell(row=r_env, column=5, value=f"습도: {latest_humid} %" if latest_humid else "습도: [   ] %").alignment = align_center
    ws.cell(row=r_env, column=5).font = font_body
    ws.cell(row=r_env, column=5).border = thin_border

    col_widths = [26, 45, 45, 18, 30]
    for idx, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    report_filename = f'위험물일일점검표_{display_company}_{display_month.replace(" ", "_")}.xlsx'
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
    month = request.args.get('month', '')
    report_filename = generate_monthly_excel(company, month)
    return send_file(report_filename, as_attachment=True)

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
