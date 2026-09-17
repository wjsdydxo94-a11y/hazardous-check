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
        try:
            df = pd.read_csv(DATA_FILE)
            if not all(col in df.columns for col in ["건축물구조_상태", "온도", "습도"]):
                df_new = pd.DataFrame(columns=required_cols)
                df_new.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
        except Exception:
            df = pd.DataFrame(columns=required_cols)
            df.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')

# 월별 기록 관리 대장 생성
def generate_monthly_excel(company_name, target_month=None):
    init_db()
    df_db = pd.read_csv(DATA_FILE)
    
    if company_name and company_name != 'all' and '업체명' in df_db.columns:
        df_db = df_db[df_db['업체명'] == company_name]
    
    if target_month and target_month != 'all' and '점검일시' in df_db.columns:
        df_db = df_db[df_db['점검일시'].astype(str).str.startswith(target_month)]
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "월별기록관리대장"
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

    ws.merge_cells("A1:J1")
    ws["A1"] = f"위험물 저장소/제조소 월별 기록 관리 대장 ({display_company} / {display_month})"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center
    ws.row_dimensions[1].height = 40
    ws.row_dimensions[2].height = 10

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

    report_filename = f'위험물저장소_제조소_월별기록관리대장_{display_company}_{display_month.replace(" ", "_")}.xlsx'
    wb.save(report_filename)
    return report_filename

# 개별 상세 주간점검표 생성 (기준표 포함)
def generate_single_excel_by_index(idx):
    init_db()
    df_db = pd.read_csv(DATA_FILE)
    
    if idx < 0 or idx >= len(df_db):
        return None
    
    row = df_db.iloc[idx]
    company_name = str(row.get('업체명', '아로마솔루션'))
    dt_str = str(row.get('점검일시', datetime.now().strftime('%Y-%m-%d %H:%M:%S')))
    
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "위험물주간점검표"
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

    ws.merge_cells("A1:E1")
    ws["A1"] = f"위험물 저장소/제조소 주간점검 항목별 상세 절차 및 기준표 ({company_name})"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center
    ws.row_dimensions[1].height = 40

    ws.merge_cells("A2:E2")
    ws["A2"] = "* 본 점검표는 초보자도 쉽게 이해할 수 있도록 4대 점검 항목의 세부 절차와 정상 기준을 명시하고 즉시 점검 결과를 기록하는 법정 주간 서식입니다."
    ws["A2"].font = Font(name="맑은 고딕", size=9, italic=True, color="595959")
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 20

    metadata = [
        ("사업장명", f"주식회사 {company_name}", "점검일시", dt_str),
        ("점검대상", "옥내저장소/제조소", "점검자", str(row.get('점검자', '')))
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

    def format_cell_status(status_val, remark_val):
        s = str(status_val).strip()
        r = str(remark_val).strip() if pd.notna(remark_val) and str(remark_val) != 'nan' else ''
        res = "양호 (O)" if ('양호' in s or '정상' in s or 'O' in s) else "정비요함 (X)"
        if r:
            res += f"\n({r})"
        return res

    items_data = [
        (
            "1. 건축물 구조 및 피뢰설비",
            "• 저장소 외벽, 지붕, 바닥, 출입문에 균열이나 변형이 있는지 육안으로 살핍니다.\n• 건물 최상단에 설치된 피뢰침(피뢰도체)과 건물 외벽을 따라 내려오는 접지선이 끊어지거나 부식되지 않았는지 확인합니다.",
            "• 건물 구조체에 누수나 심한 균열이 없습니다.\n• 피뢰침이 건물의 가장 높은 곳에 단단히 고정되어 있고, 접지선이 땅속까지 끊김 없이 안전하게 연결되어 있습니다.",
            format_cell_status(row.get('건축물구조_상태', ''), row.get('건축물구조_비고', ''))
        ),
        (
            "2. 소방 및 환기·배출설비",
            "• 비치된 소화기의 압력계 바늘 위치를 확인합니다.\n• 환기팬(배기팬) 스위치를 켜서 정상적으로 회전하는지 확인합니다.",
            "• 소화기 압력계 바늘이 녹색(정상) 영역에 정확히 위치해 있습니다.\n• 환기팬 가동 시 이상 소음 없이 인화성 증기를 밖으로 원활하게 배출합니다.",
            format_cell_status(row.get('소방환기_상태', ''), row.get('소방환기_비고', ''))
        ),
        (
            "3. 표지판 및 위험물 저장·취급",
            "• 출입구에 부착된 '위험물 옥내저장소', '화기엄금', '금연' 표지판이 잘 보이는지 확인합니다.\n• 저장소 내부 바닥과 통행로를 둘러봅니다.",
            "• 표지판이 훼손되거나 글씨가 지워지지 않고 선명하게 부착되어 있습니다.\n• 저장소 내부에 위험물 용기 외에 종이박스, 쓰레기 등 가연성 폐기물이 일절 없습니다.",
            format_cell_status(row.get('표지저장_상태', ''), row.get('표지저장_비고', ''))
        ),
        (
            "4. 온습도 및 누출·비산 방지",
            "• 저장소 내부에 부착된 온·습도계 수치를 확인합니다.\n• 드럼 및 용기 하단부, 바닥 턱(방유제) 주변을 확인합니다.",
            "• 원료 변질을 유발하는 극단적인 고온·다습을 피해 적정 범위 내로 유지됩니다.\n• 바닥이나 용기 하단에 액체가 흘러내린 흔적(누유, 누수)이나 미세 누출이 전혀 없습니다.",
            format_cell_status(row.get('온습도누출_상태', ''), row.get('온습도누출_비고', ''))
        )
    ]

    for idx, item in enumerate(items_data, 7):
        ws.row_dimensions[idx].height = 70
        
        c1 = ws.cell(row=idx, column=1, value=item[0])
        c2 = ws.cell(row=idx, column=2, value=item[1])
        c3 = ws.cell(row=idx, column=3, value=item[2])
        c4 = ws.cell(row=idx, column=4, value=item[3])
        c5 = ws.cell(row=idx, column=5, value="")

        for cell in [c1, c2, c3, c4, c5]:
            cell.font = font_body
            cell.border = thin_border
            cell.alignment = align_center if cell != c2 and cell != c3 else align_left
            if idx % 2 == 1:
                cell.fill = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

    latest_temp = str(row.get('온도', ''))
    latest_humid = str(row.get('습도', ''))

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

    col_widths = [26, 45, 45, 22, 25]
    for idx, width in enumerate(col_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = width

    report_filename = f'위험물주간점검표_{company_name}_{dt_str[:10]}.xlsx'
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

@app.route('/admin_add', methods=['POST'])
def admin_add():
    init_db()
    try:
        company = request.form.get('company', '아로마솔루션')
        inspector = request.form.get('inspector', '관리자')
        
        custom_date = request.form.get('inspection_date')
        if custom_date:
            try:
                dt_obj = datetime.strptime(custom_date, '%Y-%m-%dT%H:%M')
                now = dt_obj.strftime('%Y-%m-%d %H:%M:%S')
            except:
                now = (datetime.utcnow() + timedelta(hours=9)).strftime('%Y-%m-%d %H:%M:%S')
        else:
            now = (datetime.utcnow() + timedelta(hours=9)).strftime('%Y-%m-%d %H:%M:%S')
        
        item1_status = request.form.get('item1_status', '양호')
        item1_remark = request.form.get('item1_remark', '')
        item2_status = request.form.get('item2_status', '양호')
        item2_remark = request.form.get('item2_remark', '')
        item3_status = request.form.get('item3_status', '양호')
        item3_remark = request.form.get('item3_remark', '')
        item4_status = request.form.get('item4_status', '양호')
        item4_remark = request.form.get('item4_remark', '')
        
        temp = request.form.get('temp', '20')
        humidity = request.form.get('humidity', '50')
        
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
    except Exception as e:
        print("Add error:", e)
        import traceback
        traceback.print_exc()
        
    return redirect(url_for('admin'))

@app.route('/admin_edit', methods=['POST'])
def admin_edit():
    init_db()
    try:
        idx = int(request.form.get('index', -1))
        df = pd.read_csv(DATA_FILE)
        
        if 0 <= idx < len(df):
            custom_date = request.form.get('inspection_date')
            if custom_date:
                try:
                    dt_obj = datetime.strptime(custom_date, '%Y-%m-%dT%H:%M')
                    df.at[idx, '점검일시'] = dt_obj.strftime('%Y-%m-%d %H:%M:%S')
                except Exception as e:
                    print("Date parse error:", e)

            df.at[idx, '업체명'] = request.form.get('company', df.at[idx, '업체명'])
            df.at[idx, '점검자'] = request.form.get('inspector', df.at[idx, '점검자'])
            df.at[idx, '건축물구조_상태'] = request.form.get('item1_status', '양호')
            df.at[idx, '건축물구조_비고'] = request.form.get('item1_remark', '')
            df.at[idx, '소방환기_상태'] = request.form.get('item2_status', '양호')
            df.at[idx, '소방환기_비고'] = request.form.get('item2_remark', '')
            df.at[idx, '표지저장_상태'] = request.form.get('item3_status', '양호')
            df.at[idx, '표지저장_비고'] = request.form.get('item3_remark', '')
            df.at[idx, '온습도누출_상태'] = request.form.get('item4_status', '양호')
            df.at[idx, '온습도누출_비고'] = request.form.get('item4_remark', '')
            df.at[idx, '온도'] = request.form.get('temp', '')
            df.at[idx, '습도'] = request.form.get('humidity', '')
            df.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
    except Exception as e:
        print("Edit error:", e)
        import traceback
        traceback.print_exc()
        
    return redirect(url_for('admin'))

@app.route('/admin')
def admin():
    init_db()
    df = pd.read_csv(DATA_FILE)
    records = []
    for idx, row in df.iterrows():
        r_dict = row.to_dict()
        r_dict['row_index'] = idx
        records.append(r_dict)
    return render_template('admin.html', records=records)

@app.route('/download')
def download():
    mode = request.args.get('mode', 'monthly')
    if mode == 'single':
        idx = int(request.args.get('index', 0))
        report_filename = generate_single_excel_by_index(idx)
    else:
        company = request.args.get('company', 'all')
        month = request.args.get('month', 'all')
        report_filename = generate_monthly_excel(company, month)
    return send_file(report_filename, as_attachment=True)

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
