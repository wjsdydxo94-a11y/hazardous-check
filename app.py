from flask import Flask, render_template, request, redirect, url_for, send_file
import pandas as pd
from datetime import datetime, timedelta
import os
import requests
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
DATA_FILE = 'inspection_db.csv'

POWER_AUTOMATE_WEBHOOK_URL = "YOUR_POWER_AUTOMATE_HTTP_URL_HERE"

def init_db():
    if not os.path.exists(DATA_FILE):
        df = pd.DataFrame(columns=[
            "점검일시", "업체명", "점검자", "건축물구조", "소방환기", 
            "표지저장", "온습도누출", "온도", "습도", "특이사항"
        ])
        df.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')

def generate_monthly_excel(company_name):
    init_db()
    df_db = pd.read_csv(DATA_FILE)
    
    if '업체명' in df_db.columns:
        df_db = df_db[df_db['업체명'] == company_name]
    
    # 가장 최근 제출된 데이터 가져오기
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

    # 1. Title
    ws.merge_cells("A1:E1")
    ws["A1"] = f"위험물 저장소 일일점검 항목별 상세 절차 및 기준표 ({company_name})"
    ws["A1"].font = font_title
    ws["A1"].alignment = align_center
    ws.row_dimensions[1].height = 40

    # 2. Sub-note
    ws.merge_cells("A2:E2")
    ws["A2"] = "* 본 점검표는 초보자도 쉽게 이해할 수 있도록 4대 점검 항목의 세부 절차와 정상 기준을 명시하고 즉시 점검 결과를 기록하는 법정 서식입니다."
    ws["A2"].font = Font(name="맑은 고딕", size=9, italic=True, color="595959")
    ws["A2"].alignment = Alignment(horizontal="left", vertical="center")
    ws.row_dimensions[2].height = 20

    # 3. Metadata Header Box (안전관리자란 공백 수기 작성용)
    current_date_str = datetime.now().strftime('%Y년 %m월 %d일')
    inspector_val = str(latest_row['점검자']) if latest_row is not None and '점검자' in latest_row else ""
    
    metadata = [
        ("사업장명", f"주식회사 {company_name}", "점검일자", current_date_str),
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

    # 4. Table Headers (사진 기준표 + 체크란 + 특이사항 란)
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

    # 5. Checklist Data Rows
    def get_status_mark(val):
        if latest_row is None:
            return "[   ]"
        val_str = str(val).strip()
        if any(kw in val_str for kw in ['양호', '정상', '이상없음', '적합']):
            return "양호 (O)"
        else:
            return "정비요함 (X)"

    items_data = [
        (
            "1. 건축물 구조 및 피뢰설비",
            "• 저장소 외벽, 지붕, 바닥, 출입문에 균열이나 변형이 있는지 육안으로 살핍니다.\n• 건물 최상단에 설치된 피뢰침(피뢰도체)과 건물 외벽을 따라 내려오는 접지선이 끊어지거나 부식되지 않았는지 확인합니다.",
            "• 건물 구조체에 누수나 심한 균열이 없습니다.\n• 피뢰침이 건물의 가장 높은 곳에 단단히 고정되어 있고, 접지선이 땅속까지 끊김 없이 안전하게 연결되어 있습니다.",
            get_status_mark(latest_row.get('건축물구조', '')) if latest_row is not None else "[   ]"
        ),
        (
            "2. 소방 및 환기·배출설비",
            "• 비치된 소화기의 압력계 바늘 위치를 확인합니다.\n• 환기팬(배기팬) 스위치를 켜서 정상적으로 회전하는지 확인합니다.",
            "• 소화기 압력계 바늘이 녹색(정상) 영역에 정확히 위치해 있습니다.\n• 환기팬 가동 시 이상 소음 없이 인화성 증기를 밖으로 원활하게 배출합니다.",
            get_status_mark(latest_row.get('소방환기', '')) if latest_row is not None else "[   ]"
        ),
        (
            "3. 표지판 및 위험물 저장·취급",
            "• 출입구에 부착된 '위험물 옥내저장소', '화기엄금', '금연' 표지판이 잘 보이는지 확인합니다.\n• 저장소 내부 바닥과 통행로를 둘러봅니다.",
            "• 표지판이 훼손되거나 글씨가 지워지지 않고 선명하게 부착되어 있습니다.\n• 저장소 내부에 위험물 용기 외에 종이박스, 쓰레기 등 가연성 폐기물이 일절 없습니다.",
            get_status_mark(latest_row.get('표지저장', '')) if latest_row is not None else "[   ]"
        ),
        (
            "4. 온습도 및 누출·비산 방지",
            "• 저장소 내부에 부착된 온·습도계 수치를 확인합니다.\n• 드럼 및 용기 하단부, 바닥 턱(방유제) 주변을 확인합니다.",
            "• 원료 변질을 유발하는 극단적인 고온·다습을 피해 적정 범위 내로 유지됩니다.\n• 바닥이나 용기 하단에 액체가 흘러내린 흔적(누유, 누수)이나 미세 누출이 전혀 없습니다.",
            get_status_mark(latest_row.get('온습도누출', '')) if latest_row is not None else "[   ]"
        )
    ]

    latest_remark = str(latest_row.get('특이사항', '-')) if latest_row is not None and '특이사항' in latest_row else ""
    latest_temp = str(latest_row.get('온도', '')) if latest_row is not None else ""
    latest_humid = str(latest_row.get('습도', '')) if latest_row is not None else ""
    
    remark_text = latest_remark
    if latest_temp or latest_humid:
        remark_text = f"[측정 온습도] 온도: {latest_temp}℃, 습도: {latest_humid}%\n[특이사항] {latest_remark}"

    for idx, item in enumerate(items_data, 7):
        ws.row_dimensions[idx].height = 70
        
        c1 = ws.cell(row=idx, column=1, value=item[0])
        c2 = ws.cell(row=idx, column=2, value=item[1])
        c3 = ws.cell(row=idx, column=3, value=item[2])
        c4 = ws.cell(row=idx, column=4, value=item[3])
        
        # 첫 번째 행에만 특이사항 입력, 나머지는 빈칸 또는 병합
        c5 = ws.cell(row=idx, column=5, value=remark_text if idx == 7 else "")

        for cell in [c1, c2, c3, c4, c5]:
            cell.font = font_body
            cell.border = thin_border
            cell.alignment = align_center if cell != c2 and cell != c3 and cell != c5 else align_left
            if idx % 2 == 1:
                cell.fill = PatternFill(start_color="FAFAFA", end_color="FAFAFA", fill_type="solid")

    # 온도/습도 기록용 추가 행
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

    report_filename = f'위험물일일점검표_상세지침형_{company_name}.xlsx'
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
    item1 = request.form.get('item1', '양호')
    item2 = request.form.get('item2', '양호')
    item3 = request.form.get('item3', '양호')
    item4 = request.form.get('item4', '양호')
    temp = request.form.get('temp')
    humidity = request.form.get('humidity')
    remarks = request.form.get('remarks', '-')
    
    now = (datetime.utcnow() + timedelta(hours=9)).strftime('%Y-%m-%d %H:%M:%S')
    
    new_row = pd.DataFrame([{
        "점검일시": now,
        "업체명": company,
        "점검자": inspector,
        "건축물구조": item1,
        "소방환기": item2,
        "표지저장": item3,
        "온습도누출": item4,
        "온도": temp,
        "습도": humidity,
        "특이사항": remarks
    }])
    
    df = pd.read_csv(DATA_FILE)
    df = pd.concat([df, new_row], ignore_index=True)
    df.to_csv(DATA_FILE, index=False, encoding='utf-8-sig')
    
    if POWER_AUTOMATE_WEBHOOK_URL != "YOUR_POWER_AUTOMATE_HTTP_URL_HERE":
        payload = {
            "점검일시": now, "업체명": company, "점검자": inspector,
            "건축물구조": item1, "소방환기": item2, "표지저장": item3,
            "온습도누출": item4, "온도": str(temp), "습도": str(humidity), "특이사항": remarks
        }
        try:
            requests.post(POWER_AUTOMATE_WEBHOOK_URL, json=payload)
        except Exception as e:
            print("OneDrive sync failed:", e)
    
    return render_template('success.html')

@app.route('/admin')
def admin():
    init_db()
    df = pd.read_csv(DATA_FILE)
    records = df.to_dict(orient='records')
    return render_template('admin.html', records=records)

@app.route('/download')
def download():
    company = request.args.get('company', '아로마솔루션')
    report_filename = generate_monthly_excel(company)
    return send_file(report_filename, as_attachment=True)

if __name__ == '__main__':
    init_db()
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port)
