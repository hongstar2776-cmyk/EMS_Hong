import io
import requests
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Font

app = Flask(__name__)
# 프론트엔드(GitHub Pages)에서 보내는 요청을 허용합니다.
CORS(app)

# 깃허브에 올려두신 엑셀 템플릿 주소입니다.
TEMPLATE_URL = "https://hongstar2776-cmyk.github.io/My-Dashboard/resource/template_estmate.xlsx"

# 1. 엑셀의 각 행에 데이터를 입력하고 '수식'을 넣는 함수
def write_row(ws, row_idx, data):
    ws[f'A{row_idx}'] = data.get('category', '')
    ws[f'B{row_idx}'] = data.get('name', '')
    ws[f'C{row_idx}'] = data.get('spec', '')
    ws[f'D{row_idx}'] = data.get('unit', '')

    # 만약 파란띠(헤더) 행이라면 글씨를 굵게 하고 종료
    if data.get('_type') == 'header':
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{row_idx}'].font = Font(bold=True)
        return

    # 안전하게 숫자로 변환 (빈칸이면 0)
    qty = float(data.get('qty', 0) or 0)
    mat_up = float(data.get('mat_up', 0) or 0)
    lab_up = float(data.get('lab_up', 0) or 0)
    exp_up = float(data.get('exp_up', 0) or 0)

    # 단가 및 수량 입력
    ws[f'E{row_idx}'] = qty
    ws[f'F{row_idx}'] = mat_up
    ws[f'H{row_idx}'] = lab_up
    ws[f'J{row_idx}'] = exp_up
    
    # ★ 엑셀 수식(Formula) 적용 부분
    ws[f'G{row_idx}'] = f"=E{row_idx}*F{row_idx}"         # 자재금액
    ws[f'I{row_idx}'] = f"=E{row_idx}*H{row_idx}"         # 노무금액
    ws[f'K{row_idx}'] = f"=E{row_idx}*J{row_idx}"         # 경비금액
    ws[f'L{row_idx}'] = f"=F{row_idx}+H{row_idx}+J{row_idx}" # 합계단가
    ws[f'M{row_idx}'] = f"=G{row_idx}+I{row_idx}+K{row_idx}" # 합계금액
    ws[f'N{row_idx}'] = data.get('note', '')

# 2. 엑셀 소계(SUM 함수)를 입력하는 함수
def write_subtotal(ws, row_idx, start_row, end_row, category):
    ws[f'A{row_idx}'] = category
    ws[f'B{row_idx}'] = f"[{category} 소계]"
    
    # 세로 합계(SUM) 수식 적용
    ws[f'G{row_idx}'] = f"=SUM(G{start_row}:G{end_row})"
    ws[f'I{row_idx}'] = f"=SUM(I{start_row}:I{end_row})"
    ws[f'K{row_idx}'] = f"=SUM(K{start_row}:K{end_row})"
    ws[f'M{row_idx}'] = f"=SUM(M{start_row}:M{end_row})"

    # 소계행 폰트를 굵게 처리
    for col in ['A', 'B', 'G', 'I', 'K', 'M']:
        ws[f'{col}{row_idx}'].font = Font(bold=True)

# 3. 메인 통신(API) 라우터
@app.route('/api/export', methods=['POST'])
def export_excel():
    try:
        # 프론트엔드에서 보낸 탭 데이터(JSON) 덩어리를 받음
        payload = request.json
        tabs = payload.get('tabs', [])

        if not tabs:
            return jsonify({"error": "출력할 데이터가 없습니다."}), 400

        # 1) 깃허브에서 템플릿 다운로드
        response = requests.get(TEMPLATE_URL)
        response.raise_for_status()
        
        # 2) 엑셀 메모리에 올리기
        wb = openpyxl.load_workbook(io.BytesIO(response.content))
        base_sheet = wb.worksheets[0] # 첫 번째 원본 시트 (내역서)
        
        # 3) 탭(세트) 개수만큼 반복하며 시트 복사 및 데이터 입력
        for tab in tabs:
            tab_name = tab.get('name', '새 내역서')
            tab_data = tab.get('data', [])
            
            # ★ 핵심 1: 원본 시트 복사하고, 복사한 시트의 이름을 '내역서 1' 등으로 변경
            new_sheet = wb.copy_worksheet(base_sheet)
            new_sheet.title = tab_name

            # 공종번호 기준으로 그룹화
            groups = {}
            for r in tab_data:
                cat = r.get('category', '').strip() or '미지정'
                if cat not in groups:
                    groups[cat] = []
                groups[cat].append(r)

            current_row = 5 # 엑셀 A5 셀부터 시작
            
            # ★ 핵심 2: 23행 분할 알고리즘
            for cat, rows in groups.items():
                # 맨 앞에 파란띠(헤더) 추가
                items_to_print = [{'_type': 'header', 'category': cat, 'name': cat}] + rows
                
                while items_to_print:
                    start_data_row = current_row
                    
                    if len(items_to_print) <= 20:
                        page_items = items_to_print[:]
                        items_to_print = []
                        for r in page_items:
                            write_row(new_sheet, current_row, r)
                            current_row += 1
                            
                        # 20행 꽉 채울때까지 건너뜀
                        while current_row < start_data_row + 20:
                            current_row += 1
                        
                        current_row += 1 # 21행(빈행) 통과
                        write_subtotal(new_sheet, current_row, start_data_row, start_data_row + 19, cat) # 22행(소계)
                        current_row += 2 # 23행(빈행) 통과 후 다음 시작행으로

                    elif len(items_to_print) == 21:
                        page_items = items_to_print[:21]
                        items_to_print = items_to_print[21:]
                        for r in page_items:
                            write_row(new_sheet, current_row, r)
                            current_row += 1
                        
                        write_subtotal(new_sheet, current_row, start_data_row, start_data_row + 20, cat)
                        current_row += 2

                    else: # 22행 이상 남았을 때 (페이지 이월)
                        page_items = items_to_print[:21]
                        items_to_print = items_to_print[21:]
                        for r in page_items:
                            write_row(new_sheet, current_row, r)
                            current_row += 1
                        
                        # 22, 23행에 이월 표시
                        new_sheet[f'B{current_row}'] = "[...다음 장으로 이어짐]"
                        current_row += 1
                        new_sheet[f'B{current_row}'] = "[...다음 장으로 이어짐]"
                        current_row += 1

        # ★ 핵심 3: 기존 템플릿의 원본 "내역서" 시트는 보이지 않게 숨기기!
        base_sheet.sheet_state = 'hidden'

        # 완성된 엑셀 파일을 바이너리로 변환
        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        # 엑셀 파일(File) 형태로 프론트엔드에 응답 전송!
        return send_file(
            output, 
            as_attachment=True, 
            download_name="내역서_멀티탭_출력.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        print("서버 에러 발생:", str(e))
        return jsonify({"error": str(e)}), 500
