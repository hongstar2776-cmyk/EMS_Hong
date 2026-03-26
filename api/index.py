import io
import requests
from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
import openpyxl
from openpyxl.styles import Font

app = Flask(__name__)
CORS(app)

TEMPLATE_URL = "https://hongstar2776-cmyk.github.io/My-Dashboard/resource/template_estmate.xlsx"

def write_row(ws, row_idx, data):
    ws[f'A{row_idx}'] = data.get('category', '')
    ws[f'B{row_idx}'] = data.get('name', '')
    ws[f'C{row_idx}'] = data.get('spec', '')
    ws[f'D{row_idx}'] = data.get('unit', '')

    if data.get('_type') == 'header':
        for col in ['A', 'B', 'C', 'D']:
            ws[f'{col}{row_idx}'].font = Font(bold=True)
        return

    qty = float(data.get('qty', 0) or 0)
    mat_up = float(data.get('mat_up', 0) or 0)
    lab_up = float(data.get('lab_up', 0) or 0)
    exp_up = float(data.get('exp_up', 0) or 0)

    ws[f'E{row_idx}'] = qty
    ws[f'F{row_idx}'] = mat_up
    ws[f'H{row_idx}'] = lab_up
    ws[f'J{row_idx}'] = exp_up
    
    ws[f'G{row_idx}'] = f"=E{row_idx}*F{row_idx}"
    ws[f'I{row_idx}'] = f"=E{row_idx}*H{row_idx}"
    ws[f'K{row_idx}'] = f"=E{row_idx}*J{row_idx}"
    ws[f'L{row_idx}'] = f"=F{row_idx}+H{row_idx}+J{row_idx}"
    ws[f'M{row_idx}'] = f"=G{row_idx}+I{row_idx}+K{row_idx}"
    ws[f'N{row_idx}'] = data.get('note', '')

def write_subtotal(ws, row_idx, start_row, end_row, category):
    ws[f'A{row_idx}'] = category
    ws[f'B{row_idx}'] = f"[{category} 소계]"
    
    ws[f'G{row_idx}'] = f"=SUM(G{start_row}:G{end_row})"
    ws[f'I{row_idx}'] = f"=SUM(I{start_row}:I{end_row})"
    ws[f'K{row_idx}'] = f"=SUM(K{start_row}:K{end_row})"
    ws[f'M{row_idx}'] = f"=SUM(M{start_row}:M{end_row})"

    for col in ['A', 'B', 'G', 'I', 'K', 'M']:
        ws[f'{col}{row_idx}'].font = Font(bold=True)

@app.route('/api/export', methods=['POST'])
def export_excel():
    try:
        payload = request.json
        tabs = payload.get('tabs', [])

        if not tabs:
            return jsonify({"error": "출력할 데이터가 없습니다."}), 400

        response = requests.get(TEMPLATE_URL)
        response.raise_for_status()
        
        wb = openpyxl.load_workbook(io.BytesIO(response.content))
        
        base_est_sheet = wb.worksheets[0]
        base_sum_sheet = None
        for sheet in wb.worksheets:
            if "합계표" in sheet.title or "공종별" in sheet.title:
                base_sum_sheet = sheet
                break
        
        for i, tab in enumerate(tabs):
            tab_name = tab.get('name', f'내역서 {i+1}')
            tab_data = tab.get('data', [])
            
            # 내역서 시트 복사 및 설정
            new_est_sheet = wb.copy_worksheet(base_est_sheet)
            new_est_sheet.title = tab_name
            # [추가] 반복할 행(1~4행) 인쇄 옵션 고정
            new_est_sheet.print_title_rows = '1:4'
            
            new_sum_sheet = None
            if base_sum_sheet:
                new_sum_sheet = wb.copy_worksheet(base_sum_sheet)
                idx_str = tab_name.split(' ')[-1] if ' ' in tab_name else str(i+1)
                new_sum_sheet.title = f"공종별합계표 {idx_str}"
                # [추가] 합계표 시트 인쇄 영역 고정
                new_sum_sheet.print_area = "B1:N27"

            groups = {}
            for r in tab_data:
                cat = r.get('category', '').strip() or '미지정'
                if cat not in groups:
                    groups[cat] = []
                groups[cat].append(r)

            current_row = 5 
            summary_data = [] 
            
            for cat, rows in groups.items():
                cat_ranges = [] 
                items_to_print = [{'_type': 'header', 'category': cat, 'name': cat}] + rows
                
                while items_to_print:
                    start_chunk_row = current_row
                    
                    if len(items_to_print) <= 20:
                        page_items = items_to_print[:]
                        items_to_print = []
                        
                        first_data = -1
                        last_data = -1
                        
                        for r in page_items:
                            write_row(new_est_sheet, current_row, r)
                            if r.get('_type') != 'header': 
                                if first_data == -1: first_data = current_row
                                last_data = current_row
                            current_row += 1
                            
                        if first_data != -1:
                            cat_ranges.append((first_data, last_data))
                        
                        while current_row < start_chunk_row + 20:
                            current_row += 1
                            
                        current_row += 1 
                        write_subtotal(new_est_sheet, current_row, start_chunk_row, start_chunk_row + 19, cat)
                        current_row += 2 
                        
                    elif len(items_to_print) == 21:
                        page_items = items_to_print[:21]
                        items_to_print = []
                        
                        first_data = -1
                        last_data = -1
                        
                        for r in page_items:
                            write_row(new_est_sheet, current_row, r)
                            if r.get('_type') != 'header':
                                if first_data == -1: first_data = current_row
                                last_data = current_row
                            current_row += 1
                        
                        if first_data != -1:
                            cat_ranges.append((first_data, last_data))
                        
                        write_subtotal(new_est_sheet, current_row, start_chunk_row, start_chunk_row + 20, cat)
                        current_row += 2 
                        
                    else: 
                        page_items = items_to_print[:21]
                        items_to_print = items_to_print[21:]
                        
                        first_data = -1
                        last_data = -1
                        
                        for r in page_items:
                            write_row(new_est_sheet, current_row, r)
                            if r.get('_type') != 'header':
                                if first_data == -1: first_data = current_row
                                last_data = current_row
                            current_row += 1
                            
                        if first_data != -1:
                            cat_ranges.append((first_data, last_data))
                        
                        new_est_sheet[f'B{current_row}'] = "[...다음 장으로 이어짐]"
                        current_row += 1
                        new_est_sheet[f'B{current_row}'] = "[...다음 장으로 이어짐]"
                        current_row += 1
                
                summary_data.append({
                    'category': cat,
                    'ranges': cat_ranges
                })
            
            # [추가] 새 내역서 시트의 인쇄 영역 동적 설정 (4 + 23 * 페이지수와 정확히 일치)
            last_row = current_row - 1
            new_est_sheet.print_area = f"B1:N{last_row}"
            
            if new_sum_sheet:
                sum_row = 5
                for s_data in summary_data:
                    cat = s_data['category']
                    ranges = s_data['ranges']
                    
                    new_sum_sheet[f'A{sum_row}'] = cat
                    new_sum_sheet[f'B{sum_row}'] = f"[{cat} 소계]"
                    
                    if ranges:
                        est_sheet_name = new_est_sheet.title
                        g_parts = [f"'{est_sheet_name}'!G{s}:G{e}" for s,e in ranges]
                        i_parts = [f"'{est_sheet_name}'!I{s}:I{e}" for s,e in ranges]
                        k_parts = [f"'{est_sheet_name}'!K{s}:K{e}" for s,e in ranges]
                        m_parts = [f"'{est_sheet_name}'!M{s}:M{e}" for s,e in ranges]
                        
                        new_sum_sheet[f'G{sum_row}'] = f"=SUM({','.join(g_parts)})"
                        new_sum_sheet[f'I{sum_row}'] = f"=SUM({','.join(i_parts)})"
                        new_sum_sheet[f'K{sum_row}'] = f"=SUM({','.join(k_parts)})"
                        new_sum_sheet[f'M{sum_row}'] = f"=SUM({','.join(m_parts)})"
                    
                    for col in ['A', 'B', 'G', 'I', 'K', 'M']:
                        new_sum_sheet[f'{col}{sum_row}'].font = Font(bold=True)
                        
                    sum_row += 1

        # [수정] 원본 템플릿 시트들은 보이지 않게 숨김 처리
        base_est_sheet.sheet_state = 'hidden'
        if base_sum_sheet:
            base_sum_sheet.sheet_state = 'hidden'

        # [추가] 숨김 처리 후, 엑셀을 열었을 때 포커스가 맞춰질 활성 시트(새로 만든 첫번째 시트)를 명시적으로 지정
        for idx, sheet_name in enumerate(wb.sheetnames):
            if wb[sheet_name].sheet_state == 'visible':
                wb.active = idx
                break

        output = io.BytesIO()
        wb.save(output)
        output.seek(0)

        return send_file(
            output, 
            as_attachment=True, 
            download_name="내역서_멀티탭_출력.xlsx",
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    except Exception as e:
        print("서버 에러 발생:", str(e))
        return jsonify({"error": str(e)}), 500
