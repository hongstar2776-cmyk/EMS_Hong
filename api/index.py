from flask import Flask, request, send_file
from flask_cors import CORS
import openpyxl
import io

app = Flask(__name__)
# GitHub Pages에서 이 API를 호출할 수 있도록 허용(CORS)
CORS(app) 

@app.route('/api/generate-excel', methods=['POST'])
def generate_excel():
    # 1. 프론트엔드에서 보낸 JSON 데이터(주문서) 받기
    data = request.json
    items = data.get('items', []) # 예: [{name: "철근", qty: 10, price: 50000}, ...]

    # 2. 새로운 엑셀 워크북 만들기 (나중에는 만들어둔 템플릿을 불러올 수도 있습니다)
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "내역서"

    # 3. 엑셀 첫 줄(헤더) 쓰기
    ws.append(["품명", "수량", "단가", "금액"])

    # 4. 프론트에서 넘어온 데이터 한 줄씩 엑셀에 적기
    for item in items:
        # 금액은 수량 * 단가
        total_price = item['qty'] * item['price']
        ws.append([item['name'], item['qty'], item['price'], total_price])

    # 5. 완성된 엑셀을 파일 형태로 변환해서 프론트로 보내기 (컴퓨터에 저장하지 않고 메모리에서 바로 전송)
    out = io.BytesIO()
    wb.save(out)
    out.seek(0)
    
    return send_file(out, download_name="산출내역서.xlsx", as_attachment=True)