from flask import Flask, request, jsonify
from flask_cors import CORS

app = Flask(__name__)
# CORS 설정: 깃허브 페이지(프론트엔드)에서 오는 요청을 허락해줍니다.
CORS(app) 

# '/api/export' 라는 주소로 프론트엔드가 데이터를(POST) 보내면 이 함수가 실행됩니다!
@app.route('/api/export', methods=['POST'])
def export_excel():
    try:
        # 1. 프론트엔드(캔버스의 시스템)에서 보낸 내역서 탭 데이터를 받습니다.
        received_data = request.json
        
        # 2. 버셀 서버 로그에 잘 들어왔는지 출력해봅니다.
        print("프론트엔드에서 데이터가 잘 도착했습니다!")
        
        # 3. 프론트엔드 쪽으로 "통신 성공!"이라는 메시지를 다시 던져줍니다.
        # (다음 단계에서는 이 부분에 엑셀 23행 분할 파이썬 코드를 넣을 겁니다!)
        return jsonify({"message": "Vercel API와 성공적으로 연결되었습니다!", "status": "success"}), 200

    except Exception as e:
        return jsonify({"error": str(e)}), 500
