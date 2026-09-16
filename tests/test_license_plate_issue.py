import requests
import json
import base64
from pathlib import Path

def main():
    if not Path('supplement_data/ai_config.json').exists():
        return
    cfg = json.load(open('supplement_data/ai_config.json', encoding='utf-8'))
    key = cfg.get('gemini_api_key', '')
    img_path = Path('supplement_data/supplement_staging/761b72634556.jpg')
    if not img_path.exists():
        return
    img_bytes = img_path.read_bytes()
    b64 = base64.b64encode(img_bytes).decode('utf-8')

    target_date = "2026-07-08"
    target_time = "07:59:50"

    prompt = f"""Bạn là chuyên gia xử lý ảnh chuyên về watermark ngày giờ trên ảnh chấm công (Timestamp Camera, GPS Map Camera, Timekeeper).
Nhiệm vụ: Tìm chính xác vị trí watermark ngày giờ camera chấm công trên ảnh và tạo text thay thế theo ngày {target_date} và giờ {target_time}.

CỰC KỲ QUAN TRỌNG - QUY TẮC PHÂN BIỆT:
1. TUYỆT ĐỐI KHÔNG nhận diện biển số xe máy, biển số xe cộ, số hiệu xe, tem xe, bảng hiệu hay chữ trên đồng phục/thẻ tên nhân viên (ví dụ: các số như 59-C2, 107.59, 59-C1, 86-C1, 92-G1... LÀ BIỂN SỐ XE, CẤM SỬA!).
2. Watermark ngày giờ camera chấm công là chữ do ứng dụng chụp ảnh đóng dấu lên ảnh:
   - Thường nằm ở rìa/mép ảnh: góc trên cùng, góc dưới cùng, góc trái hoặc góc phải của bức ảnh.
   - Thường có bóng đổ (drop shadow) hoặc viền tương phản (đen/trắng).
   - Nội dung có cả ngày tháng và giờ phút giây (ví dụ: '18 Jan 2026 at 07.59.50' hoặc '21 Th1, 2026 05:52:34' hoặc '08/07/2026 07:59:50').

YÊU CẦU ĐẦU RA (JSON duy nhất):
{{
  "box_2d": [ymin, xmin, ymax, xmax],
  "original_text": "...",
  "new_text": "...",
  "text_color": "#FFFFFF",
  "has_shadow": true,
  "shadow_color": "#000000",
  "has_stroke": false,
  "stroke_color": "#000000",
  "font_style": "sans-serif",
  "font_weight": "bold"
}}
Trong đó:
- "new_text": Chuỗi thay thế phải giữ đúng 100% phong cách ngôn ngữ và định dạng của "original_text".
  + Nếu original_text dạng Tiếng Anh '18 Jan 2026 at 07.59.50' -> new_text BẮT BUỘC là '08 Jul 2026 at {target_time.replace(':', '.')}' hoặc '08 Jul 2026 at {target_time}'.
  + Nếu original_text dạng Tiếng Việt '21 Th1, 2026 05:52:34' -> new_text BẮT BUỘC là '08 Th7, 2026 {target_time}'.
  + Nếu original_text dạng 'DD/MM/YYYY HH:MM:SS' -> new_text BẮT BUỘC là '08/07/2026 {target_time}'.
"""

    for model in ['gemini-3.5-flash', 'gemini-3.6-flash', 'gemini-flash-latest']:
        try:
            r = requests.post(f'https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={key}', json={
                'contents': [{'parts': [{'text': prompt}, {'inline_data': {'mime_type': 'image/jpeg', 'data': b64}}]}],
                'generationConfig': {'response_mime_type': 'application/json', 'temperature': 0.1}
            }, timeout=25)
            if r.status_code == 200:
                print(f"Success with {model}:")
                ans = r.json()['candidates'][0]['content']['parts'][0]['text']
                print(ans)
                break
            else:
                print(f"Failed {model}: {r.status_code}")
        except Exception as e:
            print(f"Error {model}: {e}")

if __name__ == '__main__':
    main()
