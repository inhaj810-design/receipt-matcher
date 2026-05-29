import streamlit as st
import pandas as pd
import google.generativeai as genai
from PIL import Image
import json
import math
import io
import base64
from datetime import datetime
from collections import defaultdict
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="영수증 HCP 매칭", page_icon="🧾", layout="wide")

# ─── CSS ───────────────────────────────────────────────
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Noto+Sans+KR:wght@300;400;500;700&family=DM+Mono:wght@400;500&display=swap');

html, body, [class*="css"] { font-family: 'Noto Sans KR', sans-serif; }

.main { background: #f8f9fc; }

.stApp { background: #f8f9fc; }

.hero {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 50%, #0f3460 100%);
    border-radius: 16px;
    padding: 40px 48px;
    margin-bottom: 32px;
    color: white;
    position: relative;
    overflow: hidden;
}
.hero::before {
    content: '';
    position: absolute;
    top: -50%;
    right: -10%;
    width: 400px;
    height: 400px;
    background: radial-gradient(circle, rgba(83,166,255,0.15) 0%, transparent 70%);
    border-radius: 50%;
}
.hero h1 { font-size: 2rem; font-weight: 700; margin: 0 0 8px 0; letter-spacing: -0.5px; }
.hero p { font-size: 0.95rem; opacity: 0.7; margin: 0; font-weight: 300; }

.step-card {
    background: white;
    border-radius: 12px;
    padding: 24px;
    margin-bottom: 16px;
    border: 1px solid #e8ecf0;
    box-shadow: 0 2px 8px rgba(0,0,0,0.04);
}
.step-num {
    display: inline-block;
    background: #0f3460;
    color: white;
    width: 28px;
    height: 28px;
    border-radius: 50%;
    text-align: center;
    line-height: 28px;
    font-size: 13px;
    font-weight: 700;
    margin-right: 10px;
    font-family: 'DM Mono', monospace;
}
.step-title { font-size: 1rem; font-weight: 600; color: #1a1a2e; }

.result-card {
    background: white;
    border-radius: 12px;
    padding: 20px 24px;
    border-left: 4px solid #0f3460;
    box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    margin-bottom: 12px;
}
.tag {
    display: inline-block;
    padding: 2px 10px;
    border-radius: 20px;
    font-size: 12px;
    font-weight: 500;
    margin-right: 6px;
}
.tag-hospital { background: #e8f0fe; color: #1a73e8; }
.tag-grade-s { background: #fce8e6; color: #d93025; }
.tag-grade-a { background: #e6f4ea; color: #137333; }
.tag-grade-b { background: #fef7e0; color: #b06000; }
.tag-grade-c { background: #f1f3f4; color: #5f6368; }

.warn-box {
    background: #fff8e1;
    border: 1px solid #ffd54f;
    border-radius: 8px;
    padding: 12px 16px;
    font-size: 13px;
    color: #795548;
}
.success-box {
    background: #e8f5e9;
    border: 1px solid #a5d6a7;
    border-radius: 8px;
    padding: 12px 16px;
    font-size: 13px;
    color: #2e7d32;
}
</style>
""", unsafe_allow_html=True)

# ─── HERO ──────────────────────────────────────────────
st.markdown("""
<div class="hero">
    <h1>🧾 영수증 HCP 매칭 시스템</h1>
    <p>영수증 사진과 방문기록 엑셀을 업로드하면 자동으로 HCP를 매칭해드립니다</p>
</div>
""", unsafe_allow_html=True)

# ─── STEP 1: API KEY ────────────────────────────────────
st.markdown("""<div class="step-card">
<span class="step-num">1</span><span class="step-title">Gemini API 키 입력</span>
</div>""", unsafe_allow_html=True)

api_key = st.text_input("Gemini API Key", type="password", placeholder="AIza...")
if api_key:
    genai.configure(api_key=api_key)
    st.success("✅ API 키 연결됨")

# ─── STEP 2: 규정 설정 ─────────────────────────────────
st.markdown("""<div class="step-card">
<span class="step-num">2</span><span class="step-title">매칭 규정 설정</span>
</div>""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    min_days = st.number_input("동일 HCP 최소 간격 (일)", min_value=1, max_value=30, value=7)
with col2:
    max_per_month = st.number_input("월 최대 배정 횟수", min_value=1, max_value=10, value=4)

# ─── STEP 3: 파일 업로드 ────────────────────────────────
st.markdown("""<div class="step-card">
<span class="step-num">3</span><span class="step-title">파일 업로드</span>
</div>""", unsafe_allow_html=True)

col1, col2 = st.columns(2)
with col1:
    receipt_files = st.file_uploader(
        "영수증 사진 (여러 장 가능)",
        type=["png", "jpg", "jpeg"],
        accept_multiple_files=True
    )
    if receipt_files:
        st.info(f"📸 {len(receipt_files)}장 업로드됨")

with col2:
    excel_file = st.file_uploader("방문기록 엑셀", type=["xlsx", "xls"])
    if excel_file:
        st.info(f"📊 {excel_file.name} 업로드됨")

# ─── 영수증 분석 함수 ────────────────────────────────────
def analyze_receipt_with_gemini(image_file):
    model = genai.GenerativeModel('gemini-2.5-flash')
    img = Image.open(image_file)
    prompt = """이 영수증 사진에서 다음 정보를 추출해서 JSON으로만 응답해줘. 다른 텍스트 없이 JSON만.
{
  "shop_name": "가맹점명 (영수증에 적힌 그대로)",
  "address": "주소 (있으면)",
  "date": "YYYY-MM-DD",
  "time": "HH:MM",
  "amount": 숫자만(원),
  "approval_number": "승인번호"
}"""
    response = model.generate_content([prompt, img])
    text = response.text.strip()
    if text.startswith("```"):
        text = text.split("```")[1]
        if text.startswith("json"):
            text = text[4:]
    return json.loads(text.strip())

# ─── 방문기록 읽기 함수 ──────────────────────────────────
def load_visit_records(excel_file):
    df = pd.read_excel(excel_file, sheet_name=0, header=None)
    visits = []
    for i in range(2, len(df)):
        row = df.iloc[i]
        if pd.notna(row[5]) and pd.notna(row[10]) and pd.notna(row[11]):
            visit_date = str(row[5])[:10]
            hospital = str(row[10])
            hcp = str(row[11])
            dept = str(row[13]) if pd.notna(row[13]) else ''
            grade = str(row[12]) if pd.notna(row[12]) else ''
            if hospital not in ['출근', '퇴근', '']:
                visits.append({
                    '방문일자': visit_date, '거래처': hospital,
                    'HCP': hcp, '진료과': dept, '등급': grade
                })
    return visits

# ─── 병원 매핑 함수 (주소/가맹점명 기반) ─────────────────
def guess_hospital(shop_name, address, date, visit_dates_hospitals):
    """가맹점명과 주소로 근처 병원 추측"""
    keywords = {
        '안동성소': ['안동성소병원'],
        '안동병원': ['의료법인안동병원'],
        '가톨릭대': ['대구가톨릭대학교병원'],
        '가톨릭병원': ['대구가톨릭대학교병원'],
        '계명대': ['계명대학교동산병원'],
        '동산병원': ['계명대학교동산병원'],
        '칠곡경북대': ['칠곡경북대학교병원'],
        '경북대병원': ['칠곡경북대학교병원'],
        '구미': ['순천향대학교 부속 구미병원'],
        '순천향': ['순천향대학교 부속 구미병원'],
        '곽병원': ['곽병원'],
    }
    text = (shop_name or '') + (address or '')
    for kw, hospitals in keywords.items():
        if kw in text:
            return hospitals

    # 주소에서 지역 파악
    if address:
        if '안동' in address:
            # 안동에서 해당 날짜 방문한 병원
            visited = visit_dates_hospitals.get(date, [])
            andong = [h for h in visited if '안동' in h]
            if andong:
                return andong
        if '구미' in address:
            return ['순천향대학교 부속 구미병원']
        if '대구' in address:
            visited = visit_dates_hospitals.get(date, [])
            daegu = [h for h in visited if '대구' in h or '계명' in h or '칠곡' in h or '곽' in h]
            if daegu:
                return daegu

    # 해당 날짜 방문한 모든 병원 반환
    return visit_dates_hospitals.get(date, [])

# ─── 인원수 계산 ─────────────────────────────────────────
def calc_persons(amount):
    return math.ceil(amount / 10000)

# ─── HCP 매칭 ────────────────────────────────────────────
def match_hcp(receipts, visits, min_days, max_per_month):
    grade_order = {'S':0,'A1':1,'A2':2,'B1':3,'B2':4,'B3':5,'C1':6,'C2':7,'D1':8,'D2':9}
    def grade_key(v): return grade_order.get(v['등급'], 99)

    visit_by_date_hospital = defaultdict(list)
    for v in visits:
        visit_by_date_hospital[(v['방문일자'], v['거래처'])].append(v)

    visit_dates_hospitals = defaultdict(list)
    for v in visits:
        if v['거래처'] not in visit_dates_hospitals[v['방문일자']]:
            visit_dates_hospitals[v['방문일자']].append(v['거래처'])

    hcp_assign_dates = defaultdict(list)

    def can_assign(hcp, date_str):
        d = datetime.strptime(date_str, '%Y-%m-%d')
        for pd_ in hcp_assign_dates[hcp]:
            if abs((d - pd_).days) < min_days:
                return False
        mc = sum(1 for pd_ in hcp_assign_dates[hcp]
                 if pd_.month == d.month and pd_.year == d.year)
        return mc < max_per_month

    results = []
    for r in receipts:
        date = r['date']
        n = r['persons']
        hospitals = r.get('hospitals') or guess_hospital(
            r.get('shop_name',''), r.get('address',''), date, visit_dates_hospitals)

        candidates = []
        for hosp in (hospitals or []):
            for c in visit_by_date_hospital.get((date, hosp), []):
                candidates.append((hosp, c))
        candidates.sort(key=lambda x: grade_key(x[1]))

        selected = []
        for hosp, v in candidates:
            if len(selected) >= n: break
            if can_assign(v['HCP'], date):
                selected.append((hosp, v))
                hcp_assign_dates[v['HCP']].append(datetime.strptime(date, '%Y-%m-%d'))

        shortage = n - len(selected)
        warn = f'⚠ {n}명 필요 / {len(selected)}명 배정' if shortage > 0 else ''

        if selected:
            for idx, (hosp, v) in enumerate(selected):
                results.append({
                    'No': r['no'], '승인일자': date, '시간': r['time'],
                    '가맹점': r['shop_name'], '주소': r.get('address',''),
                    '승인금액': r['amount'], '선정인원': n, '실배정': len(selected),
                    '병원': hosp, 'HCP': v['HCP'], '진료과': v['진료과'], '등급': v['등급'],
                    '비고': warn if idx == 0 else ''
                })
        else:
            results.append({
                'No': r['no'], '승인일자': date, '시간': r['time'],
                '가맹점': r['shop_name'], '주소': r.get('address',''),
                '승인금액': r['amount'], '선정인원': n, '실배정': 0,
                '병원': ', '.join(hospitals or ['미확인']),
                'HCP': '(매칭 없음)', '진료과': '', '등급': '', '비고': warn or '⚠ HCP 없음'
            })
    return results

# ─── 엑셀 생성 ───────────────────────────────────────────
def make_excel(results):
    wb = Workbook()
    ws = wb.active
    ws.title = '매칭결과'
    hf = Font(name='Malgun Gothic', bold=True, color='FFFFFF', size=10)
    hfill = PatternFill('solid', start_color='1a1a2e')
    wfill = PatternFill('solid', start_color='FFF8E1')
    fills = [PatternFill('solid', start_color='FFFFFF'),
             PatternFill('solid', start_color='EEF2FA')]
    c = Alignment(horizontal='center', vertical='center', wrap_text=True)
    l = Alignment(horizontal='left', vertical='center', wrap_text=True)
    thin = Side(style='thin', color='CCCCCC')
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    headers = ['No','승인일자','시간','가맹점','주소','승인금액','선정인원','실배정','병원','HCP','진료과','등급','비고']
    widths =  [5,   13,      8,    24,    28,    12,      9,       9,      22,    10,    12,    8,    28]
    for ci, (h, w) in enumerate(zip(headers, widths), 1):
        cell = ws.cell(row=1, column=ci, value=h)
        cell.font = hf; cell.fill = hfill; cell.alignment = c; cell.border = bd
        ws.column_dimensions[get_column_letter(ci)].width = w
    ci_toggle = 0; prev_no = None; rn = 2
    for r in results:
        if r['No'] != prev_no:
            ci_toggle = (ci_toggle + 1) % 2; prev_no = r['No']
        fill = wfill if r['비고'] else fills[ci_toggle]
        vals = [r['No'],r['승인일자'],r['시간'],r['가맹점'],r['주소'],
                r['승인금액'],r['선정인원'],r['실배정'],r['병원'],r['HCP'],r['진료과'],r['등급'],r['비고']]
        for ci, val in enumerate(vals, 1):
            cell = ws.cell(row=rn, column=ci, value=val)
            cell.fill = fill; cell.border = bd
            cell.font = Font(name='Malgun Gothic', size=10)
            cell.alignment = c if ci in [1,2,3,6,7,8,12] else l
            if ci == 6: cell.number_format = '#,##0'
        if ws.cell(row=rn, column=12).value in ['S','A1','A2']:
            ws.cell(row=rn, column=12).font = Font(name='Malgun Gothic', size=10, bold=True, color='0f3460')
        rn += 1
    ws.freeze_panes = 'A2'
    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf

# ─── 영수증 합본 이미지 생성 ─────────────────────────────
def make_collage(images, cols=5):
    THUMB_W = 500; GAP = 8; BG = (235, 237, 242)
    rows = math.ceil(len(images) / cols)
    imgs = []
    for img in images:
        r = THUMB_W / img.width
        imgs.append(img.resize((THUMB_W, int(img.height * r)), Image.LANCZOS))
    row_heights = [max(imgs[r*cols:(r+1)*cols], key=lambda x: x.height).height for r in range(rows)]
    W = cols * THUMB_W + (cols+1)*GAP
    H = sum(row_heights) + (rows+1)*GAP
    canvas = Image.new('RGB', (W, H), BG)
    for idx, img in enumerate(imgs):
        row = idx // cols
        col = idx % cols
        x = GAP + (cols-1-col)*(THUMB_W+GAP)
        y = GAP + sum(row_heights[:row]) + row*GAP
        canvas.paste(img, (x, y))
    return canvas

# ─── 메인 실행 ───────────────────────────────────────────
if st.button("🚀 매칭 시작", type="primary", use_container_width=True,
             disabled=not(api_key and receipt_files and excel_file)):

    st.divider()
    st.subheader("📋 처리 결과")

    # 방문기록 로드
    with st.spinner("방문기록 엑셀 읽는 중..."):
        visits = load_visit_records(excel_file)
        st.success(f"✅ 방문기록 {len(visits)}건 로드 완료")

    # 영수증 분석
    st.write("**영수증 분석 중...**")
    receipt_data = []
    sorted_files = sorted(receipt_files, key=lambda f: f.name)
    prog = st.progress(0)
    for i, f in enumerate(sorted_files):
        with st.spinner(f"분석 중: {f.name}"):
            try:
                f.seek(0)
                info = analyze_receipt_with_gemini(f)
                info['no'] = i + 1
                info['persons'] = calc_persons(info.get('amount', 0))
                receipt_data.append(info)
                st.write(f"  ✅ {f.name} → {info.get('shop_name','')} / {info.get('amount',0):,}원 / {info.get('persons',0)}명")
            except Exception as e:
                st.warning(f"  ⚠ {f.name} 분석 실패: {e}")
        prog.progress((i+1)/len(sorted_files))

    # 날짜+시간 순 정렬
    receipt_data.sort(key=lambda x: (x.get('date',''), x.get('time','')))
    for i, r in enumerate(receipt_data):
        r['no'] = i + 1

    # HCP 매칭
    with st.spinner("HCP 매칭 중..."):
        results = match_hcp(receipt_data, visits, min_days, max_per_month)

    # 결과 표시
    st.success(f"✅ 총 {len(results)}건 매칭 완료")
    df_result = pd.DataFrame(results)
    st.dataframe(df_result, use_container_width=True)

    # 엑셀 다운로드
    excel_buf = make_excel(results)
    st.download_button("📥 결과 엑셀 다운로드", excel_buf,
                       file_name="HCP_매칭결과.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # 영수증 합본 이미지
    st.write("**영수증 합본 이미지 생성 중...**")
    pil_images = []
    for f in sorted_files:
        f.seek(0)
        pil_images.append(Image.open(f).convert('RGB'))

    for batch_idx in range(0, len(pil_images), 10):
        batch = pil_images[batch_idx:batch_idx+10]
        collage = make_collage(batch)
        buf = io.BytesIO()
        collage.save(buf, 'PNG')
        buf.seek(0)
        st.image(collage, caption=f"합본 {batch_idx//10 + 1}", use_column_width=True)
        st.download_button(
            f"📥 합본 이미지 {batch_idx//10 + 1} 다운로드",
            buf, file_name=f"영수증_합본_{batch_idx//10+1}.png",
            mime="image/png"
        )
