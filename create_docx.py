from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

doc = Document()

# 기본 스타일 설정
style = doc.styles['Normal']
style.font.name = '맑은 고딕'
style.font.size = Pt(11)

def add_heading(doc, text, level=1):
    h = doc.add_heading(text, level=level)
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run = h.runs[0]
    if level == 1:
        run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)
        run.font.size = Pt(18)
    elif level == 2:
        run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
        run.font.size = Pt(14)
    elif level == 3:
        run.font.color.rgb = RGBColor(0x2E, 0x74, 0xB5)
        run.font.size = Pt(12)
    return h

def add_table(doc, headers, rows):
    table = doc.add_table(rows=1+len(rows), cols=len(headers))
    table.style = 'Table Grid'
    # 헤더
    hdr = table.rows[0].cells
    for i, h in enumerate(headers):
        hdr[i].text = h
        hdr[i].paragraphs[0].runs[0].bold = True
        hdr[i].paragraphs[0].alignment = WD_ALIGN_PARAGRAPH.CENTER
    # 데이터
    for r_idx, row in enumerate(rows):
        cells = table.rows[r_idx+1].cells
        for c_idx, val in enumerate(row):
            cells[c_idx].text = val
    doc.add_paragraph()

# ── 타이틀 ──
title = doc.add_paragraph()
title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run = title.add_run('스마트 화장지 솔루션 기획서')
run.bold = True
run.font.size = Pt(24)
run.font.color.rgb = RGBColor(0x1F, 0x49, 0x7D)

sub = doc.add_paragraph()
sub.alignment = WD_ALIGN_PARAGRAPH.CENTER
sub.add_run('상태: 설계 확정  |  최종 업데이트: 2026-05-09').italic = True

doc.add_paragraph()

# ── 1. 아이디어 개요 ──
add_heading(doc, '1. 아이디어 개요', 1)
doc.add_paragraph('제품명 (가칭): 물티슈 없는 스마트 화장실 — 폼 분사 모듈')
doc.add_paragraph(
    '한 줄 정의: 공중화장실 점보롤 디스펜서에 부착하는 모듈형 장치로, '
    '화장지를 뽑는 물리적 힘으로 기계식 펌프를 작동시켜 거품(폼) 세정제를 자동 분사함으로써 '
    '전기 없이도 물티슈 수준의 위생 경험을 제공한다.'
)

# ── 2. 문제 정의 ──
add_heading(doc, '2. 문제 정의 (Pain Point)', 1)
doc.add_paragraph('공중화장실에서 대변 후 일반 화장지만으로는 청결감이 부족하지만:')
for item in [
    '물티슈는 변기 막힘 유발 (배관 문제)',
    '물티슈 쓰레기가 화장실 쓰레기통에 넘침',
    '비데는 설치 비용·공간 문제로 공중화장실 보급 어려움',
]:
    p = doc.add_paragraph(item, style='List Bullet')

doc.add_paragraph('→ "깨끗하게 닦고 싶지만 물티슈는 쓰기 꺼려지는" 마찰(Friction) 지점')

# ── 3. 시장 검증 ──
add_heading(doc, '3. 시장 검증 (블라인드 앱 설문, n=18)', 1)
add_table(doc,
    ['응답', '비율', '인원'],
    [
        ['물티슈가 모든 화장실에 있다면 쓴다 (= 잠재 고객)', '50.0%', '9명'],
        ['마른 휴지로도 충분하다', '33.3%', '6명'],
        ['물티슈도 귀찮아서 안 쓴다', '16.7%', '3명'],
    ]
)
p = doc.add_paragraph()
run = p.add_run('핵심 인사이트: ')
run.bold = True
p.add_run('응답자 절반이 더 쾌적한 대안이 있다면 기꺼이 사용할 의향 있음.')

# ── 4. 솔루션 설계 ──
add_heading(doc, '4. 솔루션 설계', 1)

add_heading(doc, '4-1. 하드웨어 구조', 2)
structure = doc.add_paragraph()
structure.add_run(
    '[ 점보롤 디스펜서 ]  ───  기존 그대로 유지\n'
    '        ↓\n'
    '[ 부착형 모듈 ]\n'
    '  ├─ 기계식 펌프: 휴지 인출 시 레버 연동 → 폼 미량 분사\n'
    '  ├─ 폼 카트리지: 교체형 (OEM 폼 → 추후 자체 포뮬러)\n'
    '  ├─ 분사 노즐: 화장지 뽑히는 지점에 위치\n'
    '  └─ 사용자 Lock 스위치: 원하지 않으면 직접 끔'
).font.name = 'Courier New'

p = doc.add_paragraph()
run = p.add_run('핵심 원리: ')
run.bold = True
p.add_run('휴지를 당기는 물리적 힘 → 기계식 펌프 압축 → 폼 분사. ')
p.add_run('전기·배터리 불필요.').bold = True

add_heading(doc, '4-2. 주요 결정 사항', 2)
add_table(doc,
    ['항목', '결정', '근거'],
    [
        ['분사 물질', '거품(폼) 세정제', '화장지 찢어짐 방지, 세정력 향상'],
        ['제품 형태', '모듈형 (기존 디스펜서 부착)', 'B2B 도입 장벽 최소화'],
        ['트리거', '기계식 레버 연동', '전원 불필요, 자연스러운 사용 흐름'],
        ['전원', '없음 (순수 기계식)', '설치·유지 비용 제로'],
        ['Lock', '사용자 직접 온/오프', '원하지 않는 사용자 선택권 보장'],
        ['폼 공급', 'OEM 시작 → 자체 개발', '빠른 출시 + 장기 차별화'],
        ['파트너십', '디스펜서 제조사 공동개발', '호환성 해결 + 유통망 확보'],
    ]
)

# ── 5. B2B 타겟 ──
add_heading(doc, '5. B2B 타겟 고객', 1)
for item in ['고속도로 휴게소', '공공기관 (정부·지자체 청사)', '대형 쇼핑몰·복합시설', '기업 오피스 빌딩']:
    doc.add_paragraph(item, style='List Bullet')

add_heading(doc, 'B2B 설득 포인트', 2)
add_table(doc,
    ['예상 반론', '방어 논리'],
    [
        ['"휴지 사용량 늘어나서 비용 증가 아닌가?"', '물티슈 변기 막힘 해결 비용 절감이 더 큼'],
        ['"설비 교체 비용이 부담스럽다"', '기존 디스펜서 그대로, 모듈만 부착'],
        ['"관리가 번거롭지 않나?"', '폼 카트리지 정기 구독으로 관리 일원화'],
        ['"전기 공사가 필요한가?"', '순수 기계식, 전원 불필요'],
    ]
)

# ── 6. 비즈니스 모델 ──
add_heading(doc, '6. 비즈니스 모델 & 수익 구조', 1)
phases = [
    ('Phase 1 (0~12개월): 디스펜서 제조사 파트너십',
     '공동개발 → 제조사 기존 유통망으로 시장 진입\n수익: 모듈 공급가 + 카트리지 초도물량'),
    ('Phase 2 (12~24개월): 카트리지 구독 정착',
     'B2B 구독 (월정액 or 카트리지 소진 시 자동 배송)\n데이터 확보 → 교체 주기·사용량 최적화'),
    ('Phase 3 (24개월+): 자체 폼 포뮬러 + 자체 브랜드',
     '특허 출원 → 경쟁사 진입 장벽\n독립 브랜드로 직접 B2B 영업 확장'),
]
for title_text, desc in phases:
    p = doc.add_paragraph()
    p.add_run(title_text).bold = True
    doc.add_paragraph(desc)

add_heading(doc, '수익 구조 요약', 2)
add_table(doc,
    ['항목', '수익원'],
    [
        ['모듈', '초기 판매 (원가 근접, 진입 유도)'],
        ['폼 카트리지', '정기 구독 (핵심 수익)'],
        ['파트너십', '라이선스 or 공동 수익 배분'],
    ]
)

# ── 7. 다음 단계 ──
add_heading(doc, '7. 다음 단계', 1)
for item in [
    '디스펜서 제조사 리스트업 및 파트너십 접촉',
    '기계식 펌프 프로토타입 제작',
    'OEM 폼 세정제 공급사 탐색',
    '특허 선행 조사 (기계식 펌프 + 디스펜서 연동)',
    '규제/인증 검토 (식약처 등 세정제 관련)',
    'MVP PoC 시범 설치 장소 확보',
]:
    doc.add_paragraph(f'☐  {item}')

output_path = '/home/user/Idea_2-Toilet-/docs/superpowers/specs/2026-05-09-smart-toilet-paper-design.docx'
doc.save(output_path)
print(f'저장 완료: {output_path}')
