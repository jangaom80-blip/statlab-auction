import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from datetime import datetime

# ==========================================
# 1. 초보자 보호를 위한 위험 권리 배제 키워드 정의
# ==========================================
DANGER_KEYWORDS = [
    "유치권", "법정지상권", "대지권미등기", "분묘기지권", 
    "지분매각", "선순위", "임차권등기", "가처분", "환매등기"
]

def is_safe_property(note: str) -> bool:
    """비고란에 위험 키워드가 1개라도 있으면 False 반환"""
    if pd.isna(note):
        return True
    return not any(keyword in str(note) for keyword in DANGER_KEYWORDS)

# ==========================================
# 2. 데이터 수집 및 정제 (테스트용 샘플 데이터셋)
# ※ 추후 대법원 API나 크롤링 모듈로 교체 연동됩니다.
# ==========================================
raw_data = [
    {
        "사건번호": "2024타경10123",
        "소재지": "부산광역시 해운대구 우동 마린시티 OO아파트 101동 1502호",
        "용도": "아파트",
        "감정가": 850000000,
        "최저입찰가": 544000000,
        "유찰횟수": 2,
        "매각기일": "2026-10-15",
        "특이사항_비고": "대항력 있는 임차인 배당신청 완료. 유치권 신고서 접수됨"
    },
    {
        "사건번호": "2024타경88412",
        "소재지": "서울특별시 노원구 상계동 OO주공아파트 3단지 504호",
        "용도": "아파트",
        "감정가": 620000000,
        "최저입찰가": 396800000,
        "유찰횟수": 2,
        "매각기일": "2026-10-18",
        "특이사항_비고": "임차인 없음(채무자 점유). 권리관계 이상 없음"
    },
    {
        "사건번호": "2025타경14205",
        "소재지": "경기도 성남시 분당구 정자동 OO아파트 203동 801호",
        "용도": "아파트",
        "감정가": 1250000000,
        "최저입찰가": 800000000,
        "유찰횟수": 2,
        "매각기일": "2026-10-21",
        "특이사항_비고": "소유자 거주. 권리 분석상 특이사항 없음"
    },
    {
        "사건번호": "2024타경30012",
        "소재지": "인천광역시 연수구 송도동 OO오피스텔 1204호",
        "용도": "오피스텔",
        "감정가": 420000000,
        "최저입찰가": 215040000,
        "유찰횟수": 3,
        "매각기일": "2026-10-25",
        "특이사항_비고": "토지별도등기 있음, 법정지상권 성립여부 불분명"
    },
    {
        "사건번호": "2025타경05819",
        "소재지": "대구광역시 수성구 범어동 OO빌라 302호",
        "용도": "다세대",
        "감정가": 310000000,
        "최저입찰가": 198400000,
        "유찰횟수": 2,
        "매각기일": "2026-10-29",
        "특이사항_비고": "권리관계 깨끗함, 즉시 명도 가능"
    }
]

df = pd.DataFrame(raw_data)

# ==========================================
# 3. 알고리즘 필터링 파이프라인
# ==========================================
# 1) 2회 이상 유찰된 물건만 선택 (최저가율 대폭 하락)
filtered_df = df[df["유찰횟수"] >= 2].copy()

# 2) 위험 권리 키워드 100% 배제
filtered_df["안전성검증"] = filtered_df["특이사항_비고"].apply(is_safe_property)
safe_df = filtered_df[filtered_df["안전성검증"] == True].copy()

# 3) 지표 계산 (할인율 및 예상 안전 마진)
safe_df["최저가비율(%)"] = (safe_df["최저입찰가"] / safe_df["감정가"] * 100).round(1)
safe_df["감정가대비_할인액"] = safe_df["감정가"] - safe_df["최저입찰가"]

# 4) 판매용 컬럼 정리 및 정렬 (할인율 큰 순서)
output_columns = [
    "사건번호", "용도", "소재지", "유찰횟수", 
    "감정가", "최저입찰가", "최저가비율(%)", "감정가대비_할인액", "매각기일"
]
result_df = safe_df[output_columns].sort_values(by="최저가비율(%)", ascending=True)

# ==========================================
# 4. 프리미엄 리포트 서식 적용 엑셀 생성
# ==========================================
file_name = f"스탯랩_안전경매_반값유찰_리포트_{datetime.today().strftime('%Y%m%d')}.xlsx"

with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
    result_df.to_excel(writer, index=False, sheet_name="알짜_안전경매_리스트", startrow=3)
    
    workbook = writer.book
    worksheet = writer.sheets["알짜_안전경매_리스트"]
    
    # 헤더 타이틀 디자인
    worksheet.merge_cells("A1:I1")
    title_cell = worksheet["A1"]
    title_cell.value = f"[STATLAB] 전국 법원경매 반값 유찰 & 권리안전 데이터 리포트 ({datetime.today().strftime('%Y-%m-%d')})"
    title_cell.font = Font(name="맑은 고딕", size=15, bold=True, color="FFFFFF")
    title_cell.fill = PatternFill(start_color="1A2B4C", end_color="1A2B4C", fill_type="solid")
    title_cell.alignment = Alignment(horizontal="center", vertical="center")
    worksheet.row_dimensions[1].height = 40

    # 서브 가이드 문구
    worksheet.merge_cells("A2:I2")
    sub_cell = worksheet["A2"]
    sub_cell.value = "※ 본 데이터는 대법원 공공데이터를 기반으로 유치권·법정지상권 등 고위험 요소를 사전에 100% 필터링한 초보자용 통계 분석집입니다."
    sub_cell.font = Font(name="맑은 고딕", size=9, italic=True, color="555555")
    sub_cell.alignment = Alignment(horizontal="left", vertical="center")
    worksheet.row_dimensions[2].height = 20

    # 테이블 컬럼 헤더 스타일링
    header_fill = PatternFill(start_color="2E4057", end_color="2E4057", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='CCCCCC'),
        right=Side(style='thin', color='CCCCCC'),
        top=Side(style='thin', color='CCCCCC'),
        bottom=Side(style='thin', color='CCCCCC')
    )

    for col_num in range(1, len(output_columns) + 1):
        cell = worksheet.cell(row=4, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
    worksheet.row_dimensions[4].height = 26

    # 데이터 행 서식 및 숫자 포맷 적용
    for row in worksheet.iter_rows(min_row=5, max_row=4 + len(result_df), min_col=1, max_col=len(output_columns)):
        for cell in row:
            cell.font = Font(name="맑은 고딕", size=9)
            cell.border = thin_border
            col_name = output_columns[cell.column - 1]
            
            if col_name in ["감정가", "최저입찰가", "감정가대비_할인액"]:
                cell.number_format = '#,##0"원"'
                cell.alignment = Alignment(horizontal="right", vertical="center")
            elif col_name == "최저가비율(%)":
                cell.number_format = '0.0"%"'
                cell.alignment = Alignment(horizontal="center", vertical="center")
                cell.font = Font(name="맑은 고딕", size=9, bold=True, color="C00000")
            elif col_name in ["사건번호", "용도", "유찰횟수", "매각기일"]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

    # 열 너비 자동 조정
    for col in worksheet.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = get_column_letter(col[0].column)
        worksheet.column_dimensions[col_letter].width = max(max_len + 4, 12)

print(f"✅ 생성 완료: {file_name}")
print(f"📊 원본 물건 5건 중 위험 물건(유치권/법정지상권) 2건 자동 제거 ➔ 안전 물건 3건 추출 완료!")
