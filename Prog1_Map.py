import pandas as pd
import re

# 1. 엑셀 파일 열기
def load_excel(file_path):
    """엑셀 파일을 읽어오는 함수"""
    return pd.read_excel(file_path)

# 2. URL에서 위도와 경도 추출
def extract_lat_lng(url):
    """Google Maps URL에서 위도와 경도를 추출하는 함수"""
    try:
        match = re.search(r"@([-.\d]+),([-.\d]+)", url)
        if match:
            lat, lng = match.groups()
            return float(lat), float(lng)
        else:
            return None, None
    except Exception as e:
        print(f"Error processing URL {url}: {e}")
        return None, None

# 3. URL로부터 위도와 경도를 데이터프레임에 추가
def add_lat_lng_to_df(df, url_column):
    """URL 열에서 위도와 경도를 추출해 데이터프레임에 추가"""
    df[['위도', '경도']] = df[url_column].apply(lambda x: pd.Series(extract_lat_lng(x)))
    return df

# 4. 결과를 저장
def save_to_excel(df, output_path):
    """결과를 엑셀 파일로 저장"""
    df.to_excel(output_path, index=False)
    print(f"Results saved to {output_path}")

# 5. 메인 실행 함수
def main(input_file, output_file, url_column):
    """전체 흐름을 실행하는 메인 함수"""
    # Step 1: 엑셀 파일 로드
    df = load_excel(input_file)
    
    # Step 2: URL에서 위도와 경도 추출
    df = add_lat_lng_to_df(df, url_column)
    
    # Step 3: 결과 저장
    save_to_excel(df, output_file)

# 실행: 파일 경로 및 열 이름 설정
if __name__ == "__main__":
    input_file = "Honeymoon_Locations.xlsx"  # 입력 파일 경로
    output_file = "Honeymoon_Locations_with_Coordinates.xlsx"  # 출력 파일 경로
    url_column = "url"  # URL이 있는 열 이름

    main(input_file, output_file, url_column)


###############################################################

import pandas as pd
import folium

# 1. 엑셀 파일 로드
input_file = "Honeymoon_Locations_with_Coordinates.xlsx"  # 위도/경도 포함 엑셀 파일
df = pd.read_excel(input_file)

# 2. 지도 초기화
world_map = folium.Map(location=[0, 0], zoom_start=2)

# 3. 카테고리별 색상 설정
color_dict = {
    "공항": "blue",
    "숙박": "green",
    "관광지": "red",
    "기타": "purple"
}

# 4. 중복 숫자 마커 처리를 위한 설정
offsets = [0, 20, 30, 40]  # 중복 마커의 숫자 위치를 조정하기 위한 오프셋 값
seen_coordinates = {}

# 5. 마커 추가
for idx, row in df.iterrows():
    if not pd.isna(row['위도']) and not pd.isna(row['경도']):  # 유효한 위도/경도만 처리
        lat, lng = row['위도'], row['경도']
        coord_key = f"{lat},{lng}"  # 좌표를 문자열로 변환하여 키로 사용

        # 중복된 위치인지 확인
        if coord_key in seen_coordinates:
            count = seen_coordinates[coord_key]
            offset_y = offsets[count % len(offsets)]  # 순환적으로 오프셋 값 선택
            seen_coordinates[coord_key] += 1
        else:
            offset_y = 0  # 중복이 없으면 오프셋 없음
            seen_coordinates[coord_key] = 1

        category = row.get('분류', '기타')  # 카테고리가 없으면 '기타'

        # 팝업 내용을 HTML 스타일로 지정
        popup_content = f"""
        <div style="
            font-size:14px;
            text-align:left;
            white-space:nowrap;
            max-width:200px;">
            {row['장소명']} ({category})
        </div>
        """

        # 카테고리 색상 마커
        folium.Marker(
            location=[lat, lng],
            popup=folium.Popup(popup_content, max_width=250),  # 한 줄로 표현되도록 설정
            icon=folium.Icon(color=color_dict.get(category, "gray")),
        ).add_to(world_map)
        
        # 순서 번호를 텍스트로 표시 (중복된 경우만 위치 조정)
        folium.Marker(
            location=[lat, lng],
            icon=folium.DivIcon(
                html=f"""
                <div style="
                    font-size:14px; 
                    font-weight:bold; 
                    color:black; 
                    background-color:white; 
                    border:1px solid black; 
                    border-radius:50%; 
                    padding:5px; 
                    text-align:center; 
                    width:30px; 
                    height:30px;
                    position: relative;
                    top: {offset_y}px;">
                    {idx + 1}
                </div>"""
            ),
        ).add_to(world_map)

# 6. 경로 추가
route_coords = df[['위도', '경도']].dropna().values.tolist()  # NaN 제거 후 리스트로 변환

# 지도에 PolyLine 추가
folium.PolyLine(
    locations=route_coords,
    color="blue",  # 경로 색상
    weight=3,      # 선 두께
    opacity=0.8    # 선 투명도
).add_to(world_map)

# 7. 지도를 HTML 파일로 저장
output_map = "index.html"
world_map.save(output_map)

print(f"지도 생성 완료: {output_map}")
