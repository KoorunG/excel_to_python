from pathlib import Path
import pandas as pd

root = Path()
file_name = "서울특별시 중랑구 연령별 인구수 현황_20230427.xlsx"
io = root / file_name

# 엑셀에서 데이터프레임 가져오기
df1: pd.DataFrame = pd.read_excel(io, sheet_name=0, skiprows=1, header=0, index_col=[1, 2]).iloc[1:, 1:]

# unstack() 으로 인덱스 -> 컬럼으로 옮기기
df2 = df1.unstack(level=[0, 1]).reset_index()

# 컬럼명 재정의하기
df2.columns = ["지역", "연령", "성별", "인구수"]

# 성별이 "계"인 데이터 필터링, 인덱스 리셋
df3: pd.DataFrame = df2.iloc[3:].loc[df2["성별"] != "계"].reset_index(drop=True)

# pivot_table()으로 피봇팅
df4 = df3.pivot_table(index=["성별", "연령"], columns="지역", values="인구수")

# 프린트 & 클립보드에 복사
df4.to_clipboard()
print(df4)
