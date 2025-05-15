import os
import pandas as pd
import openai
from collections import defaultdict

# 📁 엑셀 파일 디렉토리 경로
directory = r"C:\Users\user\Desktop\thesis"

# 📄 엑셀 파일 목록 불러오기
excel_files = [f for f in os.listdir(directory) if f.endswith(('.xlsx', '.xls'))]
print("엑셀 파일 목록:")
for idx, file in enumerate(excel_files):
    print(f"{idx + 1}. {file}")

# 🔢 사용자 입력으로 파일 선택
file_num = int(input("\n분석할 파일의 번호를 입력하세요: "))
selected_file = os.path.join(directory, excel_files[file_num - 1])

# 📥 파일 불러오기
df = pd.read_excel(selected_file)

# 🔍 컬럼명 확인 및 변경
print("컬럼명 확인:", df.columns.tolist())
df = df.reset_index().rename(columns={"index": "Number"})
df.rename(columns={"제목": "Title", "국문 초록 (Abstract)": "Abstract"}, inplace=True)

# 🧹 중복 논문 제거
duplicates_dict = defaultdict(list)
for idx, row in df.iterrows():
    title = row.get("Title", "")
    if pd.isna(title):
        continue
    duplicates_dict[title].append(row["Number"])

duplicated_pairs = [tuple(numbers) for numbers in duplicates_dict.values() if len(numbers) > 1]
print(f"\n중복된 논문 쌍: {duplicated_pairs}")
print(f"중복된 논문 총 개수: {sum(len(pair) for pair in duplicated_pairs)}개")

df = df.drop_duplicates(subset=["Title"], keep="first")

# 🕳 초록이 없는 논문
no_abstract = df[df["Abstract"].isnull() | (df["Abstract"].astype(str).str.strip() == "")]["Number"].tolist()
print(f"\n초록이 없는 논문 번호: {no_abstract}")
print(f"초록이 없는 논문 총 개수: {len(no_abstract)}개")

# 🔍 키워드 필터링
def is_chatbot_study(abstract):
    chatbot_keywords = ["챗봇", "Chatbot"]
    return any(keyword.lower() in abstract.lower() for keyword in chatbot_keywords)

no_chatbot_papers = []
for idx, row in df.iterrows():
    abstract = str(row["Abstract"]) if pd.notna(row["Abstract"]) else ""
    if row["Number"] not in no_abstract and not is_chatbot_study(abstract):
        no_chatbot_papers.append(row["Number"])

print(f"\n챗봇 관련 키워드가 없는 논문 번호: {no_chatbot_papers}")
print(f"챗봇 관련 키워드가 없는 논문 개수: {len(no_chatbot_papers)}개")

# 🔐 OpenAI API 키 환경 변수
os.environ["OPENAI_API_KEY"] = "sk-..."  # 실제 키로 교체하거나 환경변수 설정
client = openai.OpenAI()

# 🤖 GPT 관련성 판단 함수
def check_relevance(abstract):
    prompt = f"""
다음 초록을 읽고, 이 논문이 '챗봇', 'chatbot'을 연구의 핵심 주제로 다루는지 판단해주세요.

🛑 주의: 아래 기준을 엄격히 적용하세요.
✅ '관련 있음':
1. 연구의 핵심 주제가 챗봇과 직접적으로 관련
2. 챗봇 관련 정책, 제도, 프로그램, 사례 연구 포함
3. 실험, 분석, 모델 중심에 챗봇이 있음

❌ '단순 언급' 또는 '관련 없음':
- 배경 설명 수준에서만 등장
- 결론/서론에서 잠깐 언급
- 주요 연구 초점과 무관

반드시 다음 중 하나로만 답해주세요:
- 관련 있음
- 단순 언급
- 관련 없음

초록:
\"\"\"{abstract}\"\"\"
"""
    try:
        response = client.chat.completions.create(
            model="gpt-4o",
            messages=[
                {"role": "system", "content": "당신은 논문 초록에서 챗봇 관련성을 분류하는 AI입니다."},
                {"role": "user", "content": prompt}
            ],
            temperature=0
        )
        return response.choices[0].message.content.strip()
    except Exception as e:
        print(f"GPT 판단 실패 (Number={row['Number']}): {e}")
        return "오류"

# 🧠 GPT 필터링
unrelated_papers = []
minor_mention_papers = []

for index, row in df.iterrows():
    number = row["Number"]
    if number in no_abstract or number in no_chatbot_papers:
        continue

    abstract = str(row["Abstract"]).strip()
    result = check_relevance(abstract)

    if result == "관련 없음":
        unrelated_papers.append(number)
    elif result == "단순 언급":
        minor_mention_papers.append(number)

print(f"\nGPT - 관련 없음: {unrelated_papers}")
print(f"GPT - 단순 언급: {minor_mention_papers}")

# ✅ 최종 필터링 결과
excluded_numbers = set(no_abstract + no_chatbot_papers + unrelated_papers + minor_mention_papers)
filtered_df = df[~df["Number"].isin(excluded_numbers)]

print(f"\n최종 필터링된 논문 수: {len(filtered_df)}개")

# 💾 저장
output_path = os.path.join(directory, "filtered_output_step.xlsx")
filtered_df.to_excel(output_path, index=False)
print(f"\n✅ 필터링 완료. 저장 경로: {output_path}")
