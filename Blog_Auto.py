import random
import requests
import pandas as pd
from pytrends.request import TrendReq
import openai
from dotenv import load_dotenv
import os
import tkinter as tk
from tkinter import filedialog, messagebox

# Load .env file
load_dotenv()

# Step 1: Generate Keywords
def generate_keywords(topic):
    pytrends = TrendReq(hl='ko', tz=360)
    pytrends.build_payload([topic], cat=0, timeframe='now 7-d', geo='KR')  # 최근 일주일 데이터
    try:
        related_queries = pytrends.related_queries()
        print("관련 키워드 데이터:", related_queries)  # 디버깅: 관련 키워드 데이터 출력
        if related_queries and topic in related_queries and related_queries[topic]['top'] is not None:
            top_queries = related_queries[topic]['top']
            if not top_queries.empty:
                keywords = top_queries['query'].tolist()[:10]
                print("생성된 키워드:", keywords)  # 디버깅: 생성된 키워드 출력
                return keywords
        print("관련 키워드를 찾을 수 없습니다. 기본 키워드를 반환합니다.")
        return ["여행 준비", "국내 여행", "서울 여행 팁"]
    except Exception as e:
        print(f"키워드 생성 중 오류 발생: {e}")
        return ["여행 준비", "국내 여행", "서울 여행 팁"]

# Step 2: Generate Blog Content
def generate_content(keyword):
    openai.api_key = os.getenv("OPENAI_API_KEY")
    print("사용 중인 OpenAI API 키:", openai.api_key)  # 디버깅: API 키 확인
    messages = [
        {"role": "system", "content": "당신은 SEO 전문가로서 블로그 글을 작성하는 AI입니다."},
        {"role": "user", "content": f"'{keyword}'를 주제로 SEO 친화적인 블로그 글을 작성해주세요. 글 길이는 2000자 이상이며, 질문과 답변, 예시를 포함하세요."}
    ]
    try:
        response = openai.ChatCompletion.create(
            model="gpt-4",
            messages=messages,
            max_tokens=1500,  # 토큰 제한 줄임
            temperature=0.7
        )
        content = response['choices'][0]['message']['content']
        print("생성된 본문 일부:", content[:200])  # 디버깅: 본문 일부 출력
        return content
    except Exception as e:
        print(f"본문 생성 중 오류 발생: {e}")
        return "본문 생성 실패."

# Step 3: Fetch Image
def fetch_image(keyword):
    access_key = os.getenv("UNSPLASH_ACCESS_KEY")
    print("사용 중인 Unsplash API 키:", access_key)  # 디버깅: API 키 확인
    url = f"https://api.unsplash.com/search/photos?query={keyword}&client_id={access_key}"
    try:
        response = requests.get(url).json()
        print("이미지 API 응답:", response)  # 디버깅: API 응답 출력
        if response['results']:
            return response['results'][0]['urls']['regular']
        else:
            print("이미지를 찾을 수 없습니다. 기본 이미지를 반환합니다.")
            return "https://via.placeholder.com/800x400?text=No+Image+Available"
    except Exception as e:
        print(f"이미지 가져오기 오류 발생: {e}")
        return "https://via.placeholder.com/800x400?text=Error"

# Step 4: Save Data to Excel
def save_to_excel(title, keyword, image_url, content, filename="blog_data.xlsx"):
    data = {
        "제목": [title],
        "키워드": [keyword],
        "이미지 URL": [image_url],
        "본문": [content]
    }
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"엑셀 파일 저장 완료: {filename}")  # 디버깅: 파일 저장 확인

# Test OpenAI API Connection
def test_openai():
    try:
        response = openai.ChatCompletion.acreate(  # 최신 비동기 방식 지원
            model="gpt-4",
            messages=[{"role": "user", "content": "테스트 문장을 생성해 주세요."}]
        )
        print("테스트 응답:", response.choices[0].message["content"])
    except Exception as e:
        print(f"OpenAI API 테스트 실패: {e}")

# Tkinter UI
def create_ui():
    def generate_blog():
        topic = topic_entry.get()
        if not topic:
            messagebox.showerror("오류", "주제를 입력해주세요.")
            return

        try:
            print("블로그 생성 작업 시작...")
            keywords = generate_keywords(topic)
            selected_keyword = keywords[0]
            print("선택된 키워드:", selected_keyword)  # 디버깅: 선택된 키워드 출력

            title = f"{selected_keyword} 준비 팁"
            content = generate_content(selected_keyword)
            image_url = fetch_image(selected_keyword)

            file_path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", 
                filetypes=[("Excel files", "*.xlsx")]
            )
            if file_path:
                save_to_excel(title, selected_keyword, image_url, content, file_path)
                messagebox.showinfo("완료", "블로그 콘텐츠가 성공적으로 저장되었습니다!")
        except Exception as e:
            messagebox.showerror("오류", f"작업 중 오류 발생: {e}")

    root = tk.Tk()
    root.title("블로그 콘텐츠 생성기")

    tk.Label(root, text="블로그 주제 입력:").pack(pady=5)
    topic_entry = tk.Entry(root, width=40)
    topic_entry.pack(pady=5)

    generate_button = tk.Button(root, text="콘텐츠 생성", command=generate_blog)
    generate_button.pack(pady=20)

    test_button = tk.Button(root, text="OpenAI 테스트", command=test_openai)
    test_button.pack(pady=5)

    root.mainloop()

if __name__ == "__main__":
    create_ui()
