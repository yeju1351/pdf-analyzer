import streamlit as st
import pdfplumber
import google.generativeai as genai
from docx import Document
import io
import olefile
import zlib

# 🚨 본인의 API 키를 꼭 다시 넣어주세요!
API_KEY = "AIzaSyDmK3GTgvmA2cupO-FVI1MoUv8wfQR52Cs"
genai.configure(api_key=API_KEY)

# 1. PDF 텍스트 추출 함수
def get_text_from_pdf(file):
    all_text = ""
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            extracted = page.extract_text()
            if extracted:
                all_text += extracted + "\n"
    return all_text

# 2. 한글(HWP) 텍스트 추출 함수 (새로 추가됨)
def get_text_from_hwp(file):
    try:
        file_bytes = file.read()
        ole = olefile.OleFileIO(file_bytes)
        
        dirs = ole.listdir()
        valid_dirs = [d for d in dirs if d[0] == "BodyText"]
        
        text = ""
        for d in valid_dirs:
            stream = ole.openstream(d)
            data = stream.read()
            try:
                # HWP 내부의 압축된 텍스트를 해독합니다.
                decoded = zlib.decompress(data, -15)
                text += decoded.decode('utf-16le', errors='ignore')
            except Exception:
                pass
        
        # 파일 포인터를 다시 처음으로 되돌려 놓습니다.
        file.seek(0)
        return text
    except Exception as e:
        return f"\n[HWP 추출 오류: {e}]\n"

# 3. 파일 종류에 따라 알아서 추출해주는 통합 함수
def get_text_from_file(file):
    file_name = file.name.lower()
    if file_name.endswith('.pdf'):
        return get_text_from_pdf(file)
    elif file_name.endswith('.hwp'):
        return get_text_from_hwp(file)
    else:
        return ""

def analyze_text_with_ai(text):
    model = genai.GenerativeModel('gemini-2.5-flash')
    
    # 💡 이 부분의 지시사항을 엑셀 표 스타일로 완벽하게 바꿨습니다!
    prompt = f"""
    다음은 사업 공고문, 세부평가기준, 작성안내서 등의 전체 문서 내용입니다. 
    이 내용을 종합적으로 분석하여, 아래의 [출력 양식]과 완벽하게 동일한 형태의 표(마크다운 Table)로 정리해 주세요.
    HWP 파일 텍스트 추출 과정에서 이상한 기호가 섞여 있더라도 무시하고 전체 문맥을 파악해 주세요.

    [분석 지시사항]
    1. 문서를 꼼꼼히 읽고 '실격, 평가제외, 감점, 최하점, 불가, 탈락, 주의사항'에 해당하는 조건들을 모두 찾아내세요.
    2. 찾은 조건들을 '구분, 조건유형, 세부 내용, 출처' 4개의 열로 분류하여 표 형식으로 작성하세요.
    3. '출처'에는 해당 내용이 어느 문서의 어느 부분(예: 공고문 p.2, 세부평가기준 Ⅳ.6)에 있는지 최대한 정확히 기재하세요.

    [출력 양식 예시]
    | 구분 | 조건유형 | 세부 내용 | 출처 |
    | :--- | :--- | :--- | :--- |
    | 공동도급 | 주의사항 | 자격1) + 자격2) 업체의 공동도급 형태로만 참여 가능(단독 참여 불가) | 공고문 p.1 |
    | 책임기술인 | 감점 | A4로 작성해야 하며, A3로 작성한 경우 2Page로 계산. 초과 시 0.5점 감점 | 작성안내서 붙임1 |

    [전체 문서 내용]
    {text}
    """
    response = model.generate_content(prompt)
    return response.text

def create_word_file(result_text):
    doc = Document()
    doc.add_heading('📄 AI 사업 공고 통합 분석 보고서', 0)
    doc.add_paragraph(result_text)
    bio = io.BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- 웹 화면(UI) ---
st.set_page_config(page_title="통합 문서 분석기", page_icon="📚")

st.title("📚 AI 통합 문서 분석기")
st.write("공고문, 세부지침 등 쪼개져 있는 **PDF와 한글(HWP)** 파일들을 한 번에 올리면 종합 분석해 드립니다!")

# 🌟 핵심 변경: type=["pdf", "hwp"] 로 설정하여 두 가지 파일 모두 허용
uploaded_files = st.file_uploader("관련 파일들을 모두 드래그해서 올려주세요 (PDF, HWP 동시 업로드 가능)", type=["pdf", "hwp"], accept_multiple_files=True)

if uploaded_files:
    if st.button("🚀 통합 분석 시작"):
        
        combined_text = ""
        
        with st.spinner('AI가 PDF와 한글(HWP) 문서들을 하나로 취합하여 꼼꼼히 읽고 있습니다... 🧠'):
            
            # 올려준 파일들을 하나씩 확인하며 텍스트를 합칩니다.
            for file in uploaded_files:
                combined_text += f"\n\n--- [{file.name} 문서 내용] ---\n\n"
                combined_text += get_text_from_file(file)
            
            result = analyze_text_with_ai(combined_text)
            
            st.success("✅ 문서 통합 분석이 완료되었습니다!")
            st.markdown(result)
            
            st.divider()
            st.subheader("💾 통합 분석 보고서 저장하기")
            col1, col2 = st.columns(2)
            
            with col1:
                st.download_button(
                    label="📝 텍스트(.txt)로 저장",
                    data=result,
                    file_name="사업공고_통합분석.txt",
                    mime="text/plain"
                )
            with col2:
                word_data = create_word_file(result)
                st.download_button(
                    label="📄 워드(.docx)로 저장",
                    data=word_data,
                    file_name="사업공고_통합분석.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                )