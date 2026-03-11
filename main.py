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
    # HWP 파일에서 추출 시 기호가 섞일 수 있어 AI에게 무시하라는 지시를 추가했습니다.
    prompt = f"""
    다음은 특정 사업에 대한 공고문 및 관련 지침서들의 전체 내용을 하나로 합친 것입니다. 
    이 내용을 종합적으로 분석해서 아래 양식에 맞게 요약하고 정리해 주세요.
    HWP 파일 텍스트 추출 과정에서 이상한 기호가 섞여 있더라도 무시하고 전체 문맥을 파악해 주세요.
    
    [분석 양식]
    1. 기본 정보: 사업명, 주관 기관, 제출 마감 일시
    2. 지원 자격: 신청 가능 대상 및 필수 조건
    3. 제한 및 결격 사유 (특이사항 1): 지원할 수 없는 대상이나 조건
    4. 제출 및 작성 지침 (특이사항 2): 필수 제출 서류 목록, 분량 제한 등
    5. 평가 기준: 주요 평가 항목 및 배점
    6. 감점 요인 및 탈락 기준 (핵심): 감점되거나 탈락되는 경우

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