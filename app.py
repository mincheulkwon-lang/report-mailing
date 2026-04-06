import streamlit as st
import tempfile
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from core.report import build_html_from_file

st.set_page_config(page_title="리포트 대시보드", layout="wide")

st.title("📊 리포트 메일링 대시보드")

# =========================
# secrets에서 발송 계정 불러오기
# =========================
sender_email = st.secrets["EMAIL"]
sender_pw = st.secrets["PASSWORD"]

# =========================
# 사이드바 설정
# =========================
with st.sidebar:
    st.header("📧 메일 설정")
    st.caption(f"발송 계정: {sender_email}")
    receiver_email = st.text_input("받는 이메일 (여러 명이면 쉼표로 구분)")

# =========================
# 파일 업로드
# =========================
uploaded_files = st.file_uploader(
    "엑셀 파일 업로드",
    type=["xlsx", "xlsm"],
    accept_multiple_files=True
)

# =========================
# 리포트 생성
# =========================
if st.button("리포트 생성"):
    results = []

    if not uploaded_files:
        st.warning("엑셀 파일을 먼저 업로드해주세요.")
    else:
        for f in uploaded_files:
            try:
                with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
                    tmp.write(f.read())
                    tmp_path = tmp.name

                html, display_tables, comment_source_tables = build_html_from_file(tmp_path)

                results.append({
                    "file_name": f.name,
                    "html": html,
                    "display_tables": display_tables,
                    "comment_source_tables": comment_source_tables
                })

            except Exception as e:
                st.error(f"{f.name} 처리 중 오류 발생: {e}")

        st.session_state["results"] = results

# =========================
# 리포트 미리보기
# =========================
if "results" in st.session_state and st.session_state["results"]:
    st.divider()
    st.subheader("리포트 미리보기")

    for r in st.session_state["results"]:
        st.markdown(f"### 📁 {r['file_name']}")
        st.components.v1.html(r["html"], height=800, scrolling=True)

    # =========================
    # 메일 발송
    # =========================
    if st.button("📧 메일 발송"):
        if not receiver_email.strip():
            st.error("받는 이메일을 입력해주세요.")
        else:
            try:
                server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
                server.login(sender_email, sender_pw)

                for r in st.session_state["results"]:
                    msg = MIMEMultipart()
                    msg["Subject"] = r["file_name"]
                    msg["From"] = sender_email
                    msg["To"] = receiver_email

                    msg.attach(MIMEText(r["html"], "html"))
                    server.send_message(msg)

                server.quit()
                st.success("메일 발송 완료")

            except Exception as e:
                st.error(f"메일 발송 중 오류 발생: {e}")