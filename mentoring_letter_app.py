# main.py
# Smilegate HRD | Mentoring Letter Auto Generator (Streamlit App)
# 실행: streamlit run main.py

import io
from datetime import date

import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE
from pptx.dml.color import RGBColor

APP_TITLE = "멘토링 Letter 자동 생성기"
FIRST_SENTENCE_TEMPLATE = "{mentor} 멘토님, {mentee} 멘티의 멘토링 지원을 잘 부탁드립니다."

# 기본 요청사항
DEFAULT_REQUEST_TEXT = """1) 조직, 회사에 대한 이해
  - 조직의 방향성 및 구성에 대한 빠른 학습
  - 안정적으로 팀 문화에 적응할 수 있도록 도와주세요.
  - 업무적으로 편안하게 질문 할 수 있는 관계 형성이 되면 좋겠습니다.

2) 성장 및 업무 관련 지원
  - 팀 업무를 위해 사용 필요한 각종 시스템 및 프로세스에 대해 알려주세요.
  - 앞으로 맡아서 진행할 프로젝트 내 역할 분담"""

# 기본 멘토 활동 후기 설명
DEFAULT_MENTOR_NOTE = """▶ 리더 요청 사항 기반 활동한 내용을 간단하게 작성해주세요
▶ 추가적으로 조직장이 F/U이 필요한 사항을 작성해주세요.
   (ex 멘토링 활동간 멘티 궁금해 했으나, 답변을 못한 부분 or 요청한 사항)"""

THEME_COLOR = "#0B2B4C"  # 네이비 톤
RIGHT_BG = (237, 233, 226)
FONT_NAME = "Malgun Gothic"


def _add_textbox(slide, left_in, top_in, width_in, height_in, title, body,
                 font_size_title=28, font_size_body=18, bold_title=True):
    left = Inches(left_in)
    top = Inches(top_in)
    width = Inches(width_in)
    height = Inches(height_in)
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.word_wrap = True

    # 제목
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.size = Pt(font_size_title)
    run.font.bold = bold_title
    run.font.name = FONT_NAME

    p = tf.add_paragraph()
    p.text = ""
    p.space_after = Pt(4)

    # 본문
    for line in (body or "").splitlines():
        p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size_body)
        p.font.name = FONT_NAME
    return shape


def _add_rect(slide, left_in, top_in, width_in, height_in, fill_rgb=None,
              line_rgb=(180, 180, 180), line_width_pt=1.25):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left_in), Inches(top_in),
        Inches(width_in), Inches(height_in)
    )
    if fill_rgb:
        shape.fill.solid()
        shape.fill.fore_color.rgb = RGBColor(*fill_rgb)
    else:
        shape.fill.background()
    if line_rgb:
        shape.line.color.rgb = RGBColor(*line_rgb)
        shape.line.width = Pt(line_width_pt)
    else:
        shape.line.fill.background()
    return shape


def build_ppt(mentor, mentee, manager, first_sentence_template, request_text,
              use_default_request, qna_text, hide_qna_if_empty, mentor_note_text,
              logo_bytes, theme_color_hex):

    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 테두리
    _add_rect(slide, 0.3, 0.3, 12.7, 6.9, None, (60, 60, 60), 1.5)

    # 로고
    if logo_bytes:
        slide.shapes.add_picture(io.BytesIO(logo_bytes), Inches(0.5), Inches(0.5), height=Inches(0.55))

    # 상단 문구
    header = slide.shapes.add_textbox(Inches(1.0), Inches(0.5), Inches(11.4), Inches(0.6))
    tf = header.text_frame
    p = tf.paragraphs[0]
    r = p.add_run()
    r.text = "멘토링 Letter는 멘토/멘티가 유의미한 멘토링이 되도록 참고할 수 있는 내용을 리더가 멘토에게 보내는 메시지 입니다."
    r.font.size = Pt(16)
    r.font.bold = True
    r.font.name = FONT_NAME

    # 섹션 제목
    for (text, x) in [("멘토에게", 1.0), ("활동 후기", 7.2)]:
        box = slide.shapes.add_textbox(Inches(x), Inches(1.1), Inches(6.0), Inches(0.5))
        tfb = box.text_frame
        r = tfb.paragraphs[0].add_run()
        r.text = text
        r.font.size = Pt(24)
        r.font.bold = True
        r.font.name = FONT_NAME

    # 첫 문장
    sentence = first_sentence_template.format(mentor=mentor.strip(), mentee=mentee.strip())
    box = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(11.4), Inches(0.6))
    tf = box.text_frame
    r = tf.paragraphs[0].add_run()
    r.text = sentence
    r.font.size = Pt(18)
    r.font.name = FONT_NAME

    left_x, top_y = 1.0, 2.1
    col_w, col_h = 6.0, 4.9
    _add_rect(slide, left_x + col_w + 0.2, top_y, col_w, col_h, RIGHT_BG, (180, 180, 180), 1.25)

    # 좌측: 요청사항
    req = (request_text or "").strip()
    if use_default_request or len(req) < 5:
        req = DEFAULT_REQUEST_TEXT
    _add_textbox(slide, left_x, top_y, col_w, 2.6, "조직장 요청사항", req)

    # 좌측: 질문·고민
    if not (hide_qna_if_empty and not qna_text.strip()):
        qna = qna_text.strip() or "(멘티 작성 예정)"
        _add_textbox(slide, left_x, top_y + 2.7, col_w, 2.3, "멘티 질문·고민", qna)

    # 우측: 활동 후기
    _add_textbox(slide, left_x + col_w + 0.25, top_y + 0.15,
                 col_w - 0.5, col_h - 0.3, "멘토 활동 후기", mentor_note_text)

    # 푸터
    footer = slide.shapes.add_textbox(Inches(0.7), Inches(7.1), Inches(12.0), Inches(0.3))
    tf = footer.text_frame
    p = tf.paragraphs[0]
    p.alignment = PP_ALIGN.RIGHT
    r = p.add_run()
    r.text = f"Mentor: {mentor}  |  Mentee: {mentee}  |  Date: {date.today():%Y.%m.%d}"
    r.font.size = Pt(12)
    r.font.name = FONT_NAME

    bio = io.BytesIO()
    prs.save(bio)
    bio.seek(0)
    return bio


def ui():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧡", layout="wide")
    st.title(APP_TITLE)

    with st.sidebar:
        st.header("브랜딩 설정")
        theme = st.color_picker("포인트 색상", THEME_COLOR)
        logo_file = st.file_uploader("로고 업로드 (PNG 권장)", type=["png", "jpg", "jpeg"])
        st.caption("폰트는 시스템의 'Malgun Gothic'을 사용합니다.")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("인적 정보")
        mentor = st.text_input("멘토 이름")
        mentee = st.text_input("멘티 이름")
        manager = st.text_input("조직장(선택)")
        first_sentence_template = st.text_input("첫 문장 템플릿", value=FIRST_SENTENCE_TEMPLATE)

        st.subheader("조직장 요청사항")
        request_text = st.text_area("요청사항 입력", height=200)
        use_default_request = st.checkbox("비어있거나 짧으면 기본 양식 사용", value=True)

        st.subheader("멘티 질문·고민")
        qna_text = st.text_area("질문·고민 입력", height=140)
        hide_qna_if_empty = st.checkbox("질문·고민이 없으면 해당 영역 삭제", value=True)

    with col2:
        st.subheader("멘토 활동 후기")
        mentor_note_text = st.text_area("후기 가이드", value=DEFAULT_MENTOR_NOTE, height=260)

        if mentor and mentee:
            st.markdown(f"**미리보기:** {first_sentence_template.format(mentor=mentor, mentee=mentee)}")
        else:
            st.caption("멘토/멘티 이름을 입력하면 첫 문장을 미리볼 수 있어요.")

    if st.button("PPT 생성 (다운로드)"):
        if not mentor or not mentee:
            st.error("멘토/멘티 이름은 필수입니다.")
            return
        logo_bytes = logo_file.read() if logo_file else None
        ppt_bytes = build_ppt(mentor, mentee, manager, first_sentence_template,
                              request_text, use_default_request, qna_text,
                              hide_qna_if_empty, mentor_note_text, logo_bytes, theme)
        st.download_button(
            "PPT 다운로드",
            ppt_bytes,
            f"Mentoring_Letter_{mentee}_{mentor}.pptx",
            "application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )


if __name__ == "__main__":
    ui()
