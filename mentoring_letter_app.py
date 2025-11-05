# mentoring_letter_app.py
# Streamlit app to automate mentoring letter PPT creation (Smilegate HRD Mentoring Letter)

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

# --- 요청사항 기본 양식 (삼중따옴표: \n 불필요) ---
DEFAULT_REQUEST_TEXT = """1) 조직, 회사에 대한 이해
  - 조직의 방향성 및 구성에 대한 빠른 학습
  - 안정적으로 팀 문화에 적응할 수 있도록 도와주세요.
  - 업무적으로 편안하게 질문 할 수 있는 관계 형성이 되면 좋겠습니다.

2) 성장 및 업무 관련 지원
  - 팀 업무를 위해 사용 필요한 각종 시스템 및 프로세스에 대해 알려주세요.
  - 앞으로 맡아서 진행할 프로젝트 내 역할 분담"""

# --- 멘토 활동 후기 가이드 (삼중따옴표) ---
DEFAULT_MENTOR_NOTE = """▶ 리더 요청 사항 기반 활동한 내용을 간단하게 작성해주세요
▶ 추가적으로 조직장이 F/U이 필요한 사항을 작성해주세요.
   (ex 멘토링 활동간 멘티 궁금해 했으나, 답변을 못한 부분 or 요청한 사항)"""

THEME_COLOR = "#0B2B4C"         # 네이비 톤(시안 느낌)
RIGHT_BG = (237, 233, 226)      # 우측 카드 배경
FONT_NAME = "Malgun Gothic"     # 배포 환경 폰트 설치 필요


def _add_textbox(slide, left_in, top_in, width_in, height_in,
                 title, body, font_size_title=28, font_size_body=18, bold_title=True):
    left = Inches(left_in)
    top = Inches(top_in)
    width = Inches(width_in)
    height = Inches(height_in)
    shape = slide.shapes.add_textbox(left, top, width, height)
    tf = shape.text_frame
    tf.word_wrap = True

    # Title
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = title
    run.font.size = Pt(font_size_title)
    run.font.bold = bold_title
    run.font.name = FONT_NAME

    # Spacer
    p = tf.add_paragraph()
    p.text = ""
    p.space_after = Pt(4)

    # Body (에디터/OS 상관없이 줄 나눔 처리)
    for line in (body or "").splitlines():
        p = tf.add_paragraph()
        p.text = line
        p.font.size = Pt(font_size_body)
        p.font.name = FONT_NAME
    return shape


def _add_rect(slide, left_in, top_in, width_in, height_in,
              fill_rgb=None, line_rgb=(180, 180, 180), line_width_pt=1.25):
    shape = slide.shapes.add_shape(
        MSO_SHAPE.RECTANGLE,
        Inches(left_in), Inches(top_in), Inches(width_in), Inches(height_in)
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


def build_ppt(
    mentor: str,
    mentee: str,
    manager: str | None,
    first_sentence_template: str,
    request_text: str | None,
    use_default_request: bool,
    qna_text: str | None,
    hide_qna_if_empty: bool,
    mentor_note_text: str,
    logo_bytes: bytes | None,
    theme_color_hex: str,
):
    prs = Presentation()
    blank_layout = prs.slide_layouts[6]
    slide = prs.slides.add_slide(blank_layout)

    # 전체 테두리
    _add_rect(slide, 0.3, 0.3, 12.7, 6.9, fill_rgb=None, line_rgb=(60, 60, 60), line_width_pt=1.5)

    # 로고
    if logo_bytes is not None:
        slide.shapes.add_picture(io.BytesIO(logo_bytes), Inches(0.5), Inches(0.5), height=Inches(0.55))

    # 상단 설명 헤더
    header = slide.shapes.add_textbox(Inches(1.0), Inches(0.5), Inches(11.4), Inches(0.6))
    tf = header.text_frame
    tf.clear()
    p = tf.paragraphs[0]
    run = p.add_run()
    run.text = "멘토링 Letter는 멘토/멘티가 유의미한 멘토링이 되도록 참고할 수 있는 내용을 리더가 멘토에게 보내는 메시지 입니다."
    run.font.size = Pt(16)
    run.font.bold = True
    run.font.name = FONT_NAME

    # 섹션 제목
    lt = slide.shapes.add_textbox(Inches(1.0), Inches(1.1), Inches(6.0), Inches(0.5))
    ltf = lt.text_frame; ltf.clear()
    lrun = ltf.paragraphs[0].add_run()
    lrun.text = "멘토에게"
    lrun.font.size = Pt(24); lrun.font.bold = True; lrun.font.name = FONT_NAME

    rt = slide.shapes.add_textbox(Inches(7.2), Inches(1.1), Inches(5.0), Inches(0.5))
    rtf = rt.text_frame; rtf.clear()
    rrun = rtf.paragraphs[0].add_run()
    rrun.text = "활동 후기"
    rrun.font.size = Pt(24); rrun.font.bold = True; rrun.font.name = FONT_NAME

    # 첫 문장
    first_sentence = first_sentence_template.format(mentor=mentor.strip(), mentee=mentee.strip())
    sub = slide.shapes.add_textbox(Inches(1.0), Inches(1.6), Inches(11.4), Inches(0.6))
    tf2 = sub.text_frame
    p2 = tf2.paragraphs[0]
    run2 = p2.add_run()
    run2.text = first_sentence
    run2.font.size = Pt(18)
    run2.font.name = FONT_NAME

    # 좌/우 영역
    left_x, top_y = 1.0, 2.1
    col_w, col_h = 6.0, 4.9

    # 우측 카드 배경
    _add_rect(slide, left_x + col_w + 0.2, top_y, col_w, col_h,
              fill_rgb=RIGHT_BG, line_rgb=(180, 180, 180), line_width_pt=1.25)

    # 좌측: 요청사항
    req_body = (request_text or "").strip()
    if use_default_request or len(req_body) < 5:
        req_body = DEFAULT_REQUEST_TEXT
    _add_textbox(slide, left_in=left_x, top_in=top_y, width_in=col_w, height_in=2.6,
                 title="조직장 요청사항", body=req_body)

    # 좌측: 질문·고민
    if not (hide_qna_if_empty and (not qna_text or len(qna_text.strip()) == 0)):
        qna_body = (qna_text or "").strip() or "(멘티 작성 예정)"
        _add_textbox(slide, left_in=left_x, top_in=top_y + 2.7, width_in=col_w, height_in=2.3,
                     title="멘티 질문·고민", body=qna_body)

    # 우측: 활동 후기 가이드
    _add_textbox(slide, left_in=left_x + col_w + 0.25, top_in=top_y + 0.15,
                 width_in=col_w - 0.5, height_in=col_h - 0.3,
                 title="멘토 활동 후기", body=mentor_note_text)

    # 푸터
    footer = slide.shapes.add_textbox(Inches(0.7), Inches(7.1), Inches(12.0), Inches(0.3))
    tf3 = footer.text_frame; tf3.clear()
    p3 = tf3.paragraphs[0]; p3.alignment = PP_ALIGN.RIGHT
    r3 = p3.add_run()
    today = date.today().strftime("%Y.%m.%d")
    r3.text = f"Mentor: {mentor}  |  Mentee: {mentee}  |  Date: {today}"
    r3.font.size = Pt(12); r3.font.name = FONT_NAME

    bio = io.BytesIO()
    prs.save(bio); bio.seek(0)
    return bio


def ui():
    st.set_page_config(page_title=APP_TITLE, page_icon="🧡", layout="wide")
    st.title(APP_TITLE)

    with st.sidebar:
        st.header("브랜딩 설정")
        theme = st.color_picker("포인트 색상", THEME_COLOR)
        logo_file = st.file_uploader("로고 업로드 (PNG 권장)", type=["png", "jpg", "jpeg"])
        st.markdown("—")
        st.caption("폰트는 시스템의 'Malgun Gothic'을 사용합니다. 배포 환경의 폰트 설치를 확인하세요.")

    col1, col2 = st.columns(2)
    with col1:
        st.subheader("인적 정보")
        mentor = st.text_input("멘토 이름", placeholder="홍길동")
        mentee = st.text_input("멘티 이름", placeholder="김스마일")
        manager = st.text_input("조직장(선택)")
        first_sentence_template = st.text_input(
            "첫 문장 템플릿",
            value=FIRST_SENTENCE_TEMPLATE,
            help="{mentor}, {mentee} 플레이스홀더 사용"
        )

        st.subheader("조직장 요청사항")
        request_text = st.text_area("요청사항 입력 (비워두면 기본 양식 사용)", height=200, value="")
        use_default_request = st.checkbox("비어있거나 짧으면 기본 양식 자동 적용", value=True)

        st.subheader("멘티 질문·고민")
        qna_text = st.text_area("질문·고민 입력", height=140)
        hide_qna_if_empty = st.checkbox("질문·고민이 없으면 해당 영역 삭제", value=True)

    with col2:
        st.subheader("멘토 활동 후기")
        mentor_note_text = st.text_area("후기 가이드/초안", value=DEFAULT_MENTOR_NOTE, height=260)
        st.info("우측 영역은 멘토가 활동 종료 후 작성합니다. 가이드 문구를 커스터마이즈 할 수 있어요.")

        st.subheader("미리보기")
        if mentor and mentee:
            preview_first = first_sentence_template.format(mentor=mentor, mentee=mentee)
            st.markdown(f"**첫 문장:** {preview_first}")
        else:
            st.caption("멘토/멘티 이름을 입력하면 첫 문장을 미리볼 수 있어요.")

    st.markdown("—")
    if st.button("PPT 생성 (다운로드)"):
        if not mentor or not mentee:
            st.error("멘토/멘티 이름은 필수입니다.")
            return
        logo_bytes = logo_file.read() if logo_file else None
        ppt_bytes = build_ppt(
            mentor=mentor,
            mentee=mentee,
            manager=manager,
            first_sentence_template=first_sentence_template,
            request_text=request_text,
            use_default_request=use_default_request,
            qna_text=qna_text,
            hide_qna_if_empty=hide_qna_if_empty,
            mentor_note_text=mentor_note_text,
            logo_bytes=logo_bytes,
            theme_color_hex=theme,
        )
        st.download_button(
            label="PPT 다운로드",
            data=ppt_bytes,
            file_name=f"Mentoring_Letter_{mentee}_{mentor}.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )


if __name__ == "__main__":
    ui()
