import hmac
import io
import os
import time

import streamlit as st
from anthropic import APIError, APIStatusError

import problem_proposal as pp

MAX_FILES = 8
MAX_FILE_BYTES = 30 * 1024
MAX_RUNS_PER_SESSION = 5
MAX_LOGIN_FAILS = 5

st.set_page_config(page_title="부서 공통 문제 해결 제안서", page_icon="📑", layout="centered")


def secret(name: str):
    """Streamlit secrets 또는 환경변수에서 값을 읽는다. 없으면 None."""
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:  # secrets.toml이 없는 로컬 실행
        pass
    return os.environ.get(name)


def check_access() -> bool:
    """앱 비밀번호(APP_PASSWORD)를 아는 사람만 사용할 수 있게 한다.

    비밀번호가 설정되지 않았으면 앱을 열지 않는다. (설정 누락으로 API 키가 무방비로 노출되는 것을 막음)
    """
    expected = secret("APP_PASSWORD")
    if not expected:
        st.error("앱 비밀번호(APP_PASSWORD)가 설정되지 않아 사용할 수 없습니다. 관리자에게 문의하세요.")
        return False
    if st.session_state.get("authed"):
        return True

    fails = st.session_state.get("login_fails", 0)
    if fails >= MAX_LOGIN_FAILS:
        st.error("비밀번호를 여러 번 틀렸습니다. 페이지를 새로 고친 뒤 다시 시도하세요.")
        return False

    with st.form("login"):
        code = st.text_input("앱 비밀번호", type="password")
        submitted = st.form_submit_button("들어가기")
    if submitted:
        if hmac.compare_digest(code.encode(), str(expected).encode()):
            st.session_state["authed"] = True
            st.session_state["login_fails"] = 0
            st.rerun()
        else:
            st.session_state["login_fails"] = fails + 1
            time.sleep(1)  # 무차별 대입 지연
            st.error(f"비밀번호가 맞지 않습니다. ({fails + 1}/{MAX_LOGIN_FAILS})")
    return False


def decode_text(raw: bytes) -> str:
    for enc in ("utf-8-sig", "cp949"):
        try:
            return raw.decode(enc)
        except UnicodeDecodeError:
            continue
    raise RuntimeError("텍스트 인코딩을 읽을 수 없습니다. UTF-8 또는 CP949 형식으로 저장해주세요.")


def reports_from_uploads(files):
    if len(files) < 2:
        raise RuntimeError("공통 문제를 찾으려면 보고서가 2개 이상 필요합니다.")
    if len(files) > MAX_FILES:
        raise RuntimeError(f"보고서는 최대 {MAX_FILES}개까지 올릴 수 있습니다.")
    reports = []
    for f in files:
        raw = f.getvalue()
        if len(raw) > MAX_FILE_BYTES:
            raise RuntimeError(f"'{f.name}' 파일이 너무 큽니다. (최대 {MAX_FILE_BYTES // 1024}KB)")
        content = decode_text(raw).strip()
        if not content:
            raise RuntimeError(f"'{f.name}' 파일 내용이 비어 있습니다.")
        reports.append({"filename": os.path.basename(f.name), "content": content})
    names = [r["filename"] for r in reports]
    if len(set(names)) != len(names):
        raise RuntimeError("파일 이름이 같은 보고서가 있습니다. 이름을 다르게 바꿔주세요.")
    return reports


def resolve_api_key() -> str | None:
    key = secret("ANTHROPIC_API_KEY")
    if key:
        return key
    return st.session_state.get("user_key") or None


# ------------------------------------------------------------------ 화면

st.title("부서 공통 문제 해결 제안서")
st.caption("여러 부서의 보고서(.txt)를 올리면 공통 문제를 찾아 해결 방안을 담은 PPT 제안서를 만듭니다.")

if not check_access():
    st.stop()

st.info(
    "올린 보고서 내용은 제안서 작성을 위해 Anthropic API로 전송됩니다. "
    "개인정보나 대외비 자료는 올리지 마세요. 서버에는 저장하지 않습니다."
)

if not secret("ANTHROPIC_API_KEY"):
    st.text_input(
        "Anthropic API 키",
        type="password",
        key="user_key",
        help="서버에 저장하지 않고 이 브라우저 세션에서만 사용합니다.",
    )

files = st.file_uploader(
    f"부서 보고서 (.txt, 2~{MAX_FILES}개, 파일당 {MAX_FILE_BYTES // 1024}KB 이하)",
    type=["txt"],
    accept_multiple_files=True,
)

runs = st.session_state.get("runs", 0)
if st.button("제안서 만들기", type="primary", disabled=not files):
    api_key = resolve_api_key()
    try:
        if not api_key:
            raise RuntimeError("API 키가 설정되어 있지 않습니다.")
        if runs >= MAX_RUNS_PER_SESSION:
            raise RuntimeError(f"한 세션에서는 최대 {MAX_RUNS_PER_SESSION}번까지 만들 수 있습니다. 페이지를 새로 고침해주세요.")
        reports = reports_from_uploads(files)
        with st.spinner("보고서를 분석하고 제안서를 만드는 중입니다. (1~2분)"):
            data = pp.generate_proposal(api_key, reports)
            warnings = pp.check_proposal(data, reports)
            buf = io.BytesIO()
            slide_count = pp.build_pptx(data, reports, buf)
        st.session_state["runs"] = runs + 1
        st.session_state["result"] = {
            "pptx": buf.getvalue(),
            "warnings": warnings,
            "slides": slide_count,
            "data": data,
        }
    except (RuntimeError, APIError, APIStatusError) as e:
        st.session_state.pop("result", None)
        st.error(f"제안서를 만들지 못했습니다: {e}")
    except Exception:
        st.session_state.pop("result", None)
        st.error("예상하지 못한 오류가 발생했습니다. 잠시 후 다시 시도해주세요.")

result = st.session_state.get("result")
if result:
    st.success(f"제안서를 만들었습니다. (슬라이드 {result['slides']}장)")
    st.download_button(
        "PPT 내려받기",
        data=result["pptx"],
        file_name="공통문제_해결제안서.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
    )
    if result["warnings"]:
        st.warning("아래 항목은 사람이 직접 확인해주세요.")
        for w in result["warnings"]:
            st.write(f"- {w}")
    else:
        st.caption("자동 점검: 문제별 근거·인용 일치, 보고서 인용, '제안' 표시 확인 완료.")

    with st.expander("확인된 문제 미리보기"):
        for i, p in enumerate(result["data"]["problems"], 1):
            st.markdown(f"**{i}. {p['title']}** ({', '.join(p['departments'])})")
            for e in p["evidence"]:
                st.markdown(f"- `{e['report']}` — {e['quote']}")
            note = (p.get("note") or "").strip()
            if note:
                st.markdown(f"  - 확인 필요: {note}")

    st.caption("AI가 보고서 원문만 바탕으로 작성한 초안입니다. 인용과 일정은 원문과 대조한 뒤 사용하세요.")
