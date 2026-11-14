"""AI Podcast Generator - Streamlit user interface.  Run with:  streamlit run app.py"""
from __future__ import annotations

import hashlib
import hmac
import os
import traceback

import streamlit as st

import podcast_core as core

st.set_page_config(page_title="AI Podcast Generator", page_icon="🎙️", layout="centered")

CSS = """
<style>
.block-container {max-width: 820px; padding-top: 2rem;}
.app-title {font-size: 2.4rem; font-weight: 800; margin-bottom: 0;}
.app-subtitle {color: #6b7280; margin-top: 0.2rem; margin-bottom: 1.2rem; font-size: 1.05rem;}
.step-label {font-weight: 700; font-size: 1.15rem; margin: 1.4rem 0 0.2rem 0;}
</style>
"""


def md_escape(text: str) -> str:
    """Stop Streamlit's markdown from treating $, *, _ etc. in script text as formatting."""
    for ch in ("\\", "$", "*", "_", "`", "~", "#"):
        text = text.replace(ch, "\\" + ch)
    return text


def render_header() -> None:
    st.markdown(CSS, unsafe_allow_html=True)
    st.markdown('<p class="app-title">🎙️ AI Podcast Generator</p>', unsafe_allow_html=True)
    st.markdown(
        '<p class="app-subtitle">Turn any document into a two-host podcast you can play and download.</p>',
        unsafe_allow_html=True,
    )
    if not core.has_api_key():
        st.error(
            "**OpenAI API key not found.** Running locally: create a file named `.env` next to `app.py` "
            "containing `OPENAI_API_KEY=your_key_here`, then restart. Running on Streamlit Cloud: add "
            "`OPENAI_API_KEY` under the app's Settings → Secrets. (See README.md / DEPLOY.md.)"
        )


def render_upload():
    st.markdown('<div class="step-label">1 · Upload your document</div>', unsafe_allow_html=True)
    uploaded = st.file_uploader(
        "Choose a PDF, Word (.docx) or text (.txt) file",
        type=["pdf", "docx", "txt"],
        help=f"Up to {core.MAX_FILE_MB} MB. Scanned (image-only) PDFs are not supported.",
    )
    if uploaded is None:
        st.session_state.pop("doc_cache", None)
        st.session_state.pop("result", None)
        return None

    data = uploaded.getvalue()
    key = hashlib.md5(data).hexdigest() + "|" + uploaded.name
    cache = st.session_state.get("doc_cache")
    if cache and cache["key"] == key:
        doc = cache["doc"]
    else:
        try:
            doc = core.load_document(uploaded.name, data)
        except core.PodcastError as exc:
            st.session_state.pop("doc_cache", None)
            st.error(str(exc))
            return None
        except Exception as exc:  # never show a raw traceback for a bad upload
            st.session_state.pop("doc_cache", None)
            st.error(f"Sorry, that file could not be read ({exc}).")
            return None
        st.session_state["doc_cache"] = {"key": key, "doc": doc}
        st.session_state.pop("result", None)  # a new document invalidates the old podcast

    st.success(f"**{doc.filename}** uploaded successfully.")
    c1, c2, c3 = st.columns(3)
    c1.metric("File type", doc.kind.upper())
    c2.metric("Words", f"{doc.word_count:,}")
    c3.metric("Characters", f"{doc.char_count:,}")
    st.caption("Preview of the extracted text")
    st.text_area("Preview", doc.preview(), height=170, disabled=True, label_visibility="collapsed")
    if doc.char_count > core.DIRECT_LIMIT:
        st.info("This is a long document. Each section will be summarised first, so generation may take a little longer.")
    return doc


def render_settings():
    st.markdown('<div class="step-label">2 · Customise your podcast</div>', unsafe_allow_html=True)
    c1, c2 = st.columns(2)
    with c1:
        length = st.radio("Podcast length", list(core.LENGTHS.keys()), index=0)
    with c2:
        tone = st.radio("Tone", list(core.TONES.keys()), index=0,
                        help="Conversational is casual and friendly. Educational is structured and ends with a recap.")
    return length, tone


def run_generation(doc, length: str, tone: str) -> None:
    bar = st.progress(0.0, text="Starting…")

    def on_progress(fraction: float, message: str) -> None:
        bar.progress(min(max(fraction, 0.0), 1.0), text=message)

    try:
        client = core.make_client()
        result = core.generate_podcast(client, doc.text, length, tone, progress=on_progress)
    except core.PodcastError as exc:
        bar.empty()
        st.error(str(exc))
        return
    except Exception as exc:  # unexpected bug: show something useful but friendly
        bar.empty()
        st.error("Something unexpected went wrong while generating the podcast. Please try again.")
        with st.expander("Technical details"):
            st.code("".join(traceback.format_exception(exc)))
        return
    bar.empty()
    st.session_state["result"] = {"result": result, "stem": core.safe_filename(doc.filename)}


def render_result() -> None:
    saved = st.session_state.get("result")
    if not saved:
        return
    result, stem = saved["result"], saved["stem"]

    st.markdown('<div class="step-label">3 · Your podcast</div>', unsafe_allow_html=True)
    st.success(f"**{result.script.title}** — about {result.duration_seconds / 60:.1f} minutes")
    for w in result.warnings:
        st.warning(w)

    st.audio(result.wav_bytes, format="audio/wav")

    cols = st.columns(3 if result.mp3_bytes else 2)
    cols[0].download_button("⬇️ Download Podcast (WAV)", data=result.wav_bytes,
                            file_name=f"{stem}_podcast.wav", mime="audio/wav", key="dl_wav")
    idx = 1
    if result.mp3_bytes:
        cols[1].download_button("⬇️ Download Podcast (MP3)", data=result.mp3_bytes,
                                file_name=f"{stem}_podcast.mp3", mime="audio/mpeg", key="dl_mp3")
        idx = 2
    cols[idx].download_button("📄 Download Script", data=result.script.to_text(),
                              file_name=f"{stem}_script.txt", mime="text/plain", key="dl_script")

    st.subheader("Podcast script")
    with st.container(border=True):
        for turn in result.script.turns:
            st.markdown(f"**{turn.host.name}:** {md_escape(turn.text)}")


def check_password() -> bool:
    """Optional access gate: only active when an APP_PASSWORD secret / environment variable is set."""
    expected = os.getenv("APP_PASSWORD")
    if not expected or st.session_state.get("authed"):
        return True
    st.info("This app is password protected.")
    entered = st.text_input("Password", type="password")
    if entered:
        if hmac.compare_digest(entered.encode(), expected.encode()):
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def main() -> None:
    render_header()
    if not check_password():
        return
    doc = render_upload()
    length, tone = render_settings()

    st.markdown('<div class="step-label">&nbsp;</div>', unsafe_allow_html=True)
    can_run = doc is not None and core.has_api_key()
    if st.button("🎧 Generate Podcast", type="primary", disabled=not can_run):
        run_generation(doc, length, tone)
    if doc is None:
        st.caption("Upload a document to enable the Generate Podcast button.")

    render_result()

    with st.expander("How does this work?"):
        st.markdown(
            "1. Your document's text is extracted.\n"
            "2. If it's long, each section is summarised so nothing important is missed.\n"
            "3. An AI language model writes a conversation between two hosts, sticking to the document.\n"
            "4. An AI text-to-speech model gives each host a different voice.\n"
            "5. The voices are joined into one audio file you can play and download.\n\n"
            "*Your document text is sent to OpenAI to create the podcast, and each run uses a small amount of API credit.*"
        )


main()
