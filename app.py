import streamlit as st
import os
import time
import wave
import threading
import warnings
import datetime
import smtplib
import asyncio
import sys
import tempfile
import shutil
import re
from email.message import EmailMessage

import numpy as np
import soundfile as sf

# pyaudiowpatch: patched PyAudio with WASAPI loopback support on Windows.
# pip install pyaudiowpatch  (do NOT have plain pyaudio installed alongside it)
import pyaudiowpatch as pyaudio

from playwright.sync_api import sync_playwright
from google import genai
from google.genai import types
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT
from dotenv import load_dotenv

# ── Windows asyncio fix ────────────────────────────────────────────────────────
if sys.platform == "win32":
    asyncio.set_event_loop_policy(asyncio.WindowsProactorEventLoopPolicy())

load_dotenv()
warnings.filterwarnings("ignore")

# ── Config ─────────────────────────────────────────────────────────────────────
st.set_page_config(page_title="Auto MOM Generator", page_icon="📝", layout="centered")

OUTPUT_AUDIO    = "meeting_recording.wav"
OUTPUT_PDF      = "meeting_mom.pdf"
BOT_PROFILE_DIR = "./bot_profile"
GEMINI_MODEL    = "gemini-3-flash-preview"
TODAY_DATE      = datetime.date.today().isoformat()
BROWSER_WIDTH   = 1280
BROWSER_HEIGHT  = 720
SILENCE_RMS_THRESHOLD = 0.0005

# ── Prompts ────────────────────────────────────────────────────────────────────

TRANSCRIPTION_PROMPT = """
You are a transcription engine. Your ONLY job is to convert the audio into a verbatim transcript.

STRICT RULES — NO EXCEPTIONS:
1. Write EXACTLY what is spoken. Word for word.
2. Do NOT summarize, interpret, rephrase, or invent anything.
3. If a word is unclear write [unclear]. If audio is silent write [silence].
4. Include filler words (uh, um, like) exactly as spoken.
5. Label speakers: assign Person1 to the first voice, Person2 to the second, etc.
   - Only use a real name if it is clearly and directly spoken aloud (e.g. "Hi Rahul", "Thanks Priya").
   - Never infer or guess a name from context.
6. Format: one line per speaker turn → "Person1: <exact words>"
7. Output ONLY the transcript lines. No headers, no explanations, nothing else.
"""

MOM_PROMPT = f"""
You are a Minutes of Meeting (MOM) writer.
You receive a verbatim transcript. Extract and structure ONLY what is in the transcript.

ABSOLUTE RULES:
1. NEVER invent, guess, or add anything not present word-for-word in the transcript.
2. If a section has no content in the transcript → write exactly: Not discussed
3. If no action items → write exactly: None recorded
4. If no deadline spoken → write: not specified
5. Use speaker labels from the transcript (Person1, Person2, or spoken names) exactly as they appear.
6. Do NOT rewrite or paraphrase action items — copy them from the transcript.
7. Plain text only. No Markdown. No asterisks. No bold. No bullet symbols other than "-".

DATE: If spoken in transcript use it (YYYY-MM-DD). Otherwise: {TODAY_DATE}
TITLE: Derive from actual topics discussed. If unclear use: Meeting {TODAY_DATE}

Output format — use EXACTLY these section headers, nothing more, nothing less:

Title: <value>
Date: <YYYY-MM-DD>
Participants: <Person1, Person2, ... or spoken names>

Agenda:
- <item from transcript or "Not discussed">

Key Discussion Points:
- <PersonX said: exact quote or close paraphrase — only from transcript>

Decisions:
- <decision explicitly stated in transcript, or "None recorded">

Action Items:
- <PersonX>: <task as stated> (deadline: <date or "not specified">)
"""


# ── WASAPI Loopback Audio Recorder ────────────────────────────────────────────

def _find_loopback_device(p: pyaudio.PyAudio):
    """
    Use pyaudiowpatch's built-in helper to find the default WASAPI loopback device.
    This captures exactly what plays through your speakers — including Google Meet audio.
    """
    try:
        # pyaudiowpatch adds get_default_wasapi_loopback() specifically for this purpose
        device = p.get_default_wasapi_loopback()
        return device
    except Exception as e:
        print(f"[Audio] get_default_wasapi_loopback failed: {e}")

    # Manual fallback: scan all devices for WASAPI loopback flag
    try:
        wasapi_info = p.get_host_api_info_by_type(pyaudio.paWASAPI)
        for i in range(p.get_device_count()):
            dev = p.get_device_info_by_index(i)
            # loopback devices have maxInputChannels > 0 and name contains "Loopback"
            if (dev.get("hostApi") == wasapi_info["index"]
                    and dev.get("maxInputChannels", 0) > 0
                    and "loopback" in dev.get("name", "").lower()):
                return dev
    except Exception as e:
        print(f"[Audio] Manual loopback scan failed: {e}")

    raise RuntimeError(
        "No WASAPI loopback device found.\n"
        "Fix: Right-click the speaker icon → Sounds → Recording tab → "
        "right-click blank area → 'Show Disabled Devices' → Enable 'Stereo Mix'.\n"
        "Then restart the app."
    )


def check_audio_health(wav_path: str):
    """Return (has_speech: bool, rms: float, duration_s: int)."""
    try:
        data, sr = sf.read(wav_path, dtype="float32")
        if data.ndim > 1:
            data = data.mean(axis=1)
        dur = int(len(data) / sr)
        rms = float(np.sqrt(np.mean(data ** 2)))
        return rms > SILENCE_RMS_THRESHOLD and dur > 2, rms, dur
    except Exception:
        return False, 0.0, 0


class AudioRecorder:
    """
    Records system audio using Windows WASAPI loopback via pyaudiowpatch.
    This is the only reliable way to capture Google Meet / browser audio on Windows.
    """

    def __init__(self, filename: str):
        self.filename    = filename
        self._stop       = threading.Event()
        self._thread     = None
        self.error       = None
        self.source_name = "unknown"

    def start(self):
        self._stop.clear()
        self.error = None
        self._thread = threading.Thread(target=self._loop, daemon=True)
        self._thread.start()
        # Give the thread 2 s to initialise; surface early errors fast
        time.sleep(2)
        if self.error:
            raise RuntimeError(f"Audio recorder failed to start: {self.error}")

    def stop(self):
        self._stop.set()
        if self._thread and self._thread.is_alive():
            self._thread.join(timeout=10)

    def _loop(self):
        p = None
        stream = None
        wf = None
        try:
            p = pyaudio.PyAudio()
            device = _find_loopback_device(p)
            self.source_name = device.get("name", "loopback")

            sample_rate = int(device.get("defaultSampleRate", 44100))
            channels    = min(int(device.get("maxInputChannels", 2)), 2)
            chunk       = 1024

            print(f"[AudioRecorder] device='{self.source_name}' "
                  f"rate={sample_rate} ch={channels}")

            wf = wave.open(self.filename, "wb")
            wf.setnchannels(channels)
            wf.setsampwidth(p.get_sample_size(pyaudio.paInt16))
            wf.setframerate(sample_rate)

            stream = p.open(
                format=pyaudio.paInt16,
                channels=channels,
                rate=sample_rate,
                frames_per_buffer=chunk,
                input=True,
                input_device_index=device["index"],
                # WASAPI exclusive/shared mode flag
                input_host_api_specific_stream_info=pyaudio.PaMacCoreStreamInfo()
                    if sys.platform == "darwin" else None,
            )

            while not self._stop.is_set():
                try:
                    data = stream.read(chunk, exception_on_overflow=False)
                    wf.writeframes(data)
                except OSError:
                    pass  # buffer overflow — skip frame, keep going

        except Exception as exc:
            self.error = str(exc)
            print(f"[AudioRecorder] FATAL: {exc}")
        finally:
            if stream:
                try:
                    stream.stop_stream()
                    stream.close()
                except Exception:
                    pass
            if wf:
                try:
                    wf.close()
                except Exception:
                    pass
            if p:
                try:
                    p.terminate()
                except Exception:
                    pass


# ── Meet bot helpers ───────────────────────────────────────────────────────────

_END_SELECTORS = [
    "text=/You were removed/i",
    "text=/You left the meeting/i",
    "text=/Meeting ended/i",
    "text=/Meet ended/i",
    "text=/You were disconnected/i",
    "text=/You've been removed/i",
    "text=/The meeting has ended/i",
    "text=/Left the meeting/i",
    "[jsname='RWcFGd']",
]

_IN_MEETING_SELECTORS = [
    "button[aria-label*='microphone']",
    "button[aria-label*='camera']",
    "button[aria-label*='Leave call']",
    "button[aria-label*='Leave meeting']",
    "[data-participant-id]",
    "[data-is-muted]",
]


def _meeting_has_ended(page) -> bool:
    # 1. Explicit end-screen text/elements
    for sel in _END_SELECTORS:
        try:
            loc = page.locator(sel).first
            if loc.count() and loc.is_visible(timeout=300):
                return True
        except Exception:
            pass
    # 2. URL navigated away from meet
    try:
        url = page.url
        if "meet.google.com" not in url:
            return True
        if re.fullmatch(r"https://meet\.google\.com/?", url):
            return True
    except Exception:
        pass
    # 3. No in-meeting UI visible after page is fully loaded
    in_meeting = False
    for sel in _IN_MEETING_SELECTORS:
        try:
            loc = page.locator(sel).first
            if loc.count() and loc.is_visible(timeout=200):
                in_meeting = True
                break
        except Exception:
            pass
    if not in_meeting:
        try:
            if page.evaluate("document.readyState") == "complete":
                return True
        except Exception:
            pass
    return False


def _get_participant_count(page):
    for sel in ['div[aria-label^="Participants"]',
                'button[aria-label^="Participants"]',
                '[aria-label*="participant"]']:
        try:
            loc = page.locator(sel).first
            if loc.count():
                label = loc.get_attribute("aria-label") or ""
                m = re.search(r"\((\d+)\)", label)
                if m:
                    return int(m.group(1))
        except Exception:
            pass
    return None


# ── Main bot runner ────────────────────────────────────────────────────────────

def run_meet_bot(meet_url: str, status_container) -> bool:
    if not meet_url.startswith("http"):
        meet_url = "https://" + meet_url

    recorder = AudioRecorder(OUTPUT_AUDIO)
    temp_profile_dir = None
    exit_reason = "unknown"

    with sync_playwright() as p:
        status_container.write("🌐 Launching browser...")

        launch_args = [
            "--start-maximized",
            "--use-fake-ui-for-media-stream",
            "--disable-blink-features=AutomationControlled",
            "--disable-infobars",
            "--no-sandbox",
        ]

        def _launch(profile_dir):
            return p.chromium.launch_persistent_context(
                user_data_dir=profile_dir,
                headless=False,
                no_viewport=True,
                args=launch_args,
            )

        try:
            browser = _launch(BOT_PROFILE_DIR)
        except Exception as e:
            status_container.warning(f"Persistent profile locked ({e}). Using temp profile.")
            temp_profile_dir = tempfile.mkdtemp(prefix="bot_profile_")
            browser = _launch(temp_profile_dir)

        page = browser.pages[0] if browser.pages else browser.new_page()

        # Sync viewport to real screen size
        try:
            sz = page.evaluate("()=>({w:window.screen.availWidth,h:window.screen.availHeight})")
            page.set_viewport_size({"width": sz["w"], "height": sz["h"]})
        except Exception:
            pass

        status_container.write("🔗 Navigating to Meet...")
        page.goto(meet_url, wait_until="domcontentloaded", timeout=30_000)
        page.wait_for_timeout(4000)

        # Dismiss cookie/permission popups
        for txt in ["Accept all", "I agree", "Got it", "OK"]:
            try:
                b = page.locator(f"button:has-text('{txt}')").first
                if b.is_visible(timeout=800):
                    b.click()
                    page.wait_for_timeout(400)
            except Exception:
                pass

        # Mute mic and camera before joining
        for key in ["Control+d", "Control+e"]:
            try:
                page.keyboard.press(key)
                page.wait_for_timeout(300)
            except Exception:
                pass

        # Click join button
        joined = False
        for attempt in range(12):
            for label in ["Join now", "Ask to join", "Join meeting", "Join and use", "Join"]:
                try:
                    b = page.locator(f"button:has-text('{label}')").first
                    if b.is_visible(timeout=700):
                        b.click()
                        joined = True
                        break
                except Exception:
                    pass
            if joined:
                break
            page.wait_for_timeout(1500)

        if not joined:
            status_container.warning("⚠️ Join button not found — please click manually.")
        else:
            status_container.write("✅ Joined the meeting.")

        # Wait for meeting UI to stabilise
        page.wait_for_timeout(3000)

        # Start WASAPI loopback recording
        status_container.write("🎙️ Starting audio capture (WASAPI loopback)...")
        try:
            recorder.start()
            status_container.success(
                f"🎙️ Recording started — source: {recorder.source_name}"
            )
        except RuntimeError as e:
            status_container.error(str(e))
            browser.close()
            return False

        # Monitor loop
        grace_until = time.time() + 20
        try:
            while True:
                if time.time() > grace_until:
                    if page.is_closed():
                        exit_reason = "page closed"
                        break
                    if _meeting_has_ended(page):
                        exit_reason = "meeting ended / removed"
                        break
                    cnt = _get_participant_count(page)
                    if cnt is not None and cnt <= 1:
                        exit_reason = f"only {cnt} participant(s) left"
                        break
                time.sleep(0.5)
        finally:
            recorder.stop()

            has_speech, rms, dur = check_audio_health(OUTPUT_AUDIO)
            if recorder.error:
                status_container.error(f"⚠️ Recording error: {recorder.error}")
            elif has_speech:
                status_container.write(
                    f"⏹️ Stopped ({exit_reason}). Audio OK: {dur}s, RMS={rms:.4f} ✅"
                )
            else:
                status_container.warning(
                    f"⚠️ Audio recorded but appears silent (dur={dur}s, RMS={rms:.6f}). "
                    f"Source: {recorder.source_name}. "
                    "Ensure 'Stereo Mix' is enabled in Windows Sound → Recording devices."
                )

            try:
                browser.close()
            except Exception:
                pass
            if temp_profile_dir:
                shutil.rmtree(temp_profile_dir, ignore_errors=True)

    return True


# ── Gemini: 2-step transcribe → structure ─────────────────────────────────────

def _upload_and_wait(client, audio_path: str, status_fn):
    status_fn("📤 Uploading audio to Gemini...")
    uploaded = client.files.upload(file=audio_path)
    poll = 0
    while uploaded.state.name == "PROCESSING":
        poll += 1
        status_fn(f"⏳ Gemini processing audio... ({poll * 3}s)")
        time.sleep(3)
        uploaded = client.files.get(name=uploaded.name)
    if uploaded.state.name == "FAILED":
        raise ValueError("Gemini failed to process the audio file.")
    return uploaded


def generate_mom_from_audio(audio_path: str, api_key: str, status_fn=None):
    if status_fn is None:
        status_fn = print

    # Pre-flight audio check
    has_speech, rms, dur = check_audio_health(audio_path)
    if not has_speech:
        raise ValueError(
            f"Audio file appears silent (duration={dur}s, RMS={rms:.6f}). "
            "Nothing to transcribe. Check your loopback/recording setup."
        )
    status_fn(f"🔊 Audio health OK: {dur}s, RMS={rms:.4f}")

    client = genai.Client(api_key=api_key)
    uploaded = _upload_and_wait(client, audio_path, status_fn)

    # ── Step A: Verbatim transcription ──────────────────────────────────────
    status_fn("🔤 Step A — Transcribing audio verbatim...")
    r1 = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=[
            TRANSCRIPTION_PROMPT,
            uploaded,
            (
                "Transcribe every word spoken. "
                "Label speakers Person1, Person2, etc. "
                "Use a real name ONLY if clearly spoken aloud. "
                "Write [unclear] for inaudible words. "
                "Output ONLY the transcript lines, nothing else."
            ),
        ],
        config=types.GenerateContentConfig(temperature=0, top_p=1, top_k=1),
    )
    transcript = r1.text.strip() if r1.text else ""

    if not transcript or len(transcript) < 10:
        transcript = "[No speech detected in recording]"
    status_fn(f"✅ Transcript ready ({len(transcript)} chars).")

    # ── Step B: Structure transcript → MOM (NO audio, text only) ───────────
    status_fn("📝 Step B — Structuring MOM from transcript...")
    r2 = client.models.generate_content(
        model=GEMINI_MODEL,
        contents=[
            MOM_PROMPT,
            (
                f"VERBATIM TRANSCRIPT:\n\n"
                f"--- START ---\n{transcript}\n--- END ---\n\n"
                "Extract the MOM strictly from the transcript above. "
                "Do not add anything not present in the transcript. "
                "Plain text. No Markdown. No asterisks."
            ),
        ],
        config=types.GenerateContentConfig(temperature=0, top_p=1, top_k=1),
    )
    mom = r2.text.strip() if r2.text else "Error: No output from model."
    return transcript, mom


# ── PDF generation ─────────────────────────────────────────────────────────────

def text_to_pdf(text: str, output_path: str):
    doc = SimpleDocTemplate(
        output_path, pagesize=letter,
        rightMargin=0.75*inch, leftMargin=0.75*inch,
        topMargin=0.75*inch,   bottomMargin=0.75*inch,
    )
    styles = getSampleStyleSheet()
    h_style = ParagraphStyle("h", parent=styles["Heading2"], fontSize=12, spaceAfter=4)
    b_style = ParagraphStyle("b", parent=styles["Normal"],   fontSize=11, spaceAfter=2,
                             alignment=TA_LEFT)
    elems = []
    for line in text.split("\n"):
        s = line.strip()
        if not s:
            elems.append(Spacer(1, 0.1*inch))
        elif s.endswith(":") or any(s.startswith(p) for p in
                                    ("Title:", "Date:", "Participants:")):
            elems.append(Paragraph(s, h_style))
        elif s.startswith("- "):
            elems.append(Paragraph("• " + s[2:], b_style))
        else:
            elems.append(Paragraph(s, b_style))
    doc.build(elems)


# ── Email ──────────────────────────────────────────────────────────────────────

def send_email(sender, password, receiver, pdf_path, mom_text):
    if not (sender and password and receiver):
        return False, "Missing credentials."
    msg = EmailMessage()
    msg["Subject"] = f"Minutes of Meeting — {TODAY_DATE}"
    msg["From"]    = sender
    msg["To"]      = receiver
    msg.set_content(
        f"Hi,\n\nAuto-generated Minutes of Meeting attached and pasted below.\n\n"
        f"{'='*60}\n{mom_text}\n{'='*60}\n\n"
        f"Please review for any inaccuracies.\n\nRegards,\nAI Meeting Assistant"
    )
    with open(pdf_path, "rb") as f:
        msg.add_attachment(f.read(), maintype="application", subtype="pdf",
                           filename=f"MOM_{TODAY_DATE}.pdf")
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as srv:
            srv.ehlo(); srv.starttls(); srv.ehlo()
            srv.login(sender, password)
            srv.send_message(msg)
        return True, "Email sent."
    except Exception as e:
        return False, str(e)


# ── Streamlit UI ───────────────────────────────────────────────────────────────

st.title("🤖 Auto MOM Generator")
st.markdown(
    "Joins Google Meet → records system audio via **WASAPI loopback** → "
    "transcribes verbatim → structures MOM → emails PDF."
)

with st.sidebar:
    st.header("⚙️ Configuration")
    api_key_input = st.text_input(
        "Gemini API Key", value=os.getenv("GEMINI_API_KEY", ""), type="password",
        help="https://aistudio.google.com/app/apikey"
    )
    st.divider()
    st.subheader("📧 Email")
    sender_email = st.text_input("Sender Gmail", value=os.getenv("SENDER_EMAIL", ""),
                                 placeholder="you@gmail.com")
    sender_pass  = st.text_input("Gmail App Password", value=os.getenv("SENDER_PASSWORD", ""),
                                 type="password")
    st.info("Use a Gmail App Password (myaccount.google.com/apppasswords), not your login password.")
    st.divider()
    st.subheader("🔊 Audio Setup")
    st.markdown(
        "**Required on Windows:**\n"
        "1. Right-click speaker icon → Sounds\n"
        "2. Recording tab → right-click blank area\n"
        "3. Show Disabled Devices\n"
        "4. Enable **Stereo Mix** (or it uses WASAPI loopback automatically)"
    )

col1, col2 = st.columns(2)
with col1:
    meet_url = st.text_input("Google Meet URL", placeholder="meet.google.com/abc-defg-hij")
with col2:
    receiver_email = st.text_input("Receiver Email", placeholder="manager@example.com")

st.markdown("---")
start_btn = st.button("🚀 Join Meeting & Start Agent", type="primary", use_container_width=True)

if start_btn:
    if not meet_url:
        st.error("Meeting URL is required.")
        st.stop()
    if not api_key_input:
        st.error("Gemini API Key is required.")
        st.stop()

    status_box = st.status("🔄 Initialising...", expanded=True)

    try:
        # Step 1 — Join & Record
        status_box.write("**Step 1/4** — Joining meeting and recording...")
        success = run_meet_bot(meet_url, status_box)
        if not success:
            status_box.update(label="❌ Bot failed", state="error")
            st.stop()

        # Step 2 — Transcribe + MOM
        status_box.write("**Step 2/4** — Transcribing & building MOM...")
        if not os.path.exists(OUTPUT_AUDIO):
            status_box.update(label="❌ Audio file missing", state="error")
            st.stop()

        transcript, mom_text = generate_mom_from_audio(
            OUTPUT_AUDIO, api_key_input, status_fn=status_box.write
        )
        st.session_state["transcript"] = transcript
        st.session_state["mom_text"]   = mom_text
        status_box.write("✅ MOM ready.")

        # Step 3 — PDF
        status_box.write("**Step 3/4** — Creating PDF...")
        text_to_pdf(mom_text, OUTPUT_PDF)
        status_box.write("✅ PDF created.")

        # Step 4 — Email
        if receiver_email and sender_email and sender_pass:
            status_box.write("**Step 4/4** — Sending email...")
            ok, msg = send_email(sender_email, sender_pass, receiver_email,
                                 OUTPUT_PDF, mom_text)
            status_box.write(f"{'✅' if ok else '❌'} {msg}")
        else:
            status_box.warning("⚠️ Email skipped — credentials incomplete.")

        status_box.update(label="✅ Done!", state="complete", expanded=False)
        st.success("🎉 Complete!")

        tab_mom, tab_tx = st.tabs(["📋 Minutes of Meeting", "🔤 Raw Transcript"])

        with tab_mom:
            st.text_area("MOM", value=mom_text, height=420, label_visibility="collapsed")
            with open(OUTPUT_PDF, "rb") as f:
                st.download_button("⬇️ Download PDF", f.read(),
                                   file_name=f"MOM_{TODAY_DATE}.pdf",
                                   mime="application/pdf", use_container_width=True)

        with tab_tx:
            st.caption(
                "Verbatim transcript extracted from the audio. "
                "MOM is built strictly from this text — nothing is invented."
            )
            st.text_area("Transcript", value=transcript, height=420,
                         label_visibility="collapsed")

    except Exception as exc:
        status_box.update(label="❌ Error", state="error")
        st.exception(exc)
