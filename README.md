



# 🤖 Auto MOM Generator

An automated AI agent that joins Google Meet sessions, records system audio, transcribes the conversation verbatim, and generates structured **Minutes of Meeting (MOM)** delivered straight to your inbox.

Built with **Streamlit**, **Playwright**, and **Google Gemini 3 Flash**.

---
# Body

<img width="1914" height="786" alt="Image" src="https://github.com/user-attachments/assets/39d9e7c3-89c3-4848-8f96-ef9d15e3ab76" />

---
## ✨ Features

* **Browser Automation**: Uses Playwright to navigate Google Meet, bypasses permission popups, and joins calls automatically.
* **System Audio Capture**: Implements `pyaudiowpatch` to utilize **WASAPI Loopback**, capturing crystal-clear meeting audio directly from the system output (no microphone echo).
* **Two-Step AI Processing**:
    * **Step A**: High-fidelity verbatim transcription using `gemini-3-flash-preview`.
    * **Step B**: Strict MOM extraction (Agenda, Discussion, Decisions, Action Items) ensuring zero AI hallucinations.
* **Professional Reporting**: Generates a clean PDF using `ReportLab`.
* **Automated Delivery**: Sends the MOM PDF to a designated email address via Gmail SMTP.

---

## 🛠️ Technical Architecture



1.  **Frontend**: Streamlit Dashboard for configuration and real-time status monitoring.
2.  **Audio Engine**: `pyaudiowpatch` (Windows WASAPI) for internal loopback recording.
3.  **Inference**: Google Generative AI (Gemini API) for audio-to-text and text-to-structured-MOM.
4.  **Export**: `ReportLab` for PDF orchestration and `smtplib` for email dispatch.

---

## 🚀 Getting Started

### 1. Prerequisites
* **OS**: Windows (Required for WASAPI loopback support).
* **Python**: 3.9+ 
* **Google Gemini API Key**: [Get it here](https://aistudio.google.com/app/apikey).
* **Gmail App Password**: Required for the email feature. [Generate one here](https://myaccount.google.com/apppasswords).

### 2. Installation

First, ensure you do **not** have standard `pyaudio` installed, as it conflicts with the patched version.

```bash
# Clone the repository
git clone https://github.com/your-username/auto-mom-generator.git
cd auto-mom-generator

# Install dependencies
pip install streamlit playwright pyaudiowpatch google-genai reportlab python-dotenv numpy soundfile
pip install -r requirements.txt
python -m playwright install

# Install Playwright browser binaries
playwright install chromium
```

### 3. Windows Audio Setup
To ensure the bot can "hear" the meeting:
1.  Right-click the **Speaker Icon** in your Taskbar → **Sounds**.
2.  Navigate to the **Recording** tab.
3.  Right-click a blank area and select **"Show Disabled Devices"**.
4.  If **Stereo Mix** appears, right-click and **Enable** it. 
    *(Note: The code uses WASAPI loopback, so it may work even without Stereo Mix enabled, provided the loopback device is detected.)*

### 4. Environment Variables
Create a `.env` file in the root directory:
```env
GEMINI_API_KEY=your_api_key_here
SENDER_EMAIL=your_gmail@gmail.com
SENDER_PASSWORD=your_app_password
```
code = dkeh zazi rxtd esmu 
---

## 🖥️ Usage

1.  Run the Streamlit application:
    ```bash
    streamlit run app.py
    ```
2.  Paste the **Google Meet URL**.
3.  Enter the **Receiver Email**.
4.  Click **"Join Meeting & Start Agent"**.
5.  The bot will launch a browser, join the call, and wait for the meeting to end (or for participants to leave) before generating the report.

---

## ⚠️ Important Notes

* **Ethics & Privacy**: Always inform meeting participants that an automated recording/transcription tool is present.
* **Headless Mode**: The current configuration runs in `headless=False` so you can monitor the bot's interactions. You can toggle this in the `run_meet_bot` function.
* **Profile Locking**: If your default browser profile is in use, the bot will automatically generate a temporary profile to avoid conflicts.

---

## 🤝 Contributing
Feel free to fork this project and submit Pull Requests. For major changes, please open an issue first to discuss what you would like to change.

---
