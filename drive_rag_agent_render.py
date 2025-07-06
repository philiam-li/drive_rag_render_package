import os
import io
import fitz  # PyMuPDF
import docx
from fastapi import FastAPI
from docx import Document
from docx.shared import Pt
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import google.generativeai as genai

app = FastAPI()

GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
GOOGLE_SERVICE_ACCOUNT_FILE = os.getenv("GOOGLE_SERVICE_ACCOUNT_FILE", "service_account.json")
INPUT_FOLDER_NAME = "spec-inbox"
OUTPUT_FOLDER_NAME = "spec-outbox"
TEMP_FOLDER = "temp"

genai.configure(api_key=GEMINI_API_KEY)
gemini = genai.GenerativeModel("gemini-2.0-flash")

SCOPES = ["https://www.googleapis.com/auth/drive"]
credentials = service_account.Credentials.from_service_account_file(
    GOOGLE_SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
drive_service = build("drive", "v3", credentials=credentials)

def get_folder_id_by_name(folder_name):
    results = drive_service.files().list(
        q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder'",
        spaces="drive",
        fields="files(id, name)",
    ).execute()
    items = results.get("files", [])
    return items[0]["id"] if items else None

def list_files_in_folder(folder_id):
    results = drive_service.files().list(
        q=f"'{folder_id}' in parents",
        fields="files(id, name, mimeType)",
    ).execute()
    return results.get("files", [])

def save_file_to_temp(file_id, file_name):
    request = drive_service.files().get_media(fileId=file_id)
    fh = io.BytesIO()
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while not done:
        status, done = downloader.next_chunk()
    os.makedirs(TEMP_FOLDER, exist_ok=True)
    local_path = os.path.join(TEMP_FOLDER, file_name)
    with open(local_path, "wb") as f:
        f.write(fh.getbuffer())
    return local_path

def extract_text_from_file(file_path):
    if file_path.endswith(".txt"):
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    elif file_path.endswith(".docx"):
        doc = docx.Document(file_path)
        return "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    elif file_path.endswith(".pdf"):
        try:
            doc = fitz.open(file_path)
            return "\n".join([page.get_text() for page in doc])
        except Exception as e:
            print(f"❌ 無法解析 PDF：{e}")
            return ""
    print("⚠️ 不支援的檔案格式：", file_path)
    return ""

def analyze_text_with_gemini(text):
    if not text.strip():
        return "⚠️ 無法從文件中擷取任何內容，請確認格式或重新上傳。"
    prompt = f"""
你是一位專業的系統分析師。請根據以下規格文件，提出其中潛在的風險條款或需進一步釐清的點，並用條列式中文回答：

---
{text[:10000]}
---

請用繁體中文簡潔明確列出「風險」、「建議」與「需補充的資訊」。
"""
    print("📄 傳送給 Gemini 的 Prompt:\n", prompt[:500], "\n...\n（內容截斷）")
    response = gemini.generate_content(prompt)
    return response.text.strip()

def write_summary_to_docx(summary, output_path, original_filename=None):
    doc = Document()
    doc.add_heading('系統規格風險分析報告', level=1)
    if original_filename:
        doc.add_paragraph(f"📄 來源檔案：{original_filename}")
    doc.add_heading("📌 Gemini 分析結果", level=2)
    for line in summary.splitlines():
        doc.add_paragraph(line, style='List Bullet')
    doc.save(output_path)

def upload_file_to_folder(folder_id, local_path, file_name):
    file_metadata = {"name": file_name, "parents": [folder_id]}
    media = MediaFileUpload(local_path, resumable=True)
    drive_service.files().create(body=file_metadata, media_body=media, fields="id").execute()

@app.post("/webhook")
def run_agent():
    input_folder_id = get_folder_id_by_name(INPUT_FOLDER_NAME)
    output_folder_id = get_folder_id_by_name(OUTPUT_FOLDER_NAME)
    if not input_folder_id or not output_folder_id:
        return {"error": "❌ 找不到 Google Drive 資料夾"}

    files = list_files_in_folder(input_folder_id)
    if not files:
        return {"message": "📂 沒有可處理的檔案"}

    os.makedirs(TEMP_FOLDER, exist_ok=True)

    for file in files:
        file_name = file["name"]
        print(f"📥 處理中：{file_name}")
        local_path = save_file_to_temp(file["id"], file_name)
        text = extract_text_from_file(local_path)
        summary = analyze_text_with_gemini(text)
        docx_path = os.path.join(TEMP_FOLDER, f"報告_{file_name}.docx")
        write_summary_to_docx(summary, docx_path, original_filename=file_name)
        upload_file_to_folder(output_folder_id, docx_path, os.path.basename(docx_path))
        print(f"✅ 已完成並上傳報告：{os.path.basename(docx_path)}")

    return {"message": "🎉 所有檔案處理完成"}