import io
import os
import json
import requests
from google.oauth2.service_account import Credentials
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import fitz  # PyMuPDF
from PIL import Image, ImageChops
from reportlab.pdfgen import canvas
from reportlab.lib.utils import ImageReader
from reportlab.lib import colors
from datetime import datetime
import boto3

# --- CREDENTIALS & SECRETS ---
# Google Service Account JSON string will be passed as a secret
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON")

SHEET_ID = os.environ.get("SHEET_ID", "1_67K2zI1XRFAx8PB9NkV9lbwfT84uGOAyKCRrWxprxU")

# Cloudflare R2
R2_ACCOUNT_ID = os.environ.get("R2_ACCOUNT_ID")
R2_ACCESS_KEY = os.environ.get("R2_ACCESS_KEY")
R2_SECRET_KEY = os.environ.get("R2_SECRET_KEY")
R2_BUCKET_NAME = os.environ.get("R2_BUCKET_NAME", "whatsapp-reports")
R2_PUBLIC_BASE = os.environ.get("R2_PUBLIC_BASE", "https://pub-dd91ca5649e84cfe93f20e9fb4468ed8.r2.dev")

# AiSensy WhatsApp
AISENSY_API_KEY = os.environ.get("AISENSY_API_KEY")
CAMPAIGN_NAME = os.environ.get("CAMPAIGN_NAME", "Central Analytics PDF Automation")

# Parse comma-separated destinations, or fallback to default
dests_env = os.environ.get("DESTINATIONS", "916303054457")
DESTINATIONS = [d.strip() for d in dests_env.split(",")]

TODAY = datetime.now().strftime("%d %B %Y")
TIMESTAMP = datetime.now().strftime("%Y%m%d_%H%M%S")
FILE_NAME = "VD_Report.pdf"

print("✅ Environment Variables Loaded")

def get_google_creds():
    if not GOOGLE_CREDENTIALS_JSON:
        raise Exception("GOOGLE_CREDENTIALS_JSON environment variable is missing")
    info = json.loads(GOOGLE_CREDENTIALS_JSON)
    creds = Credentials.from_service_account_info(
        info,
        scopes=["https://www.googleapis.com/auth/drive.readonly", "https://www.googleapis.com/auth/spreadsheets.readonly"]
    )
    creds.refresh(Request())
    return creds

def get_sheet_gid(creds, sheet_name):
    service = build("sheets", "v4", credentials=creds)
    meta = service.spreadsheets().get(spreadsheetId=SHEET_ID).execute()
    for sheet in meta["sheets"]:
        if sheet["properties"]["title"] == sheet_name:
            return str(sheet["properties"]["sheetId"])
    raise Exception(f"Sheet {sheet_name} not found")

def trim_white_space(pil_img):
    bg = Image.new(pil_img.mode, pil_img.size, (255, 255, 255))
    diff = ImageChops.difference(pil_img, bg)
    diff = ImageChops.add(diff, diff, 2.0, -100)
    bbox = diff.getbbox()
    if bbox:
        padding = 15
        b_x0 = max(0, bbox[0] - padding)
        b_y0 = max(0, bbox[1] - padding)
        b_x1 = min(pil_img.size[0], bbox[2] + padding)
        b_y1 = min(pil_img.size[1], bbox[3] + padding)
        return pil_img.crop((b_x0, b_y0, b_x1, b_y1))
    return pil_img

def export_range_image(creds, sheet_name, range_name):
    sheet_gid = get_sheet_gid(creds, sheet_name)

    export_url = (
        f"https://docs.google.com/spreadsheets/d/{SHEET_ID}/export"
        f"?format=pdf&gid={sheet_gid}&range={range_name}&size=A2&portrait=true&fitw=true"
        f"&scale=2&sheetnames=false&printtitle=false&pagenumbers=false&gridlines=false"
        f"&fzr=false&top_margin=0&bottom_margin=0&left_margin=0&right_margin=0"
    )

    response = requests.get(export_url, headers={"Authorization": f"Bearer {creds.token}"}, timeout=120)
    response.raise_for_status()

    if not response.content.startswith(b"%PDF"):
        raise Exception("Invalid PDF returned from Google")

    doc = fitz.open(stream=response.content, filetype="pdf")
    page = doc[0]
    pix = page.get_pixmap(dpi=450)
    img_bytes = pix.tobytes("png")
    
    pil_img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    cropped_img = trim_white_space(pil_img)
    
    cropped_bytes = io.BytesIO()
    cropped_img.save(cropped_bytes, format="PNG")
    cropped_bytes.seek(0)
    
    return ImageReader(cropped_bytes), cropped_img.width, cropped_img.height

def generate_dynamic_single_page_clean():
    creds = get_google_creds()

    sections = [
        ("VD Report", "A26:U32", "Hello Team, Yesterday's Leader level sales view summary along with their day target's."),
        ("VD Top Batch Day View", "A5:F20", "#Top Batch Day Sales View with 2-year comparison and YoY growth."),
        ("VD Top Batch Day View", "H6:M20", "#Top Batch YTD Sales View with 2-year comparison and YoY growth.")
    ]

    images_data = []
    total_h = 0
    PAGE_WIDTH = 1800
    MARGIN = 70 
    USABLE_WIDTH = PAGE_WIDTH - (MARGIN * 2)

    print("📄 Capturing regions from Google Sheets...")
    for sheet_name, range_name, description in sections:
        print(f"   -> {sheet_name} ({range_name})")
        img_reader, w, h = export_range_image(creds, sheet_name, range_name)
        
        scale = USABLE_WIDTH / w
        target_w = USABLE_WIDTH
        target_h = h * scale
        
        total_h += target_h + 150  
        images_data.append((img_reader, target_w, target_h, description))
        
    PAGE_HEIGHT = total_h + MARGIN
    buffer = io.BytesIO()
    c = canvas.Canvas(buffer, pagesize=(PAGE_WIDTH, PAGE_HEIGHT))
    current_y = PAGE_HEIGHT - MARGIN

    for img_reader, target_w, target_h, description in images_data:
        current_y -= 50
        c.setFont("Helvetica-Bold", 32)
        c.setFillColor(colors.black)
        c.drawCentredString(PAGE_WIDTH / 2.0, current_y, description)
        current_y -= (target_h + 40)
        c.drawImage(img_reader, MARGIN, current_y, width=target_w, height=target_h, preserveAspectRatio=True, mask='auto')
        current_y -= 80

    c.save()
    buffer.seek(0)
    print("✅ FINAL: Large, UHD, Single Page Dynamic PDF Generated")
    return buffer

def upload_to_r2(pdf_buffer):
    s3 = boto3.client(
        service_name="s3",
        endpoint_url=f"https://{R2_ACCOUNT_ID}.r2.cloudflarestorage.com",
        aws_access_key_id=R2_ACCESS_KEY,
        aws_secret_access_key=R2_SECRET_KEY,
        region_name="auto"
    )

    s3.put_object(
        Bucket=R2_BUCKET_NAME,
        Key=FILE_NAME,
        Body=pdf_buffer.read(),
        ContentType="application/pdf"
    )

    public_url = f"{R2_PUBLIC_BASE}/{FILE_NAME}"

    print("✅ Uploaded to Cloudflare R2")
    print("🔗 Public URL:", public_url)

    return public_url

def send_to_aisensy(url):
    endpoint = "https://backend.aisensy.com/campaign/t1/api/v2"

    for dest in DESTINATIONS:
        payload = {
            "apiKey": AISENSY_API_KEY,
            "campaignName": CAMPAIGN_NAME,
            "destination": dest,
            "userName": "PW Online- Analytics",
            "templateParams": [TODAY],
            "source": "r2-centered",
            "media": {
                "url": url,
                "filename": FILE_NAME
            }
        }

        response = requests.post(
            endpoint,
            json=payload,
            headers={"Content-Type": "application/json"}
        )

        print(f"📱 Sent to WhatsApp ({dest}):", response.status_code, response.text)

if __name__ == "__main__":
    try:
        # Validate that essential secrets exist
        missing_vars = []
        for v in ["GOOGLE_CREDENTIALS_JSON", "R2_ACCOUNT_ID", "R2_ACCESS_KEY", "R2_SECRET_KEY", "AISENSY_API_KEY"]:
            if not os.environ.get(v):
                missing_vars.append(v)
        
        if missing_vars:
            raise Exception(f"Missing required environment variables: {', '.join(missing_vars)}")
            
        print("🚀 Starting Dynamically Scaled PDF Generation...")
        pdf_buffer = generate_dynamic_single_page_clean()

        public_pdf_url = upload_to_r2(pdf_buffer)

        send_to_aisensy(public_pdf_url)

        print("🎉 Successfully completed dynamic PDF automation via GitHub Actions!")
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"❌ Error occurred: {e}")
        exit(1)
fy27_orders
