import os
import base64
from io import BytesIO
from typing import Dict, Optional

from fastapi import FastAPI, HTTPException, Response
from fastapi.responses import StreamingResponse
from pydantic import BaseModel
import boto3
from botocore.config import Config
from docx import Document
from docx.shared import Inches

app = FastAPI(title="IDS Template Filler", version="1.0.0")

# S3 Configuration
S3_ENDPOINT = "https://nyc3.digitaloceanspaces.com"
S3_BUCKET = "fam.workspace"
S3_REGION = "nyc3"

# Initialize S3 client
s3_client = boto3.client(
    "s3",
    endpoint_url=S3_ENDPOINT,
    aws_access_key_id=os.getenv("S3_ACCESS_KEY"),
    aws_secret_access_key=os.getenv("S3_SECRET_KEY"),
    region_name=S3_REGION,
    config=Config(s3={'addressing_style': 'path'})  # Handle dot in bucket name
)

# Image width configuration
IMAGE_WIDTHS = {
    "IMAGE_SOURCES_USES": 6.5,
    "IMAGE_CAPITAL_STACK_CLOSING": 6.5,
    "IMAGE_LOAN_TO_COST": 6.0,
    "IMAGE_LTV_LTC": 6.0,
    "IMAGE_AERIAL_MAP": 5.0,
    "IMAGE_LOCATION_MAP": 5.0,
    "IMAGE_REGIONAL_MAP": 5.0,
    "IMAGE_SITE_PLAN": 5.5,
    "IMAGE_PILOT_SCHEDULE": 6.0,
    "IMAGE_TAKEOUT_SIZING": 6.0,
}

class FillRequest(BaseModel):
    placeholders: Dict[str, str] = {}
    images: Dict[str, str] = {}
    template_key: str = "_Templates/IDS_Template_Fairbridge.docx"
    output_filename: str = "IDS_Generated.docx"

class FillAndUploadRequest(BaseModel):
    placeholders: Dict[str, str] = {}
    images: Dict[str, str] = {}
    template_key: str = "_Templates/IDS_Template_Fairbridge.docx"
    output_key: str

@app.get("/health")
async def health_check():
    return {"status": "ok"}

def download_template(template_key: str) -> bytes:
    """Download template from S3."""
    try:
        response = s3_client.get_object(Bucket=S3_BUCKET, Key=template_key)
        return response['Body'].read()
    except Exception as e:
        raise HTTPException(status_code=404, detail=f"Template not found: {template_key}")

def upload_to_s3(content: bytes, key: str) -> str:
    """Upload content to S3 and return URL."""
    try:
        s3_client.put_object(
            Bucket=S3_BUCKET,
            Key=key,
            Body=content,
            ContentType='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )
        return f"{S3_ENDPOINT}/{S3_BUCKET}/{key}"
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to upload to S3: {str(e)}")

def replace_in_paragraph(paragraph, placeholders: Dict[str, str]) -> bool:
    """Replace placeholders in a paragraph, handling split runs."""
    full_text = ''.join(run.text for run in paragraph.runs)
    
    modified = False
    for key, value in placeholders.items():
        placeholder = "{{" + key + "}}"
        if placeholder in full_text:
            full_text = full_text.replace(placeholder, str(value))
            modified = True
    
    if modified and paragraph.runs:
        paragraph.runs[0].text = full_text
        for run in paragraph.runs[1:]:
            run.text = ""
    
    return modified

def replace_image_in_paragraph(paragraph, images: Dict[str, str]) -> bool:
    """Replace image placeholders in a paragraph."""
    full_text = ''.join(run.text for run in paragraph.runs)
    
    for key, image_data in images.items():
        placeholder = "{{" + key + "}}"
        if placeholder in full_text:
            # Clear the paragraph text
            for run in paragraph.runs:
                run.text = ""
            
            # Decode the base64 image
            try:
                image_bytes = base64.b64decode(image_data)
                image_stream = BytesIO(image_bytes)
                
                # Add the image to the first run
                width = IMAGE_WIDTHS.get(key, 6.0)
                paragraph.runs[0].add_picture(image_stream, width=Inches(width))
                
                return True
            except Exception as e:
                raise HTTPException(status_code=400, detail=f"Failed to decode image {key}: {str(e)}")
    
    return False

def process_paragraphs(paragraphs, placeholders: Dict[str, str], images: Dict[str, str]):
    """Process paragraphs for text and image replacements."""
    for paragraph in paragraphs:
        # Try text replacement first
        replace_in_paragraph(paragraph, placeholders)
        # Then try image replacement
        replace_image_in_paragraph(paragraph, images)

def fill_template(template_bytes: bytes, placeholders: Dict[str, str], images: Dict[str, str]) -> bytes:
    """Fill the template with placeholders and images."""
    doc = Document(BytesIO(template_bytes))
    
    # Process main document paragraphs
    process_paragraphs(doc.paragraphs, placeholders, images)
    
    # Process tables
    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                process_paragraphs(cell.paragraphs, placeholders, images)
    
    # Process headers and footers
    for section in doc.sections:
        # Header
        process_paragraphs(section.header.paragraphs, placeholders, images)
        # Footer
        process_paragraphs(section.footer.paragraphs, placeholders, images)
    
    # Save to bytes
    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()

@app.post("/fill")
async def fill_template_endpoint(request: FillRequest):
    """Fill template and return as download."""
    # Download template
    template_bytes = download_template(request.template_key)
    
    # Fill template
    filled_bytes = fill_template(template_bytes, request.placeholders, request.images)
    
    # Return as download
    return StreamingResponse(
        BytesIO(filled_bytes),
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": f"attachment; filename={request.output_filename}"}
    )

@app.post("/fill-and-upload")
async def fill_and_upload_endpoint(request: FillAndUploadRequest):
    """Fill template and upload to S3."""
    # Download template
    template_bytes = download_template(request.template_key)
    
    # Fill template
    filled_bytes = fill_template(template_bytes, request.placeholders, request.images)
    
    # Upload to S3
    output_url = upload_to_s3(filled_bytes, request.output_key)
    
    return {
        "success": True,
        "output_key": request.output_key,
        "output_url": output_url
    }

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)