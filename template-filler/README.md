# IDS Template Filler Microservice

A FastAPI microservice for filling Word document templates with dynamic content and images.

## Features

- Fill Word documents with text placeholders
- Insert base64-encoded images into documents
- Download filled documents directly
- Upload filled documents to S3/DigitalOcean Spaces
- Preserves document formatting

## Environment Variables

```bash
S3_ACCESS_KEY=your_access_key
S3_SECRET_KEY=your_secret_key
```

## API Endpoints

### Health Check
```
GET /health
```

### Fill Template
```
POST /fill
```
Fill template and download the result.

Request body:
```json
{
  "placeholders": {
    "LOAN_AMOUNT": "$25,650,000",
    "PROPERTY_ADDRESS": "89 Montauk Highway, East Moriches, NY"
  },
  "images": {
    "IMAGE_SOURCES_USES": "<base64_encoded_png>"
  },
  "template_key": "_Templates/IDS_Template_Fairbridge.docx",
  "output_filename": "IDS_Generated.docx"
}
```

### Fill and Upload Template
```
POST /fill-and-upload
```
Fill template and upload to S3.

Request body:
```json
{
  "placeholders": {
    "LOAN_AMOUNT": "$25,650,000"
  },
  "images": {
    "IMAGE_SOURCES_USES": "<base64_encoded_png>"
  },
  "template_key": "_Templates/IDS_Template_Fairbridge.docx",
  "output_key": "outputs/IDS_Generated.docx"
}
```

Response:
```json
{
  "success": true,
  "output_key": "outputs/IDS_Generated.docx",
  "output_url": "https://fam.workspace.nyc3.digitaloceanspaces.com/outputs/IDS_Generated.docx"
}
```

## Local Development

1. Build the Docker image:
```bash
docker build -t template-filler .
```

2. Run the container:
```bash
docker run -p 8000:8000 \
  -e S3_ACCESS_KEY=your_key \
  -e S3_SECRET_KEY=your_secret \
  template-filler
```

3. Test the service:
```bash
# Health check
curl http://localhost:8000/health

# Fill template
curl -X POST http://localhost:8000/fill \
  -H "Content-Type: application/json" \
  -d '{
    "placeholders": {
      "LOAN_AMOUNT": "$25,000,000",
      "PROPERTY_ADDRESS": "123 Main St"
    }
  }' \
  --output test.docx
```

## Deployment on Coolify

1. Push this repository to GitHub
2. In Coolify:
   - Create new service from GitHub repository
   - Set environment variables:
     - `S3_ACCESS_KEY`
     - `S3_SECRET_KEY`
   - Set health check path: `/health`
   - Port: `8000`

## Placeholder Format

Text placeholders: `{{PLACEHOLDER_NAME}}`
Image placeholders: `{{IMAGE_*}}`

## Supported Image Placeholders

- `IMAGE_SOURCES_USES` (6.5 inches)
- `IMAGE_CAPITAL_STACK_CLOSING` (6.5 inches)
- `IMAGE_LTV_LTC` (6.0 inches)
- `IMAGE_AERIAL_MAP` (5.0 inches)
- `IMAGE_LOCATION_MAP` (5.0 inches)
- `IMAGE_REGIONAL_MAP` (5.0 inches)
- `IMAGE_SITE_PLAN` (5.5 inches)
- `IMAGE_PILOT_SCHEDULE` (6.0 inches)
- `IMAGE_TAKEOUT_SIZING` (6.0 inches)