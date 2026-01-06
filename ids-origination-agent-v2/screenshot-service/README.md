# Excel Screenshot Service

A microservice that converts Excel spreadsheet ranges into PNG screenshots for the IDS Origination Agent.

## Features

- Converts specific Excel sheet ranges to PNG images
- Preserves formatting (bold, italic, currency, percentages)
- Case-insensitive sheet name matching
- Handles large Excel files via base64 encoding
- Health check endpoint for monitoring

## API Endpoints

### Health Check
```
GET /health
```
Returns service status.

### Convert Excel to PNG
```
POST /convert
```

Request body:
```json
{
  "excelBase64": "<base64 encoded Excel file>",
  "sheetName": "Sources & Uses",
  "range": "A1:H30",
  "filename": "output.png"  // optional
}
```

Response:
```json
{
  "success": true,
  "filename": "output.png",
  "image": "<base64 encoded PNG>",
  "mimeType": "image/png"
}
```

## Local Development

```bash
npm install
npm start
```

Test with curl:
```bash
# Health check
curl http://localhost:3000/health

# Convert Excel (example)
curl -X POST http://localhost:3000/convert \
  -H "Content-Type: application/json" \
  -d '{
    "excelBase64": "...",
    "sheetName": "Summary",
    "range": "A1:F20"
  }'
```

## Deployment on Coolify

This service is deployed as a Docker container on Coolify. See main project README for deployment instructions.