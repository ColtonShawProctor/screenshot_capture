# Excel Screenshot Service

A microservice that converts Excel spreadsheet ranges into PNG screenshots for the IDS Origination Agent.

## Features

- Converts specific Excel sheet ranges to PNG images
- Automatic table detection by header text
- Preserves formatting (bold, italic, currency, percentages)
- Case-insensitive sheet name matching
- Fuzzy matching for table names
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

### Detect and Capture Table
```
POST /detect-and-capture
```

Automatically finds a table by its header text and captures it as a PNG.

Request body:
```json
{
  "excelBase64": "<base64 encoded Excel file>",
  "tableName": "Sources and Uses",
  "searchSheets": ["S&U", "LTC and LTV Calcs"],  // optional - if omitted, search all sheets
  "padding": 2,  // optional - extra rows/cols around detected table (default: 2)
  "filename": "sources_uses.png"  // optional
}
```

Success Response:
```json
{
  "success": true,
  "filename": "sources_uses.png",
  "image": "<base64 encoded PNG>",
  "mimeType": "image/png",
  "detected": {
    "sheet": "S&U",
    "range": "B5:H24",
    "headerCell": "B5",
    "confidence": "exact"  // or "fuzzy"
  }
}
```

Not Found Response (404):
```json
{
  "success": false,
  "error": "Table 'Sources and Uses' not found",
  "searchedSheets": ["S&U", "LTC and LTV Calcs"],
  "suggestions": ["Sources & Uses (B7)", "Capital Stack at Closing (B51)"]
}
```

## Supported Table Names

The service recognizes these common table types in Fairbridge S&U Excel files:

| Table Name | Common Variations | Typical Location |
|------------|-------------------|------------------|
| Sources and Uses | "Sources & Uses", "Source and Use" | S&U sheet |
| Take Out Loan Sizing | "Takeout Loan Sizing", "Loan Sizing" | S&U sheet |
| Capital Stack at Closing | "Capital Stack", "Cap Stack" | S&U sheet |
| Loan to Cost | "LTC", "Loan-to-Cost" | LTC and LTV Calcs sheet |
| Loan to Value | "LTV", "Loan-to-Value" | LTC and LTV Calcs sheet |
| PILOT Schedule | "PILOT", "Payment in Lieu of Taxes" | Various sheets |
| Occupancy | "Unit Mix", "Occupancy Schedule" | Summary/I&E sheet |

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

# Detect and capture table (example)
curl -X POST http://localhost:3000/detect-and-capture \
  -H "Content-Type: application/json" \
  -d '{
    "excelBase64": "...",
    "tableName": "Sources and Uses"
  }'
```

## Deployment on Coolify

This service is deployed as a Docker container on Coolify. See main project README for deployment instructions.