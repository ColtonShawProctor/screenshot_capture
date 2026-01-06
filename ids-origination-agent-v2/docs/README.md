# IDS Origination Agent V2 - Deployment Guide

Automated Initial Deal Summary generation for commercial real estate loan origination.

## 📋 Overview

This n8n workflow automates the IDS creation process by:
- Processing deal folders from S3
- Extracting data from Quick Sheets, Term Sheets, and Financial Models
- Capturing Excel table screenshots automatically
- Performing property and sponsor research
- Analyzing risks and assigning borrower grades
- Generating completed IDS documents

## 🏗️ Architecture

```
LibreChat (MCP) → n8n Workflow → Screenshot Service
                      ↓
                  S3 Storage ← APIs (Gemini, Perplexity)
```

## 📦 Deployment Steps

### 1. Deploy Screenshot Microservice on Coolify

1. **Create new Docker Compose resource** in Coolify
2. **Point to this repository** (screenshot-service folder)
3. **Use this compose configuration:**

```yaml
services:
  excel-screenshot:
    build:
      context: ./screenshot-service
      dockerfile: Dockerfile
    environment:
      - NODE_ENV=production
    # No external ports - internal only
```

4. **Deploy the service**

### 2. Configure Networking

⚠️ **Critical Step:** Both services must communicate internally.

1. **Enable cross-service networking:**
   - Go to your existing n8n service → Settings
   - Enable **"Connect to Predefined Networks"**
   - Note the **Destination**

2. **Enable networking on screenshot service:**
   - Go to screenshot service → Settings  
   - Enable **"Connect to Predefined Networks"**
   - Select **same Destination** as n8n

3. **Redeploy both services** to apply network changes

### 3. Add Environment Variables to n8n

In Coolify, add these to your **existing n8n service**:

```bash
OPENROUTER_API_KEY=your_openrouter_key
PERPLEXITY_API_KEY=your_perplexity_key
```

### 4. Create n8n Credentials

In n8n UI, create these credentials:

| Name | Type | Configuration |
|------|------|---------------|
| `openRouterApiKey` | HTTP Header Auth | Header: `Authorization`, Value: `Bearer your_key` |
| `perplexityApiKey` | HTTP Header Auth | Header: `Authorization`, Value: `Bearer your_key` |
| `digitalOceanSpaces` | AWS S3 | Access Key + Secret for DO Spaces |

### 5. Upload IDS Template

Upload `IDS_Template_Fairbridge_Generic__5_.docx` to:
```
s3://your-bucket/templates/IDS_Template_Fairbridge_Generic__5_.docx
```

### 6. Import Workflow

1. In n8n UI → Templates → Import
2. Upload `ids-origination-agent-v2.json`
3. **Activate** the workflow

## 🧪 Testing

### Test Internal Networking

Create a test HTTP Request node in n8n:
```
Method: GET
URL: http://excel-screenshot:3000/health
```

Should return: `{"status": "healthy", "service": "excel-screenshot"}`

### Test Webhook Trigger

Send a POST request to your n8n webhook URL:

```bash
curl -X POST https://your-n8n-domain.com/webhook/ids-origination-webhook \
  -H "Content-Type: application/json" \
  -d '{
    "deal_folder_path": "s3://your-bucket/deals/2025/test-deal/",
    "deal_name": "Test Property",
    "originator": "John Smith"
  }'
```

## 📊 Workflow Stages

1. **Trigger & Validation** - MCP webhook receives deal folder path
2. **Document Discovery** - Lists and identifies source documents
3. **Data Extraction** - Gemini extracts deal facts from Quick Sheet
4. **Table Screenshots** - Puppeteer captures S&U and LTV/LTC tables
5. **Research** - Perplexity researches property location and sponsor
6. **Risk Analysis** - Gemini identifies top 5 risks with mitigants
7. **Document Generation** - Populates IDS template with all data
8. **Delivery** - Uploads final IDS and returns download URL

## 🔧 Configuration

### Document Identification Patterns

The workflow identifies documents by filename:

| Document | Pattern |
|----------|---------|
| Quick Sheet | `*quick*sheet*`, `*quicksheet*` |
| Term Sheet | `*term*sheet*`, `*terms*` |
| Financial Model | `*model*`, `*proforma*` |
| Appraisal | `*appraisal*` |
| Rent Roll | `*rent*roll*` |

### Excel Sheet Names

Screenshot service looks for these sheets:
- **S&U Table:** "Sources & Uses" (range A1:H30)
- **LTV/LTC Table:** "Summary" (range A1:F20)

Names are matched case-insensitively with partial matching.

## 🚨 Troubleshooting

### Screenshot Service Issues

1. **Service won't start:**
   ```bash
   # Check logs in Coolify
   docker logs excel-screenshot
   ```

2. **n8n can't reach screenshot service:**
   - Verify both services have "Connect to Predefined Networks" enabled
   - Ensure same Destination is selected
   - Redeploy both services

3. **Excel conversion fails:**
   - Check file is valid Excel format
   - Verify sheet name exists
   - Ensure range is valid (e.g., "A1:H30")

### API Issues

1. **Gemini extraction fails:**
   - Check OpenRouter API key
   - Verify model name: `google/gemini-2.5-flash`
   - Check file size limits

2. **Perplexity research fails:**
   - Verify API key
   - Check rate limits (add delays if needed)

3. **S3 upload/download fails:**
   - Verify DO Spaces credentials
   - Check bucket permissions
   - Ensure file paths are correct

### Workflow Errors

1. **Missing documents:**
   - Workflow continues with available files
   - Check filename patterns match your files
   - Review missing_documents in response

2. **Template population fails:**
   - Ensure template is uploaded to S3
   - Check placeholder format in template
   - Review extracted data structure

## 📈 Performance

- **Average execution time:** 3-5 minutes per deal
- **Rate limiting:** 1-second delays between API calls
- **Concurrency:** Process multiple deals simultaneously
- **Timeout settings:** 30s for API calls, 60s for screenshots

## 🔐 Security

- API keys stored as n8n credentials
- Internal service communication (no external ports)
- S3 presigned URLs with 24-hour expiry
- No sensitive data logged

## 📝 Response Format

Successful completion returns:

```json
{
  "status": "success",
  "deal_name": "Example Property",
  "ids_document_url": "https://...",
  "summary": {
    "property": "123 Main St",
    "loan_amount": "$5,000,000",
    "borrower_grade": "B",
    "top_risk": "Market risk description"
  },
  "processing_time_seconds": 180,
  "documents_processed": ["quickSheet", "termSheet", "financialModel"],
  "missing_documents": []
}
```

## 🔄 Updates

To update the workflow:
1. Modify workflow in n8n UI, or
2. Export updated JSON and re-import
3. Update screenshot service by pushing to Git repo

## 📞 Support

For issues:
1. Check n8n execution logs
2. Review Coolify service logs
3. Verify API key quotas/limits
4. Test individual workflow nodes