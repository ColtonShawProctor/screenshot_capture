# IDS Origination Agent V2

**Automated Initial Deal Summary generation for Fairbridge Commercial Lending**

Version 2.0 introduces automated table screenshots, enhanced research capabilities, and improved error handling.

## 📁 Project Structure

```
ids-origination-agent-v2/
├── ids-origination-agent-v2.json     # n8n workflow - import into existing instance
├── docker-compose.yml               # Coolify deployment for screenshot service
├── screenshot-service/              # Excel to PNG microservice
│   ├── Dockerfile
│   ├── package.json
│   ├── server.js
│   └── README.md
└── docs/
    └── README.md                    # Detailed deployment guide
```

## 🚀 Quick Start

### 1. Deploy Screenshot Service
```bash
# In Coolify: New Resource → Docker Compose
# Point to this repository
# Select docker-compose.yml
```

### 2. Configure Networking
- Enable "Connect to Predefined Networks" on both n8n and screenshot services
- Use same Destination for both services

### 3. Import Workflow
- Upload `ids-origination-agent-v2.json` to your existing n8n instance
- Configure API credentials (OpenRouter, Perplexity, S3)
- Activate workflow

### 4. Test
```bash
curl -X POST https://your-n8n.com/webhook/ids-origination-webhook \
  -H "Content-Type: application/json" \
  -d '{
    "deal_folder_path": "s3://bucket/deals/2025/test-deal/",
    "deal_name": "Test Property",
    "originator": "John Smith"
  }'
```

## ✨ Key Features

- **🔄 Automated Processing:** End-to-end IDS generation from deal folders
- **📊 Table Screenshots:** Puppeteer-powered Excel table capture
- **🔍 Enhanced Research:** Property location analysis and sponsor background checks  
- **⚖️ Risk Assessment:** Automated risk identification with mitigation strategies
- **📋 Borrower Grading:** AI-powered sponsor risk grading (A-F scale)
- **🛡️ Error Handling:** Graceful degradation with partial results
- **📈 Performance:** 3-5 minute execution time per deal

## 🔧 Requirements

### Infrastructure
- Existing n8n instance on Coolify ✓
- Digital Ocean Spaces (S3-compatible storage)
- Docker environment for screenshot service

### API Keys
- **OpenRouter:** For Gemini 2.5 Flash/Pro access
- **Perplexity:** For sonar-pro research model
- **Digital Ocean Spaces:** For file storage

### Input Documents
- Quick Sheet (Excel)
- Term Sheet (PDF) 
- Financial Model (Excel with S&U and Summary sheets)
- Appraisal (PDF) - optional
- Rent Roll (Excel/PDF) - optional

## 📋 Workflow Overview

1. **Trigger** - MCP webhook from LibreChat
2. **Discovery** - Identify documents in S3 deal folder
3. **Extraction** - Extract deal facts via Gemini
4. **Screenshots** - Capture Excel tables as PNGs
5. **Research** - Analyze property location and sponsor background
6. **Risk Analysis** - Identify top 5 risks with mitigants
7. **Generation** - Populate IDS template with all data
8. **Delivery** - Return presigned download URL

## 🎯 Output

**IDS Document:** Fully populated Word document with:
- Deal summary and terms
- Property analysis with location research
- Sponsor background and risk grade
- Sources & Uses table screenshot
- LTV/LTC metrics table screenshot  
- Top 5 risks with recommended mitigants

**Response JSON:** Summary data for LibreChat integration

## 📚 Documentation

- **[Deployment Guide](docs/README.md)** - Complete setup instructions
- **[Screenshot Service](screenshot-service/README.md)** - Microservice documentation

## 🔄 Version History

| Version | Date | Key Changes |
|---------|------|-------------|
| V1 | Dec 2024 | Basic extraction and manual screenshots |
| V2 | Jan 2025 | Automated screenshots, enhanced research, borrower grading |

## 📞 Support

For deployment issues:
1. Check [Troubleshooting Guide](docs/README.md#troubleshooting)
2. Review n8n execution logs
3. Verify API credentials and quotas

---

**Target Release:** January 9th, 2025  
**Client:** Fairbridge Commercial Lending