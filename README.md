# PDF Table Extractor API

Backend API for extracting tables from PDF files with OCR support.

## Quick Start

### Local Development
```bash
pip install -r requirements.txt
python app.py
```

### Deploy to Render
1. Push to GitHub
2. Connect to Render
3. Deploy with `./render-build.sh`

## API Endpoints
- `GET /api/health` - Health check
- `POST /api/extract` - Extract PDF tables
- `POST /api/recalculate` - Recalculate data
- `POST /api/download-excel` - Generate Excel

## Environment
- Python 3.11
- Flask + Gunicorn
- Camelot + EasyOCR
