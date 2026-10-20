# PaperFinder Pro — Backend API

Flask backend for PaperFinder Pro, providing PDF extraction, AI-powered academic paper generation (via Groq), search clarification, and chat-with-PDF capabilities.

## Features
- **PDF Extraction**: Extracts text and metadata using PyMuPDF (`fitz`) and `pdfplumber`.
- **AI Paper Generation**: Synthesizes and drafts IEEE-formatted `.docx` documents using Groq LLM (`llama-3.3-70b-versatile`).
- **PDF Chat**: Context-aware Q&A over research papers using Groq LLM.
- **Search Clarification**: Smart question & keyword suggestions for literature review queries.
- **Health Check**: `/health` endpoint for uptime monitoring and health checks.

## Environment Variables
Create a `.env` file in the root directory:
```env
GROQ_API_KEY=gsk_your_groq_api_key_here
PORT=8080
```

## Local Development
1. Create and activate a virtual environment:
   ```bash
   python3 -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   ```
2. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```
3. Run the development server:
   ```bash
   python app.py
   ```
   Or using Gunicorn:
   ```bash
   gunicorn app:app --bind 0.0.0.0:8080
   ```

## Deploying to Render (Free Tier)
1. Push this folder to GitHub.
2. In [Render Dashboard](https://dashboard.render.com), click **New +** → **Web Service**.
3. Connect your GitHub repository (`paperfinder`).
4. Settings:
   - **Name**: `paperfinder`
   - **Runtime**: `Python 3`
   - **Build Command**: `pip install -r requirements.txt`
   - **Start Command**: `gunicorn app:app --bind 0.0.0.0:${PORT:-8080} --timeout 120 --workers 2`
   - **Instance Type**: `Free`
5. Under **Environment Variables**, add:
   - `GROQ_API_KEY`: Your Groq API key (`gsk_...`)
   - `PYTHON_VERSION`: `3.11.9`
6. Click **Create Web Service**.
7. Once live, your backend URL will be: `https://<your-service-name>.onrender.com`.
   Verify at `https://<your-service-name>.onrender.com/health`.
