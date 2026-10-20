# PaperFinder Pro — Backend API

Flask backend for PaperFinder Pro, providing PDF extraction, AI-powered academic paper generation (via Groq), search clarification, and chat-with-PDF capabilities.

## Features
- **PDF Extraction**: Extracts text and metadata using PyMuPDF (`fitz`) and `pdfplumber`.
- **AI Paper Generation**: Synthesizes and drafts IEEE-formatted `.docx` documents using Groq LLM (`llama-3.3-70b-versatile`).
- **PDF Chat**: Context-aware Q&A over research papers using Groq LLM.
- **Search Clarification**: Smart question & keyword suggestions for literature review queries.
- **Health Check**: `/health` endpoint for uptime monitoring and Railway health checks.

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

## Deploying to Railway
1. Push this folder/repository to GitHub.
2. In [Railway](https://railway.app), click **New Project** → **Deploy from GitHub repo**.
3. Select this repository.
4. Go to **Variables** tab in your Railway service and add:
   - `GROQ_API_KEY`: Your Groq API key (`gsk_...`)
5. In **Settings** → **Networking**, click **Generate Domain** to get your public backend URL (e.g. `https://your-service.up.railway.app`).
6. Copy that URL and set it in your frontend as `VITE_BACKEND_URL`.
