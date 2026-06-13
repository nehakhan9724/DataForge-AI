# 📊 Visual Data Assistant

An AI-powered document extraction and analysis platform that transforms unstructured documents into structured, actionable data using Google's Gemini AI.

## 🚀 Overview

Visual Data Assistant helps users extract meaningful information from documents, images, spreadsheets, and text files. The application leverages Generative AI to identify key-value pairs, tables, summaries, and insights, presenting the results in an easy-to-use interface with Excel export support.

Whether you're working with invoices, reports, forms, or business documents, Visual Data Assistant automates data extraction and reduces manual effort.

---

## ✨ Features

### 📄 Multi-Format Document Support
- PDF Files
- Images (PNG, JPG, JPEG, WEBP, BMP, GIF)
- Excel Files (XLS, XLSX)
- Word Documents (DOCX)
- Text Files (TXT, MD, XML)
- Direct Text Input

### 🤖 AI-Powered Data Extraction
- Intelligent document understanding
- Key-value pair extraction
- Automatic table recognition
- Document summarization
- Structured JSON output generation

### 📈 Data Management
- Interactive table view
- Clean structured data presentation
- Excel export functionality
- Easy review and validation

### 🎨 User Experience
- Responsive UI
- Drag-and-drop file upload
- Real-time processing feedback
- Intuitive workflow

---

## 🏗️ Architecture

```text
User Upload
     │
     ▼
Next.js Frontend
     │
     ▼
FastAPI Backend
     │
     ▼
Gemini AI Processing
     │
     ▼
Structured JSON Response
     │
     ├── Table Visualization
     ├── Key Insights
     └── Excel Export
```

---

## 🛠️ Tech Stack

### Frontend
- Next.js
- React
- JavaScript
- XLSX Library

### Backend
- Python
- FastAPI
- Pandas
- Python-Docx

### AI & Data Processing
- Google Gemini 2.5 Flash Lite
- Document Parsing
- Structured Data Extraction
- Intelligent Table Recognition

---

## 📂 Project Structure

```text
visual-data-assistant/
│
├── frontend/
│   ├── components/
│   ├── pages/
│   ├── public/
│   └── utils/
│
├── backend/
│   ├── api/
│   ├── services/
│   ├── models/
│   └── utils/
│
├── uploads/
├── exports/
├── requirements.txt
├── package.json
└── README.md
```

---

## ⚙️ Installation

### Clone Repository

```bash
git clone https://github.com/yourusername/visual-data-assistant.git
cd visual-data-assistant
```

### Backend Setup

```bash
cd backend

python -m venv venv

# Windows
venv\Scripts\activate

# Linux/Mac
source venv/bin/activate

pip install -r requirements.txt
```

### Frontend Setup

```bash
cd frontend

npm install

npm run dev
```

---

## 🔑 Environment Variables

Create a `.env` file in the backend directory:

```env
GEMINI_API_KEY=your_gemini_api_key
```

---

## ▶️ Running the Application

### Start Backend

```bash
uvicorn main:app --reload
```

Backend runs on:

```text
http://localhost:8000
```

### Start Frontend

```bash
npm run dev
```

Frontend runs on:

```text
http://localhost:3000
```

---

## 📋 Example Workflow

1. Upload a document (PDF/Image/Excel/Word).
2. AI processes the content.
3. Extracted data is converted into structured format.
4. Results are displayed in tables and summaries.
5. Export results to Excel for further analysis.

---

## 🎯 Key Learning Outcomes

- Built a full-stack AI-powered application.
- Integrated Google Gemini API for document intelligence.
- Developed REST APIs using FastAPI.
- Implemented structured JSON extraction pipelines.
- Created scalable frontend architecture using Next.js.
- Designed exportable analytics workflows.

---

## 🔮 Future Enhancements

- Retrieval-Augmented Generation (RAG)
- Multi-document comparison
- OCR optimization
- Cloud deployment on Azure/AWS
- SharePoint integration
- Microsoft Power Platform integration
- Workflow automation agents

---

## 👩‍💻 Author

**Neha Khan**

Aspiring Automation & AI Engineer | Python Developer | AI & LLM Enthusiast

- LinkedIn: https://www.linkedin.com/in/midha-nehal-pathan-815b7a264
- GitHub: https://github.com/nehakhan9724

---

## 📄 License

This project is licensed under the MIT License.

---

⭐ If you found this project useful, consider giving it a star!
