# 📊 Excel AI Analyzer

**Ask questions about your Excel and CSV files using a local AI — 100% private, no data leaves your computer.**

This project is a beginner-friendly demonstration of how to build a simple web application that combines:

- A **Python Flask backend** (`server.py`) that reads and converts your spreadsheet into structured text
- A **single-page HTML frontend** (`excel-ai-analyzer.html`) that provides the chat interface
- **Ollama** — a free, local AI engine that runs entirely on your own machine

Because everything runs locally, your spreadsheet data is never sent to any cloud service or third-party API.

---

## 🗂️ Project Files

```
excel-ai-analyzer/
├── server.py                  # Python backend (Flask + Pandas + Ollama)
├── excel-ai-analyzer.html     # Browser-based frontend (open directly)
└── README.md                  # This file
```

---

## ✅ Prerequisites

Before you begin, make sure you have the following installed:

| Requirement | Purpose | Download |
|---|---|---|
| **Python 3.9+** | Runs the backend server | https://www.python.org |
| **Conda** | Manages your Python environment (recommended) | https://docs.conda.io/en/latest/miniconda.html |
| **Ollama** | Runs the local AI model | https://ollama.com |

---

## 🚀 Step-by-Step Setup

### Step 1 — Install Ollama

Go to [https://ollama.com](https://ollama.com) and download the installer for your operating system (Mac, Windows, or Linux). Follow the on-screen instructions to install it.

Once installed, open a terminal and pull the AI model the app uses by default:

```bash
ollama pull qwen2.5
```

> This will download the model (around 4–5 GB). You only need to do this once.

---

### Step 2 — Create a Conda Environment

Using a Conda environment keeps your project dependencies isolated and avoids conflicts with other Python projects on your computer.

Open a terminal (or Anaconda Prompt on Windows) and run:

```bash
conda create -n excel-ai python=3.11 -y
conda activate excel-ai
```

You should see `(excel-ai)` appear at the start of your terminal prompt. This confirms you are inside the environment.

---

### Step 3 — Install Python Dependencies

With your environment activated, install the required libraries:

```bash
pip install flask flask-cors pandas openpyxl xlrd requests numpy
```

Here is what each package does:

| Package | Purpose |
|---|---|
| `flask` | The lightweight web server that runs the Python backend |
| `flask-cors` | Allows the HTML frontend to communicate with the backend |
| `pandas` | Reads and processes Excel and CSV files |
| `openpyxl` | Engine for reading modern Excel files (`.xlsx`, `.xlsm`) |
| `xlrd` | Engine for reading legacy Excel files (`.xls`) |
| `requests` | Sends your data and questions to the Ollama AI engine |
| `numpy` | Used for clean number formatting in the data summary |

> **Having trouble with pip?** See the FAQ below for the recommended fix.

---

### Step 4 — Start the Ollama AI Engine

In a terminal, run:

```bash
ollama serve
```

Leave this terminal window open. Ollama needs to be running in the background while you use the app.

> On macOS, Ollama may already be running as a system service after installation — you can skip this step if you see the Ollama icon in your menu bar.

---

### Step 5 — Start the Python Backend

Open a **new terminal window**, activate your environment, then start the server:

```bash
conda activate excel-ai
python server.py
```

You should see output like this:

```
============================================================
  Excel AI Analyzer — Backend
  Ollama URL    : http://localhost:11434
  Default model : qwen2.5
  Server        : http://localhost:3000
============================================================
```

Leave this terminal open. The backend server must keep running while you use the app.

---

### Step 6 — Open the Frontend

Open the file `excel-ai-analyzer.html` directly in your web browser (Chrome or Firefox recommended). You can double-click it in your file explorer or drag it into a browser window.

> No web server is needed for the frontend — it runs directly from your local file system.

---

### Step 7 — Use the App

1. Click **Upload File** and select an Excel (`.xlsx`, `.xls`, `.xlsm`) or CSV (`.csv`) file
2. Wait for the file to be processed — you will see a confirmation with the sheet name and row count
3. Type a question in the chat box, for example:
   - *"What are the column names?"*
   - *"Summarize the data for me."*
   - *"Which row has the highest value in column Sales?"*
4. Press **Enter** or click **Send** and the AI will respond based on your data

---

## ⚙️ How It Works (Under the Hood)

```
Your Excel File
      ↓
  server.py reads it with Pandas
      ↓
  Converts to structured plain text (column stats, data sample)
      ↓
  Text is sent to Ollama (running locally on your machine)
      ↓
  Ollama's AI model answers your question
      ↓
  Answer streams back to the browser
```

Your data stays on your machine at every step.

---

## 🔧 Configuration (Optional)

You can change the default AI model or Ollama URL by setting environment variables before running the server:

```bash
# Use a different Ollama model
export OLLAMA_MODEL=llama3.2

# Use Ollama running on a different machine
export OLLAMA_BASE_URL=http://192.168.1.100:11434

python server.py
```

On Windows (Command Prompt), use `set` instead of `export`:

```cmd
set OLLAMA_MODEL=llama3.2
python server.py
```

---

## 📌 Supported File Formats

| Format | Extension |
|---|---|
| Excel (modern) | `.xlsx`, `.xlsm` |
| Excel (legacy) | `.xls` |
| CSV | `.csv` |

---

## 🛑 Stopping the App

To stop the backend server, go to its terminal window and press `Ctrl + C`.

To deactivate the Conda environment when you are done:

```bash
conda deactivate
```

---

## ❓ Frequently Asked Questions

See [FAQ.md](FAQ.md) for troubleshooting help.

---

## 📄 License

This project is released for educational purposes. Feel free to modify and share it.
