# ❓ FAQ — Excel AI Analyzer

---

### Q: The app isn't responding. What should I check first?

Make sure all three components are running at the same time:

1. **Ollama** — run `ollama serve` in a terminal (or check for the Ollama icon in your Mac menu bar)
2. **Python backend** — run `python server.py` in a terminal with your `excel-ai` conda environment activated
3. **Frontend** — open `excel-ai-analyzer.html` in your browser

All three must be running simultaneously for the app to work.

---

### Q: I see a "Cannot reach Ollama" error in the chat

This means the Python backend cannot connect to Ollama. Try the following:

1. Open a terminal and run `ollama serve`
2. Refresh the browser page
3. Try uploading your file again and asking your question

---

### Q: I get a port conflict error when starting `server.py`

The backend runs on **port 3000** by default. If you already have another application using port 3000, you will see an error like:

```
OSError: [Errno 48] Address already in use
```

**Fix:** Open `server.py` in a text editor and scroll to the very last line. Change the port number:

```python
app.run(host="0.0.0.0", port=3000, debug=False)
```

Change `3000` to `5001` (or any unused port number):

```python
app.run(host="0.0.0.0", port=5001, debug=False)
```

Then open `excel-ai-analyzer.html` in a text editor and search for `localhost:3000`. Replace every occurrence with `localhost:5001`. Save the file and restart the server.

---

### Q: What is the difference between pip and conda? Why does it matter?

Both `pip` and `conda` are package installers for Python, but they work differently and can sometimes conflict with each other.

- **pip** is Python's built-in installer. It installs packages from the PyPI repository (Python Package Index).
- **conda** is Anaconda's installer. It manages both Python packages and system-level dependencies, and is better at avoiding conflicts.

When you are inside a Conda environment, mixing pip and conda installs can sometimes cause problems — packages installed by one may overwrite or conflict with packages installed by the other. The safest approach is to **install as much as possible through conda first**, then use pip only for packages that conda does not carry.

---

### Q: `pip install flask` fails or gives an error

This is one of the most common issues when using Conda. There are a few variations of this problem:

**Symptom 1 — Permission error:**
```
ERROR: Could not install packages due to an OSError
```

**Symptom 2 — pip installs to the wrong Python:**
Flask installs without error but `import flask` still fails when you run the server. This usually means pip installed Flask into your system Python instead of your conda environment.

**Symptom 3 — Dependency conflict:**
```
ERROR: pip's dependency resolver does not currently take into account all the packages that are installed.
```

**The recommended fix for all of the above** is to install Flask through conda instead of pip:

```bash
conda install -c conda-forge flask flask-cors -y
```

Then install the remaining packages with pip (these are less likely to conflict):

```bash
pip install pandas openpyxl xlrd requests numpy
```

---

### Q: I ran `pip install flask` but Python still says `ModuleNotFoundError: No module named 'flask'`

This almost always means pip installed Flask into a different Python than the one your conda environment is using. Here is how to check and fix it:

**Check which Python you are using:**
```bash
which python        # Mac/Linux
where python        # Windows
```

The path should contain `excel-ai` (your environment name), for example:
```
/Users/yourname/miniconda3/envs/excel-ai/bin/python
```

If it shows a system path instead (e.g. `/usr/bin/python`), your environment is not activated. Run:
```bash
conda activate excel-ai
```

Then reinstall Flask through conda:
```bash
conda install -c conda-forge flask flask-cors -y
```

---

### Q: How do I do a clean reinstall if nothing else works?

If your environment is in a broken state, the safest option is to delete it and start fresh. This takes about 5 minutes and almost always resolves pip/conda conflicts:

```bash
# Step 1 — Exit the broken environment
conda deactivate

# Step 2 — Delete it completely
conda remove -n excel-ai --all -y

# Step 3 — Create a fresh one
conda create -n excel-ai python=3.11 -y
conda activate excel-ai

# Step 4 — Install Flask through conda first (avoids the most common conflict)
conda install -c conda-forge flask flask-cors -y

# Step 5 — Install the rest through pip
pip install pandas openpyxl xlrd requests numpy

# Step 6 — Verify everything is in place
python -c "import flask, pandas, openpyxl, xlrd, requests, numpy; print('All packages OK')"
```

If the last line prints `All packages OK`, you are ready to run `python server.py`.

---

### Q: I forgot to activate my Conda environment and got import errors

You will see errors like `ModuleNotFoundError: No module named 'flask'` if the environment is not active. Always run this before starting the server:

```bash
conda activate excel-ai
```

Then run `python server.py` again.

---

### Q: The model is taking a very long time to respond

The first response after starting Ollama is usually the slowest because the model needs to load into memory. Subsequent questions in the same session will be faster. Response time also depends on:

- The size of your spreadsheet (more rows = more text for the AI to read)
- The speed of your computer's CPU or GPU
- Which Ollama model you are using (`qwen2.5` is a good balance of speed and quality)

If responses are consistently too slow, try a smaller and faster model:

```bash
ollama pull llama3.2:3b
```

Then set it as the default before starting the server:

```bash
export OLLAMA_MODEL=llama3.2:3b
python server.py
```

---

### Q: My Excel file uploaded but the AI gives wrong or odd answers

The AI reads a **text summary** of your data — column names, statistics, and a sample of the first 100 rows. Keep in mind:

- For very large files, the AI only sees the first 100 rows as a sample (though statistics like sum and average are calculated across all rows)
- Merged cells and heavily formatted Excel files may not parse perfectly
- Try asking more specific questions, for example: *"What is the sum of the Revenue column?"* rather than *"Tell me about this file."*

---

### Q: Can I use a different AI model?

Yes. Any model available in Ollama will work. First pull the model you want:

```bash
ollama pull mistral
```

Then start the server with that model:

```bash
export OLLAMA_MODEL=mistral
python server.py
```

You can also switch models directly in the app's dropdown menu in the top bar of the browser interface.

---

### Q: Is my data safe? Does it get sent to the internet?

No. Everything runs locally on your machine. Your spreadsheet data is:

1. Read by the Python backend running on your computer
2. Converted to text on your computer
3. Sent to Ollama, which is also running on your computer

No data is transmitted to any external server or cloud API. This is the primary design goal of this project.

---

### Q: Does this work on Windows?

Yes. The setup is the same, with a few small differences:

- Use **Anaconda Prompt** instead of Terminal for running conda commands
- Use `set` instead of `export` for environment variables:
  ```cmd
  set OLLAMA_MODEL=llama3.2
  python server.py
  ```
- To open the HTML file, just double-click it in File Explorer

---

### Q: The browser page loads but shows no status or the status indicator is red

The status indicator in the top-right corner of the app reflects whether the backend server is reachable. A red indicator means the browser cannot connect to `http://localhost:3000`.

Check that:
- `server.py` is still running in your terminal (it should show no errors)
- You have not accidentally closed the terminal window running the server
- Your browser is not blocking `localhost` connections (try Chrome or Firefox)

---

### Q: How do I completely uninstall everything?

```bash
# Remove the conda environment
conda deactivate
conda remove -n excel-ai --all -y

# Remove the Ollama model (optional, frees disk space)
ollama rm qwen2.5

# Uninstall Ollama
# Mac: drag Ollama from Applications to Trash
# Windows: uninstall via Control Panel
# Linux: rm -rf ~/.ollama
```
