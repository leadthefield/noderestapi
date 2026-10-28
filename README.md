<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Word Document Merger</title>
  <script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.0/Sortable.min.js"></script>
  <style>
    body { font-family: sans-serif; padding: 24px; max-width: 600px; margin: auto; }
    #drop-zone { border: 2px dashed #ccc; padding: 20px; text-align: center; cursor: pointer; }
    #file-list { list-style: none; padding: 0; margin-top: 16px; }
    .file-item { display: flex; align-items: center; justify-content: space-between; padding: 10px; margin-bottom: 8px; background: #f4f4f4; border: 1px solid #ddd; cursor: grab; }
    .file-item.disabled { opacity: 0.5; }
    .actions { margin-top: 20px; }
    button { padding: 10px 18px; cursor: pointer; background: #0066cc; color: #fff; border: none; border-radius: 4px; }
    button:disabled { background: #999; }
  </style>
</head>
<body>
  <h2>Word Document Assembler</h2>
  <div id="drop-zone" onclick="document.getElementById('file-input').click()">
    Drop .docx files here or click to browse
    <input type="file" id="file-input" multiple accept=".docx" style="display:none" onchange="handleFiles(this.files)">
  </div>
  <ul id="file-list"></ul>
  <div class="actions">
    <button id="merge-btn" onclick="executeMerge()" disabled>Merge & Download</button>
  </div>

  <script>
    let files = [];
    const listEl = document.getElementById('file-list');
    const mergeBtn = document.getElementById('merge-btn');

    new Sortable(listEl, {
      animation: 150,
      onEnd: (evt) => {
        const movedItem = files.splice(evt.oldIndex, 1)[0];
        files.splice(evt.newIndex, 0, movedItem);
      }
    });

    function handleFiles(uploaded) {
      for (const f of uploaded) {
        if (!f.name.endsWith('.docx')) continue;
        files.push({ id: crypto.randomUUID(), file: f, enabled: true });
      }
      render();
    }

    function render() {
      listEl.innerHTML = '';
      files.forEach((item, index) => {
        const li = document.createElement('li');
        li.className = `file-item ${item.enabled ? '' : 'disabled'}`;
        li.innerHTML = `
          <div>
            <input type="checkbox" ${item.enabled ? 'checked' : ''} onchange="toggle(${index})">
            <span>${item.file.name}</span>
          </div>
          <button onclick="removeFile(${index})" style="background:#cc0000;padding:4px 8px;">✕</button>
        `;
        listEl.appendChild(li);
      });
      mergeBtn.disabled = files.filter(f => f.enabled).length < 2;
    }

    function toggle(index) {
      files[index].enabled = !files[index].enabled;
      render();
    }

    function removeFile(index) {
      files.splice(index, 1);
      render();
    }

    async function executeMerge() {
      mergeBtn.disabled = true;
      mergeBtn.innerText = "Processing...";
      const activeFiles = files.filter(f => f.enabled).map(f => f.file);

      const formData = new FormData();
      activeFiles.forEach((file) => formData.append('documents', file));

      const res = await fetch('/api/merge', { method: 'POST', body: formData });
      if (!res.ok) throw new Error('Merge failed');

      const blob = await res.blob();
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = "merged_output.docx";
      a.click();
      URL.revokeObjectURL(url);

      mergeBtn.innerText = "Merge & Download";
      mergeBtn.disabled = false;
    }
  </script>
</body>
</html>


import io
from fastapi import FastAPI, UploadFile, File
from fastapi.responses import StreamingResponse
from docx import Document
from docxcompose.composer import Composer

app = FastAPI()

@app.post("/api/merge")
async def merge_docx(documents: list[UploadFile] = File(...)):
    if not documents:
        return {"error": "No files provided"}

    first_stream = io.BytesIO(await documents[0].read())
    base_doc = Document(first_stream)
    composer = Composer(base_doc)

    for upload in documents[1:]:
        sub_stream = io.BytesIO(await upload.read())
        sub_doc = Document(sub_stream)
        composer.append(sub_doc)

    out_stream = io.BytesIO()
    composer.save(out_stream)
    out_stream.seek(0)

    return StreamingResponse(
        out_stream,
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        headers={"Content-Disposition": "attachment; filename=merged_output.docx"}
    )
```Alright. Let me know if you need any other details for your project.
