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
```