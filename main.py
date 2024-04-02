import io
import os
import zipfile
from typing import Union

from fastapi import FastAPI, File, UploadFile
from fastapi.params import Query
from starlette.responses import JSONResponse, StreamingResponse
from traits.trait_types import Enum

from utils import pdf_to_jpg, pdf_to_images, pdf_to_word, image_to_pdf, word_to_pdf

app = FastAPI()


# From Word, PNG, JPG to PDF

@app.post("/convert/pdf", response_model=dict)
async def convert_pdf(file: UploadFile = File(...), q: str = Query("word", enum=["word", "png", "jpg"])):
    image_output_folder = "images"
    word_output_folder = 'words'
    if not os.path.exists(image_output_folder):
        os.makedirs(image_output_folder)
    if not os.path.exists(word_output_folder):
        os.makedirs(word_output_folder)
    if not q:
        return {'error': 'Please specify file type to convert'}
    pdf_bytes = await file.read()
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        if q == 'jpg':
            paths = pdf_to_jpg(pdf_bytes, image_output_folder)
            for path in paths:
                zipf.write(path, os.path.basename(path))
        elif q == 'png':
            paths = pdf_to_images(pdf_bytes, image_output_folder)
            for path in paths:
                zipf.write(path, os.path.basename(path))
        elif q == 'word':
            paths = pdf_to_word(pdf_bytes, word_output_folder + '/' + file.filename.split(".")[0] + ".docx")
            for path in paths:
                zipf.write(path, os.path.basename(path))
    zip_buffer.seek(0)

    return StreamingResponse(zip_buffer, media_type="application/zip",
                             headers={"Content-Disposition": "attachment; filename=converted_files.zip"})


@app.post("/convert-to/pdf", response_model=dict)
async def convert_to_pdf(file: UploadFile = File(...), q: str = Query("word", enum=["word", "png", "jpg"])):
    image_output_folder = "images"
    word_output_folder = 'words'
    if not os.path.exists(image_output_folder):
        os.makedirs(image_output_folder)
    if not os.path.exists(word_output_folder):
        os.makedirs(word_output_folder)
    if not q:
        return {'error': 'Please specify file type to convert'}
    pdf_bytes = await file.read()
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w') as zipf:
        if q == 'jpg':
            paths = image_to_pdf(pdf_bytes, 'output_jpg_to_pdf.pdf')
            for path in paths:
                zipf.write(path, os.path.basename(path))
        elif q == 'png':
            paths = image_to_pdf(pdf_bytes, image_output_folder)
            for path in paths:
                zipf.write(path, os.path.basename(path))
        elif q == 'word':
            paths = word_to_pdf(pdf_bytes, word_output_folder)
            for path in paths:
                zipf.write(path, os.path.basename(path))
    zip_buffer.seek(0)

    return StreamingResponse(zip_buffer, media_type="application/zip",
                             headers={"Content-Disposition": "attachment; filename=converted_files.zip"})

