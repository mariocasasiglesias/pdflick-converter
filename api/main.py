"""
API de conversión de PDFlick.

Solo hace las conversiones que no se pueden hacer bien en el navegador:
  POST /v1/pdf-to-word   PDF -> Word (.docx) de alta calidad, con pdf2docx.
  POST /v1/pdf-to-pptx   PDF -> PowerPoint (.pptx) con el texto editable (convert.py).
  POST /v1/pdf-to-xlsx   PDF -> Excel (.xlsx), con las tablas de cada página (convert.py).
  POST /v1/office-to-pdf Word, Excel, PowerPoint u OpenDocument -> PDF, con Gotenberg (LibreOffice).

Principios:
  - No se guarda nada: cada archivo se procesa en una carpeta temporal que se borra al terminar.
  - No se registra el contenido ni el nombre de los archivos, solo ruta, estado y duración.
  - Cada conversión corre en un proceso aparte con límite de tiempo: si se cuelga, se mata.
  - Tamaño, páginas y conversiones simultáneas limitados para no afectar a otros servicios.

Este servicio usa PyMuPDF (AGPL-3.0), así que se distribuye bajo AGPL-3.0: su código fuente
está disponible en la dirección que devuelve GET /source.
"""

import asyncio
import logging
import multiprocessing as mp
import os
import re
import tempfile
import time

import httpx
from fastapi import FastAPI, File, Form, HTTPException, Request, UploadFile
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse, Response

from convert import UserError, pdf_to_pptx, pdf_to_xlsx

ALLOWED_ORIGINS = [o.strip() for o in os.environ.get('ALLOWED_ORIGINS', 'https://pdflick.app').split(',') if o.strip()]
MAX_BYTES = int(os.environ.get('MAX_UPLOAD_MB', '25')) * 1024 * 1024
MAX_PAGES = int(os.environ.get('MAX_PAGES', '150'))
TIMEOUT = int(os.environ.get('CONVERT_TIMEOUT', '120'))
GOTENBERG_URL = os.environ.get('GOTENBERG_URL', 'http://gotenberg:3000')
SOURCE_URL = os.environ.get('SOURCE_URL', '')
SLOTS = asyncio.Semaphore(int(os.environ.get('CONCURRENCY', '2')))

DOCX = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
PPTX = 'application/vnd.openxmlformats-officedocument.presentationml.presentation'
XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
OFFICE_EXT = {'.doc', '.docx', '.odt', '.rtf', '.xls', '.xlsx', '.ods', '.csv', '.ppt', '.pptx', '.odp'}

logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s %(message)s')
log = logging.getLogger('pdflick-api')

app = FastAPI(docs_url=None, redoc_url=None, openapi_url=None)
app.add_middleware(
    CORSMiddleware,
    allow_origins=ALLOWED_ORIGINS,
    allow_methods=['GET', 'POST'],
    allow_headers=['*'],
    max_age=86400,
)


@app.middleware('http')
async def access_log(request: Request, call_next):
    # Registro mínimo: sin nombres de archivo, sin IP, sin contenido.
    start = time.monotonic()
    response = await call_next(request)
    log.info('%s %s %s %.1fs', request.method, request.url.path, response.status_code, time.monotonic() - start)
    return response


def check_origin(request: Request):
    """Solo PDFlick puede usar la API desde el navegador (además de CORS, que no frena a otros clientes)."""
    if request.headers.get('origin') not in ALLOWED_ORIGINS:
        raise HTTPException(403, 'Origen no permitido.')


async def read_limited(file: UploadFile) -> bytes:
    data = await file.read(MAX_BYTES + 1)
    if len(data) > MAX_BYTES:
        raise HTTPException(413, f'El archivo supera el máximo de {MAX_BYTES // (1024 * 1024)} MB.')
    if not data:
        raise HTTPException(400, 'El archivo está vacío.')
    return data


# ---------- Conversión aislada en otro proceso ----------

def _pdf_to_docx(data: bytes) -> bytes:
    import pymupdf
    from pdf2docx import Converter

    with tempfile.TemporaryDirectory() as tmp:
        src = os.path.join(tmp, 'entrada.pdf')
        dst = os.path.join(tmp, 'salida.docx')
        with open(src, 'wb') as f:
            f.write(data)
        with pymupdf.open(src) as doc:
            if doc.needs_pass:
                raise UserError('El PDF está protegido con contraseña. Desbloquéalo antes de convertirlo.')
            if doc.page_count > MAX_PAGES:
                raise UserError(f'El PDF tiene {doc.page_count} páginas; el máximo es {MAX_PAGES}.')
        cv = Converter(src)
        try:
            cv.convert(dst, multi_processing=False)
        finally:
            cv.close()
        with open(dst, 'rb') as f:
            return f.read()


def _worker(conn, fn, args):
    try:
        conn.send(('ok', fn(*args)))
    except UserError as e:
        conn.send(('user', str(e)))
    except Exception as e:  # noqa: BLE001 - cualquier fallo de la librería se devuelve como error genérico
        conn.send(('error', type(e).__name__))
    finally:
        conn.close()


# «forkserver» crea cada proceso desde un proceso limpio (no desde el servidor, que tiene hilos),
# con las librerías de conversión ya cargadas para que cada conversión arranque rápido.
CTX = mp.get_context('forkserver')
CTX.set_forkserver_preload(['pymupdf', 'pdf2docx', 'pptx', 'openpyxl', 'convert'])


def _run_isolated(fn, args: tuple) -> bytes:
    parent, child = CTX.Pipe(duplex=False)
    proc = CTX.Process(target=_worker, args=(child, fn, args), daemon=True)
    proc.start()
    child.close()
    try:
        if not parent.poll(TIMEOUT):
            raise UserError(f'La conversión ha tardado más de {TIMEOUT} s y se ha cancelado. Prueba con menos páginas.')
        kind, value = parent.recv()
    except EOFError:
        # El proceso murió sin responder (normalmente por falta de memoria).
        raise UserError('El documento es demasiado pesado para convertirlo. Prueba con menos páginas.')
    finally:
        parent.close()
        if proc.is_alive():
            proc.kill()
        proc.join(5)
    if kind == 'ok':
        return value
    if kind == 'user':
        raise UserError(value)
    log.warning('fallo de conversión: %s', value)
    raise HTTPException(500, 'No se ha podido convertir este PDF.')


async def run_isolated(fn, *args) -> bytes:
    async with SLOTS:
        try:
            return await asyncio.get_running_loop().run_in_executor(None, _run_isolated, fn, args)
        except UserError as e:
            raise HTTPException(422, str(e))


# ---------- Rutas ----------

@app.get('/health')
async def health():
    return {'ok': True}


@app.get('/source')
async def source():
    return {'license': 'AGPL-3.0', 'source': SOURCE_URL}


@app.post('/v1/pdf-to-word')
async def pdf_to_word(request: Request, file: UploadFile = File(...)):
    check_origin(request)
    data = await read_limited(file)
    if not data.startswith(b'%PDF'):
        raise HTTPException(415, 'El archivo no es un PDF.')
    docx = await run_isolated(_pdf_to_docx, data)
    return Response(docx, media_type=DOCX, headers={'Cache-Control': 'no-store'})


async def read_pdf(request: Request, file: UploadFile) -> bytes:
    check_origin(request)
    data = await read_limited(file)
    if not data.startswith(b'%PDF'):
        raise HTTPException(415, 'El archivo no es un PDF.')
    return data


@app.post('/v1/pdf-to-pptx')
async def pdf_to_powerpoint(request: Request, file: UploadFile = File(...), mode: str = Form('editable')):
    if mode not in ('editable', 'image'):
        raise HTTPException(400, 'Modo no válido.')
    data = await read_pdf(request, file)
    pptx = await run_isolated(pdf_to_pptx, data, mode, MAX_PAGES)
    return Response(pptx, media_type=PPTX, headers={'Cache-Control': 'no-store'})


@app.post('/v1/pdf-to-xlsx')
async def pdf_to_excel(request: Request, file: UploadFile = File(...), lang: str = Form('es')):
    data = await read_pdf(request, file)
    xlsx = await run_isolated(pdf_to_xlsx, data, 'en' if lang == 'en' else 'es', MAX_PAGES)
    return Response(xlsx, media_type=XLSX, headers={'Cache-Control': 'no-store'})


@app.post('/v1/office-to-pdf')
async def office_to_pdf(request: Request, file: UploadFile = File(...)):
    check_origin(request)
    ext = os.path.splitext(file.filename or '')[1].lower()
    if ext not in OFFICE_EXT:
        raise HTTPException(415, 'Formato no admitido. Usa Word, Excel, PowerPoint u OpenDocument.')
    data = await read_limited(file)
    # A LibreOffice solo le llega un nombre genérico con la extensión correcta.
    name = 'documento' + re.sub(r'[^a-z.]', '', ext)
    async with SLOTS:
        try:
            async with httpx.AsyncClient(timeout=TIMEOUT) as client:
                r = await client.post(f'{GOTENBERG_URL}/forms/libreoffice/convert', files={'files': (name, data)})
        except httpx.HTTPError:
            raise HTTPException(503, 'El servicio de conversión no está disponible ahora mismo.')
    if r.status_code != 200:
        log.warning('gotenberg respondió %s', r.status_code)
        raise HTTPException(422, 'No se ha podido convertir este documento.')
    return Response(r.content, media_type='application/pdf', headers={'Cache-Control': 'no-store'})


@app.exception_handler(HTTPException)
async def http_error(_: Request, exc: HTTPException):
    # Mensaje en español en un campo fijo, para mostrarlo tal cual en PDFlick.
    return JSONResponse({'error': exc.detail}, status_code=exc.status_code, headers={'Cache-Control': 'no-store'})
