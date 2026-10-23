# API de conversión de PDFlick

Servicio pequeño para las conversiones que no se pueden hacer bien en el navegador:

| Ruta | Qué hace | Motor |
|---|---|---|
| `POST /v1/pdf-to-word` | PDF → Word (.docx) de alta calidad: tablas, imágenes, columnas | [pdf2docx](https://github.com/ArtifexSoftware/pdf2docx) + PyMuPDF |
| `POST /v1/pdf-to-pptx` | PDF → PowerPoint (.pptx) con el texto editable (`mode=editable`) o como imagen exacta (`mode=image`) | PyMuPDF + [python-pptx](https://github.com/scanny/python-pptx) |
| `POST /v1/pdf-to-xlsx` | PDF → Excel (.xlsx): una hoja por página, con sus tablas (`lang=es` o `en` para leer los números) | PyMuPDF + [openpyxl](https://openpyxl.readthedocs.io) |
| `POST /v1/office-to-pdf` | Word, Excel, PowerPoint, OpenDocument → PDF | [Gotenberg](https://gotenberg.dev) (LibreOffice) |
| `GET /health` | Comprobación de que está activo | |
| `GET /source` | Dónde está el código (licencia AGPL) | |

**Privacidad:** no guarda archivos (todo pasa por `/tmp`, que está en memoria, y se borra al terminar), no registra nombres ni contenido, y nginx no guarda registros de acceso.

**Licencia:** [AGPL-3.0](LICENSE), porque PyMuPDF lo es. El código que se ejecuta en https://api.pdflick.app se publica en https://github.com/mariocasasiglesias/pdflick-converter (dirección que la API devuelve en `GET /source`). La web de PDFlick solo usa esta API por HTTP y no se ve afectada por la licencia.

---

## Instalación en el VPS (Ubuntu 24.04)

Todo como `root`. Pensado para convivir con otras webs del mismo servidor sin tocarlas. Antes de empezar, comprueba que el puerto 8787 está libre (`ss -ltnp | grep :8787` no debe mostrar nada).

### 1. Memoria de intercambio (2 GB)
Evita que una conversión muy grande deje al sistema sin memoria.
```bash
fallocate -l 2G /swapfile && chmod 600 /swapfile && mkswap /swapfile && swapon /swapfile
echo '/swapfile none swap sw 0 0' >> /etc/fstab
```

### 2. Docker
```bash
apt-get update && apt-get install -y ca-certificates curl
install -m 0755 -d /etc/apt/keyrings
curl -fsSL https://download.docker.com/linux/ubuntu/gpg -o /etc/apt/keyrings/docker.asc
echo "deb [arch=$(dpkg --print-architecture) signed-by=/etc/apt/keyrings/docker.asc] https://download.docker.com/linux/ubuntu $(. /etc/os-release && echo $VERSION_CODENAME) stable" > /etc/apt/sources.list.d/docker.list
apt-get update && apt-get install -y docker-ce docker-ce-cli containerd.io docker-compose-plugin
```

### 3. Copiar el código y arrancar
En el VPS, descárgalo del repositorio público:
```bash
git clone https://github.com/mariocasasiglesias/pdflick-converter /opt/pdflick-api
```
O cópialo desde tu ordenador (en el repositorio de PDFlick es la carpeta `server/`), con la IP de tu VPS:
```bash
scp -r server root@IP_DEL_VPS:/opt/pdflick-api
```
En el VPS:
```bash
cd /opt/pdflick-api && docker compose up -d --build
```
La primera vez tarda unos minutos (descarga Python, LibreOffice y las librerías). Comprueba:
```bash
docker compose ps
curl -s http://127.0.0.1:8787/health          # {"ok":true}
```

### 4. DNS en Cloudflare
En Cloudflare › pdflick.app › **DNS** › **Add record**:
- Tipo **A**, nombre **api**, contenido **IP del VPS**, proxy **activado** (nube naranja).

### 5. Certificado HTTPS (certificado de origen de Cloudflare)
1. Cloudflare › pdflick.app › **SSL/TLS** › **Origin Server** › **Create Certificate** (deja los valores por defecto, 15 años).
2. En el VPS, guarda el certificado y la clave que te muestra:
   ```bash
   mkdir -p /etc/ssl/cloudflare && chmod 700 /etc/ssl/cloudflare
   nano /etc/ssl/cloudflare/pdflick-origin.pem    # pega el certificado
   nano /etc/ssl/cloudflare/pdflick-origin.key    # pega la clave privada
   chmod 600 /etc/ssl/cloudflare/pdflick-origin.key
   ```
3. Cloudflare › **SSL/TLS** › **Overview**: modo **Full (strict)**.

> La clave privada no se comparte con nadie ni se sube a ningún repositorio.

### 6. nginx
```bash
cp /opt/pdflick-api/nginx/cloudflare-realip.conf /etc/nginx/snippets/
cp /opt/pdflick-api/nginx/api.pdflick.app.conf /etc/nginx/sites-available/
ln -s /etc/nginx/sites-available/api.pdflick.app.conf /etc/nginx/sites-enabled/
nginx -t && systemctl reload nginx
```
`nginx -t` debe decir «syntax is ok» antes de recargar; si no, no se toca la configuración que ya funciona.

### 7. Comprobación final
Desde cualquier ordenador:
```bash
curl -s https://api.pdflick.app/health
curl -s -o /dev/null -w "%{http_code}\n" -X POST https://api.pdflick.app/v1/pdf-to-word   # 403: sin origen permitido, correcto
```

---

## Mantenimiento

| Tarea | Comando (en `/opt/pdflick-api`) |
|---|---|
| Ver estado y memoria | `docker compose ps` y `docker stats --no-stream` |
| Ver registros | `docker compose logs --tail 100 api` |
| Actualizar tras cambios | `docker compose up -d --build` |
| Parar | `docker compose down` |

Límites (en `docker-compose.yml`): cada servicio usa como máximo 2 CPU y 2 GB de RAM; hasta 2 conversiones a la vez; archivos de hasta 25 MB y 150 páginas; 120 s por conversión. nginx limita a 10 conversiones por minuto y persona.
