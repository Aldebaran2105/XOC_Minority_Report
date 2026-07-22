# XOC Minority Report - Guía de integración al API real

Este módulo es el POC para generar **Minority Report XOC**, que es el reporte orientado a cliente.

La idea productiva no es que el usuario escriba el informe manualmente. El flujo esperado es:

```text
Frontend cliente
  -> API principal consulta BD real del tenant
  -> API arma snapshot estructurado
  -> Azure AI Foundry / Azure OpenAI genera JSON del reporte
  -> Python genera DOCX usando plantilla Word corporativa
  -> API devuelve/almacena el DOCX
```

La IA **no debe generar el DOCX directamente**. Azure solo debe generar contenido estructurado en JSON. El Word final lo genera Python con `python-docx`.

## Estado actual del POC

Actualmente el POC ya hace lo siguiente:

- Tiene una sección de mini-front para **Minority Report cliente**.
- Permite elegir una empresa mock.
- Lee una BD mock local.
- Genera gráficos con `matplotlib`.
- Envía el snapshot de datos a Azure.
- Azure devuelve JSON estructurado.
- Python genera el DOCX usando la portada de la plantilla actual.
- El cuerpo del informe intenta seguir el estilo del Word ejemplo corporativo.
- El archivo final se guarda en `XOC_Minority_Report/output/`.

Importante: el botón de Minority Report ahora está configurado para usar Azure obligatoriamente. No usa fallback local silencioso.

## Archivos Word incluidos

Dentro de esta carpeta hay dos documentos Word importantes:

```text
Plantilla Minority Report.docx
TXDX-WR-2606-26083-JOCKEY MINORITY_REPORT_SEMANAL_26_JUNIO.docx
```

### `Plantilla Minority Report.docx`

Es la plantilla base actual que se usa para la portada.

El generador toma este documento como base inicial porque contiene la portada que debe mantenerse.

La portada contiene placeholders como:

```text
Change for prepared for
Change for date
TXDXSECURE
```

El script reemplaza esos valores con:

- nombre del cliente/tenant;
- periodo del reporte;
- prepared by.

### `TXDX-WR-2606-26083-JOCKEY MINORITY_REPORT_SEMANAL_26_JUNIO.docx`

Es el documento ejemplo real de la empresa.

No se usa para copiar datos de Jockey. Se usa como referencia de estilo/estructura.

El objetivo visual es que el contenido generado se parezca a este Word:

- títulos tipo `Heading 1`, `Heading 2`, `Heading 3`;
- texto justificado;
- listas estilo `List Paragraph`;
- tablas con encabezado azul y letras blancas;
- secciones como `Datos generales`, `Resumen ejecutivo del dominio`, `Seguridad por Dominio`, etc.

## Estructura de carpeta

```text
XOC_Minority_Report/
├── Plantilla Minority Report.docx
├── TXDX-WR-2606-26083-JOCKEY MINORITY_REPORT_SEMANAL_26_JUNIO.docx
├── azure_minority_client.py
├── generate_minority_report.py
├── minority_mock_data.py
├── requirements.txt
├── .env.example
├── .gitignore
├── README.md
├── mock_db/
│   ├── minority_mock_db.json
│   └── minority_report_history.json
└── output/
```

Notas:

- `.env` puede existir localmente, pero está ignorado por git.
- `output/` está ignorado por git.
- `mock_db/minority_report_history.json` está ignorado por git porque contiene historial local y rutas locales.
- `mock_db/minority_mock_db.json` es la BD mock que simula data real para demo.

## Dependencias

Instalación local:

```powershell
cd .\XOC_Minority_Report
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

Dependencias principales:

```text
openai
azure-identity
python-docx
python-dotenv
lxml
Pillow
matplotlib
```

## Variables de entorno

El código carga variables desde:

```text
XOC_Minority_Report/.env
```

Si el `.env` no existe, entonces busca variables ya definidas en el entorno del sistema/servidor.

No depende de `prueba-script-xoc`. Ese fallback fue eliminado.

Variables esperadas:

```env
USE_AZURE_FOUNDRY=true
AZURE_FOUNDRY_PROJECT_ENDPOINT=https://<recurso>.services.ai.azure.com/api/projects/<project>
AZURE_FOUNDRY_OPENAI_ENDPOINT=https://<recurso>.services.ai.azure.com/openai/v1
AZURE_FOUNDRY_MODEL_DEPLOYMENT=gpt-5-mini
AZURE_FOUNDRY_API_KEY=<NO_COMMIT>
REPORT_MAX_IMAGE_MB=10
MINORITY_MAX_OUTPUT_TOKENS=9000
MINORITY_JSON_SCHEMA=true
```

En producción se recomienda no usar `.env` plano. Usar variables seguras del entorno, Key Vault, secret manager o identidad administrada si aplica.

## Flujo actual desde mini-front

El mini-front está fuera de esta carpeta:

```text
../mini-front/
```

Endpoint relevante:

```text
POST /api/minority/generate
```

Archivo donde vive el endpoint:

```text
mini-front/server.py
```

Función relevante:

```python
MiniFrontHandler._handle_minority_generate()
```

Flujo interno actual:

```text
1. Recibe companyKey desde el frontend.
2. Llama build_mock_minority_report(company_key).
3. Lee data mock de mock_db/minority_mock_db.json.
4. Genera charts con matplotlib.
5. Arma structured_data con daily_records, métricas, hallazgos, dominios y figuras.
6. Llama generate_minority_payload(... use_azure=True, allow_local_fallback=False).
7. Azure devuelve JSON.
8. Se agregan datos backend-controlados como document_code, client_name y period.
9. Python genera DOCX con _build_minority_docx().
10. Se registra historial local con register_report().
11. El mini-front devuelve downloadUrl.
```

Empresas mock disponibles:

```text
al-ink
interbank
supercel
hermanos-mario
```

El frontend muestra:

- AL ink
- Interbank
- Supercel
- Hermanos Mario

## Cómo probar localmente

Desde la raíz del workspace:

```powershell
python .\mini-front\server.py
```

Abrir:

```text
http://127.0.0.1:8765
```

En la sección **Minority Report cliente**:

1. Elegir empresa.
2. Presionar `Generar Reporte`.

Si Azure está bien configurado, se generará un DOCX en:

```text
XOC_Minority_Report/output/
```

Nombre esperado de ejemplo:

```text
TXDX-WR-2607-26092-INTERBANK_MINORITY_REPORT_MENSUAL_22_JULIO.docx
```

Si Azure no está configurado, el botón debe fallar con un error claro. Eso es intencional.

## Arquitectura de archivos Python

### `azure_minority_client.py`

Responsabilidad:

- Construir prompt para Azure.
- Enviar texto, datos estructurados e imágenes.
- Pedir JSON válido.
- Validar JSON.
- Normalizar payload.
- Guardar payload en `output/generated_minority_report_payload.json`.

Funciones principales:

```python
generate_minority_payload(...)
build_prompt(...)
_foundry_openai_client()
_normalize_payload(...)
```

Parámetros importantes de `generate_minority_payload`:

```python
generate_minority_payload(
    client_name: str,
    period: str,
    analyst_text: str = "",
    image_paths: list[str | Path] | None = None,
    image_descriptions: list[str] | None = None,
    structured_data: dict | None = None,
    reference_markdown: str = "",
    use_azure: bool | None = None,
    allow_local_fallback: bool = True,
    output_path: Path = DEFAULT_PAYLOAD_PATH,
)
```

Para el flujo cliente productivo se debe usar:

```python
use_azure=True
allow_local_fallback=False
```

Así se evita generar reportes falsos si Azure falla.

### `minority_mock_data.py`

Responsabilidad:

- Simular una BD mensual.
- Armar un snapshot de datos para una empresa.
- Generar gráficos con `matplotlib`.
- Registrar historial local de reportes.

Funciones principales:

```python
build_mock_minority_report(company_key, range_mode="month_to_date", as_of=None)
generate_charts(company_name, records, report_id)
register_report(company_key, range_mode, payload, docx_path)
company_options()
```

En producción, este archivo debe reemplazarse o adaptarse para consultar la BD/API real.

Lo importante es mantener el concepto de salida:

```python
{
  "payload": {...borrador/local/meta...},
  "charts": [
    {"path": Path(...), "description": "..."}
  ],
  "records": [...daily_records...],
  "last_snapshot": {...}
}
```

### `generate_minority_report.py`

Responsabilidad:

- Cargar plantilla Word.
- Aplicar estilo de cuerpo similar al Word ejemplo.
- Reemplazar portada.
- Limpiar contenido anterior de la plantilla.
- Crear índice TOC real.
- Construir cuerpo del reporte.
- Insertar tablas.
- Insertar imágenes/gráficos.
- Guardar y validar DOCX.

Funciones principales:

```python
base_template_path()
apply_example_body_style(document)
update_cover_and_footer(document, payload)
clear_template_body_after_cover(document)
build_report_body(document, payload)
place_evidence_images(document, images, descriptions)
validate_docx(path)
```

## Contrato JSON que debe devolver Azure

Azure debe devolver solo JSON válido. Sin Markdown, sin bloques de código, sin texto antes o después.

Estructura esperada:

```json
{
  "title": "",
  "client_name": "",
  "prepared_by": "",
  "period": "",
  "service_name": "",
  "tools": [
    {
      "name": "",
      "description": ""
    }
  ],
  "data_base": "",
  "executive_summary": "",
  "vulnerability_comparison": {
    "summary": "",
    "severity_rows": [
      {
        "severity": "",
        "previous": "",
        "current": ""
      }
    ]
  },
  "histogram_summary": "",
  "results_and_next_actions": "",
  "results_obtained": [],
  "next_actions": [],
  "requirements": [],
  "security_domains": [
    {
      "name": "",
      "summary": "",
      "findings": [
        {
          "id": "",
          "vulnerability": "",
          "affected_hosts": "",
          "severity": ""
        }
      ]
    }
  ],
  "weekly_actions": [],
  "reinforced_security": "",
  "pending_findings": [],
  "security_news": [
    {
      "title": "",
      "date": "",
      "source": "",
      "links": [],
      "summary": "",
      "recommendation": ""
    }
  ],
  "limitations": [],
  "image_citations": [
    {
      "label": "",
      "description": "",
      "used_in_sections": []
    }
  ]
}
```

El cliente `azure_minority_client.py` usa schema estricto cuando:

```env
MINORITY_JSON_SCHEMA=true
```

Si Azure devuelve campos extra o faltan campos, `_normalize_payload` falla.

## Campos agregados por backend/Python

Algunos campos no deberían depender de Azure:

- `document_code`
- `mock_meta`
- nombre de archivo final
- ruta de descarga
- historial de generación

En el POC, después de recibir el JSON de Azure, `mini-front/server.py` vuelve a anexar:

```python
payload["client_name"] = draft_payload.get("client_name")
payload["period"] = draft_payload.get("period")
payload["prepared_by"] = draft_payload.get("prepared_by") or "TXDXSECURE"
payload["document_code"] = draft_payload.get("document_code")
payload["mock_meta"] = draft_payload.get("mock_meta")
```

Esto es importante porque Azure no debe inventar nombres oficiales, códigos internos o periodos.

## Nombre del archivo DOCX

El generador usa `document_code` si existe.

Ejemplo:

```text
TXDX-WR-2607-26092-INTERBANK_MINORITY_REPORT_MENSUAL_22_JULIO.docx
```

Si no existe `document_code`, cae a un nombre informativo basado en cliente y periodo.

## Índice del Word

El DOCX genera un índice tipo TOC real de Word:

```text
TOC \o "1-3" \h \z \u
```

Además se activa:

```text
updateFields=true
```

Al abrir el DOCX en Microsoft Word, puede ser necesario actualizar campos:

```text
Ctrl + A
F9
```

o clic derecho sobre el índice y seleccionar `Actualizar campo`.

Para producción, si se quiere entregar el DOCX con índice ya renderizado, se recomienda agregar un paso con Microsoft Word/LibreOffice en servidor o un worker que abra/actualice/guarde el archivo. Python-docx crea el campo, pero no renderiza números de página.

## Tablas

Las tablas generadas actualmente tienen:

- primera fila con fondo azul `#006D9F`;
- texto blanco;
- bordes suaves;
- resto de filas sin relleno especial.

Tablas principales:

- comparación de severidades;
- hallazgos por dominio;
- noticias/referencias cuando aplica.

## Gráficos

Los gráficos se generan con `matplotlib` en:

```text
output/charts/
```

Actualmente se generan figuras como:

- tendencia de hallazgos por severidad;
- distribución actual;
- actividad de nuevos hallazgos vs remediados.

Estas imágenes se insertan en el DOCX como `Figura 1`, `Figura 2`, `Figura 3`.

En producción, los gráficos pueden seguir generándose desde Python con data real de BD. No es necesario que Azure genere imágenes.

## Cómo adaptar al API real

La adaptación recomendada es reemplazar el uso de `mock_db` por consultas reales al backend/BD.

### Paso 1: crear endpoint productivo

Ejemplo conceptual:

```text
POST /reports/minority
```

Body mínimo:

```json
{
  "tenant_id": "uuid-o-id",
  "period_mode": "month_to_date"
}
```

En un caso cliente logueado, el `tenant_id` puede venir del JWT/sesión y no del body.

### Paso 2: consultar BD real

El backend debe reunir datos como:

- tenant/cliente;
- periodo del reporte;
- activos monitoreados;
- dominios o categorías;
- hallazgos por severidad;
- cambios vs periodo anterior;
- acciones trabajadas;
- remediaciones;
- hallazgos pendientes;
- disponibilidad/eventos;
- noticias o inteligencia si existe;
- evidencias/figuras si aplica.

### Paso 3: armar snapshot estructurado

Debe construirse un objeto parecido a:

```json
{
  "source": "production_db",
  "client_name": "Cliente Real",
  "period": "Del 1 de julio al 22 de julio de 2026",
  "document_code": "TXDX-WR-2607-XXXXX-CLIENTE_MINORITY_REPORT_MENSUAL_22_JULIO",
  "tools": [],
  "aggregated_metrics": {},
  "daily_records": [],
  "security_domains": [],
  "pending_findings": [],
  "weekly_actions": [],
  "chart_evidence": []
}
```

Ese snapshot es lo que se manda a Azure en `structured_data`.

### Paso 4: generar gráficos

Con la data real, llamar una función equivalente a `generate_charts`.

La salida esperada para el generador DOCX es:

```python
chart_images = [
    {
        "path": Path(".../figura_1.png"),
        "description": "Tendencia de hallazgos abiertos por severidad."
    }
]
```

### Paso 5: llamar Azure

Usar:

```python
payload = generate_minority_payload(
    client_name=client_name,
    period=period,
    analyst_text=instruction_text,
    image_paths=[image["path"] for image in chart_images],
    image_descriptions=[image["description"] for image in chart_images],
    structured_data=structured_data,
    use_azure=True,
    allow_local_fallback=False,
    output_path=payload_output_path,
)
```

### Paso 6: anexar campos controlados por backend

Después de Azure:

```python
payload["client_name"] = client_name
payload["period"] = period
payload["prepared_by"] = "TXDXSECURE"
payload["document_code"] = document_code
```

### Paso 7: generar DOCX

La lógica equivalente a `_build_minority_docx` debe:

```python
document = Document(base_template_path())
apply_example_body_style(document)
update_cover_and_footer(document, payload)
clear_template_body_after_cover(document)
build_report_body(document, payload)
place_evidence_images(document, image_paths, image_descriptions)
document.save(result_path)
validate_docx(result_path)
```

### Paso 8: entregar respuesta

Respuesta sugerida:

```json
{
  "ok": true,
  "report_type": "minority_report",
  "tenant_id": "...",
  "document_code": "...",
  "docx_url": "...",
  "generated_at": "...",
  "period": {
    "start": "2026-07-01",
    "end": "2026-07-22"
  }
}
```

## Manejo de snapshots/reportes generados

El POC registra historial local en:

```text
mock_db/minority_report_history.json
```

En producción esto debería ir a una tabla real, por ejemplo:

```text
report_generations
```

Campos sugeridos:

```text
id
tenant_id
report_type
period_start
period_end
generated_at
generated_by
document_code
docx_storage_path
payload_storage_path
status
error_message
```

Recomendación de producto:

- El botón cliente debería generar el reporte acumulado hasta la fecha actual.
- No recortar automáticamente por el último reporte, porque si dos usuarios generan el mismo día, el segundo puede salir casi vacío.
- Guardar snapshots para auditoría y trazabilidad.
- Si se quiere un modo incremental, exponerlo como opción interna/admin, no como default cliente.

## Seguridad y privacidad

Reglas importantes:

- No loguear imágenes sensibles.
- No loguear payload completo con datos del cliente en producción.
- No subir `.env`.
- No hardcodear API keys.
- No copiar datos del Word ejemplo.
- Azure debe usar solo evidencia enviada por backend.
- Si falta información, el JSON debe incluirlo en `limitations`.
- Los archivos temporales deben limpiarse después de generar el reporte.
- Validar tamaño y extensión de imágenes.

## Qué NO debe hacer Azure

Azure no debe:

- generar el DOCX;
- calcular nombres oficiales de archivo;
- inventar fechas;
- inventar IPs;
- inventar activos;
- inventar acciones realizadas;
- inventar severidades;
- copiar datos del reporte Jockey;
- decidir rutas de almacenamiento;
- registrar auditoría.

Azure sí debe:

- resumir la data;
- redactar contenido ejecutivo/técnico;
- estructurar hallazgos;
- generar recomendaciones basadas en evidencia;
- devolver JSON válido.

## Diferencia entre POC y producción

POC actual:

```text
mini-front -> mock_db JSON -> Azure -> Python DOCX -> output local
```

Producción deseada:

```text
frontend real -> API real -> BD real -> Azure -> Python DOCX -> storage/API response
```

Lo que se reemplaza:

- `mock_db/minority_mock_db.json`
- `minority_mock_data.build_mock_minority_report`
- almacenamiento local en `output/`
- historial local JSON

Lo que se conserva:

- `azure_minority_client.generate_minority_payload`
- contrato JSON esperado;
- generación DOCX con `generate_minority_report.py`;
- plantilla Word;
- lógica de tablas, índice, portada, imágenes y validación.

## Comandos útiles

Compilar/validar sintaxis:

```powershell
python -m py_compile .\azure_minority_client.py .\generate_minority_report.py .\minority_mock_data.py
```

Levantar mini-front:

```powershell
cd ..
python .\mini-front\server.py
```

Probar generación directa local sin Azure, solo para validar DOCX:

```powershell
cd .\XOC_Minority_Report
python .\generate_minority_report.py `
  --client-name "Cliente Demo" `
  --period "Del 1 de julio al 22 de julio de 2026" `
  --text "Prueba técnica de generación DOCX." `
  --no-azure
```

Nota: ese modo directo local sirve para probar Word, no para validar el flujo cliente real.

## Checklist para integración al API real

- [ ] Definir endpoint backend para Minority Report.
- [ ] Resolver tenant desde sesión/JWT.
- [ ] Consultar BD real.
- [ ] Armar snapshot estructurado.
- [ ] Generar gráficos con data real.
- [ ] Configurar credenciales Azure en entorno seguro.
- [ ] Llamar `generate_minority_payload(... use_azure=True, allow_local_fallback=False)`.
- [ ] Anexar `document_code`, `client_name`, `period` desde backend.
- [ ] Generar DOCX con plantilla.
- [ ] Actualizar/renderizar TOC si se requiere.
- [ ] Guardar DOCX en storage.
- [ ] Registrar auditoría/snapshot.
- [ ] Devolver URL/archivo al frontend.
- [ ] Limpiar temporales.

## Advertencias actuales

- `XOC_Minority_Report/.env` puede existir localmente para pruebas, pero no debe subirse.
- `.env.example` no tiene secretos; solo documenta variables.
- El índice Word se crea como campo TOC; para verlo con páginas reales, Word debe actualizar campos.
- El POC usa empresas mock. En producción, esas empresas no deben estar hardcodeadas.
- El archivo ejemplo Jockey es referencia visual, no fuente de datos.

