"""Cliente Azure AI Foundry multimodal para Minority Report XOC.

La IA devuelve JSON. El DOCX se genera luego con Python.
"""

from __future__ import annotations

import base64
import json
import mimetypes
import os
import re
from pathlib import Path
from typing import Any

from azure.identity import DefaultAzureCredential
from dotenv import load_dotenv

ROOT = Path(__file__).resolve().parent
OUTPUT_DIR = ROOT / "output"
DEFAULT_PAYLOAD_PATH = OUTPUT_DIR / "generated_minority_report_payload.json"
SUPPORTED_IMAGE_EXTENSIONS = {".png", ".jpg", ".jpeg", ".webp"}

MINORITY_KEYS = {
    "title",
    "client_name",
    "prepared_by",
    "period",
    "service_name",
    "tools",
    "data_base",
    "executive_summary",
    "vulnerability_comparison",
    "histogram_summary",
    "results_and_next_actions",
    "results_obtained",
    "next_actions",
    "requirements",
    "security_domains",
    "weekly_actions",
    "reinforced_security",
    "pending_findings",
    "security_news",
    "limitations",
    "image_citations",
}

MINORITY_JSON_SCHEMA = {
    "type": "object",
    "additionalProperties": False,
    "required": sorted(MINORITY_KEYS),
    "properties": {
        "title": {"type": "string"},
        "client_name": {"type": "string"},
        "prepared_by": {"type": "string"},
        "period": {"type": "string"},
        "service_name": {"type": "string"},
        "tools": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["name", "description"],
                "properties": {
                    "name": {"type": "string"},
                    "description": {"type": "string"},
                },
            },
        },
        "data_base": {"type": "string"},
        "executive_summary": {"type": "string"},
        "vulnerability_comparison": {
            "type": "object",
            "additionalProperties": False,
            "required": ["summary", "severity_rows"],
            "properties": {
                "summary": {"type": "string"},
                "severity_rows": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "additionalProperties": False,
                        "required": ["severity", "previous", "current"],
                        "properties": {
                            "severity": {"type": "string"},
                            "previous": {"type": "string"},
                            "current": {"type": "string"},
                        },
                    },
                },
            },
        },
        "histogram_summary": {"type": "string"},
        "results_and_next_actions": {"type": "string"},
        "results_obtained": {"type": "string"},
        "next_actions": {"type": "array", "items": {"type": "string"}},
        "requirements": {"type": "array", "items": {"type": "string"}},
        "security_domains": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["name", "summary", "findings"],
                "properties": {
                    "name": {"type": "string"},
                    "summary": {"type": "string"},
                    "findings": {
                        "type": "array",
                        "items": {
                            "type": "object",
                            "additionalProperties": False,
                            "required": ["id", "vulnerability", "affected_hosts", "severity"],
                            "properties": {
                                "id": {"type": "string"},
                                "vulnerability": {"type": "string"},
                                "affected_hosts": {"type": "string"},
                                "severity": {"type": "string"},
                            },
                        },
                    },
                },
            },
        },
        "weekly_actions": {"type": "array", "items": {"type": "string"}},
        "reinforced_security": {"type": "string"},
        "pending_findings": {"type": "array", "items": {"type": "string"}},
        "security_news": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["title", "date", "source", "links", "summary", "recommendation"],
                "properties": {
                    "title": {"type": "string"},
                    "date": {"type": "string"},
                    "source": {"type": "string"},
                    "links": {"type": "array", "items": {"type": "string"}},
                    "summary": {"type": "string"},
                    "recommendation": {"type": "string"},
                },
            },
        },
        "limitations": {"type": "array", "items": {"type": "string"}},
        "image_citations": {
            "type": "array",
            "items": {
                "type": "object",
                "additionalProperties": False,
                "required": ["label", "description", "used_in_sections"],
                "properties": {
                    "label": {"type": "string"},
                    "description": {"type": "string"},
                    "used_in_sections": {"type": "array", "items": {"type": "string"}},
                },
            },
        },
    },
}

PROMPT_BASE = """Eres un analista senior XOC y debes generar el contenido de un Minority Report para cliente.
El reporte es ejecutivo-técnico, claro, formal y orientado a valor para el cliente.

Reglas obligatorias:
- Usa únicamente la evidencia entregada: texto, datos estructurados, referencia de formato e imágenes.
- No inventes fechas, IPs, activos, hallazgos, severidades, acciones, resultados, herramientas ni noticias.
- Si algo no se puede confirmar, agrégalo en limitations.
- Si no hay acciones o resultados confirmados, deja arrays vacíos o texto vacío y explica la limitación.
- Cita imágenes como Figura 1, Figura 2, etc. cuando se usen como evidencia.
- No generes DOCX.
- Devuelve SOLO JSON válido, sin markdown ni bloques de código.
- Mantén el estilo de Minority Report: ejecutivo, ordenado, con dominios de seguridad y seguimiento semanal.

Estructura ideal del Minority Report:
1. Datos generales
  1.1 Servicio de Monitoreo
  1.2 Periodo
  1.3 Herramientas
  1.4 Datos Base
2. Resumen ejecutivo del dominio
  2.1 Análisis Comparativo de Vulnerabilidades Semanales
  2.2 Histograma de la seguridad
  2.3 Resultados obtenidos y próximas acciones
  2.4 Resultados obtenidos
  2.5 Próximas acciones
    2.5.1 Requerimiento
3. Seguridad por Dominio
4. Reporte de acciones trabajadas durante la semana
5. Resultados obtenidos
  5.1 Seguridad Reforzada
  5.2 Hallazgos pendientes
6. Noticias de seguridad
"""


def load_env() -> None:
    load_dotenv(ROOT / ".env")


def image_to_data_url(image_path: Path) -> str:
    path = image_path.resolve()
    if not path.is_file():
        raise FileNotFoundError(f"No existe la imagen: {image_path}")
    if path.suffix.lower() not in SUPPORTED_IMAGE_EXTENSIONS:
        raise ValueError(f"Formato no soportado: {path.suffix}. Use PNG, JPG, JPEG o WEBP.")
    max_mb = float(os.environ.get("REPORT_MAX_IMAGE_MB", "10"))
    size_mb = path.stat().st_size / (1024 * 1024)
    if size_mb > max_mb:
        raise ValueError(f"La imagen excede REPORT_MAX_IMAGE_MB={max_mb:g}: {path.name}")
    mime = mimetypes.guess_type(path.name)[0] or "application/octet-stream"
    encoded = base64.b64encode(path.read_bytes()).decode("ascii")
    return f"data:{mime};base64,{encoded}"


def _strip_json_fence(raw: str) -> str:
    raw = raw.strip()
    if raw.startswith("```"):
        raw = re.sub(r"^```(?:json)?\s*|\s*```$", "", raw, flags=re.IGNORECASE)
    return raw.strip()


def _extract_json_object(raw: str) -> str:
    start = raw.find("{")
    end = raw.rfind("}")
    if start == -1 or end == -1 or end <= start:
        return raw
    return raw[start : end + 1]


def _repair_common_json_issues(raw: str) -> str:
    repaired = _extract_json_object(raw)
    repaired = repaired.replace("\ufeff", "").replace("“", '"').replace("”", '"')
    repaired = re.sub(r"//.*?$", "", repaired, flags=re.MULTILINE)
    repaired = re.sub(r"/\*.*?\*/", "", repaired, flags=re.DOTALL)
    repaired = re.sub(r",\s*([}\]])", r"\1", repaired)
    return repaired.strip()


def _loads_json_with_repair(raw: str) -> dict[str, Any]:
    try:
        parsed = json.loads(raw)
    except json.JSONDecodeError as first_exc:
        repaired = _repair_common_json_issues(raw)
        if repaired != raw:
            try:
                parsed = json.loads(repaired)
            except json.JSONDecodeError:
                raise first_exc
        else:
            raise
    if not isinstance(parsed, dict):
        raise RuntimeError("Azure no devolvió un objeto JSON")
    return parsed


def _as_list(value: Any) -> list[Any]:
    if isinstance(value, list):
        return value
    return [value] if value not in (None, "") else []


def _clean_string(value: Any) -> str:
    return str(value or "").strip()


def _normalize_payload(payload: dict[str, Any]) -> dict[str, Any]:
    missing = MINORITY_KEYS - set(payload)
    extra = set(payload) - MINORITY_KEYS
    if missing or extra:
        raise RuntimeError(
            "JSON inesperado para Minority Report. "
            f"Faltan: {', '.join(sorted(missing)) or 'ninguna'}. "
            f"Sobran: {', '.join(sorted(extra)) or 'ninguna'}."
        )

    normalized = {key: payload.get(key) for key in MINORITY_KEYS}
    for key in ("title", "client_name", "prepared_by", "period", "service_name", "data_base", "executive_summary", "histogram_summary", "results_and_next_actions", "results_obtained", "reinforced_security"):
        normalized[key] = _clean_string(normalized[key])

    normalized["tools"] = [
        {"name": _clean_string(item.get("name")), "description": _clean_string(item.get("description"))}
        for item in _as_list(normalized["tools"])
        if isinstance(item, dict) and (_clean_string(item.get("name")) or _clean_string(item.get("description")))
    ]
    comparison = normalized["vulnerability_comparison"] if isinstance(normalized["vulnerability_comparison"], dict) else {}
    normalized["vulnerability_comparison"] = {
        "summary": _clean_string(comparison.get("summary")),
        "severity_rows": [
            {
                "severity": _clean_string(row.get("severity")),
                "previous": _clean_string(row.get("previous")),
                "current": _clean_string(row.get("current")),
            }
            for row in _as_list(comparison.get("severity_rows"))
            if isinstance(row, dict) and _clean_string(row.get("severity"))
        ],
    }
    for key in ("next_actions", "requirements", "weekly_actions", "pending_findings", "limitations"):
        normalized[key] = [_clean_string(item) for item in _as_list(normalized[key]) if _clean_string(item)]

    domains = []
    for domain in _as_list(normalized["security_domains"]):
        if not isinstance(domain, dict):
            continue
        findings = []
        for finding in _as_list(domain.get("findings")):
            if not isinstance(finding, dict):
                continue
            findings.append(
                {
                    "id": _clean_string(finding.get("id")),
                    "vulnerability": _clean_string(finding.get("vulnerability")),
                    "affected_hosts": _clean_string(finding.get("affected_hosts")),
                    "severity": _clean_string(finding.get("severity")),
                }
            )
        name = _clean_string(domain.get("name"))
        summary = _clean_string(domain.get("summary"))
        if name or summary or findings:
            domains.append({"name": name, "summary": summary, "findings": findings})
    normalized["security_domains"] = domains

    news_items = []
    for news in _as_list(normalized["security_news"]):
        if not isinstance(news, dict):
            continue
        title = _clean_string(news.get("title"))
        if not title:
            continue
        news_items.append(
            {
                "title": title,
                "date": _clean_string(news.get("date")),
                "source": _clean_string(news.get("source")),
                "links": [_clean_string(link) for link in _as_list(news.get("links")) if _clean_string(link)],
                "summary": _clean_string(news.get("summary")),
                "recommendation": _clean_string(news.get("recommendation")),
            }
        )
    normalized["security_news"] = news_items

    normalized["image_citations"] = [
        {
            "label": _clean_string(item.get("label")),
            "description": _clean_string(item.get("description")),
            "used_in_sections": [_clean_string(section) for section in _as_list(item.get("used_in_sections")) if _clean_string(section)],
        }
        for item in _as_list(normalized["image_citations"])
        if isinstance(item, dict) and (_clean_string(item.get("label")) or _clean_string(item.get("description")))
    ]
    return normalized


def parse_and_validate_json(raw: str) -> dict[str, Any]:
    clean = _strip_json_fence(raw or "")
    if not clean:
        raise RuntimeError("Azure no devolvió texto JSON visible.")
    try:
        parsed = _loads_json_with_repair(clean)
    except json.JSONDecodeError as exc:
        sample = re.sub(r"\s+", " ", clean[:220]).strip()
        raise RuntimeError(
            "Azure devolvió una respuesta que no es JSON válido "
            f"({exc.msg}, línea {exc.lineno}, columna {exc.colno}). "
            f"Inicio seguro: {sample!r}"
        ) from exc
    return _normalize_payload(parsed)


def save_payload(payload: dict[str, Any], output_path: Path = DEFAULT_PAYLOAD_PATH) -> Path:
    output_path.parent.mkdir(parents=True, exist_ok=True)
    output_path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return output_path


def _content_blocks(prompt: str, image_paths: list[Path]) -> list[dict[str, Any]]:
    blocks: list[dict[str, Any]] = [{"type": "input_text", "text": prompt}]
    for path in image_paths:
        blocks.append({"type": "input_image", "image_url": image_to_data_url(path)})
    return blocks


def _credential_token() -> str:
    credential = DefaultAzureCredential()
    return credential.get_token("https://ai.azure.com/.default").token


def _foundry_openai_client() -> Any:
    from openai import OpenAI

    endpoint = (
        os.environ.get("AZURE_FOUNDRY_OPENAI_ENDPOINT", "").strip()
        or os.environ.get("AZURE_OPENAI_ENDPOINT", "").strip()
        or os.environ.get("AZURE_FOUNDRY_PROJECT_ENDPOINT", "").strip()
    ).rstrip("/")
    if not endpoint:
        raise RuntimeError(
            "Falta configurar Azure para Minority Report. Defina AZURE_FOUNDRY_OPENAI_ENDPOINT "
            "o AZURE_FOUNDRY_PROJECT_ENDPOINT en variables de entorno o en XOC_Minority_Report/.env."
        )
    if endpoint.endswith("/openai/v1"):
        base_url = f"{endpoint}/"
    elif "/api/projects/" in endpoint:
        resource_root = endpoint.split("/api/projects/", 1)[0].rstrip("/")
        base_url = f"{resource_root}/openai/v1/"
    else:
        base_url = f"{endpoint}/openai/v1/"
    api_key = (
        os.environ.get("AZURE_FOUNDRY_API_KEY", "").strip()
        or os.environ.get("AZURE_OPENAI_API_KEY", "").strip()
        or _credential_token()
    )
    return OpenAI(base_url=base_url, api_key=api_key)


def _extract_response_text(response: Any) -> str:
    output_text = str(getattr(response, "output_text", "") or "").strip()
    if output_text:
        return output_text
    pieces: list[str] = []
    for item in getattr(response, "output", []) or []:
        for content in getattr(item, "content", []) or []:
            text = getattr(content, "text", None)
            if text:
                pieces.append(str(text))
            if isinstance(content, dict):
                value = content.get("text") or content.get("value")
                if value:
                    pieces.append(str(value))
    text = "\n".join(piece.strip() for piece in pieces if piece.strip())
    if text:
        return text
    raise RuntimeError("Azure respondió, pero no devolvió texto visible para convertir a JSON.")


def build_prompt(
    *,
    client_name: str,
    period: str,
    analyst_text: str,
    structured_data: dict[str, Any] | None,
    reference_markdown: str,
    image_metadata: list[dict[str, str]],
) -> str:
    structured = structured_data or {}
    return (
        f"{PROMPT_BASE}\n\n"
        f"Cliente objetivo: {client_name or 'No especificado'}\n"
        f"Periodo objetivo: {period or 'No especificado'}\n\n"
        "Texto del analista:\n"
        f"{analyst_text.strip() or 'No se proporcionó texto del analista.'}\n\n"
        "Metadata de imágenes recibidas:\n"
        f"{json.dumps(image_metadata, ensure_ascii=False, indent=2)}\n\n"
        "Datos estructurados opcionales:\n"
        f"{json.dumps(structured, ensure_ascii=False, indent=2)}\n\n"
        "Referencia de formato Minority Report. Úsala solo como guía estructural, no copies datos de cliente si no corresponden:\n"
        f"{reference_markdown[:18000] if reference_markdown else 'No se proporcionó referencia markdown.'}"
    )


def local_fallback_payload(
    *,
    client_name: str,
    period: str,
    analyst_text: str,
    image_metadata: list[dict[str, str]],
    structured_data: dict[str, Any] | None,
) -> dict[str, Any]:
    structured = structured_data or {}
    return _normalize_payload(
        {
            "title": structured.get("title") or "MINORITY REPORT - XOC",
            "client_name": client_name or structured.get("client_name") or "Cliente",
            "prepared_by": structured.get("prepared_by") or "TXDXSECURE",
            "period": period or structured.get("period") or "Periodo no especificado",
            "service_name": structured.get("service_name") or "SERVICIO DE MONITOREO PROACTIVO XOC",
            "tools": structured.get("tools") or [],
            "data_base": structured.get("data_base") or analyst_text.strip(),
            "executive_summary": structured.get("executive_summary") or analyst_text.strip(),
            "vulnerability_comparison": structured.get("vulnerability_comparison")
            or {"summary": "", "severity_rows": []},
            "histogram_summary": structured.get("histogram_summary") or "",
            "results_and_next_actions": structured.get("results_and_next_actions") or "",
            "results_obtained": structured.get("results_obtained") or "",
            "next_actions": structured.get("next_actions") or [],
            "requirements": structured.get("requirements") or [],
            "security_domains": structured.get("security_domains") or [],
            "weekly_actions": structured.get("weekly_actions") or [],
            "reinforced_security": structured.get("reinforced_security") or "",
            "pending_findings": structured.get("pending_findings") or [],
            "security_news": structured.get("security_news") or [],
            "limitations": [
                f"Azure Foundry no fue utilizado. Se recibieron {len(image_metadata)} imagen(es), pero no fueron interpretadas por un modelo multimodal."
            ],
            "image_citations": [
                {
                    "label": item.get("label", f"Figura {index}"),
                    "description": item.get("description") or item.get("file_name") or "Evidencia visual proporcionada.",
                    "used_in_sections": [],
                }
                for index, item in enumerate(image_metadata, start=1)
            ],
        }
    )


def generate_minority_payload(
    *,
    client_name: str,
    period: str,
    analyst_text: str = "",
    image_paths: list[str | Path] | None = None,
    image_descriptions: list[str] | None = None,
    structured_data: dict[str, Any] | None = None,
    reference_markdown: str = "",
    use_azure: bool | None = None,
    allow_local_fallback: bool = True,
    output_path: Path = DEFAULT_PAYLOAD_PATH,
) -> dict[str, Any]:
    load_env()
    images = [Path(path) for path in (image_paths or [])]
    descriptions = image_descriptions or []
    image_metadata = [
        {
            "label": f"Figura {index}",
            "file_name": path.name,
            "description": descriptions[index - 1] if index - 1 < len(descriptions) else "",
        }
        for index, path in enumerate(images, start=1)
    ]
    if use_azure is None:
        use_azure = os.environ.get("USE_AZURE_FOUNDRY", "true").strip().lower() not in {"0", "false", "no"}

    if not use_azure:
        payload = local_fallback_payload(
            client_name=client_name,
            period=period,
            analyst_text=analyst_text,
            image_metadata=image_metadata,
            structured_data=structured_data,
        )
        save_payload(payload, output_path)
        return payload

    prompt = build_prompt(
        client_name=client_name,
        period=period,
        analyst_text=analyst_text,
        structured_data=structured_data,
        reference_markdown=reference_markdown,
        image_metadata=image_metadata,
    )
    try:
        client = _foundry_openai_client()
        deployment = (
            os.environ.get("AZURE_FOUNDRY_MODEL_DEPLOYMENT", "").strip()
            or os.environ.get("AZURE_OPENAI_DEPLOYMENT", "").strip()
            or "gpt-5-mini"
        )
        request: dict[str, Any] = {
            "model": deployment,
            "input": [{"role": "user", "content": _content_blocks(prompt, images)}],
            "max_output_tokens": int(os.environ.get("MINORITY_MAX_OUTPUT_TOKENS", "9000")),
        }
        if os.environ.get("MINORITY_JSON_SCHEMA", "true").strip().lower() not in {"0", "false", "no"}:
            request["text"] = {
                "format": {
                    "type": "json_schema",
                    "name": "xoc_minority_report_payload",
                    "schema": MINORITY_JSON_SCHEMA,
                    "strict": True,
                }
            }
        try:
            response = client.responses.create(**request)
        except Exception:
            if "text" not in request:
                raise
            request.pop("text", None)
            response = client.responses.create(**request)
        payload = parse_and_validate_json(_extract_response_text(response))
    except Exception:
        if not allow_local_fallback:
            raise
        payload = local_fallback_payload(
            client_name=client_name,
            period=period,
            analyst_text=analyst_text,
            image_metadata=image_metadata,
            structured_data=structured_data,
        )

    save_payload(payload, output_path)
    return payload
