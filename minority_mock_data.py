"""Mock DB y agregador para Minority Report.

Este modulo simula una base de datos mensual por cliente. La idea del POC es
que el boton de Minority Report no dependa de texto manual: lee datos, resume,
genera tablas/graficos y entrega un payload compatible con el generador DOCX.
"""

from __future__ import annotations

import json
import random
import re
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402

ROOT = Path(__file__).resolve().parent
MOCK_DB_DIR = ROOT / "mock_db"
MOCK_DB_PATH = MOCK_DB_DIR / "minority_mock_db.json"
REPORT_HISTORY_PATH = MOCK_DB_DIR / "minority_report_history.json"
CHARTS_ROOT = ROOT / "output" / "charts"

DEFAULT_AS_OF = date(2026, 7, 22)
MOCK_MONTH_START = date(2026, 7, 1)
MOCK_MONTH_END = date(2026, 7, 31)

SEVERITIES = ["Crítica", "Alta", "Media", "Baja", "Informativa"]
MONTHS_ES = {
    1: "ENERO",
    2: "FEBRERO",
    3: "MARZO",
    4: "ABRIL",
    5: "MAYO",
    6: "JUNIO",
    7: "JULIO",
    8: "AGOSTO",
    9: "SEPTIEMBRE",
    10: "OCTUBRE",
    11: "NOVIEMBRE",
    12: "DICIEMBRE",
}

COMPANIES: dict[str, dict[str, Any]] = {
    "al-ink": {
        "name": "AL ink",
        "code": "ALINK",
        "numeric": "26091",
        "domains": ["Red corporativa", "Aplicaciones internas", "Endpoints", "Servicios publicados"],
        "base": {"Crítica": 5, "Alta": 17, "Media": 58, "Baja": 31, "Informativa": 138},
    },
    "interbank": {
        "name": "Interbank",
        "code": "INTERBANK",
        "numeric": "26092",
        "domains": ["Core bancario", "Canales digitales", "Perímetro", "Endpoints"],
        "base": {"Crítica": 7, "Alta": 24, "Media": 71, "Baja": 42, "Informativa": 165},
    },
    "supercel": {
        "name": "Supercel",
        "code": "SUPERCEL",
        "numeric": "26093",
        "domains": ["Infraestructura OT", "Red de tiendas", "ERP", "Servicios cloud"],
        "base": {"Crítica": 4, "Alta": 15, "Media": 49, "Baja": 28, "Informativa": 112},
    },
    "hermanos-mario": {
        "name": "Hermanos Mario",
        "code": "HERMANOS_MARIO",
        "numeric": "26094",
        "domains": ["Red administrativa", "Puntos de venta", "Servidores", "Correo corporativo"],
        "base": {"Crítica": 3, "Alta": 13, "Media": 45, "Baja": 25, "Informativa": 96},
    },
}

VULN_NAMES = [
    "TLS con cifrados débiles habilitados",
    "Servicio expuesto sin restricción de origen",
    "Componente web con versión fuera de soporte",
    "Credenciales por defecto detectadas",
    "Servidor con parches pendientes",
    "Cabeceras de seguridad HTTP incompletas",
    "Puerto administrativo accesible desde segmento no esperado",
    "Certificado próximo a vencer",
    "Endpoint con configuración insegura",
    "Agente de monitoreo sin actualización reciente",
]

ACTION_TEMPLATES = [
    "Validación de hallazgos críticos y priorización con el equipo responsable.",
    "Correlación de eventos con actividad de red para descartar falsos positivos.",
    "Revisión de exposición externa y actualización de matriz de riesgo.",
    "Seguimiento de remediaciones aplicadas durante la semana.",
    "Verificación de disponibilidad de servicios monitoreados.",
    "Actualización de inventario técnico con activos observados en monitoreo.",
    "Revisión de controles de hardening y recomendaciones pendientes.",
]

NEWS_POOL = [
    {
        "title": "Campañas de phishing con adjuntos maliciosos dirigidos a áreas administrativas",
        "source": "Boletín XOC",
        "recommendation": "Reforzar filtros de correo y campañas internas de concientización.",
    },
    {
        "title": "Incremento de explotación sobre servicios VPN y portales publicados",
        "source": "Threat Intelligence XOC",
        "recommendation": "Validar MFA, firmware y restricciones por origen en accesos remotos.",
    },
    {
        "title": "Nuevas variantes de ransomware enfocadas en credenciales privilegiadas",
        "source": "Monitoreo de amenazas",
        "recommendation": "Revisar cuentas privilegiadas y aplicar rotación de credenciales críticas.",
    },
    {
        "title": "Uso activo de vulnerabilidades en componentes web legacy",
        "source": "Alertas de seguridad",
        "recommendation": "Priorizar actualización de componentes web expuestos a Internet.",
    },
]


def _daterange(start: date, end: date) -> list[date]:
    return [start + timedelta(days=offset) for offset in range((end - start).days + 1)]


def _iso(value: date) -> str:
    return value.isoformat()


def _parse_date(value: str | None, fallback: date = DEFAULT_AS_OF) -> date:
    if not value:
        return fallback
    return date.fromisoformat(value[:10])


def _safe_code(value: str) -> str:
    value = re.sub(r"[^A-Za-z0-9]+", "_", value.upper()).strip("_")
    return value or "CLIENTE"


def _period_text(start: date, end: date) -> str:
    if start == end:
        return f"{start.day} de {MONTHS_ES[start.month].lower()} de {start.year}"
    return f"Del {start.day} de {MONTHS_ES[start.month].lower()} al {end.day} de {MONTHS_ES[end.month].lower()} de {end.year}"


def _period_suffix(mode: str, end: date) -> str:
    if mode == "last_7_days":
        return f"SEMANAL_{end.day}_{MONTHS_ES[end.month]}"
    if mode == "since_last_report":
        return f"ON_DEMAND_{end.day}_{MONTHS_ES[end.month]}"
    return f"MENSUAL_{end.day}_{MONTHS_ES[end.month]}"


def company_options() -> list[dict[str, str]]:
    return [{"key": key, "name": value["name"]} for key, value in COMPANIES.items()]


def seed_mock_database(force: bool = False) -> Path:
    MOCK_DB_DIR.mkdir(parents=True, exist_ok=True)
    if MOCK_DB_PATH.exists() and not force:
        return MOCK_DB_PATH

    data: dict[str, Any] = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "month": MOCK_MONTH_START.strftime("%Y-%m"),
        "companies": {},
    }
    for company_index, (key, config) in enumerate(COMPANIES.items(), start=1):
        rng = random.Random(20260700 + company_index)
        current = dict(config["base"])
        daily: list[dict[str, Any]] = []
        for day in _daterange(MOCK_MONTH_START, MOCK_MONTH_END):
            remediated = {sev: rng.randint(0, 3 if sev in {"Crítica", "Alta"} else 6) for sev in SEVERITIES}
            new_items = {sev: rng.randint(0, 2 if sev in {"Crítica", "Alta"} else 7) for sev in SEVERITIES}
            for severity in SEVERITIES:
                current[severity] = max(0, current[severity] - remediated[severity] + new_items[severity])

            findings = []
            for finding_idx in range(rng.randint(3, 8)):
                severity = rng.choices(SEVERITIES, weights=[5, 14, 34, 23, 24], k=1)[0]
                domain = rng.choice(config["domains"])
                findings.append(
                    {
                        "id": f"{config['code']}-{day.strftime('%m%d')}-{finding_idx + 1:02d}",
                        "domain": domain,
                        "vulnerability": rng.choice(VULN_NAMES),
                        "affected_hosts": str(rng.randint(1, 14)),
                        "severity": severity,
                        "status": rng.choices(["Pendiente", "En tratamiento", "Mitigado"], weights=[45, 35, 20], k=1)[0],
                    }
                )

            daily.append(
                {
                    "date": _iso(day),
                    "severity_counts": dict(current),
                    "new_findings": sum(new_items.values()),
                    "remediated_findings": sum(remediated.values()),
                    "events_correlated": rng.randint(180, 950),
                    "availability_percent": round(rng.uniform(98.1, 99.99), 2),
                    "actions": rng.sample(ACTION_TEMPLATES, k=rng.randint(2, 4)),
                    "findings": findings,
                }
            )
        data["companies"][key] = {
            "key": key,
            "name": config["name"],
            "code": config["code"],
            "numeric": config["numeric"],
            "domains": config["domains"],
            "tools": [
                {"name": "MonEvents", "description": "Correlación y monitoreo de eventos de seguridad."},
                {"name": "MonVulE", "description": "Escaneo de vulnerabilidades en rangos de gestión y red."},
                {"name": "MonVulC", "description": "Análisis de vulnerabilidades en servidores y activos críticos."},
                {"name": "MonApps", "description": "Monitoreo de disponibilidad de aplicaciones."},
                {"name": "MonNet / MonInfra", "description": "Monitoreo de rendimiento e infraestructura."},
            ],
            "daily": daily,
        }
    MOCK_DB_PATH.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    if not REPORT_HISTORY_PATH.exists():
        REPORT_HISTORY_PATH.write_text("[]", encoding="utf-8")
    return MOCK_DB_PATH


def load_mock_database() -> dict[str, Any]:
    seed_mock_database()
    return json.loads(MOCK_DB_PATH.read_text(encoding="utf-8"))


def load_report_history() -> list[dict[str, Any]]:
    if not REPORT_HISTORY_PATH.exists():
        return []
    try:
        data = json.loads(REPORT_HISTORY_PATH.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return []
    return data if isinstance(data, list) else []


def last_report_for(company_key: str) -> dict[str, Any] | None:
    history = [item for item in load_report_history() if item.get("company_key") == company_key]
    if not history:
        return None
    return sorted(history, key=lambda item: item.get("generated_at", ""))[-1]


def _resolve_range(company_key: str, mode: str, as_of: date) -> tuple[date, date, dict[str, Any] | None]:
    as_of = min(max(as_of, MOCK_MONTH_START), MOCK_MONTH_END)
    last = last_report_for(company_key)
    if mode == "last_7_days":
        start = max(MOCK_MONTH_START, as_of - timedelta(days=6))
    elif mode == "since_last_report" and last:
        # Se incluye el día del último snapshot para que una regeneración el mismo
        # día no quede vacía; el documento indicará que ya existía un reporte previo.
        start = max(MOCK_MONTH_START, _parse_date(last.get("period_end"), MOCK_MONTH_START))
    else:
        start = MOCK_MONTH_START
    return start, as_of, last


def _records_for(company: dict[str, Any], start: date, end: date) -> list[dict[str, Any]]:
    return [row for row in company["daily"] if start <= _parse_date(row["date"]) <= end]


def _sum_counts(records: list[dict[str, Any]], key: str) -> int:
    return sum(int(row.get(key, 0)) for row in records)


def _top_pending_findings(records: list[dict[str, Any]], limit: int = 12) -> list[dict[str, Any]]:
    severity_rank = {"Crítica": 0, "Alta": 1, "Media": 2, "Baja": 3, "Informativa": 4}
    findings = [
        finding
        for row in records
        for finding in row.get("findings", [])
        if finding.get("status") != "Mitigado"
    ]
    findings.sort(key=lambda item: (severity_rank.get(item.get("severity"), 9), item.get("id", "")))
    return findings[:limit]


def _domain_sections(company: dict[str, Any], records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    by_domain: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for finding in _top_pending_findings(records, 40):
        by_domain[finding["domain"]].append(finding)

    sections = []
    for domain in company["domains"]:
        items = by_domain.get(domain, [])
        if not items:
            sections.append(
                {
                    "name": domain,
                    "summary": "No se registraron hallazgos pendientes de prioridad alta para este dominio en el rango evaluado.",
                    "findings": [],
                }
            )
            continue
        counter = Counter(item["severity"] for item in items)
        sections.append(
            {
                "name": domain,
                "summary": (
                    f"Se mantienen {len(items)} hallazgos pendientes asociados al dominio. "
                    f"Distribución principal: "
                    + ", ".join(f"{sev}: {counter[sev]}" for sev in SEVERITIES if counter.get(sev))
                    + "."
                ),
                "findings": [
                    {
                        "id": item["id"],
                        "vulnerability": item["vulnerability"],
                        "affected_hosts": item["affected_hosts"],
                        "severity": item["severity"],
                    }
                    for item in items[:8]
                ],
            }
        )
    return sections


def _make_chart_dir(report_id: str) -> Path:
    chart_dir = CHARTS_ROOT / report_id
    chart_dir.mkdir(parents=True, exist_ok=True)
    return chart_dir


def generate_charts(company_name: str, records: list[dict[str, Any]], report_id: str) -> list[dict[str, Any]]:
    chart_dir = _make_chart_dir(report_id)
    dates = [_parse_date(row["date"]) for row in records]
    labels = [f"{item.day:02d}/{item.month:02d}" for item in dates]
    colors = {
        "Crítica": "#D72638",
        "Alta": "#F46036",
        "Media": "#F3A712",
        "Baja": "#2E86AB",
        "Informativa": "#7A8B99",
    }

    trend_path = chart_dir / "figura_1_tendencia_severidades.png"
    plt.figure(figsize=(9.5, 4.2), dpi=170)
    for severity in SEVERITIES:
        values = [row["severity_counts"][severity] for row in records]
        plt.plot(labels, values, marker="o", linewidth=2, label=severity, color=colors[severity])
    plt.title(f"Tendencia de vulnerabilidades por severidad - {company_name}")
    plt.xlabel("Fecha")
    plt.ylabel("Hallazgos abiertos")
    plt.xticks(rotation=45, ha="right")
    plt.grid(axis="y", alpha=0.25)
    plt.legend(ncol=3, fontsize=8)
    plt.tight_layout()
    plt.savefig(trend_path, bbox_inches="tight")
    plt.close()

    current = records[-1]["severity_counts"]
    severity_path = chart_dir / "figura_2_distribucion_actual.png"
    plt.figure(figsize=(8.2, 4.0), dpi=170)
    bars = plt.bar(SEVERITIES, [current[sev] for sev in SEVERITIES], color=[colors[sev] for sev in SEVERITIES])
    plt.title(f"Distribución actual de hallazgos - {company_name}")
    plt.ylabel("Cantidad")
    plt.grid(axis="y", alpha=0.2)
    for bar in bars:
        height = int(bar.get_height())
        plt.text(bar.get_x() + bar.get_width() / 2, height + 1, str(height), ha="center", fontsize=9)
    plt.tight_layout()
    plt.savefig(severity_path, bbox_inches="tight")
    plt.close()

    activity_path = chart_dir / "figura_3_actividad_remediacion.png"
    plt.figure(figsize=(8.8, 4.0), dpi=170)
    new_values = [row["new_findings"] for row in records]
    rem_values = [row["remediated_findings"] for row in records]
    x = range(len(labels))
    plt.bar(x, new_values, label="Nuevos", color="#00AEEF", alpha=0.8)
    plt.bar(x, [-value for value in rem_values], label="Remediados", color="#00FF9F", alpha=0.8)
    plt.axhline(0, color="#1C2B39", linewidth=0.8)
    plt.title(f"Actividad de hallazgos - {company_name}")
    plt.xticks(list(x), labels, rotation=45, ha="right")
    plt.ylabel("Nuevos / Remediados")
    plt.legend(fontsize=8)
    plt.tight_layout()
    plt.savefig(activity_path, bbox_inches="tight")
    plt.close()

    return [
        {"path": trend_path, "description": "Tendencia de hallazgos abiertos por severidad durante el periodo evaluado."},
        {"path": severity_path, "description": "Distribución actual de vulnerabilidades por severidad al cierre del reporte."},
        {"path": activity_path, "description": "Comparativo diario entre nuevos hallazgos y hallazgos remediados."},
    ]


def build_mock_minority_report(company_key: str, range_mode: str = "month_to_date", as_of: str | None = None) -> dict[str, Any]:
    db = load_mock_database()
    if company_key not in db["companies"]:
        raise ValueError(f"Empresa mock no soportada: {company_key}")
    company = db["companies"][company_key]
    end_date = _parse_date(as_of, DEFAULT_AS_OF)
    start_date, end_date, last_snapshot = _resolve_range(company_key, range_mode, end_date)
    records = _records_for(company, start_date, end_date)
    if not records:
        raise ValueError("No hay data mock para el rango solicitado.")

    first = records[0]["severity_counts"]
    current = records[-1]["severity_counts"]
    total_new = _sum_counts(records, "new_findings")
    total_remediated = _sum_counts(records, "remediated_findings")
    availability_avg = sum(float(row["availability_percent"]) for row in records) / len(records)
    events_total = _sum_counts(records, "events_correlated")
    pending = _top_pending_findings(records, 10)
    critical_delta = current["Crítica"] - first["Crítica"]
    high_delta = current["Alta"] - first["Alta"]

    mode_label = {
        "last_7_days": "semanal",
        "since_last_report": "on demand",
        "month_to_date": "mensual acumulado",
    }.get(range_mode, "mensual acumulado")
    period = _period_text(start_date, end_date)
    document_code = f"TXDX-WR-{end_date:%y%m}-{company['numeric']}-{company['code']}_MINORITY_REPORT_{_period_suffix(range_mode, end_date)}"
    report_id = f"{company_key}_{range_mode}_{end_date.isoformat()}_{int(datetime.now().timestamp())}"
    charts = generate_charts(company["name"], records, report_id)

    actions = []
    for row in records[-8:]:
        actions.extend(row.get("actions", [])[:2])
    actions = list(dict.fromkeys(actions))[:12]

    limitations = [
        "La información corresponde a data mock local para validar el flujo del POC.",
        "El rango no se recorta automáticamente por el último reporte para evitar reportes vacíos generados por usuarios distintos.",
    ]
    if last_snapshot:
        limitations.append(
            f"Existe un snapshot previo generado el {last_snapshot.get('generated_at')} para el periodo "
            f"{last_snapshot.get('period_start')} a {last_snapshot.get('period_end')}; este reporte puede considerarse una regeneración/ampliación."
        )

    payload = {
        "title": f"Minority Report XOC - {company['name']}",
        "client_name": company["name"],
        "prepared_by": "TXDXSECURE",
        "period": period,
        "service_name": "Servicio de monitoreo proactivo XOC",
        "tools": company["tools"],
        "data_base": (
            f"Se consolidaron {len(records)} días de registros mock para {company['name']}, "
            f"incluyendo {events_total:,} eventos correlacionados, {total_new} nuevos hallazgos, "
            f"{total_remediated} hallazgos remediados y disponibilidad promedio de {availability_avg:.2f}%."
        ),
        "executive_summary": (
            f"Durante el periodo {period}, el servicio XOC mantuvo seguimiento {mode_label} sobre los dominios monitoreados de "
            f"{company['name']}. Al cierre se observan {current['Crítica']} hallazgos críticos, {current['Alta']} altos, "
            f"{current['Media']} medios, {current['Baja']} bajos y {current['Informativa']} informativos. "
            f"La variación frente al inicio del rango es de {critical_delta:+d} críticos y {high_delta:+d} altos."
        ),
        "vulnerability_comparison": {
            "summary": (
                "El análisis comparativo toma como referencia el primer día del rango frente al último día disponible. "
                "La lectura permite identificar cambios de postura y priorizar la continuidad de remediaciones."
            ),
            "severity_rows": [
                {"severity": sev, "previous": str(first[sev]), "current": str(current[sev])}
                for sev in SEVERITIES
            ],
        },
        "histogram_summary": (
            "La Figura 1 muestra la evolución diaria por severidad, la Figura 2 resume la distribución actual "
            "y la Figura 3 compara nuevos hallazgos contra remediaciones ejecutadas."
        ),
        "results_and_next_actions": (
            f"Se remediaron {total_remediated} hallazgos y se registraron {total_new} nuevos elementos para análisis. "
            "Las próximas acciones deben concentrarse en reducir exposición crítica/alta y sostener la validación de activos con mayor recurrencia."
        ),
        "results_obtained": [
            f"Disponibilidad promedio observada: {availability_avg:.2f}%.",
            f"Eventos correlacionados durante el rango: {events_total:,}.",
            f"Hallazgos remediados registrados: {total_remediated}.",
        ],
        "next_actions": [
            "Priorizar tratamiento de hallazgos críticos y altos aún abiertos.",
            "Revisar controles de exposición externa en dominios con mayor recurrencia.",
            "Validar ventanas de remediación con responsables técnicos del cliente.",
        ],
        "requirements": [
            "Confirmar responsables por dominio para los hallazgos pendientes.",
            "Validar evidencia de remediación antes del siguiente corte de reporte.",
        ],
        "security_domains": _domain_sections(company, records),
        "weekly_actions": actions,
        "reinforced_security": [
            "Se mantuvo correlación continua de eventos y seguimiento de vulnerabilidades por dominio.",
            "Se consolidó priorización por criticidad para facilitar el tratamiento operativo.",
            "Se generaron visualizaciones de tendencia para explicar cambios de postura al cliente.",
        ],
        "pending_findings": [
            f"{item['id']} - {item['vulnerability']} ({item['severity']}, {item['affected_hosts']} hosts)"
            for item in pending
        ],
        "security_news": [
            {
                "title": news["title"],
                "date": end_date.isoformat(),
                "source": news["source"],
                "links": ["Referencia mock para POC"],
                "summary": "Noticia simulada para completar la sección de inteligencia y recomendaciones del Minority Report.",
                "recommendation": news["recommendation"],
            }
            for news in NEWS_POOL[:3]
        ],
        "limitations": limitations,
        "image_citations": [
            {"label": f"Figura {idx}", "description": chart["description"], "used_in_sections": ["2.2 Histograma de la seguridad"]}
            for idx, chart in enumerate(charts, start=1)
        ],
        "document_code": document_code,
        "mock_meta": {
            "company_key": company_key,
            "range_mode": range_mode,
            "period_start": start_date.isoformat(),
            "period_end": end_date.isoformat(),
            "last_snapshot": last_snapshot,
        },
    }
    return {"payload": payload, "charts": charts, "records": records, "last_snapshot": last_snapshot}


def register_report(company_key: str, range_mode: str, payload: dict[str, Any], docx_path: Path) -> Path:
    MOCK_DB_DIR.mkdir(parents=True, exist_ok=True)
    history = load_report_history()
    meta = payload.get("mock_meta") or {}
    history.append(
        {
            "company_key": company_key,
            "company_name": payload.get("client_name"),
            "range_mode": range_mode,
            "period_start": meta.get("period_start"),
            "period_end": meta.get("period_end"),
            "generated_at": datetime.now().isoformat(timespec="seconds"),
            "document_code": payload.get("document_code"),
            "docx_path": str(docx_path),
        }
    )
    REPORT_HISTORY_PATH.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")
    return REPORT_HISTORY_PATH

