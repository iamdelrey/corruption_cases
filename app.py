from __future__ import annotations

import html
import io
import json
import mimetypes
import os
import re
import secrets
import sqlite3
import textwrap
import uuid
from datetime import datetime
from http import HTTPStatus
from http.cookies import SimpleCookie
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from pathlib import Path
from typing import Any
from urllib.parse import parse_qs, quote, urlencode, urlparse

from docx import Document
from email.parser import BytesParser
from email.policy import default
from reportlab.lib.pagesizes import A4
from reportlab.lib.utils import simpleSplit
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.pdfgen import canvas

BASE_DIR = Path(__file__).resolve().parent
DB_PATH = BASE_DIR / "data.sqlite3"
STATIC_DIR = BASE_DIR / "static"
UPLOADS_DIR = BASE_DIR / "uploads"
PHOTOS_DIR = UPLOADS_DIR / "photos"
IMPORTS_DIR = UPLOADS_DIR / "imports"
SEED_JSON_PATH = BASE_DIR / "seed_cases.json"
FONT_PATH = "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"
HOST = os.environ.get("CASES_HOST", "127.0.0.1")
PORT = int(os.environ.get("CASES_PORT", "8080"))
ADMIN_LOGIN = os.environ.get("ADMIN_LOGIN", "admin")
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD", "admin123")
SESSION_COOKIE = "corruption_cases_session"
SESSIONS: dict[str, dict[str, Any]] = {}

SECTION_LABELS = {
    "russia": "Коррупционные кейсы в России",
    "foreign": "Коррупционные кейсы в иностранных государствах",
    "intl-orgs": "Коррупционные кейсы в международных организациях",
}
SECTION_SHORT = {
    "russia": "Россия",
    "foreign": "Иностранные государства",
    "intl-orgs": "Международные организации",
}
HEADER_NAV_LABELS = {
    "about": "О проекте",
    "russia": "Россия",
    "foreign": "Иностранные государства",
    "intl-orgs": "Международные организации",
}
VIOLATION_PRESETS = [
    "взяточничество",
    "злоупотребление полномочиями",
    "конфликт интересов",
    "хищение / растрата",
    "мошенничество",
    "незаконное обогащение",
    "лоббизм / влияние в обход процедур",
    "иное",
]
DOCX_HEADINGS = {
    "фабула дела": "case_summary",
    "правовая квалификация": "legal_qualification",
    "ход дела": "case_progress",
    "последствия": "consequences",
    "институциональные эффекты": "institutional_effects",
    "выводы и уроки для антикоррупционной политики": "policy_lessons",
    "источники": "sources",
    "базовые данные": "basic_data",
}
FIELD_LABELS = {
    "full_name": "ФИО",
    "short_description": "Краткое описание",
    "year_or_period": "Год / период",
    "amount": "Сумма",
    "country": "Страна",
    "organization": "Организация",
    "jurisdiction": "Юрисдикция",
    "governance_level": "Уровень управления",
    "risk_sector": "Отрасль риска",
    "violation_type": "Тип нарушения",
    "case_summary": "Фабула дела",
    "legal_qualification": "Правовая квалификация",
    "case_progress": "Ход дела",
    "consequences": "Последствия",
    "institutional_effects": "Институциональные эффекты",
    "policy_lessons": "Выводы и уроки для антикоррупционной политики",
}
TRANSLIT = {
    "а": "a", "б": "b", "в": "v", "г": "g", "д": "d", "е": "e", "ж": "zh", "з": "z", "и": "i",
    "й": "y", "к": "k", "л": "l", "м": "m", "н": "n", "о": "o", "п": "p", "р": "r", "с": "s",
    "т": "t", "у": "u", "ф": "f", "х": "h", "ц": "ts", "ч": "ch", "ш": "sh", "щ": "sch", "ъ": "",
    "ы": "y", "ь": "", "э": "e", "ю": "yu", "я": "ya",
}


def now_iso() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def html_escape(value: Any) -> str:
    return html.escape("" if value is None else str(value), quote=True)


def normalize_spaces(value: str) -> str:
    return re.sub(r"\s+", " ", value or "").strip()


def transliterate(value: str) -> str:
    result = []
    for ch in value.lower():
        result.append(TRANSLIT.get(ch, ch))
    return "".join(result)


def slugify(value: str) -> str:
    value = transliterate(value)
    value = re.sub(r"[^a-z0-9]+", "-", value.lower())
    value = re.sub(r"-+", "-", value).strip("-")
    if not value:
        value = f"case-{uuid.uuid4().hex[:8]}"
    return value


CASE_TITLE_RE = re.compile(r"^Кейс\s*(?:№)?\s*\d+\s*[:：]\s*(.+)$", re.I)
DOCX_SECTION_RE = re.compile(
    r"^(?:\d+\.\s*)?(Базовые данные|Фабула дела|Правовая квалификация|Ход дела|Последствия|Институциональные эффекты|Выводы и уроки(?: для антикоррупционной политики)?|Источники)\s*:?$",
    re.I,
)
DOCX_FIELD_MAP = {
    "страна/юрисдикция": "country",
    "страна": "country",
    "международная организация": "organization",
    "юрисдикция уголовного разбирательства": "jurisdiction",
    "годы развития дела": "year_or_period",
    "годы": "year_or_period",
    "уровень власти": "governance_level",
    "уровень управления": "governance_level",
    "отрасль риска": "risk_sector",
    "тип коррупционного поведения": "violation_type",
}
AMOUNT_RE = re.compile(r"(\d[\d\s,.]*(?:млн|миллион|млрд)?\s*(?:руб\.?|рублей|USD|доллар[а-я]* США|евро))", re.I)


def cleanup_case_title(title: str) -> str:
    title = normalize_spaces(title)
    title = re.sub(r"^(?:уголовное\s+)?дело\s+", "", title, flags=re.I)
    return title.strip()


def extract_url(value: str) -> str:
    match = re.search(r"https?://\S+", value or "")
    if not match:
        return ""
    return match.group(0).rstrip(").,")


def build_short_description(value: str) -> str:
    value = normalize_spaces(value)
    if not value:
        return ""
    parts = re.split(r"(?<=[.!?])\s+", value)
    sentence = parts[0]
    if len(sentence) < 90 and len(parts) > 1:
        sentence = f"{sentence} {parts[1]}"
    if len(sentence) > 220:
        sentence = sentence[:217].rstrip(" ,;:") + "..."
    return sentence


def empty_case_payload() -> dict[str, Any]:
    return {
        "slug": "",
        "section": "russia",
        "full_name": "",
        "short_description": "",
        "year_or_period": "",
        "amount": "",
        "country": "",
        "organization": "",
        "jurisdiction": "",
        "governance_level": "",
        "risk_sector": "",
        "violation_type": "",
        "case_summary": "",
        "legal_qualification": "",
        "case_progress": "",
        "consequences": "",
        "institutional_effects": "",
        "policy_lessons": "",
        "sources": [],
        "status": "published",
    }


def parse_case_chunks_from_paragraphs(paragraphs: list[str], default_section: str = "russia") -> list[dict[str, Any]]:
    normalized = [normalize_spaces(p.replace("\xa0", " ")) for p in paragraphs if normalize_spaces(p)]
    if not normalized:
        return []

    chunks: list[dict[str, Any]] = []
    current: dict[str, Any] | None = None
    for line in normalized:
        title_match = CASE_TITLE_RE.match(line)
        if title_match:
            if current:
                chunks.append(current)
            current = {"title": cleanup_case_title(title_match.group(1)), "lines": []}
            continue
        if current is None:
            current = {"title": cleanup_case_title(normalized[0]), "lines": []}
        current["lines"].append(line)
    if current:
        chunks.append(current)

    parsed_cases: list[dict[str, Any]] = []
    for chunk in chunks:
        data = empty_case_payload()
        data["full_name"] = chunk["title"] or "Кейс"
        data["slug"] = slugify(data["full_name"])
        data["section"] = default_section
        blocks: dict[str, list[str]] = {}
        current_block: str | None = None

        for line in chunk["lines"]:
            section_match = DOCX_SECTION_RE.match(line)
            if section_match:
                current_block = section_match.group(1).lower().replace(" для антикоррупционной политики", "")
                blocks.setdefault(current_block, [])
                continue
            if current_block:
                blocks.setdefault(current_block, []).append(line)

        for line in blocks.get("базовые данные", []):
            if ":" not in line:
                continue
            key, value = line.split(":", 1)
            target = DOCX_FIELD_MAP.get(normalize_spaces(key.lower()))
            if not target:
                continue
            data[target] = normalize_spaces(value).rstrip(".")

        if data["section"] == "russia" and data["country"] and "россий" in data["country"].lower():
            data["country"] = "Россия"
        if data["organization"] and data["section"] == "russia":
            data["section"] = "intl-orgs"
            data["country"] = ""

        block_mapping = {
            "фабула дела": "case_summary",
            "правовая квалификация": "legal_qualification",
            "ход дела": "case_progress",
            "последствия": "consequences",
            "институциональные эффекты": "institutional_effects",
            "выводы и уроки": "policy_lessons",
        }
        for source_key, target_key in block_mapping.items():
            data[target_key] = "\n\n".join(blocks.get(source_key, []))

        for line in blocks.get("источники", []):
            cleaned = re.sub(r"^\d+[\).]?\s*", "", line).strip()
            if cleaned:
                data["sources"].append({"gost_text": cleaned, "url": extract_url(cleaned)})

        amount_source = " ".join([
            data["case_summary"],
            data["legal_qualification"],
            data["case_progress"],
            data["consequences"],
        ])
        amount_match = AMOUNT_RE.search(amount_source)
        if amount_match:
            data["amount"] = normalize_spaces(amount_match.group(1))
        elif "в особо крупном размере" in amount_source.lower():
            data["amount"] = "особо крупный размер"

        data["short_description"] = build_short_description(
            data["case_summary"] or data["violation_type"] or data["risk_sector"] or data["full_name"]
        )
        parsed_cases.append(data)

    return parsed_cases


def load_seed_cases() -> list[dict[str, Any]]:
    if not SEED_JSON_PATH.exists():
        return []
    try:
        payload = json.loads(SEED_JSON_PATH.read_text(encoding="utf-8"))
    except Exception:
        return []
    result: list[dict[str, Any]] = []
    for item in payload:
        if not isinstance(item, dict) or not item.get("full_name"):
            continue
        prepared = empty_case_payload()
        prepared.update(item)
        prepared["slug"] = slugify(prepared.get("slug") or prepared["full_name"])
        prepared["status"] = prepared.get("status") or "published"
        result.append(prepared)
    return result




class SimpleUploadedFile:
    def __init__(self, filename: str, content: bytes) -> None:
        self.file_name = filename
        self.size = len(content)
        self.file_object = io.BytesIO(content)

def section_to_placeholder(section: str, country: str | None, organization: str | None) -> str:
    if section == "russia":
        return "РФ"
    if section == "foreign":
        return (country or "Флаг")[:18]
    return (organization or "Орг.")[:18]


class Database:
    def __init__(self, path: Path):
        self.path = path

    def connect(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.path)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA foreign_keys = ON")
        return conn

    def init(self) -> None:
        conn = self.connect()
        try:
            conn.executescript(
                """
                CREATE TABLE IF NOT EXISTS cases (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    slug TEXT NOT NULL UNIQUE,
                    section TEXT NOT NULL,
                    full_name TEXT NOT NULL,
                    short_description TEXT NOT NULL,
                    photo_path TEXT,
                    year_or_period TEXT,
                    amount TEXT,
                    country TEXT,
                    organization TEXT,
                    jurisdiction TEXT,
                    governance_level TEXT,
                    risk_sector TEXT,
                    violation_type TEXT,
                    case_summary TEXT,
                    legal_qualification TEXT,
                    case_progress TEXT,
                    consequences TEXT,
                    institutional_effects TEXT,
                    policy_lessons TEXT,
                    status TEXT NOT NULL,
                    published_at TEXT,
                    created_at TEXT NOT NULL,
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS sources (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    case_id INTEGER NOT NULL,
                    gost_text TEXT NOT NULL,
                    url TEXT,
                    sort_order INTEGER NOT NULL DEFAULT 0,
                    FOREIGN KEY(case_id) REFERENCES cases(id) ON DELETE CASCADE
                );

                CREATE TABLE IF NOT EXISTS about_page (
                    id INTEGER PRIMARY KEY CHECK (id = 1),
                    goal TEXT NOT NULL DEFAULT '',
                    methodology TEXT NOT NULL DEFAULT '',
                    contacts TEXT NOT NULL DEFAULT '',
                    education_note TEXT NOT NULL DEFAULT '',
                    updated_at TEXT NOT NULL
                );

                CREATE TABLE IF NOT EXISTS countries (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE
                );

                CREATE TABLE IF NOT EXISTS organizations (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE
                );

                CREATE TABLE IF NOT EXISTS violation_types (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    name TEXT NOT NULL UNIQUE
                );
                """
            )
            about = conn.execute("SELECT id FROM about_page WHERE id = 1").fetchone()
            if not about:
                conn.execute(
                    "INSERT INTO about_page (id, goal, methodology, contacts, education_note, updated_at) VALUES (1, ?, ?, ?, ?, ?)",
                    (
                        "Учебно-просветительский проект, посвященный структурированному анализу коррупционных кейсов.",
                        "Кейсы собираются в унифицированной структуре: базовые данные, фабула, квалификация, ход дела, последствия, институциональные эффекты и выводы.",
                        "E-mail: project@example.org",
                        "Проект носит учебно-просветительский характер и не является юридической консультацией.",
                        now_iso(),
                    ),
                )
            for item in VIOLATION_PRESETS:
                conn.execute("INSERT OR IGNORE INTO violation_types(name) VALUES (?)", (item,))
            self.seed_demo(conn)
            conn.commit()
        finally:
            conn.close()

    
    def seed_demo(self, conn: sqlite3.Connection) -> None:
        seeds = load_seed_cases()
        if seeds:
            existing_slugs = {row["slug"] for row in conn.execute("SELECT slug FROM cases").fetchall()}
            for sample in seeds:
                if sample["slug"] in existing_slugs:
                    continue
                created_at = now_iso()
                updated_at = created_at
                published_at = created_at if sample.get("status") == "published" else None
                conn.execute(
                    textwrap.dedent(
                        """
                        INSERT INTO cases (
                            slug, section, full_name, short_description, photo_path, year_or_period, amount,
                            country, organization, jurisdiction, governance_level, risk_sector, violation_type,
                            case_summary, legal_qualification, case_progress, consequences,
                            institutional_effects, policy_lessons, status, published_at, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """
                    ),
                    (
                        sample["slug"], sample["section"], sample["full_name"], sample["short_description"], None,
                        sample.get("year_or_period"), sample.get("amount"), sample.get("country"), sample.get("organization"),
                        sample.get("jurisdiction"), sample.get("governance_level"), sample.get("risk_sector"),
                        sample.get("violation_type"), sample.get("case_summary"), sample.get("legal_qualification"),
                        sample.get("case_progress"), sample.get("consequences"), sample.get("institutional_effects"),
                        sample.get("policy_lessons"), sample.get("status") or "published", published_at, created_at, updated_at,
                    ),
                )
                case_id = conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
                for index, source in enumerate(sample.get("sources", []), start=1):
                    conn.execute(
                        "INSERT INTO sources(case_id, gost_text, url, sort_order) VALUES (?, ?, ?, ?)",
                        (case_id, source["gost_text"], source.get("url"), index),
                    )
                if sample.get("country"):
                    conn.execute("INSERT OR IGNORE INTO countries(name) VALUES (?)", (sample["country"],))
                if sample.get("organization"):
                    conn.execute("INSERT OR IGNORE INTO organizations(name) VALUES (?)", (sample["organization"],))
                if sample.get("violation_type"):
                    conn.execute("INSERT OR IGNORE INTO violation_types(name) VALUES (?)", (sample["violation_type"],))
            return

        existing = conn.execute("SELECT COUNT(*) AS cnt FROM cases").fetchone()["cnt"]
        if existing:
            return
        samples = [
            {
                "slug": "aleksey-petrov",
                "section": "russia",
                "full_name": "Алексей Петров",
                "short_description": "Региональный чиновник, обвиненный в получении крупной взятки при распределении контрактов.",
                "country": "Россия",
                "year_or_period": "2019-2021",
                "amount": "35 млн руб.",
                "jurisdiction": "национальный суд",
                "governance_level": "региональный",
                "risk_sector": "госзакупки",
                "violation_type": "взяточничество",
                "case_summary": "Следствие установило систематическое получение незаконного вознаграждения при выборе подрядчиков для инфраструктурных проектов.",
                "legal_qualification": "Получение взятки в особо крупном размере, злоупотребление должностными полномочиями.",
                "case_progress": "Возбуждение дела, предъявление обвинения, рассмотрение в суде первой инстанции.",
                "consequences": "Назначено наказание, конфискована часть имущества, пересмотрены процедуры размещения заказов.",
                "institutional_effects": "Усилен внутренний контроль за конкурсными процедурами.",
                "policy_lessons": "Необходимы прозрачные процедуры закупок и цифровой след принятия решений.",
                "status": "published",
                "published_at": now_iso(),
            },
            {
                "slug": "carlos-mendes",
                "section": "foreign",
                "full_name": "Карлос Мендес",
                "short_description": "Корпоративный посредник в деле о трансграничных платежах и фиктивных консалтинговых договорах.",
                "country": "Бразилия",
                "year_or_period": "2016-2019",
                "amount": "5 млн USD",
                "jurisdiction": "национальный суд",
                "governance_level": "федеральный",
                "risk_sector": "инфраструктура",
                "violation_type": "мошенничество",
                "case_summary": "Фиктивные договоры использовались для сокрытия незаконных платежей и обхода процедур контроля.",
                "legal_qualification": "Мошенничество, отмывание средств, коррупционные платежи.",
                "case_progress": "Следственные действия, международные запросы, соглашения о сотрудничестве.",
                "consequences": "Штрафы, запрет на участие в тендерах, репутационные потери.",
                "institutional_effects": "Уточнены требования к проверке посредников и бенефициаров.",
                "policy_lessons": "Проверка цепочек подрядчиков должна быть обязательной на ранней стадии.",
                "status": "published",
                "published_at": now_iso(),
            },
            {
                "slug": "fifa-procurement-case",
                "section": "intl-orgs",
                "full_name": "Комитет закупок FIFA",
                "short_description": "Кейс о непрозрачных процедурах выбора подрядчиков в международной спортивной организации.",
                "organization": "FIFA",
                "year_or_period": "2015-2018",
                "amount": "не раскрыта",
                "jurisdiction": "международный орган",
                "governance_level": "международный",
                "risk_sector": "спорт",
                "violation_type": "конфликт интересов",
                "case_summary": "Проверка выявила пересечение личных интересов и решений по контрактам.",
                "legal_qualification": "Конфликт интересов, нарушение внутренних регламентов.",
                "case_progress": "Внутренняя проверка, дисциплинарные меры, пересмотр регламентов.",
                "consequences": "Увольнения, изменение состава комитетов, новые требования к раскрытию интересов.",
                "institutional_effects": "Созданы дополнительные механизмы раскрытия аффилированности.",
                "policy_lessons": "Даже при мягкой правовой квалификации институциональные выводы могут быть значимыми.",
                "status": "published",
                "published_at": now_iso(),
            },
        ]
        for sample in samples:
            created_at = now_iso()
            updated_at = created_at
            conn.execute(
                textwrap.dedent(
                    """
                    INSERT INTO cases (
                        slug, section, full_name, short_description, photo_path, year_or_period, amount,
                        country, organization, jurisdiction, governance_level, risk_sector, violation_type,
                        case_summary, legal_qualification, case_progress, consequences,
                        institutional_effects, policy_lessons, status, published_at, created_at, updated_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                    """
                ),
                (
                    sample["slug"], sample["section"], sample["full_name"], sample["short_description"], None,
                    sample.get("year_or_period"), sample.get("amount"), sample.get("country"), sample.get("organization"),
                    sample.get("jurisdiction"), sample.get("governance_level"), sample.get("risk_sector"),
                    sample.get("violation_type"), sample.get("case_summary"), sample.get("legal_qualification"),
                    sample.get("case_progress"), sample.get("consequences"), sample.get("institutional_effects"),
                    sample.get("policy_lessons"), sample["status"], sample.get("published_at"), created_at, updated_at,
                ),
            )
            case_id = conn.execute("SELECT id FROM cases WHERE slug = ?", (sample["slug"],)).fetchone()["id"]
            conn.execute(
                "INSERT INTO sources(case_id, gost_text, url, sort_order) VALUES (?, ?, ?, ?)",
                (case_id, f"Демонстрационный источник по кейсу «{sample['full_name']}».", "https://example.org", 1),
            )
            if sample.get("country"):
                conn.execute("INSERT OR IGNORE INTO countries(name) VALUES (?)", (sample["country"],))
            if sample.get("organization"):
                conn.execute("INSERT OR IGNORE INTO organizations(name) VALUES (?)", (sample["organization"],))


DB = Database(DB_PATH)



def ensure_dirs() -> None:
    STATIC_DIR.mkdir(parents=True, exist_ok=True)
    PHOTOS_DIR.mkdir(parents=True, exist_ok=True)
    IMPORTS_DIR.mkdir(parents=True, exist_ok=True)


class CasesRepository:
    @staticmethod
    def get_case_by_slug(slug: str, include_hidden: bool = False) -> sqlite3.Row | None:
        conn = DB.connect()
        try:
            if include_hidden:
                return conn.execute("SELECT * FROM cases WHERE slug = ?", (slug,)).fetchone()
            return conn.execute("SELECT * FROM cases WHERE slug = ? AND status = 'published'", (slug,)).fetchone()
        finally:
            conn.close()

    @staticmethod
    def get_case_by_id(case_id: int) -> sqlite3.Row | None:
        conn = DB.connect()
        try:
            return conn.execute("SELECT * FROM cases WHERE id = ?", (case_id,)).fetchone()
        finally:
            conn.close()

    @staticmethod
    def get_sources(case_id: int) -> list[sqlite3.Row]:
        conn = DB.connect()
        try:
            return conn.execute("SELECT * FROM sources WHERE case_id = ? ORDER BY sort_order, id", (case_id,)).fetchall()
        finally:
            conn.close()

    @staticmethod
    def upsert_case(data: dict[str, Any], case_id: int | None = None) -> int:
        conn = DB.connect()
        try:
            created_at = now_iso()
            updated_at = created_at
            published_at = data.get("published_at")
            if data.get("status") == "published" and not published_at:
                published_at = created_at
            payload = (
                data["slug"], data["section"], data["full_name"], data["short_description"], data.get("photo_path"),
                data.get("year_or_period"), data.get("amount"), data.get("country"), data.get("organization"),
                data.get("jurisdiction"), data.get("governance_level"), data.get("risk_sector"), data.get("violation_type"),
                data.get("case_summary"), data.get("legal_qualification"), data.get("case_progress"), data.get("consequences"),
                data.get("institutional_effects"), data.get("policy_lessons"), data["status"], published_at,
            )
            if case_id is None:
                conn.execute(
                    textwrap.dedent(
                        """
                        INSERT INTO cases (
                            slug, section, full_name, short_description, photo_path, year_or_period, amount,
                            country, organization, jurisdiction, governance_level, risk_sector, violation_type,
                            case_summary, legal_qualification, case_progress, consequences,
                            institutional_effects, policy_lessons, status, published_at, created_at, updated_at
                        ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                        """
                    ),
                    payload + (created_at, updated_at),
                )
                case_id = conn.execute("SELECT last_insert_rowid() AS id").fetchone()["id"]
            else:
                conn.execute(
                    textwrap.dedent(
                        """
                        UPDATE cases SET
                            slug = ?, section = ?, full_name = ?, short_description = ?, photo_path = ?, year_or_period = ?, amount = ?,
                            country = ?, organization = ?, jurisdiction = ?, governance_level = ?, risk_sector = ?, violation_type = ?,
                            case_summary = ?, legal_qualification = ?, case_progress = ?, consequences = ?,
                            institutional_effects = ?, policy_lessons = ?, status = ?, published_at = ?, updated_at = ?
                        WHERE id = ?
                        """
                    ),
                    payload + (updated_at, case_id),
                )
                conn.execute("DELETE FROM sources WHERE case_id = ?", (case_id,))
            for index, source in enumerate(data.get("sources", []), start=1):
                conn.execute(
                    "INSERT INTO sources(case_id, gost_text, url, sort_order) VALUES (?, ?, ?, ?)",
                    (case_id, source["gost_text"], source.get("url"), index),
                )
            if data.get("country"):
                conn.execute("INSERT OR IGNORE INTO countries(name) VALUES (?)", (data["country"],))
            if data.get("organization"):
                conn.execute("INSERT OR IGNORE INTO organizations(name) VALUES (?)", (data["organization"],))
            if data.get("violation_type"):
                conn.execute("INSERT OR IGNORE INTO violation_types(name) VALUES (?)", (data["violation_type"],))
            conn.commit()
            return int(case_id)
        finally:
            conn.close()

    @staticmethod
    def change_status(case_id: int, status: str) -> None:
        conn = DB.connect()
        try:
            published_at = None
            if status == "published":
                published_at = now_iso()
            conn.execute(
                "UPDATE cases SET status = ?, published_at = COALESCE(?, published_at), updated_at = ? WHERE id = ?",
                (status, published_at, now_iso(), case_id),
            )
            conn.commit()
        finally:
            conn.close()

    @staticmethod
    def list_public(section: str | None = None, q: str = "", country: str = "", violation_type: str = "", year: str = "", sort: str = "new") -> list[sqlite3.Row]:
        conn = DB.connect()
        try:
            clauses = ["status = 'published'"]
            params: list[Any] = []
            if section:
                clauses.append("section = ?")
                params.append(section)
            if q:
                like = f"%{q.strip()}%"
                search_fields = " OR ".join([
                    "full_name LIKE ?", "short_description LIKE ?", "country LIKE ?", "organization LIKE ?", "violation_type LIKE ?",
                    "case_summary LIKE ?", "legal_qualification LIKE ?", "case_progress LIKE ?", "consequences LIKE ?", "policy_lessons LIKE ?",
                ])
                clauses.append(f"({search_fields})")
                params.extend([like] * 10)
            if country:
                clauses.append("country = ?")
                params.append(country)
            if violation_type:
                clauses.append("violation_type = ?")
                params.append(violation_type)
            if year:
                clauses.append("year_or_period LIKE ?")
                params.append(f"%{year}%")
            order_by = {
                "new": "published_at DESC, updated_at DESC",
                "alpha": "full_name COLLATE NOCASE ASC",
                "country": "country COLLATE NOCASE ASC, full_name COLLATE NOCASE ASC",
            }.get(sort, "published_at DESC, updated_at DESC")
            sql = f"SELECT * FROM cases WHERE {' AND '.join(clauses)} ORDER BY {order_by}"
            return conn.execute(sql, params).fetchall()
        finally:
            conn.close()

    @staticmethod
    def list_admin(q: str = "", status: str = "") -> list[sqlite3.Row]:
        conn = DB.connect()
        try:
            clauses = ["1=1"]
            params: list[Any] = []
            if q:
                like = f"%{q.strip()}%"
                clauses.append("(full_name LIKE ? OR short_description LIKE ? OR slug LIKE ?)")
                params.extend([like, like, like])
            if status:
                clauses.append("status = ?")
                params.append(status)
            sql = f"SELECT * FROM cases WHERE {' AND '.join(clauses)} ORDER BY updated_at DESC"
            return conn.execute(sql, params).fetchall()
        finally:
            conn.close()

    @staticmethod
    def stats() -> dict[str, int]:
        conn = DB.connect()
        try:
            result = {}
            result["all"] = conn.execute("SELECT COUNT(*) AS cnt FROM cases").fetchone()["cnt"]
            for status in ["published", "draft", "hidden"]:
                result[status] = conn.execute("SELECT COUNT(*) AS cnt FROM cases WHERE status = ?", (status,)).fetchone()["cnt"]
            for section in SECTION_LABELS:
                key = section.replace("-", "_")
                result[key] = conn.execute("SELECT COUNT(*) AS cnt FROM cases WHERE section = ?", (section,)).fetchone()["cnt"]
            return {k: int(v) for k, v in result.items()}
        finally:
            conn.close()

    @staticmethod
    def get_about() -> sqlite3.Row:
        conn = DB.connect()
        try:
            row = conn.execute("SELECT * FROM about_page WHERE id = 1").fetchone()
            if row is None:
                raise RuntimeError("about page not initialized")
            return row
        finally:
            conn.close()

    @staticmethod
    def update_about(goal: str, methodology: str, contacts: str, education_note: str) -> None:
        conn = DB.connect()
        try:
            conn.execute(
                "UPDATE about_page SET goal = ?, methodology = ?, contacts = ?, education_note = ?, updated_at = ? WHERE id = 1",
                (goal, methodology, contacts, education_note, now_iso()),
            )
            conn.commit()
        finally:
            conn.close()

    @staticmethod
    def list_dictionary(table: str) -> list[str]:
        conn = DB.connect()
        try:
            rows = conn.execute(f"SELECT name FROM {table} ORDER BY name COLLATE NOCASE ASC").fetchall()
            return [row["name"] for row in rows]
        finally:
            conn.close()

    @staticmethod
    def add_dictionary_value(table: str, name: str) -> None:
        conn = DB.connect()
        try:
            conn.execute(f"INSERT OR IGNORE INTO {table}(name) VALUES (?)", (name,))
            conn.commit()
        finally:
            conn.close()


def parse_sources_text(sources_text: str) -> list[dict[str, str]]:
    items = []
    for raw_line in (sources_text or "").splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if "|" in line:
            gost_text, url = line.split("|", 1)
            items.append({"gost_text": gost_text.strip(), "url": url.strip()})
        else:
            items.append({"gost_text": line, "url": ""})
    return items


def save_uploaded_file(file_obj, folder: Path, prefix: str) -> str:
    folder.mkdir(parents=True, exist_ok=True)
    original_name = Path(file_obj.file_name.decode("utf-8", errors="ignore") if isinstance(file_obj.file_name, bytes) else (file_obj.file_name or "upload.bin")).name
    ext = Path(original_name).suffix or ".bin"
    name = f"{prefix}-{uuid.uuid4().hex}{ext}"
    target = folder / name
    file_obj.file_object.seek(0)
    with open(target, "wb") as output:
        output.write(file_obj.file_object.read())
    return name



def parse_docx_bytes(content: bytes) -> dict[str, Any]:
    doc = Document(io.BytesIO(content))
    cases = parse_case_chunks_from_paragraphs([p.text for p in doc.paragraphs], default_section="russia")
    if not cases:
        return empty_case_payload()
    return cases[0]


def text_to_paragraphs(value: str) -> str:
    blocks = [block.strip() for block in re.split(r"\n\n+", value or "") if block.strip()]
    if not blocks:
        return '<p class="muted">Нет данных.</p>'
    rendered: list[str] = []
    for block in blocks:
        lines = [line.strip("•·- ").strip() for line in block.split("\n") if line.strip()]
        if len(lines) > 1 and all(len(line) < 220 for line in lines):
            rendered.append("<ul>" + "".join(f"<li>{html_escape(line)}</li>" for line in lines) + "</ul>")
        else:
            rendered.append(f"<p>{html_escape(block)}</p>")
    return "".join(rendered)


def field_input(value: str) -> str:
    parts = [p.strip() for p in (value or "").split("\n") if p.strip()]
    if not parts:
        return '<p class="muted">Нет данных.</p>'
    return "".join(f"<p>{html_escape(part)}</p>" for part in parts)


def field_input(name: str, label: str, value: str = "", input_type: str = "text", required: bool = False, datalist: str = "") -> str:
    attrs = " required" if required else ""
    list_attr = f' list="{datalist}"' if datalist else ""
    return f'''<label class="form-field"><span>{html_escape(label)}</span><input type="{input_type}" name="{html_escape(name)}" value="{html_escape(value)}"{attrs}{list_attr}></label>'''


def textarea_input(name: str, label: str, value: str = "", rows: int = 5) -> str:
    return f'''<label class="form-field"><span>{html_escape(label)}</span><textarea name="{html_escape(name)}" rows="{rows}">{html_escape(value)}</textarea></label>'''


def render_datalists() -> str:
    countries = CasesRepository.list_dictionary("countries")
    orgs = CasesRepository.list_dictionary("organizations")
    vtypes = CasesRepository.list_dictionary("violation_types")
    def build(name: str, values: list[str]) -> str:
        options = "".join(f"<option value=\"{html_escape(v)}\"></option>" for v in values)
        return f"<datalist id=\"{name}\">{options}</datalist>"
    return build("countries-list", countries) + build("orgs-list", orgs) + build("violation-list", vtypes)



def render_public_layout(title: str, body: str, current_section: str = "") -> bytes:
    header_nav = [
        ("about", HEADER_NAV_LABELS["about"], "/about"),
        ("russia", HEADER_NAV_LABELS["russia"], "/cases/russia"),
        ("foreign", HEADER_NAV_LABELS["foreign"], "/cases/foreign"),
        ("intl-orgs", HEADER_NAV_LABELS["intl-orgs"], "/cases/intl-orgs"),
    ]
    nav_links = "".join(
        f'<a href="{url}" class="{"active" if current_section == slug else ""}">{html_escape(label)}</a>'
        for slug, label, url in header_nav
    )
    html_doc = f"""<!doctype html>
<html lang=\"ru\">
<head>
  <meta charset=\"utf-8\">
  <meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">
  <title>{html_escape(title)} - Открытая библиотека коррупционных кейсов</title>
  <link rel=\"stylesheet\" href=\"/static/style.css\">
</head>
<body>
  <div class=\"shell-bg\"></div>
  <header class=\"site-shell\">
    <div class=\"container site-header\">
      <a href=\"/\" class=\"brand\">
        <span class=\"brand-mark\">К</span>
        <span class=\"brand-copy\">
          <strong>Открытая библиотека коррупционных кейсов</strong>
          <span>и сравнительной практики</span>
        </span>
      </a>
      <nav class=\"main-nav\">
        {nav_links}
      </nav>
      <div class=\"header-actions\">
        <form action=\"/search\" method=\"get\" class=\"header-search\">
          <input type=\"search\" name=\"q\" placeholder=\"Поиск по кейсам\">
          <button type=\"submit\">Найти</button>
        </form>
      </div>
    </div>
  </header>
  <main class=\"container page-stack\">{body}</main>
  <footer class=\"site-footer\">
    <div class=\"container footer-grid\">
      <div>
        <strong>Библиотека коррупционных кейсов</strong>
        <p>Учебно-просветительская витрина для систематизации российских и международных антикоррупционных материалов.</p>
      </div>
      <div class=\"footer-links\">
        <a href=\"/about\">О проекте</a>
        <a href=\"/cases/russia\">Россия</a>
        <a href=\"/cases/foreign\">Иностранные государства</a>
        <a href=\"/cases/intl-orgs\">Международные организации</a>
        <a href=\"/admin/login\" class=\"footer-admin-link\">Вход для администратора</a>
      </div>
    </div>
  </footer>
</body>
</html>"""
    return html_doc.encode("utf-8")


def render_admin_layout(title: str, body: str, flash: str = "") -> bytes:
    flash_html = f'<div class="flash">{html_escape(flash)}</div>' if flash else ""
    html_doc = f'''<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>{html_escape(title)} - Админка</title>
  <link rel="stylesheet" href="/static/style.css">
</head>
<body class="admin-body">
  <aside class="admin-sidebar">
    <div class="admin-brand">Админ-панель</div>
    <nav>
      <a href="/admin">Статистика</a>
      <a href="/admin/cases">Кейсы</a>
      <a href="/admin/import">Импорт из Word</a>
      <a href="/admin/dictionaries">Справочники</a>
      <a href="/admin/about">О проекте</a>
    </nav>
    <form method="post" action="/admin/logout"><button type="submit" class="ghost full">Выйти</button></form>
  </aside>
  <main class="admin-main">
    <div class="admin-inner">{flash_html}{body}</div>
  </main>
</body>
</html>'''
    return html_doc.encode("utf-8")



def render_case_card(case: sqlite3.Row) -> str:
    if case["photo_path"]:
        media = f'<img class="case-photo" src="/uploads/photos/{quote(case["photo_path"])}" alt="{html_escape(case["full_name"])}">'
    else:
        placeholder = section_to_placeholder(case["section"], case["country"], case["organization"])
        media = f'<div class="case-photo placeholder section-{html_escape(case["section"])}"><span>{html_escape(placeholder)}</span></div>'

    chips: list[str] = []
    if case["country"]:
        chips.append(case["country"])
    elif case["organization"]:
        chips.append(case["organization"])
    if case["year_or_period"]:
        chips.append(case["year_or_period"])
    if case["violation_type"]:
        chips.append(case["violation_type"])
    chips_html = "".join(f'<span class="chip">{html_escape(item)}</span>' for item in chips[:3])

    return f"""
    <article class="case-card">
      <a href="/case/{html_escape(case['slug'])}">
        <div class="case-media">{media}</div>
        <div class="case-card-body">
          <div class="card-section">{html_escape(SECTION_SHORT.get(case["section"], case["section"]))}</div>
          <h3>{html_escape(case['full_name'])}</h3>
          <p>{html_escape(case['short_description'])}</p>
          <div class="chip-row">{chips_html}</div>
        </div>
      </a>
    </article>
    """


def build_public_filters(section: str, q: str, country: str, violation_type: str, year: str, sort: str) -> str:
    countries = CasesRepository.list_dictionary("countries")
    violation_types = CasesRepository.list_dictionary("violation_types")
    country_options = '<option value="">Все страны</option>' + "".join(
        f'<option value="{html_escape(item)}" {"selected" if item == country else ""}>{html_escape(item)}</option>' for item in countries
    )
    violation_options = '<option value="">Все типы нарушений</option>' + "".join(
        f'<option value="{html_escape(item)}" {"selected" if item == violation_type else ""}>{html_escape(item)}</option>' for item in violation_types
    )
    sort_options = [
        ("new", "Сначала новые"),
        ("alpha", "По алфавиту"),
        ("country", "По стране"),
    ]
    sort_html = "".join(
        f'<option value="{key}" {"selected" if sort == key else ""}>{label}</option>' for key, label in sort_options
    )
    return f"""
    <form method="get" class="filters-panel">
      <div class="search-slot"><input type="search" name="q" value="{html_escape(q)}" placeholder="Поиск внутри раздела"></div>
      <select name="country">{country_options}</select>
      <select name="violation_type">{violation_options}</select>
      <input type="text" name="year" value="{html_escape(year)}" placeholder="Год или период">
      <select name="sort">{sort_html}</select>
      <button type="submit">Найти</button>
      <a class="ghost" href="/cases/{section}">Сбросить</a>
    </form>
    """


def build_case_form(case: dict[str, Any]) -> str:
    sources_value = "\n".join(
        f"{item['gost_text']} | {item.get('url', '')}" if item.get("url") else item["gost_text"]
        for item in case.get("sources", [])
    )
    section_options = "".join(
        f'<option value="{key}" {"selected" if case.get("section") == key else ""}>{html_escape(label)}</option>'
        for key, label in SECTION_LABELS.items()
    )
    photo_note = f'<div class="muted">Текущее фото: {html_escape(case.get("photo_path") or "не загружено")}</div>' if case.get("photo_path") else '<div class="muted">Фото не загружено.</div>'
    return f'''
      {render_datalists()}
      <div class="form-grid two">
        <label class="form-field"><span>Раздел</span><select name="section">{section_options}</select></label>
        {field_input("slug", "Slug", case.get("slug", ""), required=True)}
        {field_input("full_name", "ФИО", case.get("full_name", ""), required=True)}
        {field_input("short_description", "Краткое описание", case.get("short_description", ""), required=True)}
        {field_input("year_or_period", "Год / период", case.get("year_or_period", ""))}
        {field_input("amount", "Сумма", case.get("amount", ""))}
        {field_input("country", "Страна", case.get("country", ""), datalist="countries-list")}
        {field_input("organization", "Организация", case.get("organization", ""), datalist="orgs-list")}
        {field_input("jurisdiction", "Юрисдикция", case.get("jurisdiction", ""))}
        {field_input("governance_level", "Уровень управления", case.get("governance_level", ""))}
        {field_input("risk_sector", "Отрасль риска", case.get("risk_sector", ""))}
        {field_input("violation_type", "Тип нарушения", case.get("violation_type", ""), datalist="violation-list")}
      </div>
      <label class="form-field"><span>Фото</span><input type="file" name="photo" accept="image/*">{photo_note}</label>
      {textarea_input("case_summary", "Фабула дела", case.get("case_summary", ""), rows=6)}
      {textarea_input("legal_qualification", "Правовая квалификация", case.get("legal_qualification", ""), rows=5)}
      {textarea_input("case_progress", "Ход дела", case.get("case_progress", ""), rows=5)}
      {textarea_input("consequences", "Последствия", case.get("consequences", ""), rows=5)}
      {textarea_input("institutional_effects", "Институциональные эффекты", case.get("institutional_effects", ""), rows=5)}
      {textarea_input("policy_lessons", "Выводы и уроки для антикоррупционной политики", case.get("policy_lessons", ""), rows=5)}
      {textarea_input("sources_text", "Источники (по одному на строку, формат: запись | ссылка)", sources_value, rows=6)}
      <div class="form-actions">
        <button type="submit">Сохранить</button>
      </div>
    '''


def generate_case_pdf(case: sqlite3.Row, sources: list[sqlite3.Row]) -> bytes:
    buffer = io.BytesIO()
    if Path(FONT_PATH).exists():
        pdfmetrics.registerFont(TTFont("DejaVuSans", FONT_PATH))
        font_name = "DejaVuSans"
    else:
        font_name = "Helvetica"
    pdf = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4
    margin_x = 45
    y = height - 50

    def write_block(title: str, text: str) -> None:
        nonlocal y
        lines = simpleSplit(text or "Нет данных.", font_name, 10, width - 2 * margin_x)
        needed = 20 + 14 * (len(lines) + 1)
        if y - needed < 50:
            pdf.showPage()
            pdf.setFont(font_name, 11)
            y = height - 50
        pdf.setFont(font_name, 12)
        pdf.drawString(margin_x, y, title)
        y -= 18
        pdf.setFont(font_name, 10)
        for line in lines:
            pdf.drawString(margin_x, y, line)
            y -= 14
        y -= 6

    pdf.setTitle(case["full_name"])
    pdf.setFont(font_name, 16)
    pdf.drawString(margin_x, y, case["full_name"])
    y -= 24
    basics = [
        f"Раздел: {SECTION_SHORT.get(case['section'], case['section'])}",
        f"Год / период: {case['year_or_period'] or '—'}",
        f"Сумма: {case['amount'] or '—'}",
        f"Страна: {case['country'] or '—'}",
        f"Организация: {case['organization'] or '—'}",
        f"Юрисдикция: {case['jurisdiction'] or '—'}",
        f"Уровень управления: {case['governance_level'] or '—'}",
        f"Отрасль риска: {case['risk_sector'] or '—'}",
        f"Тип нарушения: {case['violation_type'] or '—'}",
    ]
    write_block("Базовые данные", "\n".join(basics))
    write_block("Краткое описание", case["short_description"])
    write_block("Фабула дела", case["case_summary"])
    write_block("Правовая квалификация", case["legal_qualification"])
    write_block("Ход дела", case["case_progress"])
    write_block("Последствия", case["consequences"])
    write_block("Институциональные эффекты", case["institutional_effects"])
    write_block("Выводы и уроки для антикоррупционной политики", case["policy_lessons"])
    write_block("Источники", "\n".join([s["gost_text"] + (f" ({s['url']})" if s["url"] else "") for s in sources]) or "Нет источников.")
    pdf.save()
    return buffer.getvalue()


class AppHandler(BaseHTTPRequestHandler):
    server_version = "CorruptionCasesMVP/1.0"

    def log_message(self, format: str, *args: Any) -> None:
        print("%s - - [%s] %s" % (self.address_string(), self.log_date_time_string(), format % args))

    def do_GET(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        query = parse_qs(parsed.query)
        try:
            if path == "/":
                return self.handle_home()
            if path.startswith("/static/"):
                return self.handle_static(path)
            if path.startswith("/uploads/"):
                return self.handle_uploads(path)
            if path.startswith("/cases/"):
                section = path.split("/", 2)[2]
                return self.handle_section(section, query)
            if path.startswith("/case/") and path.endswith("/pdf"):
                slug = path.removeprefix("/case/").removesuffix("/pdf").strip("/")
                return self.handle_case_pdf(slug)
            if path.startswith("/case/"):
                slug = path.split("/", 2)[2]
                return self.handle_case(slug)
            if path == "/about":
                return self.handle_about()
            if path == "/search":
                return self.handle_search(query)
            if path == "/admin/login":
                return self.handle_admin_login_page()
            if path == "/admin":
                return self.handle_admin_dashboard()
            if path == "/admin/cases":
                return self.handle_admin_cases(query)
            if path == "/admin/case/new":
                return self.handle_admin_case_form(None)
            if path.startswith("/admin/case/"):
                case_id = int(path.split("/")[-1])
                return self.handle_admin_case_form(case_id)
            if path == "/admin/import":
                return self.handle_admin_import_page()
            if path == "/admin/about":
                return self.handle_admin_about_page()
            if path == "/admin/dictionaries":
                return self.handle_admin_dictionaries_page()
            return self.respond_text("Не найдено", status=404)
        except Exception as exc:
            return self.respond_text(f"Внутренняя ошибка: {exc}", status=500)

    def do_POST(self) -> None:
        parsed = urlparse(self.path)
        path = parsed.path
        try:
            if path == "/admin/login":
                return self.handle_admin_login()
            if path == "/admin/logout":
                return self.handle_admin_logout()
            if path == "/admin/import":
                return self.handle_admin_import_submit()
            if path == "/admin/about":
                return self.handle_admin_about_save()
            if path == "/admin/dictionaries":
                return self.handle_admin_dictionary_add()
            if path == "/admin/case/new":
                return self.handle_admin_case_save(None)
            if path.startswith("/admin/case/") and path.endswith("/status"):
                case_id = int(path.split("/")[-2])
                return self.handle_admin_case_status(case_id)
            if path.startswith("/admin/case/"):
                case_id = int(path.split("/")[-1])
                return self.handle_admin_case_save(case_id)
            return self.respond_text("Не найдено", status=404)
        except Exception as exc:
            return self.respond_text(f"Внутренняя ошибка: {exc}", status=500)

    def parse_form_data(self) -> tuple[dict[str, str], dict[str, Any]]:
        content_type = self.headers.get("Content-Type", "")
        content_length = int(self.headers.get("Content-Length", "0") or "0")
        fields: dict[str, str] = {}
        files: dict[str, Any] = {}
        body = self.rfile.read(content_length)
        if content_type.startswith("multipart/form-data"):
            header_blob = (f"Content-Type: {content_type}\r\nMIME-Version: 1.0\r\n\r\n").encode("utf-8")
            message = BytesParser(policy=default).parsebytes(header_blob + body)
            for part in message.iter_parts():
                disposition = part.get_content_disposition()
                if disposition != "form-data":
                    continue
                name = part.get_param("name", header="content-disposition")
                filename = part.get_filename()
                payload = part.get_payload(decode=True) or b""
                if not name:
                    continue
                if filename:
                    files[name] = SimpleUploadedFile(filename=filename, content=payload)
                else:
                    fields[name] = payload.decode("utf-8", errors="ignore")
        else:
            parsed = parse_qs(body.decode("utf-8", errors="ignore"), keep_blank_values=True)
            fields = {key: values[0] if values else "" for key, values in parsed.items()}
        return fields, files

    def set_session(self) -> str:
        token = secrets.token_urlsafe(24)
        SESSIONS[token] = {"created_at": now_iso()}
        return token

    def get_session_token(self) -> str | None:
        raw = self.headers.get("Cookie")
        if not raw:
            return None
        cookie = SimpleCookie()
        cookie.load(raw)
        morsel = cookie.get(SESSION_COOKIE)
        if not morsel:
            return None
        token = morsel.value
        if token in SESSIONS:
            return token
        return None

    def require_admin(self) -> bool:
        token = self.get_session_token()
        if token:
            return True
        self.redirect("/admin/login")
        return False

    def respond_html(self, payload: bytes, status: int = 200, headers: dict[str, str] | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", "text/html; charset=utf-8")
        self.send_header("Content-Length", str(len(payload)))
        if headers:
            for key, value in headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(payload)

    def respond_bytes(self, payload: bytes, content_type: str, status: int = 200, headers: dict[str, str] | None = None) -> None:
        self.send_response(status)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(payload)))
        if headers:
            for key, value in headers.items():
                self.send_header(key, value)
        self.end_headers()
        self.wfile.write(payload)

    def respond_text(self, text: str, status: int = 200) -> None:
        self.respond_bytes(text.encode("utf-8"), "text/plain; charset=utf-8", status=status)

    def redirect(self, location: str, headers: dict[str, str] | None = None) -> None:
        self.send_response(302)
        self.send_header("Location", location)
        if headers:
            for key, value in headers.items():
                self.send_header(key, value)
        self.end_headers()

    def handle_static(self, path: str) -> None:
        target = (STATIC_DIR / path.removeprefix("/static/")).resolve()
        if STATIC_DIR not in target.parents and target != STATIC_DIR:
            return self.respond_text("Forbidden", status=403)
        if not target.exists() or not target.is_file():
            return self.respond_text("Not found", status=404)
        content_type, _ = mimetypes.guess_type(target.name)
        self.respond_bytes(target.read_bytes(), content_type or "application/octet-stream")

    def handle_uploads(self, path: str) -> None:
        target = (UPLOADS_DIR / path.removeprefix("/uploads/")).resolve()
        if UPLOADS_DIR not in target.parents and target != UPLOADS_DIR:
            return self.respond_text("Forbidden", status=403)
        if not target.exists() or not target.is_file():
            return self.respond_text("Not found", status=404)
        content_type, _ = mimetypes.guess_type(target.name)
        self.respond_bytes(target.read_bytes(), content_type or "application/octet-stream")


    def handle_home(self) -> None:
        stats = CasesRepository.stats()
        latest = CasesRepository.list_public(sort="new")[:6]
        highlights = "".join(render_case_card(case) for case in latest)
        section_cards = "".join(
            f"""
            <a class="feature-section section-{slug}" href="/cases/{slug}">
              <div>
                <span>{html_escape(SECTION_SHORT.get(slug, slug))}</span>
                <strong>{html_escape(label)}</strong>
                <p>{stats[slug.replace('-', '_')]} материалов в библиотеке</p>
              </div>
            </a>
            """
            for slug, label in SECTION_LABELS.items()
        )
        about = CasesRepository.get_about()
        audience_cards = """
        <div class="audience-grid">
          <article class="info-card">
            <h3>Студенты и исследователи</h3>
            <p>Быстрый вход в тему: единая структура кейсов, ясная фабула, квалификация, последствия и выводы.</p>
          </article>
          <article class="info-card">
            <h3>Юристы и аналитики</h3>
            <p>Сравнение правовых подходов, институциональных эффектов и отраслей риска в разных юрисдикциях.</p>
          </article>
          <article class="info-card">
            <h3>Журналисты и преподаватели</h3>
            <p>Готовая база для обзоров, лекций, дискуссий и учебных материалов без ручной раскладки документов.</p>
          </article>
        </div>
        """
        body = f"""
        <section class="hero-panel">
          <div class="hero-copy">
            <div class="eyebrow">Открытая библиотека коррупционных кейсов</div>
            <h1>Российская, зарубежная и международная сравнительная практика</h1>
            <p class="lead">Цифровой ресурс для изучения коррупционных кейсов в единой структуре: базовые данные, фабула, квалификация, ход дела, последствия и институциональные выводы.</p>
            <div class="hero-actions">
              <a class="primary-link" href="/cases/russia">Перейти к кейсам</a>
              <a class="ghost" href="/about">О проекте</a>
            </div>
          </div>
          <div class="hero-side stat-rail">
            <article class="glass-card"><span>Всего кейсов</span><strong>{stats["all"]}</strong></article>
            <article class="glass-card"><span>Опубликовано</span><strong>{stats["published"]}</strong></article>
            <article class="glass-card"><span>Международные организации</span><strong>{stats["intl_orgs"]}</strong></article>
          </div>
        </section>

        <section class="search-band">
          <form action="/search" method="get" class="search-band-form">
            <input type="search" name="q" placeholder="Поиск по ФИО, стране, организации, квалификации, фабуле">
            <button type="submit">Найти</button>
          </form>
        </section>

        <section class="section-block">
          <div class="section-head"><h2>Разделы библиотеки</h2><p class="muted">Единый интерфейс для России, иностранных государств и международных организаций.</p></div>
          <div class="feature-grid">{section_cards}</div>
        </section>

        <section class="section-block">
          <div class="section-head"><h2>Для кого полезен ресурс</h2><p class="muted">Сайт собран как витрина знаний, а не как архив файлов.</p></div>
          {audience_cards}
        </section>

        <section class="section-block about-surface">
          <div class="section-head"><h2>О проекте</h2><a href="/about">Открыть страницу</a></div>
          <div class="info-surface">
            <p>{html_escape(about['goal'])}</p>
            <p>{html_escape(about['methodology'])}</p>
          </div>
        </section>

        <section class="section-block">
          <div class="section-head"><h2>Последние опубликованные кейсы</h2><a href="/cases/russia">Смотреть каталог</a></div>
          <div class="cards-grid">{highlights}</div>
        </section>
        """
        self.respond_html(render_public_layout("Главная", body))


    def handle_section(self, section: str, query: dict[str, list[str]]) -> None:
        if section not in SECTION_LABELS:
            return self.respond_text("Раздел не найден", status=404)
        q = query.get("q", [""])[0]
        country = query.get("country", [""])[0]
        violation_type = query.get("violation_type", [""])[0]
        year = query.get("year", [""])[0]
        sort = query.get("sort", ["new"])[0]
        items = CasesRepository.list_public(section=section, q=q, country=country, violation_type=violation_type, year=year, sort=sort)
        cards = "".join(render_case_card(case) for case in items) or '<div class="empty">По заданным условиям кейсы не найдены.</div>'
        body = f"""
          <section class="page-hero slim">
            <div>
              <div class="eyebrow">{html_escape(SECTION_SHORT.get(section, section))}</div>
              <h1>{html_escape(SECTION_LABELS[section])}</h1>
              <p class="lead">Поиск внутри раздела, фильтры по стране, типу нарушения и году, сортировка по ключевым сценариям MVP.</p>
            </div>
          </section>
          {build_public_filters(section, q, country, violation_type, year, sort)}
          <section class="cards-grid">{cards}</section>
        """
        self.respond_html(render_public_layout(SECTION_LABELS[section], body, current_section=section))


    def handle_case(self, slug: str) -> None:
        case = CasesRepository.get_case_by_slug(slug)
        if case is None:
            return self.respond_text("Кейс не найден", status=404)
        sources = CasesRepository.get_sources(case["id"])
        related = [item for item in CasesRepository.list_public(section=case["section"], sort="new") if item["id"] != case["id"]][:3]
        related_html = "".join(
            f'<a class="related-item" href="/case/{html_escape(item["slug"])}"><strong>{html_escape(item["full_name"])}</strong><span>{html_escape(item["short_description"])}</span></a>'
            for item in related
        ) or '<div class="muted">Похожие кейсы появятся по мере наполнения библиотеки.</div>'

        source_items = "".join(
            f'<li><div>{html_escape(src["gost_text"])}</div>' + (f'<a href="{html_escape(src["url"])}" target="_blank" rel="noopener">Открыть источник</a>' if src["url"] else "") + '</li>'
            for src in sources
        ) or '<li>Источники не добавлены.</li>'

        basic_fields = [
            ("Раздел", SECTION_SHORT.get(case["section"], case["section"])),
            ("Страна", case["country"]),
            ("Организация", case["organization"]),
            ("Год / период", case["year_or_period"]),
            ("Сумма", case["amount"]),
            ("Юрисдикция", case["jurisdiction"]),
            ("Уровень управления", case["governance_level"]),
            ("Отрасль риска", case["risk_sector"]),
            ("Тип нарушения", case["violation_type"]),
        ]
        basics_html = "".join(
            f'<div class="basic-item"><span>{html_escape(label)}</span><strong>{html_escape(value or "—")}</strong></div>'
            for label, value in basic_fields
        )

        chips = [case["country"] or case["organization"], case["year_or_period"], case["violation_type"], case["risk_sector"]]
        chip_html = "".join(f'<span class="chip">{html_escape(item)}</span>' for item in chips if item)

        body = f"""
        <article class="case-layout">
          <section class="case-main">
            <div class="page-hero case-top">
              <div>
                <div class="eyebrow">{html_escape(SECTION_SHORT.get(case['section'], case['section']))}</div>
                <h1>{html_escape(case['full_name'])}</h1>
                <div class="chip-row">{chip_html}</div>
                <p class="lead">{html_escape(case['short_description'])}</p>
              </div>
            </div>

            <nav class="case-anchor-nav">
              <a href="#basic">Базовые данные</a>
              <a href="#summary">Фабула</a>
              <a href="#qualification">Правовая квалификация</a>
              <a href="#progress">Ход дела</a>
              <a href="#effects">Последствия</a>
              <a href="#lessons">Выводы</a>
            </nav>

            <section id="basic" class="case-section"><h2>Базовые данные</h2><div class="basic-grid">{basics_html}</div></section>
            <section id="summary" class="case-section"><h2>Фабула дела</h2>{text_to_paragraphs(case['case_summary'])}</section>
            <section id="qualification" class="case-section"><h2>Правовая квалификация</h2>{text_to_paragraphs(case['legal_qualification'])}</section>
            <section id="progress" class="case-section"><h2>Ход дела</h2>{text_to_paragraphs(case['case_progress'])}</section>
            <section id="effects" class="case-section"><h2>Последствия</h2>{text_to_paragraphs(case['consequences'])}</section>
            <section class="case-section"><h2>Институциональные эффекты</h2>{text_to_paragraphs(case['institutional_effects'])}</section>
            <section id="lessons" class="case-section accent-section"><h2>Выводы и уроки для антикоррупционной политики</h2>{text_to_paragraphs(case['policy_lessons'])}</section>
            <section class="case-section"><h2>Источники</h2><ol class="sources-list">{source_items}</ol></section>
          </section>

          <aside class="case-sidebar">
            <div class="sidebar-card">
              <h3>Паспорт кейса</h3>
              <div class="sidebar-meta">
                <div><span>Раздел</span><strong>{html_escape(SECTION_SHORT.get(case['section'], case['section']))}</strong></div>
                <div><span>Год</span><strong>{html_escape(case['year_or_period'] or '—')}</strong></div>
                <div><span>Юрисдикция</span><strong>{html_escape(case['jurisdiction'] or '—')}</strong></div>
              </div>
              <a class="primary-link full" href="/case/{html_escape(case['slug'])}/pdf">Скачать PDF</a>
            </div>
            <div class="sidebar-card">
              <h3>Ключевой риск</h3>
              <p>{html_escape(case['risk_sector'] or case['violation_type'] or 'Коррупционный риск не указан.')}</p>
            </div>
            <div class="sidebar-card">
              <h3>Похожие кейсы</h3>
              <div class="related-list">{related_html}</div>
            </div>
          </aside>
        </article>
        """
        self.respond_html(render_public_layout(case["full_name"], body, current_section=case["section"]))

    def handle_case_pdf(self, slug: str) -> None:
        case = CasesRepository.get_case_by_slug(slug)
        if case is None:
            return self.respond_text("PDF доступен только для опубликованного кейса", status=404)
        sources = CasesRepository.get_sources(case["id"])
        payload = generate_case_pdf(case, sources)
        self.respond_bytes(payload, "application/pdf", headers={"Content-Disposition": f'inline; filename="{slug}.pdf"'})


    def handle_about(self) -> None:
        about = CasesRepository.get_about()
        body = f"""
        <section class="page-hero slim">
          <div>
            <div class="eyebrow">О проекте</div>
            <h1>Методика и принципы библиотеки</h1>
            <p class="lead">Страница редактируется из административной панели без участия разработчика и может меняться по мере развития библиотеки.</p>
          </div>
        </section>
        <section class="about-grid">
          <article class="info-card"><h2>Цель проекта</h2>{text_to_paragraphs(about['goal'])}</article>
          <article class="info-card"><h2>Методология</h2>{text_to_paragraphs(about['methodology'])}</article>
          <article class="info-card"><h2>Контакты</h2>{text_to_paragraphs(about['contacts'])}</article>
          <article class="info-card"><h2>Учебно-просветительский характер</h2>{text_to_paragraphs(about['education_note'])}</article>
        </section>
        """
        self.respond_html(render_public_layout("О проекте", body, current_section="about"))


    def handle_search(self, query: dict[str, list[str]]) -> None:
        q = query.get("q", [""])[0].strip()
        results_by_section = {slug: [] for slug in SECTION_LABELS}
        if q:
            for case in CasesRepository.list_public(q=q, sort="new"):
                results_by_section[case["section"]].append(case)
        groups = []
        for slug, label in SECTION_LABELS.items():
            cards = "".join(render_case_card(case) for case in results_by_section[slug]) or '<div class="empty small">Нет результатов.</div>'
            groups.append(f'<section class="section-block"><div class="section-head"><h2>{html_escape(label)}</h2><span class="muted">{len(results_by_section[slug])} найдено</span></div><div class="cards-grid">{cards}</div></section>')
        body = f"""
        <section class="page-hero slim">
          <div>
            <div class="eyebrow">Глобальный поиск</div>
            <h1>Результаты по всей библиотеке</h1>
            <p class="lead">Результаты сгруппированы по трем основным разделам сайта.</p>
          </div>
        </section>
        <form method="get" class="filters-panel single-search">
          <div class="search-slot"><input type="search" name="q" value="{html_escape(q)}" placeholder="Введите запрос"></div>
          <button type="submit">Искать</button>
        </form>
        {''.join(groups)}
        """
        self.respond_html(render_public_layout("Поиск", body))

    def handle_admin_login_page(self) -> None:
        body = '''
        <div class="login-card">
          <h1>Вход администратора</h1>
          <form method="post" class="stack">
            <label class="form-field"><span>Логин</span><input type="text" name="login" required></label>
            <label class="form-field"><span>Пароль</span><input type="password" name="password" required></label>
            <button type="submit">Войти</button>
          </form>
          <p class="muted">По умолчанию: admin / admin123</p>
        </div>
        '''
        self.respond_html(render_public_layout("Вход", body))

    def handle_admin_login(self) -> None:
        fields, _ = self.parse_form_data()
        if fields.get("login") == ADMIN_LOGIN and fields.get("password") == ADMIN_PASSWORD:
            token = self.set_session()
            headers = {"Set-Cookie": f"{SESSION_COOKIE}={token}; HttpOnly; Path=/; SameSite=Lax"}
            return self.redirect("/admin", headers=headers)
        self.respond_html(render_public_layout("Вход", '<div class="login-card"><h1>Вход администратора</h1><div class="flash">Неверный логин или пароль.</div><a class="ghost" href="/admin/login">Попробовать снова</a></div>'), status=401)

    def handle_admin_logout(self) -> None:
        token = self.get_session_token()
        if token and token in SESSIONS:
            del SESSIONS[token]
        headers = {"Set-Cookie": f"{SESSION_COOKIE}=deleted; Max-Age=0; Path=/; SameSite=Lax"}
        self.redirect("/admin/login", headers=headers)

    def handle_admin_dashboard(self) -> None:
        if not self.require_admin():
            return
        stats = CasesRepository.stats()
        cards = [
            ("Всего кейсов", stats["all"]),
            ("Опубликовано", stats["published"]),
            ("Черновики", stats["draft"]),
            ("Скрыто", stats["hidden"]),
            ("Россия", stats["russia"]),
            ("Иностранные государства", stats["foreign"]),
            ("Международные организации", stats["intl_orgs"]),
        ]
        cards_html = "".join(f'<div class="stat-card"><span>{html_escape(title)}</span><strong>{value}</strong></div>' for title, value in cards)
        body = f'''
        <section class="page-head"><h1>Статистика</h1><p class="muted">Базовые счетчики по кейсам и статусам.</p></section>
        <div class="stats-grid">{cards_html}</div>
        '''
        self.respond_html(render_admin_layout("Статистика", body))

    def handle_admin_cases(self, query: dict[str, list[str]]) -> None:
        if not self.require_admin():
            return
        q = query.get("q", [""])[0]
        status = query.get("status", [""])[0]
        items = CasesRepository.list_admin(q=q, status=status)
        rows = "".join(
            f'''<tr>
                <td>{html_escape(item['full_name'])}</td>
                <td>{html_escape(SECTION_SHORT.get(item['section'], item['section']))}</td>
                <td><span class="badge status-{html_escape(item['status'])}">{html_escape(item['status'])}</span></td>
                <td>{html_escape(item['updated_at'])}</td>
                <td><a href="/admin/case/{item['id']}">Открыть</a></td>
            </tr>'''
            for item in items
        ) or '<tr><td colspan="5">Ничего не найдено.</td></tr>'
        selected_all = "selected" if not status else ""
        selected_published = "selected" if status == "published" else ""
        selected_draft = "selected" if status == "draft" else ""
        selected_hidden = "selected" if status == "hidden" else ""
        body = f'''
        <section class="page-head admin-head"><div><h1>Список кейсов</h1><p class="muted">Поиск, фильтрация, создание и изменение статуса.</p></div><a class="ghost" href="/admin/case/new">Создать кейс</a></section>
        <form method="get" class="filters single admin-filter">
          <input type="search" name="q" value="{html_escape(q)}" placeholder="Поиск по названию, описанию или slug">
          <select name="status">
            <option value="" {selected_all}>Все статусы</option>
            <option value="published" {selected_published}>published</option>
            <option value="draft" {selected_draft}>draft</option>
            <option value="hidden" {selected_hidden}>hidden</option>
          </select>
          <button type="submit">Фильтровать</button>
        </form>
        <div class="table-wrap">
          <table>
            <thead><tr><th>Кейс</th><th>Раздел</th><th>Статус</th><th>Обновлен</th><th></th></tr></thead>
            <tbody>{rows}</tbody>
          </table>
        </div>
        '''
        self.respond_html(render_admin_layout("Кейсы", body))

    def handle_admin_case_form(self, case_id: int | None, flash: str = "") -> None:
        if not self.require_admin():
            return
        case_payload: dict[str, Any]
        case_row = None
        sources = []
        if case_id is None:
            case_payload = {
                "section": "russia",
                "slug": "",
                "full_name": "",
                "short_description": "",
                "year_or_period": "",
                "amount": "",
                "country": "Россия",
                "organization": "",
                "jurisdiction": "",
                "governance_level": "",
                "risk_sector": "",
                "violation_type": "",
                "case_summary": "",
                "legal_qualification": "",
                "case_progress": "",
                "consequences": "",
                "institutional_effects": "",
                "policy_lessons": "",
                "status": "draft",
                "sources": [],
                "photo_path": "",
            }
            title = "Новый кейс"
        else:
            case_row = CasesRepository.get_case_by_id(case_id)
            if case_row is None:
                return self.respond_text("Кейс не найден", status=404)
            sources = CasesRepository.get_sources(case_id)
            case_payload = dict(case_row)
            case_payload["sources"] = [{"gost_text": src["gost_text"], "url": src["url"] or ""} for src in sources]
            title = f"Редактирование: {case_row['full_name']}"
        status_buttons = ""
        if case_id is not None and case_row is not None:
            status_buttons = f'''
            <div class="status-actions">
              <form method="post" action="/admin/case/{case_id}/status"><input type="hidden" name="status" value="draft"><button class="ghost" type="submit">В черновик</button></form>
              <form method="post" action="/admin/case/{case_id}/status"><input type="hidden" name="status" value="published"><button type="submit">Опубликовать</button></form>
              <form method="post" action="/admin/case/{case_id}/status"><input type="hidden" name="status" value="hidden"><button class="ghost danger" type="submit">Скрыть</button></form>
            </div>
            '''
        body = f'''
        <section class="page-head admin-head"><div><h1>{html_escape(title)}</h1><p class="muted">Форма кейса со всеми основными блоками и источниками.</p></div>{status_buttons}</section>
        <form method="post" enctype="multipart/form-data" class="stack card-form">
          {build_case_form(case_payload)}
        </form>
        '''
        self.respond_html(render_admin_layout(title, body, flash=flash))

    def handle_admin_case_save(self, case_id: int | None) -> None:
        if not self.require_admin():
            return
        fields, files = self.parse_form_data()
        full_name = fields.get("full_name", "").strip()
        short_description = fields.get("short_description", "").strip()
        slug = fields.get("slug", "").strip() or slugify(full_name)
        if not full_name or not short_description:
            return self.handle_admin_case_form(case_id, flash="ФИО и краткое описание обязательны.")
        existing = CasesRepository.get_case_by_id(case_id) if case_id else None
        photo_path = existing["photo_path"] if existing else None
        if "photo" in files and getattr(files["photo"], "size", 0):
            photo_path = save_uploaded_file(files["photo"], PHOTOS_DIR, "photo")
        data = {
            "slug": slugify(slug),
            "section": fields.get("section", "russia"),
            "full_name": full_name,
            "short_description": short_description,
            "photo_path": photo_path,
            "year_or_period": fields.get("year_or_period", "").strip(),
            "amount": fields.get("amount", "").strip(),
            "country": fields.get("country", "").strip(),
            "organization": fields.get("organization", "").strip(),
            "jurisdiction": fields.get("jurisdiction", "").strip(),
            "governance_level": fields.get("governance_level", "").strip(),
            "risk_sector": fields.get("risk_sector", "").strip(),
            "violation_type": fields.get("violation_type", "").strip(),
            "case_summary": fields.get("case_summary", "").strip(),
            "legal_qualification": fields.get("legal_qualification", "").strip(),
            "case_progress": fields.get("case_progress", "").strip(),
            "consequences": fields.get("consequences", "").strip(),
            "institutional_effects": fields.get("institutional_effects", "").strip(),
            "policy_lessons": fields.get("policy_lessons", "").strip(),
            "status": existing["status"] if existing else "draft",
            "published_at": existing["published_at"] if existing else None,
            "sources": parse_sources_text(fields.get("sources_text", "")),
        }
        saved_id = CasesRepository.upsert_case(data, case_id=case_id)
        self.redirect(f"/admin/case/{saved_id}")

    def handle_admin_case_status(self, case_id: int) -> None:
        if not self.require_admin():
            return
        fields, _ = self.parse_form_data()
        status = fields.get("status", "draft")
        if status not in {"draft", "published", "hidden"}:
            return self.respond_text("Неверный статус", status=400)
        CasesRepository.change_status(case_id, status)
        self.redirect(f"/admin/case/{case_id}")

    def handle_admin_import_page(self, flash: str = "") -> None:
        if not self.require_admin():
            return
        body = '''
        <section class="page-head"><h1>Импорт из Word</h1><p class="muted">Best effort: система пробует разложить .docx по структурным блокам и открыть редактор с заполненными полями.</p></section>
        <form method="post" enctype="multipart/form-data" class="stack card-form narrow">
          <label class="form-field"><span>Файл .docx</span><input type="file" name="docx_file" accept=".docx" required></label>
          <button type="submit">Импортировать</button>
        </form>
        '''
        self.respond_html(render_admin_layout("Импорт из Word", body, flash=flash))

    def handle_admin_import_submit(self) -> None:
        if not self.require_admin():
            return
        _, files = self.parse_form_data()
        file_obj = files.get("docx_file")
        if not file_obj or not getattr(file_obj, "size", 0):
            return self.handle_admin_import_page(flash="Нужно выбрать .docx файл.")
        file_obj.file_object.seek(0)
        content = file_obj.file_object.read()
        import_name = save_uploaded_file(file_obj, IMPORTS_DIR, "import")
        parsed = parse_docx_bytes(content)
        parsed["slug"] = slugify(parsed["full_name"] or import_name)
        parsed["status"] = "draft"
        parsed["short_description"] = parsed.get("short_description") or "Кейс импортирован из Word и требует проверки администратора."
        case_id = CasesRepository.upsert_case(parsed)
        self.redirect(f"/admin/case/{case_id}")

    def handle_admin_about_page(self, flash: str = "") -> None:
        if not self.require_admin():
            return
        about = CasesRepository.get_about()
        body = f'''
        <section class="page-head"><h1>Редактирование страницы «О проекте»</h1><p class="muted">Текст обновляется без участия программиста.</p></section>
        <form method="post" class="stack card-form">
          {textarea_input("goal", "Цель проекта", about['goal'], rows=5)}
          {textarea_input("methodology", "Методология", about['methodology'], rows=6)}
          {textarea_input("contacts", "Контакты", about['contacts'], rows=4)}
          {textarea_input("education_note", "Учебно-просветительский характер", about['education_note'], rows=4)}
          <div class="form-actions"><button type="submit">Сохранить</button></div>
        </form>
        '''
        self.respond_html(render_admin_layout("О проекте", body, flash=flash))

    def handle_admin_about_save(self) -> None:
        if not self.require_admin():
            return
        fields, _ = self.parse_form_data()
        CasesRepository.update_about(
            fields.get("goal", "").strip(),
            fields.get("methodology", "").strip(),
            fields.get("contacts", "").strip(),
            fields.get("education_note", "").strip(),
        )
        self.handle_admin_about_page(flash="Страница обновлена.")

    def handle_admin_dictionaries_page(self, flash: str = "") -> None:
        if not self.require_admin():
            return
        countries = CasesRepository.list_dictionary("countries")
        organizations = CasesRepository.list_dictionary("organizations")
        violations = CasesRepository.list_dictionary("violation_types")
        def block(title: str, key: str, values: list[str]) -> str:
            items = "".join(f"<li>{html_escape(value)}</li>" for value in values) or "<li>Пока пусто.</li>"
            return f'''
            <section class="dictionary-block">
              <h2>{html_escape(title)}</h2>
              <form method="post" class="inline-form">
                <input type="hidden" name="dictionary" value="{html_escape(key)}">
                <input type="text" name="value" placeholder="Новое значение" required>
                <button type="submit">Добавить</button>
              </form>
              <ul class="dictionary-list">{items}</ul>
            </section>
            '''
        body = f'''
        <section class="page-head"><h1>Справочники</h1><p class="muted">Минимальная реализация для стран, организаций и типов нарушений.</p></section>
        <div class="dictionary-grid">
          {block('Страны', 'countries', countries)}
          {block('Организации', 'organizations', organizations)}
          {block('Типы нарушений', 'violation_types', violations)}
        </div>
        '''
        self.respond_html(render_admin_layout("Справочники", body, flash=flash))

    def handle_admin_dictionary_add(self) -> None:
        if not self.require_admin():
            return
        fields, _ = self.parse_form_data()
        dictionary = fields.get("dictionary", "")
        value = fields.get("value", "").strip()
        if dictionary not in {"countries", "organizations", "violation_types"}:
            return self.handle_admin_dictionaries_page(flash="Неизвестный справочник.")
        if not value:
            return self.handle_admin_dictionaries_page(flash="Нужно указать значение.")
        CasesRepository.add_dictionary_value(dictionary, value)
        self.handle_admin_dictionaries_page(flash="Значение добавлено.")



def create_style() -> None:
    STATIC_DIR.mkdir(parents=True, exist_ok=True)
    style_path = STATIC_DIR / "style.css"
    if style_path.exists():
        return
    style_path.write_text("", encoding="utf-8")


def main() -> None:
    ensure_dirs()
    create_style()
    DB.init()
    server = ThreadingHTTPServer((HOST, PORT), AppHandler)
    print(f"Server started on http://{HOST}:{PORT}")
    server.serve_forever()


if __name__ == "__main__":
    main()
