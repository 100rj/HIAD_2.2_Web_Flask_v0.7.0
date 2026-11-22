from __future__ import annotations

import hashlib
import json
import re
import sqlite3
import threading
import uuid
import webbrowser
import io
import zipfile
from datetime import datetime
from pathlib import Path

import pandas as pd
from flask import Flask, flash, redirect, render_template, request, send_file, url_for
from werkzeug.utils import secure_filename

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
UPLOAD_DIR = DATA_DIR / "uploads"
DB_PATH = DATA_DIR / "hiad_web.sqlite3"
CODEBOOK_PATH = BASE_DIR / "codebook.json"
ALLOWED_EXTENSIONS = {".xlsx", ".xlsm"}

PAPER_DEFAULTS = {
    "events_sheet": "EVENTS",
    "classes_sheet": "Classes for EVENTS",
    "columns": {
        "event_id": "Event ID",
        "quality_level": "Q",
        "event_title": "Event Title",
        "event_description": "Event full description",
        "summary_root_causes": "Summary root causes",
        "root_cause_analysis": "Root CAUSE analysis",
    },
}

FIELD_LABELS = {
    "event_id": "사건 ID",
    "quality_level": "자료 품질등급",
    "event_title": "사건 제목",
    "event_description": "사건 상세 설명",
    "summary_root_causes": "원본 요약 원인",
    "root_cause_analysis": "원인분석 본문",
}

EHSP_LABELS = [
    {"code": "EHSP_C01", "name": "System design error", "korean": "시스템 설계 오류"},
    {"code": "EHSP_C02", "name": "Material/ manufacturing error", "korean": "재료·제조 오류"},
    {"code": "EHSP_C03", "name": "Installation error", "korean": "설치 오류"},
    {"code": "EHSP_C04", "name": "Job factors", "korean": "작업 요인"},
    {"code": "EHSP_C05", "name": "Human factors", "korean": "인적 요인"},
    {"code": "EHSP_C06", "name": "Management factors", "korean": "관리 요인"},
    {"code": "EHSP_C07", "name": "Environment", "korean": "환경 요인"},
    {"code": "EHSP_C08", "name": "Unknown", "korean": "원인 불명"},
]

app = Flask(__name__)
app.config.update(SECRET_KEY="hiad-2-2-local-step-2", MAX_CONTENT_LENGTH=100 * 1024 * 1024)


def get_db() -> sqlite3.Connection:
    DATA_DIR.mkdir(exist_ok=True)
    connection = sqlite3.connect(DB_PATH)
    connection.row_factory = sqlite3.Row
    return connection


def init_db() -> None:
    UPLOAD_DIR.mkdir(parents=True, exist_ok=True)
    with get_db() as db:
        db.executescript(
            """
            CREATE TABLE IF NOT EXISTS projects (
                id TEXT PRIMARY KEY,
                name TEXT NOT NULL,
                description TEXT NOT NULL DEFAULT '',
                status TEXT NOT NULL DEFAULT 'created',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            );
            CREATE TABLE IF NOT EXISTS uploads (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                project_id TEXT NOT NULL UNIQUE,
                original_name TEXT NOT NULL,
                stored_path TEXT NOT NULL,
                sha256 TEXT NOT NULL,
                size_bytes INTEGER NOT NULL,
                workbook_json TEXT NOT NULL,
                uploaded_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id)
            );
            CREATE TABLE IF NOT EXISTS selections (
                project_id TEXT PRIMARY KEY,
                selection_json TEXT NOT NULL,
                baseline_json TEXT NOT NULL,
                is_paper_baseline INTEGER NOT NULL,
                confirmed_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id)
            );
            CREATE TABLE IF NOT EXISTS validations (
                project_id TEXT PRIMARY KEY,
                validation_json TEXT NOT NULL,
                status TEXT NOT NULL,
                error_count INTEGER NOT NULL,
                warning_count INTEGER NOT NULL,
                validated_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id)
            );
            CREATE TABLE IF NOT EXISTS standardizations (
                project_id TEXT PRIMARY KEY,
                result_json TEXT NOT NULL,
                rows_json TEXT NOT NULL,
                status TEXT NOT NULL,
                standardized_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id)
            );
            CREATE TABLE IF NOT EXISTS candidates (
                project_id TEXT PRIMARY KEY,
                result_json TEXT NOT NULL,
                candidates_json TEXT NOT NULL,
                evidence_json TEXT NOT NULL,
                comparisons_json TEXT NOT NULL,
                status TEXT NOT NULL,
                generated_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id)
            );
            CREATE TABLE IF NOT EXISTS comparison_reviews (
                project_id TEXT NOT NULL,
                event_id TEXT NOT NULL,
                decision TEXT NOT NULL,
                final_labels TEXT NOT NULL,
                evidence_type TEXT NOT NULL,
                reason TEXT NOT NULL,
                memo TEXT NOT NULL DEFAULT '',
                reviewer_id TEXT NOT NULL,
                reviewed_at TEXT NOT NULL,
                PRIMARY KEY(project_id, event_id),
                FOREIGN KEY(project_id) REFERENCES projects(id)
            );
            CREATE TABLE IF NOT EXISTS candidate_omission_reviews (
                project_id TEXT NOT NULL,
                event_id TEXT NOT NULL,
                decision TEXT NOT NULL,
                final_labels TEXT NOT NULL,
                evidence_type TEXT NOT NULL,
                reason TEXT NOT NULL,
                memo TEXT NOT NULL DEFAULT '',
                reviewer_id TEXT NOT NULL,
                reviewed_at TEXT NOT NULL,
                PRIMARY KEY(project_id, event_id),
                FOREIGN KEY(project_id) REFERENCES projects(id)
            );
            CREATE TABLE IF NOT EXISTS final_builds (
                project_id TEXT PRIMARY KEY,
                summary_json TEXT NOT NULL,
                generated_at TEXT NOT NULL,
                FOREIGN KEY(project_id) REFERENCES projects(id)
            );
            """
        )


def now_text() -> str:
    return datetime.now().astimezone().isoformat(timespec="seconds")


def workbook_profile(path: Path) -> dict:
    excel = pd.ExcelFile(path, engine="openpyxl")
    sheets = []
    for sheet_name in excel.sheet_names:
        frame = pd.read_excel(path, sheet_name=sheet_name, nrows=0, engine="openpyxl")
        full = pd.read_excel(path, sheet_name=sheet_name, usecols=[0], engine="openpyxl")
        sheets.append(
            {
                "name": sheet_name,
                "rows": int(len(full)),
                "columns": [str(column) for column in frame.columns],
                "column_count": int(len(frame.columns)),
            }
        )
    return {"sheets": sheets, "sheet_count": len(sheets)}


def safe_original_name(filename: str) -> str:
    name = Path(filename).name.strip()
    return name if name else "uploaded.xlsx"


def stored_name(project_id: str, filename: str) -> str:
    suffix = Path(filename).suffix.lower()
    stem = secure_filename(Path(filename).stem) or "hiad"
    return f"{project_id}_{stem}{suffix}"


def load_project(project_id: str):
    with get_db() as db:
        return db.execute("SELECT * FROM projects WHERE id = ?", (project_id,)).fetchone()


def load_upload(project_id: str):
    with get_db() as db:
        return db.execute("SELECT * FROM uploads WHERE project_id = ?", (project_id,)).fetchone()


def current_upload_path(upload) -> Path:
    """이전 버전의 data 폴더를 옮겨도 복사된 입력파일을 우선 사용한다."""
    copied_path = UPLOAD_DIR / Path(upload["stored_path"]).name
    return copied_path if copied_path.exists() else Path(upload["stored_path"])


def load_selection(project_id: str):
    with get_db() as db:
        return db.execute("SELECT * FROM selections WHERE project_id = ?", (project_id,)).fetchone()


def load_validation(project_id: str):
    with get_db() as db:
        return db.execute("SELECT * FROM validations WHERE project_id = ?", (project_id,)).fetchone()


def load_standardization(project_id: str):
    with get_db() as db:
        return db.execute("SELECT * FROM standardizations WHERE project_id = ?", (project_id,)).fetchone()


def load_candidates(project_id: str):
    with get_db() as db:
        return db.execute("SELECT * FROM candidates WHERE project_id = ?", (project_id,)).fetchone()


def load_comparison_reviews(project_id: str) -> list[sqlite3.Row]:
    with get_db() as db:
        return db.execute(
            "SELECT * FROM comparison_reviews WHERE project_id = ? ORDER BY reviewed_at, event_id",
            (project_id,),
        ).fetchall()


def review_population(comparison_rows: list[dict]) -> list[dict]:
    review_types = {"mixed_difference", "candidate_addition", "candidate_omission"}
    return [row for row in comparison_rows if row["comparison_type"] in review_types]


def split_codes(value: str) -> set[str]:
    return {code for code in (value or "").split("|") if code}


def natural_event_key(value: str):
    return (0, int(value)) if str(value).isdigit() else (1, str(value))


def is_blank(series: pd.Series) -> pd.Series:
    return series.isna() | series.astype(str).str.strip().isin({"", "nan", "None", "NaT"})


def add_check(checks: list[dict], code: str, title: str, severity: str, passed: bool,
              detail: str, count: int = 0, sample_ids: list | None = None) -> None:
    checks.append({
        "code": code, "title": title, "severity": severity, "passed": bool(passed),
        "detail": detail, "count": int(count), "sample_ids": [str(x) for x in (sample_ids or [])[:10]],
    })


def validate_input(path: Path, upload, selection: dict) -> dict:
    checks: list[dict] = []
    selected_columns = selection["columns"]
    events_sheet, classes_sheet = selection["events_sheet"], selection["classes_sheet"]
    actual_hash = hashlib.sha256(path.read_bytes()).hexdigest()
    hash_ok = actual_hash == upload["sha256"]
    add_check(checks, "file_hash", "원본 파일 무결성", "error", hash_ok,
              "업로드 당시 SHA-256과 현재 파일이 일치합니다." if hash_ok else "업로드 이후 원본 파일이 변경되었습니다.", 0 if hash_ok else 1)

    try:
        excel = pd.ExcelFile(path, engine="openpyxl")
    except Exception as exc:
        add_check(checks, "workbook_read", "Excel 파일 읽기", "error", False, f"Excel 파일을 읽을 수 없습니다: {exc}", 1)
        return build_validation_result(checks, 0, 0, 0, {})

    missing_sheets = [s for s in (events_sheet, classes_sheet) if s not in excel.sheet_names]
    add_check(checks, "required_sheets", "필수 시트", "error", not missing_sheets,
              "선택한 사건·분류기준 시트가 모두 존재합니다." if not missing_sheets else f"누락 시트: {', '.join(missing_sheets)}",
              len(missing_sheets))
    if missing_sheets:
        return build_validation_result(checks, 0, 0, 0, {})

    frame = pd.read_excel(path, sheet_name=events_sheet, engine="openpyxl")
    missing_columns = [name for name in selected_columns.values() if name not in frame.columns]
    add_check(checks, "required_columns", "필수 분석항목", "error", not missing_columns,
              "선택한 6개 분석항목이 모두 존재합니다." if not missing_columns else f"누락 항목: {', '.join(missing_columns)}",
              len(missing_columns))
    if missing_columns:
        return build_validation_result(checks, len(frame), 0, 0, {})

    total_rows = int(len(frame))
    add_check(checks, "event_rows", "사건 자료 행", "error", total_rows > 0,
              f"사건 자료 {total_rows:,}행을 확인했습니다." if total_rows else "사건 시트에 자료 행이 없습니다.", 0 if total_rows else 1)

    id_col = selected_columns["event_id"]
    id_blank = is_blank(frame[id_col])
    normalized_id = frame[id_col].astype(str).str.strip()
    id_dup = normalized_id.duplicated(keep=False) & ~id_blank
    add_check(checks, "missing_event_id", "사건 ID 결측", "error", not id_blank.any(),
              "모든 사건에 ID가 있습니다." if not id_blank.any() else f"사건 ID가 없는 행이 {int(id_blank.sum()):,}건입니다.",
              int(id_blank.sum()), (frame.index[id_blank] + 2).tolist())
    add_check(checks, "duplicate_event_id", "사건 ID 중복", "error", not id_dup.any(),
              "사건 ID가 모두 고유합니다." if not id_dup.any() else f"중복 ID 관련 행이 {int(id_dup.sum()):,}건입니다.",
              int(id_dup.sum()), frame.loc[id_dup, id_col].tolist())

    q_col = selected_columns["quality_level"]
    q_blank = is_blank(frame[q_col])
    q_num = pd.to_numeric(frame[q_col], errors="coerce")
    q_invalid = ~q_blank & ~q_num.isin([2, 3, 4, 5])
    add_check(checks, "missing_quality", "품질등급 결측", "error", not q_blank.any(),
              "모든 사건에 품질등급이 있습니다." if not q_blank.any() else f"품질등급 결측이 {int(q_blank.sum()):,}건입니다.",
              int(q_blank.sum()), frame.loc[q_blank, id_col].tolist())
    add_check(checks, "invalid_quality", "품질등급 허용값", "error", not q_invalid.any(),
              "모든 품질등급이 허용값 2·3·4·5에 해당합니다." if not q_invalid.any() else f"허용값 이외의 품질등급이 {int(q_invalid.sum()):,}건입니다.",
              int(q_invalid.sum()), frame.loc[q_invalid, id_col].tolist())

    text_keys = ["event_title", "event_description", "summary_root_causes", "root_cause_analysis"]
    missing_by_field = {}
    for key in text_keys:
        col = selected_columns[key]
        blank = is_blank(frame[col])
        count = int(blank.sum())
        missing_by_field[key] = count
        severity = "warning"
        add_check(checks, f"missing_{key}", f"{FIELD_LABELS[key]} 결측", severity, count == 0,
                  f"결측값이 없습니다." if count == 0 else f"결측값이 {count:,}건입니다. 후속 판정 시 정보 부족 여부를 확인해야 합니다.",
                  count, frame.loc[blank, id_col].tolist())

    text_present = pd.DataFrame({key: ~is_blank(frame[selected_columns[key]]) for key in text_keys}).any(axis=1)
    no_text = ~text_present
    add_check(checks, "missing_all_text", "분석 텍스트 전체 결측", "error", not no_text.any(),
              "모든 사건에 하나 이상의 분석 텍스트가 있습니다." if not no_text.any() else f"분석 텍스트가 모두 비어 있는 사건이 {int(no_text.sum()):,}건입니다.",
              int(no_text.sum()), frame.loc[no_text, id_col].tolist())

    valid_row = ~id_blank & ~id_dup & ~q_blank & ~q_invalid & text_present
    analyzable_rows = int(valid_row.sum())
    excluded_rows = total_rows - analyzable_rows
    q_distribution = {str(int(k) if float(k).is_integer() else k): int(v) for k, v in q_num.dropna().value_counts().sort_index().items()}
    return build_validation_result(checks, total_rows, analyzable_rows, excluded_rows, {
        "quality_distribution": q_distribution,
        "missing_by_field": missing_by_field,
        "column_types": {key: str(frame[col].dtype) for key, col in selected_columns.items()},
    })


def build_validation_result(checks: list[dict], total_rows: int, analyzable_rows: int,
                            excluded_rows: int, profile: dict) -> dict:
    errors = sum(1 for c in checks if not c["passed"] and c["severity"] == "error")
    warnings = sum(1 for c in checks if not c["passed"] and c["severity"] == "warning")
    return {
        "program": "HIAD 2.2 Web Dataset Builder", "step": 2,
        "status": "passed" if errors == 0 else "failed", "error_count": errors,
        "warning_count": warnings, "total_rows": total_rows, "analyzable_rows": analyzable_rows,
        "excluded_rows": excluded_rows, "checks": checks, "profile": profile, "validated_at": now_text(),
    }


def canonical_cause(value) -> str:
    if value is None or pd.isna(value):
        return ""
    return re.sub(r"\s+", " ", str(value)).strip().casefold()


def standardize_original_labels(path: Path, selection: dict) -> tuple[dict, list[dict]]:
    columns = selection["columns"]
    frame = pd.read_excel(path, sheet_name=selection["events_sheet"],
                          usecols=[columns["event_id"], columns["summary_root_causes"]], engine="openpyxl")
    lookup = {canonical_cause(item["name"]): item for item in EHSP_LABELS}
    distribution = {item["code"]: 0 for item in EHSP_LABELS}
    rows, unmapped_values = [], {}
    mapped_events = multi_events = unknown_exclusive_issues = total_positive = 0
    for _, source in frame.iterrows():
        raw_value = source[columns["summary_root_causes"]]
        raw = "" if pd.isna(raw_value) else str(raw_value)
        parts = [part.strip() for part in re.split(r"[\r\n]+", raw) if part.strip()]
        codes, names, unmapped = [], [], []
        for part in parts:
            item = lookup.get(canonical_cause(part))
            if item and item["code"] not in codes:
                codes.append(item["code"])
                names.append(item["name"])
            elif not item:
                unmapped.append(part)
                unmapped_values[part] = unmapped_values.get(part, 0) + 1
        if codes and not unmapped:
            mapped_events += 1
        if len(codes) > 1:
            multi_events += 1
        if "EHSP_C08" in codes and len(codes) > 1:
            unknown_exclusive_issues += 1
        for code in codes:
            distribution[code] += 1
            total_positive += 1
        mapping_status = "mapped" if codes and not unmapped else ("unmapped" if unmapped else "blank")
        if "EHSP_C08" in codes and len(codes) > 1:
            mapping_status = "unknown_exclusivity_issue"
        rows.append({
            "event_id": str(source[columns["event_id"]]).strip(),
            "original_summary_root_causes": raw,
            "standard_codes": "|".join(codes), "standard_names": "|".join(names),
            "label_count": len(codes),
            "mapping_status": mapping_status,
            "unmapped_values": "|".join(unmapped),
            **{item["code"]: int(item["code"] in codes) for item in EHSP_LABELS},
        })
    unmapped_events = sum(row["mapping_status"] == "unmapped" for row in rows)
    blank_events = sum(row["mapping_status"] == "blank" for row in rows)
    status = "passed" if unmapped_events == 0 and blank_events == 0 and unknown_exclusive_issues == 0 else "review_required"
    result = {
        "program": "HIAD 2.2 Web Dataset Builder", "step": 3, "status": status,
        "total_events": len(rows), "mapped_events": mapped_events, "unmapped_events": unmapped_events,
        "blank_events": blank_events, "multi_label_events": multi_events,
        "single_label_events": mapped_events - multi_events, "total_positive_labels": total_positive,
        "unknown_exclusivity_issues": unknown_exclusive_issues, "distribution": distribution,
        "unmapped_value_distribution": unmapped_values, "labels": EHSP_LABELS,
        "standardized_at": now_text(),
        "method": "공백·대소문자·줄바꿈을 정규화한 뒤 EHSP 공식 명칭과 정확히 대응",
    }
    return result, rows


def normalize_text(value) -> str:
    if value is None or pd.isna(value):
        return ""
    text = str(value).replace("_x000D_", "\n").replace("_x000B_", " ")
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r" *\n *", "\n", text)
    text = re.sub(r"\n{3,}", "\n\n", text)
    return text.strip()


def sentence_fragments(text: str) -> list[str]:
    return [part.strip() for part in re.split(r"(?<=[.!?;])\s+|\n+", text) if part.strip()]


def has_local_negation(text: str, start: int, window: int = 5) -> bool:
    tokens = re.findall(r"[a-z]+(?:'[a-z]+)?", text[:start].casefold())[-window:]
    return bool(re.search(r"\b(no|not|never|without|neither|nor|ruled out)\b", " ".join(tokens)))


def infer_candidates(analysis, description, labels: list[dict], negation_window: int) -> dict:
    sources = [("root_cause_analysis", normalize_text(analysis), "high"),
               ("event_description", normalize_text(description), "medium")]
    hits, positive, unknown_hit = [], set(), False
    for item in labels:
        for rule in item.get("rules", []):
            pattern = re.compile(rule["pattern"], re.IGNORECASE)
            for source_name, source_text, confidence in sources:
                for fragment in sentence_fragments(source_text):
                    match = pattern.search(fragment)
                    if not match or has_local_negation(fragment, match.start(), negation_window):
                        continue
                    hit = {"label_id": item["label_id"], "rule_id": rule["rule_id"],
                           "source_field": source_name, "confidence": confidence,
                           "evidence_type": rule["evidence"], "evidence_text": fragment[:500]}
                    if hit not in hits:
                        hits.append(hit)
                    if item["label_id"] == "EHSP_C08":
                        unknown_hit = True
                    else:
                        positive.add(item["label_id"])
    analysis_empty = not normalize_text(analysis)
    if positive:
        candidate_ids = sorted(positive)
    elif unknown_hit or analysis_empty:
        candidate_ids = ["EHSP_C08"]
        if analysis_empty and not unknown_hit:
            hits.append({"label_id": "EHSP_C08", "rule_id": "C08_R02",
                         "source_field": "root_cause_analysis", "confidence": "high",
                         "evidence_type": "root-cause-analysis field empty", "evidence_text": ""})
    else:
        candidate_ids = []
    return {"candidate_ids": candidate_ids, "hits": hits}


def comparison_type(confirmed: set[str], candidates: set[str]) -> str:
    if confirmed == candidates:
        return "exact_match"
    if not candidates:
        return "no_candidate"
    added, omitted = candidates - confirmed, confirmed - candidates
    if added and omitted:
        return "mixed_difference"
    return "candidate_addition" if added else "candidate_omission"


def generate_candidate_labels(path: Path, selection: dict, standardized_rows: list[dict]) -> tuple[dict, list[dict], list[dict], list[dict]]:
    codebook = json.loads(CODEBOOK_PATH.read_text(encoding="utf-8"))
    labels = codebook["labels"]
    negation_window = int(codebook.get("decision_policy", {}).get("negation_window_tokens", 5))
    columns = selection["columns"]
    frame = pd.read_excel(path, sheet_name=selection["events_sheet"],
                          usecols=[columns["event_id"], columns["event_description"], columns["root_cause_analysis"]],
                          engine="openpyxl")
    confirmed_by_id = {row["event_id"]: set(filter(None, row["standard_codes"].split("|"))) for row in standardized_rows}
    candidate_rows, evidence_rows, comparison_rows = [], [], []
    distribution = {item["label_id"]: 0 for item in labels}
    for _, source in frame.iterrows():
        event_id = str(source[columns["event_id"]]).strip()
        inference = infer_candidates(source[columns["root_cause_analysis"]], source[columns["event_description"]], labels, negation_window)
        candidate_ids = inference["candidate_ids"]
        candidate_set, confirmed = set(candidate_ids), confirmed_by_id.get(event_id, set())
        compare = comparison_type(confirmed, candidate_set)
        for code in candidate_ids:
            distribution[code] += 1
        candidate_rows.append({"event_id": event_id, "candidate_labels": "|".join(candidate_ids),
                               "candidate_count": len(candidate_ids),
                               "candidate_status": "candidate_generated" if candidate_ids else "no_candidate",
                               "evidence_hit_count": len(inference["hits"]), "codebook_version": codebook["version"]})
        for hit in inference["hits"]:
            evidence_rows.append({"event_id": event_id, **hit, "codebook_version": codebook["version"]})
        comparison_rows.append({"event_id": event_id, "confirmed_labels": "|".join(sorted(confirmed)),
                                "candidate_labels": "|".join(sorted(candidate_set)),
                                "candidate_additions": "|".join(sorted(candidate_set - confirmed)),
                                "candidate_omissions": "|".join(sorted(confirmed - candidate_set)),
                                "comparison_type": compare,
                                "review_required": int(compare != "exact_match")})
    counts = pd.Series([row["comparison_type"] for row in comparison_rows]).value_counts().to_dict()
    generated_at = now_text()
    result = {"program": "HIAD 2.2 Web Dataset Builder", "step": 4, "status": "completed",
              "total_events": len(candidate_rows),
              "candidate_events_with_labels": sum(bool(row["candidate_labels"]) for row in candidate_rows),
              "no_candidate_events": sum(not bool(row["candidate_labels"]) for row in candidate_rows),
              "candidate_evidence_hits": len(evidence_rows),
              "candidate_positive_labels": sum(row["candidate_count"] for row in candidate_rows),
              "candidate_distribution": distribution,
              "comparison_distribution": counts,
              "comparison_review_required": sum(row["review_required"] for row in comparison_rows),
              "phase4_difference_review_population": sum(row["comparison_type"] in {"mixed_difference", "candidate_addition", "candidate_omission"} for row in comparison_rows),
              "codebook_version": codebook["version"],
              "codebook_sha256": hashlib.sha256(CODEBOOK_PATH.read_bytes()).hexdigest(),
              "generated_at": generated_at,
              "method": "v1.0.0 코드북의 포함·제외·부정어·우선순위 규칙에 따른 결정론적 후보 생성"}
    return result, candidate_rows, evidence_rows, comparison_rows


@app.get("/")
def index():
    with get_db() as db:
        projects = db.execute("SELECT * FROM projects ORDER BY updated_at DESC").fetchall()
    return render_template("index.html", projects=projects)


@app.post("/projects")
def create_project():
    name = request.form.get("name", "").strip()
    if not name:
        flash("프로젝트 이름을 입력해 주세요.", "error")
        return redirect(url_for("index"))
    project_id = uuid.uuid4().hex[:12]
    timestamp = now_text()
    with get_db() as db:
        db.execute(
            "INSERT INTO projects(id, name, description, created_at, updated_at) VALUES (?, ?, ?, ?, ?)",
            (project_id, name, request.form.get("description", "").strip(), timestamp, timestamp),
        )
    return redirect(url_for("project", project_id=project_id))


@app.get("/projects/<project_id>")
def project(project_id: str):
    item = load_project(project_id)
    if item is None:
        return "프로젝트를 찾을 수 없습니다.", 404
    upload = load_upload(project_id)
    selection = None
    if upload:
        with get_db() as db:
            selection = db.execute("SELECT * FROM selections WHERE project_id = ?", (project_id,)).fetchone()
    validation = load_validation(project_id) if upload else None
    standardization = load_standardization(project_id) if validation else None
    candidates = load_candidates(project_id) if standardization else None
    comparison_reviews = load_comparison_reviews(project_id) if candidates else []
    comparison_review_total = 0
    omission_review_total = 0
    omission_reviews = []
    final_build = None
    if candidates:
        comparison_rows = json.loads(candidates["comparisons_json"])
        comparison_review_total = len(review_population(comparison_rows))
        omission_review_total = sum(row.get("comparison_type") == "no_candidate" for row in comparison_rows)
        omission_reviews = load_candidate_omission_reviews(project_id)
        with get_db() as db:
            final_build = db.execute("SELECT * FROM final_builds WHERE project_id=?", (project_id,)).fetchone()
    return render_template("project.html", project=item, upload=upload, selection=selection,
                           validation=validation, standardization=standardization, candidates=candidates,
                           comparison_review_count=len(comparison_reviews),
                           comparison_review_total=comparison_review_total,
                           omission_review_count=len(omission_reviews),
                           omission_review_total=omission_review_total, final_build=final_build)


@app.post("/projects/<project_id>/upload")
def upload(project_id: str):
    if load_project(project_id) is None:
        return "프로젝트를 찾을 수 없습니다.", 404
    file = request.files.get("excel_file")
    if not file or not file.filename:
        flash("Excel 파일을 선택해 주세요.", "error")
        return redirect(url_for("project", project_id=project_id))
    original_name = safe_original_name(file.filename)
    suffix = Path(original_name).suffix.lower()
    if suffix not in ALLOWED_EXTENSIONS:
        flash(".xlsx 또는 .xlsm 파일만 입력할 수 있습니다.", "error")
        return redirect(url_for("project", project_id=project_id))
    destination = UPLOAD_DIR / stored_name(project_id, original_name)
    file.save(destination)
    try:
        profile = workbook_profile(destination)
    except Exception as exc:
        destination.unlink(missing_ok=True)
        flash(f"Excel 구조를 읽을 수 없습니다: {exc}", "error")
        return redirect(url_for("project", project_id=project_id))
    digest = hashlib.sha256(destination.read_bytes()).hexdigest()
    timestamp = now_text()
    with get_db() as db:
        old = db.execute("SELECT stored_path FROM uploads WHERE project_id = ?", (project_id,)).fetchone()
        db.execute("DELETE FROM selections WHERE project_id = ?", (project_id,))
        db.execute("DELETE FROM comparison_reviews WHERE project_id = ?", (project_id,))
        db.execute("DELETE FROM candidates WHERE project_id = ?", (project_id,))
        db.execute("DELETE FROM standardizations WHERE project_id = ?", (project_id,))
        db.execute("DELETE FROM validations WHERE project_id = ?", (project_id,))
        db.execute("DELETE FROM uploads WHERE project_id = ?", (project_id,))
        db.execute(
            "INSERT INTO uploads(project_id, original_name, stored_path, sha256, size_bytes, workbook_json, uploaded_at) VALUES (?, ?, ?, ?, ?, ?, ?)",
            (project_id, original_name, str(destination), digest, destination.stat().st_size, json.dumps(profile, ensure_ascii=False), timestamp),
        )
        db.execute("UPDATE projects SET status='uploaded', updated_at=? WHERE id=?", (timestamp, project_id))
        if old and old["stored_path"] != str(destination):
            Path(old["stored_path"]).unlink(missing_ok=True)
    return redirect(url_for("select_columns", project_id=project_id))


@app.get("/projects/<project_id>/select")
def select_columns(project_id: str):
    item, upload = load_project(project_id), load_upload(project_id)
    if item is None or upload is None:
        flash("먼저 Excel 파일을 입력해 주세요.", "error")
        return redirect(url_for("project", project_id=project_id))
    profile = json.loads(upload["workbook_json"])
    sheets = profile["sheets"]
    names = [sheet["name"] for sheet in sheets]
    events_sheet = PAPER_DEFAULTS["events_sheet"] if PAPER_DEFAULTS["events_sheet"] in names else names[0]
    classes_sheet = PAPER_DEFAULTS["classes_sheet"] if PAPER_DEFAULTS["classes_sheet"] in names else (names[1] if len(names) > 1 else names[0])
    selected_info = next(sheet for sheet in sheets if sheet["name"] == events_sheet)
    return render_template(
        "select.html", project=item, upload=upload, sheets=sheets, events_sheet=events_sheet,
        classes_sheet=classes_sheet, selected_columns=selected_info["columns"], defaults=PAPER_DEFAULTS,
        field_labels=FIELD_LABELS,
    )


@app.post("/projects/<project_id>/select")
def save_selection(project_id: str):
    upload = load_upload(project_id)
    if upload is None:
        return "업로드 자료를 찾을 수 없습니다.", 404
    profile = json.loads(upload["workbook_json"])
    sheet_map = {sheet["name"]: sheet for sheet in profile["sheets"]}
    events_sheet = request.form.get("events_sheet", "")
    classes_sheet = request.form.get("classes_sheet", "")
    errors = []
    if events_sheet not in sheet_map:
        errors.append("사건 시트 선택값이 올바르지 않습니다.")
    if classes_sheet not in sheet_map:
        errors.append("분류기준 시트 선택값이 올바르지 않습니다.")
    columns = {key: request.form.get(f"column_{key}", "").strip() for key in FIELD_LABELS}
    if events_sheet in sheet_map:
        available = set(sheet_map[events_sheet]["columns"])
        for key, value in columns.items():
            if not value:
                errors.append(f"{FIELD_LABELS[key]} 항목을 선택해 주세요.")
            elif value not in available:
                errors.append(f"{FIELD_LABELS[key]} 항목이 선택한 사건 시트에 없습니다.")
        chosen = list(columns.values())
        if len(chosen) != len(set(chosen)):
            errors.append("하나의 항목을 여러 역할에 중복 선택할 수 없습니다.")
    if errors:
        for error in errors:
            flash(error, "error")
        return redirect(url_for("select_columns", project_id=project_id))
    selection = {"events_sheet": events_sheet, "classes_sheet": classes_sheet, "columns": columns}
    baseline = PAPER_DEFAULTS
    is_baseline = int(selection == baseline)
    timestamp = now_text()
    with get_db() as db:
        db.execute("DELETE FROM comparison_reviews WHERE project_id = ?", (project_id,))
        db.execute("DELETE FROM candidates WHERE project_id = ?", (project_id,))
        db.execute("DELETE FROM standardizations WHERE project_id = ?", (project_id,))
        db.execute("DELETE FROM validations WHERE project_id = ?", (project_id,))
        db.execute("DELETE FROM selections WHERE project_id = ?", (project_id,))
        db.execute(
            "INSERT INTO selections(project_id, selection_json, baseline_json, is_paper_baseline, confirmed_at) VALUES (?, ?, ?, ?, ?)",
            (project_id, json.dumps(selection, ensure_ascii=False), json.dumps(baseline, ensure_ascii=False), is_baseline, timestamp),
        )
        db.execute("UPDATE projects SET status='step1_complete', updated_at=? WHERE id=?", (timestamp, project_id))
    return redirect(url_for("selection_result", project_id=project_id))


@app.get("/projects/<project_id>/selection-result")
def selection_result(project_id: str):
    item, upload = load_project(project_id), load_upload(project_id)
    with get_db() as db:
        selection = db.execute("SELECT * FROM selections WHERE project_id = ?", (project_id,)).fetchone()
    if not item or not upload or not selection:
        return redirect(url_for("project", project_id=project_id))
    validation = load_validation(project_id)
    return render_template("result.html", project=item, upload=upload, selection=selection, validation=validation,
                           selected=json.loads(selection["selection_json"]), field_labels=FIELD_LABELS)


@app.route("/projects/<project_id>/validate", methods=["GET", "POST"])
def validation(project_id: str):
    item, upload, selection_row = load_project(project_id), load_upload(project_id), load_selection(project_id)
    if not item or not upload or not selection_row:
        flash("먼저 1단계 입력 구조 설정을 확정해 주세요.", "error")
        return redirect(url_for("project", project_id=project_id))
    if request.method == "POST":
        result = validate_input(current_upload_path(upload), upload, json.loads(selection_row["selection_json"]))
        with get_db() as db:
            db.execute("DELETE FROM standardizations WHERE project_id = ?", (project_id,))
            db.execute("DELETE FROM comparison_reviews WHERE project_id = ?", (project_id,))
            db.execute("DELETE FROM candidates WHERE project_id = ?", (project_id,))
            db.execute("DELETE FROM validations WHERE project_id = ?", (project_id,))
            db.execute(
                "INSERT INTO validations(project_id, validation_json, status, error_count, warning_count, validated_at) VALUES (?, ?, ?, ?, ?, ?)",
                (project_id, json.dumps(result, ensure_ascii=False), result["status"], result["error_count"], result["warning_count"], result["validated_at"]),
            )
            status = "step2_complete" if result["status"] == "passed" else "step2_action_required"
            db.execute("UPDATE projects SET status=?, updated_at=? WHERE id=?", (status, result["validated_at"], project_id))
        return redirect(url_for("validation", project_id=project_id))
    saved = load_validation(project_id)
    result = json.loads(saved["validation_json"]) if saved else None
    return render_template("validation.html", project=item, upload=upload, result=result)


@app.route("/projects/<project_id>/standardize", methods=["GET", "POST"])
def standardization(project_id: str):
    item, upload = load_project(project_id), load_upload(project_id)
    selection_row, validation_row = load_selection(project_id), load_validation(project_id)
    if not item or not upload or not selection_row or not validation_row:
        flash("먼저 2단계 입력자료 검증을 완료해 주세요.", "error")
        return redirect(url_for("project", project_id=project_id))
    validation_result = json.loads(validation_row["validation_json"])
    if validation_result["status"] != "passed":
        flash("2단계 오류를 해결해야 원본라벨 표준화를 실행할 수 있습니다.", "error")
        return redirect(url_for("validation", project_id=project_id))
    if request.method == "POST":
        result, rows = standardize_original_labels(
            current_upload_path(upload), json.loads(selection_row["selection_json"])
        )
        with get_db() as db:
            db.execute("DELETE FROM comparison_reviews WHERE project_id = ?", (project_id,))
            db.execute("DELETE FROM candidates WHERE project_id = ?", (project_id,))
            db.execute("DELETE FROM standardizations WHERE project_id = ?", (project_id,))
            db.execute(
                "INSERT INTO standardizations(project_id,result_json,rows_json,status,standardized_at) VALUES (?,?,?,?,?)",
                (project_id, json.dumps(result, ensure_ascii=False), json.dumps(rows, ensure_ascii=False),
                 result["status"], result["standardized_at"]),
            )
            status = "step3_complete" if result["status"] == "passed" else "step3_review_required"
            db.execute("UPDATE projects SET status=?, updated_at=? WHERE id=?", (status, result["standardized_at"], project_id))
        return redirect(url_for("standardization", project_id=project_id))
    saved = load_standardization(project_id)
    result = json.loads(saved["result_json"]) if saved else None
    rows = json.loads(saved["rows_json"]) if saved else []
    return render_template("standardization.html", project=item, upload=upload, result=result,
                           preview=rows[:30],
                           review_rows=[row for row in rows if row["mapping_status"] != "mapped"][:100])


@app.route("/projects/<project_id>/candidates", methods=["GET", "POST"])
def candidates(project_id: str):
    item, upload, selection_row = load_project(project_id), load_upload(project_id), load_selection(project_id)
    standardization_row = load_standardization(project_id)
    if not item or not upload or not selection_row or not standardization_row:
        flash("먼저 3단계 원본라벨 표준화를 완료해 주세요.", "error")
        return redirect(url_for("project", project_id=project_id))
    if request.method == "POST":
        result, candidate_rows, evidence_rows, comparison_rows = generate_candidate_labels(
            current_upload_path(upload), json.loads(selection_row["selection_json"]),
            json.loads(standardization_row["rows_json"]))
        with get_db() as db:
            db.execute("DELETE FROM comparison_reviews WHERE project_id = ?", (project_id,))
            db.execute("DELETE FROM candidates WHERE project_id = ?", (project_id,))
            db.execute("INSERT INTO candidates(project_id,result_json,candidates_json,evidence_json,comparisons_json,status,generated_at) VALUES (?,?,?,?,?,?,?)",
                       (project_id, json.dumps(result, ensure_ascii=False), json.dumps(candidate_rows, ensure_ascii=False),
                        json.dumps(evidence_rows, ensure_ascii=False), json.dumps(comparison_rows, ensure_ascii=False),
                        result["status"], result["generated_at"]))
            db.execute("UPDATE projects SET status='step4_complete', updated_at=? WHERE id=?", (result["generated_at"], project_id))
        return redirect(url_for("candidates", project_id=project_id))
    saved = load_candidates(project_id)
    result = json.loads(saved["result_json"]) if saved else None
    comparisons = json.loads(saved["comparisons_json"]) if saved else []
    return render_template("candidates.html", project=item, upload=upload, result=result,
                           preview=comparisons[:30], review_rows=[r for r in comparisons if r["review_required"]][:100])


def comparison_review_context(project_id: str) -> tuple[sqlite3.Row, sqlite3.Row, list[dict], dict[str, dict], dict[str, list[dict]]]:
    item, upload, selection_row, saved = (
        load_project(project_id), load_upload(project_id), load_selection(project_id), load_candidates(project_id)
    )
    if not item or not upload or not selection_row or not saved:
        raise LookupError("먼저 4단계 후보라벨 생성을 완료해 주세요.")
    comparisons = review_population(json.loads(saved["comparisons_json"]))
    comparisons.sort(key=lambda row: natural_event_key(row["event_id"]))
    selection = json.loads(selection_row["selection_json"])
    columns = selection["columns"]
    frame = pd.read_excel(
        current_upload_path(upload), sheet_name=selection["events_sheet"],
        usecols=[columns["event_id"], columns["quality_level"], columns["event_title"],
                 columns["event_description"], columns["root_cause_analysis"]], engine="openpyxl",
    )
    event_map = {}
    for _, source in frame.iterrows():
        event_id = str(source[columns["event_id"]]).strip()
        event_map[event_id] = {
            "event_id": event_id,
            "quality_level": "" if pd.isna(source[columns["quality_level"]]) else str(source[columns["quality_level"]]),
            "event_title": normalize_text(source[columns["event_title"]]),
            "event_description": normalize_text(source[columns["event_description"]]),
            "root_cause_analysis": normalize_text(source[columns["root_cause_analysis"]]),
        }
    evidence_map: dict[str, list[dict]] = {}
    for row in json.loads(saved["evidence_json"]):
        evidence_map.setdefault(row["event_id"], []).append(row)
    return item, upload, comparisons, event_map, evidence_map


@app.route("/projects/<project_id>/comparison-review", methods=["GET", "POST"])
def comparison_review(project_id: str):
    try:
        item, upload, population, event_map, evidence_map = comparison_review_context(project_id)
    except LookupError as exc:
        flash(str(exc), "error")
        return redirect(url_for("project", project_id=project_id))

    population_by_id = {row["event_id"]: row for row in population}
    if request.method == "POST":
        event_id = request.form.get("event_id", "").strip()
        comparison = population_by_id.get(event_id)
        if not comparison:
            flash("검토 대상 사건을 찾을 수 없습니다.", "error")
            return redirect(url_for("comparison_review", project_id=project_id))
        decision = request.form.get("decision", "").strip()
        allowed = {"keep_confirmed", "adopt_candidate", "merge", "custom", "exclude"}
        if decision not in allowed:
            flash("검토결정을 선택해 주세요.", "error")
            return redirect(request.form.get("return_url") or url_for("comparison_review", project_id=project_id))
        confirmed, candidate = split_codes(comparison["confirmed_labels"]), split_codes(comparison["candidate_labels"])
        if decision == "keep_confirmed":
            final_labels = confirmed
        elif decision == "adopt_candidate":
            final_labels = candidate
        elif decision == "merge":
            final_labels = confirmed | candidate
        elif decision == "custom":
            final_labels = set(request.form.getlist("final_labels"))
        else:
            final_labels = set()
        valid_codes = {item["code"] for item in EHSP_LABELS}
        reviewer_id = request.form.get("reviewer_id", "").strip()
        evidence_type = request.form.get("evidence_type", "").strip()
        reason = request.form.get("reason", "").strip()
        errors = []
        if not final_labels.issubset(valid_codes):
            errors.append("EHSP C01～C08 이외의 라벨은 저장할 수 없습니다.")
        if decision != "exclude" and not final_labels:
            errors.append("제외가 아닌 검토결정에는 하나 이상의 최종라벨이 필요합니다.")
        if "EHSP_C08" in final_labels and len(final_labels) > 1:
            errors.append("EHSP_C08 Unknown은 다른 원인라벨과 동시에 선택할 수 없습니다.")
        if not evidence_type:
            errors.append("주요 근거 유형을 선택해 주세요.")
        if not reason:
            errors.append("판정 사유를 입력해 주세요.")
        if not reviewer_id:
            errors.append("검토자 ID를 입력해 주세요.")
        if errors:
            for error in errors:
                flash(error, "error")
            return redirect(request.form.get("return_url") or url_for("comparison_review", project_id=project_id))
        reviewed_at = now_text()
        with get_db() as db:
            db.execute(
                """INSERT INTO comparison_reviews
                   (project_id,event_id,decision,final_labels,evidence_type,reason,memo,reviewer_id,reviewed_at)
                   VALUES (?,?,?,?,?,?,?,?,?)
                   ON CONFLICT(project_id,event_id) DO UPDATE SET
                   decision=excluded.decision, final_labels=excluded.final_labels,
                   evidence_type=excluded.evidence_type, reason=excluded.reason, memo=excluded.memo,
                   reviewer_id=excluded.reviewer_id, reviewed_at=excluded.reviewed_at""",
                (project_id, event_id, decision, "|".join(sorted(final_labels)), evidence_type, reason,
                 request.form.get("memo", "").strip(), reviewer_id, reviewed_at),
            )
            completed = db.execute(
                "SELECT COUNT(*) FROM comparison_reviews WHERE project_id = ?", (project_id,)
            ).fetchone()[0]
            status = "step5_complete" if completed == len(population) else "step5_in_progress"
            db.execute("UPDATE projects SET status=?, updated_at=? WHERE id=?", (status, reviewed_at, project_id))
        flash(f"사건 ID {event_id}의 비교검토 결과를 저장했습니다.", "success")
        return redirect(request.form.get("return_url") or url_for("comparison_review", project_id=project_id))

    reviews = {row["event_id"]: dict(row) for row in load_comparison_reviews(project_id)}
    for comparison in population:
        event_id = comparison["event_id"]
        comparison["event"] = event_map.get(event_id, {})
        comparison["evidence"] = evidence_map.get(event_id, [])
        comparison["review"] = reviews.get(event_id)
    status_filter = request.args.get("status", "pending")
    if status_filter == "completed":
        filtered = [row for row in population if row["review"]]
    elif status_filter == "all":
        filtered = population
    else:
        status_filter = "pending"
        filtered = [row for row in population if not row["review"]]
    per_page = 10
    try:
        page = max(1, int(request.args.get("page", 1)))
    except ValueError:
        page = 1
    page_count = max(1, (len(filtered) + per_page - 1) // per_page)
    page = min(page, page_count)
    page_rows = filtered[(page - 1) * per_page:page * per_page]
    completed_count = len(reviews)
    summary = {
        "total": len(population), "completed": completed_count,
        "pending": len(population) - completed_count,
        "progress": round(completed_count * 100 / len(population), 1) if population else 100,
    }
    codebook = json.loads(CODEBOOK_PATH.read_text(encoding="utf-8"))
    label_categories = [
        {
            "code": label.get("label_id", ""),
            "original_name": label.get("official_name", ""),
            "korean_name": label.get("korean_name", ""),
            "category": label.get("official_definition", ""),
        }
        for label in sorted(codebook.get("labels", []), key=lambda value: value.get("official_order", 999))
    ]
    return render_template(
        "comparison_review.html", project=item, upload=upload, rows=page_rows, summary=summary,
        labels=EHSP_LABELS, label_categories=label_categories,
        status_filter=status_filter, page=page, page_count=page_count,
    )



@app.post("/projects/<project_id>/comparison-review/auto-fill-test")
def comparison_review_auto_fill_test(project_id: str):
    """미완료 비교검토를 다음 단계 기능 시험용 임시 데이터로 채운다."""
    try:
        _, _, population, _, _ = comparison_review_context(project_id)
    except LookupError as exc:
        flash(str(exc), "error")
        return redirect(url_for("project", project_id=project_id))

    timestamp = now_text()
    inserted = 0
    with get_db() as db:
        existing = {
            row["event_id"]
            for row in db.execute(
                "SELECT event_id FROM comparison_reviews WHERE project_id = ?", (project_id,)
            ).fetchall()
        }
        for comparison in population:
            event_id = comparison["event_id"]
            if event_id in existing:
                continue
            confirmed = split_codes(comparison["confirmed_labels"])
            candidate = split_codes(comparison["candidate_labels"])

            # 기능 시험용 보수적 규칙: 기존 확정라벨을 우선 유지한다.
            # 기존 라벨이 없을 때만 후보라벨을 채택하고, 둘 다 없으면 제외한다.
            if confirmed:
                decision = "keep_confirmed"
                final_labels = confirmed
                reason = "다음 단계 기능 시험을 위한 임시 자동입력: 기존 확정라벨을 유지함."
            elif candidate:
                decision = "adopt_candidate"
                final_labels = candidate
                reason = "다음 단계 기능 시험을 위한 임시 자동입력: 기존 라벨이 없어 프로그램 후보라벨을 채택함."
            else:
                decision = "exclude"
                final_labels = set()
                reason = "다음 단계 기능 시험을 위한 임시 자동입력: 확정라벨과 후보라벨이 없어 판정 불가로 처리함."

            db.execute(
                """INSERT INTO comparison_reviews
                   (project_id,event_id,decision,final_labels,evidence_type,reason,memo,reviewer_id,reviewed_at)
                   VALUES (?,?,?,?,?,?,?,?,?)""",
                (
                    project_id,
                    event_id,
                    decision,
                    "|".join(sorted(final_labels)),
                    "both",
                    reason,
                    "TEST DATA: 실제 연구 검토결과로 사용하지 말 것",
                    "TEST_AUTO",
                    timestamp,
                ),
            )
            inserted += 1

        completed = db.execute(
            "SELECT COUNT(*) FROM comparison_reviews WHERE project_id = ?", (project_id,)
        ).fetchone()[0]
        status = "step5_complete" if completed == len(population) else "step5_in_progress"
        db.execute("UPDATE projects SET status=?, updated_at=? WHERE id=?", (status, timestamp, project_id))

    flash(f"미완료 {inserted}건에 기능 시험용 임시 검토결과를 저장했습니다.", "success")
    return redirect(url_for("comparison_review", project_id=project_id, status="all"))


@app.post("/projects/<project_id>/comparison-review/delete-test-data")
def comparison_review_delete_test_data(project_id: str):
    """TEST_AUTO가 입력한 임시 비교검토 결과만 삭제한다."""
    try:
        _, _, population, _, _ = comparison_review_context(project_id)
    except LookupError as exc:
        flash(str(exc), "error")
        return redirect(url_for("project", project_id=project_id))

    timestamp = now_text()
    with get_db() as db:
        cursor = db.execute(
            "DELETE FROM comparison_reviews WHERE project_id = ? AND reviewer_id = 'TEST_AUTO'",
            (project_id,),
        )
        deleted = cursor.rowcount
        completed = db.execute(
            "SELECT COUNT(*) FROM comparison_reviews WHERE project_id = ?", (project_id,)
        ).fetchone()[0]
        status = "step5_complete" if completed == len(population) else "step5_in_progress"
        db.execute("UPDATE projects SET status=?, updated_at=? WHERE id=?", (status, timestamp, project_id))

    flash(f"기능 시험용 임시 검토결과 {deleted}건을 삭제했습니다.", "success")
    return redirect(url_for("comparison_review", project_id=project_id, status="pending"))

def comparison_review_export(project_id: str) -> tuple[sqlite3.Row, list[dict], dict[str, dict]]:
    item, _, population, event_map, evidence_map = comparison_review_context(project_id)
    reviews = {row["event_id"]: dict(row) for row in load_comparison_reviews(project_id)}
    rows = []
    for index, comparison in enumerate(population, 1):
        event_id = comparison["event_id"]
        review = reviews.get(event_id)
        event = event_map.get(event_id, {})
        final = split_codes(review["final_labels"]) if review else set()
        rows.append({
            "event_id": event_id, "review_order": index, "event_title": event.get("event_title", ""),
            "quality_level": event.get("quality_level", ""), "comparison_type": comparison["comparison_type"],
            "confirmed_labels": comparison["confirmed_labels"], "candidate_labels": comparison["candidate_labels"],
            "candidate_additions": comparison["candidate_additions"], "candidate_omissions": comparison["candidate_omissions"],
            "evidence_summary": " | ".join(f'{x["label_id"]}:{x["rule_id"]}:{x["evidence_text"]}' for x in evidence_map.get(event_id, [])),
            "root_cause_analysis": event.get("root_cause_analysis", ""),
            "review_status": "completed" if review else "pending",
            "decision": review["decision"] if review else "pending",
            "final_labels": review["final_labels"] if review else "",
            **{label["code"]: (int(label["code"] in final) if review and review["decision"] != "exclude" else ("not_assessable" if review else "pending")) for label in EHSP_LABELS},
            "evidence_type": review["evidence_type"] if review else "", "reason": review["reason"] if review else "",
            "memo": review["memo"] if review else "", "reviewer_id": review["reviewer_id"] if review else "",
            "reviewed_at": review["reviewed_at"] if review else "",
        })
    return item, rows, reviews


@app.get("/projects/<project_id>/comparison-review.json")
def download_comparison_review_json(project_id: str):
    try:
        item, rows, reviews = comparison_review_export(project_id)
    except LookupError as exc:
        return str(exc), 404
    completed = len(reviews)
    payload = {
        "program": "HIAD 2.2 Web Dataset Builder", "step": 5,
        "status": "completed" if completed == len(rows) else "in_progress",
        "review_population": len(rows), "completed_reviews": completed,
        "pending_reviews": len(rows) - completed,
        "decision_distribution": pd.Series([r["decision"] for r in rows]).value_counts().to_dict(),
        "exported_at": now_text(),
    }
    output = DATA_DIR / f"{project_id}_step5_review_summary.json"
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return send_file(output, as_attachment=True, download_name=f"{item['name']}_5단계_비교검토_요약.json")


@app.get("/projects/<project_id>/comparison-review.csv")
def download_comparison_review_csv(project_id: str):
    try:
        item, rows, _ = comparison_review_export(project_id)
    except LookupError as exc:
        return str(exc), 404
    output = DATA_DIR / f"{project_id}_step5_review_results.csv"
    pd.DataFrame(rows).to_csv(output, index=False, encoding="utf-8-sig")
    return send_file(output, as_attachment=True, download_name=f"{item['name']}_5단계_비교검토_결과.csv")


def send_candidate_csv(project_id: str, field: str, suffix: str, download_label: str):
    item, saved = load_project(project_id), load_candidates(project_id)
    if not item or not saved:
        return "후보라벨 생성 결과가 없습니다.", 404
    output = DATA_DIR / f"{project_id}_step4_{suffix}.csv"
    pd.DataFrame(json.loads(saved[field])).to_csv(output, index=False, encoding="utf-8-sig")
    return send_file(output, as_attachment=True, download_name=f"{item['name']}_4단계_{download_label}.csv")


@app.get("/projects/<project_id>/candidates.json")
def download_candidates_json(project_id: str):
    item, saved = load_project(project_id), load_candidates(project_id)
    if not item or not saved:
        return "후보라벨 생성 결과가 없습니다.", 404
    output = DATA_DIR / f"{project_id}_step4_candidates.json"
    output.write_text(saved["result_json"], encoding="utf-8")
    return send_file(output, as_attachment=True, download_name=f"{item['name']}_4단계_후보라벨생성_요약.json")


@app.get("/projects/<project_id>/candidate-labels.csv")
def download_candidate_labels(project_id: str):
    return send_candidate_csv(project_id, "candidates_json", "candidate_labels", "후보라벨")


@app.get("/projects/<project_id>/candidate-evidence.csv")
def download_candidate_evidence(project_id: str):
    return send_candidate_csv(project_id, "evidence_json", "candidate_evidence", "후보근거")


@app.get("/projects/<project_id>/label-comparison.csv")
def download_label_comparison(project_id: str):
    return send_candidate_csv(project_id, "comparisons_json", "label_comparison", "원본후보비교")


@app.get("/projects/<project_id>/standardization.json")
def download_standardization_json(project_id: str):
    item, saved = load_project(project_id), load_standardization(project_id)
    if not item or not saved:
        return "원본라벨 표준화 결과가 없습니다.", 404
    output = DATA_DIR / f"{project_id}_step3_standardization.json"
    output.write_text(saved["result_json"], encoding="utf-8")
    return send_file(output, as_attachment=True, download_name=f"{item['name']}_3단계_원본라벨표준화_요약.json")


@app.get("/projects/<project_id>/standardization.csv")
def download_standardization_csv(project_id: str):
    item, saved = load_project(project_id), load_standardization(project_id)
    if not item or not saved:
        return "원본라벨 표준화 결과가 없습니다.", 404
    output = DATA_DIR / f"{project_id}_step3_standardization.csv"
    pd.DataFrame(json.loads(saved["rows_json"])).to_csv(output, index=False, encoding="utf-8-sig")
    return send_file(output, as_attachment=True, download_name=f"{item['name']}_3단계_원본라벨표준화_사건별.csv")


@app.get("/projects/<project_id>/standardization-review.csv")
def download_standardization_review_csv(project_id: str):
    item, saved = load_project(project_id), load_standardization(project_id)
    if not item or not saved:
        return "원본라벨 표준화 결과가 없습니다.", 404
    rows = [r for r in json.loads(saved["rows_json"]) if r["mapping_status"] != "mapped"]
    columns = ["event_id", "original_summary_root_causes", "mapping_status", "unmapped_values"]
    output = DATA_DIR / f"{project_id}_step3_review.csv"
    pd.DataFrame(rows, columns=columns).to_csv(output, index=False, encoding="utf-8-sig")
    return send_file(output, as_attachment=True, download_name=f"{item['name']}_3단계_미매핑검토대상.csv")


@app.get("/projects/<project_id>/validation.json")
def download_validation_json(project_id: str):
    item, saved = load_project(project_id), load_validation(project_id)
    if not item or not saved:
        return "입력자료 검증 결과가 없습니다.", 404
    output = DATA_DIR / f"{project_id}_step2_validation.json"
    output.write_text(saved["validation_json"], encoding="utf-8")
    return send_file(output, as_attachment=True, download_name=f"{item['name']}_2단계_입력자료검증.json")


@app.get("/projects/<project_id>/validation.csv")
def download_validation_csv(project_id: str):
    item, saved = load_project(project_id), load_validation(project_id)
    if not item or not saved:
        return "입력자료 검증 결과가 없습니다.", 404
    result = json.loads(saved["validation_json"])
    rows = []
    for check in result["checks"]:
        rows.append({
            "severity": check["severity"], "status": "pass" if check["passed"] else "issue",
            "issue_code": check["code"], "title": check["title"], "count": check["count"],
            "detail": check["detail"], "sample_ids": " | ".join(check["sample_ids"]),
        })
    output = DATA_DIR / f"{project_id}_step2_validation.csv"
    pd.DataFrame(rows).to_csv(output, index=False, encoding="utf-8-sig")
    return send_file(output, as_attachment=True, download_name=f"{item['name']}_2단계_입력자료검증.csv")


@app.get("/projects/<project_id>/settings.json")
def download_settings(project_id: str):
    item, upload = load_project(project_id), load_upload(project_id)
    with get_db() as db:
        selection = db.execute("SELECT * FROM selections WHERE project_id = ?", (project_id,)).fetchone()
    if not item or not upload or not selection:
        return "확정된 설정이 없습니다.", 404
    payload = {
        "program": "HIAD 2.2 Web Dataset Builder",
        "step": 1,
        "project": {"id": item["id"], "name": item["name"]},
        "input": {"filename": upload["original_name"], "sha256": upload["sha256"], "size_bytes": upload["size_bytes"]},
        "paper_baseline": json.loads(selection["baseline_json"]),
        "execution_selection": json.loads(selection["selection_json"]),
        "uses_paper_baseline": bool(selection["is_paper_baseline"]),
        "confirmed_at": selection["confirmed_at"],
    }
    output = DATA_DIR / f"{project_id}_step1_settings.json"
    output.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return send_file(output, as_attachment=True, download_name=f"{item['name']}_1단계_설정.json")


@app.get("/api/sheets/<project_id>/<path:sheet_name>")
def sheet_columns(project_id: str, sheet_name: str):
    upload = load_upload(project_id)
    if upload is None:
        return {"error": "not found"}, 404
    profile = json.loads(upload["workbook_json"])
    sheet = next((item for item in profile["sheets"] if item["name"] == sheet_name), None)
    return (sheet or {"error": "sheet not found"}, 200 if sheet else 404)


@app.errorhandler(413)
def too_large(_error):
    flash("파일 크기는 100MB 이하만 허용됩니다.", "error")
    return redirect(url_for("index"))


def load_candidate_omission_reviews(project_id: str) -> list[sqlite3.Row]:
    with get_db() as db:
        return db.execute("SELECT * FROM candidate_omission_reviews WHERE project_id=? ORDER BY reviewed_at,event_id", (project_id,)).fetchall()


def candidate_omission_context(project_id: str):
    item, upload, selection_row, saved = load_project(project_id), load_upload(project_id), load_selection(project_id), load_candidates(project_id)
    if not item or not upload or not selection_row or not saved:
        raise LookupError("먼저 4단계 후보라벨 생성을 완료해 주세요.")
    population = [r for r in json.loads(saved["comparisons_json"]) if r.get("comparison_type") == "no_candidate"]
    population.sort(key=lambda r: natural_event_key(r["event_id"]))
    selection=json.loads(selection_row["selection_json"]); c=selection["columns"]
    frame=pd.read_excel(current_upload_path(upload),sheet_name=selection["events_sheet"],usecols=[c["event_id"],c["quality_level"],c["event_title"],c["event_description"],c["root_cause_analysis"]],engine="openpyxl")
    event_map={}
    for _,src in frame.iterrows():
        eid=str(src[c["event_id"]]).strip()
        event_map[eid]={"event_id":eid,"quality_level":"" if pd.isna(src[c["quality_level"]]) else str(src[c["quality_level"]]),"event_title":normalize_text(src[c["event_title"]]),"event_description":normalize_text(src[c["event_description"]]),"root_cause_analysis":normalize_text(src[c["root_cause_analysis"]])}
    return item,upload,population,event_map


def label_category_rows():
    cb=json.loads(CODEBOOK_PATH.read_text(encoding="utf-8"))
    return [{"code":x.get("label_id",""),"original_name":x.get("official_name",""),"korean_name":x.get("korean_name",""),"category":x.get("official_definition","")} for x in sorted(cb.get("labels",[]),key=lambda v:v.get("official_order",999))]


@app.route("/projects/<project_id>/candidate-omission-review",methods=["GET","POST"])
def candidate_omission_review(project_id: str):
    try: item,upload,population,event_map=candidate_omission_context(project_id)
    except LookupError as exc:
        flash(str(exc),"error"); return redirect(url_for("project",project_id=project_id))
    by_id={r["event_id"]:r for r in population}
    if request.method=="POST":
        eid=request.form.get("event_id","").strip(); row=by_id.get(eid); decision=request.form.get("decision","").strip()
        if not row or decision not in {"keep_confirmed","custom","exclude"}:
            flash("검토대상과 검토결정을 확인해 주세요.","error"); return redirect(request.form.get("return_url") or url_for("candidate_omission_review",project_id=project_id))
        labels=split_codes(row.get("confirmed_labels","")) if decision=="keep_confirmed" else (set(request.form.getlist("final_labels")) if decision=="custom" else set())
        evidence=request.form.get("evidence_type","").strip(); reason=request.form.get("reason","").strip(); reviewer=request.form.get("reviewer_id","").strip(); errors=[]
        if decision!="exclude" and not labels: errors.append("하나 이상의 최종라벨이 필요합니다.")
        if "EHSP_C08" in labels and len(labels)>1: errors.append("EHSP_C08은 다른 라벨과 함께 선택할 수 없습니다.")
        if not evidence: errors.append("주요 근거 유형을 선택해 주세요.")
        if not reason: errors.append("판정 사유를 입력해 주세요.")
        if not reviewer: errors.append("검토자 ID를 입력해 주세요.")
        if errors:
            for e in errors: flash(e,"error")
            return redirect(request.form.get("return_url") or url_for("candidate_omission_review",project_id=project_id))
        ts=now_text()
        sql="INSERT INTO candidate_omission_reviews(project_id,event_id,decision,final_labels,evidence_type,reason,memo,reviewer_id,reviewed_at) VALUES (?,?,?,?,?,?,?,?,?) ON CONFLICT(project_id,event_id) DO UPDATE SET decision=excluded.decision,final_labels=excluded.final_labels,evidence_type=excluded.evidence_type,reason=excluded.reason,memo=excluded.memo,reviewer_id=excluded.reviewer_id,reviewed_at=excluded.reviewed_at"
        with get_db() as db:
            db.execute(sql,(project_id,eid,decision,"|".join(sorted(labels)),evidence,reason,request.form.get("memo","").strip(),reviewer,ts))
            n=db.execute("SELECT COUNT(*) FROM candidate_omission_reviews WHERE project_id=?",(project_id,)).fetchone()[0]
            db.execute("UPDATE projects SET status=?,updated_at=? WHERE id=?",("step6_complete" if n==len(population) else "step6_in_progress",ts,project_id))
        flash(f"사건 ID {eid}의 검토결과를 저장했습니다.","success"); return redirect(request.form.get("return_url") or url_for("candidate_omission_review",project_id=project_id))
    reviews={r["event_id"]:dict(r) for r in load_candidate_omission_reviews(project_id)}
    for r in population: r["event"]=event_map.get(r["event_id"],{}); r["review"]=reviews.get(r["event_id"])
    sf=request.args.get("status","pending")
    if sf=="completed": filtered=[r for r in population if r["review"]]
    elif sf=="all": filtered=population
    else: sf="pending"; filtered=[r for r in population if not r["review"]]
    try: page=max(1,int(request.args.get("page",1)))
    except ValueError: page=1
    per=10; pc=max(1,(len(filtered)+per-1)//per); page=min(page,pc); rows=filtered[(page-1)*per:page*per]; done=len(reviews)
    summary={"total":len(population),"completed":done,"pending":len(population)-done,"progress":round(done*100/len(population),1) if population else 100}
    return render_template("candidate_omission_review.html",project=item,upload=upload,rows=rows,summary=summary,labels=EHSP_LABELS,label_categories=label_category_rows(),status_filter=sf,page=page,page_count=pc)


@app.post("/projects/<project_id>/candidate-omission-review/auto-fill-test")
def candidate_omission_auto_fill_test(project_id: str):
    item,upload,population,event_map=candidate_omission_context(project_id); ts=now_text(); added=0
    with get_db() as db:
        existing={x[0] for x in db.execute("SELECT event_id FROM candidate_omission_reviews WHERE project_id=?",(project_id,))}
        for r in population:
            if r["event_id"] in existing: continue
            db.execute("INSERT INTO candidate_omission_reviews VALUES (?,?,?,?,?,?,?,?,?)",(project_id,r["event_id"],"keep_confirmed",r.get("confirmed_labels",""),"insufficient","다음 단계 기능 시험용 임시 데이터","실제 연구결과로 사용하지 않음","TEST_AUTO",ts)); added+=1
        db.execute("UPDATE projects SET status='step6_complete',updated_at=? WHERE id=?",(ts,project_id))
    flash(f"미완료 {added}건을 임시 처리했습니다.","success"); return redirect(url_for("candidate_omission_review",project_id=project_id,status="all"))


@app.post("/projects/<project_id>/candidate-omission-review/delete-test-data")
def candidate_omission_delete_test_data(project_id: str):
    _,_,population,_=candidate_omission_context(project_id); ts=now_text()
    with get_db() as db:
        n=db.execute("DELETE FROM candidate_omission_reviews WHERE project_id=? AND reviewer_id='TEST_AUTO'",(project_id,)).rowcount
        done=db.execute("SELECT COUNT(*) FROM candidate_omission_reviews WHERE project_id=?",(project_id,)).fetchone()[0]
        db.execute("UPDATE projects SET status=?,updated_at=? WHERE id=?",("step6_complete" if done==len(population) else "step6_in_progress",ts,project_id))
    flash(f"임시 데이터 {n}건을 삭제했습니다.","success"); return redirect(url_for("candidate_omission_review",project_id=project_id))


def final_dataset_context(project_id: str):
    project = load_project(project_id)
    upload = load_upload(project_id)
    selection = load_selection(project_id)
    standardization = load_standardization(project_id)
    candidates = load_candidates(project_id)
    if not all([project, upload, selection, standardization, candidates]):
        raise LookupError("1~4단계 결과가 모두 필요합니다.")
    comparisons = json.loads(candidates["comparisons_json"])
    diff_total = len(review_population(comparisons))
    omission_total = sum(r.get("comparison_type") == "no_candidate" for r in comparisons)
    diff_reviews = {r["event_id"]: dict(r) for r in load_comparison_reviews(project_id)}
    omission_reviews = {r["event_id"]: dict(r) for r in load_candidate_omission_reviews(project_id)}
    if len(diff_reviews) != diff_total:
        raise LookupError(f"5단계 비교검토가 완료되지 않았습니다. {len(diff_reviews)}/{diff_total}건")
    if len(omission_reviews) != omission_total:
        raise LookupError(f"6단계 후보 없음 검토가 완료되지 않았습니다. {len(omission_reviews)}/{omission_total}건")
    return project, upload, selection, standardization, candidates, comparisons, diff_reviews, omission_reviews


def build_final_rows(project_id: str):
    project, upload, selection, standardization, candidates, comparisons, diff_reviews, omission_reviews = final_dataset_context(project_id)
    standard_rows = {str(r["event_id"]): r for r in json.loads(standardization["rows_json"])}
    candidate_rows = {str(r["event_id"]): r for r in json.loads(candidates["candidates_json"])}
    comparison_map = {str(r["event_id"]): r for r in comparisons}
    evidence_rows = json.loads(candidates["evidence_json"])
    evidence_map = {}
    for e in evidence_rows:
        evidence_map.setdefault(str(e["event_id"]), []).append(e)

    event_map = {}
    try:
        path = current_upload_path(upload)
        sel = json.loads(selection["selection_json"])
        frame = pd.read_excel(path, sheet_name=sel["events_sheet"], engine="openpyxl")
        cols = sel["columns"]
        for _, row in frame.iterrows():
            eid = str(row.get(cols["event_id"], "")).strip()
            if eid.endswith('.0') and eid[:-2].isdigit(): eid = eid[:-2]
            def clean(key):
                value = row.get(cols[key], "")
                return "" if pd.isna(value) else str(value).strip()
            event_map[eid] = {
                "quality_level": clean("quality_level"), "event_title": clean("event_title"),
                "event_description": clean("event_description"),
                "summary_root_causes": clean("summary_root_causes"),
                "root_cause_analysis": clean("root_cause_analysis"),
            }
    except Exception:
        event_map = {}

    final_rows = []
    for eid in sorted(standard_rows, key=natural_event_key):
        std = standard_rows[eid]
        cand = candidate_rows.get(eid, {})
        comp = comparison_map.get(eid, {})
        ctype = comp.get("comparison_type", "")
        review = None
        review_stage = "자동 확정"
        if ctype == "no_candidate":
            review = omission_reviews.get(eid); review_stage = "6단계"
        elif ctype in {"mixed_difference", "candidate_addition", "candidate_omission"}:
            review = diff_reviews.get(eid); review_stage = "5단계"
        if review:
            final_labels = review.get("final_labels", "")
            decision = review.get("decision", "")
            excluded = 1 if decision == "exclude" else 0
        else:
            final_labels = comp.get("confirmed_labels", std.get("standard_codes", ""))
            decision = "exact_match_auto" if ctype == "exact_match" else "confirmed_auto"
            excluded = 0
        codes = sorted(split_codes(final_labels))
        ev = event_map.get(eid, {})
        row = {
            "event_id": eid,
            "quality_level": ev.get("quality_level", ""),
            "event_title": ev.get("event_title", ""),
            "event_description": ev.get("event_description", ""),
            "summary_root_causes": ev.get("summary_root_causes", std.get("original_summary_root_causes", "")),
            "root_cause_analysis": ev.get("root_cause_analysis", ""),
            "original_labels": std.get("standard_codes", ""),
            "candidate_labels": cand.get("candidate_labels", ""),
            "comparison_type": ctype,
            "final_labels": "|".join(codes),
            "final_label_count": len(codes),
            "excluded": excluded,
            "finalization_stage": review_stage,
            "review_decision": decision,
            "evidence_type": review.get("evidence_type", "") if review else "",
            "review_reason": review.get("reason", "") if review else "",
            "review_memo": review.get("memo", "") if review else "",
            "reviewer_id": review.get("reviewer_id", "") if review else "",
            "reviewed_at": review.get("reviewed_at", "") if review else "",
            "test_data_flag": 1 if review and review.get("reviewer_id") == "TEST_AUTO" else 0,
            "candidate_evidence_count": len(evidence_map.get(eid, [])),
        }
        for label in EHSP_LABELS:
            row[label["code"]] = 1 if label["code"] in codes and not excluded else 0
        final_rows.append(row)
    included = [r for r in final_rows if not r["excluded"]]
    distribution = {lab["code"]: sum(r[lab["code"]] for r in included) for lab in EHSP_LABELS}
    summary = {
        "program": "HIAD 2.2 Web Dataset Builder", "version": "0.7.0", "step": 7,
        "status": "completed", "project_id": project_id, "project_name": project["name"],
        "source_file": upload["original_name"], "source_sha256": upload["sha256"],
        "total_events": len(final_rows), "included_events": len(included),
        "excluded_events": len(final_rows)-len(included),
        "multi_label_events": sum(r["final_label_count"] > 1 for r in included),
        "single_label_events": sum(r["final_label_count"] == 1 for r in included),
        "unlabeled_included_events": sum(r["final_label_count"] == 0 for r in included),
        "total_positive_labels": sum(r["final_label_count"] for r in included),
        "label_distribution": distribution,
        "step5_review_records": len(diff_reviews), "step6_review_records": len(omission_reviews),
        "temporary_test_records": sum(r["test_data_flag"] for r in final_rows),
        "generated_at": now_text(),
        "warning": "TEST_AUTO 자료가 포함된 경우 실제 연구용 최종 데이터셋으로 사용하기 전에 수동 검토로 교체해야 합니다."
    }
    return project, final_rows, included, evidence_rows, summary


@app.get("/projects/<project_id>/final-dataset")
def final_dataset(project_id: str):
    try:
        project, rows, included, evidence, summary = build_final_rows(project_id)
    except LookupError as exc:
        flash(str(exc), "error")
        return redirect(url_for("project", project_id=project_id))
    with get_db() as db:
        build = db.execute("SELECT * FROM final_builds WHERE project_id=?", (project_id,)).fetchone()
    return render_template("final_dataset.html", project=project, summary=summary, build=build, labels=EHSP_LABELS)


@app.post("/projects/<project_id>/final-dataset/generate")
def generate_final_dataset(project_id: str):
    try:
        project, rows, included, evidence, summary = build_final_rows(project_id)
    except LookupError as exc:
        flash(str(exc), "error")
        return redirect(url_for("project", project_id=project_id))
    ts = summary["generated_at"]
    with get_db() as db:
        db.execute("INSERT INTO final_builds(project_id,summary_json,generated_at) VALUES(?,?,?) ON CONFLICT(project_id) DO UPDATE SET summary_json=excluded.summary_json,generated_at=excluded.generated_at", (project_id, json.dumps(summary, ensure_ascii=False), ts))
        db.execute("UPDATE projects SET status='step7_complete',updated_at=? WHERE id=?", (ts, project_id))
    flash("7단계 최종 학습 데이터셋을 생성했습니다.", "success")
    return redirect(url_for("final_dataset", project_id=project_id))


def dataframe_csv_bytes(rows):
    output = io.BytesIO()
    pd.DataFrame(rows).to_csv(output, index=False, encoding="utf-8-sig")
    output.seek(0)
    return output


@app.get("/projects/<project_id>/final-dataset/download.csv")
def download_final_dataset_csv(project_id: str):
    project, rows, included, evidence, summary = build_final_rows(project_id)
    return send_file(dataframe_csv_bytes(included), as_attachment=True, download_name=f"{project['name']}_7단계_최종_다중라벨_학습데이터셋.csv", mimetype="text/csv")


@app.get("/projects/<project_id>/final-dataset/audit.csv")
def download_final_audit_csv(project_id: str):
    project, rows, included, evidence, summary = build_final_rows(project_id)
    return send_file(dataframe_csv_bytes(rows), as_attachment=True, download_name=f"{project['name']}_7단계_최종라벨_추적표.csv", mimetype="text/csv")


@app.get("/projects/<project_id>/final-dataset/evidence.csv")
def download_final_evidence_csv(project_id: str):
    project, rows, included, evidence, summary = build_final_rows(project_id)
    return send_file(dataframe_csv_bytes(evidence), as_attachment=True, download_name=f"{project['name']}_7단계_후보근거_통합.csv", mimetype="text/csv")


@app.get("/projects/<project_id>/final-dataset/report.json")
def download_final_report_json(project_id: str):
    project, rows, included, evidence, summary = build_final_rows(project_id)
    output = io.BytesIO(json.dumps(summary, ensure_ascii=False, indent=2).encode("utf-8")); output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"{project['name']}_7단계_구축보고서.json", mimetype="application/json")


@app.get("/projects/<project_id>/final-dataset/package.zip")
def download_final_package(project_id: str):
    project, rows, included, evidence, summary = build_final_rows(project_id)
    output = io.BytesIO()
    with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("01_최종_다중라벨_학습데이터셋.csv", pd.DataFrame(included).to_csv(index=False, encoding="utf-8-sig"))
        z.writestr("02_최종라벨_추적표.csv", pd.DataFrame(rows).to_csv(index=False, encoding="utf-8-sig"))
        z.writestr("03_후보근거_통합.csv", pd.DataFrame(evidence).to_csv(index=False, encoding="utf-8-sig"))
        z.writestr("04_최종_구축보고서.json", json.dumps(summary, ensure_ascii=False, indent=2))
        z.writestr("05_EHSP_라벨정의.json", json.dumps(EHSP_LABELS, ensure_ascii=False, indent=2))
    output.seek(0)
    return send_file(output, as_attachment=True, download_name=f"{project['name']}_7단계_최종산출물.zip", mimetype="application/zip")


init_db()


def open_browser() -> None:
    """서버 시작 직후 기본 브라우저에서 로컬 화면을 연다."""
    webbrowser.open_new("http://127.0.0.1:5000")


if __name__ == "__main__":
    print("=" * 60)
    print("HIAD 2.2 웹 데이터셋 구축 프로그램 v0.7.0")
    print("서버 주소: http://127.0.0.1:5000")
    print("종료하려면 IDLE Shell에서 Ctrl+C를 누르거나 Shell을 닫으세요.")
    print("=" * 60)
    threading.Timer(1.2, open_browser).start()
    app.run(host="127.0.0.1", port=5000, debug=False)
