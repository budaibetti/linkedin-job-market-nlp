"""Step 6 — consolidate all completed months into ONE Power BI-ready CSV.

Place: <project>/scripts/06_final_merge.py
Run (from project root): python scripts/06_final_merge.py
Run (from scripts dir): python 06_final_merge.py

Inputs, for EACH completed month:
  data/04_data_with_id/<Month>/jobs_YYYY-MM.csv
  data/05_hard_skills/<Month>/job_skills_YYYY-MM.csv
  data/06_ai_mentions/<Month>/job_ai_counts_YYYY-MM.csv
  data/04_data_with_id/_registry/job_ids.sqlite

Outputs (rebuilt each run; never append blindly):
  data/final_data/merged_postings_final.csv
  data/final_data/merged_postings_manifest.json

- Exactly one master row per global job_month_id, including zero-skill jobs.
- job_id is permitted to repeat across different months for the same job.
- Never creates, changes or reassigns identifiers. Reads SQLite registry read-only.
- Any missing monthly enrichment, orphan skills, colliding IDs, or conflicting
  job/skill identities causes the operation to fail BEFORE replacing output.
- Includes original scraped skills and complete hard-skill details as JSON.
- AI mention counts are copied without treating them as a boolean flag.

Python standard library only (no pip packages required).
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import sqlite3
import sys
import tempfile
from collections import defaultdict
from pathlib import Path
from urllib.parse import quote

VERSION = "6.0.0"
MONTHS = (
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December",
)
JOB_REQUIRED = {
    "job_id", "job_month_id", "snapshot_month", "job_title", "company_name",
    "job_description", "location", "market_country", "search_city",
    "skills_json", "identity_status",
}
SKILL_REQUIRED = {
    "job_id", "job_month_id", "snapshot_month", "market_country", "search_city",
    "skill_id", "skill_name", "category", "role", "match_source",
}
AI_COLUMNS = [
    "ai_mentions_description", "ai_mentions_scraped_skills", "ai_mentions_total",
    "ai_explicit_mentions", "ai_method_mentions", "ai_tool_mentions", "has_ai_mention",
]
AI_REQUIRED = {
    "job_id", "job_month_id", "snapshot_month", "market_country", "search_city",
    *AI_COLUMNS,
}
ENRICH_COLUMNS = [
    "hard_skill_id", "hard_skill", "hard_skill_count",
    "hard_skill_ids_json", "hard_skills_json", "hard_skill_sources_json",
    "hard_skill_details_json", *AI_COLUMNS,
]
BASE_COLUMNS_FIRST = [
    "job_id", "job_month_id", "snapshot_month", "market_country", "search_city",
    "job_title", "company_name", "job_description", "location", "posted_date",
    "employment_type", "industry", "seniority_level", "skills", "skills_json",
    "identity_status", "identity_missing_fields", "same_month_source_count",
    "source_file", "source_sheet", "source_row", "source_records_json",
]
FILENAME = re.compile(r"^jobs_(\d{4})-(\d{2})\.csv$")


def fail(message: str) -> None:
    raise ValueError(message)


def positive_int(value: str, label: str) -> int:
    try:
        number = int(value)
    except (ValueError, TypeError):
        fail(f"{label}: invalid integer {value!r}")
    if number <= 0:
        fail(f"{label}: expected positive integer, got {number}")
    return number


def nonnegative_int(value: str, label: str) -> int:
    try:
        number = int(value)
    except (ValueError, TypeError):
        fail(f"{label}: invalid integer {value!r}")
    if number < 0:
        fail(f"{label}: negative count {number}")
    return number


def read_csv(path: Path, required: set[str]) -> tuple[list[str], list[dict[str, str]]]:
    if not path.is_file():
        fail(f"Missing required input: {path}")
    with path.open("r", encoding="utf-8-sig", newline="") as handle:
        reader = csv.DictReader(handle)
        columns = reader.fieldnames or []
        if len(columns) != len(set(columns)):
            fail(f"Duplicate column headers in {path}")
        missing = sorted(required - set(columns))
        if missing:
            fail(f"{path}: missing column(s): {', '.join(missing)}")
        rows = []
        for line_num, row in enumerate(reader, 2):
            if None in row or any(value is None for value in row.values()):
                fail(f"{path}: row {line_num} has an incorrect CSV column count")
            rows.append(row)
    return columns, rows


def find_months(root: Path) -> list[tuple[str, str, Path]]:
    parent = root / "data" / "04_data_with_id"
    found = []
    if not parent.is_dir():
        fail(f"Missing Step 3 folder: {parent}")
    for path in sorted(parent.glob("*/jobs_*.csv")):
        match = FILENAME.fullmatch(path.name)
        if not match:
            fail(f"Unexpected Step 3 job filename: {path}")
        year, month_number = int(match[1]), int(match[2])
        if not 2000 <= year <= 2100 or not 1 <= month_number <= 12:
            fail(f"Invalid month encoded in filename: {path}")
        label = MONTHS[month_number - 1]
        if path.parent.name != label:
            fail(f"Month folder for {path.name} must be '{label}', not '{path.parent.name}'")
        snapshot = f"{year:04d}-{month_number:02d}"
        found.append((snapshot, label, path))
    if not found:
        fail(f"No monthly jobs_YYYY-MM.csv found beneath {parent}")
    found.sort(key=lambda item: item[0])
    snapshots = [item[0] for item in found]
    if len(snapshots) != len(set(snapshots)):
        fail("Duplicate snapshot month from multiple Step 3 sources")
    return found


def load_registry(registry: Path) -> tuple[dict[int, tuple[int, str]], dict[int, str]]:
    if not registry.is_file():
        fail(
            f"Persistent ID registry missing: {registry}. Restore the original "
            "job_ids.sqlite from backup; do NOT generate a new registry."
        )
    uri = "file:" + quote(str(registry.resolve()), safe="/:") + "?mode=ro"
    try:
        conn = sqlite3.connect(uri, uri=True)
        try:
            pairs = {
                int(mid): (int(jid), month)
                for mid, jid, month in conn.execute(
                    "SELECT job_month_id, job_id, snapshot_month FROM job_months"
                )
            }
            fingerprints = {
                int(jid): fingerprint
                for jid, fingerprint in conn.execute(
                    "SELECT job_id, fingerprint FROM jobs"
                )
            }
        finally:
            conn.close()
    except sqlite3.Error as exc:
        fail(f"Cannot read persistent registry {registry}: {exc}")
    return pairs, fingerprints


def computed_fingerprint(row: dict[str, str], snapshot: str) -> str | None:
    """Reproduce Step 3 hashes when all source coordinates are available.

    For sparse fallback rows with incomplete source coordinates, Step 3 may
    include the original merged-file row number, which is not in the output.
    Those rows are still checked against the registry's job-month mapping.
    """
    fields = ("job_title", "company_name", "job_description", "location")
    values = [row[field] for field in fields]
    status = row["identity_status"]
    if status == "complete":
        payload = values
    elif status == "partial":
        market = [row["market_country"], row["search_city"]] if not row["location"].strip() else []
        payload = ["partial-v1", values, market]
    elif status == "source_fallback":
        coords = [row.get(field, "").strip() for field in ("source_file", "source_sheet", "source_row")]
        if not all(coords):
            return None
        payload = ["source-fallback-v1", snapshot, coords, values]
    else:
        fail(f"Unexpected identity_status: {status!r}")
    raw = json.dumps(payload, ensure_ascii=False, separators=(",", ":")).encode("utf-8")
    return hashlib.sha256(raw).hexdigest()


def shared_fields_check(
    enriched: dict[str, str], job: dict[str, str], path: Path, job_month_id: int
) -> None:
    for name in ("job_id", "snapshot_month", "market_country", "search_city"):
        if enriched[name] != job[name]:
            fail(f"{path}: job_month_id={job_month_id} has mismatching {name}")


def assemble(root: Path) -> tuple[list[str], list[dict[str, object]], dict[str, object]]:
    csv.field_size_limit(min(sys.maxsize, 2**31 - 1))
    months = find_months(root)
    registry_path = root / "data" / "04_data_with_id" / "_registry" / "job_ids.sqlite"
    id_pairs, registry_fingerprints = load_registry(registry_path)

    all_jobs: list[dict[str, object]] = []
    all_job_fields: set[str] = set()
    seen_month_ids: set[int] = set()
    seen_job_month: set[tuple[int, str]] = set()
    seen_job_id_fingerprints: dict[int, str] = {}
    skill_dictionary: dict[int, tuple[str, str, str]] = {}
    counts_by_month = []
    input_files = []
    uncheckable_fallback_fingerprints = 0

    for snapshot, label, jobs_path in months:
        skills_path = root / "data" / "05_hard_skills" / label / f"job_skills_{snapshot}.csv"
        ai_path = root / "data" / "06_ai_mentions" / label / f"job_ai_counts_{snapshot}.csv"
        job_columns, jobs = read_csv(jobs_path, JOB_REQUIRED)
        _, skill_rows = read_csv(skills_path, SKILL_REQUIRED)
        _, ai_rows = read_csv(ai_path, AI_REQUIRED)
        if not jobs:
            fail(f"Refusing to publish empty month: {snapshot}")
        overlap = sorted(set(job_columns) & set(ENRICH_COLUMNS))
        if overlap:
            fail(f"Step 3 columns conflict with Step 6 enrichment columns: {overlap}")
        all_job_fields.update(job_columns)
        month_jobs: dict[int, dict[str, str]] = {}

        for row in jobs:
            context = f"{jobs_path}: job_month_id={row['job_month_id']}"
            month_id = positive_int(row["job_month_id"], context)
            job_id = positive_int(row["job_id"], context)
            if row["snapshot_month"] != snapshot:
                fail(f"{context}: snapshot_month is not {snapshot}")
            if month_id in seen_month_ids:
                fail(f"COLLISION: job_month_id {month_id} occurs in multiple months/jobs")
            if (job_id, snapshot) in seen_job_month:
                fail(f"Duplicate (job_id, snapshot_month): {job_id}, {snapshot}")
            expected = id_pairs.get(month_id)
            if expected != (job_id, snapshot):
                fail(
                    f"REGISTRY MISMATCH: job_month_id={month_id} in {snapshot} "
                    f"belongs to {expected!r} in SQLite, but CSV says {(job_id, snapshot)!r}. "
                    "Restore original registry or correct Step 3 outputs."
                )
            if job_id not in registry_fingerprints:
                fail(f"REGISTRY MISMATCH: job_id {job_id} is absent from SQLite")
            fingerprint = computed_fingerprint(row, snapshot)
            if fingerprint is not None:
                if registry_fingerprints[job_id] != fingerprint:
                    fail(
                        f"REGISTRY FINGERPRINT MISMATCH for job_id={job_id} in {snapshot}. "
                        "Possibly a replaced registry or edited Step 3 job identity."
                    )
                previous = seen_job_id_fingerprints.setdefault(job_id, fingerprint)
                if previous != fingerprint:
                    fail(f"COLLISION: job_id {job_id} has different identities across months")
            else:
                uncheckable_fallback_fingerprints += 1
            try:
                raw_skills = json.loads(row["skills_json"])
                if not isinstance(raw_skills, list) or not all(isinstance(v, str) for v in raw_skills):
                    raise ValueError("expected array of strings")
            except (ValueError, TypeError) as exc:
                fail(f"{context}: invalid skills_json: {exc}")
            seen_month_ids.add(month_id)
            seen_job_month.add((job_id, snapshot))
            month_jobs[month_id] = row

        skill_map: dict[int, dict[int, dict[str, str]]] = defaultdict(dict)
        for row in skill_rows:
            context = f"{skills_path}: job_month_id={row['job_month_id']}"
            month_id = positive_int(row["job_month_id"], context)
            job = month_jobs.get(month_id)
            if job is None:
                fail(f"ORPHAN SKILL: {context} is not present in Step 3 jobs for {snapshot}")
            shared_fields_check(row, job, skills_path, month_id)
            sid = positive_int(row["skill_id"], context)
            if sid in skill_map[month_id]:
                fail(f"DUPLICATE SKILL: job_month_id={month_id} / skill_id={sid}")
            if not row["skill_name"].strip():
                fail(f"{context}: empty skill_name for {sid}")
            if row["match_source"] not in ("scraped", "regex", "both"):
                fail(f"{context}: unknown match_source {row['match_source']!r}")
            skill_identity = (row["skill_name"], row["category"], row["role"])
            previous_identity = skill_dictionary.setdefault(sid, skill_identity)
            if previous_identity != skill_identity:
                fail(
                    f"SKILL ID CONFLICT: skill_id={sid} means {previous_identity!r} "
                    f"in one month and {skill_identity!r} in another. "
                    "Regenerate older Step 4 outputs with the current taxonomy."
                )
            skill_map[month_id][sid] = row

        ai_map: dict[int, dict[str, int]] = {}
        for row in ai_rows:
            context = f"{ai_path}: job_month_id={row['job_month_id']}"
            month_id = positive_int(row["job_month_id"], context)
            job = month_jobs.get(month_id)
            if job is None:
                fail(f"ORPHAN AI COUNT: {context} absent from Step 3 jobs")
            shared_fields_check(row, job, ai_path, month_id)
            if month_id in ai_map:
                fail(f"DUPLICATE AI COUNT: job_month_id={month_id}")
            values = {column: nonnegative_int(row[column], context + "/" + column) for column in AI_COLUMNS}
            if values["ai_mentions_total"] != values["ai_mentions_description"] + values["ai_mentions_scraped_skills"]:
                fail(f"{context}: AI total != description + scraped skills")
            if values["ai_mentions_total"] != sum(values[k] for k in (
                "ai_explicit_mentions", "ai_method_mentions", "ai_tool_mentions"
            )):
                fail(f"{context}: AI total != explicit + method + tool")
            if values["has_ai_mention"] != int(values["ai_mentions_total"] > 0):
                fail(f"{context}: has_ai_mention inconsistent with AI total")
            ai_map[month_id] = values
        if set(ai_map) != set(month_jobs):
            missing = sorted(set(month_jobs) - set(ai_map))
            fail(f"MISSING AI ROW(S) in {snapshot}: {len(missing)} job(s); examples: {missing[:5]}")

        month_skills_count = 0
        for month_id, job in month_jobs.items():
            skills = [skill_map[month_id][sid] for sid in sorted(skill_map[month_id])]
            month_skills_count += len(skills)
            ids = [int(s["skill_id"]) for s in skills]
            names = [s["skill_name"] for s in skills]
            origins = [s["match_source"] for s in skills]
            details = [{
                "skill_id": int(s["skill_id"]),
                "skill_name": s["skill_name"],
                "category": s["category"],
                "role": s["role"],
                "match_source": s["match_source"],
                "source_skill_labels": s.get("source_skill_labels", ""),
                "matched_alias": s.get("matched_alias", ""),
                "evidence_snippet": s.get("evidence_snippet", ""),
                "alias_quality": s.get("alias_quality", ""),
            } for s in skills]
            enriched: dict[str, object] = {
                **job,
                "hard_skill_id": ", ".join(map(str, ids)),
                "hard_skill": ", ".join(names),
                "hard_skill_count": len(skills),
                "hard_skill_ids_json": json.dumps(ids, ensure_ascii=False),
                "hard_skills_json": json.dumps(names, ensure_ascii=False),
                "hard_skill_sources_json": json.dumps(origins, ensure_ascii=False),
                "hard_skill_details_json": json.dumps(details, ensure_ascii=False),
                **ai_map[month_id],
            }
            all_jobs.append(enriched)

        counts_by_month.append({
            "snapshot_month": snapshot,
            "jobs": len(month_jobs),
            "job_skill_pairs": month_skills_count,
            "jobs_without_hard_skills": sum(not skill_map[mid] for mid in month_jobs),
            "ai_mentions_total": sum(ai_map[mid]["ai_mentions_total"] for mid in month_jobs),
        })
        input_files.extend([str(jobs_path.relative_to(root)), str(skills_path.relative_to(root)), str(ai_path.relative_to(root))])

    all_jobs.sort(key=lambda row: (row["snapshot_month"], int(row["job_month_id"])))
    leading = [column for column in BASE_COLUMNS_FIRST if column in all_job_fields]
    extra = sorted(all_job_fields - set(leading))
    columns = leading + extra + ENRICH_COLUMNS
    manifest = {
        "step6_version": VERSION,
        "months_included": [row["snapshot_month"] for row in counts_by_month],
        "total_job_month_observations": len(all_jobs),
        "distinct_job_ids": len({int(row["job_id"]) for row in all_jobs}),
        "total_job_skill_pairs": sum(row["job_skill_pairs"] for row in counts_by_month),
        "distinct_skill_ids": len(skill_dictionary),
        "unverifiable_fallback_fingerprints": uncheckable_fallback_fingerprints,
        "monthly_summary": counts_by_month,
        "sources": input_files,
        "id_policy": "IDs are read from Step 3 CSV and checked against persistent SQLite registry; never reassigned",
        "ai_policy": "Total counts occurrences in description + scraped skills, so a concept can be counted once in each source",
        "hard_skill_policy": "Unique canonical skill IDs per job-month, as JSON arrays and human-readable text",
    }
    return columns, all_jobs, manifest


def publish(root: Path, columns: list[str], rows: list[dict[str, object]], manifest: dict[str, object]) -> tuple[Path, Path]:
    folder = root / "data" / "final_data"
    folder.mkdir(parents=True, exist_ok=True)
    output = folder / "merged_postings_final.csv"
    summary = folder / "merged_postings_manifest.json"
    temp_files = []
    try:
        for kind, path in (("csv", output), ("json", summary)):
            fd, name = tempfile.mkstemp(prefix=".step6_", suffix=path.suffix, dir=folder)
            temp = Path(name)
            temp_files.append(temp)
            with os.fdopen(fd, "w", newline="", encoding="utf-8-sig" if kind == "csv" else "utf-8") as handle:
                if kind == "csv":
                    writer = csv.DictWriter(handle, fieldnames=columns, extrasaction="raise")
                    writer.writeheader()
                    writer.writerows(rows)
                else:
                    json.dump(manifest, handle, ensure_ascii=False, indent=2)
                    handle.write("\n")
                handle.flush()
                os.fsync(handle.fileno())
        os.replace(temp_files[0], output)
        os.replace(temp_files[1], summary)
    finally:
        for temp in temp_files:
            temp.unlink(missing_ok=True)
    return output, summary


def main() -> int:
    parser = argparse.ArgumentParser(description="Build one validated Power BI master CSV from every completed month")
    parser.add_argument("--version", action="version", version=f"%(prog)s {VERSION}")
    parser.add_argument("--project-root", help="Project root (default: parent directory of this script)")
    parser.add_argument("--check-only", action="store_true", help="Validate every month but do not write final files")
    args = parser.parse_args()
    root = Path(args.project_root).expanduser().resolve() if args.project_root else Path(__file__).resolve().parent.parent
    print(f"Step 6 version: {VERSION}")
    print(f"Project root: {root}")
    try:
        columns, rows, manifest = assemble(root)
        for month in manifest["monthly_summary"]:
            print(f"  {month['snapshot_month']}: {month['jobs']} jobs, {month['job_skill_pairs']} job-skill pairs")
        print(f"VALIDATED: {len(rows)} job-month rows, {manifest['distinct_job_ids']} distinct jobs across {len(manifest['months_included'])} month(s)")
        if args.check_only:
            print("Check only: no output files were changed")
        else:
            output, summary = publish(root, columns, rows, manifest)
            print(f"SUCCESS: {output}")
            print(f"Audit summary: {summary}")
            print("Use the same master CSV path for every Power BI refresh")
        return 0
    except (OSError, ValueError, sqlite3.Error) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        print("Existing final CSV was not intentionally replaced after validation failure", file=sys.stderr)
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
