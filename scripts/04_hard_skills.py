
"""04_hard_skills.py — map original LinkedIn skills and find extra skills in descriptions.

Place this file in <project>/scripts/04_hard_skills.py.
Place the attached workbook in <project>/taxonomy/data_analyst_engineer_skill_taxonomy.xlsx.

Input:  data/04_data_with_id/September/jobs_2026-09.csv
Output: data/05_hard_skills/September/job_skills_2026-09.csv
        data/05_hard_skills/September/unmapped_skills_2026-09.csv
        data/05_hard_skills/September/taxonomy_issues_2026-09.csv
        data/05_hard_skills/September/skill_summary_2026-09.json
        data/05_hard_skills/September/job_skills_2026-09.xlsx (optional QA)

Retains original (scraped) skills and adds descriptions' newly detected skills.
Each (job_month_id, skill_id) occurs at most once. No changes to Step 3 files.
A skill mention is not evidence that the skill is a *required* qualification.

Run from the project folder:
    python scripts/04_hard_skills.py
    python scripts/04_hard_skills.py --month October --year 2026
    python scripts/04_hard_skills.py --month September --year 2026 --overwrite
    python scripts/04_hard_skills.py --validate

Dependencies: pip install openpyxl
"""
from __future__ import annotations

import argparse
import csv
import hashlib
import json
import os
import re
import sys
import tempfile
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path

from openpyxl import Workbook, load_workbook
from openpyxl.cell import WriteOnlyCell
from openpyxl.styles import Alignment, Font, PatternFill

VERSION = "4.2.0-new-skills"

MONTH_NUMBERS = dict(zip(
    ("january", "february", "march", "april", "may", "june", "july", "august",
     "september", "october", "november", "december"), range(1, 13)))

BRIDGE_COLUMNS = [
    "job_id", "job_month_id", "snapshot_month", "market_country", "search_city",
    "skill_id", "skill_name", "category", "role", "match_source",
    "source_skill_labels", "matched_alias", "evidence_snippet", "alias_quality",
]

UNMAPPED_COLUMNS = [
    "job_id", "job_month_id", "snapshot_month", "source_skill",
    "market_country", "search_city",
]

ISSUE_COLUMNS = ["skill_id", "skill_name", "issue", "detail"]

REQUIRED_JOB_COLUMNS = [
    "job_id", "job_month_id", "snapshot_month", "job_description",
    "skills_json", "market_country", "search_city",
]


@dataclass(frozen=True)
class Skill:
    skill_id: int
    name: str
    category: str
    role: str
    aliases: tuple[str, ...]
    alias_quality: str


@dataclass(frozen=True)
class Candidate:
    skill_id: int
    alias: str
    start: int
    end: int
    priority: tuple[int, int, int]


def normalized(text: str) -> str:
    """NFKC + whitespace normalization, preserving technical punctuation."""
    return re.sub(
        r"\s+", " ", unicodedata.normalize("NFKC", text)
    ).strip()


def alias_pattern(alias: str, anchored: bool = False) -> re.Pattern[str]:
    # Do not use \b around nonword-ending tokens such as C# or C++.
    parts = normalized(alias).split(" ")
    body = r"\s+".join(re.escape(part) for part in parts)
    expr = rf"(?<!\w){body}(?!\w)"
    if anchored:
        expr = rf"^{expr}$"
    return re.compile(expr, flags=re.IGNORECASE)


def load_taxonomy(path: Path) -> tuple[list[Skill], list[dict]]:
    if not path.is_file():
        raise FileNotFoundError(
            f"Taxonomy workbook missing: {path}\n"
            "Save the supplied Excel workbook under the taxonomy/ folder, "
            "or pass --taxonomy PATH"
        )

    wb = load_workbook(path, read_only=True, data_only=True)

    try:
        if "Taxonomy" not in wb.sheetnames:
            raise ValueError(
                "Taxonomy workbook needs a sheet named 'Taxonomy'"
            )

        ws = wb["Taxonomy"]
        rows = ws.iter_rows(values_only=True)
        headers = next(rows, ())

        required = (
            "Skill ID",
            "Standardized Skill",
            "Aliases (semicolon separated)",
            "Category",
            "Role",
        )

        positions = {
            str(value).strip(): i
            for i, value in enumerate(headers)
            if value is not None
        }

        missing = [name for name in required if name not in positions]

        if missing:
            raise ValueError(f"Taxonomy missing headers: {missing}")

        skills, issues = [], []
        ids, names = set(), set()

        for excel_row, raw in enumerate(rows, start=2):
            if all(
                value is None or str(value).strip() == ""
                for value in raw
            ):
                continue

            def cell(field):
                i = positions[field]
                return (
                    str(raw[i]).strip()
                    if i < len(raw) and raw[i] is not None
                    else ""
                )

            raw_id = cell("Skill ID")
            name = cell("Standardized Skill")

            if not raw_id or not name:
                raise ValueError(
                    f"Taxonomy row {excel_row}: skill ID and name required"
                )

            try:
                skill_id = int(raw_id)
            except (ValueError, TypeError) as exc:
                raise ValueError(
                    f"Taxonomy row {excel_row}: invalid skill ID {raw_id!r}"
                ) from exc

            if skill_id <= 0 or skill_id in ids:
                raise ValueError(
                    f"Taxonomy row {excel_row}: duplicate/invalid "
                    f"skill ID {skill_id}"
                )

            if name.casefold() in names:
                raise ValueError(
                    f"Taxonomy row {excel_row}: duplicate name {name!r}"
                )

            ids.add(skill_id)
            names.add(name.casefold())

            raw_aliases = cell("Aliases (semicolon separated)")

            alias_quality = (
                "curated" if raw_aliases else "canonical_fallback"
            )

            aliases = [
                s.strip()
                for s in raw_aliases.split(";")
                if s.strip()
            ]

            if not aliases:
                aliases = [name]

                issues.append({
                    "skill_id": skill_id,
                    "skill_name": name,
                    "issue": "missing_aliases",
                    "detail": (
                        "Using canonical name only; "
                        "review before production"
                    ),
                })

            category = cell("Category")
            role = cell("Role")

            if not category or not role:
                issues.append({
                    "skill_id": skill_id,
                    "skill_name": name,
                    "issue": "missing_metadata",
                    "detail": "Category and/or Role not set",
                })

            # Keep canonical display name available as a strict source-skills
            # value; don't automatically add it as a description alias if
            # curated aliases exist.
            distinct = tuple(dict.fromkeys(aliases))

            skills.append(
                Skill(
                    skill_id,
                    name,
                    category,
                    role,
                    distinct,
                    alias_quality,
                )
            )

        # Ambiguous duplicate aliases across different IDs need explicit
        # deterministic ownership.
        by_alias = defaultdict(list)

        for skill in skills:
            for alias in skill.aliases:
                by_alias[normalized(alias).casefold()].append(skill)

        for alias, entries in by_alias.items():
            if len({entry.skill_id for entry in entries}) > 1:
                issues.append({
                    "skill_id": ";".join(
                        str(s.skill_id) for s in entries
                    ),
                    "skill_name": ";".join(
                        s.name for s in entries
                    ),
                    "issue": "shared_alias",
                    "detail": (
                        f"{alias!r} belongs to multiple IDs; "
                        "prefer canonical name owner"
                    ),
                })

        return skills, issues

    finally:
        wb.close()


# Ambiguous canonical-only entries are allowed in scraped skill lists,
# but bare ordinary English words require explicit technology context
# in descriptions.

BARE_TECH_CONTEXT = re.compile(
    r"\b(?:apache|data|pipeline|engineering|developer|programming|language|script|"
    r"software|tools?|stack|platform|technology|framework|big\s+data|"
    r"cloud|backend|codebase|kubernetes|python|sql|spark|hadoop)\b",
    flags=re.IGNORECASE,
)

EXCEL_VERB_AFTER = re.compile(
    r"^\s+(?:at|in|as|within|under|through)\b",
    re.I,
)

EXCEL_VERB_BEFORE = re.compile(
    r"(?:\b(?:to|will|can|who|may|must|should)\s+|\bhelp\s+you\s+)$",
    re.I,
)

R_CONTEXT = re.compile(
    r"\b(?:python|sql|r\s+programming|r\s+language|rstudio|statistic(?:s|al)|"
    r"analytics|programming|developer|proficien(?:t|cy)|experience|knowledge|"
    r"skilled?|familiar|language|code)\b",
    re.I,
)

R_LIST_AFTER = re.compile(
    r"^\s*(?:[,;/|&+]|\band\b)\s*"
    r"(?:Python|SQL|SAS|Scala|Julia|MATLAB|SPSS)\b",
    re.I,
)

R_LIST_BEFORE = re.compile(
    r"\b(?:Python|SQL|SAS|Scala|Julia|MATLAB|SPSS)"
    r"\s*(?:[,;/|&+]|\band\b)\s*$",
    re.I,
)


def is_acceptable(
    text: str,
    skill: Skill,
    alias: str,
    start: int,
    end: int,
) -> bool:
    """Conservative disambiguation for known high-risk expressions."""

    hit = text[start:end]
    alias_lower = normalized(alias).casefold()

    left = text[max(0, start - 55):start]
    right = text[end:min(len(text), end + 55)]
    nearby = left + " " + right

    if skill.name == "R" and alias_lower == "r":
        return (
            hit == "R"
            and bool(
                R_LIST_AFTER.search(right)
                or R_LIST_BEFORE.search(left)
                or (
                    R_CONTEXT.search(nearby)
                    and re.search(
                        r"\b(?:in|with|of|for|using|knowledge|experience)\s*$",
                        left,
                        re.I,
                    )
                )
            )
        )

    if skill.name == "Excel" and alias_lower == "excel":
        if (
            EXCEL_VERB_AFTER.search(right)
            or EXCEL_VERB_BEFORE.search(left)
        ):
            return False

        # Capitalized Excel is strong evidence of the product.
        # For lowercase, demand skill/data/spreadsheet context.
        return hit[:1].isupper() or bool(
            re.search(
                r"\b(?:spreadsheet|workbook|pivot|formula|macro|skills?|"
                r"microsoft|proficient|advanced|automation|reporting|"
                r"sql|python|vba|power\s+bi|tableau)\b",
                nearby,
                re.I,
            )
        )

    if skill.name == "HTML" and alias_lower == "html":
        # HTML tags in scraped HTML do not imply HTML knowledge.
        # Still accept natural-language mentions such as "HTML and CSS".
        if (
            re.search(r"<\s*/?\s*$", left)
            or re.search(r"<!doctype\s*$", left, re.I)
        ):
            return False

        if re.match(r"\s*>", right) and left.rstrip().endswith("<"):
            return False

    if skill.name == "ORC" and alias_lower == "orc":
        return hit == "ORC"

    if (
        skill.name in ("Apache Spark", "Hive", "Apache Beam")
        and alias_lower in ("spark", "hive", "beam")
    ):
        return bool(BARE_TECH_CONTEXT.search(nearby))

    if skill.name == "Go" and alias_lower == "go":
        # Reject common verb use: go to, go through, go for, etc.
        if re.match(
            r"\s+(?:to|through|for|ahead|back|with|into|on)\b",
            right,
            re.I,
        ):
            return False

        return hit == "Go" and bool(
            re.search(
                r"\b(?:golang|programming|developer|backend|language|"
                r"microservices|code|software|engineer|python|java|"
                r"rust|typescript)\b",
                nearby,
                re.I,
            )
        )

    if skill.name == "React" and alias_lower == "react":
        return hit == "React" and bool(
            re.search(
                r"\b(?:js|javascript|typescript|frontend|front.end|"
                r"framework|developer|web|node|angular|vue|component)\b",
                nearby,
                re.I,
            )
        )

    if skill.name == "Rust" and alias_lower == "rust":
        return hit == "Rust" and bool(
            BARE_TECH_CONTEXT.search(nearby)
        )

    if skill.name == "Agile" and alias_lower == "agile":
        return bool(
            re.search(
                r"\b(?:scrum|methodolog|project|development|"
                r"software|team|sprint)\b",
                nearby,
                re.I,
            )
        )

    if skill.name == "Piwik" and alias_lower == "piwik":
        # Matomo (formerly Piwik) and Piwik PRO are separate products.
        if re.match(r"\s+PRO\b", right, re.I):
            return False

    if skill.name == "Scrum" and alias_lower == "scrum":
        # Rugby usage is not the software-delivery methodology.
        context = text[
            max(0, start - 70):min(len(text), end + 70)
        ]

        sports = re.search(
            r"\b(?:rugby|sport|players?|tackle|scrummage)\b",
            context,
            re.I,
        )

        project = re.search(
            r"\b(?:agile|software|development|project|sprint|"
            r"delivery|methodology)\b",
            context,
            re.I,
        )

        if sports and not project:
            return False

    if skill.name == "JavaScript" and alias_lower == "js":
        # Bare two-letter aliases are only safe in software context.
        if hit != "JS" and not re.search(
            r"\b(?:javascript|typescript|web|front.end|back.end|"
            r"node|react|coding|programming|developer|framework)\b",
            nearby,
            re.I,
        ):
            return False

    if skill.name == "CRM" and alias_lower == "crm":
        return hit.upper() == "CRM"

    if skill.name == "Power BI" and alias_lower == "pbi":
        return bool(BARE_TECH_CONTEXT.search(nearby))

    return True


class Matcher:
    def __init__(self, skills: list[Skill]):
        self.skills = {s.skill_id: s for s in skills}

        self.alias_owners = defaultdict(list)
        self.patterns = []

        for skill in skills:
            for alias in skill.aliases:
                self.alias_owners[
                    normalized(alias).casefold()
                ].append(skill.skill_id)

                self.patterns.append(
                    (skill, alias, alias_pattern(alias))
                )

        # Map original scraped skill-list values by exact alias equality.
        # Some curated lists omit the canonical name, so allow it too.
        self.original_patterns = []

        for skill in skills:
            for alias in list(skill.aliases) + [skill.name]:
                self.original_patterns.append(
                    (
                        skill,
                        alias,
                        alias_pattern(alias, anchored=True),
                    )
                )

    def source_matches(self, label: str) -> list[tuple[int, str]]:
        label = normalized(label)

        if not label:
            return []

        matches = [
            (skill.skill_id, alias)
            for skill, alias, pat in self.original_patterns
            if pat.fullmatch(label)
        ]

        if not matches:
            return []

        # Shared alias: prefer skill whose canonical name matches.
        candidates = sorted({sid for sid, _ in matches})

        if len(candidates) > 1:
            canonical = [
                sid
                for sid in candidates
                if self.skills[sid].name.casefold() == label.casefold()
            ]

            if len(canonical) == 1:
                candidates = canonical
            else:
                # Ambiguous label is not accepted into the bridge.
                return []

        sid = candidates[0]

        return [
            (
                sid,
                next(
                    alias
                    for match_sid, alias in matches
                    if match_sid == sid
                ),
            )
        ]

    def description_matches(
        self,
        description: str,
    ) -> dict[int, tuple[str, str]]:

        text = normalized(description)

        if not text:
            return {}

        candidates = []

        for skill, alias, pattern in self.patterns:
            for match in pattern.finditer(text):
                if not is_acceptable(
                    text,
                    skill,
                    alias,
                    match.start(),
                    match.end(),
                ):
                    continue

                # Longest / most-specific alias first.
                # Canonical owner wins ties.
                # A single span belongs to only one skill.
                canonical_owner = int(
                    normalized(alias).casefold()
                    == normalized(skill.name).casefold()
                )

                priority = (
                    match.end() - match.start(),
                    canonical_owner,
                    -skill.skill_id,
                )

                candidates.append(
                    Candidate(
                        skill.skill_id,
                        alias,
                        match.start(),
                        match.end(),
                        priority,
                    )
                )

        candidates.sort(
            key=lambda c: (
                -c.priority[0],
                -c.priority[1],
                -c.priority[2],
                c.start,
            )
        )

        used_spans = []
        found = {}

        for candidate in candidates:
            if any(
                candidate.start < end
                and start < candidate.end
                for start, end in used_spans
            ):
                continue

            used_spans.append(
                (candidate.start, candidate.end)
            )

            if candidate.skill_id not in found:
                a = max(0, candidate.start - 55)
                b = min(len(text), candidate.end + 55)

                snippet = text[a:b]

                found[candidate.skill_id] = (
                    candidate.alias,
                    snippet,
                )

        return found


def load_rows(csv_path: Path, month: str):
    csv.field_size_limit(
        min(sys.maxsize, 2**31 - 1)
    )

    with csv_path.open(
        "r",
        encoding="utf-8-sig",
        newline="",
    ) as f:

        reader = csv.DictReader(f)
        fields = reader.fieldnames or []

        if len(fields) != len(set(fields)):
            raise ValueError("Duplicate job-file headers")

        missing = [
            col for col in REQUIRED_JOB_COLUMNS
            if col not in fields
        ]

        if missing:
            raise ValueError(
                f"Step 3 output missing fields: {missing}"
            )

        seen_month_ids = set()

        for lineno, row in enumerate(reader, 2):
            if None in row or any(
                v is None for v in row.values()
            ):
                raise ValueError(
                    f"Job file row {lineno} has invalid CSV column count"
                )

            if row["snapshot_month"] != month:
                raise ValueError(
                    f"Job file row {lineno}: incorrect snapshot_month"
                )

            for key in ("job_id", "job_month_id"):
                if (
                    not row[key].isdigit()
                    or int(row[key]) < 1
                ):
                    raise ValueError(
                        f"Job file row {lineno}: invalid {key}"
                    )

            if row["job_month_id"] in seen_month_ids:
                raise ValueError(
                    f"Duplicate job_month_id: "
                    f"{row['job_month_id']}"
                )

            seen_month_ids.add(
                row["job_month_id"]
            )

            try:
                scraped = json.loads(
                    row["skills_json"]
                )
            except json.JSONDecodeError as exc:
                raise ValueError(
                    f"Job file row {lineno}: "
                    "skills_json is not valid JSON"
                ) from exc

            if (
                not isinstance(scraped, list)
                or not all(
                    isinstance(x, str)
                    for x in scraped
                )
            ):
                raise ValueError(
                    f"Job file row {lineno}: "
                    "skills_json must be a string list"
                )

            yield row, scraped


def produce_matches(
    csv_path: Path,
    month: str,
    matcher: Matcher,
):
    bridge, unmapped = [], []
    stats = Counter()
    distinct_raw_unmapped = Counter()

    for row, original_skills in load_rows(
        csv_path,
        month,
    ):
        stats["jobs"] += 1

        source_map = defaultdict(list)
        source_alias = {}

        for skill_label in original_skills:
            matches = matcher.source_matches(
                skill_label
            )

            if not matches:
                unmapped.append({
                    "job_id": int(row["job_id"]),
                    "job_month_id": int(row["job_month_id"]),
                    "snapshot_month": month,
                    "source_skill": skill_label,
                    "market_country": row["market_country"],
                    "search_city": row["search_city"],
                })

                distinct_raw_unmapped[
                    skill_label
                ] += 1

                continue

            for sid, alias in matches:
                source_map[sid].append(
                    skill_label
                )

                source_alias.setdefault(
                    sid,
                    alias,
                )

        description_hits = matcher.description_matches(
            row["job_description"]
        )

        all_skill_ids = (
            set(source_map)
            | set(description_hits)
        )

        if not all_skill_ids:
            stats["no_mapped_skills"] += 1

        for skill_id in sorted(all_skill_ids):
            skill = matcher.skills[skill_id]

            from_scrape = skill_id in source_map
            from_desc = skill_id in description_hits

            if from_scrape and from_desc:
                source = "both"
            elif from_scrape:
                source = "scraped"
            else:
                source = "regex"

            alias, evidence = description_hits.get(
                skill_id,
                (
                    source_alias.get(skill_id, ""),
                    "",
                ),
            )

            bridge.append({
                "job_id": int(row["job_id"]),
                "job_month_id": int(row["job_month_id"]),
                "snapshot_month": month,
                "market_country": row["market_country"],
                "search_city": row["search_city"],
                "skill_id": skill_id,
                "skill_name": skill.name,
                "category": skill.category,
                "role": skill.role,
                "match_source": source,
                "source_skill_labels": json.dumps(
                    source_map.get(skill_id, []),
                    ensure_ascii=False,
                ),
                "matched_alias": alias,
                "evidence_snippet": evidence,
                "alias_quality": skill.alias_quality,
            })

            stats[source] += 1

        stats["source_skill_labels"] += len(
            original_skills
        )

    assert len(bridge) == len({
        (b["job_month_id"], b["skill_id"])
        for b in bridge
    })

    stats["job_skill_pairs"] = len(bridge)
    stats["unmapped_source_labels"] = len(unmapped)

    return (
        bridge,
        unmapped,
        dict(stats),
        distinct_raw_unmapped,
    )


def write_csv(
    path: Path,
    columns: list[str],
    rows: list[dict],
):
    with path.open(
        "w",
        encoding="utf-8-sig",
        newline="",
    ) as f:
        writer = csv.DictWriter(
            f,
            fieldnames=columns,
            extrasaction="raise",
        )
        writer.writeheader()
        writer.writerows(rows)


def write_excel(
    path: Path,
    rows: list[dict],
):
    if len(rows) > 1_048_575:
        raise ValueError(
            "Excel row limit exceeded; use --csv-only"
        )

    book = Workbook(write_only=True)

    try:
        sheet = book.create_sheet(
            "Job skills"
        )

        sheet.freeze_panes = "A2"
        sheet.sheet_view.showGridLines = False

        for idx, col in enumerate(
            BRIDGE_COLUMNS,
            start=1,
        ):
            if idx <= 2:
                sheet.column_dimensions[
                    chr(64 + idx)
                ].width = 16

        headings = []

        for col in BRIDGE_COLUMNS:
            cell = WriteOnlyCell(
                sheet,
                value=col,
            )

            cell.fill = PatternFill(
                "solid",
                fgColor="17324D",
            )

            cell.font = Font(
                bold=True,
                color="FFFFFF",
            )

            cell.alignment = Alignment(
                wrap_text=True,
            )

            headings.append(cell)

        sheet.append(headings)

        for row in rows:
            cells = []

            for field in BRIDGE_COLUMNS:
                value = row[field]

                if (
                    isinstance(value, str)
                    and len(value) > 32767
                ):
                    raise ValueError(
                        f"Value in {field} exceeds "
                        "Excel 32,767 characters; "
                        "run --csv-only"
                    )

                c = WriteOnlyCell(
                    sheet,
                    value=value,
                )

                if isinstance(value, str):
                    # Prevent formula injection through scraped text.
                    c.data_type = "s"

                cells.append(c)

            sheet.append(cells)

        book.save(path)

    finally:
        book.close()


def taxonomy_validation_examples(
    path: Path,
    matcher: Matcher,
):
    wb = load_workbook(
        path,
        read_only=True,
        data_only=True,
    )

    errors = []

    try:
        if "Validation Examples" not in wb.sheetnames:
            return [
                "Workbook is missing the "
                "'Validation Examples' sheet"
            ]

        for raw in list(
            wb["Validation Examples"].values
        )[1:]:
            if not raw or raw[0] is None:
                continue

            number, snippet, expected = raw[:3]

            if not isinstance(snippet, str):
                continue

            expected_names = (
                set()
                if expected in (None, "None")
                else {
                    name.strip()
                    for name in str(expected).split(";")
                    if name.strip()
                }
            )

            observed = {
                matcher.skills[sid].name
                for sid in matcher.description_matches(
                    snippet
                )
            }

            if observed != expected_names:
                errors.append(
                    f"Test {number}: "
                    f"expected {sorted(expected_names)} "
                    f"but found {sorted(observed)}"
                )

        return errors

    finally:
        wb.close()


def run(args) -> int:
    month = args.month.strip().casefold()

    if month not in MONTH_NUMBERS:
        raise ValueError(
            f"Unknown month: {args.month}"
        )

    if not 2000 <= args.year <= 2100:
        raise ValueError(
            "Year outside 2000-2100"
        )

    root = Path(__file__).resolve().parent.parent

    label = month.capitalize()

    snapshot = (
        f"{args.year:04d}-"
        f"{MONTH_NUMBERS[month]:02d}"
    )

    input_path = (
        root
        / "data"
        / "04_data_with_id"
        / label
        / f"jobs_{snapshot}.csv"
    )

    taxonomy_path = (
        Path(args.taxonomy).expanduser().resolve()
        if args.taxonomy
        else (
            root
            / "taxonomy"
            / "data_analyst_engineer_skill_taxonomy_reviewed.xlsx"
        )
    )

    output_folder = (
        Path(args.output_dir).expanduser().resolve()
        if args.output_dir
        else (
            root
            / "data"
            / "05_hard_skills"
            / label
        )
    )

    targets = {
        "bridge": (
            output_folder
            / f"job_skills_{snapshot}.csv"
        ),
        "unmapped": (
            output_folder
            / f"unmapped_skills_{snapshot}.csv"
        ),
        "issues": (
            output_folder
            / f"taxonomy_issues_{snapshot}.csv"
        ),
        "summary": (
            output_folder
            / f"skill_summary_{snapshot}.json"
        ),
    }

    if not args.csv_only and not args.validate:
        targets["excel"] = (
            output_folder
            / f"job_skills_{snapshot}.xlsx"
        )

    print("Script version :", VERSION)
    print("Job input      :", input_path)
    print("Taxonomy       :", taxonomy_path)
    print("Output folder  :", output_folder)

    skills, issues = load_taxonomy(
        taxonomy_path
    )

    matcher = Matcher(skills)

    validation_errors = taxonomy_validation_examples(
        taxonomy_path,
        matcher,
    )

    print(
        f"Taxonomy: {len(skills)} distinct skill IDs; "
        f"{sum(s.alias_quality == 'canonical_fallback' for s in skills)} "
        "with no curated aliases"
    )

    if validation_errors:
        for failure in validation_errors:
            print(
                "VALIDATION FAIL:",
                failure,
                file=sys.stderr,
            )

        raise ValueError(
            f"{len(validation_errors)} taxonomy "
            "validation example(s) failed. "
            "No exports written."
        )

    print(
        "Taxonomy validation examples: PASS"
    )

    if args.validate:
        return 0

    if not input_path.is_file():
        raise FileNotFoundError(
            f"Step 3 job file missing: {input_path}; "
            "run 03_unique_id.py first"
        )

    existing = [
        p
        for p in targets.values()
        if p.exists()
    ]

    if existing and not args.overwrite:
        raise FileExistsError(
            f"Existing outputs: {existing}; "
            "use --overwrite to regenerate"
        )

    bridge, unmapped, counts, unmapped_top = produce_matches(
        input_path,
        snapshot,
        matcher,
    )

    output_folder.mkdir(
        parents=True,
        exist_ok=True,
    )

    summary = {
        "snapshot_month": snapshot,
        "taxonomy_file": taxonomy_path.name,
        "taxonomy_sha256": hashlib.sha256(
            taxonomy_path.read_bytes()
        ).hexdigest(),
        "unique_taxonomy_skills": len(skills),
        "taxonomy_issues": len(issues),
        **counts,
        "top_unmapped_skill_labels": (
            unmapped_top.most_common(30)
        ),
        "interpretation": (
            "Skills mentioned in scraped skills "
            "or descriptions; not necessarily required"
        ),
    }

    # Stage every output before replacing final files.
    staging = []

    try:
        for key, target in targets.items():
            fd, tempname = tempfile.mkstemp(
                dir=output_folder,
                prefix=".step4_",
                suffix=target.suffix,
            )

            os.close(fd)

            temp = Path(tempname)
            staging.append((temp, target))

            if key == "bridge":
                write_csv(
                    temp,
                    BRIDGE_COLUMNS,
                    bridge,
                )

            elif key == "unmapped":
                write_csv(
                    temp,
                    UNMAPPED_COLUMNS,
                    unmapped,
                )

            elif key == "issues":
                write_csv(
                    temp,
                    ISSUE_COLUMNS,
                    issues,
                )

            elif key == "summary":
                temp.write_text(
                    json.dumps(
                        summary,
                        ensure_ascii=False,
                        indent=2,
                    ),
                    encoding="utf-8",
                )

            elif key == "excel":
                write_excel(
                    temp,
                    bridge,
                )

        for temp, target in staging:
            os.replace(
                temp,
                target,
            )

    finally:
        for temp, _ in staging:
            temp.unlink(
                missing_ok=True
            )

    print(
        "\nSUCCESS:",
        counts["jobs"],
        "jobs processed;",
    )

    print(
        "  Unique job-month-skill pairs:",
        counts["job_skill_pairs"],
    )

    print(
        "  Origin: scraped:",
        counts.get("scraped", 0),
        "regex:",
        counts.get("regex", 0),
        "both:",
        counts.get("both", 0),
    )

    print(
        "  Jobs with no mapped skill:",
        counts.get("no_mapped_skills", 0),
        "(still retained in Step 3 job table)",
    )

    print(
        "  Unmapped source labels:",
        counts.get("unmapped_source_labels", 0),
        "(see audit CSV)",
    )

    print(
        "  Taxonomy issues:",
        len(issues),
        "(see issues CSV)",
    )

    print(
        "  Export:",
        targets["bridge"],
    )

    print(
        "Commit script + reviewed taxonomy; "
        "do not commit generated job-posting exports."
    )

    return 0


def main() -> int:
    parser = argparse.ArgumentParser(
        description=(
            "Link scraped/regex skills to canonical skill IDs"
        )
    )

    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {VERSION}",
    )

    parser.add_argument(
        "--month",
        default="September",
    )

    parser.add_argument(
        "--year",
        type=int,
        default=2026,
    )

    parser.add_argument(
        "--taxonomy",
        help=(
            "Path to approved taxonomy .xlsx "
            "(default: <project>/taxonomy/"
            "data_analyst_engineer_skill_taxonomy.xlsx)"
        ),
    )

    parser.add_argument(
        "--output-dir",
        help=(
            "Override default "
            "<project>/data/05_hard_skills/<month>"
        ),
    )

    parser.add_argument(
        "--overwrite",
        action="store_true",
    )

    parser.add_argument(
        "--csv-only",
        action="store_true",
    )

    parser.add_argument(
        "--validate",
        action="store_true",
        help=(
            "Run taxonomy's validation examples "
            "without reading jobs"
        ),
    )

    args = parser.parse_args()

    try:
        return run(args)

    except Exception as exc:
        print(
            "ERROR:",
            exc,
            file=sys.stderr,
        )
        return 1


if __name__ == "__main__":
    sys.exit(main())
