from pathlib import Path
import pandas as pd
import re


# ============================================================
# SETTINGS
# ============================================================

INPUT_FILE = Path("merged file") / "merged.xlsx"
OUTPUT_FOLDER = Path("merged file")
OUTPUT_FILE = OUTPUT_FOLDER / "hard skill flagged.xlsx"

OUTPUT_FOLDER.mkdir(exist_ok=True)


# ============================================================
# SIMPLIFIED HARD SKILL LIST
# ============================================================

HARD_SKILLS = [
    (1, "SQL"),
    (2, "Python"),
    (3, "R"),
    (4, "DAX"),
    (5, "Power Query"),
    (6, "VBA"),
    (7, "Java"),
    (8, "Scala"),
    (9, "C#"),
    (10, "C++"),
    (11, "JavaScript"),
    (12, "TypeScript"),
    (13, "Bash"),
    (14, "PowerShell"),
    (15, "Julia"),
    (16, "MATLAB"),
    (17, "SAS"),
    (18, "SPSS"),
    (19, "Groovy"),
    (20, "Rust"),
    (21, "Go"),

    (22, "PostgreSQL"),
    (23, "MySQL"),
    (24, "Microsoft SQL Server"),
    (25, "Oracle Database"),
    (26, "SQLite"),
    (27, "MariaDB"),
    (28, "IBM Db2"),
    (29, "Snowflake"),
    (30, "BigQuery"),
    (31, "Databricks"),
    (32, "Azure SQL"),
    (33, "Amazon Redshift"),
    (34, "Amazon Athena"),
    (35, "Teradata"),
    (36, "DuckDB"),
    (37, "ClickHouse"),
    (38, "MongoDB"),
    (39, "Cassandra"),
    (40, "Redis"),
    (41, "Elasticsearch"),
    (42, "Neo4j"),
    (43, "GraphQL"),

    (44, "Microsoft Excel"),
    (45, "Power BI"),
    (46, "Tableau"),
    (47, "Qlik Sense"),
    (48, "Looker"),
    (49, "Looker Studio"),
    (50, "Amazon QuickSight"),
    (51, "SAP Analytics Cloud"),
    (52, "MicroStrategy"),
    (53, "IBM Cognos"),
    (54, "Sisense"),
    (55, "Domo"),
    (56, "ThoughtSpot"),
    (57, "Apache Superset"),
    (58, "Metabase"),
    (59, "Grafana"),

    (60, "pandas"),
    (61, "NumPy"),
    (62, "SciPy"),
    (63, "Polars"),
    (64, "PyArrow"),
    (65, "Matplotlib"),
    (66, "Seaborn"),
    (67, "Plotly"),
    (68, "scikit-learn"),
    (69, "Jupyter"),
    (70, "Anaconda"),
    (71, "Streamlit"),
    (72, "PySpark"),
    (73, "Dask"),
    (74, "RStudio"),
    (75, "tidyverse"),
    (76, "dplyr"),
    (77, "ggplot2"),
    (78, "Shiny"),

    (79, "Apache Spark"),
    (80, "Apache Kafka"),
    (81, "Apache Flink"),
    (82, "Apache Airflow"),
    (83, "dbt"),
    (84, "Dagster"),
    (85, "Prefect"),
    (86, "Azure Data Factory"),
    (87, "AWS Glue"),
    (88, "Google Cloud Dataflow"),
    (89, "Fivetran"),
    (90, "Informatica"),
    (91, "Talend"),
    (92, "SSIS"),
    (93, "Alteryx"),
    (94, "KNIME"),
    (95, "Dataiku"),

    (96, "ETL"),
    (97, "ELT"),
    (98, "Data Pipelines"),
    (99, "Data Warehousing"),
    (100, "Data Lake"),
    (101, "Data Lakehouse"),
    (102, "Data Modelling"),
    (103, "Dimensional Modelling"),
    (104, "Data Marts"),
    (105, "OLAP"),
    (106, "OLTP"),
    (107, "Star Schema"),
    (108, "Snowflake Schema"),
    (109, "Data Vault"),

    (110, "Descriptive Statistics"),
    (111, "Inferential Statistics"),
    (112, "Probability"),
    (113, "Hypothesis Testing"),
    (114, "A/B Testing"),
    (115, "Regression"),
    (116, "Logistic Regression"),
    (117, "ANOVA"),
    (118, "Time Series Analysis"),
    (119, "Forecasting"),
    (120, "ARIMA"),
    (121, "Bayesian Statistics"),
    (122, "Monte Carlo Simulation"),
    (123, "Machine Learning"),
    (124, "Supervised Learning"),
    (125, "Unsupervised Learning"),
    (126, "Classification"),
    (127, "Clustering"),
    (128, "Decision Trees"),
    (129, "Random Forest"),
    (130, "Gradient Boosting"),
    (131, "XGBoost"),
    (132, "LightGBM"),
    (133, "CatBoost"),
    (134, "Support Vector Machines"),
    (135, "PCA"),
    (136, "Feature Engineering"),
    (137, "Model Evaluation"),
    (138, "Cross-Validation"),
    (139, "Hyperparameter Tuning"),
    (140, "Anomaly Detection"),
    (141, "Recommendation Systems"),

    (142, "Generative AI"),
    (143, "Large Language Models"),
    (144, "Prompt Engineering"),
    (145, "RAG"),
    (146, "Vector Databases"),
    (147, "Embeddings"),
    (148, "Semantic Search"),
    (149, "Natural Language to SQL"),
    (150, "Document Intelligence"),

    (151, "Git"),
    (152, "GitHub"),
    (153, "GitLab"),
    (154, "Bitbucket"),
    (155, "CI/CD"),
    (156, "Docker"),
    (157, "Kubernetes"),
    (158, "Terraform"),
    (159, "REST APIs"),
    (160, "GraphQL APIs"),
    (161, "OAuth"),
    (162, "API Integration"),
    (163, "Postman"),
    (164, "Web Scraping"),
    (165, "Selenium"),

    (166, "SAP"),
    (167, "Salesforce"),
    (168, "ServiceNow"),
    (169, "Workday"),
    (170, "Microsoft Dynamics 365"),
    (171, "Oracle ERP"),

    (172, "Data Governance"),
    (173, "Data Security"),
    (174, "Access Control"),
    (175, "Data Privacy"),
    (176, "Data Lineage"),
    (177, "Metadata Management"),
    (178, "Master Data Management"),
    (179, "Data Quality"),
    (180, "Data Observability"),

    (181, "Financial Modelling"),
    (182, "Financial Analysis"),
    (183, "Revenue Analysis"),
    (184, "Profitability Analysis"),
    (185, "Budget Analysis"),
    (186, "Variance Analysis"),
    (187, "Cost Analysis"),
    (188, "Pricing Analytics"),
    (189, "Sales Analytics"),
    (190, "Workforce Analytics"),
    (191, "Supply Chain Analytics"),
    (192, "Fraud Analytics"),
    (193, "Risk Analytics"),

    (194, "Google Analytics"),
    (195, "Adobe Analytics"),
    (196, "Mixpanel"),
    (197, "Amplitude"),
    (198, "Web Analytics"),
    (199, "Product Analytics"),
    (200, "Funnel Analysis"),
    (201, "Cohort Analysis"),
    (202, "Customer Segmentation"),
    (203, "Churn Analysis"),
    (204, "Attribution Analysis"),

    (205, "Power Automate"),
    (206, "Python Automation"),
    (207, "Workflow Automation"),
    (208, "API Automation"),
    (209, "UiPath"),
    (210, "Automation Anywhere"),

    (211, "Microsoft Azure"),
    (212, "Amazon Web Services"),
    (213, "Google Cloud Platform"),
    (214, "Azure Blob Storage"),
    (215, "Amazon S3"),
    (216, "Google Cloud Storage"),

    (217, "GIS"),
    (218, "Spatial Data Analysis"),
    (219, "ArcGIS"),
    (220, "QGIS"),
    (221, "PostGIS"),
    (222, "GeoPandas"),
    (223, "Google Earth Engine"),
    (224, "Mapbox"),
    (225, "Spatial SQL"),

    (226, "Data Storytelling"),
    (227, "Dashboard Development"),
    (228, "Executive Dashboards"),
    (229, "Operational Dashboards"),
    (230, "Self-Service BI"),
    (231, "Interactive Reporting"),
    (232, "Automated Reporting"),
    (233, "KPI Development"),
    (234, "KPI Reporting"),
]


# ============================================================
# SKILLS THAT MUST BE CASE-SENSITIVE
# ============================================================
#
# These skills are matched exactly as written.
#
# Example for Microsoft Excel:
#     "Excel"          -> MATCH
#     "Microsoft Excel"-> MATCH
#     "excel"          -> NO MATCH
#     "EXCEL"          -> NO MATCH
#
# Add other skills here later if you decide capitalization
# should matter for them as well.
# ============================================================

CASE_SENSITIVE_SKILLS = {
    "Microsoft Excel",
}


# ============================================================
# NORMALIZE TEXT
# ============================================================

def normalize_text(text):
    """
    Normalizes text for comparison.

    Example:
        "  Python  " -> "python"
        "PYTHON"     -> "python"
    """

    if pd.isna(text):
        return ""

    text = str(text)

    text = re.sub(r"\s+", " ", text)

    return text.strip()


# ============================================================
# NORMALIZE EXISTING SKILL
# ============================================================

def normalize_existing_skill(skill):
    """
    Normalizes an existing skill from the skills column.

    Case-sensitive skills retain their capitalization.
    All other skills are converted to lowercase.
    """

    skill = normalize_text(skill)

    if not skill:
        return ""

    # Microsoft Excel must retain capitalization
    if skill == "Microsoft Excel":
        return skill

    # Also treat the standalone "Excel" as the canonical
    # Excel wording when checking the existing skills column.
    if skill == "Excel":
        return skill

    return skill.lower()


# ============================================================
# SPLIT EXISTING SKILLS
# ============================================================

def get_existing_skills(skills_value):
    """
    Converts the existing skills cell into a set.

    Supports comma, semicolon, pipe, and newline separators.
    """

    if pd.isna(skills_value):
        return set()

    text = str(skills_value).strip()

    if not text:
        return set()

    parts = re.split(r"[,;|\n]+", text)

    return {
        normalize_existing_skill(part)
        for part in parts
        if normalize_existing_skill(part)
    }


# ============================================================
# CHECK WHETHER SKILL EXISTS IN DESCRIPTION
# ============================================================

def skill_found_in_description(skill, description):
    """
    Searches for a skill in the job description.

    Microsoft Excel is handled specially:

        Excel            -> MATCH
        Microsoft Excel  -> MATCH
        excel            -> NO MATCH
        EXCEL            -> NO MATCH

    All other skills are matched case-insensitively.
    """

    if not description:
        return False

    escaped_skill = re.escape(skill)

    # --------------------------------------------------------
    # Microsoft Excel / Excel
    # --------------------------------------------------------

    if skill == "Microsoft Excel":

        # Match "Microsoft Excel" with exact capitalization
        pattern_microsoft_excel = (
            r"(?<![A-Za-z0-9])"
            r"Microsoft Excel"
            r"(?![A-Za-z0-9])"
        )

        # Also match standalone "Excel" with exact capitalization
        pattern_excel = (
            r"(?<![A-Za-z0-9])"
            r"Excel"
            r"(?![A-Za-z0-9])"
        )

        return bool(
            re.search(
                pattern_microsoft_excel,
                description
            )
            or
            re.search(
                pattern_excel,
                description
            )
        )

    # --------------------------------------------------------
    # All other skills: case-insensitive
    # --------------------------------------------------------

    pattern = (
        r"(?<![A-Za-z0-9])"
        + escaped_skill
        + r"(?![A-Za-z0-9])"
    )

    return bool(
        re.search(
            pattern,
            description,
            flags=re.IGNORECASE
        )
    )


# ============================================================
# CHECK WHETHER SKILL IS ALREADY IN EXISTING SKILLS COLUMN
# ============================================================

def skill_already_exists(skill_name, existing_skills):
    """
    Determines whether a detected hard skill already exists
    in the row's skills column.
    """

    # --------------------------------------------------------
    # Microsoft Excel
    # --------------------------------------------------------

    if skill_name == "Microsoft Excel":

        # Only exact capitalization counts
        return (
            "Microsoft Excel" in existing_skills
            or "Excel" in existing_skills
        )

    # --------------------------------------------------------
    # All other skills
    # --------------------------------------------------------

    return normalize_text(skill_name).lower() in {
        skill.lower()
        for skill in existing_skills
    }


# ============================================================
# FIND HARD SKILLS FOR ONE ROW
# ============================================================

def find_hard_skills(job_description, existing_skills):

    description = normalize_text(job_description)

    existing = get_existing_skills(existing_skills)

    found_ids = []
    found_names = []

    # Check every canonical hard skill
    for skill_id, skill_name in HARD_SKILLS:

        # ----------------------------------------------------
        # Is it mentioned in the job description?
        # ----------------------------------------------------

        if not skill_found_in_description(
            skill_name,
            description
        ):
            continue

        # ----------------------------------------------------
        # Is it already in the skills column?
        # ----------------------------------------------------

        if skill_already_exists(
            skill_name,
            existing
        ):
            continue

        # ----------------------------------------------------
        # It is a missing hard skill -> flag it
        # ----------------------------------------------------

        found_ids.append(str(skill_id))
        found_names.append(skill_name)

    return (
        ", ".join(found_ids),
        ", ".join(found_names)
    )


# ============================================================
# CHECK INPUT FILE
# ============================================================

if not INPUT_FILE.exists():

    raise FileNotFoundError(
        f"\nInput file was not found:\n"
        f"{INPUT_FILE.resolve()}\n"
    )


# ============================================================
# OPEN WORKBOOK
# ============================================================

excel_file = pd.ExcelFile(INPUT_FILE)

print()
print("=" * 60)
print("INPUT FILE")
print("=" * 60)
print(INPUT_FILE.resolve())

print()
print("Sheets found:")

for sheet_name in excel_file.sheet_names:
    print(f"  - {sheet_name}")


# ============================================================
# PROCESS ALL SHEETS
# ============================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    for sheet_name in excel_file.sheet_names:

        print()
        print("=" * 60)
        print(f"PROCESSING: {sheet_name}")
        print("=" * 60)

        # ----------------------------------------------------
        # READ SHEET
        # ----------------------------------------------------

        df = pd.read_excel(
            INPUT_FILE,
            sheet_name=sheet_name
        )

        # ----------------------------------------------------
        # CLEAN COLUMN NAMES
        # ----------------------------------------------------

        df.columns = [
            str(column).strip()
            for column in df.columns
        ]

        # ----------------------------------------------------
        # FIND COLUMNS REGARDLESS OF COLUMN ORDER
        # ----------------------------------------------------

        job_description_column = None
        skills_column = None

        for column in df.columns:

            normalized_column = (
                str(column)
                .strip()
                .lower()
            )

            if normalized_column == "job_description":
                job_description_column = column

            elif normalized_column == "skills":
                skills_column = column

        print(
            f"job_description column: "
            f"{job_description_column}"
        )

        print(
            f"skills column: "
            f"{skills_column}"
        )

        # ----------------------------------------------------
        # CHECK FOR JOB DESCRIPTION
        # ----------------------------------------------------

        if job_description_column is None:

            print(
                "WARNING: job_description column not found."
            )

            df["hard_skill_id"] = ""
            df["hard_skill"] = ""

        else:

            # ------------------------------------------------
            # PROCESS EACH ROW
            # ------------------------------------------------

            hard_skill_ids = []
            hard_skill_names = []

            for _, row in df.iterrows():

                job_description = row[
                    job_description_column
                ]

                if skills_column is not None:
                    current_skills = row[
                        skills_column
                    ]
                else:
                    current_skills = ""

                skill_ids, skill_names = find_hard_skills(
                    job_description,
                    current_skills
                )

                hard_skill_ids.append(skill_ids)
                hard_skill_names.append(skill_names)

            # ------------------------------------------------
            # ADD / REPLACE OUTPUT COLUMNS
            # ------------------------------------------------

            df["hard_skill_id"] = hard_skill_ids
            df["hard_skill"] = hard_skill_names

        # ----------------------------------------------------
        # SAVE SHEET
        # ----------------------------------------------------

        output_sheet_name = sheet_name[:31]

        df.to_excel(
            writer,
            sheet_name=output_sheet_name,
            index=False
        )

        # ----------------------------------------------------
        # REPORT RESULTS
        # ----------------------------------------------------

        rows_with_hard_skills = sum(
            bool(value)
            for value in df["hard_skill"]
        )

        print(
            f"Rows processed: {len(df)}"
        )

        print(
            f"Rows with new hard skills: "
            f"{rows_with_hard_skills}"
        )

        print(
            f"Saved sheet: {output_sheet_name}"
        )


# ============================================================
# FINISHED
# ============================================================

print()
print("=" * 60)
print("DONE!")
print("=" * 60)
print()
print("Output file:")
print(OUTPUT_FILE.resolve())
print()