import re
import pandas as pd

# ==========================================
# 1. HARD SKILLS TAXONOMY & ALIAS MAPPING
# ==========================================

SKILL_TAXONOMY = {
    1: {"name": "SQL", "aliases": [r"\bSQL\b"]},
    2: {"name": "Python", "aliases": [r"\bPython\b"]},
    3: {
        "name": "R",
        "aliases": [r"\bRStudio\b", r"\bR Studio\b", r"\bR[- ]language\b", r"\bR[- ]programming\b", r"\bR\b"]
    },
    4: {"name": "Scala", "aliases": [r"\bScala\b"]},
    5: {"name": "Java", "aliases": [r"\bJava\b"]},
    6: {"name": "Bash", "aliases": [r"\bBash\b"]},
    7: {"name": "PowerShell", "aliases": [r"\bPowerShell\b"]},
    8: {"name": "VBA", "aliases": [r"\bVBA\b", r"\bVisual Basic\b"]},
    9: {
        "name": "Excel",
        "aliases": [r"\bMicrosoft Excel\b", r"\bMS Excel\b", r"\bExcel\b"]
    },
    10: {"name": "Power BI", "aliases": [r"\bPower BI\b", r"\bPowerBI\b", r"\bPBI\b"]},
    11: {"name": "Tableau", "aliases": [r"\bTableau\b"]},
    12: {"name": "Qlik Sense", "aliases": [r"\bQlik Sense\b", r"\bQlikSense\b"]},
    13: {"name": "QlikView", "aliases": [r"\bQlikView\b"]},
    14: {"name": "Looker", "aliases": [r"\bLooker\b"]},
    15: {"name": "Looker Studio", "aliases": [r"\bLooker Studio\b"]},
    16: {"name": "MicroStrategy", "aliases": [r"\bMicroStrategy\b"]},
    17: {"name": "SAP BusinessObjects", "aliases": [r"\bSAP BusinessObjects\b", r"\bBusinessObjects\b", r"\bSAP BO\b"]},
    18: {"name": "SSRS", "aliases": [r"\bSSRS\b"]},
    19: {"name": "SSAS", "aliases": [r"\bSSAS\b"]},
    20: {"name": "Power Query", "aliases": [r"\bPower Query\b", r"\bPowerQuery\b"]},
    21: {"name": "Power Pivot", "aliases": [r"\bPower Pivot\b", r"\bPowerPivot\b"]},
    22: {"name": "DAX", "aliases": [r"\bDAX\b"]},
    23: {"name": "MDX", "aliases": [r"\bMDX\b"]},
    24: {"name": "PostgreSQL", "aliases": [r"\bPostgreSQL\b", r"\bPostgres\b"]},
    25: {"name": "MySQL", "aliases": [r"\bMySQL\b"]},
    26: {"name": "Microsoft SQL Server", "aliases": [r"\bMicrosoft SQL Server\b", r"\bMS SQL Server\b", r"\bMSSQL\b", r"\bSQL Server\b"]},
    27: {"name": "Oracle Database", "aliases": [r"\bOracle Database\b", r"\bOracle DB\b", r"\bOracle\b"]},
    28: {"name": "Snowflake", "aliases": [r"\bSnowflake\b"]},
    29: {"name": "BigQuery", "aliases": [r"\bBigQuery\b", r"\bGoogle BigQuery\b"]},
    30: {"name": "Amazon Redshift", "aliases": [r"\bAmazon Redshift\b", r"\bRedshift\b"]},
    31: {"name": "Teradata", "aliases": [r"\bTeradata\b"]},
    32: {"name": "MongoDB", "aliases": [r"\bMongoDB\b", r"\bMongo\b"]},
    33: {"name": "Cassandra", "aliases": [r"\bCassandra\b", r"\bApache Cassandra\b"]},
    34: {"name": "DynamoDB", "aliases": [r"\bDynamoDB\b", r"\bAmazon DynamoDB\b"]},
    35: {"name": "Redis", "aliases": [r"\bRedis\b"]},
    36: {"name": "Elasticsearch", "aliases": [r"\bElasticsearch\b", r"\bElastic Search\b"]},
    37: {"name": "ClickHouse", "aliases": [r"\bClickHouse\b"]},
    38: {"name": "Azure SQL", "aliases": [r"\bAzure SQL\b"]},
    39: {"name": "Amazon RDS", "aliases": [r"\bAmazon RDS\b", r"\bAWS RDS\b", r"\bRDS\b"]},
    40: {"name": "Azure Cosmos DB", "aliases": [r"\bAzure Cosmos DB\b", r"\bCosmos DB\b", r"\bCosmosDB\b"]},
    41: {"name": "Databricks", "aliases": [r"\bDatabricks\b"]},
    42: {"name": "Azure Databricks", "aliases": [r"\bAzure Databricks\b"]},
    43: {"name": "Microsoft Fabric", "aliases": [r"\bMicrosoft Fabric\b", r"\bMS Fabric\b"]},
    44: {"name": "Apache Spark", "aliases": [r"\bApache Spark\b", r"\bSpark\b"]},
    45: {"name": "PySpark", "aliases": [r"\bPySpark\b"]},
    46: {"name": "Hadoop", "aliases": [r"\bHadoop\b", r"\bApache Hadoop\b"]},
    47: {"name": "Hive", "aliases": [r"\bApache Hive\b", r"\bHive\b"]},
    48: {"name": "Apache Kafka", "aliases": [r"\bApache Kafka\b", r"\bKafka\b"]},
    49: {"name": "Apache Flink", "aliases": [r"\bApache Flink\b", r"\bFlink\b"]},
    50: {"name": "Apache Beam", "aliases": [r"\bApache Beam\b", r"\bBeam\b"]},
    51: {"name": "Trino", "aliases": [r"\bTrino\b"]},
    52: {"name": "Presto", "aliases": [r"\bPresto\b", r"\bPrestoDB\b"]},
    53: {"name": "Apache Airflow", "aliases": [r"\bApache Airflow\b", r"\bAirflow\b"]},
    54: {"name": "dbt", "aliases": [r"\bdbt\b", r"\bdata build tool\b"]},
    55: {"name": "Azure Data Factory", "aliases": [r"\bAzure Data Factory\b", r"\bADF\b"]},
    56: {"name": "AWS Glue", "aliases": [r"\bAWS Glue\b", r"\bGlue\b"]},
    57: {"name": "Informatica", "aliases": [r"\bInformatica\b", r"\bPowerCenter\b"]},
    58: {"name": "Talend", "aliases": [r"\bTalend\b"]},
    59: {"name": "Fivetran", "aliases": [r"\bFivetran\b"]},
    60: {"name": "Matillion", "aliases": [r"\bMatillion\b"]},
    61: {"name": "Apache NiFi", "aliases": [r"\bApache NiFi\b", r"\bNiFi\b"]},
    62: {"name": "SSIS", "aliases": [r"\bSSIS\b"]},
    63: {"name": "Pentaho", "aliases": [r"\bPentaho\b", r"\bPDI\b"]},
    64: {"name": "Alteryx", "aliases": [r"\bAlteryx\b"]},
    65: {"name": "ETL", "aliases": [r"\bETL\b"]},
    66: {"name": "ELT", "aliases": [r"\bELT\b"]},
    67: {"name": "Data Pipeline", "aliases": [r"\bData Pipelines?\b"]},
    68: {"name": "Data Integration", "aliases": [r"\bData Integration\b"]},
    69: {"name": "Data Ingestion", "aliases": [r"\bData Ingestion\b"]},
    70: {"name": "Batch Processing", "aliases": [r"\bBatch Processing\b"]},
    71: {"name": "Stream Processing", "aliases": [r"\bStream Processing\b", r"\bStreaming\b"]},
    72: {"name": "Change Data Capture", "aliases": [r"\bChange Data Capture\b", r"\bCDC\b"]},
    73: {"name": "Data Warehouse", "aliases": [r"\bData Warehous(e|ing)\b", r"\bEDW\b", r"\bDWH\b"]},
    74: {"name": "Data Lake", "aliases": [r"\bData Lakes?\b"]},
    75: {"name": "Lakehouse", "aliases": [r"\bLakehouse\b", r"\bData Lakehouse\b"]},
    76: {"name": "Data Mart", "aliases": [r"\bData Marts?\b"]},
    77: {"name": "Data Modeling", "aliases": [r"\bData Model(ing|s)?\b"]},
    78: {"name": "Dimensional Modeling", "aliases": [r"\bDimensional Model(ing)?\b"]},
    79: {"name": "Star Schema", "aliases": [r"\bStar Schemas?\b"]},
    80: {"name": "Snowflake Schema", "aliases": [r"\bSnowflake Schemas?\b"]},
    81: {"name": "Data Quality", "aliases": [r"\bData Quality\b", r"\bDQ\b"]},
    82: {"name": "Data Governance", "aliases": [r"\bData Governance\b"]},
    83: {"name": "Data Lineage", "aliases": [r"\bData Lineage\b"]},
    84: {"name": "Metadata Management", "aliases": [r"\bMetadata Management\b"]},
    85: {"name": "Master Data Management", "aliases": [r"\bMaster Data Management\b", r"\bMDM\b"]},
    86: {"name": "Data Orchestration", "aliases": [r"\bData Orchestration\b"]},
    87: {"name": "Data Architecture", "aliases": [r"\bData Architecture\b"]},
    88: {"name": "Medallion Architecture", "aliases": [r"\bMedallion Architecture\b"]},
    89: {"name": "Data Mesh", "aliases": [r"\bData Mesh\b"]},
    90: {"name": "OLAP", "aliases": [r"\bOLAP\b"]},
    91: {"name": "OLTP", "aliases": [r"\bOLTP\b"]},
    92: {"name": "Delta Lake", "aliases": [r"\bDelta Lake\b"]},
    93: {"name": "Apache Iceberg", "aliases": [r"\bApache Iceberg\b", r"\bIceberg\b"]},
    94: {"name": "Apache Hudi", "aliases": [r"\bApache Hudi\b", r"\bHudi\b"]},
    95: {"name": "Microsoft Azure", "aliases": [r"\bMicrosoft Azure\b", r"\bAzure\b"]},
    96: {"name": "AWS", "aliases": [r"\bAWS\b", r"\bAmazon Web Services\b"]},
    97: {"name": "Google Cloud Platform", "aliases": [r"\bGoogle Cloud Platform\b", r"\bGCP\b", r"\bGoogle Cloud\b"]},
    98: {"name": "Azure Synapse Analytics", "aliases": [r"\bAzure Synapse Analytics\b", r"\bAzure Synapse\b", r"\bSynapse\b"]},
    99: {"name": "Azure Data Lake Storage", "aliases": [r"\bAzure Data Lake Storage\b", r"\bADLS\b"]},
    100: {"name": "Azure Blob Storage", "aliases": [r"\bAzure Blob Storage\b", r"\bAzure Blob\b"]},
    101: {"name": "Azure Functions", "aliases": [r"\bAzure Functions\b"]},
    102: {"name": "Amazon S3", "aliases": [r"\bAmazon S3\b", r"\bAWS S3\b", r"\bS3\b"]},
    103: {"name": "Amazon Athena", "aliases": [r"\bAmazon Athena\b", r"\bAthena\b"]},
    104: {"name": "Amazon EMR", "aliases": [r"\bAmazon EMR\b", r"\bAWS EMR\b", r"\bEMR\b"]},
    105: {"name": "AWS Lambda", "aliases": [r"\bAWS Lambda\b", r"\bLambda\b"]},
    106: {"name": "Amazon Kinesis", "aliases": [r"\bAmazon Kinesis\b", r"\bKinesis\b"]},
    107: {"name": "Google Cloud Storage", "aliases": [r"\bGoogle Cloud Storage\b", r"\bGCS\b"]},
    108: {"name": "Google Dataflow", "aliases": [r"\bGoogle Dataflow\b", r"\bDataflow\b"]},
    109: {"name": "Google Dataproc", "aliases": [r"\bGoogle Dataproc\b", r"\bDataproc\b"]},
    110: {"name": "Google Pub/Sub", "aliases": [r"\bGoogle Pub/Sub\b", r"\bPub/Sub\b", r"\bCloud Pub/Sub\b"]},
    111: {"name": "Cloud Composer", "aliases": [r"\bCloud Composer\b"]},
    112: {"name": "Statistics", "aliases": [r"\bStatistics\b", r"\bStatistical Analysis\b"]},
    113: {"name": "Regression", "aliases": [r"\bRegression\b"]},
    114: {"name": "Time Series Analysis", "aliases": [r"\bTime Series\b"]},
    115: {"name": "Forecasting", "aliases": [r"\bForecasting\b"]},
    116: {"name": "Hypothesis Testing", "aliases": [r"\bHypothesis Testing\b"]},
    117: {"name": "A/B Testing", "aliases": [r"\bA/B Testing\b", r"\bAB Testing\b"]},
    118: {"name": "Experimental Design", "aliases": [r"\bExperimental Design\b", r"\bDesign of Experiments\b", r"\bDOE\b"]},
    119: {"name": "Predictive Modeling", "aliases": [r"\bPredictive Model(ing)?\b"]},
    120: {"name": "Data Mining", "aliases": [r"\bData Mining\b"]},
    121: {"name": "pandas", "aliases": [r"\bpandas\b"]},
    122: {"name": "NumPy", "aliases": [r"\bNumPy\b"]},
    123: {"name": "matplotlib", "aliases": [r"\bmatplotlib\b"]},
    124: {"name": "seaborn", "aliases": [r"\bseaborn\b"]},
    125: {"name": "Plotly", "aliases": [r"\bPlotly\b"]},
    126: {"name": "scikit-learn", "aliases": [r"\bscikit-learn\b", r"\bsklearn\b"]},
    127: {"name": "Jupyter", "aliases": [r"\bJupyter\b", r"\bJupyter Notebooks?\b"]},
    128: {"name": "Git", "aliases": [r"\bGit\b"]},
    129: {"name": "GitHub", "aliases": [r"\bGitHub\b"]},
    130: {"name": "GitLab", "aliases": [r"\bGitLab\b"]},
    131: {"name": "Bitbucket", "aliases": [r"\bBitbucket\b"]},
    132: {"name": "Docker", "aliases": [r"\bDocker\b"]},
    133: {"name": "Kubernetes", "aliases": [r"\bKubernetes\b", r"\bK8s\b"]},
    134: {"name": "Terraform", "aliases": [r"\bTerraform\b"]},
    135: {"name": "Jenkins", "aliases": [r"\bJenkins\b"]},
    136: {"name": "GitHub Actions", "aliases": [r"\bGitHub Actions\b"]},
    137: {"name": "GitLab CI/CD", "aliases": [r"\bGitLab CI/CD\b", r"\bGitLab CI\b"]},
    138: {"name": "Azure DevOps", "aliases": [r"\bAzure DevOps\b", r"\bADO\b"]},
    139: {"name": "CI/CD", "aliases": [r"\bCI/CD\b", r"\bCI-CD\b"]},
    140: {"name": "REST API", "aliases": [r"\bREST APIs?\b", r"\bRESTful APIs?\b"]},
    141: {"name": "API Integration", "aliases": [r"\bAPI Integrations?\b"]},
    142: {"name": "JSON", "aliases": [r"\bJSON\b"]},
    143: {"name": "XML", "aliases": [r"\bXML\b"]},
    144: {"name": "CSV", "aliases": [r"\bCSV\b"]},
    145: {"name": "Parquet", "aliases": [r"\bParquet\b", r"\bApache Parquet\b"]},
    146: {"name": "Avro", "aliases": [r"\bAvro\b", r"\bApache Avro\b"]},
    147: {"name": "ORC", "aliases": [r"\bORC\b", r"\bApache ORC\b", r"\bOptimized Row Columnar\b"]},
    148: {"name": "YAML", "aliases": [r"\bYAML\b"]},
    149: {"name": "C#", "aliases": [r"\bC#\b", r"\bC Sharp\b"]},
    150: {"name": "C++", "aliases": [r"\bC\+\+\b"]},
    151: {"name": "JavaScript", "aliases": [r"\bJavaScript\b", r"\bJS\b"]},
    152: {"name": "TypeScript", "aliases": [r"\bTypeScript\b", r"\bTS\b"]},
    153: {"name": "Go", "aliases": [r"\bGolang\b", r"\bGo\b"]},
    154: {"name": "SPSS", "aliases": [r"\bSPSS\b"]},
    155: {"name": "Rust", "aliases": [r"\bRust\b"]},
    156: {"name": "Machine Learning", "aliases": [r"\bMachine Learning\b", r"\bML\b"]},
    157: {"name": "PyTorch", "aliases": [r"\bPyTorch\b"]},
    158: {"name": "TensorFlow", "aliases": [r"\bTensorFlow\b"]},
    159: {"name": "Salesforce", "aliases": [r"\bSalesforce\b", r"\bSFDC\b"]},
    160: {"name": "CRM", "aliases": [r"\bCRM\b"]},
    161: {"name": "React", "aliases": [r"\bReact\b", r"\bReactJS\b", r"\bReact\.js\b"]},
    162: {"name": "Node.js", "aliases": [r"\bNode\.js\b", r"\bNodeJS\b", r"\bNode\b"]},
    163: {"name": "Kotlin", "aliases": [r"\bKotlin\b"]},
    164: {"name": "PowerPoint", "aliases": [r"\bPowerPoint\b", r"\bMS PowerPoint\b", r"\bPPT\b"]},
    165: {"name": "MS Office", "aliases": [r"\bMS Office\b", r"\bMicrosoft Office\b"]},
    166: {"name": "Piwik", "aliases": [r"\bPiwik\b", r"\bPiwik PRO\b", r"\bMatomo\b"]},
    167: {"name": "UI Design", "aliases": [r"\bUI Design\b", r"\bUser Interface Design\b", r"\bUI\b"]},
    168: {"name": "UX Design", "aliases": [r"\bUX Design\b", r"\bUser Experience Design\b", r"\bUX\b"]},
    169: {"name": "Agile", "aliases": [r"\bAgile\b"]},
    170: {"name": "Scrum", "aliases": [r"\bScrum\b"]},
    171: {"name": "DevOps", "aliases": [r"\bDevOps\b"]},
    172: {"name": "Twilio", "aliases": [r"\bTwilio\b"]},
    173: {"name": "Shell Scripting", "aliases": [r"\bShell Scripting\b", r"\bShell Script\b"]}
}

# Fix 1: Sort by longest actual plain string length in aliases, not raw regex length
def get_max_alias_len(skill_entry):
    lengths = []
    for alias in skill_entry["aliases"]:
        # Strip regex control characters to measure actual word length
        clean_alias = re.sub(r'\\[bB]|\(\?:|\)|\?|\[.*?\]', '', alias)
        lengths.append(len(clean_alias))
    return max(lengths) if lengths else 0

SEARCH_ORDER = sorted(SKILL_TAXONOMY.keys(), key=lambda k: get_max_alias_len(SKILL_TAXONOMY[k]), reverse=True)


# ==========================================
# 2. DISAMBIGUATION & CONTEXT FILTER RULES
# ==========================================

def validate_r_match(text, start, end):
    """Ensure R isn't a single capitalized letter or part of regular prose."""
    matched_str = text[start:end]
    
    # Strictly reject lowercase 'r'
    if matched_str != 'R':
        return matched_str.lower() in ['rstudio', 'r studio', 'r-language', 'r-programming']
    
    # Word boundary check for character context
    left_char = text[start - 1] if start > 0 else " "
    right_char = text[end] if end < len(text) else " "
    
    if re.match(r'[A-Za-z0-9]', left_char) or re.match(r'[A-Za-z0-9#+]', right_char):
        return False

    # Require explicit technical context around standalone 'R'
    context_window = text[max(0, start - 30):min(len(text), end + 30)]
    technical_triggers = [
        r'\bPython\b', r'\bSQL\b', r'\bRStudio\b', r'\btidyverse\b', 
        r'\bggplot\b', r'\bprogramming language\b', r'\bdata science\b'
    ]
    return any(re.search(trig, context_window, re.IGNORECASE) for trig in technical_triggers)

def validate_excel_match(text, start, end):
    """Ensure Excel isn't used as an ordinary verb."""
    matched_str = text[start:end]
    if matched_str.lower() in ["microsoft excel", "ms excel"]:
        return True
    
    verb_patterns = [
        r'\bexcel\s+(at|in|as|within|under|through)\b',
        r'\b(to|will|can|who|help\s+you)\s+excel\b'
    ]
    sentence_window = text[max(0, start-30):min(len(text), end+30)]
    for vp in verb_patterns:
        if re.search(vp, sentence_window, re.IGNORECASE):
            return False
    return True

def validate_ambiguous_term(skill_id, text, start, end):
    """Disambiguate terms like Spark, Beam, Hive, ORC, Go, dbt, Scala."""
    matched_str = text[start:end]
    
    # Scala: prevent matching 'escalable' stem or plain Spanish words
    if skill_id == 4 and matched_str.lower() == 'scala':
        window = text[max(0, start-10):min(len(text), end+10)]
        if re.search(r'escalable', window, re.IGNORECASE):
            return False

    # Apache items (Spark, Beam, Hive)
    if skill_id in [44, 47, 50] and matched_str.lower() in ['spark', 'beam', 'hive']:
        window = text[max(0, start-25):min(len(text), end+25)]
        tech_words = [r'apache', r'data', r'pipeline', r'engine', r'query', r'cluster', r'big\s*data', r'sql', r'pyspark']
        return any(re.search(tw, window, re.IGNORECASE) for tw in tech_words)
    
    # ORC
    if skill_id == 147 and matched_str.lower() == 'orc':
        window = text[max(0, start-20):min(len(text), end+20)]
        return any(re.search(tw, window, re.IGNORECASE) for tw in [r'parquet', r'avro', r'file', r'format', r'columnar', r'hadoop', r'hive'])
    
    # Go (Golang)
    if skill_id == 153 and matched_str.lower() == 'go':
        window = text[max(0, start-20):min(len(text), end+20)]
        return any(re.search(tw, window, re.IGNORECASE) for tw in [r'golang', r'developer', r'language', r'python', r'java', r'backend', r'code'])

    # dbt: Must be exact case or explicitly surrounded by technical context
    if skill_id == 54 and matched_str == 'dbt':
        window = text[max(0, start-20):min(len(text), end+20)]
        return any(re.search(tw, window, re.IGNORECASE) for tw in [r'sql', r'data', r'transform', r'warehouse', r'analytics', r'snowflake', r'bigquery'])

    return True


# ==========================================
# 3. CORE PROCESSING FUNCTION
# ==========================================

def extract_hard_skills(job_desc, existing_skills_str):
    if not isinstance(job_desc, str) or not job_desc.strip():
        return "", ""
    
    existing_skills_set = set()
    if isinstance(existing_skills_str, str) and existing_skills_str.strip():
        existing_skills_set = {s.strip().lower() for s in re.split(r'[,;/|]', existing_skills_str) if s.strip()}

    matched_ids = []
    matched_names = []
    text_spans_claimed = [] 

    for sid in SEARCH_ORDER:
        skill_info = SKILL_TAXONOMY[sid]
        s_name = skill_info["name"]
        
        if s_name.lower() in existing_skills_set:
            continue

        skill_detected = False
        for pattern in skill_info["aliases"]:
            # Case sensitivity rules per skill type
            flags = re.IGNORECASE
            if sid in [3, 54, 153]: # R, dbt, Go require exact case for short acronyms
                flags = 0

            for match in re.finditer(pattern, job_desc, flags):
                start, end = match.span()

                # Fix 2: Proper interval overlap algorithm max(start, cs) < min(end, ce)
                if any(max(start, cs) < min(end, ce) for cs, ce in text_spans_claimed):
                    continue
                
                # Rule checks
                if sid == 3 and not validate_r_match(job_desc, start, end):
                    continue
                if sid == 9 and not validate_excel_match(job_desc, start, end):
                    continue
                if sid in [4, 44, 47, 50, 54, 147, 153] and not validate_ambiguous_term(sid, job_desc, start, end):
                    continue

                # Valid detection
                skill_detected = True
                text_spans_claimed.append((start, end))
                break 

            if skill_detected:
                matched_ids.append(sid)
                matched_names.append(s_name)
                break

    sorted_matches = sorted(zip(matched_ids, matched_names), key=lambda x: x[0])
    
    id_str = ", ".join(str(m[0]) for m in sorted_matches)
    name_str = ", ".join(m[1] for m in sorted_matches)
    
    return id_str, name_str


def process_excel_file(file_path, output_path):
    """Processes all sheets in the Excel workbook."""
    excel_file = pd.ExcelFile(file_path)
    sheet_writer = pd.ExcelWriter(output_path, engine='openpyxl')

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        desc_col = next((c for c in df.columns if c.lower().strip() in ['job_description', 'job description']), None)
        skills_col = next((c for c in df.columns if c.lower().strip() in ['skills', 'existing_skills', 'existing skills']), None)

        if desc_col:
            ids_list = []
            names_list = []
            
            for idx, row in df.iterrows():
                desc_text = str(row[desc_col]) if pd.notna(row[desc_col]) else ""
                exist_text = str(row[skills_col]) if skills_col and pd.notna(row[skills_col]) else ""
                
                ids, names = extract_hard_skills(desc_text, exist_text)
                ids_list.append(ids)
                names_list.append(names)
            
            df['hard_skill_id'] = ids_list
            df['hard_skill'] = names_list
        else:
            print(f"Skipping Sheet '{sheet_name}': 'job_description' column not found.")
            
        df.to_excel(sheet_writer, sheet_name=sheet_name, index=False)

    sheet_writer.close()
    print(f"Processing complete! File saved as: {output_path}")

# ==========================================
# 4. EXECUTION
# ==========================================
if __name__ == "__main__":
    INPUT_FILE = "merged.xlsx"
    OUTPUT_FILE = "merged_postings_with_skills.xlsx"
    
    process_excel_file(INPUT_FILE, OUTPUT_FILE)