import re
import pandas as pd

# ==========================================
# 1. AI REGEX PATTERNS & KEYWORDS
# ==========================================

# Multi-language full phrases (case-insensitive safe)
EXPLICIT_AI_PHRASES = [
    r"\bartificial\s+intelligence\b",
    r"\binteligencia\s+artificial\b",      # Spanish
    r"\bintelligenza\s+artificiale\b",     # Italian
    r"\bmesterséges\s+intelligencia\b",     # Hungarian
    r"\bgenerative\s+ai\b",
    r"\bgenai\b",
    r"\bgen\s+ai\b",
    r"\bllms?\b",
    r"\blarge\s+language\s+models?\b",
    r"\bmachine\s+learning\b",
    r"\bdeep\s+learning\b",
    r"\bnatural\s+language\s+processing\b",
    r"\bnlp\b",
    r"\bcomputer\s+vision\b"
]

# AI Tools, Models, Frameworks & Platforms
AI_TOOLS_AND_MODELS = [
    r"\bchatgpt\b",
    r"\bgpt-?4o?\b",
    r"\bgpt-?3(?:\.5)?\b",
    r"\bopenai\b",
    r"\bclaude(?:\s+3|\s+3\.5)?\b",
    r"\banthropic\b",
    r"\bgemini\b",
    r"\bgoogle\s+bard\b",
    r"\bcopilot\b",
    r"\bgithub\s+copilot\b",
    r"\bmidjourney\b",
    r"\bstable\s+diffusion\b",
    r"\bdall-?e\b",
    r"\bhugging\s*face\b",
    r"\blangchain\b",
    r"\bllama(?:\s*[23])?\b",
    r"\bmistral\b",
    r"\bperplexity\b",
    r"\bbert\b",
    r"\bpytorch\b",
    r"\btensorflow\b",
    r"\bscikit-learn\b",
    r"\bkeras\b"
]

# Keywords near standalone "AI" / "IA" that confirm technical intent
TECH_CONTEXT_KEYWORDS = [
    r"data", r"model", r"tool", r"solution", r"algorithm", r"engineer", 
    r"developer", r"system", r"platform", r"tech", r"analytics", r"automation",
    r"prompt", r"agent", r"smart", r"bot", r"ml", r"cloud", r"software"
]


def count_ai_mentions_in_text(text: str) -> int:
    """Counts non-overlapping, validated AI mentions in a single text string."""
    if not isinstance(text, str) or not text.strip():
        return 0

    total_count = 0
    claimed_spans = []  # Stores (start, end) tuples to prevent double-counting

    def is_overlapping(start, end):
        return any(max(start, cs) < min(end, ce) for cs, ce in claimed_spans)

    # ---------------------------------------------------------
    # Step 1: Count explicit multi-word phrases & tools (Case-Insensitive)
    # ---------------------------------------------------------
    combined_phrases = EXPLICIT_AI_PHRASES + AI_TOOLS_AND_MODELS
    for pattern in combined_phrases:
        for match in re.finditer(pattern, text, re.IGNORECASE):
            start, end = match.span()
            if not is_overlapping(start, end):
                claimed_spans.append((start, end))
                total_count += 1

    # ---------------------------------------------------------
    # Step 2: Disambiguate standalone "AI" and "IA"
    # ---------------------------------------------------------
    short_acronym_pattern = r"\b(AI|IA)\b"

    for match in re.finditer(short_acronym_pattern, text):
        start, end = match.span()
        if is_overlapping(start, end):
            continue

        # Extract context window (40 characters left and right)
        left_window = text[max(0, start - 40):start]
        right_window = text[end:min(len(text), end + 40)]
        context_window = f"{left_window} {right_window}"

        # Check if surrounding text has tech keywords or slash formats like "AI/ML"
        has_tech_context = any(
            re.search(r"\b" + kw + r"\b", context_window, re.IGNORECASE) 
            for kw in TECH_CONTEXT_KEYWORDS
        )
        has_slash_pairing = bool(re.search(r"(AI/ML|ML/AI|IA/ML|ML/IA|GenAI/AI)", context_window, re.IGNORECASE))

        if has_tech_context or has_slash_pairing:
            claimed_spans.append((start, end))
            total_count += 1

    return total_count


def count_total_row_ai_mentions(desc_text: str, skills_text: str) -> int:
    """Combines counts from job_description and skills columns for a single row."""
    desc_count = count_ai_mentions_in_text(desc_text)
    skills_count = count_ai_mentions_in_text(skills_text)
    return desc_count + skills_count


# ==========================================
# 2. EXCEL PROCESSING FUNCTION
# ==========================================

def process_ai_mentions_excel(file_path: str, output_path: str):
    """Reads all sheets in an Excel file, counts AI mentions in description and skills, and writes output."""
    excel_file = pd.ExcelFile(file_path)
    sheet_writer = pd.ExcelWriter(output_path, engine='openpyxl')

    for sheet_name in excel_file.sheet_names:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        
        # Flexibly locate both columns
        desc_col = next((c for c in df.columns if c.lower().strip() in ['job_description', 'job description']), None)
        skills_col = next((c for c in df.columns if c.lower().strip() in ['skills', 'existing_skills', 'existing skills', 'hard_skill']), None)

        if desc_col or skills_col:
            ai_counts = []
            for idx, row in df.iterrows():
                desc_text = str(row[desc_col]) if desc_col and pd.notna(row[desc_col]) else ""
                skills_text = str(row[skills_col]) if skills_col and pd.notna(row[skills_col]) else ""
                
                total_row_count = count_total_row_ai_mentions(desc_text, skills_text)
                ai_counts.append(total_row_count)
            
            df['AI_flag'] = ai_counts
            print(f"Sheet '{sheet_name}': Processed {len(df)} rows. Total AI mentions found: {sum(ai_counts)}")
        else:
            print(f"Skipping Sheet '{sheet_name}': Neither 'job_description' nor 'skills' column found.")
            
        df.to_excel(sheet_writer, sheet_name=sheet_name, index=False)

    sheet_writer.close()
    print(f"\nProcessing complete! Saved result to: {output_path}")


# ==========================================
# 3. EXECUTION
# ==========================================
if __name__ == "__main__":
    INPUT_FILE = "job_postings_with_skills.xlsx"
    OUTPUT_FILE = "merged_postings_with_skills_and_AI.xlsx"
    
    process_ai_mentions_excel(INPUT_FILE, OUTPUT_FILE)