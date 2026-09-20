# LinkedIn Job Market Analytics: Rule-Based NLP Skill Extraction

**Measuring hard-skill prevalence in Data Analytics and Engineering job postings across Southern and Central Europe using Rule-Based NLP.**

---

## Project Overview

This project analyzes more than **1,000 Data Analyst and Data Engineer job postings** collected from LinkedIn across **Spain, Italy, and Hungary**.

The goal is to identify and compare the prevalence of hard skills in Data Analytics and Data Engineering job postings.

Using **Rule-Based Natural Language Processing (NLP)**, the pipeline extracts technical skills from unstructured job descriptions while handling aliases and contextual false positives. For example, `RStudio` can be mapped to `R`, while patterns such as _"excel in"_ and _"excel at"_ are filtered to avoid incorrect detections.

The processed data is modeled in **Power BI** to analyze skill demand across countries and job roles.

---

## Research Questions

- Which hard skills are most frequently requested?
- How does skill prevalence vary across Spain, Italy, and Hungary?
- Which programming languages, tools, and technologies appear most frequently in job postings?

---

## Pipeline

```text
LinkedIn Job Postings
        │
        ▼
Data Collection
        │
        ▼
Cleaning & Merging
        │
        ▼
Rule-Based NLP
        │
        ├── Skill Detection
        ├── Boundary Checks
        ├── Alias Resolution
        └── Contextual Filtering
        │
        ▼
Processed Dataset
        │
        ▼
Power BI Data Model
        │
        ▼
Dashboard
```

---

## Dashboard

**In Progress**

The Power BI data model is complete. The remaining step is to build the interactive dashboard for exploring skill demand across countries and job roles.

The dataset will also be periodically updated with newly scraped job postings to extend the analysis over time.

---

## Tech Stack

**Python** · **Pandas** · **Regex** · **Rule-Based NLP** · **Power BI**

---

## Project Status

- [x] Job-posting data scraped
- [x] Data cleaned
- [x] Data merged
- [x] Skill detection implemented
- [x] Alias resolution implemented
- [x] Contextual false-positive filtering implemented
- [x] Power BI data model completed
- [ ] Power BI dashboard
- [ ] Future data updates

---

## Legal & Data Usage Disclaimer

The data contained in this repository was collected from publicly accessible job postings for research, educational, and non-commercial purposes only.

---

## License

This project is intended for **non-commercial use**.
