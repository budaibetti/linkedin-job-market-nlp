from pathlib import Path
import pandas as pd
import re


# ============================================================
# SETTINGS
# ============================================================

INPUT_FOLDER = Path("clean data")
OUTPUT_FOLDER = Path("merged file")
OUTPUT_FILE = OUTPUT_FOLDER / "merged.xlsx"

# Matches:
# skills/0
# skills/1
# skills/2
# skills/10
# etc.
SKILL_PATTERN = re.compile(r"^skills/\d+$", re.IGNORECASE)


# ============================================================
# CREATE OUTPUT FOLDER
# ============================================================

OUTPUT_FOLDER.mkdir(exist_ok=True)


# ============================================================
# FIND EXCEL FILES
# ============================================================

excel_files = list(INPUT_FOLDER.glob("*.xlsx"))

if not excel_files:
    raise FileNotFoundError(
        f"No Excel files found in '{INPUT_FOLDER}'."
    )

print(f"Found {len(excel_files)} Excel files.")


# ============================================================
# MERGE FILES
# ============================================================

with pd.ExcelWriter(
    OUTPUT_FILE,
    engine="openpyxl"
) as writer:

    for file in excel_files:

        print()
        print("=" * 60)
        print(f"Processing: {file.name}")

        # ----------------------------------------------------
        # READ WITHOUT HEADERS
        # ----------------------------------------------------
        #
        # This is important because skills/0, skills/1, etc.
        # are in the FIRST ROW, not actual Excel headers.
        #
        df = pd.read_excel(
            file,
            header=None
        )

        print(f"Rows: {len(df)}")
        print(f"Columns: {len(df.columns)}")

        # ----------------------------------------------------
        # FIND THE ROW CONTAINING skills/0, skills/1, etc.
        # ----------------------------------------------------

        skill_columns = []

        skill_row_index = None

        for row_index in range(len(df)):

            row = df.iloc[row_index]

            found_columns = []

            for column_index, value in enumerate(row):

                if pd.isna(value):
                    continue

                value = str(value).strip()

                if SKILL_PATTERN.match(value):
                    found_columns.append(column_index)

            if found_columns:
                skill_columns = found_columns
                skill_row_index = row_index
                break

        # ----------------------------------------------------
        # CHECK WHETHER SKILL COLUMNS WERE FOUND
        # ----------------------------------------------------

        if not skill_columns:

            print("WARNING: No skills/0, skills/1, etc. found.")

            # Still save the file
            sheet_name = file.stem[:31]
            sheet_name = re.sub(
                r"[\[\]:*?/\\]",
                "_",
                sheet_name
            )

            df.to_excel(
                writer,
                sheet_name=sheet_name,
                index=False,
                header=False
            )

            continue

        print(
            f"Found skill columns in row {skill_row_index + 1}: "
            f"{skill_columns}"
        )

        # ----------------------------------------------------
        # REMOVE THE SKILL HEADER ROW
        # ----------------------------------------------------
        #
        # The row containing:
        # skills/0 | skills/1 | skills/2
        #
        # is not actual data, so remove it.
        #

        df = df.drop(
            index=skill_row_index
        ).reset_index(drop=True)

        # ----------------------------------------------------
        # MERGE SKILLS FOR EACH ROW
        # ----------------------------------------------------

        merged_skills = []

        for _, row in df.iterrows():

            skills = []

            for column_index in skill_columns:

                # Make sure this column exists in this file
                if column_index >= len(row):
                    continue

                value = row.iloc[column_index]

                # Ignore empty cells
                if pd.isna(value):
                    continue

                value = str(value).strip()

                if not value:
                    continue

                # --------------------------------------------
                # REMOVE DUPLICATE SKILLS
                # --------------------------------------------
                #
                # Case-insensitive:
                # Python
                # python
                #
                # will be treated as the same skill.
                #

                already_exists = any(
                    skill.lower() == value.lower()
                    for skill in skills
                )

                if not already_exists:
                    skills.append(value)

            merged_skills.append(
                ", ".join(skills)
            )

        # ----------------------------------------------------
        # ADD COMBINED SKILLS COLUMN
        # ----------------------------------------------------

        df["skills"] = merged_skills

        # ----------------------------------------------------
        # REMOVE ORIGINAL SKILL COLUMNS
        # ----------------------------------------------------

        df.drop(
            columns=skill_columns,
            inplace=True
        )

        # ----------------------------------------------------
        # REMOVE DUPLICATE ROWS
        # ----------------------------------------------------

        before = len(df)

        df.drop_duplicates(
            inplace=True
        )

        after = len(df)

        print(
            f"Removed {before - after} duplicate rows."
        )

        # ----------------------------------------------------
        # CREATE COLUMN NAMES
        # ----------------------------------------------------
        #
        # Since the original file doesn't have proper headers,
        # we'll create generic column names.
        #
        # If your first row contains actual column names too,
        # see the note below.
        #

        df.columns = [
            f"column_{i + 1}"
            for i in range(len(df.columns))
        ]

        # Rename the final column to "skills"
        df.rename(
            columns={
                df.columns[-1]: "skills"
            },
            inplace=True
        )

        # ----------------------------------------------------
        # CREATE SHEET NAME
        # ----------------------------------------------------

        sheet_name = file.stem[:31]

        # Remove characters Excel doesn't allow
        sheet_name = re.sub(
            r"[\[\]:*?/\\]",
            "_",
            sheet_name
        )

        # ----------------------------------------------------
        # SAVE TAB
        # ----------------------------------------------------

        df.to_excel(
            writer,
            sheet_name=sheet_name,
            index=False
        )

        print(
            f"Created tab: {sheet_name}"
        )


# ============================================================
# FINISHED
# ============================================================

print()
print("=" * 60)
print("DONE!")
print("=" * 60)
print()
print(f"Output file:")
print(OUTPUT_FILE.resolve())