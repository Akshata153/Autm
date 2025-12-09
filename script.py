import os
import pandas as pd

# -----------------------------------------
# CONFIG
# -----------------------------------------
report_folder = "report"
bug_report_folder = "Bug report"
required_sheets = ["external", "internal", "extended", "reliablity"]
valid_status = ["DQA Verify", "Done"]
output_file = os.path.join(os.getcwd(), "output.txt")
# -----------------------------------------

def get_excel_file(folder):
    """Return the first Excel file found in the folder."""
    for file in os.listdir(folder):
        if file.endswith(".xlsx") or file.endswith(".xls"):
            return os.path.join(folder, file)
    raise FileNotFoundError(f"No Excel file found inside folder: {folder}")

def main():
    # Auto-detect Excel files
    report_excel = get_excel_file(report_folder)
    bug_excel = get_excel_file(bug_report_folder)

    print(f"Using report file    → {report_excel}")
    print(f"Using bug report file → {bug_excel}")

    # Load bug report Excel
    bug_df = pd.read_excel(bug_excel)
    bug_df["Issue key"] = bug_df["Issue key"].astype(str).str.strip()

    # Status lookup dictionary
    status_lookup = dict(zip(bug_df["Issue key"], bug_df["Status"]))

    # Output
    with open(output_file, "w") as out:
        for sheet in required_sheets:
            print(f"Processing sheet → {sheet}")
            try:
                df = pd.read_excel(report_excel, sheet_name=sheet)
            except:
                print(f"❌ Sheet '{sheet}' not found. Skipping.")
                continue

            if "Bug Ticket No." not in df.columns:
                print(f"❌ Missing 'Bug Ticket No.' column in sheet '{sheet}'. Skipping.")
                continue

            for _, row in df.iterrows():
                ticket = row["Bug Ticket No."]

                # Skip NA / empty
                if pd.isna(ticket):
                    continue

                ticket = str(ticket).strip()

                # Check in bug report
                if ticket in status_lookup:
                    status = status_lookup[ticket]
                    if status in valid_status:
                        out.write(f"{ticket} status is {status} in bug report. Please check in Jira.\n")

    print(f"\n✔ Done! Output created at: {output_file}")

if __name__ == "__main__":
    main()
