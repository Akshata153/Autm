import os
import pandas as pd

# -----------------------------------------
# CONFIG
# -----------------------------------------
report_folder = "Creport"
Subject_folder = "Sreport"
required_sheets = ["ClassA", "ClassB", "ClassC", "ClassD"]
valid_status = ["Fail", "Pending"]
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
    Sreport_excel = get_excel_file(Sreport_folder)
    Creport_excel = get_excel_file(Creport_folder)

    print(f"Using report file for classwise report    → {Creport_excel}")
    print(f"Using report file for subject report → {Sreport_excel}")

    # Load report Excel
    ddf = pd.read_excel(bug_excel)
    ddf["studentid"] = ddf["studentid"].astype(str).str.strip()

    # Status lookup dictionary
    status_lookup = dict(zip(bug_df["studentid"], bug_df["Status"]))

    # Output
    with open(output_file, "w") as out:
        for sheet in required_sheets:
            print(f"Processing sheet → {sheet}")
            try:
                df = pd.read_excel(report_excel, sheet_name=sheet)
            except:
                print(f"❌ Sheet '{sheet}' not found. Skipping.")
                continue

            if "Status" not in df.columns:
                print(f"❌ Missing 'Status' column in sheet '{sheet}'. Skipping.")
                continue

            for _, row in df.iterrows():
                Status = row["Status"]

                # Skip NA / empty
                if pd.isna(Status):
                    continue

                Status = str(Status).strip()

                if Status in status_lookup:
                    status = status_lookup[ticket]
                    if status in valid_status:
                        out.write(f"{Status} status is {status} in subject report. Please checkwait for results.\n")

    print(f"\n✔ Done! Output created at: {output_file}")

if __name__ == "__main__":
    main()
