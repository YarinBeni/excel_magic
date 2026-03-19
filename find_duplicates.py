import sys
import pandas as pd
from pathlib import Path


def find_input_file():
    script_dir = Path(__file__).parent
    xlsx_files = [
        f for f in sorted(script_dir.glob("*.xlsx"))
        if "_מסומן" not in f.name
    ]

    if len(xlsx_files) == 0:
        print("Error: No Excel file found in the folder.")
        print("Please place a .xlsx file in the folder and try again.")
        sys.exit(1)
    elif len(xlsx_files) > 1:
        print("Multiple Excel files found in the folder:")
        for i, f in enumerate(xlsx_files, 1):
            print(f"  {i}. {f.name}")
        print("\nPlease keep only one file in the folder and try again.")
        sys.exit(1)

    return xlsx_files[0]


def main():
    input_path = find_input_file()
    print(f"Reading file: {input_path.name}")

    df = pd.read_excel(input_path)

    if df.empty:
        print("Error: The file is empty.")
        sys.exit(1)

    all_cols = df.columns.tolist()

    if len(all_cols) < 10:
        print(f"Error: The file has only {len(all_cols)} columns, at least 10 are required.")
        sys.exit(1)

    # Columns to check for duplicates: all columns except A (index 0), D (index 3), J (index 9)
    exclude_indices = {0, 3, 9}
    check_cols = [col for i, col in enumerate(all_cols) if i not in exclude_indices]

    # Column D (index 3) = system timestamp — used to determine which row came first
    time_col = all_cols[3]

    print(f"Timestamp column: {time_col}")
    print(f"Checking duplicates by columns: {check_cols}")

    # Sort by timestamp ascending — the first row is the oldest
    df_sorted = df.sort_values(by=time_col, ascending=True)

    # Rank within each duplicate group (sorted by time ascending):
    # 1 = unique row (group size == 1)
    # 2 = first in a duplicate group, 3 = second, etc.
    group_sizes = df_sorted.groupby(check_cols)[time_col].transform("count")
    rank_in_group = df_sorted.groupby(check_cols).cumcount() + 1
    df_sorted["האם כפילות"] = rank_in_group + (group_sizes > 1).astype(int)

    # Restore original row order
    df_result = df_sorted.sort_index()

    output_name = input_path.stem + "_מסומן" + input_path.suffix
    output_path = input_path.parent / output_name

    try:
        df_result.to_excel(output_path, index=False)
    except PermissionError:
        print(f"\nError: Cannot save file {output_name}")
        print("The file may be open in Excel. Close it and try again.")
        sys.exit(1)

    total = len(df_result)
    unique_rows = int((df_result["האם כפילות"] == 1).sum())
    dup_rows = total - unique_rows

    print(f"\nDone!")
    print(f"Total rows:      {total}")
    print(f"Unique rows:     {unique_rows}")
    print(f"Duplicate rows:  {dup_rows}")
    print(f"\nOutput saved as: {output_name}")


if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(f"\nUnexpected error: {e}")
        sys.exit(1)
