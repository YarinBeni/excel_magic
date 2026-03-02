import pandas as pd
import sys
from pathlib import Path


def find_input_file():
    script_dir = Path(__file__).parent
    xlsx_files = [
        f for f in sorted(script_dir.glob("*.xlsx"))
        if "_מסומן" not in f.name
    ]

    if len(xlsx_files) == 0:
        print("שגיאה: לא נמצא קובץ Excel בתיקייה.")
        print("אנא הכנס קובץ .xlsx לתיקייה והרץ שנית.")
        sys.exit(1)
    elif len(xlsx_files) > 1:
        print("נמצאו מספר קבצי Excel בתיקייה:")
        for i, f in enumerate(xlsx_files, 1):
            print(f"  {i}. {f.name}")
        print("\nאנא השאר רק קובץ אחד בתיקייה והרץ שנית.")
        sys.exit(1)

    return xlsx_files[0]


def main():
    input_path = find_input_file()
    print(f"קורא קובץ: {input_path.name}")

    df = pd.read_excel(input_path)

    all_cols = df.columns.tolist()

    # עמודות לבדיקת כפילויות: כל העמודות חוץ מ-A (index 0), D (index 3), J (index 9)
    exclude_indices = {0, 3, 9}
    check_cols = [col for i, col in enumerate(all_cols) if i not in exclude_indices]

    # עמודה D (index 3) = זמן מערכת — לפיה קובעים מי ראשון בכל קבוצת כפילויות
    time_col = all_cols[3]

    print(f"עמודת זמן (קביעת ראשוניות): {time_col}")
    print(f"בודק כפילויות לפי עמודות: {check_cols}")

    # ממיינים לפי זמן מערכת בסדר עולה — הראשון הוא הישן ביותר
    df_sorted = df.sort_values(by=time_col, ascending=True)

    # keep='first' → השורה הראשונה (עם הזמן הנמוך ביותר) מקבלת False=0
    # כל השאר מקבלות True=1
    df_sorted["האם כפילות"] = df_sorted.duplicated(subset=check_cols, keep="first").astype(int)

    # מחזירים לסדר המקורי של הקובץ
    df_result = df_sorted.sort_index()

    output_name = input_path.stem + "_מסומן" + input_path.suffix
    output_path = input_path.parent / output_name

    df_result.to_excel(output_path, index=False)

    total = len(df_result)
    duplicates = int(df_result["האם כפילות"].sum())
    originals = total - duplicates

    print(f"\nסיום!")
    print(f"סה\"כ שורות:              {total}")
    print(f"שורות מקוריות (ללא כפילות): {originals}")
    print(f"שורות כפילות (מסומנות 1):   {duplicates}")
    print(f"\nהקובץ נשמר בשם: {output_name}")


if __name__ == "__main__":
    main()
