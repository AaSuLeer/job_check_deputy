import os
import pandas as pd
import openpyxl

def load_roster(file_path):
    if file_path.endswith(".csv"):
        df = pd.read_csv(file_path, header=None)
    elif file_path.endswith(".xlsx"):
        df = pd.read_excel(file_path, header=None)
    else:
        raise ValueError("Unsupported file format")

    names = df.iloc[:, 0].dropna().astype(str).tolist()
    names = [name.strip() for name in names if name.strip() != ""]
    return names


def scan_submissions(folder_path):
    return [
        f for f in os.listdir(folder_path)
        if os.path.isfile(os.path.join(folder_path, f))
    ]


def find_missing_students(names, files):
    submitted = set()

    for name in names:
        for f in files:
            if name in f:
                submitted.add(name)
                break

    missing = [name for name in names if name not in submitted]
    return submitted, missing


def save_to_excel(submitted, missing, output_file):
    """
    覆写写入Excel（等价于 w 模式）
    """
    df_sub = pd.DataFrame({"已提交": list(submitted)})
    df_miss = pd.DataFrame({"未提交": list(missing)})

    # 覆写写入
    with pd.ExcelWriter(output_file, engine="openpyxl", mode="w") as writer:
        df_sub.to_excel(writer, sheet_name="已提交", index=False)
        df_miss.to_excel(writer, sheet_name="未提交", index=False)


def main():
    roster_file = "0923201班班级成员表.csv"
    target_folder = "C:/Users/blacksheep/Desktop/0923201班级亚信安全企业考察报告"
    output_file = "result.xlsx"

    names = load_roster(roster_file)
    files = scan_submissions(target_folder)
    submitted, missing = find_missing_students(names, files)

    save_to_excel(submitted, missing, output_file)

    print("结果已写入:", output_file)
    print(f"总人数: {len(names)}")
    print(f"已提交: {len(submitted)}")
    print(f"未提交: {len(missing)}")


if __name__ == "__main__":
    main()
