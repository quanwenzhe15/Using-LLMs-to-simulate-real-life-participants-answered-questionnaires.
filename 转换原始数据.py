import pandas as pd
import os

# ===================== 文本映射规则 =====================
code_mapping = {
    # 年龄
    "age": {98: "缺失值"},
    # 性别
    "sex": {1: "男", 2: "女", 8: "缺失值"},
    # 最高教育水平
    "education": {
        1: "NO SCHOOL/SOME GRADE SCHOOL",
        2: "EIGHTH GRADE/JUNIOR HIGH SCHOOL",
        3: "SOME HIGH SCHOOL",
        4: "GED",
        5: "GRADUATED FROM HIGH SCHOOL",
        6: "1 TO 2 YEARS OF COLLEGE, NO DEGREE YET",
        7: "3 OR MORE YEARS OF COLLEGE, NO DEGREE YET",
        8: "GRAD 2 YEAR COLLEGE OR VOC SCHOOL, OR ASSOCIATES DEGREE",
        9: "GRAD 4 OR 5 YEAR COLLEGE OR BACHELORS DEGREE",
        10: "SOME GRADUATE SCHOOL",
        11: "MASTERS DEGREE",
        12: "PH.D, ED.D, MD, DDS, LLB, LLD, JD, OR OTHER PROFESSIONAL DEGREE",
        97: "缺失值", 98: "缺失值", 99: "缺失值"
    },
    # 职业
    "occupation": {
        1: "EXECUTIVE, ADMINISTRATIVE, AND MANAGERIAL",
        2: "PROFESSIONAL SPECIALTY",
        3: "TECHNICIAN AND RELATED SUPPORT",
        4: "SALES OCCUPATION",
        5: "ADMINISTRATIVE SUPPORT INCLUDING CLERICAL",
        6: "SERVICE OCCUPATION",
        7: "FARMING, FORESTRY, AND FISHING",
        8: "PRECISION PRODUCTION, CRAFTS, AND REPAIR",
        9: "OPERATOR, LABORER, AND MILITARY",
        97: "缺失值", 98: "缺失值", 99: "缺失值"
    },
    # 行业
    "industry": {
        1: "AGRICULTURE, FORESTRY, FISHING, AND MINING",
        2: "CONSTRUCTION",
        3: "MANUFACTURING",
        4: "TRANSPORTATION, COMMUNICATIONS, AND PUBLIC UTILITY",
        5: "WHOLESALE TRADE",
        6: "RETAIL TRADE",
        7: "FINANCE, INSURANCE, AND REAL ESTATE",
        8: "BUSINESS AND REPAIR SERVICES",
        9: "PERSONAL SERVICES",
        10: "ENTERTAINMENT AND RECREATIONAL SERVICES",
        11: "PROFESSIONAL AND RELATED SERVICES",
        12: "PUBLIC ADMINISTRATION",
        97: "缺失值", 98: "缺失值", 99: "缺失值"
    },
    # 家庭年总收入
    "income": {-1: "缺失值", 999999: "缺失值"}
}

# ===================== 列名重命名规则 =====================
col_rename_map = {
    "age": "年龄",
    "sex": "性别",
    "education": "最高教育水平",
    "occupation": "职业",
    "industry": "行业",
    "income": "家庭年总收入"
}

# ===================== 路径配置=====================

original_excel_path = r"C:\Users\15896\Documents\xwechat_files\wxid_6f9qxgg753bj22_781a\msg\file\2026-02\MIDUS1人口统计学.xlsx"
output_dir = r"C:\Users\15896\Desktop\我的代码文件"
output_file_name = "模拟人变量以及相应水平.xlsx"

# ===================== 心处理逻辑 =====================
# 自动创建输出文件夹，不存在就新建，避免报错
os.makedirs(output_dir, exist_ok=True)
output_full_path = os.path.join(output_dir, output_file_name)

# 读取原始Excel文件
try:
    df = pd.read_excel(original_excel_path)
    total_rows = len(df)
    print(f"✅ 成功读取原始文件，共 {total_rows} 行数据")
except FileNotFoundError:
    print(f"❌ 错误：找不到原始文件，请检查路径是否正确 → {original_excel_path}")
    exit()
except Exception as e:
    print(f"❌ 读取文件失败，错误原因：{str(e)}")
    exit()

    # 【被试ID】
if "MIDUSID" in df.columns:
    df["MIDUSID"] = range(1, total_rows + 1)  # 强制从1开始连续编号
    df.rename(columns={"MIDUSID": "被试ID"}, inplace=True)
    print("✅ 已完成【MIDUSID】列处理，重命名为【被试ID】，编号从1开始")
else:
    # 无MIDUSID列时自动生成，放在表格第一列
    df.insert(0, "被试ID", range(1, total_rows + 1))
    print("⚠️  警告：原始文件未找到【MIDUSID】列，已自动生成【被试ID】列，编号从1开始")

# 编码转文本+列名重命名（直接替换原列，不保留原始数字）
for col_name, mapping_rule in code_mapping.items():
    if col_name not in df.columns:
        print(f"⚠️  警告：原始文件中未找到【{col_name}】列，已跳过该列处理")
        continue

    # 定义转换函数，兼容数字/字符串编码、空值、缺失值
    def convert_code(code):
        if pd.isna(code):
            return "缺失值"
        # 兼容单元格里字符串格式的数字
        try:
            code_int = int(code)
        except (ValueError, TypeError):
            return code
        # 匹配映射规则，匹配不到保留原始内容
        return mapping_rule.get(code_int, code)

    df[col_name] = df[col_name].apply(convert_code)
    # 重命名为指定中文列名
    df.rename(columns={col_name: col_rename_map[col_name]}, inplace=True)
    print(f"✅ 已完成【{col_name}】列转换，重命名为【{col_rename_map[col_name]}】")

# 调整列顺序
core_cols = ["被试ID", "年龄", "性别", "最高教育水平", "职业", "行业", "家庭年总收入"]
other_cols = [col for col in df.columns if col not in core_cols]
df = df[core_cols + other_cols]

# 保存文件
try:
    df.to_excel(output_full_path, index=False, engine="openpyxl")
    print(f"\n🎉 全部处理完成！文件已保存至：\n{output_full_path}")
except Exception as e:
    print(f"❌ 保存文件失败，错误原因：{str(e)}")
    exit()