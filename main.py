import pandas as pd
import numpy as np
import os

# 等级成绩转换成为百分制数值，可自行调整
grade_mapping = {'优秀': 90, '良好': 80, '中等': 70, '及格': 60, '不及格': 0}

def process_grades(file_path):
    # 读取 Excel 文件
    df = pd.read_excel(file_path)

    # 创建输出目录
    os.makedirs('out', exist_ok=True)

    df['数值成绩'] = df['总成绩'].apply(lambda x: grade_mapping[x] if x in grade_mapping else float(x))

    # 提取学年信息
    df['学年'] = df['学年学期'].str.extract(r'(\d{4}-\d{4})')

    # 计算总必修（不含军事和体育）和总 GPA
    # 筛选必修课程且课程名不包含"体育"或"军事"
    total_required = df[(df['课程性质'] == '必修') &
                        (~df['课程名'].str.contains('体育|军事', case=False, regex=True))]

    total_gpa = np.sum(df['学分'] * df['绩点']) / np.sum(df['学分'])
    total_required_gpa = np.sum(total_required['学分'] * total_required['绩点']) / np.sum(total_required['学分'])

    # 计算每个学年的必修（不含军事和体育）和总 GPA
    yearly_stats = []
    for year, group in df.groupby('学年'):
        # 筛选当前学年的必修课程（不含军事和体育）
        year_required = group[(group['课程性质'] == '必修') &
                              (~group['课程名'].str.contains('体育|军事', case=False, regex=True))]

        year_gpa = np.sum(group['学分'] * group['绩点']) / np.sum(group['学分'])
        year_required_gpa = np.sum(year_required['学分'] * year_required['绩点']) / np.sum(year_required['学分'])
        year_avg_score = np.sum(group['数值成绩'] * group['学分']) / np.sum(group['学分'])

        yearly_stats.append({
            '学年': year,
            '总GPA': year_gpa,
            '必修GPA（不含军事体育）': year_required_gpa,
            '平均分': year_avg_score
        })

    yearly_stats_df = pd.DataFrame(yearly_stats)

    # 输出结果到控制台
    print("总体统计")
    print(f"总GPA: {total_gpa:.4f}")
    print(f"必修课程GPA（不含军事体育）: {total_required_gpa:.4f}\n")

    print("分学年统计")
    print(yearly_stats_df.to_string(index=False))

    # 准备输出到 Excel 的数据
    output_data = {
        '总体统计': pd.DataFrame({
            '指标': ['总GPA', '必修课程GPA（不含军事体育）'],
            '值': [total_gpa, total_required_gpa]
        }),
        '分学年统计': yearly_stats_df
    }

    # 输出到 Excel 文件
    output_path = 'out/grade_analysis_results.xlsx'
    with pd.ExcelWriter(output_path) as writer:
        for sheet_name, data in output_data.items():
            data.to_excel(writer, sheet_name=sheet_name, index=False)

    print(f"\n分析结果已保存到: {output_path}")

    return {
        'total_gpa': total_gpa,
        'total_required_gpa': total_required_gpa,
        'yearly_stats': yearly_stats_df
    }

if __name__ == "__main__":
    # 导入成绩表格
    process_grades("in/example.xlsx")