import pandas as pd
import numpy as np

def process_grades(file_path):
    df = pd.read_excel(file_path)

    # 转换等级成绩为数值
    grade_mapping = {'优秀': 90, '良好': 80, '中等': 70, '及格': 60, '不及格': 0}
    df['数值成绩'] = df['总成绩'].apply(lambda x: grade_mapping[x] if x in grade_mapping else float(x))

    df['学年'] = df['学年学期'].str.extract(r'(\d{4}-\d{4})')

    # 计算总必修和总GPA
    total_required = df[df['课程性质'] == '必修']
    total_gpa = np.sum(df['学分'] * df['绩点']) / np.sum(df['学分'])
    total_required_gpa = np.sum(total_required['学分'] * total_required['绩点']) / np.sum(total_required['学分'])

    # 计算每个学年的必修和总GPA
    yearly_stats = []
    for year, group in df.groupby('学年'):
        year_required = group[group['课程性质'] == '必修']
        year_gpa = np.sum(group['学分'] * group['绩点']) / np.sum(group['学分'])
        year_required_gpa = np.sum(year_required['学分'] * year_required['绩点']) / np.sum(year_required['学分'])
        year_avg_score = np.sum(group['数值成绩'] * group['学分']) / np.sum(group['学分'])

        yearly_stats.append({
            '学年': year,
            '总GPA': year_gpa,
            '必修GPA': year_required_gpa,
            '平均分': year_avg_score
        })

    yearly_stats_df = pd.DataFrame(yearly_stats)

    # 输出结果
    print("总体统计")
    print(f"总GPA: {total_gpa:.4f}")
    print(f"必修课程GPA: {total_required_gpa:.4f}\n")

    print("分学年统计")
    print(yearly_stats_df)

    return {
        'total_gpa': total_gpa,
        'total_required_gpa': total_required_gpa,
        'yearly_stats': yearly_stats_df
    }

if __name__ == "__main__":
    process_grades("in/example.xlsx")