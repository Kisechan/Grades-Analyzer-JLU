import pandas as pd
import numpy as np
import os
import glob

# 等级成绩转换成为百分制数值，可自行调整
grade_mapping = {'优秀': 90, '良好': 80, '中等': 70, '及格': 60, '不及格': 0}

def analyze_single_file(file_path):
    """分析单个成绩文件"""
    # 读取 Excel 文件
    df = pd.read_excel(file_path)

    # 从文件名获取学生标识（去除路径和扩展名）
    student_id = os.path.splitext(os.path.basename(file_path))[0]

    df['数值成绩'] = df['总成绩'].apply(lambda x: grade_mapping[x] if x in grade_mapping else float(x))

    # 提取学年信息
    df['学年'] = df['学年学期'].str.extract(r'(\d{4}-\d{4})')

    # 计算总必修（不含军事和体育）和总 GPA
    total_required = df[(df['课程性质'] == '必修') &
                        (~df['课程名'].str.contains('体育|军事', case=False, regex=True))]

    total_gpa = np.sum(df['学分'] * df['绩点']) / np.sum(df['学分'])
    total_required_gpa = np.sum(total_required['学分'] * total_required['绩点']) / np.sum(total_required['学分'])

    # 计算每个学年的必修（不含军事和体育）和总 GPA
    yearly_stats = []
    for year, group in df.groupby('学年'):
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

    # 准备输出数据
    result = {
        'student_id': student_id,
        'total_gpa': total_gpa,
        'total_required_gpa': total_required_gpa,
        'yearly_stats': yearly_stats_df
    }

    return result

def save_results_to_excel(results, output_dir='out'):
    """将所有结果保存到Excel文件"""
    # 准备合并的数据
    summary_data = []
    yearly_data = []

    for result in results:
        student_id = result['student_id']

        # 添加总体统计
        summary_data.append({
            '学号/文件名': student_id,
            '总GPA': result['total_gpa'],
            '必修GPA（不含军事体育）': result['total_required_gpa']
        })

        # 添加学年统计
        yearly_stats = result['yearly_stats']
        yearly_stats['学号/文件名'] = student_id
        yearly_data.append(yearly_stats)

    # 创建DataFrame
    summary_df = pd.DataFrame(summary_data)
    yearly_df = pd.concat(yearly_data, ignore_index=True)

    # 重新排列列顺序
    yearly_df = yearly_df[['学号/文件名', '学年', '总GPA', '必修GPA（不含军事体育）', '平均分']]

    # 保存到Excel
    output_path = os.path.join(output_dir, 'combined_analysis_results.xlsx')
    with pd.ExcelWriter(output_path) as writer:
        summary_df.to_excel(writer, sheet_name='总体统计', index=False)
        yearly_df.to_excel(writer, sheet_name='分学年统计', index=False)

    return output_path

def process_all_grades(input_dir='in', output_dir='out'):
    """处理输入目录中的所有Excel文件"""
    # 创建输出目录
    os.makedirs(output_dir, exist_ok=True)

    # 获取所有Excel文件
    excel_files = glob.glob(os.path.join(input_dir, '*.xlsx'))

    if not excel_files:
        print(f"在目录 {input_dir} 中未找到任何Excel文件")
        return

    print(f"找到 {len(excel_files)} 个Excel文件，开始分析...\n")

    all_results = []
    for file_path in excel_files:
        try:
            print(f"正在分析文件: {os.path.basename(file_path)}")
            result = analyze_single_file(file_path)
            all_results.append(result)

            # 打印简要结果
            print(f"  总GPA: {result['total_gpa']:.4f}")
            print(f"  必修GPA（不含军事体育）: {result['total_required_gpa']:.4f}")

            # 保存单个文件结果
            student_output_path = os.path.join(output_dir, f"{result['student_id']}_analysis.xlsx")
            with pd.ExcelWriter(student_output_path) as writer:
                pd.DataFrame({
                    '指标': ['总GPA', '必修GPA（不含军事体育）'],
                    '值': [result['total_gpa'], result['total_required_gpa']]
                }).to_excel(writer, sheet_name='总体统计', index=False)
                result['yearly_stats'].to_excel(writer, sheet_name='分学年统计', index=False)

            print(f"  结果已保存到: {student_output_path}\n")

        except Exception as e:
            print(f"处理文件 {os.path.basename(file_path)} 时出错: {str(e)}\n")

    # 保存合并结果
    if all_results:
        combined_path = save_results_to_excel(all_results, output_dir)
        print(f"\n所有分析完成！合并结果已保存到: {combined_path}")

if __name__ == "__main__":
    # 处理 in 目录中的所有 Excel 文件
    process_all_grades()