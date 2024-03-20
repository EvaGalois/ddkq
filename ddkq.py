import pandas as pd
from pprint import pprint
from datetime import datetime

def calculate_time_differences(records):
    """计算并返回时间差（分钟）列表"""
    time_diffs = []  # 存储时间差的列表
    for i in range(1, len(records)):
        # 将时间字符串转换为datetime对象
        time1 = datetime.strptime(records[i-1], "%H:%M")
        time2 = datetime.strptime(records[i], "%H:%M")
        # 计算相邻时间的差值，并转换为分钟
        diff = (time2 - time1).total_seconds() / 60
        time_diffs.append(diff)
    return time_diffs

def process_employee_id(employee_id):
    return str(employee_id).zfill(7)

def remove_trailing_spaces(record_list):
    if isinstance(record_list, list):
        return [item.strip() if isinstance(item, str) else item for item in record_list]
    return record_list

def extract_attendance_data(dd, dd_indices):
    attendance_data = {}
    for index in dd_indices:
        date = dd.at[index, "日期"][0] if isinstance(dd.at[index, "日期"], list) else dd.at[index, "日期"]
        attendance_record = dd.at[index, "打卡记录"]
        if isinstance(attendance_record, list):
            attendance_record = remove_trailing_spaces(attendance_record)
        else:
            attendance_record = [attendance_record]
        attendance_data[date] = attendance_record if pd.notna(attendance_record[0]) else [0]
    return attendance_data

# 定义函数计算工时时长
def calculate_work_hours(records):
    if len(records) < 4:
        return 0.0, 0.0  # 长度小于4的记录视为工时和加班工时都为0

    last_record = records[-1]
    last_time = pd.to_datetime(last_record, format="%H:%M")

    # 计算加班工时
    overtime_hours = 0.0
    if last_time > pd.to_datetime("18:00", format="%H:%M"):
        overtime_hours = round((last_time - pd.to_datetime("17:30", format="%H:%M")).seconds / 3600, 2)

    # 计算基础工时
    first_time = pd.to_datetime(records[0], format="%H:%M")
    basic_duration = 8.5  # 基础工时为8.5小时
    if first_time >= pd.to_datetime("07:00", format="%H:%M") and first_time < pd.to_datetime("08:00", format="%H:%M"):
        return basic_duration, overtime_hours

    # 处理迟到情况
    if first_time > pd.to_datetime("08:00", format="%H:%M"):
        late_duration = round((first_time - pd.to_datetime("08:00", format="%H:%M")).seconds / 3600, 2)
        basic_duration -= late_duration

    return basic_duration, overtime_hours

def calculate_daily_work_hours(unique_records, work_length):
    daily_work_hours = []  # 初始化每日工作时长列表
    
    # 计算上午工作时长
    morning_work_hours = 0
    # 检查是否存在工作时长记录且第一次打卡时间在8:00之前
    if work_length and unique_records[0] != '00:00':
        first_record_time = datetime.strptime(unique_records[0], "%H:%M")
        if first_record_time < datetime.strptime("08:00", "%H:%M"):
            if work_length[0] > 240:
                morning_work_hours = 4
            else:
                morning_work_hours = 3
        else:  # 如果第一次打卡时间不在8:00之前
            morning_work_hours = 3  # 可能需要根据具体逻辑调整
    daily_work_hours.append(morning_work_hours)
    
    # 计算下午工作时长
    afternoon_work_hours = 0
    if unique_records != [0]:
        if len(work_length) > 1 and work_length[1] > 270:
            afternoon_work_hours = 4.5
        else:
            afternoon_work_hours = 4.0
    daily_work_hours.append(afternoon_work_hours)
    
    # 计算加班工作时长
    overtime_work_hours = 0
    if len(work_length) > 2:
        overtime_work_hours = int(work_length[2] / 30) / 2
    daily_work_hours.append(overtime_work_hours)
    
    return daily_work_hours


dd_file = "table/_DY日统计_钉钉.xlsx"
# dd_file = "table/11月钉钉(1)(1)-deng.xlsx"
dd_data = pd.read_excel(dd_file, sheet_name="每日统计")
# dd_data = pd.read_excel(dd_file, sheet_name="Sheet1")
dd = dd_data[["姓名", "工号", "日期", "班次", "打卡记录"]].copy()
dd["日期"] = dd["日期"].str.split(" ")
dd["班次"] = dd["班次"].str.split(" ")
dd["打卡记录"] = dd["打卡记录"].str.split("  \n").apply(remove_trailing_spaces)
dd["工号"] = dd["工号"].apply(process_employee_id)

kq_file = "table/12月考勤 -车间.xlsx"
# kq_file = "table/11月车间考勤(1)(1)-deng.xlsx"
kq_data = pd.read_excel(kq_file, sheet_name="工时汇总")
# kq_data = pd.read_excel(kq_file, sheet_name="Sheet1")
# print(kq_data.columns)

kq = kq_data[["姓名", "工号", "总计工时", "总计天数"]].copy()
kq["工号"] = kq["工号"].apply(process_employee_id)
kq["dd索引"] = [[] for _ in range(len(kq))]
kq["考勤数据"] = [{} for _ in range(len(kq))]
names_without_attendance = []

for index, row in kq.iterrows():
    employee_id = row["工号"]
    dd_indices = dd[dd["工号"] == employee_id].index.tolist()
    kq.at[index, "dd索引"] = dd_indices
    if not dd_indices:
        names_without_attendance.append(row["姓名"])
print(names_without_attendance)
exit()
for index, row in kq.iterrows():
    dd_indices = row["dd索引"]
    attendance_data = extract_attendance_data(dd, dd_indices)
    kq.at[index, "考勤数据"] = attendance_data

print(kq)
exit()
# kq["考勤数据"].to_csv("kq_考勤数据.csv", index=False)

# print("每日统计表格的数据：")
# print(dd)

# print("工时汇总表格的数据：")
# print(kq)

for index, attendance_data in enumerate(kq["考勤数据"]):
    print(f"考勤数据 {index + 1}:")
    for date, records in attendance_data.items():
        print(f"日期: {date}")

        # 处理重复打卡问题
        unique_records = [records[0]]  # 保留第一次打卡记录
        for i in range(1, len(records)):
            current_time = pd.to_datetime(records[i], format="%H:%M")
            prev_time = pd.to_datetime(unique_records[-1], format="%H:%M")
            time_diff = (current_time - prev_time).seconds / 60  # 计算时间差（分钟）
            if time_diff > 3:
                unique_records.append(records[i])

        # print(f"去除重复打卡后的记录: {len(unique_records)} {unique_records}")

        interval = calculate_time_differences(unique_records)
        # print(f"时间间隙：{interval}")
        work_length = []
        for i in interval:
            if i < 30:
                pass
            else:
                work_length.append(i)
        # print(f"工作时长：{work_length}")

        daily_work_hours = calculate_daily_work_hours(unique_records, work_length)
        # print(f"每日工作时长: {daily_work_hours}")


        # # 在这里计算基础时长Basic_duration和加班时长Overtime_hours
        # basic_duration, overtime_hours = calculate_work_hours(unique_records)
        # print(f"基础时长: {basic_duration:.2f} 小时 加班时长: {overtime_hours:.2f} 小时")

        # 将去重打卡记录和工时信息组成一个列表
        result = [unique_records, daily_work_hours, sum(daily_work_hours)]

        # 更新attendance_data的值
        attendance_data[date] = result
        # print(result)

for index, attendance_data in enumerate(kq["考勤数据"]):
    print(f"考勤数据 {index + 1}:")
    for date, records in attendance_data.items():
        print(f"日期: {date} 工时: {records}")

# 创建一个字典来存储每个员工的总工作时长和工时天数
# 新增 "总工作时长" 和 "工时天数" 两列
kq["总工作时长"] = 0.0
kq["工时天数"] = 0.0

for index, attendance_data in enumerate(kq["考勤数据"]):
    total_hours = 0.0
    for date, records in attendance_data.items():
        daily_total_hours = records[2]
        total_hours += daily_total_hours

    # 更新 "总工作时长" 和 "工时天数"
    kq.at[index, "总工作时长"] = total_hours
    kq.at[index, "工时天数"] = total_hours / 8  # 计算工时天数

# 打印每个员工的总工作时长和工时天数
for index, row in kq.iterrows():
    print(f"工号: {row['工号']}, 总工作时长: {row['总工作时长']:.2f} 小时, 工时天数: {row['工时天数']:.2f}")


if names_without_attendance:
    names_str = "、".join(names_without_attendance)
    print(f"抱歉，没有查询到{names_str}的打卡记录")

columns_to_keep = ["姓名", "工号", "总计工时", "总计天数", "总工作时长", "工时天数"]
new_kq = kq[columns_to_keep]
# new_kq.to_csv("kq_exported.csv", index=False)
# print("已成功导出新的 CSV 文件：kq_exported.csv")

new_kq.to_excel("kq_exported11.xlsx", index=False, engine='xlsxwriter')
print("已成功导出新的 Excel 文件：考勤自动计算答案.xlsx")