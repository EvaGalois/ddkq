import tkinter as tk
from tkinter import filedialog
import pandas as pd

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

def calculate_work_hours(records):
    if len(records) < 4:
        return 0.0, 0.0  # 长度小于4的记录视为工时和加班工时都为0

    last_record = records[-1]
    last_time = pd.to_datetime(last_record, format="%H:%M")

    overtime_hours = 0.0
    if last_time > pd.to_datetime("18:00", format="%H:%M"):
        overtime_hours = round((last_time - pd.to_datetime("17:30", format="%H:%M")).seconds / 3600, 2)

    first_time = pd.to_datetime(records[0], format="%H:%M")
    basic_duration = 8.5  # 基础工时为8.5小时
    if first_time >= pd.to_datetime("07:00", format="%H:%M") and first_time < pd.to_datetime("08:00", format="%H:%M"):
        return basic_duration, overtime_hours

    if first_time > pd.to_datetime("08:00", format="%H:%M"):
        late_duration = round((first_time - pd.to_datetime("08:00", format="%H:%M")).seconds / 3600, 2)
        basic_duration -= late_duration

    return basic_duration, overtime_hours

def choose_dd_file():
    dd_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if dd_file:
        dd_filename = dd_file.split("/")[-1]  # 提取文件名
        dd_filename_label.config(text=f"选择的钉钉表: {dd_filename}")
        dd_filename_label.actual_file_path = dd_file  # 保存完整路径
        dd_button.config(state="disabled")  # 禁用钉钉表按钮
        kq_button.config(state="normal")  # 启用考勤表按钮

def choose_kq_file():
    kq_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if kq_file:
        kq_filename = kq_file.split("/")[-1]  # 提取文件名
        kq_filename_label.config(text=f"选择的考勤表: {kq_filename}")
        kq_filename_label.actual_file_path = kq_file  # 保存完整路径
        kq_button.config(state="disabled")  # 禁用考勤表按钮
        export_button.config(state="normal")  # 启用导出按钮

def export_work_hours():
    try:
        dd_file = dd_filename_label.actual_file_path
        kq_file = kq_filename_label.actual_file_path
        dd_data = pd.read_excel(dd_file, sheet_name="每日统计")
        kq_data = pd.read_excel(kq_file, sheet_name="工时汇总")
        dd = dd_data[["姓名", "工号", "日期", "班次", "打卡记录"]].copy()
        dd["日期"] = dd["日期"].str.split(" ")
        dd["班次"] = dd["班次"].str.split(" ")
        dd["打卡记录"] = dd["打卡记录"].str.split("  \n").apply(remove_trailing_spaces)
        dd["工号"] = dd["工号"].apply(process_employee_id)
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

        for index, row in kq.iterrows():
            dd_indices = row["dd索引"]
            attendance_data = extract_attendance_data(dd, dd_indices)
            kq.at[index, "考勤数据"] = attendance_data

        for index, attendance_data in enumerate(kq["考勤数据"]):
            for date, records in attendance_data.items():
                unique_records = [records[0]]
                for i in range(1, len(records)):
                    current_time = pd.to_datetime(records[i], format="%H:%M")
                    prev_time = pd.to_datetime(unique_records[-1], format="%H:%M")
                    time_diff = (current_time - prev_time).seconds / 60
                    if time_diff > 3:
                        unique_records.append(records[i])

                basic_duration, overtime_hours = calculate_work_hours(unique_records)
                result = [unique_records, [basic_duration, overtime_hours], basic_duration+overtime_hours]
                attendance_data[date] = result

        kq["总工作时长"] = 0.0
        kq["工时天数"] = 0.0

        for index, attendance_data in enumerate(kq["考勤数据"]):
            total_hours = 0.0
            for date, records in attendance_data.items():
                daily_total_hours = records[2]
                total_hours += daily_total_hours

            kq.at[index, "总工作时长"] = total_hours
            kq.at[index, "工时天数"] = total_hours / 8

        columns_to_keep = ["姓名", "工号", "总计工时", "总计天数", "总工作时长", "工时天数"]
        new_kq = kq[columns_to_keep]

        # 弹出目录选择框
        export_dir = filedialog.askdirectory()
        if export_dir:
            export_path = f"{export_dir}/kq_exported.xlsx"
            new_kq.to_excel(export_path, index=False, engine='xlsxwriter')
            result_text.insert(tk.END, f"导出成功，文件保存在: {export_path}\n")
        else:
            result_text.insert(tk.END, "导出取消\n")

        if names_without_attendance:
            names_str = "、".join(names_without_attendance)
            result_text.insert(tk.END, f"抱歉，没有查询到{names_str}的打卡记录\n")

    except Exception as e:
        result_text.insert(tk.END, f"导出失败: {str(e)}\n")

# 创建主窗口并设置宽度
root = tk.Tk()
root.title("工时计算")
root.geometry("360x240")  # 设置宽度为400像素，高度为220像素

# 创建界面组件
dd_button = tk.Button(root, text="选择钉钉表", command=choose_dd_file)
kq_button = tk.Button(root, text="选择考勤表", command=choose_kq_file)
dd_filename_label = tk.Label(root, text="")
kq_filename_label = tk.Label(root, text="")
export_button = tk.Button(root, text="导出工时计算结果", command=export_work_hours, state="disabled")
result_text = tk.Text(root, height=7, width=40)  # 创建Text组件用于显示结果

# 布局界面组件
dd_button.pack()
kq_button.pack()
dd_filename_label.pack()
kq_filename_label.pack()
export_button.pack()
result_text.pack()

# 运行主循环
root.mainloop()
