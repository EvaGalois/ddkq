import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
from datetime import datetime, timedelta
# import re

# def adjust_time_for_fieldwork(records):
#     """
#     处理含有“外勤”字样的数据，将这个元素替换为前一个时间元素加上29分钟。
#     """
#     adjusted_records = []
#     for i, record in enumerate(records):
#         # 检查是否含有"外勤"字眼
#         if "外勤" in record:
#             if i > 0:  # 确保不是第一个元素
#                 try:
#                     # 尝试解析前一个记录的时间，并加上29分钟
#                     prev_time = datetime.strptime(adjusted_records[-1], "%H:%M")
#                     new_time = prev_time + timedelta(minutes=29)
#                     # 将新时间转换回字符串格式，并替换当前元素
#                     adjusted_records.append(new_time.strftime("%H:%M"))
#                 except ValueError:
#                     # 如果前一个时间元素无法解析，保留当前元素不变
#                     adjusted_records.append(record)
#             else:
#                 # 如果是第一个元素且含有"外勤"，直接保留
#                 adjusted_records.append(record)
#         else:
#             # 如果当前元素不含"外勤"，直接添加到结果列表
#             adjusted_records.append(record)
#     return adjusted_records

# def adjust_time_for_fieldwork(records):
#     """
#     处理含有“外勤”字样的数据，将这个元素替换为前一个时间元素加上29分钟。
#     """
#     adjusted_records = []
#     for i, record in enumerate(records):
#         if "外勤" in record:
#             if i > 0:  # 确保不是第一个元素
#                 # 尝试解析前一个记录的时间，并加上29分钟
#                 prev_time = datetime.strptime(adjusted_records[-1], "%H:%M")
#                 new_time = prev_time + timedelta(minutes=29)
#                 # 将新时间转换回字符串格式，并替换当前元素
#                 adjusted_records.append(new_time.strftime("%H:%M"))
#             else:
#                 # 如果是第一个元素，且含有"外勤"，考虑跳过或用默认值处理
#                 continue  # 或者 adjusted_records.append(默认值)
#         else:
#             # 如果当前元素不含"外勤"，直接添加到结果列表
#             adjusted_records.append(record)
#     return adjusted_records

def adjust_time_for_fieldwork(records, default_start_time='08:00'):
    """
    处理含有“外勤”字样的数据，将这个元素替换为前一个时间元素加上29分钟。
    如果是第一个元素含有“外勤”，则使用默认的开始时间。
    """
    adjusted_records = []
    for i, record in enumerate(records):
        if "外勤" in record:
            if i > 0:
                try:
                    prev_time = datetime.strptime(adjusted_records[-1], "%H:%M")
                    new_time = prev_time + timedelta(minutes=29)
                    adjusted_records.append(new_time.strftime("%H:%M"))
                except ValueError as e:
                    print(f"Error parsing time: {e}")
                    # Optionally, handle the error, e.g., by logging or using a default value.
            else:
                # Use default start time for the first "外勤" record
                adjusted_records.append(default_start_time)
        else:
            adjusted_records.append(record)
    return adjusted_records



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

# def process_employee_id(employee_id):
#     return str(employee_id).zfill(7)

# def process_employee_id(employee_id):
#     # count = 1
#     # print(f"employee_id: {employee_id}, {type(employee_id)}")
#     if pd.notna(employee_id):
#         # 如果employee_id不是NaN，转换为整数然后格式化为7位数的字符串
#         # print(f"{str(int(employee_id)).zfill(7)}, {type(str(int(employee_id)).zfill(7))}")
#         return str(int(employee_id)).zfill(7)
#     else:
#         # 如果是NaN，直接返回
#         # print(count + 1)
#         return employee_id

# def process_employee_id(employee_id):
#     # 尝试转换employee_id为字符串，然后去除两端的空格
#     employee_id_str = str(employee_id).strip()
#     try:
#         # 尝试将处理过的字符串转换为整数，并格式化为7位数的字符串
#         return str(int(employee_id_str)).zfill(7)
#     except ValueError:
#         # 如果转换失败（例如，由于存在非数字字符），直接返回原始字符串
#         return employee_id_str

def process_employee_id(employee_id):
    employee_id_str = str(employee_id).strip()  # 去除两端的空格
    if employee_id_str.replace('.', '', 1).isdigit():
        # 这里检查字符串是否仅包含数字（和最多一个小数点，用于处理浮点数）
        # 如果是，尝试将其转换为整数，并格式化为7位数的字符串
        return str(int(float(employee_id_str))).zfill(7)
    else:
        # 如果包含非数字字符，直接返回处理后的字符串
        return employee_id_str



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
            attendance_record = adjust_time_for_fieldwork(attendance_record)
        else:
            attendance_record = [attendance_record]
        attendance_data[date] = attendance_record if pd.notna(attendance_record[0]) else [0]
    return attendance_data

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

def export_work_hours():
    try:
        # 确保文件已被选择
        dd_file = dd_filename_label.actual_file_path if 'actual_file_path' in dir(dd_filename_label) else ''
        kq_file = kq_filename_label.actual_file_path if 'actual_file_path' in dir(kq_filename_label) else ''
        
        if not dd_file or not kq_file:
            messagebox.showerror("错误", "请先选择文件！")
            return
        
        # 获取用户输入的sheet名称
        dd_sheet_name = dd_sheet_name_entry.get()
        kq_sheet_name = kq_sheet_name_entry.get()

        dd_data = pd.read_excel(dd_file, sheet_name=dd_sheet_name)
        kq_data = pd.read_excel(kq_file, sheet_name=kq_sheet_name)

        # 在此处进行数据处理，如您之前定义的逻辑
        # 处理dd数据
        # dd_data['工号'] = dd_data['工号'].apply(lambda x: str(int(x)).zfill(7) if pd.notna(x) else x)

        # specific_employee = dd_data[dd_data['工号'] == '0002965']
        # print(specific_employee)
        dd = dd_data[["姓名", "工号", "日期", "班次", "打卡记录"]].copy()
        dd["日期"] = dd["日期"].str.split(" ")
        dd["班次"] = dd["班次"].str.split(" ")
        dd["打卡记录"] = dd["打卡记录"].str.split("  \n").apply(remove_trailing_spaces)
        # print(dd["打卡记录"])
        dd["工号"] = dd["工号"].apply(process_employee_id)

        # 处理kq数据并计算考勤
        kq = kq_data[["姓名", "工号", "总计工时", "总计天数"]].copy()
        kq["工号"] = kq["工号"].apply(process_employee_id)
        kq["dd索引"] = [[] for _ in range(len(kq))]
        kq["考勤数据"] = [{} for _ in range(len(kq))]
        names_without_attendance = []

        # specific_employee = dd[dd['工号'] == '0002965']
        # print(specific_employee)
        
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
                
                interval = calculate_time_differences(unique_records)
                # print(f"时间间隙：{interval}")
                work_length = []
                for i in interval:
                    if i < 30:
                        pass
                    else:
                        work_length.append(i)

                daily_work_hours = calculate_daily_work_hours(unique_records, work_length)
                result = [unique_records, daily_work_hours, sum(daily_work_hours)]
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
            export_path = f"{export_dir}/自动公式计算结果.xlsx"
            new_kq.to_excel(export_path, index=False, engine='openpyxl')
            result_text.insert(tk.END, f"导出成功，文件保存在: {export_path}\n")
            messagebox.showinfo("成功", f"导出成功，文件保存在: {export_path}")
        else:
            result_text.insert(tk.END, "导出取消\n")
            messagebox.showinfo("取消", "导出取消")

        if names_without_attendance:
            names_str = "、".join(names_without_attendance)
            result_text.insert(tk.END, f"抱歉，没有查询到{names_str}的打卡记录\n")

    except Exception as e:
        result_text.insert(tk.END, f"导出失败: {str(e)}\n")

def choose_dd_file():
    global dd_sheet_name_entry  # 声明为全局变量，以便在其他地方访问
    dd_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if dd_file:
        dd_filename = dd_file.split("/")[-1]
        dd_filename_label.config(text=f"选择的钉钉表: {dd_filename}")
        dd_filename_label.actual_file_path = dd_file
        dd_button.config(state="disabled")
        kq_button.config(state="normal")
        # 显示sheet名称输入框
        dd_sheet_name_entry.pack(pady=5)

def choose_kq_file():
    global kq_sheet_name_entry  # 声明为全局变量，以便在其他地方访问
    kq_file = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if kq_file:
        kq_filename = kq_file.split("/")[-1]
        kq_filename_label.config(text=f"选择的考勤表: {kq_filename}")
        kq_filename_label.actual_file_path = kq_file
        kq_button.config(state="disabled")
        export_button.config(state="normal")
        # 显示sheet名称输入框
        kq_sheet_name_entry.pack(pady=5)

# 初始化主题
def init_dark_theme(root):
    root.configure(bg="#333333")

    style = ttk.Style()
    style.theme_use("clam")

    # 配置大号按钮样式
    style.configure("Large.TButton", background="#555555", foreground="#CCCCCC", 
                    borderwidth=1, font=('Helvetica', 12), padding=6)
    style.map("Large.TButton", background=[('active', '#666666'), ('pressed', '#555555')])

    # 配置标签和输入框的暗黑风格
    style.configure("TLabel", background="#333333", foreground="#CCCCCC", font=('Helvetica', 10))
    style.configure("TEntry", fieldbackground="#555555", foreground="#CCCCCC", borderwidth=1)

    return {"text_bg": "#555555", "text_fg": "#CCCCCC", "label_bg": "#333333", "label_fg": "#CCCCCC"}

# 调整部分代码以适配暗黑风格
root = tk.Tk()
root.title("工时计算")
root.geometry("360x385")

dd_sheet_name_entry = tk.Entry(root, width=20)
kq_sheet_name_entry = tk.Entry(root, width=20)
dd_sheet_name_entry.insert(0, "每日统计")
kq_sheet_name_entry.insert(0, "工时汇总")

# 应用暗黑风格
colors = init_dark_theme(root)

# 创建界面组件
dd_button = ttk.Button(root, text="选择钉钉表", command=choose_dd_file, style="Large.TButton")
kq_button = ttk.Button(root, text="选择考勤表", command=choose_kq_file, style="Large.TButton")
dd_filename_label = ttk.Label(root, text="", background=colors["label_bg"], foreground=colors["label_fg"])
kq_filename_label = ttk.Label(root, text="", background=colors["label_bg"], foreground=colors["label_fg"])
export_button = ttk.Button(root, text="导出工时计算结果", command=export_work_hours, style="Large.TButton")
result_text = tk.Text(root, height=7, width=40, bg=colors["text_bg"], fg=colors["text_fg"], borderwidth=0)

# 使用pack布局管理器布局组件
dd_button.pack(pady=5)
kq_button.pack(pady=5)
dd_filename_label.pack(pady=5)
kq_filename_label.pack(pady=5)

# 这里直接将输入框布局到界面上
dd_sheet_name_entry.pack(pady=5)
kq_sheet_name_entry.pack(pady=5)
export_button.pack(pady=5)
result_text.pack(pady=5)

# 运行主循环
root.mainloop()


# # 创建主窗口并设置宽度
# root = tk.Tk()
# root.title("工时计算")
# root.geometry("360x240")

# # 创建界面组件
# dd_button = tk.Button(root, text="选择钉钉表", command=choose_dd_file)
# kq_button = tk.Button(root, text="选择考勤表", command=choose_kq_file)
# dd_filename_label = tk.Label(root, text="")
# kq_filename_label = tk.Label(root, text="")
# export_button = tk.Button(root, text="导出工时计算结果", command=export_work_hours, state="disabled")
# result_text = tk.Text(root, height=7, width=40)  # 创建Text组件用于显示结果

# # 布局界面组件
# dd_button.pack()
# kq_button.pack()
# dd_filename_label.pack()
# kq_filename_label.pack()
# export_button.pack()
# result_text.pack()

# # 运行主循环
# root.mainloop()
