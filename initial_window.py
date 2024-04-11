import ttkbootstrap as ttk
from tkinter import messagebox
import subprocess
import win32com.client as win32

selected_climate = ""
selected_province = ""
climate_provinces = {"热带": ['北京', '上海'],
                       "亚热带": ['广州', '深圳'],
                       "温带": ['成都', '西安'],
                       "寒带": ['山西', '甘肃']}

root = ttk.Window(themename="solar")
# 设置窗口标题
root.title("种田")


def open_excel():
    root.destroy()
    # 定义要打开的 Excel 文件路径
    excel_file_path = "E:\Projects\Farm\land.xlsm"  # 替换为实际的 Excel 文件路径
    # 打开 Excel 文件并读取数据
    try:
        # 创建Excel应用程序对象
        excel = win32.gencache.EnsureDispatch('Excel.Application')

        # 隐藏Excel应用程序窗口（可选）
        excel.Visible = False

        # 打开或新建一个Excel工作簿
        wb = excel.Workbooks.Open(excel_file_path)  # 替换为你的文件路径
        # 或者新建一个工作簿
        # wb = excel.Workbooks.Add()

        # 选择要设置下拉列表的工作表
        ws = wb.Sheets('Sheet1')  # 替换为你要操作的工作表名称
        
        climate = ws.Range("B1")
        province = ws.Range("D1")
        climate.Value = selected_climate
        province.Value = selected_province

        # 定义下拉列表选项
        dropdown_options = [r'土豆', r'玉米', r'水稻']

        # 设置单元格的数据验证规则
        dv = ws.Range("D2").Validation  # 假设你想要在A1单元格设置下拉列表
        dv.Delete()  # 先删除可能存在的旧规则
        dv.Add(Type=win32.constants.xlValidateList, AlertStyle=win32.constants.xlValidAlertStop, Formula1='{}'.format(','.join(dropdown_options)))

        # 保存并关闭Excel文件
        wb.Save()
        wb.Close(True)  # True表示保存更改

        # 关闭Excel应用程序
        excel.Quit()
        wps_executable_path = r'D:\WPS Office\12.1.0.16412\office6\wps.exe'  # 请替换为实际路径
        subprocess.run([wps_executable_path, excel_file_path])

    except FileNotFoundError:
        messagebox.showerror("错误", "文件未找到，请检查路径是否正确！")
     

def create_window():
    
    # 获取屏幕尺寸
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # 计算窗口居中的位置
    window_width = 400  # 自定义窗口宽度
    window_height = 300  # 自定义窗口高度
    x_pos = (screen_width // 2) - (window_width // 2)
    y_pos = (screen_height // 2) - (window_height // 2)

    # 设置窗口大小和位置
    root.geometry(f"{window_width}x{window_height}+{x_pos}+{y_pos}")

    # 创建一个内部Frame用于布局在同一行的控件
    inner_frame = ttk.Frame(root, padding=10)
    inner_frame.pack()

    # 创建气候区标签并添加到Frame内
    climate_label = ttk.Label(inner_frame, text="气候区：", width=8)
    climate_label.grid(row=0, column=0, padx=(0, 5), pady=(60, 30), sticky=ttk.W)

    # 创建下拉列表框，并填充一些示例数据
    climate_options = ['热带', '亚热带', '温带', '寒带']
    climate_combobox = ttk.Combobox(inner_frame, values=climate_options, width=15)
    climate_combobox.set('请选择气候区')  # 设置默认显示的文字
    climate_combobox.grid(row=0, column=1, pady=(60, 30),sticky=ttk.W + ttk.E)
    def climate_select(event):
        # 获取当前选中的值
        global selected_climate
        selected_climate = climate_combobox.get()
        province_combobox['values'] = climate_provinces[selected_climate]
        province_combobox.set('请选择省份')
    
    climate_combobox.bind("<<ComboboxSelected>>", climate_select)

    province_label = ttk.Label(inner_frame, text="省份：", width=8)
    province_label.grid(row=1, column=0, sticky=ttk.W)

    # 创建下拉列表框，并填充一些示例数据    
    province_combobox = ttk.Combobox(inner_frame, width=15)
    province_combobox.set('请选择省份')  # 设置默认显示的文字
    province_combobox.grid(row=1, column=1, sticky=ttk.W + ttk.E)
    province_combobox['values'] = climate_provinces["热带"]

    def province_selected(event):
        global selected_province
        selected_province = province_combobox.get()

    province_combobox.bind("<<ComboboxSelected>>", province_selected)

    confirm_button = ttk.Button(root, text="确定", command=open_excel)
    confirm_button.pack(pady=(50, 20))

    root.mainloop()

create_window()