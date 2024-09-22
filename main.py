import tkinter as tk
from tkinter import messagebox, filedialog
from tkinter import ttk
import json
import os
import pyautogui  # 确保已安装：pip install pyautogui
import threading
import time
from PIL import ImageGrab  # 确保已安装：pip install pillow

# 默认配置文件目录
DEFAULT_CONFIG_DIR = os.path.join(os.path.expanduser("~"), ".tkinter_app")
SETTINGS_FILE = "settings.json"
ITEMS_FILE = "items.json"

class TkinterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("增强的 Tkinter 示例界面")
        self.root.geometry("1400x700")  # 扩大窗口大小以适应Treeview

        # 配置目录和文件路径
        self.config_dir = DEFAULT_CONFIG_DIR
        self.settings_path = os.path.join(self.config_dir, SETTINGS_FILE)
        self.items_path = os.path.join(self.config_dir, ITEMS_FILE)

        # 数据结构：items 是一个列表，每个项目是一个字典，包含 'coordinates', 'color', 'judge_color', 'click', 'delay', 'delay_time', 'remarks'
        self.items = []

        # 运行控制
        self.running = False
        self.stop_event = threading.Event()

        # 加载设置
        self.load_settings()

        # 创建主界面
        self.create_widgets()

        # 加载列表项
        self.load_items()

        # 绑定按键F3用于停止（已移除）
        # self.root.bind('<F3>', self.stop_actions)
        # self.root.bind('<f3>', self.stop_actions)

    def create_widgets(self):
        # 使用帧（Frame）来组织布局
        left_frame = tk.Frame(self.root, width=400, height=700, padx=10, pady=10)
        left_frame.pack(side=tk.LEFT, fill=tk.Y)

        right_frame = tk.Frame(self.root, width=1000, height=700, padx=10, pady=10)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        # 左侧部分：标签和按钮
        label = tk.Label(left_frame, text="欢迎使用 Tkinter GUI", font=("Arial", 16))
        label.pack(pady=20)

        # 新增“获取我的坐标”按钮
        button_get_my_coords = tk.Button(left_frame, text="获取我的坐标", command=self.get_my_coordinates)
        button_get_my_coords.pack(pady=10, fill=tk.X)

        # 新增“运行”按钮
        self.button_run = tk.Button(left_frame, text="运行", command=self.run_actions)
        self.button_run.pack(pady=10, fill=tk.X)

        # 设置配置目录按钮
        button_set_dir = tk.Button(left_frame, text="设置配置文件目录", command=self.set_config_directory)
        button_set_dir.pack(pady=10, fill=tk.X)

        # 保存配置文件按钮
        button_save_config = tk.Button(left_frame, text="保存配置文件", command=self.save_all)
        button_save_config.pack(pady=10, fill=tk.X)

        # 右侧部分：Treeview 和操作按钮
        list_label = tk.Label(right_frame, text="项目列表", font=("Arial", 16))
        list_label.pack(pady=10)

        # 创建一个滚动条
        scrollbar = tk.Scrollbar(right_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 创建Treeview
        self.tree = ttk.Treeview(right_frame, columns=("Coordinates", "RGB Color", "Judge Color", "Click", "Delay", "Delay Time", "Remarks"), show='headings', yscrollcommand=scrollbar.set)
        self.tree.pack(fill=tk.BOTH, expand=True)

        # 配置滚动条
        scrollbar.config(command=self.tree.yview)

        # 定义列
        self.tree.heading("Coordinates", text="坐标")
        self.tree.heading("RGB Color", text="RGB颜色")
        self.tree.heading("Judge Color", text="是否判断颜色")
        self.tree.heading("Click", text="是否点击")
        self.tree.heading("Delay", text="是否延时")
        self.tree.heading("Delay Time", text="延迟时间")
        self.tree.heading("Remarks", text="备注")

        self.tree.column("Coordinates", width=150, anchor='center')
        self.tree.column("RGB Color", width=120, anchor='center')
        self.tree.column("Judge Color", width=120, anchor='center')
        self.tree.column("Click", width=80, anchor='center')
        self.tree.column("Delay", width=80, anchor='center')
        self.tree.column("Delay Time", width=100, anchor='center')
        self.tree.column("Remarks", width=200, anchor='w')

        # 绑定双击事件
        self.tree.bind('<Double-1>', self.on_tree_double_click)

        # 添加、删除、复制、上移和下移按钮
        button_add = tk.Button(right_frame, text="新增一行", command=self.add_item)
        button_add.pack(pady=5, fill=tk.X)

        button_delete = tk.Button(right_frame, text="删除一行", command=self.delete_item)
        button_delete.pack(pady=5, fill=tk.X)

        button_copy = tk.Button(right_frame, text="复制一行", command=self.copy_item)
        button_copy.pack(pady=5, fill=tk.X)

        # 上移和下移按钮
        button_move_up = tk.Button(right_frame, text="上移", command=self.move_up)
        button_move_up.pack(pady=5, fill=tk.X)

        button_move_down = tk.Button(right_frame, text="下移", command=self.move_down)
        button_move_down.pack(pady=5, fill=tk.X)

    def get_my_coordinates(self):
        # 点击按钮后，等待5秒，获取鼠标位置和颜色，并自动新增到列表
        messagebox.showinfo("提示", "请在5秒内将鼠标移动到目标位置...")
        self.root.after(5000, self.retrieve_my_coordinates)

    def retrieve_my_coordinates(self):
        # 获取鼠标位置
        x, y = pyautogui.position()

        try:
            # 获取鼠标所在位置的颜色
            color = pyautogui.pixel(x, y)
        except Exception as e:
            messagebox.showerror("错误", f"获取颜色失败: {e}")
            return

        # 自动新增一行列表
        new_item = {
            "coordinates": (x, y),
            "color": color,
            "judge_color": True,  # 默认为是
            "click": False,       # 默认值
            "delay": False,       # 默认值
            "delay_time": 0,      # 默认值
            "remarks": ""         # 默认值
        }
        self.items.append(new_item)
        self.update_treeview_display()
        self.save_items()
        messagebox.showinfo("成功", f"已将坐标 ({x}, {y}) 和颜色 {color} 添加到列表中。")

    def add_item(self):
        # 弹出输入对话框，获取用户输入
        popup = tk.Toplevel(self.root)
        popup.title("新增项目")
        popup.geometry("400x550")
        popup.grab_set()  # 模态窗口

        # 坐标输入
        coordinates_label = tk.Label(popup, text="坐标 (x, y)：")
        coordinates_label.pack(pady=5)
        coordinates_frame = tk.Frame(popup)
        coordinates_frame.pack(pady=5)

        x_label = tk.Label(coordinates_frame, text="X：")
        x_label.grid(row=0, column=0, padx=5, pady=5)
        x_entry = tk.Entry(coordinates_frame, width=10)
        x_entry.grid(row=0, column=1, padx=5, pady=5)

        y_label = tk.Label(coordinates_frame, text="Y：")
        y_label.grid(row=0, column=2, padx=5, pady=5)
        y_entry = tk.Entry(coordinates_frame, width=10)
        y_entry.grid(row=0, column=3, padx=5, pady=5)

        # RGB颜色输入分为R, G, B三个输入框
        rgb_frame = tk.Frame(popup)
        rgb_frame.pack(pady=5)

        r_label = tk.Label(rgb_frame, text="R：")
        r_label.grid(row=0, column=0, padx=5, pady=5)
        r_entry = tk.Entry(rgb_frame, width=5)
        r_entry.grid(row=0, column=1, padx=5, pady=5)

        g_label = tk.Label(rgb_frame, text="G：")
        g_label.grid(row=0, column=2, padx=5, pady=5)
        g_entry = tk.Entry(rgb_frame, width=5)
        g_entry.grid(row=0, column=3, padx=5, pady=5)

        b_label = tk.Label(rgb_frame, text="B：")
        b_label.grid(row=0, column=4, padx=5, pady=5)
        b_entry = tk.Entry(rgb_frame, width=5)
        b_entry.grid(row=0, column=5, padx=5, pady=5)

        # 是否判断颜色
        judge_color_var = tk.BooleanVar(value=True)
        cb_judge_color = tk.Checkbutton(popup, text="是否判断颜色", variable=judge_color_var)
        cb_judge_color.pack(pady=5)

        # 是否点击
        click_var = tk.BooleanVar()
        cb_click = tk.Checkbutton(popup, text="是否点击", variable=click_var)
        cb_click.pack(pady=5)

        # 是否延时
        delay_var = tk.BooleanVar()
        cb_delay = tk.Checkbutton(popup, text="是否延时", variable=delay_var)
        cb_delay.pack(pady=5)

        # 延迟时间输入框
        delay_time_label = tk.Label(popup, text="延迟时间（秒）：")
        delay_time_label.pack(pady=5)
        delay_time_entry = tk.Entry(popup)
        delay_time_entry.pack(pady=5)

        # 备注输入框
        remarks_label = tk.Label(popup, text="备注：")
        remarks_label.pack(pady=5)
        remarks_entry = tk.Entry(popup)
        remarks_entry.pack(pady=5)

        # 增加到列表按钮
        def add_to_list():
            # 获取坐标
            x = x_entry.get().strip()
            y = y_entry.get().strip()
            if not x or not y:
                messagebox.showerror("错误", "坐标X和Y不能为空。")
                return
            try:
                x = int(x)
                y = int(y)
                coordinates = (x, y)
            except:
                messagebox.showerror("错误", "坐标X和Y必须是整数。")
                return

            # 获取颜色
            r = r_entry.get().strip()
            g = g_entry.get().strip()
            b = b_entry.get().strip()
            if r and g and b:
                try:
                    r = int(r)
                    g = int(g)
                    b = int(b)
                    if not (0 <= r <= 255 and 0 <= g <= 255 and 0 <= b <= 255):
                        raise ValueError
                    color = (r, g, b)
                except:
                    messagebox.showerror("错误", "RGB颜色必须是0-255之间的整数。")
                    return
            elif any([r, g, b]):
                messagebox.showerror("错误", "请完整输入R、G、B三个颜色值，或全部留空。")
                return
            else:
                color = None  # 不判断颜色

            # 获取是否判断颜色
            judge_color = judge_color_var.get()

            # 获取延迟时间
            if delay_var.get():
                delay_time = delay_time_entry.get().strip()
                if not delay_time:
                    messagebox.showerror("错误", "延迟时间不能为空。")
                    return
                try:
                    delay_time = float(delay_time)
                    if delay_time < 0:
                        raise ValueError
                except ValueError:
                    messagebox.showerror("错误", "延迟时间必须是非负数字。")
                    return
            else:
                delay_time = 0

            # 获取备注
            remarks = remarks_entry.get().strip()

            # 创建新项目
            new_item = {
                "coordinates": coordinates,
                "color": color,
                "judge_color": judge_color,
                "click": click_var.get(),
                "delay": delay_var.get(),
                "delay_time": delay_time,
                "remarks": remarks
            }
            self.items.append(new_item)
            self.update_treeview_display()
            self.save_items()
            popup.destroy()
            messagebox.showinfo("成功", "已将信息添加到列表中。")

        def cancel():
            popup.destroy()

        button_add = tk.Button(popup, text="增加到列表", command=add_to_list)
        button_add.pack(pady=10)

        button_cancel = tk.Button(popup, text="取消", command=cancel)
        button_cancel.pack(pady=5)

    def delete_item(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要删除的项目。")
            return
        # Collect indices and sort in reverse to prevent reindexing issues
        indices = sorted([self.tree.index(item) for item in selected_items], reverse=True)
        for index in indices:
            del self.items[index]
        self.update_treeview_display()
        self.save_items()

    def copy_item(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要复制的项目。")
            return
        for item_id in selected_items:
            index = self.tree.index(item_id)
            item = self.items[index].copy()  # 深拷贝以避免引用问题
            # Optionally, you can modify some fields like remarks to indicate it's a copy
            # item['remarks'] = item.get('remarks', '') + " (复制)"
            self.items.append(item)
        self.update_treeview_display()
        self.save_items()
        messagebox.showinfo("成功", "已复制选中的项目到列表底部。")

    def set_config_directory(self):
        # 弹出目录选择对话框
        new_dir = filedialog.askdirectory(title="选择配置文件目录")
        if new_dir:
            self.config_dir = new_dir
            self.settings_path = os.path.join(self.config_dir, SETTINGS_FILE)
            self.items_path = os.path.join(self.config_dir, ITEMS_FILE)
            self.save_settings()
            self.load_items()
            messagebox.showinfo("成功", f"配置文件目录已设置为：{self.config_dir}")

    def load_settings(self):
        # 如果设置文件存在，则加载配置目录
        if os.path.exists(self.settings_path):
            try:
                with open(self.settings_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                self.config_dir = settings.get("config_dir", DEFAULT_CONFIG_DIR)
                self.items_path = os.path.join(self.config_dir, ITEMS_FILE)
            except Exception as e:
                messagebox.showerror("错误", f"加载设置失败: {e}")
                self.config_dir = DEFAULT_CONFIG_DIR
                self.items_path = os.path.join(self.config_dir, ITEMS_FILE)
        else:
            # 如果配置目录不存在，则创建它
            if not os.path.exists(self.config_dir):
                os.makedirs(self.config_dir)
            self.save_settings()

    def save_settings(self):
        # 保存配置目录到设置文件
        settings = {
            "config_dir": self.config_dir
        }
        try:
            with open(self.settings_path, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=4)
        except Exception as e:
            messagebox.showerror("错误", f"保存设置失败: {e}")

    def load_items(self):
        # 清空当前Treeview
        self.tree.delete(*self.tree.get_children())
        # 如果配置目录不存在，则创建它
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)
        # 如果项目文件存在，则加载
        if os.path.exists(self.items_path):
            try:
                with open(self.items_path, 'r', encoding='utf-8') as f:
                    self.items = json.load(f)
                for item in self.items:
                    # 确保每个项目包含所有键
                    if 'coordinates' not in item:
                        item['coordinates'] = (0, 0)
                    if 'color' not in item:
                        item['color'] = None
                    if 'judge_color' not in item:
                        item['judge_color'] = True  # 默认为是
                    if 'click' not in item:
                        item['click'] = False
                    if 'delay' not in item:
                        item['delay'] = False
                    if 'delay_time' not in item:
                        item['delay_time'] = 0
                    if 'remarks' not in item:
                        item['remarks'] = ""
                self.update_treeview_display()
            except Exception as e:
                messagebox.showerror("错误", f"加载项目列表失败: {e}")
                self.items = []

    def save_items(self):
        # 保存到项目文件
        try:
            with open(self.items_path, 'w', encoding='utf-8') as f:
                json.dump(self.items, f, ensure_ascii=False, indent=4)
            self.update_treeview_display()
        except Exception as e:
            messagebox.showerror("错误", f"保存项目列表失败: {e}")

    def save_all(self):
        """
        手动保存所有配置，包括设置和列表项。
        """
        self.save_settings()
        self.save_items()
        messagebox.showinfo("保存成功", "配置文件和列表项已成功保存。")

    def on_tree_double_click(self, event):
        # 获取被点击的项
        selected_item = self.tree.selection()
        if not selected_item:
            return
        item_id = selected_item[0]
        index = self.tree.index(item_id)
        item = self.items[index]

        # 创建编辑窗口
        popup = tk.Toplevel(self.root)
        popup.title(f"编辑项目：{item['coordinates']}")
        popup.geometry("400x600")
        popup.grab_set()  # 模态窗口

        # 坐标输入（允许修改）
        coordinates_label = tk.Label(popup, text="坐标 (x, y)：")
        coordinates_label.pack(pady=5)
        coordinates_frame = tk.Frame(popup)
        coordinates_frame.pack(pady=5)

        x_label = tk.Label(coordinates_frame, text="X：")
        x_label.grid(row=0, column=0, padx=5, pady=5)
        x_entry = tk.Entry(coordinates_frame, width=10)
        x_entry.grid(row=0, column=1, padx=5, pady=5)
        x_entry.insert(0, str(item['coordinates'][0]))

        y_label = tk.Label(coordinates_frame, text="Y：")
        y_label.grid(row=0, column=2, padx=5, pady=5)
        y_entry = tk.Entry(coordinates_frame, width=10)
        y_entry.grid(row=0, column=3, padx=5, pady=5)
        y_entry.insert(0, str(item['coordinates'][1]))

        # RGB颜色输入分为R, G, B三个输入框
        rgb_frame = tk.Frame(popup)
        rgb_frame.pack(pady=5)

        r_label = tk.Label(rgb_frame, text="R：")
        r_label.grid(row=0, column=0, padx=5, pady=5)
        r_entry = tk.Entry(rgb_frame, width=5)
        r_entry.grid(row=0, column=1, padx=5, pady=5)

        g_label = tk.Label(rgb_frame, text="G：")
        g_label.grid(row=0, column=2, padx=5, pady=5)
        g_entry = tk.Entry(rgb_frame, width=5)
        g_entry.grid(row=0, column=3, padx=5, pady=5)

        b_label = tk.Label(rgb_frame, text="B：")
        b_label.grid(row=0, column=4, padx=5, pady=5)
        b_entry = tk.Entry(rgb_frame, width=5)
        b_entry.grid(row=0, column=5, padx=5, pady=5)

        if item['color']:
            r_entry.insert(0, str(item['color'][0]))
            g_entry.insert(0, str(item['color'][1]))
            b_entry.insert(0, str(item['color'][2]))

        # 是否判断颜色
        judge_color_var = tk.BooleanVar(value=item.get("judge_color", True))
        cb_judge_color = tk.Checkbutton(popup, text="是否判断颜色", variable=judge_color_var)
        cb_judge_color.pack(pady=5)

        # 是否点击
        click_var = tk.BooleanVar(value=item.get("click", False))
        cb_click = tk.Checkbutton(popup, text="是否点击", variable=click_var)
        cb_click.pack(pady=5)

        # 是否延时
        delay_var = tk.BooleanVar(value=item.get("delay", False))
        cb_delay = tk.Checkbutton(popup, text="是否延时", variable=delay_var)
        cb_delay.pack(pady=5)

        # 延迟时间输入框
        delay_time_label = tk.Label(popup, text="延迟时间（秒）：")
        delay_time_label.pack(pady=5)
        delay_time_entry = tk.Entry(popup)
        delay_time_entry.pack(pady=5)
        delay_time_entry.insert(0, str(item.get("delay_time", 0)))

        # 备注输入框
        remarks_label = tk.Label(popup, text="备注：")
        remarks_label.pack(pady=5)
        remarks_entry = tk.Entry(popup)
        remarks_entry.pack(pady=5)
        remarks_entry.insert(0, item.get("remarks", ""))

        # 保存按钮
        def save_edit():
            # 获取坐标
            x = x_entry.get().strip()
            y = y_entry.get().strip()
            if not x or not y:
                messagebox.showerror("错误", "坐标X和Y不能为空。")
                return
            try:
                x = int(x)
                y = int(y)
                coordinates = (x, y)
            except:
                messagebox.showerror("错误", "坐标X和Y必须是整数。")
                return

            # 获取颜色
            r = r_entry.get().strip()
            g = g_entry.get().strip()
            b = b_entry.get().strip()
            if r and g and b:
                try:
                    r = int(r)
                    g = int(g)
                    b = int(b)
                    if not (0 <= r <= 255 and 0 <= g <= 255 and 0 <= b <= 255):
                        raise ValueError
                    color = (r, g, b)
                except:
                    messagebox.showerror("错误", "RGB颜色必须是0-255之间的整数。")
                    return
            elif any([r, g, b]):
                messagebox.showerror("错误", "请完整输入R、G、B三个颜色值，或全部留空。")
                return
            else:
                color = None  # 不判断颜色

            # 获取是否判断颜色
            judge_color = judge_color_var.get()

            # 获取延迟时间
            if delay_var.get():
                delay_time = delay_time_entry.get().strip()
                if not delay_time:
                    messagebox.showerror("错误", "延迟时间不能为空。")
                    return
                try:
                    delay_time = float(delay_time)
                    if delay_time < 0:
                        raise ValueError
                except ValueError:
                    messagebox.showerror("错误", "延迟时间必须是非负数字。")
                    return
            else:
                delay_time = 0

            # 获取备注
            remarks = remarks_entry.get().strip()

            # 更新项目数据
            self.items[index]['coordinates'] = coordinates
            self.items[index]['color'] = color
            self.items[index]['judge_color'] = judge_color
            self.items[index]['click'] = click_var.get()
            self.items[index]['delay'] = delay_var.get()
            self.items[index]['delay_time'] = delay_time
            self.items[index]['remarks'] = remarks

            self.save_items()
            popup.destroy()
            messagebox.showinfo("成功", "项目已更新。")

        button_save = tk.Button(popup, text="保存", command=save_edit)
        button_save.pack(pady=10)

        button_cancel = tk.Button(popup, text="取消", command=popup.destroy)
        button_cancel.pack(pady=5)

    def move_up(self):
        selected_items = list(self.tree.selection())
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要上移的项目。")
            return

        # 获取选中项的索引，并排序
        indices = sorted([self.tree.index(item) for item in selected_items])
        if 0 in indices:
            messagebox.showwarning("警告", "最顶部的项目无法上移。")
            return

        for index in indices:
            # 交换项目与上方的项目
            self.items[index - 1], self.items[index] = self.items[index], self.items[index - 1]

        self.update_treeview_display()
        self.save_items()

        # 重新选择移动后的项目
        self.tree.selection_remove(*self.tree.selection())
        for index in indices:
            self.tree.selection_add(index - 1)

    def move_down(self):
        selected_items = list(self.tree.selection())
        if not selected_items:
            messagebox.showwarning("警告", "请先选择要下移的项目。")
            return

        max_index = len(self.items) - 1
        indices = sorted([self.tree.index(item) for item in selected_items], reverse=True)
        if max_index in indices:
            messagebox.showwarning("警告", "最底部的项目无法下移。")
            return

        for index in indices:
            # 交换项目与下方的项目
            self.items[index + 1], self.items[index] = self.items[index], self.items[index + 1]

        self.update_treeview_display()
        self.save_items()

        # 重新选择移动后的项目
        self.tree.selection_remove(*self.tree.selection())
        for index in indices:
            self.tree.selection_add(index + 1)

    def run_actions(self):
        if self.running:
            messagebox.showwarning("警告", "脚本正在运行中。")
            return

        # 创建运行选项窗口
        run_popup = tk.Toplevel(self.root)
        run_popup.title("运行选项")
        run_popup.geometry("350x300")
        run_popup.grab_set()  # 模态窗口

        # 循环选项
        loop_var = tk.BooleanVar()
        cb_loop = tk.Checkbutton(run_popup, text="循环（勾选后设置次数无效）", variable=loop_var)
        cb_loop.pack(pady=10, anchor='w')

        # 次数选项
        count_label = tk.Label(run_popup, text="次数：")
        count_label.pack(pady=5, anchor='w')
        count_entry = tk.Entry(run_popup)
        count_entry.pack(pady=5, fill=tk.X, padx=20)

        # 间隔时间选项
        interval_label = tk.Label(run_popup, text="间隔时间（秒）：")
        interval_label.pack(pady=5, anchor='w')
        interval_entry = tk.Entry(run_popup)
        interval_entry.pack(pady=5, fill=tk.X, padx=20)

        # 开始运行按钮
        def start_run():
            loop = loop_var.get()
            try:
                if loop:
                    count = None  # 无限循环
                else:
                    count = int(count_entry.get()) if count_entry.get() else 1
                    if count < 1:
                        raise ValueError
            except ValueError:
                messagebox.showerror("错误", "次数必须是正整数。")
                return
            try:
                interval = float(interval_entry.get()) if interval_entry.get() else 1
                if interval < 0:
                    raise ValueError
            except ValueError:
                messagebox.showerror("错误", "间隔时间必须是非负数字。")
                return

            run_popup.destroy()
            self.running = True
            self.stop_event.clear()
            self.button_run.config(state=tk.DISABLED)
            threading.Thread(target=self.execute_actions, args=(loop, count, interval), daemon=True).start()

        button_start = tk.Button(run_popup, text="开始运行", command=start_run)
        button_start.pack(pady=10)

        button_cancel = tk.Button(run_popup, text="取消", command=run_popup.destroy)
        button_cancel.pack(pady=5)

    def execute_actions(self, loop, count, interval):
        iteration = 0
        while True:
            for item in self.items:
                if self.stop_event.is_set():
                    self.running = False
                    self.button_run.config(state=tk.NORMAL)
                    messagebox.showinfo("已停止", "脚本运行已被停止。")
                    return

                coordinates = item['coordinates']

                x, y = coordinates

                color = item['color']
                judge_color = item.get('judge_color', True)
                click = item['click']
                delay = item['delay']
                delay_time = item['delay_time']
                remarks = item['remarks']

                perform_click = True




                if perform_click and click:
                    try:
                        if judge_color and color:
                            try:
                                current_color = pyautogui.pixel(x, y)
                                if current_color != color:
                                    continue
                            except Exception as e:
                                messagebox.showerror("错误", f"获取颜色失败: {e}")
                                continue  # Skip this item
                        # 移动鼠标到坐标
                        pyautogui.moveTo(x, y, duration=0.5)
                        pyautogui.click()
                    except Exception as e:
                        messagebox.showerror("错误", f"点击失败: {e}")
                        continue



                if delay:
                    # 分段延迟以便及时响应停止事件
                    total_sleep = delay_time
                    elapsed = 0
                    while elapsed < total_sleep:
                        if self.stop_event.is_set():
                            self.running = False
                            self.button_run.config(state=tk.NORMAL)
                            messagebox.showinfo("已停止", "脚本运行已被停止。")
                            return
                        time.sleep(0.1)
                        elapsed += 0.1

            iteration += 1
            if not loop:
                if count and iteration >= count:
                    break
            time.sleep(interval)

            if loop:
                # 在无限循环中, continue until stopped
                continue
            else:
                if count and iteration >= count:
                    break

        self.running = False
        self.button_run.config(state=tk.NORMAL)
        messagebox.showinfo("运行完成", "所有操作已完成。")

    # def stop_actions(self, event=None):
    #     if self.running:
    #         self.stop_event.set()
    #     else:
    #         messagebox.showinfo("提示", "当前没有正在运行的脚本。")

    def update_treeview_display(self):
        # 清空当前Treeview
        self.tree.delete(*self.tree.get_children())
        # 重新插入项目
        for item in self.items:
            coordinates = f"({item['coordinates'][0]}, {item['coordinates'][1]})"
            if item['color']:
                color = f"({item['color'][0]}, {item['color'][1]}, {item['color'][2]})"
            else:
                color = ""
            judge_color = "是" if item.get("judge_color", True) else "否"
            click = "✔" if item.get("click", False) else ""
            delay = "✔" if item.get("delay", False) else ""
            delay_time = item.get("delay_time", 0)
            remarks = item.get("remarks", "")
            self.tree.insert("", "end", values=(coordinates, color, judge_color, click, delay, delay_time, remarks))

if __name__ == "__main__":
    root = tk.Tk()
    app = TkinterApp(root)
    root.mainloop()
