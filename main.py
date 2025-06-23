import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
import win_process
import window
import json
import keyboard
import threading
import time
import os
import sys
import subprocess
import win32gui  # 添加导入用于获取当前激活窗口


# 移除日志配置以优化性能

def resource_path(relative_path):
    """获取资源绝对路径"""
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


class HotkeySettingsDialog(tk.Toplevel):
    """快捷键设置对话框（修改为按键监听方式）"""

    def __init__(self, parent, current_hotkeys):
        super().__init__(parent)
        self.title("设置快捷键")
        self.geometry("350x250")  # 增加高度以适应新控件
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()

        self.result = None
        self.current_hotkeys = current_hotkeys
        self.listening_for = None  # 当前正在监听的快捷键类型
        self.key_pressed = None  # 用户按下的键

        # 添加上一个角色快捷键设置区域
        first_frame = ttk.Frame(self)
        first_frame.grid(row=0, column=0, padx=10, pady=10, sticky="ew")

        ttk.Label(first_frame, text="上一个角色快捷键:").pack(side=tk.LEFT)
        self.first_var = tk.StringVar(value=current_hotkeys.get('first', ''))
        self.first_label = ttk.Label(first_frame, textvariable=self.first_var, width=15, relief="sunken", padding=5)
        self.first_label.pack(side=tk.LEFT, padx=5)

        self.first_button = ttk.Button(first_frame, text="设置", command=lambda: self.start_listening('first'))
        self.first_button.pack(side=tk.LEFT)

        # 添加下一个角色快捷键设置区域
        next_frame = ttk.Frame(self)
        next_frame.grid(row=1, column=0, padx=10, pady=10, sticky="ew")

        ttk.Label(next_frame, text="下一个角色快捷键:").pack(side=tk.LEFT)
        self.next_var = tk.StringVar(value=current_hotkeys.get('next', ''))
        self.next_label = ttk.Label(next_frame, textvariable=self.next_var, width=15, relief="sunken", padding=5)
        self.next_label.pack(side=tk.LEFT, padx=5)

        self.next_button = ttk.Button(next_frame, text="设置", command=lambda: self.start_listening('next'))
        self.next_button.pack(side=tk.LEFT)

        # 添加按键提示
        self.instruction_var = tk.StringVar()
        ttk.Label(self, textvariable=self.instruction_var, foreground="blue").grid(row=2, column=0, pady=10)

        # 添加按钮区域
        button_frame = ttk.Frame(self)
        button_frame.grid(row=3, column=0, pady=10)

        ttk.Button(button_frame, text="确定", command=self.on_ok).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="取消", command=self.destroy).pack(side=tk.LEFT, padx=5)

        # 绑定键盘事件
        self.bind("<KeyPress>", self.on_key_press)

    def start_listening(self, hotkey_type):
        """开始监听按键"""
        self.listening_for = hotkey_type
        self.instruction_var.set(f"请按下要设置为 {'上一个角色' if hotkey_type == 'first' else '下一个角色'} 的按键...")

        # 禁用按钮防止重复点击
        self.first_button.config(state=tk.DISABLED)
        self.next_button.config(state=tk.DISABLED)

    def stop_listening(self):
        """停止监听按键"""
        self.listening_for = None
        self.instruction_var.set("")
        self.first_button.config(state=tk.NORMAL)
        self.next_button.config(state=tk.NORMAL)

    def on_key_press(self, event):
        """处理按键事件"""
        if not self.listening_for:
            return

        # 获取按下的键（排除修饰键）
        key = event.keysym
        if key in ["Control_L", "Control_R", "Shift_L", "Shift_R", "Alt_L", "Alt_R"]:
            return

        self.key_pressed = key

        # 更新显示
        if self.listening_for == 'first':
            self.first_var.set(key)
        else:
            self.next_var.set(key)

        self.stop_listening()

    def on_ok(self):
        """验证并返回快捷键设置"""
        first_hotkey = self.first_var.get().strip()
        next_hotkey = self.next_var.get().strip()

        if not first_hotkey or not next_hotkey:
            messagebox.showerror("错误", "快捷键不能为空")
            return

        if first_hotkey == next_hotkey:
            messagebox.showerror("错误", "两个快捷键不能相同")
            return

        self.result = {
            'first': first_hotkey,
            'next': next_hotkey
        }
        self.destroy()


class MainWindow:
    def __init__(self, root, config):
        self.root = root
        self.config = config
        self.hotkeys = config.get('hotkeys', {})
        self.characters = config.get('characters', [])
        self.characters_info = []
        self.update_interval = 5  # 增加更新间隔到10秒
        self.drag_data = {"item": None}  # 用于拖拽功能

        # 设置应用图标
        try:
            icon_path = resource_path('app_icon.ico')
            root.iconbitmap(icon_path)
        except:
            pass

        self.create_widgets()
        self.setup_hotkeys()
        self.update_character_info()
        self.start_update_thread()

    def create_widgets(self):
        """创建优化后的GUI界面"""
        self.root.title("EVE多开切换器 v1.0.1")
        self.root.geometry("800x600")  # 增大窗口尺寸确保显示完整
        self.root.resizable(True, True)

        # 使用更现代的clam主题
        style = ttk.Style()
        style.theme_use('clam')

        # 设置背景色
        self.root.configure(background='#f5f5f5')

        # 创建主框架 - 使用带背景色的Frame
        main_frame = ttk.Frame(self.root, padding=15)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        main_frame.configure(relief=tk.RAISED, borderwidth=1)

        # 标题区域
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 15))

        # 应用图标和标题
        try:
            icon_path = resource_path('app_icon.ico')
            img = tk.PhotoImage(file=icon_path)
            icon_label = ttk.Label(header_frame, image=img)
            icon_label.image = img  # 保持引用
            icon_label.pack(side=tk.LEFT, padx=(0, 10))
        except:
            pass

        title_label = ttk.Label(header_frame,
                                text="星华军团版 反馈找幻象 Nany 1419",
                                font=("微软雅黑", 16, "bold"),
                                foreground="#2c3e50")
        title_label.pack(side=tk.LEFT, pady=5)


        # 快捷键信息卡片
        hotkey_card = ttk.LabelFrame(main_frame, text="快捷键设置", padding=10)
        hotkey_card.pack(fill=tk.X, pady=10)

        # 使用网格布局使快捷键信息更整齐
        ttk.Label(hotkey_card, text="上一个角色快捷键:", font=("微软雅黑", 10)).grid(row=0, column=0, sticky="w",
                                                                                     padx=5, pady=3)
        self.hotkey_label1 = ttk.Label(hotkey_card,
                                       text=self.hotkeys.get('first', '未设置'),
                                       font=("Consolas", 10, "bold"),
                                       foreground="#3498db",
                                       width=15,
                                       relief="solid",
                                       padding=3)
        self.hotkey_label1.grid(row=0, column=1, padx=5, pady=3)

        ttk.Label(hotkey_card, text="下一个角色快捷键:", font=("微软雅黑", 10)).grid(row=1, column=0, sticky="w",
                                                                                     padx=5, pady=3)
        self.hotkey_label2 = ttk.Label(hotkey_card,
                                       text=self.hotkeys.get('next', '未设置'),
                                       font=("Consolas", 10, "bold"),
                                       foreground="#3498db",
                                       width=15,
                                       relief="solid",
                                       padding=3)
        self.hotkey_label2.grid(row=1, column=1, padx=5, pady=3)

        # 添加设置按钮
        settings_btn = ttk.Button(hotkey_card, text="修改快捷键", command=self.open_hotkey_settings, width=15)
        settings_btn.grid(row=0, column=2, rowspan=2, padx=10, pady=5, sticky="ns")

        # 角色状态表格区域
        char_frame = ttk.LabelFrame(main_frame,
                                    text="角色状态管理 (拖拽调整顺序 | 双击切换启用状态)",
                                    padding=10)
        char_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # 添加表格说明
        table_note = ttk.Label(char_frame,
                               text="状态说明: 激活 - 当前前台窗口 | 运行中 - 后台运行 | 未运行 - 客户端未启动",
                               foreground="#7f8c8d",
                               font=("微软雅黑", 9))
        table_note.pack(anchor="w", pady=(0, 5))

        # 表格列配置
        columns = ("name", "status", "hwnd", "enabled")
        self.tree = ttk.Treeview(char_frame, columns=columns, show="headings", height=8)

        # 设置列 - 增加名称列宽度
        self.tree.heading("name", text="角色名称", anchor=tk.W)
        self.tree.heading("status", text="运行状态", anchor=tk.CENTER)
        self.tree.heading("hwnd", text="窗口句柄", anchor=tk.CENTER)
        self.tree.heading("enabled", text="启用状态", anchor=tk.CENTER)

        self.tree.column("name", width=220, minwidth=200)  # 增加名称列宽度
        self.tree.column("status", width=100, minwidth=80, anchor=tk.CENTER)
        self.tree.column("hwnd", width=100, minwidth=80, anchor=tk.CENTER)
        self.tree.column("enabled", width=80, minwidth=70, anchor=tk.CENTER)

        # 绑定事件
        self.tree.bind("<ButtonPress-1>", self.on_drag_start)
        self.tree.bind("<ButtonRelease-1>", self.on_drag_end)
        self.tree.bind("<B1-Motion>", self.on_drag_motion)
        self.tree.bind("<Double-1>", self.on_toggle_enabled)

        # 添加滚动条
        scrollbar = ttk.Scrollbar(char_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # 按钮区域 - 使用Frame分组
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 5))

        # 左侧操作按钮组
        left_btn_frame = ttk.Frame(button_frame)
        left_btn_frame.pack(side=tk.LEFT, fill=tk.X, expand=True)

        ttk.Button(left_btn_frame, text="刷新状态", command=self.update_character_info, width=12).pack(side=tk.LEFT,
                                                                                                       padx=3)
        ttk.Button(left_btn_frame, text="切换下一个", command=self.on_next_hotkey, width=12).pack(side=tk.LEFT, padx=3)
        ttk.Button(left_btn_frame, text="重启应用", command=self.restart_application, width=12).pack(side=tk.LEFT,
                                                                                                     padx=3)

        # 右侧操作按钮组
        right_btn_frame = ttk.Frame(button_frame)
        right_btn_frame.pack(side=tk.RIGHT)

        ttk.Button(right_btn_frame, text="退出", command=self.root.quit, width=10).pack(side=tk.RIGHT, padx=3)

        # 状态栏和信息区域
        footer_frame = ttk.Frame(main_frame)
        footer_frame.pack(fill=tk.X, pady=(5, 0))

        # 提示信息
        tip_label = ttk.Label(footer_frame,
                              text="提示: 如果切换功能异常，请设置EVE为窗口模式或尝试重启应用",
                              foreground="#e74c3c",
                              font=("微软雅黑", 9))
        tip_label.pack(side=tk.LEFT, anchor="w")

        # 状态栏
        self.status_var = tk.StringVar(value="就绪 | 快捷键: " +
                                             f"{self.hotkeys.get('first', '未设置')}/{self.hotkeys.get('next', '未设置')}")
        status_bar = ttk.Label(self.root,
                               textvariable=self.status_var,
                               relief=tk.SUNKEN,
                               anchor=tk.W,
                               font=("微软雅黑", 9),
                               background="#ecf0f1")
        status_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # 开发者信息
        dev_frame = ttk.Frame(self.root, height=25)
        dev_frame.pack(side=tk.BOTTOM, fill=tk.X)
        dev_frame.configure(relief=tk.RAISED, borderwidth=1)

        dev_label = ttk.Label(dev_frame,
                              text="本工具由Nany 1419开发 | 欢迎支持ISK捐助",
                              foreground="#3498db",
                              font=("微软雅黑", 9, "italic"))
        dev_label.pack(side=tk.RIGHT, padx=10, pady=3)

    def setup_hotkeys(self):
        """设置快捷键"""
        try:
            # 移除旧的热键
            keyboard.unhook_all()

            # 注册新的热键
            keyboard.add_hotkey(self.hotkeys.get('first'), self.on_first_hotkey)
            keyboard.add_hotkey(self.hotkeys.get('next'), self.on_next_hotkey)

            # 更新界面显示
            self.hotkey_label1.config(text=f"上一个角色: {self.hotkeys.get('first')}")
            self.hotkey_label2.config(text=f"下一个角色: {self.hotkeys.get('next')}")

            self.status_var.set(f"快捷键已设置: {self.hotkeys.get('first')}/{self.hotkeys.get('next')}")
        except Exception as e:
            self.status_var.set(f"快捷键设置失败: {e}")

    def restart_application(self):
        """重启应用程序"""
        self.save_config()
        python = sys.executable
        # 启动新进程
        subprocess.Popen([python] + sys.argv)
        # 关闭当前应用
        self.root.destroy()
        sys.exit()

    def open_hotkey_settings(self):
        """打开快捷键设置对话框"""
        dialog = HotkeySettingsDialog(self.root, self.hotkeys)
        self.root.wait_window(dialog)

        if dialog.result:
            self.hotkeys = dialog.result
            self.config['hotkeys'] = self.hotkeys
            self.save_config()
            self.setup_hotkeys()
            self.status_var.set("快捷键设置已保存并更新")

    def update_character_info(self):
        """更新角色信息（添加启用状态列）"""
        self.tree.delete(*self.tree.get_children())
        self.characters_info = []
        process_list = win_process.get_new_game_processes_list()

        for idx, character in enumerate(self.characters):
            name = character.get('name')
            enabled = character.get('enabled', True)  # 默认为启用
            char_data = win_process.find_character_data(process_list, name)

            if char_data:
                hwnd = char_data.get('hwnd')
                status = "激活" if window.window_is_active(hwnd) else "运行中"
                self.characters_info.append({
                    "name": name,
                    "hwnd": hwnd,
                    "status": status,
                    "enabled": enabled,
                    "index": idx
                })

                # 标记当前激活角色
                tags = ('active',) if status == "激活" else ()
                self.tree.insert("", tk.END, values=(name, status, hwnd, "是" if enabled else "否"), tags=tags)
            else:
                self.tree.insert("", tk.END, values=(name, "未运行", "", "是" if enabled else "否"))

        # 高亮显示当前激活窗口
        self.tree.tag_configure('active', background='#e6f7ff')
        # 禁用状态的行显示为灰色
        self.tree.tag_configure('disabled', foreground='gray')

    def on_toggle_enabled(self, event):
        """双击切换角色启用状态"""
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.tree.identify_column(event.x)
            if column == "#4":  # 启用状态列
                item = self.tree.selection()[0]
                item_values = self.tree.item(item, 'values')
                enabled_text = item_values[3]

                # 切换启用状态
                new_enabled = enabled_text == "否"
                new_enabled_text = "是" if new_enabled else "否"

                # 更新显示
                values = list(item_values)
                values[3] = new_enabled_text
                self.tree.item(item, values=values)

                # 更新配置
                character_name = item_values[0]
                for char in self.characters:
                    if char['name'] == character_name:
                        char['enabled'] = new_enabled
                        break

                # 更新标签
                if new_enabled:
                    self.tree.item(item, tags=())
                else:
                    self.tree.item(item, tags=('disabled',))

                # 保存配置
                self.save_config()
                self.status_var.set(f"已{'启用' if new_enabled else '禁用'}角色: {character_name}")

    def start_update_thread(self):
        """启动状态更新线程"""

        def update_loop():
            while True:
                time.sleep(self.update_interval)
                self.root.after(0, self.update_character_info)

        thread = threading.Thread(target=update_loop, daemon=True)
        thread.start()

    def switch_to_character(self, hwnd):
        """切换到指定窗口句柄"""
        if window.window_active(hwnd):
            self.status_var.set(f"已切换到窗口: {hwnd}")
            self.update_character_info()
        else:
            self.status_var.set(f"切换失败: {hwnd}")

    def is_current_window_in_list(self):
        """检查当前激活窗口是否在角色列表中"""
        # 获取当前激活窗口
        fg_hwnd = win32gui.GetForegroundWindow()

        # 检查是否在角色列表中
        for char in self.characters_info:
            if char['hwnd'] == fg_hwnd:
                return True
        return False

    def on_first_hotkey(self):
        """切换到上一个角色 - 基于当前激活窗口位置"""
        self.status_var.set(f"按下了快捷键: {self.hotkeys.get('first')}")
        self.root.update()

        # 检查当前激活窗口是否在角色列表中
        if not self.is_current_window_in_list():
            self.status_var.set("当前激活窗口不在角色列表中，忽略切换")
            return

        # 获取当前激活窗口
        fg_hwnd = win32gui.GetForegroundWindow()

        # 查找当前激活的角色窗口
        current_char = None
        for char in self.characters_info:
            if char['hwnd'] == fg_hwnd:
                current_char = char
                break

        if current_char:
            # 查找当前角色在列表中的位置
            current_index = self.characters_info.index(current_char)

            # 计算上一个角色的索引
            prev_index = current_index - 1
            if prev_index < 0:
                prev_index = len(self.characters_info) - 1

            # 获取上一个角色
            prev_char = self.characters_info[prev_index]

            # 如果角色被禁用，继续向前查找
            while not prev_char.get('enabled', True):
                prev_index -= 1
                if prev_index < 0:
                    prev_index = len(self.characters_info) - 1
                prev_char = self.characters_info[prev_index]

                # 避免无限循环
                if prev_index == current_index:
                    break

            # 切换到上一个角色
            self.switch_to_character(prev_char['hwnd'])
        else:
            # 如果没有找到当前角色，切换到第一个启用的角色
            for char in self.characters_info:
                if char.get('enabled', True):
                    self.switch_to_character(char['hwnd'])
                    break

    def on_next_hotkey(self):
        """切换到下一个角色 - 基于当前激活窗口位置"""
        self.status_var.set(f"按下了快捷键: {self.hotkeys.get('next')}")
        self.root.update()

        # 检查当前激活窗口是否在角色列表中
        if not self.is_current_window_in_list():
            self.status_var.set("当前激活窗口不在角色列表中，忽略切换")
            return

        # 获取当前激活窗口
        fg_hwnd = win32gui.GetForegroundWindow()

        # 查找当前激活的角色窗口
        current_char = None
        for char in self.characters_info:
            if char['hwnd'] == fg_hwnd:
                current_char = char
                break

        if current_char:
            # 查找当前角色在列表中的位置
            current_index = self.characters_info.index(current_char)

            # 计算下一个角色的索引
            next_index = current_index + 1
            if next_index >= len(self.characters_info):
                next_index = 0

            # 获取下一个角色
            next_char = self.characters_info[next_index]

            # 如果角色被禁用，继续向后查找
            while not next_char.get('enabled', True):
                next_index += 1
                if next_index >= len(self.characters_info):
                    next_index = 0
                next_char = self.characters_info[next_index]

                # 避免无限循环
                if next_index == current_index:
                    break

            # 切换到下一个角色
            self.switch_to_character(next_char['hwnd'])
        else:
            # 如果没有找到当前角色，切换到第一个启用的角色
            for char in self.characters_info:
                if char.get('enabled', True):
                    self.switch_to_character(char['hwnd'])
                    break

    def on_drag_start(self, event):
        """开始拖拽"""
        region = self.tree.identify("region", event.x, event.y)
        if region == "cell":
            item = self.tree.identify_row(event.y)
            if item:
                self.drag_data["item"] = item
                self.tree.selection_set(item)

    def on_drag_end(self, event):
        """结束拖拽"""
        self.drag_data["item"] = None

    def on_drag_motion(self, event):
        """拖拽过程中"""
        if not self.drag_data["item"]:
            return

        target_item = self.tree.identify_row(event.y)
        if target_item and target_item != self.drag_data["item"]:
            # 获取拖拽项和目标项的位置
            drag_index = self.tree.index(self.drag_data["item"])
            target_index = self.tree.index(target_item)

            # 移动角色顺序
            if drag_index < target_index:
                # 向下移动
                self.characters.insert(target_index + 1, self.characters.pop(drag_index))
            else:
                # 向上移动
                self.characters.insert(target_index, self.characters.pop(drag_index))

            # 保存新顺序
            self.config['characters'] = self.characters
            self.save_config()

            # 更新显示
            self.update_character_info()

            # 更新拖拽项
            new_items = self.tree.get_children()
            self.drag_data["item"] = new_items[target_index if drag_index < target_index else target_index]
            self.tree.selection_set(self.drag_data["item"])

    def save_config(self):
        """保存配置到文件"""
        try:
            with open('config.json', 'w') as f:
                json.dump(self.config, f, indent=4)
        except Exception as e:
            messagebox.showerror("错误", f"保存配置失败: {e}")


def load_config():
    """加载配置文件（添加启用状态字段）"""
    config_path = 'config.json'
    default_config = {
        "hotkeys": {
            "first": "ctrl+shift+[",
            "next": "ctrl+shift+]"
        },
        "characters": []
    }

    # 获取当前运行的EVE客户端角色
    process_list = win_process.get_new_game_processes_list()
    current_characters = [item['title'] for item in process_list]

    # 如果配置文件不存在，创建默认配置
    if not os.path.exists(config_path):
        # 添加当前检测到的角色
        for char_name in current_characters:
            default_config['characters'].append({"name": char_name, "enabled": True})

        with open(config_path, 'w') as f:
            json.dump(default_config, f, indent=4)
        return default_config

    # 加载配置文件
    try:
        with open(config_path, 'r') as f:
            config = json.load(f)

            # 确保每个角色都有enabled字段
            for char in config.get('characters', []):
                if 'enabled' not in char:
                    char['enabled'] = True

            # 添加新检测到的角色（如果不存在）
            existing_names = {char['name'] for char in config['characters']}
            for char_name in current_characters:
                if char_name not in existing_names:
                    config['characters'].append({"name": char_name, "enabled": True})

            return config
    except Exception as e:
        messagebox.showerror("配置错误", f"无法加载配置文件: {e}")
        # 使用当前检测到的角色作为默认配置
        for char_name in current_characters:
            default_config['characters'].append({"name": char_name, "enabled": True})
        return default_config


if __name__ == '__main__':
    # 请求管理员权限（根据需要取消注释）
    # if not win_process.run_as_admin():
    #     sys.exit()

    # 加载配置
    config = load_config()

    # 创建主窗口
    root = tk.Tk()
    app = MainWindow(root, config)
    root.mainloop()