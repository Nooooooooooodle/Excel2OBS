import pandas as pd
import openpyxl  # 用于读取带宏的Excel文件
import websocket
import json
from tkinter import Tk, filedialog, Label, Entry, Button, Frame, Checkbutton, IntVar, OptionMenu, StringVar, Scrollbar, Canvas
import logging
import os
import threading
import time
import json

# 设置日志
logging.basicConfig(level=logging.DEBUG, format='%(asctime)s - %(levelname)s - %(message)s')

# OBS WebSocket 地址和端口
obs_ws_url = "ws://localhost:4444"

class ExcelToOBS:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel2OBS 作者 B站:直播说 由@姓面名包君 修改 ")

        # 设置窗口图标
        icon_path = 'icon.ico'
        if os.path.exists(icon_path):
            self.root.iconbitmap(icon_path)
        else:
            logging.warning(f"Icon file not found: {icon_path}")

        self.file_path = None
        self.inputs = []
        self.previous_values = {}
        self.obs_connected = False  # 新增：OBS连接状态

        # 配置网格权重，使中间的画布区域可以扩展
        self.root.grid_rowconfigure(2, weight=1)  # 第3行(索引2)获得所有额外空间
        self.root.grid_columnconfigure(0, weight=1)  # 第1列获得所有额外空间
        self.root.grid_columnconfigure(1, weight=1)  # 第2列获得所有额外空间
        self.root.grid_columnconfigure(2, weight=1)  # 第3列获得所有额外空间
        self.root.grid_columnconfigure(3, weight=1)  # 第4列获得所有额外空间

        # 新增：OBS连接状态指示器
        self.status_frame = Frame(root)
        self.status_frame.grid(row=0, column=3, sticky='e', padx=5, pady=5)
        
        self.status_label = Label(self.status_frame, text="OBS Status: ", font=("Arial", 10))
        self.status_label.pack(side="left")
        
        self.status_indicator = Label(self.status_frame, text="Disconnected", fg="red", font=("Arial", 10, "bold"))
        self.status_indicator.pack(side="left")

        Label(root, text="Excel File:").grid(row=0, column=0, sticky='w', padx=5, pady=5)
        self.file_entry = Entry(root)
        self.file_entry.grid(row=0, column=1, sticky='ew', padx=5, pady=5)
        Button(root, text="Browse", command=self.choose_file).grid(row=0, column=2, sticky='ew', padx=5, pady=5)

        Label(root, text="Sheet Name:").grid(row=1, column=0, sticky='w', padx=5, pady=5)
        self.sheet_entry = Entry(root)
        self.sheet_entry.grid(row=1, column=1, sticky='ew', padx=5, pady=5)

        # 创建画布和滚动条
        self.canvas_frame = Frame(root)  # 新增一个框架来包含画布和滚动条
        self.canvas_frame.grid(row=2, column=0, columnspan=4, sticky='nsew')
        self.canvas_frame.grid_rowconfigure(0, weight=1)  # 让画布框架中的画布扩展
        self.canvas_frame.grid_columnconfigure(0, weight=1)  # 让画布框架中的画布扩展

        self.canvas = Canvas(self.canvas_frame)
        self.canvas.grid(row=0, column=0, sticky='nsew')

        self.scrollbar = Scrollbar(self.canvas_frame, command=self.canvas.yview)
        self.scrollbar.grid(row=0, column=1, sticky='ns')
        self.canvas.configure(yscrollcommand=self.scrollbar.set)

        self.inputs_frame = Frame(self.canvas)
        self.canvas_frame_id = self.canvas.create_window((0, 0), window=self.inputs_frame, anchor='nw')

        self.add_input()

        Button(root, text="Add More", command=self.add_input).grid(row=3, column=0, columnspan=4, sticky='ew', padx=5, pady=5)
        Button(root, text="Update Text", command=lambda: self.update_text(check_changes=False)).grid(row=4, column=0, columnspan=4, sticky='ew', padx=5, pady=5)
        Button(root, text="Save Configuration", command=self.save_configuration).grid(row=5, column=0, columnspan=4, sticky='ew', padx=5, pady=5)
        Button(root, text="Load Configuration", command=self.load_configuration).grid(row=6, column=0, columnspan=4, sticky='ew', padx=5, pady=5)

        # 新增：连接测试按钮
        self.connect_button = Button(root, text="测试OBS连接", command=self.test_obs_connection)
        self.connect_button.grid(row=7, column=0, columnspan=4, sticky='ew', padx=5, pady=5)

        self.update_interval = 0.5  # 每0.5秒检测一次
        self.running = True
        self.start_update_thread()
        
        # 新增：启动OBS连接状态检测线程
        self.start_obs_status_thread()

        # 绑定画布滚动事件和窗口大小变化事件
        self.inputs_frame.bind("<Configure>", self._on_frame_configure)
        self.canvas.bind("<Configure>", self._on_canvas_configure)

    def _on_frame_configure(self, event):
        """当输入框架大小改变时，更新画布的滚动区域"""
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))

    def _on_canvas_configure(self, event):
        """当画布大小改变时，调整输入框架的宽度"""
        # 获取画布的宽度
        canvas_width = event.width
        # 设置输入框架的宽度与画布相同
        self.canvas.itemconfig(self.canvas_frame_id, width=canvas_width)

    def choose_file(self):
        """选择Excel文件"""
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx;*.xlsm")])
        self.file_entry.delete(0, 'end')
        self.file_entry.insert(0, file_path)
        self.file_path = file_path
        logging.info(f'Selected file: {file_path}')

    def add_input(self):
        """添加新的行输入"""
        row_index = len(self.inputs)
        data_type_var = StringVar(self.inputs_frame)
        data_type_var.set("Text")  # 默认值
        row_entry = Entry(self.inputs_frame, width=5)  # 设置宽度为5
        column_entry = Entry(self.inputs_frame, width=5)  # 设置宽度为5
        name_entry = Entry(self.inputs_frame)
        value_label = Label(self.inputs_frame, text="N/A")
        check_var = IntVar()
        check_button = Checkbutton(self.inputs_frame, variable=check_var)
        data_type_menu = OptionMenu(self.inputs_frame, data_type_var, "Text", "Image")

        Label(self.inputs_frame, text=f"Input {row_index + 1}:").grid(row=row_index, column=0, padx=5, pady=5)
        data_type_menu.grid(row=row_index, column=1, padx=5, pady=5)
        name_entry.grid(row=row_index, column=2, sticky='ew', padx=5, pady=5)
        row_entry.grid(row=row_index, column=3, padx=5, pady=5)
        column_entry.grid(row=row_index, column=4, padx=5, pady=5)
        value_label.grid(row=row_index, column=5, padx=5, pady=5)
        check_button.grid(row=row_index, column=6, padx=5, pady=5)

        # 配置name_entry列的权重，使其能够扩展
        self.inputs_frame.grid_columnconfigure(2, weight=1)

        row_entry.bind("<KeyRelease>", lambda event: self.update_value_label(row_entry, column_entry, value_label))
        column_entry.bind("<KeyRelease>", lambda event: self.update_value_label(row_entry, column_entry, value_label))

        self.inputs.append((data_type_var, row_entry, column_entry, name_entry, value_label, check_var))

    def update_value_label(self, row_entry, column_entry, value_label):
        """更新值标签"""
        row_str = row_entry.get().strip()
        column_str = column_entry.get().strip()

        if not self.file_path or not os.path.exists(self.file_path):
            logging.error("No valid Excel file selected.")
            return

        sheet_name = self.sheet_entry.get()
        if not sheet_name:
            logging.error("No sheet name provided.")
            return

        if not row_str.isdigit() or not column_str.isdigit():
            logging.error(f"Invalid row or column input: Row - {row_str}, Column - {column_str}. Row and column must be numbers.")
            return

        row = int(row_str) - 1
        column = int(column_str) - 1

        try:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
            if row < 0 or column < 0 or row >= len(df) or column >= len(df.columns):
                logging.error(f"Row or column out of range. Row: {row + 1}, Column: {column + 1}")
                return

            value = df.iloc[row, column]
            if isinstance(value, float) and value.is_integer():
                value = int(value)
            logging.info(f'Read value from Excel: {value}')
            value_label.config(text=str(value))
        except Exception as e:
            logging.error(f'Error reading from Excel: {e}')

    def start_update_thread(self):
        """启动后台线程定期更新数据"""
        threading.Thread(target=self.periodic_update, daemon=True).start()

    def periodic_update(self):
        """定期更新数据"""
        while self.running:
            self.update_text(check_changes=True)
            time.sleep(self.update_interval)

    def update_text(self, check_changes=False):
        """从Excel读取数据并更新到OBS"""
        if not self.file_path:
            logging.error("No Excel file selected.")
            return

        sheet_name = self.sheet_entry.get()
        if not sheet_name:
            logging.error("No sheet name provided.")
            return

        try:
            df = pd.read_excel(self.file_path, sheet_name=sheet_name, engine='openpyxl', header=None)
            for i, (data_type_var, row_entry, column_entry, name_entry, value_label, check_var) in enumerate(self.inputs):
                row_str = row_entry.get().strip()
                column_str = column_entry.get().strip()
                source_name = name_entry.get().strip()

                if not row_str.isdigit() or not column_str.isdigit():
                    logging.error(f"Invalid row or column input: Row - {row_str}, Column - {column_str}. Row and column must be numbers.")
                    continue

                row = int(row_str) - 1
                column = int(column_str) - 1

                if row < 0 or column < 0 or row >= len(df) or column >= len(df.columns):
                    logging.error(f"Row or column out of range. Row: {row + 1}, Column: {column + 1}")
                    continue

                logging.debug(f'User Input - Row: {row + 1}, Column: {column + 1}')
                logging.debug(f'Calculated Index - Row: {row}, Column: {column}')

                try:
                    value = df.iloc[row, column]
                    if isinstance(value, float) and value.is_integer():
                        value = int(value)
                    logging.info(f'Read value from Excel: {value}')
                    value_label.config(text=str(value))

                    if source_name:
                        if check_changes:
                            if check_var.get():
                                previous_value = self.previous_values.get((row, column))
                                if previous_value != value:
                                    logging.info(f'Value changed from {previous_value} to {value}')
                                    self.send_update_to_obs(data_type_var.get(), value, source_name)
                                self.previous_values[(row, column)] = value
                        else:
                            self.send_update_to_obs(data_type_var.get(), value, source_name)
                            self.previous_values[(row, column)] = value
                except Exception as e:
                    logging.error(f'Error reading from Excel: {e}')
        except Exception as e:
            logging.error(f'Error loading Excel file: {e}')

    def send_update_to_obs(self, data_type, value, source_name):
        """根据数据类型将更新发送到OBS"""
        if not self.obs_connected:
            logging.warning("无法发送数据到OBS：未连接到OBS服务器")
            self.update_obs_status(False)
            return
            
        if data_type == "Image":
            logging.info(f"Updating image source '{source_name}' with file path: {self.clean_file_path(value)}")
            self.update_obs_image_source(self.clean_file_path(value), source_name)
        else:
            logging.info(f"Updating text source '{source_name}' with text: {str(value)}")
            self.update_obs_text_source(str(value), source_name)

    def clean_file_path(self, file_path):
        """清理文件路径中的不可见字符"""
        cleaned_path = file_path.strip()
        logging.debug(f'Original file path: {file_path}')
        # 移除不可见字符
        cleaned_path = ''.join(c for c in cleaned_path if c.isprintable())
        # 进一步处理路径中的特殊字符
        cleaned_path = cleaned_path.replace('\u202a', '').replace('\u202c', '')
        logging.debug(f'Cleaned file path: {cleaned_path}')
        return cleaned_path

    def update_obs_text_source(self, text, source_name):
        """发送文本数据到OBS的指定文本源，适用于OBS WebSocket 5.x"""
        try:
            ws = websocket.create_connection(obs_ws_url)

            identify_message = {
                "op": 1,
                "d": {
                    "rpcVersion": 1
                }
            }
            ws.send(json.dumps(identify_message))
            response = ws.recv()
            logging.info(f'Received identify response: {response}')

            update_message = {
                "op": 6,
                "d": {
                    "requestType": "SetInputSettings",
                    "requestData": {
                        "inputName": source_name,
                        "inputSettings": {
                            "text": text
                        }
                    },
                    "requestId": str(int(time.time()))
                }
            }
            ws.send(json.dumps(update_message))
            response = ws.recv()
            ws.close()
        except Exception as e:
            logging.error(f'Error sending data to OBS: {e}')
            self.update_obs_status(False)

    def update_obs_image_source(self, image_path, source_name):
        """发送图像数据到OBS的指定图像源，适用于OBS WebSocket 5.x"""
        try:
            ws = websocket.create_connection(obs_ws_url)

            identify_message = {
                "op": 1,
                "d": {
                    "rpcVersion": 1
                }
            }
            ws.send(json.dumps(identify_message))
            response = ws.recv()
            logging.info(f'Received identify response: {response}')

            update_message = {
                "op": 6,
                "d": {
                    "requestType": "SetInputSettings",
                    "requestData": {
                        "inputName": source_name,
                        "inputSettings": {
                            "file": image_path
                        }
                    },
                    "requestId": str(int(time.time()))
                }
            }
            ws.send(json.dumps(update_message))
            response = ws.recv()
            ws.close()
        except Exception as e:
            logging.error(f'Error sending image data to OBS: {e}')
            self.update_obs_status(False)

    def save_configuration(self):
        """保存当前配置信息到文件"""
        config = {
            "file_path": self.file_entry.get(),
            "sheet_name": self.sheet_entry.get(),
            "inputs": []
        }

        for data_type_var, row_entry, column_entry, name_entry, _, check_var in self.inputs:
            config["inputs"].append({
                "data_type": data_type_var.get(),
                "row": row_entry.get(),
                "column": column_entry.get(),
                "name": name_entry.get(),
                "checked": check_var.get()
            })

        file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
        if file_path:
            try:
                with open(file_path, 'w') as f:
                    json.dump(config, f, indent=4)
                logging.info(f"Configuration saved to {file_path}")
            except Exception as e:
                logging.error(f"Error saving configuration: {e}")

    def load_configuration(self):
        """从文件加载配置信息"""
        file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json")])
        if not file_path:
            return

        try:
            with open(file_path, 'r') as f:
                config = json.load(f)

            # 清除当前输入
            for widget in self.inputs_frame.winfo_children():
                widget.destroy()
            self.inputs = []

            # 加载配置
            self.file_entry.delete(0, 'end')
            self.file_entry.insert(0, config.get('file_path', ''))
            self.file_path = config.get('file_path', '')

            self.sheet_entry.delete(0, 'end')
            self.sheet_entry.insert(0, config.get('sheet_name', ''))

            # 添加输入行
            for input_data in config.get('inputs', []):
                self.add_input()
                index = len(self.inputs) - 1
                data_type_var, row_entry, column_entry, name_entry, _, check_var = self.inputs[index]

                data_type_var.set(input_data.get('data_type', 'Text'))
                row_entry.delete(0, 'end')
                row_entry.insert(0, input_data.get('row', ''))
                column_entry.delete(0, 'end')
                column_entry.insert(0, input_data.get('column', ''))
                name_entry.delete(0, 'end')
                name_entry.insert(0, input_data.get('name', ''))
                check_var.set(input_data.get('checked', 0))

            logging.info(f"Configuration loaded from {file_path}")
        except Exception as e:
            logging.error(f"Error loading configuration: {e}")

    # 新增：测试OBS连接
    def test_obs_connection(self):
        threading.Thread(target=self._test_obs_connection_thread, daemon=True).start()

    def _test_obs_connection_thread(self):
        try:
            ws = websocket.create_connection(obs_ws_url, timeout=2)
            
            identify_message = {
                "op": 1,
                "d": {
                    "rpcVersion": 1
                }
            }
            ws.send(json.dumps(identify_message))
            response = ws.recv()
            ws.close()
            
            logging.info("成功连接到OBS服务器")
            self.update_obs_status(True)
        except Exception as e:
            logging.error(f"无法连接到OBS服务器: {e}")
            self.update_obs_status(False)

    # 新增：启动OBS连接状态检测线程
    def start_obs_status_thread(self):
        threading.Thread(target=self._check_obs_status_loop, daemon=True).start()

    def _check_obs_status_loop(self):
        while self.running:
            try:
                # 使用短超时来快速检测连接状态
                ws = websocket.create_connection(obs_ws_url, timeout=1)
                ws.close()
                if not self.obs_connected:  # 状态变化时才更新UI
                    self.root.after(0, lambda: self.update_obs_status(True))
            except:
                if self.obs_connected:  # 状态变化时才更新UI
                    self.root.after(0, lambda: self.update_obs_status(False))
            time.sleep(2)  # 每2秒检查一次

    # 新增：更新OBS连接状态显示
    def update_obs_status(self, connected):
        self.obs_connected = connected
        if connected:
            self.status_indicator.config(text="Connected", fg="green")
            self.connect_button.config(text="OBS 已连接", state="disabled")
        else:
            self.status_indicator.config(text="Disconnected", fg="red")
            self.connect_button.config(text="测试OBS连接", state="normal")

if __name__ == "__main__":
    root = Tk()
    app = ExcelToOBS(root)
    root.mainloop()