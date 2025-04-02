import sys
import requests
import schedule
import time
from datetime import datetime, timedelta
from chinese_calendar import is_workday
import winreg
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QTimeEdit, QCheckBox, QPushButton, QSystemTrayIcon, 
                            QMenu, QLabel, QFrame, QHBoxLayout, QLineEdit)
from PyQt6.QtCore import QTime, Qt, QThread, QTimer
from PyQt6.QtGui import QIcon, QAction
from winotify import Notification, audio
import os
import json

class WorkerThread(QThread):
    def __init__(self):
        super().__init__()
        self.running = True

    def run(self):
        while self.running:
            schedule.run_pending()
            time.sleep(1)

    def stop(self):
        self.running = False

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.config_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
        self.load_config()  # 加载配置
        
        self.setWindowTitle("考勤提醒助手")
        self.setFixedSize(300, 300)  # 增加窗口高度
        
        # 获取图标路径
        self.icon_path = self.get_icon_path()
        self.setWindowIcon(QIcon(self.icon_path))  # 设置窗口图标
        
        # 创建主窗口部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)

        # 添加状态显示区域
        self.status_frame = QFrame()
        self.status_frame.setFrameStyle(QFrame.Shape.Box | QFrame.Shadow.Sunken)
        status_layout = QVBoxLayout(self.status_frame)
        
        self.workday_label = QLabel()
        self.last_workday_label = QLabel()
        self.need_reminder_label = QLabel()
        self.sign_day_label = QLabel()  # 新增考勤签到日标签
        
        status_layout.addWidget(self.workday_label)
        status_layout.addWidget(self.last_workday_label)
        status_layout.addWidget(self.need_reminder_label)
        status_layout.addWidget(self.sign_day_label)  # 添加到布局中
        
        layout.addWidget(self.status_frame)

        # 添加分隔线
        separator = QFrame()
        separator.setFrameShape(QFrame.Shape.HLine)
        separator.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(separator)

        # 原有的控件
        time_container = QWidget()
        time_layout = QHBoxLayout(time_container)  # 改为水平布局
        time_layout.setContentsMargins(0, 0, 0, 0)  # 移除边距
        
        time_label = QLabel("提醒时间：")
        time_layout.addWidget(time_label)
        
        self.time_edit = QTimeEdit()
        self.time_edit.setTime(QTime.fromString(self.config['reminder_time'], "HH:mm"))
        time_layout.addWidget(self.time_edit)
        time_layout.addStretch()  # 添加弹性空间，使控件靠左对齐
        
        layout.addWidget(time_container)

        # 添加自动启动复选框
        self.auto_start_checkbox = QCheckBox("开机自动启动")
        self.auto_start_checkbox.stateChanged.connect(self.toggle_auto_start)
        layout.addWidget(self.auto_start_checkbox)

        # 添加启用提醒复选框
        self.enable_reminder_checkbox = QCheckBox("启用提醒")
        self.enable_reminder_checkbox.stateChanged.connect(self.toggle_reminder)
        layout.addWidget(self.enable_reminder_checkbox)

        # 添加测试按钮
        self.test_button = QPushButton("测试提醒")
        self.test_button.clicked.connect(self.send_notification)
        layout.addWidget(self.test_button)

        # 添加 Bark 链接配置区域
        bark_container = QWidget()
        bark_layout = QHBoxLayout(bark_container)
        bark_layout.setContentsMargins(0, 0, 0, 0)
        
        bark_label = QLabel("Bark链接：")
        bark_layout.addWidget(bark_label)
        
        self.bark_url_edit = QLineEdit()
        self.bark_url_edit.setText(self.config.get('bark_url', 'https://api.day.app/aqnckBntzqSMzySnCkourD'))
        self.bark_url_edit.textChanged.connect(self.bark_url_changed)
        bark_layout.addWidget(self.bark_url_edit)
        
        layout.addWidget(bark_container)

        # 添加分隔线
        separator2 = QFrame()
        separator2.setFrameShape(QFrame.Shape.HLine)
        separator2.setFrameShadow(QFrame.Shadow.Sunken)
        layout.addWidget(separator2)

        # 设置系统托盘
        self.setup_tray()
        
        # 创建工作线程
        self.worker = WorkerThread()
        self.worker.start()

        # 创建定时器，每分钟更新一次状态
        self.status_timer = QTimer(self)
        self.status_timer.timeout.connect(self.update_status)
        self.status_timer.start(60000)  # 60000毫秒 = 1分钟
        
        # 初始更新状态
        self.update_status()

        # 设置控件初始状态
        self.auto_start_checkbox.setChecked(self.config['auto_start'])
        self.enable_reminder_checkbox.setChecked(self.config['enable_reminder'])

        # 如果是开机自启，则启动时隐藏窗口
        if len(sys.argv) > 1 and sys.argv[1] == '--startup':
            self.hide()
            self.show_action.setText("显示")
            self.notify_daily_status()  # 显示当日状态通知

    def get_icon_path(self):
        # 获取当前脚本所在目录
        current_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(current_dir, "app.ico")
        
        # 如果图标文件不存在，创建一个默认图标
        if not os.path.exists(icon_path):
            self.create_default_icon(icon_path)
        
        return icon_path

    def create_default_icon(self, icon_path):
        try:
            from PIL import Image, ImageDraw
            
            # 创建一个 32x32 的图像
            img = Image.new('RGBA', (32, 32), color=(0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            
            # 绘制一个简单的图标（一个蓝色圆形）
            draw.ellipse([4, 4, 28, 28], fill='#0078D7')
            
            # 保存为 ICO 文件
            img.save(icon_path, format='ICO')
        except ImportError:
            # 如果没有安装 Pillow，提示用户
            print("请安装 Pillow 库以创建默认图标：pip install Pillow")
            raise

    def setup_tray(self):
        self.tray = QSystemTrayIcon(self)
        self.tray.setIcon(QIcon(self.icon_path))  # 使用相同的图标
        
        # 创建托盘菜单
        self.tray_menu = QMenu()
        self.show_action = QAction("隐藏", self)  # 默认显示"隐藏"
        quit_action = QAction("退出", self)
        
        self.show_action.triggered.connect(self.toggle_window)
        quit_action.triggered.connect(self.quit_app)
        
        self.tray_menu.addAction(self.show_action)
        self.tray_menu.addAction(quit_action)
        
        self.tray.setContextMenu(self.tray_menu)
        self.tray.show()
        
        # 添加托盘双击事件
        self.tray.activated.connect(self.tray_icon_activated)

    def toggle_window(self):
        if self.isVisible():
            self.hide()
            self.show_action.setText("显示")
        else:
            self.show()
            self.show_action.setText("隐藏")
            self.activateWindow()  # 激活窗口到最前

    def tray_icon_activated(self, reason):
        # 双击托盘图标时触发显示/隐藏
        if reason == QSystemTrayIcon.ActivationReason.DoubleClick:
            self.toggle_window()

    def toggle_auto_start(self, state):
        key = winreg.HKEY_CURRENT_USER
        run_key = r"Software\Microsoft\Windows\CurrentVersion\Run"
        try:
            with winreg.OpenKey(key, run_key, 0, winreg.KEY_ALL_ACCESS) as registry_key:
                if state == Qt.CheckState.Checked.value:
                    # 添加 --startup 参数用于标识开机自启动
                    winreg.SetValueEx(registry_key, "AttendanceReminder", 0, 
                                    winreg.REG_SZ, f'"{sys.argv[0]}" --startup')
                else:
                    winreg.DeleteValue(registry_key, "AttendanceReminder")
            self.save_config()  # 保存配置
        except WindowsError as e:
            self.tray.showMessage("错误", f"设置开机自启动失败：{str(e)}", 
                                QSystemTrayIcon.MessageIcon.Critical)

    def toggle_reminder(self, state):
        if state == Qt.CheckState.Checked.value:
            schedule.every().day.at(self.time_edit.time().toString("HH:mm")).do(
                self.check_and_notify)
        else:
            schedule.clear()
        self.update_status()
        self.save_config()  # 保存配置

    def is_last_workday_of_month(self):
        today = datetime.now()
        # 获取下一天
        if today.month == 12:
            next_month = datetime(today.year + 1, 1, 1)
        else:
            next_month = datetime(today.year, today.month + 1, 1)
        
        current_date = today
        while current_date < next_month:
            if is_workday(current_date) and current_date != today:
                return False
            current_date = current_date + timedelta(days=1)
        
        return is_workday(today)

    def is_24th_of_month(self):
        today = datetime.now()
        return today.day == 24

    def check_and_notify(self):
        today = datetime.now()
        if self.is_last_workday_of_month():
            self.send_notification("请记得填写考勤")
        elif self.is_24th_of_month() and is_workday(today):
            self.send_notification("记得考勤签字")

    def send_notification(self, message):
        try:
            # 使用配置的 Bark URL
            response = requests.get(f"{self.bark_url_edit.text()}/{message}")
            
            # 发送 Windows 通知
            toast = Notification(
                app_id="考勤提醒助手",
                title="考勤提醒",
                msg=message,
                duration="short"
            )
            toast.set_audio(audio.Default, loop=False)
            toast.show()

            if response.status_code == 200:
                self.tray.showMessage("成功", "提醒发送成功！", 
                                    QSystemTrayIcon.MessageIcon.Information)
            else:
                self.tray.showMessage("错误", "提醒发送失败！", 
                                    QSystemTrayIcon.MessageIcon.Warning)
        except Exception as e:
            self.tray.showMessage("错误", f"发送失败：{str(e)}", 
                                QSystemTrayIcon.MessageIcon.Critical)

    def quit_app(self):
        self.worker.stop()
        self.worker.wait()
        QApplication.quit()

    def closeEvent(self, event):
        event.ignore()
        self.hide()
        self.show_action.setText("显示")
        self.save_config()  # 保存配置

    def update_status(self):
        today = datetime.now()
        is_today_workday = is_workday(today)
        is_last = self.is_last_workday_of_month()
        is_sign_day = self.is_24th_of_month() and is_today_workday  # 判断是否是考勤签到日
        need_reminder = (is_today_workday and (is_last or is_sign_day) and 
                        self.enable_reminder_checkbox.isChecked())
        
        self.workday_label.setText(f"今日是否工作日: {'是' if is_today_workday else '否'}")
        self.last_workday_label.setText(f"是否本月最后工作日: {'是' if is_last else '否'}")
        self.need_reminder_label.setText(f"今日是否需要提醒: {'是' if need_reminder else '否'}")
        self.sign_day_label.setText(f"今日是否为考勤签到日: {'是' if is_sign_day else '否'}")  # 新增考勤签到日状态
        
        # 设置样式
        style_enabled = "color: green; font-weight: bold;"
        style_disabled = "color: red;"
        
        self.workday_label.setStyleSheet(style_enabled if is_today_workday else style_disabled)
        self.last_workday_label.setStyleSheet(style_enabled if is_last else style_disabled)
        self.need_reminder_label.setStyleSheet(style_enabled if need_reminder else style_disabled)
        self.sign_day_label.setStyleSheet(style_enabled if is_sign_day else style_disabled)  # 设置考勤签到日样式

    def load_config(self):
        default_config = {
            'reminder_time': '17:00',
            'auto_start': False,
            'enable_reminder': False,
            'bark_url': 'https://api.day.app/aqnckBntzqSMzySnCkourD'  # 添加默认 Bark URL
        }
        
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    self.config = json.load(f)
            else:
                self.config = default_config
        except Exception:
            self.config = default_config

    def save_config(self):
        config = {
            'reminder_time': self.time_edit.time().toString("HH:mm"),
            'auto_start': self.auto_start_checkbox.isChecked(),
            'enable_reminder': self.enable_reminder_checkbox.isChecked(),
            'bark_url': self.bark_url_edit.text()  # 保存 Bark URL
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=4)
        except Exception as e:
            self.tray.showMessage("错误", f"保存配置失败：{str(e)}", 
                                QSystemTrayIcon.MessageIcon.Critical)

    def notify_daily_status(self):
        today = datetime.now()
        is_today_workday = is_workday(today)
        is_last = self.is_last_workday_of_month()
        need_reminder = (is_today_workday and is_last and 
                        self.enable_reminder_checkbox.isChecked())
        
        toast = Notification(
            app_id="考勤提醒助手",
            title="每日状态",
            msg=f"今日是否需要提醒: {'是' if need_reminder else '否'}",
            duration="short"
        )
        toast.set_audio(audio.Default, loop=False)
        toast.show()

    def time_edit_changed(self):
        if self.enable_reminder_checkbox.isChecked():
            schedule.clear()
            schedule.every().day.at(self.time_edit.time().toString("HH:mm")).do(
                self.check_and_notify)
        self.save_config()  # 保存配置

    def bark_url_changed(self):
        """当 Bark URL 改变时保存配置"""
        self.save_config()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainWindow()
    # 仅在非开机自启动时显示窗口
    if len(sys.argv) <= 1 or sys.argv[1] != '--startup':
        window.show()
    sys.exit(app.exec()) 