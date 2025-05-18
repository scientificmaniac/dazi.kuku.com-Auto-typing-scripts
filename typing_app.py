import sys
import time
import random
import os
from PyQt5.QtWidgets import (QApplication, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
                             QSlider, QPushButton, QGroupBox, QMessageBox, QSpinBox,
                             QRadioButton, QTabWidget, QLineEdit, QCheckBox, QButtonGroup)
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.common.exceptions import WebDriverException
from PyQt5.QtCore import Qt, QThread, pyqtSignal
from selenium.webdriver.chrome.service import Service

class TypingThread(QThread):
    """后台线程用于执行自动输入任务"""
    finished = pyqtSignal(str)  # 任务完成信号
    progress = pyqtSignal(int)  # 进度更新信号
    wait_for_user = pyqtSignal(str)  # 等待用户操作信号，附带提示信息

    def __init__(self, mode, speed, accuracy, test_time=None, test_type=None, url_suffix=None,
                 username=None, password=None, use_wechat_login=False):
        super().__init__()
        self.mode = mode  # 模式：normal或exam
        self.speed = speed
        self.accuracy = accuracy
        self.test_time = test_time  # 测试时间（分钟，普通模式使用）
        self.test_type = test_type  # 测试类型
        self.url_suffix = url_suffix  # 网址后缀
        self.username = username  # 用户名
        self.password = password  # 密码
        self.running = True  # 控制线程运行状态
        self.user_confirmed = False  # 用户确认标志
        self.use_wechat_login = use_wechat_login  # 是否使用微信登录

    def is_browser_closed(self, browser):
        """检查浏览器窗口是否已关闭"""
        try:
            # 尝试获取当前URL，如果浏览器已关闭会抛出异常
            browser.current_url
            return False
        except WebDriverException:
            return True

    def wait_for_confirmation(self, message):
        """等待用户确认"""
        self.wait_for_user.emit(message)
        while self.running and not self.user_confirmed:
            time.sleep(0.1)
        self.user_confirmed = False  # 重置确认标志
        return self.running

    def run(self):
        """线程执行的主函数"""
        backspace_count = 0
        total_chars = 0

        try:
            if getattr(sys, 'frozen', False):
                # 如果是打包后的可执行文件
                base_path = sys._MEIPASS
            else:
                # 如果是开发环境
                base_path = os.path.dirname(os.path.abspath(__file__))

            # 指定chromedriver的路径
            driver_path = os.path.join(base_path, 'chromedriver.exe')
            # 修改为使用 Service 类
            service = Service(driver_path)
            browser = webdriver.Chrome(service=service)

            if self.mode == "normal":
                # 普通模式
                url = 'https://dazi.kukuw.com/'
                browser.get(url)

                if self.use_wechat_login:
                    # 微信登录流程
                    browser.find_element(By.XPATH, '//*[@id="form"]/ul[1]/li[2]/span/span').click()

                    # 等待用户确认微信登录完成
                    if not self.wait_for_confirmation("请使用微信扫描二维码登录，登录完成后点击确认按钮"):
                        self.finished.emit("用户取消了操作")
                        return

                # 选择测试类型
                try:
                    if self.test_type == 'english':
                        browser.find_element(By.XPATH, '//*[@id="radio_en"]').click()
                    elif self.test_type == 'chinese':
                        browser.find_element(By.XPATH, '//*[@id="radio_cn"]').click()
                    time.sleep(0.5)
                except Exception as e:
                    self.finished.emit(f"无法选择测试类型: {str(e)}")
                    return

                # 设置测试时间
                try:
                    time_input = browser.find_element(By.XPATH, '//*[@id="time"]')
                    time_input.clear()
                    time_input.send_keys(str(self.test_time))
                    time.sleep(0.5)
                except Exception as e:
                    self.finished.emit(f"无法设置测试时间: {str(e)}")
                    return

                # 等待用户在浏览器中选择文章
                if not self.wait_for_confirmation("请在浏览器中选择打字文章，完成后点击确认按钮"):
                    self.finished.emit("用户取消了操作")
                    return

                # 点击开始测试按钮
                browser.find_element(By.XPATH, '//*[@id="form"]/ul[6]/li[2]/input').click()

                # 普通模式使用用户设置的时间
                loop_count = int(self.test_time * self.speed * 10)

            elif self.mode == "exam":
                # 考试模式 - 使用HTTP协议
                url = f'http://ks.kukuw.com/{self.url_suffix}'  # 修改为HTTP协议
                browser.get(url)

                # 等待用户处理安全警告
                if not self.wait_for_confirmation("请在浏览器中处理安全警告，完成后点击确认按钮"):
                    self.finished.emit("用户取消了操作")
                    return

                # 填入用户名和密码
                try:
                    username_input = browser.find_element(By.XPATH, '//*[@id="user"]')
                    password_input = browser.find_element(By.XPATH, '//*[@id="pass"]')

                    username_input.send_keys(self.username)
                    time.sleep(0.5)

                    password_input.send_keys(self.password)
                    time.sleep(0.5)

                    # 点击登录按钮
                    login_button = browser.find_element(By.XPATH, '//*[@id="form"]/div[2]/table/tbody/tr[5]/td/input')
                    login_button.click()
                    time.sleep(2)  # 等待登录完成

                except Exception as e:
                    self.finished.emit(f"登录失败: {str(e)}")
                    return

                # 考试模式使用固定的循环次数，不依赖用户设置的时间
                loop_count = int(30 * self.speed * 10)  # 默认30分钟考试时间

            # 计算每个字符的延迟时间
            delay = 60 / self.speed

            # 字符映射表
            error_map = {
                'a': 'qwsz', 'b': 'vghn', 'c': 'xdfv', 'd': 'erfcxs', 'e': 'rdsw',
                'f': 'rtgvcd', 'g': 'tyhbvf', 'h': 'yujnbg', 'i': 'ujko', 'j': 'uikmnh',
                'k': 'ijolm', 'l': 'kop', 'm': 'njk', 'n': 'bhjm', 'o': 'iklp',
                'p': 'ol', 'q': 'wa', 'r': 'edft', 's': 'awedxz', 't': 'rfgy',
                'u': 'yhji', 'v': 'cfgb', 'w': 'qase', 'x': 'zsdc', 'y': 'tghu', 'z': 'asx'
            }

            for i in range(0, loop_count):
                if not self.running or self.is_browser_closed(browser):  # 检查是否需要停止或浏览器已关闭
                    break

                try:
                    # 两种模式使用相同的元素定位
                    element = browser.find_element(By.XPATH, f'//*[@id="i_{i}"]/div/span')

                    if not element.is_displayed():
                        # 如果元素不可见，尝试下一个
                        continue

                    x_v = element.text + ' '
                    total_chars += len(x_v)

                    # 获取输入框元素（两种模式相同）
                    input_element = browser.find_element(By.XPATH, f'//*[@id="i_{i}"]/input[2]')

                    for j in x_v:
                        if not self.running or self.is_browser_closed(browser):  # 检查是否需要停止或浏览器已关闭
                            raise Exception("用户停止输入或浏览器已关闭")

                        # 随机决定是否输入错误
                        if random.randint(1, 100) > self.accuracy and j.lower() in error_map:
                            wrong_chars = error_map[j.lower()]
                            wrong_char = random.choice(wrong_chars)
                            if j.isupper():
                                wrong_char = wrong_char.upper()

                            input_element.send_keys(wrong_char)
                            time.sleep(delay)

                            time.sleep(random.uniform(0.3, 0.8))

                            input_element.send_keys(Keys.BACKSPACE)
                            backspace_count += 1
                            time.sleep(delay)

                        # 输入正确字符
                        input_element.send_keys(j)
                        time.sleep(delay)

                        # 发送进度更新
                        self.progress.emit(int((i / loop_count) * 100))

                except Exception as e:
                    # 如果找不到元素，可能测试已完成或页面结构有变化
                    print(f"处理第{i}个字符时出错: {e}")
                    break

            # 等待测试完成
            time.sleep(2)

            # 计算准确率百分比
            correct_chars = total_chars - backspace_count
            accuracy_percentage = (correct_chars / total_chars) * 100 if total_chars > 0 else 0

            mode_text = "打字测试" if self.mode == "normal" else "考试"
            test_type_text = f"{'英文' if self.test_type == 'english' else '中文'}" if self.mode == "normal" else ""

            # 显示不同的结果信息
            if self.mode == "normal":
                result = (f"模式: {mode_text}{test_type_text}\n"
                          f"时间: {self.test_time} 分钟\n"
                          f"总输入字数: {total_chars}\n"
                          f"退格修正次数: {backspace_count}\n"
                          f"实际准确率: {accuracy_percentage:.2f}%")
            else:
                result = (f"模式: {mode_text}\n"
                          f"总输入字数: {total_chars}\n"
                          f"退格修正次数: {backspace_count}\n"
                          f"实际准确率: {accuracy_percentage:.2f}%")

            browser.quit()
            self.finished.emit(result)

        except Exception as e:
            # 确保浏览器被关闭
            try:
                if 'browser' in locals() and browser:
                    browser.quit()
            except:
                pass
            self.finished.emit(f"发生错误: {str(e)}")


class TypingApp(QWidget):
    def __init__(self):
        super().__init__()
        self.speed = 50  # 默认输入速度 (kpm)
        self.accuracy = 95  # 默认准确率 (%)
        self.test_time = 1  # 默认测试时间 (分钟，普通模式使用)
        self.test_type = 'english'  # 默认测试类型为英文
        self.typing_thread = None  # 后台线程
        self.use_wechat_login = False  # 默认不使用微信登录
        self.initUI()

    def initUI(self):
        # 创建主布局
        main_layout = QVBoxLayout()

        # 创建选项卡控件
        self.tab_widget = QTabWidget()

        # 添加普通模式选项卡
        normal_tab = QWidget()
        self.init_normal_tab(normal_tab)
        self.tab_widget.addTab(normal_tab, "打字测试")

        # 添加考试模式选项卡
        exam_tab = QWidget()
        self.init_exam_tab(exam_tab)
        self.tab_widget.addTab(exam_tab, "考试模式")

        main_layout.addWidget(self.tab_widget)
        self.setLayout(main_layout)

        # 设置窗口属性
        self.setWindowTitle('智能打字训练助手')
        self.setGeometry(300, 300, 600, 500)
        self.show()

    def init_normal_tab(self, parent):
        """初始化普通模式选项卡"""
        layout = QVBoxLayout(parent)

        # 控制面板区域
        control_group = QGroupBox("控制设置")
        control_layout = QHBoxLayout()

        # 速度控制区域
        speed_layout = QVBoxLayout()
        speed_layout.addWidget(QLabel("输入速度控制", self))

        self.speed_slider = QSlider(self)
        self.speed_slider.setOrientation(1)  # 垂直滑块
        self.speed_slider.setMinimum(10)  # 最小速度 10 kpm
        self.speed_slider.setMaximum(200)  # 最大速度 200 kpm
        self.speed_slider.setValue(50)  # 默认速度 50 kpm
        self.speed_slider.valueChanged.connect(self.update_speed)
        speed_layout.addWidget(self.speed_slider)

        self.speed_label = QLabel(f"输入速度: {self.speed} kpm", self)
        speed_layout.addWidget(self.speed_label)

        control_layout.addLayout(speed_layout)

        # 准确率控制区域
        accuracy_layout = QVBoxLayout()
        accuracy_layout.addWidget(QLabel("输入准确率控制", self))

        self.accuracy_slider = QSlider(self)
        self.accuracy_slider.setOrientation(1)  # 垂直滑块
        self.accuracy_slider.setMinimum(50)  # 最低准确率 50%
        self.accuracy_slider.setMaximum(100)  # 最高准确率 100%
        self.accuracy_slider.setValue(95)  # 默认准确率 95%
        self.accuracy_slider.valueChanged.connect(self.update_accuracy)
        accuracy_layout.addWidget(self.accuracy_slider)

        self.accuracy_label = QLabel(f"输入准确率: {self.accuracy}%", self)
        accuracy_layout.addWidget(self.accuracy_label)

        control_layout.addLayout(accuracy_layout)

        # 测试时间设置
        time_layout = QVBoxLayout()
        time_layout.addWidget(QLabel("测试时间设置", self))

        time_input_layout = QHBoxLayout()
        time_input_layout.addWidget(QLabel("测试时间 (分钟):", self))

        self.time_spinbox = QSpinBox(self)
        self.time_spinbox.setMinimum(1)  # 最小测试时间 1 分钟
        self.time_spinbox.setMaximum(50)  # 最大测试时间50分钟
        self.time_spinbox.setValue(1)  # 默认测试时间 1 分钟
        self.time_spinbox.valueChanged.connect(self.update_test_time)
        time_input_layout.addWidget(self.time_spinbox)

        time_layout.addLayout(time_input_layout)

        # 测试类型选择
        test_type_layout = QVBoxLayout()
        test_type_layout.addWidget(QLabel("测试类型选择", self))

        test_type_group = QButtonGroup(self)

        self.english_radio = QRadioButton("英文测试", self)
        self.english_radio.setChecked(True)  # 默认选择英文测试
        self.english_radio.toggled.connect(lambda: self.update_test_type('english'))
        test_type_layout.addWidget(self.english_radio)

        self.chinese_radio = QRadioButton("中文测试", self)
        self.chinese_radio.toggled.connect(lambda: self.update_test_type('chinese'))
        test_type_layout.addWidget(self.chinese_radio)

        # 微信登录选择
        self.wechat_login_checkbox = QCheckBox("使用微信登录", self)
        self.wechat_login_checkbox.toggled.connect(self.update_wechat_login)
        test_type_layout.addWidget(self.wechat_login_checkbox)

        time_layout.addLayout(test_type_layout)
        control_layout.addLayout(time_layout)

        control_group.setLayout(control_layout)
        layout.addWidget(control_group)

        # 按钮区域
        button_layout = QHBoxLayout()

        # 开始按钮
        self.start_button = QPushButton('开始自动输入', self)
        self.start_button.setMinimumHeight(40)
        self.start_button.setStyleSheet("background-color: #4CAF50; color: white; font-size: 16px;")
        self.start_button.clicked.connect(lambda: self.start_typing("normal"))
        button_layout.addWidget(self.start_button)

        # 确认按钮（用于普通模式）
        self.confirm_button = QPushButton('确认已完成操作', self)
        self.confirm_button.setMinimumHeight(40)
        self.confirm_button.setStyleSheet("background-color: #FF9800; color: white; font-size: 16px;")
        self.confirm_button.setEnabled(False)  # 初始禁用
        self.confirm_button.clicked.connect(self.confirm_article_selection)
        button_layout.addWidget(self.confirm_button)

        # 停止按钮
        self.stop_button = QPushButton('停止输入', self)
        self.stop_button.setMinimumHeight(40)
        self.stop_button.setStyleSheet("background-color: #f44336; color: white; font-size: 16px;")
        self.stop_button.setEnabled(False)  # 初始禁用
        self.stop_button.clicked.connect(self.stop_typing)
        button_layout.addWidget(self.stop_button)

        layout.addLayout(button_layout)

        # 状态标签
        self.status_label = QLabel("准备就绪", self)
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-size: 14px; color: #666;")
        layout.addWidget(self.status_label)

        parent.setLayout(layout)

    def init_exam_tab(self, parent):
        """初始化考试模式选项卡"""
        layout = QVBoxLayout(parent)

        # 考试设置区域
        exam_group = QGroupBox("考试设置")
        exam_layout = QVBoxLayout()

        # 网址后缀
        url_layout = QHBoxLayout()
        url_layout.addWidget(QLabel("网址后缀 (http://ks.kukuw.com/", self))  # 修改提示文本
        self.url_suffix = QLineEdit(self)
        self.url_suffix.setPlaceholderText("例如：123456")
        url_layout.addWidget(self.url_suffix)
        url_layout.addWidget(QLabel(")", self))
        exam_layout.addLayout(url_layout)

        # 用户名和密码
        username_layout = QHBoxLayout()
        username_layout.addWidget(QLabel("用户名:", self))
        self.username = QLineEdit(self)
        username_layout.addWidget(self.username)
        exam_layout.addLayout(username_layout)

        password_layout = QHBoxLayout()
        password_layout.addWidget(QLabel("密码:", self))
        self.password = QLineEdit(self)
        self.password.setEchoMode(QLineEdit.Password)
        password_layout.addWidget(self.password)
        exam_layout.addLayout(password_layout)

        # 速度和准确率控制
        speed_accuracy_layout = QHBoxLayout()

        # 速度控制
        speed_layout = QVBoxLayout()
        speed_layout.addWidget(QLabel("输入速度", self))

        self.exam_speed_slider = QSlider(self)
        self.exam_speed_slider.setOrientation(1)  # 垂直滑块
        self.exam_speed_slider.setMinimum(10)  # 最小速度 10 kpm
        self.exam_speed_slider.setMaximum(200)  # 最大速度 200 kpm
        self.exam_speed_slider.setValue(50)  # 默认速度 50 kpm
        self.exam_speed_slider.valueChanged.connect(self.update_exam_speed)
        speed_layout.addWidget(self.exam_speed_slider)

        self.exam_speed_label = QLabel(f"速度: {self.speed} kpm", self)
        speed_layout.addWidget(self.exam_speed_label)

        speed_accuracy_layout.addLayout(speed_layout)

        # 准确率控制
        accuracy_layout = QVBoxLayout()
        accuracy_layout.addWidget(QLabel("输入准确率", self))

        self.exam_accuracy_slider = QSlider(self)
        self.exam_accuracy_slider.setOrientation(1)  # 垂直滑块
        self.exam_accuracy_slider.setMinimum(50)  # 最低准确率 50%
        self.exam_accuracy_slider.setMaximum(100)  # 最高准确率 100%
        self.exam_accuracy_slider.setValue(95)  # 默认准确率 95%
        self.exam_accuracy_slider.valueChanged.connect(self.update_exam_accuracy)
        accuracy_layout.addWidget(self.exam_accuracy_slider)

        self.exam_accuracy_label = QLabel(f"准确率: {self.accuracy}%", self)
        accuracy_layout.addWidget(self.exam_accuracy_label)

        speed_accuracy_layout.addLayout(accuracy_layout)

        # 移除考试时间设置，考试模式不需要用户设置时间

        exam_layout.addLayout(speed_accuracy_layout)

        exam_group.setLayout(exam_layout)
        layout.addWidget(exam_group)

        # 按钮区域
        button_layout = QHBoxLayout()

        # 开始考试按钮
        self.start_exam_button = QPushButton('开始考试', self)
        self.start_exam_button.setMinimumHeight(40)
        self.start_exam_button.setStyleSheet("background-color: #2196F3; color: white; font-size: 16px;")
        self.start_exam_button.clicked.connect(lambda: self.start_typing("exam"))
        button_layout.addWidget(self.start_exam_button)

        # 确认按钮（用于考试模式）
        self.exam_confirm_button = QPushButton('已处理安全警告，继续', self)
        self.exam_confirm_button.setMinimumHeight(40)
        self.exam_confirm_button.setStyleSheet("background-color: #FF9800; color: white; font-size: 16px;")
        self.exam_confirm_button.setEnabled(False)  # 初始禁用
        self.exam_confirm_button.clicked.connect(self.confirm_article_selection)
        button_layout.addWidget(self.exam_confirm_button)

        # 停止按钮
        self.stop_exam_button = QPushButton('停止考试', self)
        self.stop_exam_button.setMinimumHeight(40)
        self.stop_exam_button.setStyleSheet("background-color: #f44336; color: white; font-size: 16px;")
        self.stop_exam_button.setEnabled(False)  # 初始禁用
        self.stop_exam_button.clicked.connect(self.stop_typing)
        button_layout.addWidget(self.stop_exam_button)

        layout.addLayout(button_layout)

        # 状态标签
        self.exam_status_label = QLabel("准备就绪", self)
        self.exam_status_label.setAlignment(Qt.AlignCenter)
        self.exam_status_label.setStyleSheet("font-size: 14px; color: #666;")
        layout.addWidget(self.exam_status_label)

        parent.setLayout(layout)

    def update_speed(self, value):
        self.speed = value
        self.speed_label.setText(f"输入速度: {self.speed} kpm")
        self.exam_speed_label.setText(f"速度: {self.speed} kpm")

    def update_accuracy(self, value):
        self.accuracy = value
        self.accuracy_label.setText(f"输入准确率: {self.accuracy}%")
        self.exam_accuracy_label.setText(f"准确率: {self.accuracy}%")

    def update_test_time(self, value):
        self.test_time = value

    def update_test_type(self, test_type):
        """更新测试类型"""
        self.test_type = test_type

    def update_exam_speed(self, value):
        self.speed = value
        self.exam_speed_label.setText(f"速度: {self.speed} kpm")

    def update_exam_accuracy(self, value):
        self.accuracy = value
        self.exam_accuracy_label.setText(f"准确率: {self.accuracy}%")

    def update_exam_time(self, value):
        # 考试模式不再使用此函数
        pass

    def update_wechat_login(self, state):
        """更新微信登录选项"""
        self.use_wechat_login = state

    def start_typing(self, mode):
        """开始自动输入或考试"""
        if mode == "normal":
            self.start_button.setEnabled(False)
            self.stop_button.setEnabled(True)
            self.status_label.setText("正在准备...")
            self.confirm_button.setEnabled(False)

            # 创建并启动后台线程（普通模式）
            self.typing_thread = TypingThread(
                mode=mode,
                speed=self.speed,
                accuracy=self.accuracy,
                test_time=self.test_time,
                test_type=self.test_type,
                use_wechat_login=self.use_wechat_login
            )
            self.typing_thread.wait_for_user.connect(self.on_wait_for_user)
            self.typing_thread.finished.connect(self.on_typing_finished)
            self.typing_thread.start()

        else:  # exam
            # 检查考试设置
            if not self.url_suffix.text().strip():
                QMessageBox.warning(self, "警告", "请输入网址后缀")
                return

            if not self.username.text().strip():
                QMessageBox.warning(self, "警告", "请输入用户名")
                return

            if not self.password.text().strip():
                QMessageBox.warning(self, "警告", "请输入密码")
                return

            self.start_exam_button.setEnabled(False)
            self.stop_exam_button.setEnabled(True)
            self.exam_confirm_button.setEnabled(True)  # 启用确认按钮
            self.exam_status_label.setText("请处理浏览器中的安全警告，完成后点击确认按钮")

            # 创建并启动后台线程（考试模式）
            self.typing_thread = TypingThread(
                mode=mode,
                speed=self.speed,
                accuracy=self.accuracy,
                url_suffix=self.url_suffix.text().strip(),
                username=self.username.text().strip(),
                password=self.password.text().strip()
            )
            self.typing_thread.wait_for_user.connect(self.on_wait_for_user)
            self.typing_thread.finished.connect(self.on_typing_finished)
            self.typing_thread.start()

    def confirm_article_selection(self):
        """确认文章选择或微信登录完成或安全警告处理完成"""
        if self.typing_thread:
            self.typing_thread.user_confirmed = True
            # 根据当前选项卡禁用相应的确认按钮
            if self.tab_widget.currentIndex() == 0:  # 普通模式
                self.confirm_button.setEnabled(False)
            else:  # 考试模式
                self.exam_confirm_button.setEnabled(False)

    def on_wait_for_user(self, message):
        """处理等待用户操作信号"""
        # 根据当前选项卡更新相应的状态标签
        if self.tab_widget.currentIndex() == 0:  # 普通模式
            self.status_label.setText(message)
            self.confirm_button.setEnabled(True)
        else:  # 考试模式
            self.exam_status_label.setText(message)
            self.exam_confirm_button.setEnabled(True)

    def stop_typing(self):
        """停止自动输入或考试"""
        if self.typing_thread and self.typing_thread.isRunning():
            self.typing_thread.running = False  # 设置标志位停止线程
            self.typing_thread.wait(2000)  # 等待最多2秒

            # 无论线程是否成功停止，都恢复UI状态
            self.on_typing_finished("操作已停止")

    def on_typing_finished(self, result):
        """处理输入完成事件"""
        self.start_button.setEnabled(True)
        self.stop_button.setEnabled(False)
        self.confirm_button.setEnabled(False)
        self.start_exam_button.setEnabled(True)
        self.stop_exam_button.setEnabled(False)
        self.exam_confirm_button.setEnabled(False)

        # 根据当前选项卡更新相应的状态标签
        if self.tab_widget.currentIndex() == 0:  # 普通模式
            self.status_label.setText("准备就绪")
        else:  # 考试模式
            self.exam_status_label.setText("准备就绪")

        # 显示结果
        QMessageBox.information(self, "操作完成", result)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = TypingApp()
    sys.exit(app.exec_())
