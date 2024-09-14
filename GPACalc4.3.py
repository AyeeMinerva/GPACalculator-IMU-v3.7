import os
import sys
import time
import configparser
from selenium import webdriver
from selenium.webdriver.edge.service import Service
import ddddocr
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout, QPushButton,
    QLabel, QTableWidget, QTableWidgetItem, QCheckBox, QMessageBox, QFileDialog
)
from PyQt5.QtCore import Qt
from datetime import datetime

LOGIN_URL = "https://jwxt.imu.edu.cn/login"
GRADES_URL = "https://jwxt.imu.edu.cn/student/integratedQuery/scoreQuery/allPassingScores/index?mobile=false"
COURSES_FILE = "courses.txt"
IGNORE_FILE = "ignore_words.txt"
CREDENTIALS_FILE = "credentials.txt"
MAX_REFRESH_TIMES = 4
DRIVER_PATH =  "msedgedriver.exe"

USE_DIRECT_POINT = False


class Lesson:
    def __init__(self, name: str, times: int, marks: float, point: float = 0.0):
        self.name = name
        self.marks = marks
        self.times = times
        self.point = point if USE_DIRECT_POINT else self.calculate_point(marks)

    def calculate_point(self, marks: float):
        if 90 <= marks <= 100:
            return 4.0
        elif 85 <= marks <= 89:
            return 3.7
        elif 82 <= marks <= 84:
            return 3.3
        elif 78 <= marks <= 81:
            return 3.0
        elif 75 <= marks <= 77:
            return 2.7
        elif 72 <= marks <= 74:
            return 2.3
        elif 68 <= marks <= 71:
            return 2.0
        elif 65 <= marks <= 67:
            return 1.7
        elif 62 <= marks <= 64:
            return 1.3
        elif 60 <= marks <= 61:
            return 1.0
        else:
            return 0.0

    def calculate_marks(self, point: float):
        if point == 4.0:
            return 90
        elif point == 3.7:
            return 85
        elif point == 3.3:
            return 82
        elif point == 3.0:
            return 78
        elif point == 2.7:
            return 75
        elif point == 2.3:
            return 72
        elif point == 2.0:
            return 68
        elif point == 1.7:
            return 65
        elif point == 1.3:
            return 62
        elif point == 1.0:
            return 60
        else:
            return 0


class GPAApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("GPA Calculator")
        self.setGeometry(100, 100, 700, 500)

        self.courses = []
        self.modified = False
        self.driver = None
        self.ocr = ddddocr.DdddOcr(beta=True)

        main_layout = QVBoxLayout()
        control_layout = QHBoxLayout()
        self.config_layout = QVBoxLayout()
        display_layout = QHBoxLayout()

        self.load_button = QPushButton("从本地配置读入")
        self.scrape_button = QPushButton("爬取到本地")
        self.refresh_button = QPushButton("刷新")
        self.export_button = QPushButton("导出到日志")
        self.switch_point_method = QCheckBox("使用绩点")
        self.switch_point_method.setChecked(False)
        self.default_load_checkbox = QCheckBox("默认读取相对路径的 courses.txt")
        self.default_load_checkbox.setChecked(True)
        self.add_course_button = QPushButton("添加课程")
        self.remove_course_button = QPushButton("删除选中课程")

        self.load_button.clicked.connect(self.load_from_file)
        self.scrape_button.clicked.connect(self.start_scrape)
        self.refresh_button.clicked.connect(self.refresh_gpa)
        self.export_button.clicked.connect(self.export_to_log)
        self.switch_point_method.stateChanged.connect(self.toggle_point_method)
        self.add_course_button.clicked.connect(self.add_course)
        self.remove_course_button.clicked.connect(self.remove_course)

        control_layout.addWidget(self.load_button)
        control_layout.addWidget(self.scrape_button)
        control_layout.addWidget(self.refresh_button)
        control_layout.addWidget(self.export_button)
        control_layout.addWidget(self.switch_point_method)
        control_layout.addWidget(self.default_load_checkbox)
        control_layout.addWidget(self.add_course_button)
        control_layout.addWidget(self.remove_course_button)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["课程名", "学分", "成绩", "绩点", "影响"])
        self.config_layout.addWidget(self.table)

        self.table.cellChanged.connect(self.handle_cell_changed)

        self.gpa_label = QLabel("GPA: 0.00 (0.00000)")
        self.total_credits_label = QLabel("总学分: 0")
        display_layout.addWidget(self.gpa_label)
        display_layout.addWidget(self.total_credits_label)

        main_layout.addLayout(control_layout)
        main_layout.addLayout(self.config_layout)
        main_layout.addLayout(display_layout)

        self.setLayout(main_layout)

        if self.default_load_checkbox.isChecked():
            self.courses = self.read_courses_from_file(COURSES_FILE)
            self.update_table()

    def handle_cell_changed(self, row, column):
        if column == 2:  # 修改成绩时更新绩点
            marks_item = self.table.item(row, 2)
            if marks_item:
                marks = float(marks_item.text())
                point = Lesson("", 0, marks).calculate_point(marks)
                self.table.blockSignals(True)
                self.table.setItem(row, 3, QTableWidgetItem(str(point)))
                self.table.blockSignals(False)
        elif column == 3:  # 修改绩点时更新成绩
            point_item = self.table.item(row, 3)
            if point_item:
                point = float(point_item.text())
                marks = Lesson("", 0, 0).calculate_marks(point)
                self.table.blockSignals(True)
                self.table.setItem(row, 2, QTableWidgetItem(str(marks)))
                self.table.blockSignals(False)

        self.refresh_gpa()

    def load_from_file(self):
        if self.default_load_checkbox.isChecked():
            filename = COURSES_FILE
        else:
            filename, _ = QFileDialog.getOpenFileName(self, "选择课程配置文件", "", "Text Files (*.txt)")
        if filename:
            self.courses = self.read_courses_from_file(filename)
            self.update_table()

    def read_courses_from_file(self, filename):
        courses = []
        ignore_words = self.load_ignore_words(IGNORE_FILE)
        try:
            with open(filename, 'r', encoding='utf-8') as file:
                for line in file:
                    line = line.strip()
                    if not line or any(word in line for word in ignore_words):
                        continue

                    parts = line.split()
                    if len(parts) < 5:
                        continue

                    course_name = parts[0].strip().replace(" ", "-")  # 替换空格为连字符
                    course_attribute = parts[1].strip()
                    times = int(parts[2].strip())
                    marks = float(parts[3].strip())
                    points = float(parts[4].strip())
                    courses.append(Lesson(course_name, times, marks, points))

        except FileNotFoundError:
            QMessageBox.warning(self, "错误", "无法打开文件！")
        
        return courses

    def load_ignore_words(self, filename):
        ignore_words = set()
        if os.path.exists(filename):
            with open(filename, 'r', encoding='utf-8') as f:
                for line in f:
                    ignore_words.add(line.strip())
        return ignore_words

    def update_table(self):
        self.table.setRowCount(0)
        for course in self.courses:
            row_position = self.table.rowCount()
            self.table.insertRow(row_position)
            self.table.setItem(row_position, 0, QTableWidgetItem(course.name))
            self.table.setItem(row_position, 1, QTableWidgetItem(str(course.times)))
            self.table.setItem(row_position, 2, QTableWidgetItem(str(course.marks)))
            self.table.setItem(row_position, 3, QTableWidgetItem(str(course.point)))
            affect_item = QTableWidgetItem("0.00")
            affect_item.setFlags(affect_item.flags() & ~Qt.ItemIsEditable)
            self.table.setItem(row_position, 4, affect_item)

        self.refresh_gpa()

    def refresh_gpa(self):
        total_credits = 0
        gp_sum = 0
        for row in range(self.table.rowCount()):
            credits_item = self.table.item(row, 1)
            point_item = self.table.item(row, 3)
            if credits_item is None or point_item is None:
                continue

            try:
                credits = int(credits_item.text())
                point = float(point_item.text())
            except ValueError:
                continue

            total_credits += credits
            gp_sum += credits * point

        gpa = gp_sum / total_credits if total_credits > 0 else 0
        self.gpa_label.setText(f"GPA: {gpa:.2f} ({gpa:.5f})")
        self.total_credits_label.setText(f"总学分: {total_credits}")

        for row in range(self.table.rowCount()):
            current_point_item = self.table.item(row, 3)
            current_credits_item = self.table.item(row, 1)
            if current_point_item is None or current_credits_item is None:
                continue

            try:
                current_point = float(current_point_item.text())
                current_credits = int(current_credits_item.text())
            except ValueError:
                continue

            if total_credits - current_credits > 0:
                gpa_without = (gp_sum - current_point * current_credits) / (total_credits - current_credits)
                influence = gpa - gpa_without
            else:
                influence = 0.0

            influence_item = self.table.item(row, 4)
            if influence_item:
                influence_item.setText(f"{influence:.2f}")

    def toggle_point_method(self, state):
        global USE_DIRECT_POINT
        USE_DIRECT_POINT = state == Qt.Checked
        self.refresh_gpa()

    def add_course(self):
        row_position = self.table.rowCount()
        self.table.insertRow(row_position)
        self.table.setItem(row_position, 0, QTableWidgetItem("新课程"))
        self.table.setItem(row_position, 1, QTableWidgetItem("0"))
        self.table.setItem(row_position, 2, QTableWidgetItem("0.0"))
        self.table.setItem(row_position, 3, QTableWidgetItem("0.0"))
        self.table.setItem(row_position, 4, QTableWidgetItem("0.00"))
        self.refresh_gpa()

    def remove_course(self):
        selected_rows = self.table.selectionModel().selectedRows()
        for selected in selected_rows:
            self.table.removeRow(selected.row())
        self.refresh_gpa()

    def export_to_log(self):
        filename = datetime.now().strftime("%Y%m%d_%H%M%S.log")
        try:
            with open(filename, 'w') as file:
                with open(COURSES_FILE, 'r', encoding='utf-8') as courses_file:
                    file.write(courses_file.read())
                    file.write("\n")

                file.write(f"GPA: {self.gpa_label.text()}\n")
                file.write(f"总学分: {self.total_credits_label.text()}\n")
                file.write("课程列表:\n")
                for row in range(self.table.rowCount()):
                    name = self.table.item(row, 0).text()
                    credits = self.table.item(row, 1).text()
                    marks = self.table.item(row, 2).text()
                    point = self.table.item(row, 3).text()
                    influence = self.table.item(row, 4).text()
                    file.write(f"{name}, 学分: {credits}, 成绩: {marks}, 绩点: {point}, 影响: {influence}\n")

            QMessageBox.information(self, "成功", f"GPA日志已保存为 {filename}")
        except Exception as e:
            QMessageBox.warning(self, "错误", f"保存日志失败: {e}")

    def start_scrape(self):
        try:
            username, password = self.prompt_login_credentials()
            self.total(username, password)
        except Exception as e:
            QMessageBox.warning(self, "错误", f"爬取失败: {e}")

    def prompt_login_credentials(self):
        if os.path.exists(CREDENTIALS_FILE):
            try:
                with open(CREDENTIALS_FILE, 'r', encoding='utf-8') as file:
                    lines = file.readlines()
                    if len(lines) >= 2:
                        username = lines[0].strip()
                        password = lines[1].strip()
                        return username, password
                    else:
                        print("账号或密码格式不正确")
                        return None, None
            except Exception as e:
                print(f"读取账号密码文件时出错: {e}")
                return None, None
        else:
            print(f"找不到账号密码文件 {CREDENTIALS_FILE}")
            return None, None

    def total(self, username, password):
        if self.driver is None:
            try:
                service = Service(executable_path=DRIVER_PATH)
                self.driver = webdriver.Edge(service=service)
            except Exception as e:
                print(e)
                QMessageBox.warning(self, "警告", "无法启动程序自带驱动浏览器，正在尝试启动本地Edge WebDriver:")
                try:
                    self.driver = webdriver.Edge()
                except Exception as e:
                    print(f"无法启动Edge WebDriver: {e}, 请安装最新版Edge WebDriver")
                    return
               
        self.input_username_password(username, password)
        self.input_captcha(0)
        if not self.click_login_button():
            with open(COURSES_FILE, "a", encoding='utf-8') as file:
                file.write("账号或密码错误\n")
            return
        self.get_scores(0, username, password)

    def input_username_password(self, username, password):
        self.driver.get(LOGIN_URL)
        username_input = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "input_username"))
        )
        username_input.send_keys(username)

        password_input = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "input_password"))
        )
        password_input.send_keys(password)

    def input_captcha(self, attempts):
        if attempts >= MAX_REFRESH_TIMES:
            raise Exception("验证码识别失败，超过最大重试次数")

        captcha_code = self.refresh_captcha()
        if captcha_code and len(captcha_code) == 4:
            captcha_input = WebDriverWait(self.driver, 10).until(
                EC.visibility_of_element_located((By.ID, "input_checkcode"))
            )
            captcha_input.send_keys(captcha_code)
        else:
            self.input_captcha(attempts + 1)

    def refresh_captcha(self):
        self.driver.find_element(By.ID, "captchaImg").click()
        captcha_element = WebDriverWait(self.driver, 10).until(
            EC.visibility_of_element_located((By.ID, "captchaImg"))
        )
        time.sleep(1)
        captcha_image = captcha_element.screenshot_as_png
        return self.ocr.classification(captcha_image)

    def click_login_button(self):
        time.sleep(0.5)
        login_button = WebDriverWait(self.driver, 10).until(
            EC.element_to_be_clickable((By.ID, "loginButton"))
        )
        login_button.click()
        return not ('errorCode=badCredentials' in self.driver.current_url)

    def get_scores(self, refresh_times, username, password):
        if refresh_times > MAX_REFRESH_TIMES:
            with open(COURSES_FILE, "a", encoding='utf-8') as file:
                file.write("您可能没有成绩，或者教务系统崩溃，请进入教务系统确认\n")
            return

        try:
            self.driver.get(GRADES_URL)
            time.sleep(1)
            tabs = self.driver.find_elements(By.CSS_SELECTOR, '[id^="tab"]')
            currentPageUrl = self.driver.current_url

            if not tabs and refresh_times <= MAX_REFRESH_TIMES:
                if currentPageUrl == GRADES_URL:
                    self.driver.refresh()
                else:
                    self.total(username, password)
                self.get_scores(refresh_times + 1, username, password)
            else:
                with open(COURSES_FILE, 'w', encoding='utf-8') as file:
                    for tab in tabs:
                        tab_name = tab.find_element(By.TAG_NAME, 'h4').text
                        file.write(f"{tab_name}\n")

                        table = tab.find_element(By.TAG_NAME, 'table')
                        rows = table.find_elements(By.TAG_NAME, 'tr')

                        for row in rows:
                            cells = row.find_elements(By.TAG_NAME, 'td')
                            course_info = []
                            for cell in cells:
                                text = cell.text.strip()
                                if text:
                                    course_info.append(text)
                            if len(course_info) == 9:
                                no, course_number, course_sequence_number, course_name, course_attribute, credits, grades, grade_points, english_course_name = course_info
                                course_name = course_name.replace(" ", "-")  # 替换空格为连字符
                                line = f"{course_name} {course_attribute} {credits} {grades} {grade_points}\n"
                                file.write(line)

        except Exception as e:
            print(f"获取成绩出现错误，错误码: {e}")
        
        finally:
            self.driver.quit()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = GPAApp()
    window.show()
    sys.exit(app.exec_())
