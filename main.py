import sys
import time
import random
import os
import ctypes

# 解决 macOS 上 Qt 平台插件 "cocoa" 未找到的问题：
try:
    from PyQt5.QtCore import QLibraryInfo
    _plugins_path = QLibraryInfo.location(QLibraryInfo.PluginsPath) or ""
    _libs_path = QLibraryInfo.location(QLibraryInfo.LibrariesPath) or ""
    if _plugins_path:
        os.environ.setdefault("QT_QPA_PLATFORM_PLUGIN_PATH", _plugins_path)
        os.environ.setdefault("QT_PLUGIN_PATH", _plugins_path)
    if _libs_path:
        os.environ.setdefault("DYLD_FRAMEWORK_PATH", _libs_path)
        os.environ.setdefault("DYLD_LIBRARY_PATH", _libs_path)
except Exception:
    pass

from PyQt5.QtWidgets import (
    QApplication,
    QWidget,
    QVBoxLayout,
    QHBoxLayout,
    QLabel,
    QTextEdit,
    QDoubleSpinBox,
    QSpinBox,
    QPushButton,
    QComboBox
)
from PyQt5.QtCore import Qt, QThread, pyqtSignal

# 尝试导入pynput
class KeyboardController:
    def __init__(self):
        self.method = "none"
        self.controller = None
        self._init_controllers()
    
    def _init_controllers(self):
        # 尝试初始化各种输入方法
        try:
            from pynput.keyboard import Controller as PynputController, Key
            self.controller = PynputController()
            self.Key = Key
            self.method = "pynput"
            return
        except ImportError:
            pass
        
        # 尝试Windows API
        try:
            self._init_winapi()
            self.method = "winapi"
            return
        except Exception:
            pass
    
    def _init_winapi(self):
        """初始化Windows API控制器"""
        self.user32 = ctypes.WinDLL('user32', use_last_error=True)
        self.INPUT_KEYBOARD = 1
        self.KEYEVENTF_KEYUP = 0x0002
        self.KEYEVENTF_UNICODE = 0x0004
        self.VK_RETURN = 0x0D
        
        # 简化的INPUT结构
        class INPUT(ctypes.Structure):
            _fields_ = [
                ('type', ctypes.c_ulong),
                ('ki', ctypes.c_ubyte * 24),  # 足够大的空间
            ]
        
        class KEYBDINPUT(ctypes.Structure):
            _fields_ = [
                ('wVk', ctypes.c_ushort),
                ('wScan', ctypes.c_ushort),
                ('dwFlags', ctypes.c_ulong),
                ('time', ctypes.c_ulong),
                ('dwExtraInfo', ctypes.c_void_p),
            ]
        
        self.INPUT = INPUT
        self.KEYBDINPUT = KEYBDINPUT
        self.user32.SendInput.argtypes = (
            ctypes.c_uint,
            ctypes.POINTER(INPUT),
            ctypes.c_int
        )
        self.user32.SendInput.restype = ctypes.c_uint
    
    def type(self, text):
        """输入文本"""
        if self.method == "pynput":
            try:
                self.controller.type(text)
                return True
            except Exception:
                return False
        elif self.method == "winapi":
            try:
                for char in text:
                    if char == '\n':
                        self._press_key(self.VK_RETURN)
                    else:
                        self._send_unicode(char)
                return True
            except Exception:
                return False
        return False
    
    def _press_key(self, vk_code):
        """按下并释放按键"""
        inp = self.INPUT(type=self.INPUT_KEYBOARD)
        ki = self.KEYBDINPUT(
            wVk=vk_code,
            wScan=0,
            dwFlags=0,
            time=0,
            dwExtraInfo=None
        )
        ctypes.memmove(inp.ki, ctypes.addressof(ki), ctypes.sizeof(ki))
        self.user32.SendInput(1, ctypes.pointer(inp), ctypes.sizeof(inp))
        
        # 释放按键
        ki.dwFlags = self.KEYEVENTF_KEYUP
        ctypes.memmove(inp.ki, ctypes.addressof(ki), ctypes.sizeof(ki))
        self.user32.SendInput(1, ctypes.pointer(inp), ctypes.sizeof(inp))
    
    def _send_unicode(self, char):
        """发送Unicode字符"""
        inp = self.INPUT(type=self.INPUT_KEYBOARD)
        ki = self.KEYBDINPUT(
            wVk=0,
            wScan=ord(char),
            dwFlags=self.KEYEVENTF_UNICODE,
            time=0,
            dwExtraInfo=None
        )
        ctypes.memmove(inp.ki, ctypes.addressof(ki), ctypes.sizeof(ki))
        self.user32.SendInput(1, ctypes.pointer(inp), ctypes.sizeof(inp))
        
        # 释放
        ki.dwFlags = self.KEYEVENTF_UNICODE | self.KEYEVENTF_KEYUP
        ctypes.memmove(inp.ki, ctypes.addressof(ki), ctypes.sizeof(ki))
        self.user32.SendInput(1, ctypes.pointer(inp), ctypes.sizeof(inp))

class TypingWorker(QThread):
    progress = pyqtSignal(int)
    result = pyqtSignal(bool, str)

    def __init__(self, text, cps, start_delay_ms, jitter_ms, input_method):
        super().__init__()
        self.text = text
        self.cps = max(0.1, float(cps))
        self.start_delay_ms = max(0, int(start_delay_ms))
        self.jitter_ms = max(0, int(jitter_ms))
        self.input_method = input_method
        self._stopping = False
        self.controller = KeyboardController()

    def stop(self):
        self._stopping = True

    def run(self):
        try:
            # 开始前延迟
            if self.start_delay_ms > 0:
                slept = 0
                step = 50
                while slept < self.start_delay_ms and not self._stopping:
                    time.sleep(step / 1000.0)
                    slept += step
            
            if self._stopping:
                self.result.emit(False, "用户中止")
                return
            
            # 输入文本
            count = 0
            base_delay = 1.0 / self.cps
            
            for i, ch in enumerate(self.text):
                if self._stopping:
                    self.result.emit(False, "用户中止")
                    return
                
                # 输入字符
                success = self.controller.type(ch)
                if not success:
                    self.result.emit(False, "输入失败，请检查权限")
                    return
                
                count += 1
                self.progress.emit(count)
                
                # 随机抖动
                jitter = 0.0
                if self.jitter_ms > 0:
                    jitter = random.uniform(-self.jitter_ms, self.jitter_ms) / 1000.0
                delay = max(0.0, base_delay + jitter)
                time.sleep(delay)
            
            self.result.emit(True, "已完成输入")
        except Exception as e:
            self.result.emit(False, f"发生错误: {str(e)}")

class TypingSimulator(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("打字输入模拟器")
        self.resize(520, 450)
        self.worker = None

        layout = QVBoxLayout()

        # 提示信息
        self.hint = QLabel(
            "将光标置于目标输入框后，点击“开始模拟”。\n"
            "支持本地和虚拟机环境的输入模拟。"
        )
        self.hint.setWordWrap(True)
        layout.addWidget(self.hint)

        # 文本输入
        layout.addWidget(QLabel("要输入的文本："))
        self.text_edit = QTextEdit()
        self.text_edit.setPlaceholderText("在此粘贴文本（将被模拟为逐字输入）")
        layout.addWidget(self.text_edit)

        # 控制选项
        controls = QHBoxLayout()
        
        # 速度控制
        cps_box = QVBoxLayout()
        cps_label = QLabel("速度（字符/秒）：")
        self.cps = QDoubleSpinBox()
        self.cps.setRange(0.1, 50.0)
        self.cps.setSingleStep(0.5)
        self.cps.setValue(10.0)
        cps_box.addWidget(cps_label)
        cps_box.addWidget(self.cps)

        # 延迟控制
        start_delay_box = QVBoxLayout()
        start_delay_label = QLabel("开始前延迟（毫秒）：")
        self.start_delay = QSpinBox()
        self.start_delay.setRange(0, 10000)
        self.start_delay.setSingleStep(100)
        self.start_delay.setValue(2000)
        start_delay_box.addWidget(start_delay_label)
        start_delay_box.addWidget(self.start_delay)

        # 抖动控制
        jitter_box = QVBoxLayout()
        jitter_label = QLabel("随机抖动（±毫秒）：")
        self.jitter = QSpinBox()
        self.jitter.setRange(0, 500)
        self.jitter.setSingleStep(10)
        self.jitter.setValue(50)
        jitter_box.addWidget(jitter_label)
        jitter_box.addWidget(self.jitter)

        controls.addLayout(cps_box)
        controls.addLayout(start_delay_box)
        controls.addLayout(jitter_box)
        layout.addLayout(controls)

        # 输入方法选择
        method_box = QHBoxLayout()
        method_label = QLabel("输入方法：")
        self.method_combo = QComboBox()
        self.method_combo.addItems(["自动选择", "Pynput", "Windows API"])
        method_box.addWidget(method_label)
        method_box.addWidget(self.method_combo)
        layout.addLayout(method_box)

        # 进度显示
        self.progress_label = QLabel("进度：0 / 0")
        layout.addWidget(self.progress_label)

        # 按钮
        buttons = QHBoxLayout()
        self.start_btn = QPushButton("开始模拟")
        self.start_btn.clicked.connect(self.on_start)
        self.stop_btn = QPushButton("停止")
        self.stop_btn.clicked.connect(self.on_stop)
        self.stop_btn.setEnabled(False)
        buttons.addWidget(self.start_btn)
        buttons.addWidget(self.stop_btn)
        layout.addLayout(buttons)

        # 状态信息
        self.status_label = QLabel("就绪")
        self.status_label.setStyleSheet("color: #666")
        layout.addWidget(self.status_label)

        self.setLayout(layout)

    def on_start(self):
        text = self.text_edit.toPlainText()
        if not text:
            self.progress_label.setText("进度：0 / 0（请输入文本）")
            return
        
        self.progress_label.setText(f"进度：0 / {len(text)}")
        self.start_btn.setEnabled(False)
        self.stop_btn.setEnabled(True)
        self.status_label.setText("准备中...")

        input_method = self.method_combo.currentText()
        
        self.worker = TypingWorker(
            text=text,
            cps=self.cps.value(),
            start_delay_ms=self.start_delay.value(),
            jitter_ms=self.jitter.value(),
            input_method=input_method
        )
        self.worker.progress.connect(self.on_progress)
        self.worker.result.connect(self.on_result)
        self.worker.finished.connect(self.on_thread_finished)
        self.worker.start()

    def on_stop(self):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()
        self.stop_btn.setEnabled(False)
        self.start_btn.setEnabled(True)
        self.status_label.setText("已停止")

    def on_progress(self, count):
        total = len(self.text_edit.toPlainText())
        self.progress_label.setText(f"进度：{count} / {total}")

    def on_result(self, ok, msg):
        self.progress_label.setText(msg)
        self.status_label.setText("完成" if ok else "错误")
        if ok:
            self.stop_btn.setEnabled(False)
            self.start_btn.setEnabled(True)

    def on_thread_finished(self):
        if self.worker:
            try:
                self.worker.wait()
            except Exception:
                pass
            self.worker = None

    def closeEvent(self, event):
        if self.worker and self.worker.isRunning():
            self.worker.stop()
            self.worker.wait()
        event.accept()

def main():
    app = QApplication(sys.argv)
    w = TypingSimulator()
    w.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()