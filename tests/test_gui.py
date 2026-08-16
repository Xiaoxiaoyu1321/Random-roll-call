# -*- coding: utf-8 -*-
"""GUI 冒烟测试:窗口构造、点击抽取流程、悬浮球点击/拖动区分(offscreen 平台)"""
import os
import sys
import time
import tempfile
import unittest
import atexit

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_DIR)
# 需要真实窗口平台:qfluentwidgets 的 qframelesswindow 在 macOS 的
# offscreen 平台无法获取原生 NSWindow,必须用 cocoa;Windows 上用默认平台
if sys.platform == 'darwin':
    os.environ.setdefault('QT_QPA_PLATFORM', 'cocoa')

# 与 test_logic.py 相同:保证 import main 走"创建"分支,结束后清理
root_cfg = os.path.join(PROJECT_DIR, 'config.json')


def _del_cfg():
    if os.path.exists(root_cfg):
        os.remove(root_cfg)


atexit.register(_del_cfg)
_del_cfg()

import main  # noqa: E402

from PyQt5 import QtWidgets  # noqa: E402
from PyQt5.QtCore import Qt  # noqa: E402
from PyQt5.QtTest import QTest  # noqa: E402


class FakeTray(object):
    """替代 SEEWO_Tool,避免 closeEvent 触发真实托盘"""

    def showMessage(self, *a, **k):
        pass

    def notify(self, *a, **k):
        pass

    def stop(self):
        pass


class GuiSmokeTest(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication([])
        cls.tmp = tempfile.mkdtemp()
        main.name_pro = os.path.join(cls.tmp, 'name.pro')
        main.config_file = os.path.join(cls.tmp, 'config.json')
        main.file_manager = main.NewList()
        main.SEEWO_Tool = FakeTray()

    def setUp(self):
        main.name_list = []
        main.counted_list = []
        main.name_password = ''
        main.quiet_boot = False

    def tearDown(self):
        # 结束残留的 worker 线程
        win = getattr(main, 'mWindow', None)
        if win is not None:
            worker = getattr(win.homeInterface, 'worker', None)
            if worker is not None and worker.isRunning():
                worker.wait(3000)
            try:
                win.close()
            except Exception:
                pass

    def test_welcome_window_construction(self):
        main.mWindow = main.WelcomeWindow()
        self.assertIsNotNone(main.mWindow.homeInterface)
        self.assertIsNotNone(main.mWindow.setInterface)
        self.assertIsNotNone(main.mWindow.aboutInterface)

    def test_quiet_boot_no_crash(self):
        """修复 S1:quiet_boot 下构造主窗口不应再 NameError 崩溃"""
        main.quiet_boot = True
        try:
            main.mWindow = main.WelcomeWindow()
        finally:
            main.quiet_boot = False
        self.assertIsNotNone(main.mWindow)

    def test_start_button_draw_flow(self):
        """修复 S2:点击开始后通过 worker 抽选,UI 状态正确恢复"""
        main.mWindow = main.WelcomeWindow()
        win = main.mWindow.homeInterface  # 主功能窗口(MainWindow)
        main.name_list = ['学生%d' % i for i in range(30)]
        main.counted_list = []
        total_before = 30

        win.StartButton_do()
        # 抽选期间按钮应被禁用
        self.assertFalse(win.Start_button.isEnabled())

        # 等待 worker 完成(按钮恢复为"开始")
        deadline = time.time() + 15
        while time.time() < deadline:
            self.app.processEvents()
            if win.Start_button.isEnabled() and win.Start_button.text() == '开始':
                break
            time.sleep(0.05)

        self.assertTrue(win.Start_button.isEnabled())
        self.assertEqual(win.Start_button.text(), '开始')
        self.assertEqual(len(main.counted_list), 1)
        self.assertEqual(len(main.name_list), total_before - 1)
        self.assertNotEqual(win.name_label.text(), '未选定')
        self.assertNotEqual(win.name_label.text(), '文件为空')

    def test_start_button_empty_roster(self):
        main.mWindow = main.WelcomeWindow()
        win = main.mWindow.homeInterface
        main.name_list = []
        win.StartButton_do()
        self.assertEqual(win.name_label.text(), '文件为空')

    def test_floating_ball_click_opens_window(self):
        """短按(<0.5s)视为点击,应打开主窗口"""
        main.mWindow = main.WelcomeWindow()
        ball = main.FloatingBall()
        ball.show()
        main.mWindow.hide()
        self.app.processEvents()
        self.assertFalse(main.mWindow.isVisible())

        QTest.mousePress(ball, Qt.LeftButton)
        QTest.mouseRelease(ball, Qt.LeftButton)
        self.app.processEvents()
        self.assertTrue(main.mWindow.isVisible())
        ball.close()

    def test_floating_ball_drag_does_not_open(self):
        """修复 S4:长按(拖动)后松开不应弹出主窗口"""
        main.mWindow = main.WelcomeWindow()
        ball = main.FloatingBall()
        ball.show()
        main.mWindow.hide()
        self.app.processEvents()

        QTest.mousePress(ball, Qt.LeftButton)
        time.sleep(0.6)  # 按住超过 0.5 秒
        QTest.mouseRelease(ball, Qt.LeftButton)
        self.app.processEvents()
        self.assertFalse(main.mWindow.isVisible())
        ball.close()


if __name__ == '__main__':
    unittest.main()
