# -*- coding: utf-8 -*-
"""纯逻辑单元测试:名单文件存取、密码、抽取逻辑(不依赖 GUI 窗口)"""
import os
import sys
import json
import base64
import tempfile
import unittest
import atexit

PROJECT_DIR = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, PROJECT_DIR)

# import main 会执行多开检测并写 config.json 到项目根,
# 先删除旧文件保证 import 走"创建"分支,测试结束后清理
root_cfg = os.path.join(PROJECT_DIR, 'config.json')


def _del_cfg():
    if os.path.exists(root_cfg):
        os.remove(root_cfg)


atexit.register(_del_cfg)
_del_cfg()

import main  # noqa: E402


class NewListTest(unittest.TestCase):
    """名单文件存取与密码逻辑"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        main.name_pro = os.path.join(self.tmp, 'name.pro')
        main.name_file = os.path.join(self.tmp, 'name.wow')
        main.config_file = os.path.join(self.tmp, 'config.json')
        main.file_manager = main.NewList()

    def test_save_load_roundtrip(self):
        names = ['张三', '李四', '王五']
        main.file_manager.save(names)
        main.file_manager.load()
        self.assertEqual(main.name_list, names)
        self.assertEqual(main.name_password, '')

    def test_save_cleans_old_student_keys(self):
        """修复 L5:保存更少的人后,旧 student 键不应残留"""
        main.file_manager.save(['a', 'b', 'c', 'd'])
        main.file_manager.save(['x', 'y'])
        content = main.file_manager.file_load()
        student_keys = [k for k in content.keys() if k.startswith('student')]
        self.assertEqual(len(student_keys), 2)
        self.assertEqual(content['num'], 2)

    def test_corrupt_file_backed_up(self):
        """修复 L7:损坏的名单文件应被备份而不是静默清空"""
        with open(main.name_pro, 'wb') as f:
            f.write(b'not-base64!!!')
        result = main.file_manager.file_load()
        self.assertEqual(result['num'], 0)
        self.assertTrue(os.path.exists(main.name_pro + '.bak'))

    def test_passwd_set_and_clear(self):
        main.file_manager.save(['a'])
        # 无密码时设置密码
        self.assertTrue(main.file_manager.passwd('', 'secret'))
        main.file_manager.load()
        self.assertEqual(main.name_password, 'secret')
        # 密码错误
        self.assertFalse(main.file_manager.passwd('wrong', 'x'))
        # 清除密码
        self.assertTrue(main.file_manager.passwd('secret', ''))
        main.file_manager.load()
        self.assertEqual(main.name_password, '')

    def test_load_missing_student_keys_no_crash(self):
        """修复:文件 num 与键不一致时不应 KeyError 崩溃"""
        content = {'num': 5, 'student0': 'a', 'password_exist': False}
        with open(main.name_pro, 'wb') as f:
            f.write(base64.b64encode(json.dumps(content).encode('utf-8')))
        main.file_manager.load()
        self.assertEqual(main.name_list, ['a'])


class GetListNewTest(unittest.TestCase):
    """核心抽取逻辑"""

    def setUp(self):
        main.name_list = []
        main.counted_list = []

    def test_large_roster_consumes_one(self):
        main.name_list = ['学生%d' % i for i in range(30)]
        result = main.get_list_new()
        self.assertEqual(len(main.counted_list), 1)
        self.assertEqual(len(main.name_list), 29)
        self.assertEqual(result[-1], main.counted_list[-1])  # 落点即被抽中者
        self.assertNotIn(result[-1], main.name_list)

    def test_small_roster_consumes_one(self):
        """修复 L1:人数不足 13 时也要真实消耗 1 人"""
        main.name_list = ['学生%d' % i for i in range(8)]
        result = main.get_list_new()
        self.assertEqual(len(main.counted_list), 1)
        self.assertEqual(len(main.name_list), 7)
        self.assertEqual(result[-1], main.counted_list[-1])
        self.assertNotIn(result[-1], main.name_list)

    def test_empty_roster(self):
        result = main.get_list_new()
        self.assertEqual(result, [])
        self.assertEqual(main.counted_list, [])

    def test_duplicate_names_consume_one(self):
        """修复 L3:重名场景数量上仍正确消耗 1 人,不崩溃"""
        main.name_list = ['张三', '张三', '李四'] * 7  # 21 人
        total_before = len(main.name_list)
        main.get_list_new()
        self.assertEqual(len(main.counted_list), 1)
        self.assertEqual(len(main.name_list), total_before - 1)

    def test_repeated_draws_never_repeat(self):
        """不重复点名:抽满为止,被抽中的人不会再次出现"""
        main.name_list = ['学生%d' % i for i in range(10)]
        drawn = []
        for _ in range(10):
            main.get_list_new()
            drawn.append(main.counted_list[-1])
        self.assertEqual(len(set(drawn)), 10)
        self.assertEqual(main.name_list, [])
        self.assertEqual(len(main.counted_list), 10)


class MiscTest(unittest.TestCase):
    def test_checkfile(self):
        self.assertTrue(main.checkfile(__file__))
        self.assertFalse(main.checkfile('/nonexistent/path/xyz'))

    def test_advanced_shuffle(self):
        items = list(range(100))
        result = main.advanced_shuffle(items)
        self.assertEqual(sorted(result), items)
        self.assertIsInstance(result, list)
        self.assertEqual(len(result), 100)


if __name__ == '__main__':
    unittest.main()
