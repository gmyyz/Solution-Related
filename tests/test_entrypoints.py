import py_compile
import sys
import unittest
from pathlib import Path


ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / '工作区域'))


class EntrypointCompileTests(unittest.TestCase):
    def test_pyw_entrypoints_compile(self):
        entrypoints = [
            ROOT / '工作区域' / '运行-1.工资奖金自动化处理平台.pyw',
            ROOT / '工作区域' / '运行-2.工时周报生成.pyw',
            ROOT / '工作区域' / '原始数据' / '工时数据' / '运行-工时周报生成.pyw',
        ]
        for path in entrypoints:
            with self.subTest(path=path):
                py_compile.compile(str(path), doraise=True)


if __name__ == '__main__':
    unittest.main()
