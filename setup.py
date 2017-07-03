#coding=utf-8
from distutils.core import setup
import py2exe
setup(console=["ExcelDiffer.py"], options={"py2exe":{"includes":["sip"]}}) 