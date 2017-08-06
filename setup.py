# -*- coding: utf-8 -*-
from setuptools import find_packages, setup
from glob import glob

from uiautomation.version import VERSION

setup(
    name='uiautomation',
    version=VERSION,
    description='Python UIAutomation for Windows',
    license='Apache 2.0',
    author='yinkaisheng',
    author_email='yinkaisheng@foxmail.com',
    keywords="windows ui automation uiautomation inspect",
    url='https://github.com/yinkaisheng/Python-UIAutomation-for-Windows',
    platforms='Windows Only',
    packages=find_packages(),
    data_files=[("bin", glob('bin/*.dll'))],
    scripts=["scripts/automation.py", 'scripts/automation.py'],
    long_description='Python UIAutomation for Windows. Supports py2, py3, x86, x64',
)
