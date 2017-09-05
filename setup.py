# -*- coding: utf-8 -*-
from setuptools import find_packages, setup

from uiautomation import VERSION, AUTHOR_MAIL

setup(
    name='uiautomation',
    version=VERSION,
    description='Python UIAutomation for Windows',
    license='Apache 2.0',
    author='yinkaisheng',
    author_email=AUTHOR_MAIL,
    keywords="windows ui automation uiautomation inspect",
    url='https://github.com/yinkaisheng/Python-UIAutomation-for-Windows',
    platforms='Windows Only',
    packages=find_packages(),
    include_package_data=True,
    scripts=["scripts/automation.py", 'scripts/automation.py'],
    long_description='Python UIAutomation for Windows. Supports py2, py3, x86, x64',
    zip_safe=False,
)
