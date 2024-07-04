# -*- coding: utf-8 -*-
import sys
from setuptools import find_packages, setup
from uiautomation.version import VERSION

requires = ['comtypes>=1.2.1'] if sys.version_info >= (3, 7) else ['comtypes==1.2.1']

if sys.version_info < (3, 5):
    requires.append('typing')

setup(
    name='uiautomation',
    version=VERSION,
    description='Python UIAutomation for Windows',
    license='Apache 2.0',
    author='yinkaisheng',
    author_email='yinkaisheng@foxmail.com',
    keywords='windows ui automation uiautomation inspect',
    url='https://github.com/yinkaisheng/Python-UIAutomation-for-Windows',
    platforms='Windows Only',
    packages=find_packages(),
    include_package_data=True,
    scripts=['scripts/automation.py', 'scripts/automation.py'],
    long_description='Python UIAutomation for Windows. Supports Python3.4+, x86, x64',
    zip_safe=False,
    classifiers=[
        "Programming Language :: Python :: 3.4",
        "Programming Language :: Python :: 3.5",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
        "Programming Language :: Python :: 3.9",
        "Programming Language :: Python :: 3.10",
        "Programming Language :: Python :: 3.11",
        "Programming Language :: Python :: 3.12",
    ],
    install_requires=requires
)
