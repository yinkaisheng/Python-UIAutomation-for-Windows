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
    long_description='Python UIAutomation for Windows. Supports Python3.4+, x86, x64',
    zip_safe=False,
    classifiers=[
        "Programming Language :: Python :: 3.4",
        "Programming Language :: Python :: 3.5",
        "Programming Language :: Python :: 3.6",
        "Programming Language :: Python :: 3.7",
        "Programming Language :: Python :: 3.8",
    ],
    install_requires=[
        'typing',
        'comtypes>=1.1.7',
        ]
)
