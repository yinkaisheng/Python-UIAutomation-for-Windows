del /q dist\*
python setup.py bdist_wheel
pause
twine upload dist\*
pause