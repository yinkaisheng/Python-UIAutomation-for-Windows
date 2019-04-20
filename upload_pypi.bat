del /q dist\*
python setup.py bdist_wheel
python rename_wheel.py
pause
twine upload dist\*
pause