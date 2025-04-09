del /q dist\*
SET python=D:\Python38x64\python.exe
%python% setup.py bdist_wheel
%python% rename_wheel.py
pause
%python% -m twine upload dist\*
pause