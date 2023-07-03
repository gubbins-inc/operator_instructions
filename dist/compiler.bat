cd C:\Users\gcs\Documents\GitHub\operator_instructions\operator_instructions
pyinstaller --onefile  --icon=op-ins.ico --hiddenimport pkgutil OICC.py
cd dist
timeout /t 3
OICC.exe --firstset
timeout /t 3
copy "OICC.exe" "\\NT4\Client_Files\Public\PEOI\Codebase\OICC_FILES"
timeout /t 10
cd C:\Program Files (x86)\Inno Setup 6
timeout /t 5
iscc \\NT4\Client_Files\Public\PEOI\Codebase\OICC_FILES\update_oicc.iss
timeout /t 5
iscc \\NT4\Client_Files\Public\PEOI\Codebase\OICC_FILES\full_oicc.iss