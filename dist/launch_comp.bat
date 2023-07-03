cd C:\Users\gcs\Documents\GitHub\operator_instructions\operator_instructions
pyinstaller --onefile  --icon=launcher.ico --hiddenimport pkgutil Launch_OICC.py
timeout /t 1
cd dist
timeout /t 3
copy "Launch_OICC.exe" "\\NT4\Client_Files\Public\PEOI\Codebase\OICC_FILES"
timeout /t 10
cd C:\Program Files (x86)\Inno Setup 6
timeout /t 5
iscc \\NT4\Client_Files\Public\PEOI\Codebase\OICC_FILES\update_oicc.iss
timeout /t 5
iscc \\NT4\Client_Files\Public\PEOI\Codebase\OICC_FILES\full_oicc.iss