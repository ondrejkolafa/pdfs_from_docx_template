@ECHO OFF
CLS
 
:: Check for Python Installation
 
python --version 2>NUL
if ERRORLEVEL 1 goto errorNoPython
 
 
:: Reaching here means Python is installed.
 
ECHO Python s installed
GOTO download_script
 
 
:errorNoPython
 
ECHO.
ECHO Error^: Python not installed
ECHO Installing python 3.12 from MS Store
winget install 9NCVDN91XZQP  --accept-source-agreements --accept-package-agreements
ECHO.
pip install docx2pdf
pip install docxtpl
pip install openpyxl
pip install pandas
pip install python-docx
GOTO download_script
 
 
:download_script
 
ECHO.
if exist create_documents.py (
    ECHO Python script found
) else (
    ECHO Downloading python script
    powershell -command wget https://raw.githubusercontent.com/ondrejkolafa/pdfs_from_docx_template/main/create_documents.py -OutFile create_documents.py
)
GOTO eof
 
 
:eof
 
CLS
python create_documents.py
 
PAUSE