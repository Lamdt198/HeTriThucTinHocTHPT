@echo off
chcp 65001 > nul
echo ========================================================
echo   HE THONG CANH BAO SOM NGUY CO HOC TAP MON TIN HOC
echo      De tai Cao hoc CNTT K20.2 - GDPT 2018
echo ========================================================
echo.
echo Dang khoi dong ung dung Web Streamlit...
python -m streamlit run app/app.py --server.port 8501 --server.headless false
pause
