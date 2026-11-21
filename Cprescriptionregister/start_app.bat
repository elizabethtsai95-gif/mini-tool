@echo off
title Prescription Register App - keep this window open
cd /d C:\Cprescriptionregister
start "" cmd /c "timeout /t 4 >nul & start http://localhost:8501"
python -m streamlit run Doc2us.py --server.headless true --server.address 0.0.0.0 --server.port 8501
pause
