@echo off
title Flipkart Bulk Listing Automation Tool
echo Starting Flipkart Bulk Lister...
if exist venv\Scripts\python.exe (
    venv\Scripts\python.exe run.py
) else (
    python run.py
)
pause
