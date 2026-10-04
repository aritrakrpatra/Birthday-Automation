@echo off
cd /d "%~dp0"
python birthday_automation.py >> scheduler_run.log 2>&1
