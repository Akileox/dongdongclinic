@echo off
echo ==========================================
echo   Student Report Generator Server Start
echo ==========================================
echo Starting server in WSL...
wsl -d Ubuntu -u akileo -e bash -c "cd /home/akileo/workspace/dongdongclinic && source myenv/bin/activate && python3 app.py"
pause
