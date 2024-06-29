C:\Users\22965\anaconda3\envs\autorun\python "C:\Users\22965\LocalPC\UI_autorun\myauto\bjl.py"
@echo off
setlocal enabledelayedexpansion
set /a countdown=10
echo cmd will be closed, please note...
:countdown
<nul set /p =Remaining Time: !countdown! seconds
set /a countdown-=1
timeout /t 1 >nul
if !countdown! gtr 0 (
    echo.
    goto countdown
)

echo.
echo Countdown finished, the window will close now.
timeout /t 2 >nul
exit
