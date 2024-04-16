C:\Users\22965\anaconda3\envs\autorun\python "C:\Users\22965\LocalPC\UI_autorun\myauto\bjl.py"
@echo off
set /a countdown=60
echo cmd will be closed, please note...
:countdown
echo Remaining Time: %countdown% seconds
set /a countdown-=1
timeout /t 1 >nul
if %countdown% leq 0 goto exit
goto countdown

:exit
echo Countdown finished, the window will close now.
timeout /t 1 >nul
exit