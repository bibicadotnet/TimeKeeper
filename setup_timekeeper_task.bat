@echo off
setlocal enabledelayedexpansion

:: Kiểm tra quyền Administrator
net session >nul 2>&1
if %errorlevel% NEQ 0 (
    echo This script requires Administrator privileges.
    echo Please right-click and select "Run as Administrator"
    pause
    exit /b 1
)

:: Thiết lập các biến
set "TASK_NAME=TimeKeeperTask"
set "DEST_DIR=C:\TimeKeeper"
set "DEST_VBS_PATH=%DEST_DIR%\TimeKeeper.vbs"
set "SCRIPT_DIR=%~dp0"
set "SOURCE_VBS_PATH=%SCRIPT_DIR%TimeKeeper.vbs"
set "XML_PATH=%TEMP%\TimeKeeperTask.xml"

:: Lấy thời gian hiện tại cho StartBoundary
for /f "tokens=2 delims==" %%I in ('wmic os get localdatetime /value') do set datetime=%%I
set "YEAR=%datetime:~0,4%"
set "MONTH=%datetime:~4,2%"
set "DAY=%datetime:~6,2%"
set "HOUR=%datetime:~8,2%"
set "MINUTE=%datetime:~10,2%"
set "SECOND=%datetime:~12,2%"

:: Tạo thư mục đích
if not exist "%DEST_DIR%" mkdir "%DEST_DIR%"

:: Kiểm tra file nguồn
if not exist "%SOURCE_VBS_PATH%" (
    echo Error: TimeKeeper.vbs not found in: %SCRIPT_DIR%
    pause
    exit /b 1
)

:: Sao chép file
copy /Y "%SOURCE_VBS_PATH%" "%DEST_VBS_PATH%" >nul

:: Xóa task cũ nếu tồn tại
schtasks /Query /TN "%TASK_NAME%" >nul 2>&1
if %errorlevel% EQU 0 (
    echo Removing existing task...
    schtasks /Delete /TN "%TASK_NAME%" /F >nul
)

:: Tạo XML cho task với nhiều trigger
echo ^<?xml version="1.0" encoding="UTF-16"?^>> "%XML_PATH%"
echo ^<Task version="1.4" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task"^>>> "%XML_PATH%"
echo   ^<RegistrationInfo^>>> "%XML_PATH%"
echo     ^<Description^>Synchronizes system time every 5 minutes and at startup^</Description^>>> "%XML_PATH%"
echo   ^</RegistrationInfo^>>> "%XML_PATH%"
echo   ^<Triggers^>>> "%XML_PATH%"
echo     ^<BootTrigger^>>> "%XML_PATH%"
echo       ^<Enabled^>true^</Enabled^>>> "%XML_PATH%"
echo     ^</BootTrigger^>>> "%XML_PATH%"
echo     ^<TimeTrigger^>>> "%XML_PATH%"
echo       ^<StartBoundary^>%YEAR%-%MONTH%-%DAY%T%HOUR%:%MINUTE%:%SECOND%^</StartBoundary^>>> "%XML_PATH%"
echo       ^<Repetition^>>> "%XML_PATH%"
echo         ^<Interval^>PT5M^</Interval^>>> "%XML_PATH%"
echo         ^<StopAtDurationEnd^>false^</StopAtDurationEnd^>>> "%XML_PATH%"
echo       ^</Repetition^>>> "%XML_PATH%"
echo       ^<Enabled^>true^</Enabled^>>> "%XML_PATH%"
echo     ^</TimeTrigger^>>> "%XML_PATH%"
echo   ^</Triggers^>>> "%XML_PATH%"
echo   ^<Principals^>>> "%XML_PATH%"
echo     ^<Principal id="Author"^>>> "%XML_PATH%"
echo       ^<RunLevel^>HighestAvailable^</RunLevel^>>> "%XML_PATH%"
echo       ^<UserId^>S-1-5-18^</UserId^>>> "%XML_PATH%"
echo     ^</Principal^>>> "%XML_PATH%"
echo   ^</Principals^>>> "%XML_PATH%"
echo   ^<Settings^>>> "%XML_PATH%"
echo     ^<MultipleInstancesPolicy^>IgnoreNew^</MultipleInstancesPolicy^>>> "%XML_PATH%"
echo     ^<DisallowStartIfOnBatteries^>false^</DisallowStartIfOnBatteries^>>> "%XML_PATH%"
echo     ^<StopIfGoingOnBatteries^>false^</StopIfGoingOnBatteries^>>> "%XML_PATH%"
echo     ^<AllowHardTerminate^>true^</AllowHardTerminate^>>> "%XML_PATH%"
echo     ^<StartWhenAvailable^>true^</StartWhenAvailable^>>> "%XML_PATH%"
echo     ^<RunOnlyIfNetworkAvailable^>false^</RunOnlyIfNetworkAvailable^>>> "%XML_PATH%"
echo     ^<IdleSettings^>>> "%XML_PATH%"
echo       ^<StopOnIdleEnd^>false^</StopOnIdleEnd^>>> "%XML_PATH%"
echo       ^<RestartOnIdle^>false^</RestartOnIdle^>>> "%XML_PATH%"
echo     ^</IdleSettings^>>> "%XML_PATH%"
echo     ^<AllowStartOnDemand^>true^</AllowStartOnDemand^>>> "%XML_PATH%"
echo     ^<Enabled^>true^</Enabled^>>> "%XML_PATH%"
echo     ^<Hidden^>false^</Hidden^>>> "%XML_PATH%"
echo     ^<RunOnlyIfIdle^>false^</RunOnlyIfIdle^>>> "%XML_PATH%"
echo     ^<DisallowStartOnRemoteAppSession^>false^</DisallowStartOnRemoteAppSession^>>> "%XML_PATH%"
echo     ^<UseUnifiedSchedulingEngine^>true^</UseUnifiedSchedulingEngine^>>> "%XML_PATH%"
echo     ^<WakeToRun^>false^</WakeToRun^>>> "%XML_PATH%"
echo     ^<ExecutionTimeLimit^>PT10M^</ExecutionTimeLimit^>>> "%XML_PATH%"
echo     ^<Priority^>7^</Priority^>>> "%XML_PATH%"
echo   ^</Settings^>>> "%XML_PATH%"
echo   ^<Actions Context="Author"^>>> "%XML_PATH%"
echo     ^<Exec^>>> "%XML_PATH%"
echo       ^<Command^>wscript.exe^</Command^>>> "%XML_PATH%"
echo       ^<Arguments^>//B "%DEST_VBS_PATH%"^</Arguments^>>> "%XML_PATH%"
echo     ^</Exec^>>> "%XML_PATH%"
echo   ^</Actions^>>> "%XML_PATH%"
echo ^</Task^>>> "%XML_PATH%"

:: Tạo task từ XML
schtasks /Create /TN "%TASK_NAME%" /XML "%XML_PATH%" /F
if %errorlevel% NEQ 0 (
    echo Failed to create scheduled task.
    del "%XML_PATH%"
    pause
    exit /b 1
)

:: Chạy task ngay lập tức
schtasks /Run /TN "%TASK_NAME%"

:: Xóa file XML tạm
del "%XML_PATH%"

echo Task setup completed successfully!
echo - Task will run at system startup
echo - Task will repeat every 5 minutes
echo - Task will run with SYSTEM privileges
echo.
echo You can verify the task in Task Scheduler
pause