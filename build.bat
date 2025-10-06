@echo off
REM Build script for Multi Serial Monitor on Windows

echo Cleaning up old files...
if exist SerialMonitor.exe (
    del SerialMonitor.exe
    if exist SerialMonitor.exe (
        echo Error: Could not delete existing SerialMonitor.exe. Please close any running instances.
        pause
        exit /b 1
    )
)
if exist SerialMonitor_res.o del SerialMonitor_res.o

echo Compiling resource file...
windres SerialMonitor.rc -o SerialMonitor_res.o

if %errorlevel% neq 0 (
    echo Error compiling resource file
    pause
    exit /b 1
)

echo Compiling and linking executable...
g++ -o SerialMonitor.exe main.cpp SerialMonitor_res.o -mwindows -lcomctl32 -lriched20 -lgdi32 -luser32 -lkernel32 -ladvapi32

if %errorlevel% neq 0 (
    echo Error compiling executable
    pause
    exit /b 1
)

echo Build successful! Run SerialMonitor.exe to start the application.
pause