@echo off

rem # Friendly dos menu system for dashboard-flow project

rem # set environment variables

rem # adjust the values below for your environment
set PROJECT_DIR=D:/Devel/automation/source/dashboard-flow
set PYTHON_BIN=D:/Devel/tools/WPy64-3820/python-3.8.2.amd64
set PYTHONPATH=%PROJECT_DIR%/src
set PATH=%PATH%;%PYTHON_BIN%
set DATA_DIR=D:/Devel/automation/apps/dashboard-flow

rem # the variables below do not need to be changed
set MODE_CREATE_MANIFEST_ONLY="CREATE_MANIFEST_ONLY"
set MODE_BUILD_DATASET_ONLY="BUILD_DATASET_ONLY"
set MODE_PUBLISH_DATASET_ONLY="PUBLISH_DATASET_ONLY"
set MODE_CREATE_BUILD_PUBLISH="CREATE_BUILD_PUBLISH"
set GOOGLE_SHEET_ID="1t5oKburSle0kg74mMp4Y7I0f78PJGho49NrSeZ4x_Sc"
set BLACK_GREEN=02
set BLACK_BLUE=09
set BLACK_AQUA=03
set BLACK_RED=04
set BLACK_PURPLE=05
set BLACK_YELLOW=06
set BLACK_WHITE=07

rem # create menu system
cls
:MENU
color %BLACK_GREEN%
echo.
echo ................................................................................
echo.
echo "  ______          _     _                         _  ______ _                 "
echo "  |  _  \        | |   | |                       | | |  ___| |                "
echo "  | | | |__ _ ___| |__ | |__   ___   __ _ _ __ __| | | |_  | | _____      __  "
echo "  | | | / _` / __| '_ \| '_ \ / _ \ / _` | '__/ _` | |  _| | |/ _ \ \ /\ / /  "
echo "  | |/ / (_| \__ \ | | | |_) | (_) | (_| | | | (_| | | |   | | (_) \ V  V /   "
echo "  |___/ \__,_|___/_| |_|_.__/ \___/ \__,_|_|  \__,_| \_|   |_|\___/ \_/\_/    "
echo.                                                                         
echo ................................................................................
echo.
echo.
echo.
echo press 1, 2, 3, 4 to select your task, or 5 to exit.
echo ................................................................................
echo.
echo 1 - Create manifest.json
echo 2 - Build dataset
echo 3 - Publish dataset to Google Data Studio
echo 4 - Create manifest.json, build and publish dataset to Google Data Studio
echo 5 - Exit
echo.
set /p m=type 1, 2, 3, 4 or 5 then press enter: 
if %m%==1 goto CREATE_MANIFEST
if %m%==2 goto BUILD_DATASET
if %m%==3 goto PUBLISH_DATASET
if %m%==4 goto CREATE_BUILD_PUBLISH
if %m%==5 goto eof
:CREATE_MANIFEST
cd %PROJECT_DIR%
echo "Executing main.py with datadir: %DATA_DIR% and mode %MODE_CREATE_MANIFEST_ONLY%"
python %PROJECT_DIR%/src/main.py --data-dir %DATA_DIR% --mode %MODE_CREATE_MANIFEST_ONLY% --sheet-id %GOOGLE_SHEET_ID%
pause
goto MENU
:BUILD_DATASET
cd %PROJECT_DIR%
echo "Executing main.py with datadir: %DATA_DIR% and mode %MODE_BUILD_DATASET_ONLY%"
python %PROJECT_DIR%/src/main.py --data-dir %DATA_DIR% --mode %MODE_BUILD_DATASET_ONLY% --sheet-id %GOOGLE_SHEET_ID%
pause
goto MENU
:PUBLISH_DATASET
cd %PROJECT_DIR%
echo "Executing main.py with datadir: %DATA_DIR% and mode %MODE_PUBLISH_DATASET_ONLY%"
python %PROJECT_DIR%/src/main.py --data-dir %DATA_DIR% --mode %MODE_PUBLISH_DATASET_ONLY% --sheet-id %GOOGLE_SHEET_ID%
pause
goto MENU
:CREATE_BUILD_PUBLISH
cd %PROJECT_DIR%
echo "Executing main.py with datadir: %DATA_DIR% and mode %MODE_CREATE_BUILD_PUBLISH%"
python %PROJECT_DIR%/src/main.py --data-dir %DATA_DIR% --mode %MODE_CREATE_BUILD_PUBLISH% --sheet-id %GOOGLE_SHEET_ID%
pause
goto MENU
