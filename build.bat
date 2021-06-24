@echo off

rem Copyright 2016 Google Inc. All Rights Reserved.

rem Licensed under the Apache License, Version 2.0 (the "License");
rem you may not use this file except in compliance with the License.
rem You may obtain a copy of the License at

rem     http://www.apache.org/licenses/LICENSE-2.0

rem Unless required by applicable law or agreed to in writing, software
rem distributed under the License is distributed on an "AS IS" BASIS,
rem WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
rem See the License for the specific language governing permissions and
rem limitations under the License.

echo Build PasswordSyncSupportTool's EXE files using Aut2Exe (https://www.autoitscript.com/site/autoit/downloads/)

echo Liron Newman lironn@google.com

rem Switch the the script's directory.
pushd "%~dp0"
rem Create a build directory, ignore errors.
mkdir build 2>nul
rem Find the version number by searching for the version set in the VBS file.
for /f "tokens=4" %%a IN ('findstr /R /B /C:"Const Ver = \"[0-9]*\.[0-9]*\.[0-9]*\.[0-9]*\"" gspstool.vbs') DO set ver=%%a
rem Strip the double quotes (") from the version number.
set ver=%ver:"=%
echo Going to build PasswordSyncSupportTool version %ver%...

rem Generate the build commandline.
set buildcmd=..\..\tools\autoit\Aut2Exe\Aut2Exe.exe /in gspstool.au3 /icon pictographs-fire_extinguisher_inv.ico /comp 4 /pack /gui /execlevel requireadministrator /companyname Google /productname "Password Sync support tool" /fileversion %ver% /productversion %ver%

echo Build command base: %buildcmd%

echo Building the x86 EXE.
%buildcmd% /out build\PasswordSyncSupportTool.exe

echo Building the x64 EXE.
%buildcmd% /out build\PasswordSyncSupportTool_x64.exe /x64

echo Done.
popd