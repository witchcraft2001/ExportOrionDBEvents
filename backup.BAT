rem @echo off
rem set d=%date:~0,2%
rem set m=%date:~3,2%
rem set y=%date:~6,4%
rem set arch_date=%y%.%m%.%d%
setlocal enabledelayedexpansion
set rarexe="%ProgramFiles%"\WinRAR\Rar.exe
if exist = %rarexe% goto backup
set rarexe="%ProgramFiles(x86)%"\WinRAR\Rar.exe
if exist = %rarexe% goto backup
set rarexe="%ProgramW6432%"\WinRAR\Rar.exe
if exist = %rarexe% goto backup
@echo "Ошибка! Не найден WinRAR!!!"
pause 0
goto end

:backup

set name=%cd%
pushd ..
rem cd ..
rem @echo !name:%cd%\=!
set arcName=!name:%cd%\=!
popd

if not exist "..\backup" (
	mkdir "..\backup"
)

%rarexe% a -agYYYYMMDD-HHMM -dh -r -x*.*log -x*.rar -x*.pdb -x*.idb -x*.tmp -x*.obj ..\backup\"%arcName%_" *.*

:end
endlocal
rem pause 0