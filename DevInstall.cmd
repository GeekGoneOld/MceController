@ECHO OFF
ECHO.
ECHO.Usage: Command.InstallAndRegister.cmd [/u][/debug]
ECHO.

set CompanyName=Media Center Network Controller
set AssemblyName=VmcController.Add-In
set RegistrationFilename=VmcController.Add-In.xml
set GACUtilPath=C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\Bin

set ReleaseType=Release
rem set ReleaseType=Debug

ECHO %GACUtilPath%

REM Determine whether we are on an 32 or 64 bit machine
if "%PROCESSOR_ARCHITECTURE%"=="x86" if "%PROCESSOR_ARCHITEW6432%"=="" goto x86

ECHO.On an x64 machine
set ProgramFilesPath=%ProgramFiles%
ECHO.

goto unregister

:x86

	ECHO.On an x86 machine
	set ProgramFilesPath=%ProgramFiles%
	ECHO.

:unregister

	ECHO.Unregister and delete previously installed files (which may fail if nothing is registered)
	ECHO.

	ECHO.Unregister the application entry points
	%windir%\ehome\RegisterMCEApp.exe /allusers "%ProgramFilesPath%\%CompanyName%\%RegistrationFilename%" /u
	ECHO.

	ECHO.Remove the DLLs from the Global Assembly cache
	"%GACUtilPath%\gacutil.exe" /u "%AssemblyName%"
	"%GACUtilPath%\gacutil.exe" /u Newtonsoft.Json
	"%GACUtilPath%\gacutil.exe" /u Interop.WMPLib
	ECHO.

	ECHO.Delete the folders containing the DLLs and supporting files (silent if successful)
	rd /s /q "%ProgramFilesPath%\%CompanyName%"
	ECHO.

	REM Exit out if the /u uninstall argument is provided, leaving no trace of the program files.
	if "%1"=="/u" goto exit
	
:releasetype

	REM evaluate the second argument
	if "%1"=="/debug" goto debug
	
	ECHO.Using the release version of the binaries
	set ReleaseType=Release
	ECHO.
	
	goto checkbin
	
:debug

	ECHO.Using the Debug version of the binaries
	set ReleaseType=Debug
	ECHO.
	
:checkbin

    if exist ".\Add-In\bin\%ReleaseType%\%AssemblyName%.dll" goto register
    ECHO.Cannot find %ReleaseType% binaries.
    ECHO.Build solution as %ReleaseType% and run script again. 
    goto exit

:register

	REM Copying and registering assemblies

	ECHO.Create the path for the binaries and supporting files (silent if successful)
	md "%ProgramFilesPath%\%CompanyName%"
	ECHO.
	
	ECHO.Copy the add-in assembly to Program Files
	copy /y ".\Add-In\bin\%ReleaseType%\*.dll" "%ProgramFilesPath%\%CompanyName%\"
	
	ECHO.Copy the status sink assembly to Program Files
	copy /y ".\Status\bin\%ReleaseType%\*.dll" "%ProgramFilesPath%\%CompanyName%\"		

	ECHO.Copy the registration XML to program files
	copy /y ".\Add-In\%RegistrationFilename%" "%ProgramFilesPath%\%CompanyName%\"
	ECHO.

	ECHO.Register the DLL with the global assembly cache
	"%GACUtilPath%\gacutil.exe" /if "%ProgramFilesPath%\%CompanyName%\%AssemblyName%.dll"
	ECHO.
	
	ECHO.Register the Interop DLL with the global assembly cache
	"%GACUtilPath%\gacutil.exe" /if "%ProgramFilesPath%\%CompanyName%\Interop.WMPLib.dll"
	ECHO.

	ECHO.Register the Newtonsoft.Json DLL with the global assembly cache
	"%GACUtilPath%\gacutil.exe" /if "%ProgramFilesPath%\%CompanyName%\Newtonsoft.Json.dll"
	ECHO.

	ECHO.Register the application with Windows Media Center
	%windir%\ehome\RegisterMCEApp.exe /allusers "%ProgramFilesPath%\%CompanyName%\%RegistrationFilename%"
	ECHO.



:exit
pause