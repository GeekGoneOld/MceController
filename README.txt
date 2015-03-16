WMCController
=============

Windows Media Center TCP/IP Controller (Fork: vmccontroller.codeplex.com).

Started this project because the slow support/development on vmccontroller.codeplex.com 

Based on vmccontroller 47386 with patch 6442.

Development on W7-32 using VS2012 Express with patch and WiX3.8
	Build .sln in root for ANY_CPU and X64.
	Edit build.bat in setup folder to change WIX_BUILD_LOCATION and SRC_PATH
	Run build.bat
	Repeat last two instructions for setup x64 folder

Known bugs:
1.	If computer is very busy when loaded (e.g. at Windows Startup), it may fail to load.
	If so, stop MC and restart when Windows has settled down
2.	Very first button press command doesn't work the very first time after each startup.

