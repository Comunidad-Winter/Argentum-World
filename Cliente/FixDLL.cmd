cd /d "%~dp0"

DEL /f "C:/Windows/SysWow64/Aurora.Network.dll"
regsvr32 Aurora.Network.dll