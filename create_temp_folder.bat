IF EXIST C:\Temp GOTO EXIT
mkdir c:\Temp
icacls C:\Temp /inheritance:e /grant *S-1-1-0:(F)
:EXIT