Attribute VB_Name = "Module1"
Public Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Public Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Declare Function mciSendString Lib "mmsystem" (ByVal lpstrCommand$, ByVal lpstrReturnStr As Any, ByVal wReturnLen%, ByVal hCallBack%) As Long
