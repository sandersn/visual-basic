Attribute VB_Name = "GetWindowsDir"
Option Explicit

Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" _
    (ByVal lpBuffer As String, ByVal nSize As Long) As Long

Public Function GetWinDirectory() As String
Dim lResult As Long
Dim strBuffer As String * 255
    lResult = GetWindowsDirectory(strBuffer, Len(strBuffer))
    strBuffer = Left(strBuffer, lResult)
    GetWinDirectory = strBuffer
End Function
