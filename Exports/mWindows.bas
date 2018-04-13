Attribute VB_Name = "mWindows"
Option Explicit

' https://www.techrepublic.com/blog/10-things/10-plus-of-my-favorite-windows-api-functions-to-use-in-office-applications/

Declare Function GetSystemMetrics32 Lib "User32" Alias "GetSystemMetrics" _
                    (ByVal nIndex As Long) As Long
Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" _
                    (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" _
                    (ByVal lpBuffer As String, nSize As Long) As Long
 


Function apicGetUserName() As String
  'Call to apiGetUserName returns current user.
  Dim lngResponse As Long
  Dim strUserName As String * 32
  lngResponse = GetUserName(strUserName, 32)
  apicGetUserName = Left(strUserName, InStr(strUserName, Chr$(0)) - 1)
End Function


Function apicGetComputerName() As String
  'Call to apiGetUserName returns current user.
  Dim lngResponse As Long
  Dim strUserName As String * 32
  lngResponse = GetComputerName(strUserName, 32)
  apicGetComputerName = Left(strUserName, InStr(strUserName, Chr$(0)) - 1)
End Function

