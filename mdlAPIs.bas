Attribute VB_Name = "mdlAPIs"
Option Explicit

#If VBA7 Then
  'Object Window Apis \/
  Public Declare PtrSafe Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  Public Declare PtrSafe Function FindWindowExA Lib "user32.dll" (ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
  Public Declare PtrSafe Function PostMessageA Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal wMsg As LongPtr, ByVal wParam As LongPtr, ByVal lParama As LongPtr) As Long
  Public Declare PtrSafe Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
  Public Declare PtrSafe Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
  Public Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  Public Declare PtrSafe Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
  Public Declare PtrSafe Function GetWindowLongA Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal nIndex As Integer) As Long
  Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
  Public Declare PtrSafe Function SendMessageByString Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
  'Mouse APIS \/
  Public Declare PtrSafe Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
  Public Declare PtrSafe Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
  Public Declare PtrSafe Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
  'Download API\/
  Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long

#Else
  'Object Window Apis \/
  Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
  Public Declare Function FindWindowExA Lib "user32.dll" (ByVal hwndParent As LongPtr, ByVal hwndChildAfter As LongPtr, ByVal lpszClass As String, ByVal lpszWindow As String) As Long
  Public Declare Function PostMessageA Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal wMsg As LongPtr, ByVal wParam As LongPtr, ByVal lParama As LongPtr) As Long
  Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
  Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
  Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
  Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
  Public Declare Function GetWindowLongA Lib "user32.dll" (ByVal hwnd As LongPtr, ByVal nIndex As Integer) As Long
  Public Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
  Public Declare Function SendMessageByString Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
  'Mouse APIS \/
  Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
  Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
  Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
  'Download API\/
  Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
#End If


Public Type POINTAPI
  X_Pos As Long
  Y_Pos As Long
End Type
