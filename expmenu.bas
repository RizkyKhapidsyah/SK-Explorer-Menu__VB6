Attribute VB_Name = "Module1"
Option Explicit
        
Public Type POINTAPI
        x As Long
        y As Long
End Type
        
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function DoExplorerMenu Lib "cfexpmnu.dll" (ByVal hWnd As Long, ByVal sFilePath As String, ByVal x As Long, ByVal y As Long) As Boolean

            

