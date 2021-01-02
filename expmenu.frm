VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2985
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4965
   LinkTopic       =   "Form1"
   ScaleHeight     =   2985
   ScaleWidth      =   4965
   StartUpPosition =   3  'Windows Default
   Begin VB.DirListBox Dir1 
      Height          =   1890
      Left            =   240
      TabIndex        =   2
      Top             =   240
      Width           =   1815
   End
   Begin VB.FileListBox File1 
      Height          =   1845
      Left            =   2280
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Explorer Menu"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   4575
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim pt As POINTAPI
    Dim dl As Long
    Dim lCurX As Long
    Dim lCurY As Long
    Dim sFilePath As String
    On Error Resume Next
    
    dl = GetCursorPos(pt)
    
    sFilePath = Dir1.Path
    If Right$(sFilePath, 1) <> "\" Then sFilePath = sFilePath & "\"
    sFilePath = sFilePath & File1.FileName & Chr$(0)
    
    dl = DoExplorerMenu((Me.hWnd), sFilePath, pt.x, pt.y)

End Sub


Private Sub Drive1_Change()
End Sub


Private Sub Dir1_Change()
    File1.Path = Dir1.Path
    If File1.ListCount > 0 Then
        File1.ListIndex = 0
        Command1.Enabled = True
    Else
        Command1.Enabled = False
    End If
    
End Sub


Private Sub Form_Load()
    Dir1.Path = App.Path
End Sub


