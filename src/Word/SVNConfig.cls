VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SVNConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private sTOrtoiseproc    As String


Private Sub Class_Initialize()
    Dim oShell As Object
    Set oShell = CreateObject("WScript.Shell")
    
    'sTOrtoiseproc = """" & Environ("ProgramFiles") & "\TortoiseSVN\bin\TortoiseProc.exe"""
    
    On Error GoTo NoTortoiseSVN
    sTOrtoiseproc = oShell.RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\TortoiseSVN\ProcPath")
    Exit Sub
    
NoTortoiseSVN:
    MsgBox "Can't find TortoiseSVN", vbCritical, "TortoiseSVN Error"
End Sub

Public Property Get TortoisePath() As String
    TortoisePath = sTOrtoiseproc
End Property

