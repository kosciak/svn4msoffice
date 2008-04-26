Attribute VB_Name = "SVN"
Option Explicit

Dim oAppClass As New SVNEventCatcher
Private conf As SVNConfig
Private oShell As Object
Private filesystemobj As Object

Public files As SVNFiles
Public gui As SVNGUI

Public Const version = "0.3"

Sub AutoExec()
    'MsgBox "AutoExec"
    SVNInit
End Sub

Private Sub SVNInit()
    'MsgBox "SVNInit()"
    Set files = New SVNFiles
    Set conf = New SVNConfig
    Set gui = New SVNGUI

    Set oAppClass.oApp = Word.Application
    Set oShell = CreateObject("Wscript.Shell")
    Set filesystemobj = CreateObject("Scripting.FileSystemObject")
    
End Sub

Private Sub SVNInstall()
    gui.install
End Sub

Private Sub SVNUninstall()
    gui.uninstall
    ActiveDocument.AttachedTemplate.Save
End Sub

Sub svnAdd()
    If files.isFileVersioned = True Then
        MsgBox "File is already under version control!", vbExclamation
        Exit Sub
    End If
    
    RunTortoiseCommand "add"
    gui.checkButtons
End Sub

Sub svnCommit()

    If forceSave("Do you want to save document before Commit?") = True Then
        If files.isFileVersioned = False Then
            Dim response
            response = MsgBox("Do you want to Add File before Commit?", vbOKCancel + vbQuestion, "File not added")
            If response = vbCancel Then Exit Sub
            
            svnAdd
        End If
        RunTortoiseCommand "commit"
    End If
End Sub


Private Sub Show()
CommandBars("SVN").Enabled = True
    gui.showMenuAndToolbar True
    
End Sub

Sub svnRevert()
    If forceAdd() = False Then Exit Sub
    
    Dim response
    response = MsgBox("All your current and local changes will be lost!", vbOKCancel + vbExclamation + vbDefaultButton2, "Warning!")
    If response = vbCancel Then Exit Sub
    
    Dim name
    name = ActiveDocument.FullName
    
    ActiveDocument.Close False
    RunTortoiseCommandOnPath "revert", name
    Documents.Open name
End Sub

Sub svnUpdate()

    If forceAdd() = False Then Exit Sub

    If forceSave("All your current changes will be lost!" & vbCrLf & "Do you want to save document before Update?") = False Then Exit Sub
    
    Dim name
    name = ActiveDocument.FullName
    
    ActiveDocument.Close , False
    RunTortoiseCommandOnPath "update", name
    Documents.Open name
End Sub

Sub svnLock()
    
    If forceAdd() = False Then Exit Sub

    If files.isLocked() = False Then
        files.lockFile
    End If
    
    RunTortoiseCommand "lock"
End Sub

Sub svnUnLock()
    
    If forceAdd() = False Then Exit Sub

    If files.isLocked() = True Then files.unlockFile
    
    RunTortoiseCommand "unlock"
End Sub

Sub svnLog()

    If forceAdd() = False Then Exit Sub
    
    RunTortoiseCommand "log"
End Sub

'Sub svnCopy()
'    RunTortoiseCommand "copy", ActiveDocument.FullName
'End Sub

'Sub svnRename()
'    RunTortoiseCommand "rename", ActiveDocument.FullName
'End Sub

Sub svnStatus()
    
    If forceAdd() = False Then Exit Sub

    RunTortoiseCommand "repostatus"
End Sub

Sub svnDiff()

    If forceAdd() = False Then Exit Sub
    
    If forceSave("Do you want to save document before Diff?") = True Then RunTortoiseCommand "diff"
End Sub


Sub svnAbout()
    MsgBox "SVN4MSOffice version " & version & vbCrLf & "http://code.google.com/p/svn4msoffice/" & vbCrLf & vbCrLf & "Author: Wojciech 'KosciaK' Pietrzok", , "About"
End Sub



Private Sub RunTortoiseCommand(sCommand)
    RunTortoiseCommandOnPath sCommand, ActiveDocument.FullName
End Sub


''
' Runs tortoise commands
''
Private Sub RunTortoiseCommandOnPath(sCommand, sPath)
    Dim x
   
    On Error GoTo NoSVN
    x = oShell.Run("""" & conf.TortoisePath & """ /command:" & sCommand & " /path:""" & sPath & """ /notempfile /closeonend:4", , True)
    Exit Sub
    
NoSVN:
    MsgBox "TortoiseSVN not installed or not working correctly!", vbCritical, "TortoiseSVN Error"
End Sub

Private Function forceSave(sMsg As String) As Boolean
    Dim response
    forceSave = True
    If ActiveDocument.Saved = False Then
        response = MsgBox(sMsg, vbYesNoCancel + vbQuestion, "File not saved")
        If response = vbYes Then ActiveDocument.Save
        If response = vbCancel Then forceSave = False
    End If
End Function


Private Function forceAdd() As Boolean
    forceAdd = True
    
    If files.isFileVersioned = False Then
        MsgBox "File is not under version control!", vbExclamation
        forceAdd = False
    End If

End Function


