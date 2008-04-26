VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SVNEventCatcher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents oApp As Word.Application
Attribute oApp.VB_VarHelpID = -1

Private Sub oApp_DocumentChange()
    On Error GoTo ADError
    If ActiveDocument.name = "" Then Exit Sub
    'MsgBox "EventCatcher.oApp_DocumentChange"
    gui.showMenuAndToolbar files.isFolderVersioned()

ADError:
    ' There's no ActiveDocument - just exit the sub
End Sub

'Private Sub oApp_DocumentOpen(ByVal Doc As Document)
    'MsgBox "EventCatcher.oApp_DocumentOpen"
    'gui.showMenuAndToolbar files.isFolderVersioned()
'End Sub

Private Sub oApp_DocumentBeforeClose(ByVal Doc As Document, _
 Cancel As Boolean)
    'MsgBox "EventCatcher.oApp_DocumentClose"
    If files.isFolderVersioned = True Then
 
        Dim response
        response = MsgBox("Do you want to Commit document before closing?", vbYesNoCancel + vbQuestion, "File not saved")
        If response = vbYes Then svnCommit
        If response = vbCancel Then GoTo CancelClose:
        
        If files.isLocked() = True Then
            response = MsgBox("Unlock document before closing?", vbYesNo + vbQuestion, "File locked")
            If response = vbYes Then files.unlockFile
        End If
    
    End If
    
    If Application.Documents.count = 1 Then gui.showMenuAndToolbar False
    Exit Sub
    
CancelClose:
    Cancel = True
End Sub


Private Sub oApp_Quit()
    'MsgBox "EventCatcher.oApp_Quit"
    gui.uninstall
End Sub


