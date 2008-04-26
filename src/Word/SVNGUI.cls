VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SVNGUI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Sub install()
    MsgBox "Installing"
    createMenu
    populateMenu
    createToolbar
End Sub

Sub uninstall()
    On Error Resume Next
    CommandBars("Menu Bar").Controls("&SVN").Delete
    CommandBars("SVN").Delete
End Sub


Sub showMenuAndToolbar(Show As Boolean)
    'MsgBox "showMenuAndToolbar " & show
ShowThem:
    On Error GoTo InstallThem
    CommandBars("Menu Bar").Controls("&SVN").Visible = Show
    CommandBars("SVN").Visible = Show
    checkButtons
    Exit Sub

InstallThem:
    ' There's no need to install the menu and toolbar if it won't be shown anyway
    If Show = True Then
        install
        GoTo ShowThem
    End If
End Sub


Sub checkButtons()
    Dim i, count As Integer
    Dim enable As Boolean
    
    enable = files.isFolderVersioned And files.isFileVersioned
    
    count = CommandBars("SVN").Controls.count
    
    CommandBars("SVN").Controls(1).Enabled = files.isFolderVersioned And Not enable
    CommandBars("Menu Bar").Controls("&SVN").CommandBar.Controls(1).Enabled = files.isFolderVersioned And Not enable
    
    For i = 2 To count
        CommandBars("SVN").Controls(i).Enabled = enable
        CommandBars("Menu Bar").Controls("&SVN").CommandBar.Controls(i).Enabled = enable
    Next i
 
End Sub


Sub createMenu()

    Dim svnMenu As CommandBarPopup
    
    Set svnMenu = CommandBars("Menu Bar").Controls.Add(Type:=msoControlPopup)
    With svnMenu
            .Caption = "&SVN"
            .BeginGroup = True
            .Visible = False
    
    End With

End Sub


Sub populateMenu()
    Dim svnMenu As CommandBarPopup
    
    ' You have to create the toolbar first and save it in the .dot file
    '
    'Set svnMenu = CommandBars("Menu Bar").Controls.Add(Type:=msoControlPopup)
    'With svnMenu
    '        .Caption = "&SVN"
    '        .BeginGroup = True
    '        .Visible = False
    '
    'End With
    
    Set svnMenu = CommandBars("Menu Bar").Controls("&SVN")
    
    Dim btnAdd, btnUpdate, btnCommit, btnRevert, btnLog, btnStatus, btnRename, btnCopy, btnLock, btnUnlock, btnDiff, btnAbout
    
    Set btnAdd = svnMenu.Controls.Add(Type:=msoControlButton)
    With btnAdd
            .FaceId = 137
            .OnAction = "svnAdd"
            .Caption = "&Add to Repository"
            .Enabled = False
    End With
    
    Set btnCommit = svnMenu.Controls.Add(Type:=msoControlButton)
    With btnCommit
            .FaceId = 270
            .OnAction = "svnCommit"
            .Caption = "&Commit File"
            .BeginGroup = True
            .Enabled = False
    End With
    
    Set btnUpdate = svnMenu.Controls.Add(Type:=msoControlButton)
    With btnUpdate
            .FaceId = 271
            .OnAction = "svnUpdate"
            .Caption = "&Update File"
            .Enabled = False
    End With

    Set btnRevert = svnMenu.Controls.Add(Type:=msoControlButton)
    With btnRevert
            .FaceId = 1921
            .OnAction = "svnRevert"
            .Caption = "&Revert"
            .Enabled = False
    End With
    

    Set btnLog = svnMenu.Controls.Add(Type:=msoControlButton)
    With btnLog
            .FaceId = 587
            .OnAction = "svnLog"
            .Caption = "Show &Log"
            .BeginGroup = True
            .Enabled = False
    End With

    Set btnStatus = svnMenu.Controls.Add(Type:=msoControlButton)
    With btnStatus
            .FaceId = 487
            .OnAction = "svnStatus"
            .Caption = "&Status"
            .Enabled = False
    End With
    
    Set btnDiff = svnMenu.Controls.Add(Type:=msoControlButton)
    With btnDiff
            .FaceId = 278
            .OnAction = "svnDiff"
            .Caption = "&Diff"
            .Enabled = False
    End With

'    Set btnRename = svnToolbar.Controls.Add(Type:=msoControlButton)
'    With btnRename
'            '.FaceId =
'            .OnAction = "svnRename"
'            .Caption = "Rename"
'            .Style = msoButtonIcon
'            .TooltipText = "Rename file"
'            .BeginGroup = True
'    End With
'

'    Set btnCopy = svnToolbar.Controls.Add(Type:=msoControlButton)
'    With btnCopy
'            '.FaceId =
'            .OnAction = "svnCopy"
'            .Caption = "Copy"
'            .Style = msoButtonIcon
'            .TooltipText = "Copy file"
'    End With
    

    Set btnLock = svnMenu.Controls.Add(Type:=msoControlButton)
    With btnLock
            .FaceId = 225
            .OnAction = "svnLock"
            .Caption = "&Lock file"
            .BeginGroup = True
            .Enabled = False
    End With
    

    Set btnUnlock = svnMenu.Controls.Add(Type:=msoControlButton)
    With btnUnlock
            .FaceId = 277
            .OnAction = "svnUnlock"
            .Caption = "U&nlock file"
            .Enabled = False
    End With
    
    Set btnAbout = svnMenu.Controls.Add(Type:=msoControlButton)
    With btnAbout
            '.FaceId =
            .OnAction = "svnAbout"
            .Caption = "A&bout"
            .BeginGroup = True
    End With

  
End Sub


Sub createToolbar()
    Dim svnToolbar As CommandBar
    Set svnToolbar = CommandBars.Add
    With svnToolbar
            .name = "SVN"
            '.Top = 0
            '.Left = 0
            .Visible = False
    End With
                    
    Dim btnAdd, btnUpdate, btnCommit, btnRevert, btnLog, btnStatus, btnRename, btnCopy, btnLock, btnUnlock, btnDiff
    
    Set btnAdd = svnToolbar.Controls.Add(Type:=msoControlButton)
    With btnAdd
            .FaceId = 137
            .OnAction = "svnAdd"
            .Caption = "Add"
            .Style = msoButtonIcon
            .TooltipText = "Add to Repository"
            .Enabled = False
    End With
    
    Set btnCommit = svnToolbar.Controls.Add(Type:=msoControlButton)
    With btnCommit
            .FaceId = 270
            .OnAction = "svnCommit"
            .Caption = "Commit"
            .Style = msoButtonIcon
            .TooltipText = "Commit File"
            .BeginGroup = True
            .Enabled = False
    End With
    
    Set btnUpdate = svnToolbar.Controls.Add(Type:=msoControlButton)
    With btnUpdate
            .FaceId = 271
            .OnAction = "svnUpdate"
            .Caption = "Update"
            .Style = msoButtonIcon
            .TooltipText = "Update File"
            .Enabled = False
    End With

    Set btnRevert = svnToolbar.Controls.Add(Type:=msoControlButton)
    With btnRevert
            .FaceId = 1921
            .OnAction = "svnRevert"
            .Caption = "Revert"
            .Style = msoButtonIcon
            .TooltipText = "Revert"
            .Enabled = False
    End With
    

    Set btnLog = svnToolbar.Controls.Add(Type:=msoControlButton)
    With btnLog
            .FaceId = 587
            .OnAction = "svnLog"
            .Caption = "Log"
            .Style = msoButtonIcon
            .TooltipText = "Show log"
            .BeginGroup = True
            .Enabled = False
    End With

    Set btnStatus = svnToolbar.Controls.Add(Type:=msoControlButton)
    With btnStatus
            .FaceId = 487
            .OnAction = "svnStatus"
            .Caption = "Status"
            .Style = msoButtonIcon
            .TooltipText = "Check for Modifications"
            .Enabled = False
    End With
    
    Set btnDiff = svnToolbar.Controls.Add(Type:=msoControlButton)
    With btnDiff
            .FaceId = 278
            .OnAction = "svnDiff"
            .Caption = "Diff"
            .Style = msoButtonIcon
            .TooltipText = "Diff"
            .Enabled = False
    End With

'    Set btnRename = svnToolbar.Controls.Add(Type:=msoControlButton)
'    With btnRename
'            '.FaceId =
'            .OnAction = "svnRename"
'            .Caption = "Rename"
'            .Style = msoButtonIcon
'            .TooltipText = "Rename file"
'            .BeginGroup = True
'    End With
'

'    Set btnCopy = svnToolbar.Controls.Add(Type:=msoControlButton)
'    With btnCopy
'            '.FaceId =
'            .OnAction = "svnCopy"
'            .Caption = "Copy"
'            .Style = msoButtonIcon
'            .TooltipText = "Copy file"
'    End With
    

    Set btnLock = svnToolbar.Controls.Add(Type:=msoControlButton)
    With btnLock
            .FaceId = 225
            .OnAction = "svnLock"
            .Caption = "Lock"
            .Style = msoButtonIcon
            .TooltipText = "Lock file"
            .BeginGroup = True
            .Enabled = False
    End With
    

    Set btnUnlock = svnToolbar.Controls.Add(Type:=msoControlButton)
    With btnUnlock
            .FaceId = 277
            .OnAction = "svnUnlock"
            .Caption = "Unlock"
            .Style = msoButtonIcon
            .TooltipText = "Unlock file"
            .Enabled = False
    End With
        
End Sub

