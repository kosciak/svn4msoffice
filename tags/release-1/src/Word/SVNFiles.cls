VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SVNFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private locked As String
Private filesystemobj As Object

Private Sub Class_Initialize()
    locked = ""
    Set filesystemobj = CreateObject("Scripting.FileSystemObject")
End Sub

Function isLocked() As Boolean
    Dim arr() As String
    Dim el
    
    isLocked = False
    
    arr = Split(locked, "??")
    For Each el In arr
        If el = ActiveDocument.FullName Then isLocked = True
    Next el
End Function

Sub unlockFile()
    Dim arr() As String
    Dim newArr As String
    Dim el
    
    arr = Split(locked, "??")
    For Each el In arr
        If el <> ActiveDocument.FullName Then newArr = newArr + "??" + el
    Next el
    
    locked = newArr
End Sub

Sub lockFile()
    If isLocked = False Then locked = locked + "??" + ActiveDocument.FullName
End Sub


Function isFolderVersioned() As Boolean
    isFolderVersioned = filesystemobj.FolderExists(ActiveDocument.Path & "\.svn")
End Function

Function isFileVersioned() As Boolean
    isFileVersioned = filesystemobj.fileexists(ActiveDocument.Path & "\.svn\text-base\" & ActiveDocument.name & ".svn-base")
End Function
