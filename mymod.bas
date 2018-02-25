Attribute VB_Name = "mymod"
Option Explicit

Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    Flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Const OFN_EXPLORER = &H80000

Public Declare Function GetOFN Lib "comdlg32" Alias "GetOpenFileNameA" (ByRef ofn As OPENFILENAME) As Long
Public Declare Function GetActWindow Lib "user32" Alias "GetActiveWindow" () As Long

Public ofn As OPENFILENAME
Public fso As Object
Public SrcDBPath As String
Public bSuperSave As Boolean


Sub InitOFN()
   With ofn
      .lStructSize = Len(ofn)
      .hwndOwner = GetActWindow
      .hInstance = 0
      .lpstrInitialDir = fs.buildpath(App.Path, "\TUITION")
      .lpstrFilter = "TUITION DATABASE" & Chr(0) & "TUITION*.mdb" & Chr(0) & Chr(0)
      .nMaxFile = 255
      .lpstrFile = String(255, 0)
      .Flags = OFN_EXPLORER
   End With
End Sub

Function OpenUp() As Boolean
    Dim spath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
   
      OpenUp = False
      
      InitOFN
        
      If GetOFN(ofn) <> 0 Then
        spath = fso.GetParentFolderName(ofn.lpstrFile)
        SrcDBPath = fso.buildpath(spath, fso.GetFileName(ofn.lpstrFile))
        OpenUp = True
      End If
      
End Function


