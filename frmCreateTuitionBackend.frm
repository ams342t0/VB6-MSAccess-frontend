VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCreateNewDatabase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CREATE NEW DATABASE"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   4755
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2400
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog commdlg 
      Left            =   3840
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "CREATE"
      Height          =   735
      Left            =   240
      TabIndex        =   4
      Top             =   1560
      Width           =   4335
   End
   Begin VB.ComboBox cbSemester 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   3
      Text            =   "cbSemester"
      Top             =   840
      Width           =   2895
   End
   Begin VB.TextBox txtSchoolyear 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   1680
      TabIndex        =   1
      Text            =   "txtSchoolyear"
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "SEMESTER"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "SCHOOL YEAR"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1815
   End
End
Attribute VB_Name = "frmCreateNewDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fso As Object
Dim txttemplatepath As String
Dim semstring As String

Private Sub cbSemester_Click()
    Select Case cbSemester.ListIndex
        Case 0
            semstring = "FIRST SEMESTER"
        
        Case 1
            semstring = "SECOND SEMESTER"
    End Select
End Sub

Private Sub cmdCreate_Click()
    
    commdlg.FileName = "TUITION-" & txtSchoolyear.Text & "-" & cbSemester.ListIndex + 1
    commdlg.InitDir = fso.BuildPath(App.Path, "TUITION")
    commdlg.Filter = "ACCESS DATABASE|*.mdb"
    commdlg.CancelError = True
    
    On Error Resume Next
    
    commdlg.ShowSave
    
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical + vbOKOnly, "ERROR CREATING DATABASE"
    Else
        Err.Clear
        
        CopyTemplate
        
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbCritical + vbOKOnly, "ERROR CREATING DATABASE"
        End If
        
        Err.Clear
        ChangeGlobalData
        
        If Err.Number <> 0 Then
            MsgBox Err.Description, vbCritical + vbOKOnly, "ERROR CREATING DATABASE"
        End If
        
    End If
End Sub

Private Sub Form_Load()
    Set fso = CreateObject("Scripting.Filesystemobject")
    
    cbSemester.Clear
    cbSemester.AddItem "FIRST"
    cbSemester.AddItem "SECOND"
    cbSemester.ListIndex = 0
    txtSchoolyear.Text = "2011-2012"
    
End Sub

Sub CopyTemplate()
    txttemplate = fso.BuildPath(App.Path, "TUITIONTEMPLATE.mdb")
    
    If Not fso.FileExists(txttemplate) Then
        If MsgBox("Cannot find template. Look it up?", vbYesNo + vbExclamation) = vbYes Then
            commdlg.InitDir = fso.BuildPath(App.Path, "TUITION")
            commdlg.Filter = "BACKEND TEMPLATE|TUITIONTEMPLATE.mdb"
            commdlg.CancelError = True
            
            On Error Resume Next
            
            commdlg.ShowOpen
            
            If Err.Number = 0 Then
                txttemplate = commdlg.FileName
            End If
        End If
    End If
    
    If fso.FileExists(commdlg.FileName) Then
        If MsgBox("A file with the same name already exists. Continue and overwrite existing file?", vbExclamation + vbYesNo, "CREATE DATABASE") = vbNo Then
            Exit Sub
        End If
    End If
    
    fso.CopyFile txttemplate, commdlg.FileName, True
    
    MsgBox "Template copied", vbOKOnly, "CREATE DATABASE"
End Sub

Sub ChangeGlobalData()
    Dim cn As ADODB.Connection
    
    Set cn = CreateObject("ADODB.CONNECTION")
    
    cn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & commdlg.FileName
    cn.CursorLocation = adUseClient
    cn.Mode = adModeReadWrite
    
    cn.Open
    
    cn.Execute "INSERT INTO globals VALUES('" & "S.Y. " & txtSchoolyear.Text & "','')"
    cn.Execute "INSERT INTO globals VALUES('" & semstring & "','')"
    
    cn.Close
    
    Set cn = Nothing
End Sub
