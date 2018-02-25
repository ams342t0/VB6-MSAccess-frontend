VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "BOOK ENTRIES"
   ClientHeight    =   5520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6675
   BeginProperty Font 
      Name            =   "Arial Narrow"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5520
   ScaleWidth      =   6675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSave 
      Caption         =   "&SAVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4455
      TabIndex        =   5
      Top             =   720
      Width           =   1320
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "&REMOVE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4455
      TabIndex        =   6
      Top             =   1350
      Width           =   1320
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "&ADD"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   510
      Left            =   4455
      TabIndex        =   4
      Top             =   135
      Width           =   1320
   End
   Begin VB.Frame Frame2 
      Caption         =   "TUITION FEES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3480
      Left            =   45
      TabIndex        =   8
      Top             =   2025
      Width           =   6585
      Begin MSComctlLib.ListView lvTuition 
         Height          =   3165
         Left            =   90
         TabIndex        =   7
         Top             =   225
         Width           =   6405
         _ExtentX        =   11298
         _ExtentY        =   5583
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Book Info"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1905
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   3840
      Begin VB.ComboBox cbLevel 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1125
         TabIndex        =   1
         Top             =   225
         Width           =   2445
      End
      Begin VB.TextBox txtAmount 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1125
         TabIndex        =   3
         Top             =   990
         Width           =   1410
      End
      Begin VB.TextBox txtDescription 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1125
         TabIndex        =   2
         Top             =   585
         Width           =   2445
      End
      Begin VB.Label txtTotalAmount 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   495
         TabIndex        =   12
         Top             =   1395
         Width           =   3075
      End
      Begin VB.Label Label2 
         Caption         =   "Tuition Fee"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   180
         TabIndex        =   11
         Top             =   315
         Width           =   870
      End
      Begin VB.Label Label4 
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   495
         TabIndex        =   10
         Top             =   1035
         Width           =   555
      End
      Begin VB.Label Label1 
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   180
         TabIndex        =   9
         Top             =   630
         Width           =   870
      End
   End
End
Attribute VB_Name = "frmBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dbRSet As Object
Dim bItemSelect As Boolean
Dim bDirty As Boolean

Sub UpdateButtons()
    Dim v As Boolean
    
    v = (Len(cbLevel.Text) > 0 And Len(txtDescription.Text) > 0 And Len(txtAmount.Text) > 0)
    cmdAdd.Enabled = v
    cmdSave.Enabled = v And bItemSelect And bDirty
    cmdRemove.Enabled = v And bItemSelect

End Sub

Sub LoadLevelList()
    
    cbLevel.Clear
    
    Set dbRSet = cnTuition.Execute("select TFText from tblTuitionID order by TFID")
    
    dbRSet.MoveFirst
    
    While Not dbRSet.EOF
        cbLevel.AddItem dbRSet.Fields("TFText")
        dbRSet.MoveNext
    Wend
    
    Set dbRSet = Nothing
    
    Exit Sub
    
    cbLevel.AddItem "Toddler Tuition Fees"
    cbLevel.AddItem "Nursery Tuition Fees"
    cbLevel.AddItem "Kinder I Tuition Fees"
    cbLevel.AddItem "Kinder II Tuition Fees"
    cbLevel.AddItem "Grade 1 Tuition Fees"
    cbLevel.AddItem "Grade 2 Tuition Fees"
    cbLevel.AddItem "Grade 3 Tuition Fees"
    cbLevel.AddItem "Grade 4 Tuition Fees"
    cbLevel.AddItem "Grade 5 Tuition Fees"
    cbLevel.AddItem "Grade 6 Tuition Fees"
    cbLevel.AddItem "First Year Tuition Fees"
    cbLevel.AddItem "Second Year Tuition Fees"
    cbLevel.AddItem "Third Year Tuition Fees"
    cbLevel.AddItem "Fourth Year Tuition Fees"
    cbLevel.AddItem "Nursery-NEW Tuition Fees"
    cbLevel.AddItem "Kinder I-NEW Tuition Fees"
    cbLevel.AddItem "Kinder II-NEW Tuition Fees"
    cbLevel.AddItem "Grade 1-NEW Tuition Fees"
    cbLevel.AddItem "Grade 2-NEW Tuition Fees"
    cbLevel.AddItem "Grade 3-NEW Tuition Fees"
    cbLevel.AddItem "Grade 4-NEW Tuition Fees"
    cbLevel.AddItem "Grade 5-NEW Tuition Fees"
    cbLevel.AddItem "Grade 6-NEW Tuition Fees"
    cbLevel.AddItem "First Year-NEW Tuition Fees"
    cbLevel.AddItem "Second Year-NEW Tuition Fees"
    cbLevel.AddItem "Third Year-NEW Tuition Fees"
    cbLevel.AddItem "Fourth Year-NEW Tuition Fees"
    
End Sub

Sub InitTuitionList()
    lvTuition.View = lvwReport
    lvTuition.FullRowSelect = True
    lvTuition.GridLines = True
    lvTuition.LabelEdit = lvwManual
    lvTuition.HideSelection = False
    lvTuition.ColumnHeaders.Add 1, , "Tuition Fee", 3000
    lvTuition.ColumnHeaders.Add 2, , "Amount", 1000, lvwColumnRight
End Sub

Private Sub cbLevel_Change()
    UpdateButtons
    bDirty = True
End Sub

Sub ClearFields()
    txtDescription = ""
    txtAmount = ""
End Sub

Private Sub cbLevel_Click()
    UpdateButtons
    ListTuition
    Me.txtTotalAmount.Caption = "TOTAL: " & Format$(Me.GetBooksTotal, "P #,#.00")
    ClearFields
    bItemSelect = False
End Sub

Private Sub cmdAdd_Click()
    AddTuition
    ListTuition
    MsgBox "Tuition added", vbOKOnly + vbApplicationModal, "TUITION ADD"
    bDirty = False
    ClearFields
    UpdateButtons
End Sub

Private Sub cmdRemove_Click()
    If MsgBox("Delete this tuition entry?", vbYesNo + vbExclamation + vbApplicationModal, "DELETE TUITION") = vbYes Then
        RemoveTuition
        ListTuition
        MsgBox "Tuition Removed", vbOKOnly + vbApplicationModal, "TUITION DELETE"
        ClearFields
        UpdateButtons
    End If
End Sub

Private Sub cmdSave_Click()
    If MsgBox("Save changes?", vbYesNo + vbExclamation + vbApplicationModal, "SAVE TUITION FEE") = vbYes Then
        SaveTuition
        ListTuition
        MsgBox "Saved", vbOKOnly + vbApplicationModal, "TUITION SAVE"
        bDirty = False
        UpdateButtons
    End If
End Sub

Private Sub Form_Load()
    LoadLevelList
    InitTuitionList
    
    cmdAdd.Enabled = False
    cmdSave.Enabled = False
    cmdRemove.Enabled = False
    bItemSelect = False
    bDirty = False
End Sub

Private Sub lvTuition_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdSave.Enabled = True
    cmdRemove.Enabled = True
    
    txtDescription.Text = Item.Text
    txtAmount.Text = Format$(Item.SubItems(1), "#,#.00")
    bItemSelect = True
    bDirty = False
    UpdateButtons
End Sub

Private Sub txtAmount_Change()
    UpdateButtons
    bDirty = True
End Sub

Private Sub txtQuantity_Change()
    UpdateButtons
    bDirty = True
End Sub

Private Sub txtDescription_Change()
    UpdateButtons
    bDirty = True
End Sub


Sub AddTuition()
                    
    cnTuition.Execute "INSERT INTO tblTuitionFee (TFREFNUMBER,TFTEXT,TFTYPE,TFPRICE) VALUES (" & _
                    cbLevel.ListIndex + 1 & ",""" & txtDescription.Text & """," & _
                    0 & "," & Val(txtAmount.Text) & ")"
                    
End Sub

Sub SaveTuition()
    cnTuition.Execute "UPDATE tblTuitionFee SET " & _
                   "TFTEXT = """ & txtDescription & """," & _
                   "TFPRICE = " & Val(txtAmount) & _
                   " WHERE TFREFNUMBER= " & cbLevel.ListIndex + 1 & " AND TFTEXT = """ & txtDescription & """"
                                            
End Sub

Sub RemoveTuition()
    cnTuition.Execute "DELETE FROM tblTuitionFee WHERE TFREFNUMBER = " & cbLevel.ListIndex + 1 & " AND TFTEXT LIKE """ & txtDescription & """"
End Sub


Sub ListTuition()
    Dim idx As Long
    
    Set dbRSet = cnTuition.Execute("SELECT * FROM tblTuitionfee WHERE TFREFNUMBER = " & cbLevel.ListIndex + 1 & " ORDER BY LISTORDER")
    
    lvTuition.ListItems.Clear
    
    idx = 1
    If dbRSet.RecordCount > 0 Then
        While Not dbRSet.EOF
            lvTuition.ListItems.Add idx, , dbRSet.Fields("TFTEXT")
            lvTuition.ListItems(idx).SubItems(1) = Format$(dbRSet.Fields("TFPRICE"), "#,#.00")
            lvTuition.ListItems(idx).Tag = dbRSet.Fields("TFREFNUMBER")
            idx = idx + 1
            dbRSet.MoveNext
        Wend
    End If
    
    Set dbRSet = Nothing
End Sub

Function GetBooksTotal() As Double
    
    On Error Resume Next

    Set dbRSet = cnTuition.Execute("SELECT SUM(TFPRICE) as x FROM tblTuitionfee WHERE TFREFNUMBER = " & cbLevel.ListIndex + 1)
    
    If dbRSet.EOF And dbRSet.BOF Then
        GetBooksTotal = 0
    Else
        GetBooksTotal = dbRSet.Fields("x")
    End If
    
    Set dbRSet = Nothing
    
End Function
