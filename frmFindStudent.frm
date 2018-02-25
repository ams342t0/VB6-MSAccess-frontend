VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "FIND"
   ClientHeight    =   5295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5295
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      Caption         =   "&CLOSE"
      Height          =   465
      Left            =   4815
      TabIndex        =   7
      Top             =   630
      Width           =   1455
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&FIND"
      Height          =   465
      Left            =   4800
      TabIndex        =   6
      Top             =   90
      Width           =   1455
   End
   Begin MSComctlLib.ListView lvList 
      Height          =   3885
      Left            =   45
      TabIndex        =   5
      Top             =   1350
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   6853
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Frame Frame1 
      Caption         =   "FIND"
      Height          =   1140
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   4560
      Begin VB.TextBox txtInvoiceNumber 
         Height          =   330
         Left            =   1440
         TabIndex        =   4
         Top             =   630
         Width           =   3030
      End
      Begin VB.TextBox txtName 
         Height          =   330
         Left            =   1440
         TabIndex        =   3
         Top             =   270
         Width           =   3030
      End
      Begin VB.OptionButton opInvoiceNumber 
         Caption         =   "IN&VOICE NO."
         Height          =   375
         Left            =   135
         TabIndex        =   2
         Top             =   630
         Width           =   1275
      End
      Begin VB.OptionButton opName 
         Caption         =   "N&AME"
         Height          =   285
         Left            =   135
         TabIndex        =   1
         Top             =   315
         Width           =   915
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sqlstring As String

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim sqlfindstring As String
    Dim pvrset As ADODB.Recordset
    Dim idx As Long
    
    If opName Then
        sqlfindstring = " WHERE instr(1,engname,""" & txtName.Text & """)>0"
    Else
        sqlfindstring = " WHERE iv.invoicenumber = " & Val(Me.txtInvoiceNumber)
    End If
    
    On Error Resume Next

    Set pvrset = cnTuition.Execute(sqlstring & sqlfindstring)
    
    With pvrset
        If .EOF And .BOF Then
            MsgBox "No match found."
            Set pvrset = Nothing
            Exit Sub
        End If
        
        lvList.ListItems.Clear
        idx = 1
        
        .MoveFirst
        
        While Not .EOF
            
            lvList.ListItems.Add idx, , Format$(.Fields("STUDREFNUMBER"), "00000")
            lvList.ListItems(idx).SubItems(1) = .Fields("ENGNAME")
            lvList.ListItems(idx).SubItems(2) = Format$(.Fields("INVOICENUMBER"), "00000")
            lvList.ListItems(idx).SubItems(3) = Format$(.Fields("AMOUNTDUE"), "P #,#.00")
            lvList.ListItems(idx).SubItems(4) = Format$(.Fields("STAMPTIME"), "MM-DD-YYYY")
            lvList.ListItems(idx).SubItems(5) = Format$(.Fields("PAYAMOUNT"), "P #,#.00")
            lvList.ListItems(idx).SubItems(6) = Format$(.Fields("TIMESTAMP"), "MM-DD-YYYY")
            
            idx = idx + 1
            .MoveNext
        Wend
        
    End With
        
End Sub

Private Sub Form_Load()
    sqlstring = "select ml.studrefnumber,ml.engname,iv.invoicenumber,iv.amountdue,iv.stamptime,pl.payamount,pl.timestamp from (tblmasterlist as ml left join tblinvoice as iv on ml.studrefnumber=iv.studrefnumber) left join tblPayments as pl on ml.studrefnumber=pl.studrefnumber"
    opName.Value = True
    
    With Me.lvList
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .LabelEdit = lvwManual
        .ColumnHeaders.Add 1, , "Ref. No.", 700
        .ColumnHeaders.Add 2, , "Name", 2500
        .ColumnHeaders.Add 3, , "Invoice No.", 700
        .ColumnHeaders.Add 4, , "Amount Due", 1100
        .ColumnHeaders.Add 5, , "Invoice Date", 1000
        .ColumnHeaders.Add 6, , "Amount Paid", 1100
        .ColumnHeaders.Add 7, , "Date Paid", 1000
    End With
End Sub

Private Sub opInvoiceNumber_Click()
    txtName.Enabled = False
    txtInvoiceNumber.Enabled = True
End Sub

Private Sub opName_Click()
    txtName.Enabled = True
    txtInvoiceNumber.Enabled = False
End Sub
