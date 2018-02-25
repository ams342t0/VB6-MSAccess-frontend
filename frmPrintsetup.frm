VERSION 5.00
Begin VB.Form frmPrintsetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Print Setup"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3315
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1365
   ScaleWidth      =   3315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdPrint 
      Caption         =   "&PRINT"
      Height          =   510
      Left            =   90
      TabIndex        =   3
      Top             =   735
      Width           =   1500
   End
   Begin VB.CommandButton cmdClosePrint 
      Caption         =   "&CLOSE"
      Height          =   510
      Left            =   1710
      TabIndex        =   2
      Top             =   735
      Width           =   1500
   End
   Begin VB.ComboBox cbPrinter 
      Height          =   315
      Left            =   105
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   300
      Width           =   3120
   End
   Begin VB.Label Label1 
      Caption         =   "Printer name"
      Height          =   285
      Left            =   120
      TabIndex        =   1
      Top             =   75
      Width           =   1455
   End
End
Attribute VB_Name = "frmPrintsetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cbPrinter_Click()
    strPrinter = cbPrinter.Text
    SaveSetting "VBENROLL", "SETTING", "TARGETPRINTER", strPrinter
End Sub

Private Sub cbPrinter_DropDown()
    Dim pvrset As ADODB.Recordset
    
    cbPrinter.Clear
    
    Set pvrset = cnTuition.Execute("SELECT printer from tblPrinter")
    
    With pvrset
        .MoveFirst
        While Not .EOF
            cbPrinter.AddItem pvrset.Fields("printer")
            .MoveNext
        Wend
    End With
    
    Set pvrset = Nothing
End Sub


Private Sub cmdClosePrint_Click()
    Unload Me
End Sub

Private Sub cmdPrint_Click()
    PrintReceipt nPage2InvoiceNumber
End Sub

Private Sub Form_Load()
    cbPrinter.Text = strPrinter
End Sub
