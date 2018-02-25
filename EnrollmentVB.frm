VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "ENROLLMENT"
   ClientHeight    =   8235
   ClientLeft      =   60
   ClientTop       =   -1095
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   ScaleHeight     =   727.152
   ScaleMode       =   0  'User
   ScaleWidth      =   785.436
   StartUpPosition =   1  'CenterOwner
   WindowState     =   2  'Maximized
   Begin MSComDlg.CommonDialog cdDialog 
      Left            =   6840
      Top             =   7920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.StatusBar sbStatbar 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   44
      Top             =   7800
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   7545
      Left            =   60
      TabIndex        =   0
      Top             =   165
      Width           =   11145
      Begin VB.Frame frInvoice 
         Caption         =   "Invoice No.:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   90
         TabIndex        =   43
         Top             =   180
         Width           =   5160
         Begin VB.OptionButton Option1 
            Caption         =   "(&3) Use last"
            Height          =   255
            Left            =   135
            TabIndex        =   1
            Top             =   450
            Width           =   1515
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(&2) Use existing"
            Height          =   495
            Left            =   135
            TabIndex        =   2
            Top             =   630
            Width           =   1830
         End
         Begin VB.OptionButton Option3 
            Caption         =   "(&1) New"
            Height          =   255
            Left            =   135
            TabIndex        =   4
            Top             =   1080
            Width           =   1980
         End
         Begin VB.ComboBox cbP1InvoiceNumber 
            Height          =   315
            Left            =   2160
            TabIndex        =   3
            Top             =   720
            Width           =   2880
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5775
         Left            =   90
         TabIndex        =   29
         Top             =   1665
         Width           =   5145
         Begin VB.ComboBox cbP1Student 
            Height          =   315
            Left            =   765
            TabIndex        =   5
            Top             =   225
            Width           =   4245
         End
         Begin VB.ComboBox cbP1EnglishLevel 
            Height          =   315
            Left            =   765
            TabIndex        =   6
            Top             =   630
            Width           =   3075
         End
         Begin VB.Frame Frame4 
            Caption         =   "DETAILS"
            Height          =   1275
            Left            =   90
            TabIndex        =   31
            Top             =   1035
            Width           =   4965
            Begin VB.Label txtStudID 
               Caption         =   "STUDENT ID"
               Height          =   285
               Left            =   135
               TabIndex        =   37
               Top             =   585
               Width           =   1140
            End
            Begin VB.Label txtStudName 
               Caption         =   "NAME"
               Height          =   285
               Left            =   135
               TabIndex        =   36
               Top             =   270
               Width           =   3435
            End
            Begin VB.Label txtChiName 
               Caption         =   "CHINAME"
               Height          =   285
               Left            =   3870
               TabIndex        =   35
               Top             =   270
               Width           =   1005
            End
            Begin VB.Label txtSex 
               Caption         =   "SEX"
               Height          =   285
               Left            =   1665
               TabIndex        =   34
               Top             =   585
               Width           =   915
            End
            Begin VB.Label txtEngClass 
               Caption         =   "ENGLISHCLASS"
               Height          =   330
               Left            =   135
               TabIndex        =   33
               Top             =   900
               Width           =   1500
            End
            Begin VB.Label txtChiClass 
               Caption         =   "CHINESECLASS"
               Height          =   330
               Left            =   1665
               TabIndex        =   32
               Top             =   900
               Width           =   1770
            End
         End
         Begin MSComctlLib.ListView lvP1TuitionFees 
            Height          =   2895
            Left            =   90
            TabIndex        =   30
            Top             =   2340
            Width           =   4965
            _ExtentX        =   8758
            _ExtentY        =   5106
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label1 
            Caption         =   "Name"
            Height          =   240
            Left            =   135
            TabIndex        =   42
            Top             =   270
            Width           =   555
         End
         Begin VB.Label Label2 
            Caption         =   "Level"
            Height          =   285
            Left            =   135
            TabIndex        =   41
            Top             =   675
            Width           =   600
         End
         Begin VB.Label Label9 
            Caption         =   "TOTAL"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   40
            Top             =   5265
            Width           =   1050
         End
         Begin VB.Label txtP1Total 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   3600
            TabIndex        =   39
            Top             =   5265
            Width           =   1455
         End
         Begin VB.Label txtSaved 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   375
            Left            =   135
            TabIndex        =   38
            Top             =   5265
            Width           =   2265
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "DATE FILTER"
         Height          =   825
         Left            =   5310
         TabIndex        =   27
         Top             =   225
         Width           =   5730
         Begin VB.OptionButton Option4 
            Caption         =   "ALL"
            Height          =   285
            Left            =   180
            TabIndex        =   7
            Top             =   360
            Width           =   780
         End
         Begin VB.OptionButton Option5 
            Caption         =   "Between"
            Height          =   285
            Left            =   1035
            TabIndex        =   8
            Top             =   360
            Width           =   960
         End
         Begin MSComCtl2.DTPicker dpP2DateTo 
            Height          =   330
            Left            =   4050
            TabIndex        =   10
            Top             =   315
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   582
            _Version        =   393216
            Format          =   20643841
            CurrentDate     =   39727
         End
         Begin MSComCtl2.DTPicker dpP2DateFrom 
            Height          =   330
            Left            =   2070
            TabIndex        =   9
            Top             =   315
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   582
            _Version        =   393216
            Format          =   20643841
            CurrentDate     =   39727
         End
         Begin VB.Label Label13 
            Caption         =   "and"
            Height          =   240
            Left            =   3645
            TabIndex        =   28
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame6 
         Height          =   870
         Left            =   5310
         TabIndex        =   26
         Top             =   6570
         Width           =   3300
         Begin VB.CommandButton cmdP1New 
            Caption         =   "N&EW"
            Height          =   600
            Left            =   90
            TabIndex        =   14
            Top             =   180
            Width           =   1005
         End
         Begin VB.CommandButton cmdP1Save 
            Caption         =   "&SAVE"
            Height          =   600
            Left            =   1125
            TabIndex        =   15
            Top             =   180
            Width           =   1005
         End
         Begin VB.CommandButton cmdAdmission 
            Caption         =   "&ADMISSION SLIP"
            Height          =   600
            Left            =   2160
            TabIndex        =   16
            Top             =   165
            Width           =   1050
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "INVOICE"
         Height          =   4515
         Left            =   5310
         TabIndex        =   23
         Top             =   2025
         Width           =   5730
         Begin VB.Frame Frame2 
            Caption         =   "Sort"
            Height          =   600
            Left            =   1035
            TabIndex        =   45
            Top             =   585
            Width           =   4515
            Begin VB.OptionButton opSortByName 
               Caption         =   "Name"
               Height          =   285
               Left            =   1395
               TabIndex        =   47
               Top             =   225
               Width           =   1005
            End
            Begin VB.OptionButton opSortByInvoice 
               Caption         =   "Invoice No."
               Height          =   240
               Left            =   135
               TabIndex        =   46
               Top             =   270
               Width           =   1140
            End
         End
         Begin VB.ComboBox cbP2InvoiceNumber 
            Height          =   315
            Left            =   1035
            TabIndex        =   13
            Top             =   225
            Width           =   4515
         End
         Begin MSComctlLib.ListView lvP2Invoice 
            Height          =   3165
            Left            =   90
            TabIndex        =   24
            Top             =   1260
            Width           =   5550
            _ExtentX        =   9790
            _ExtentY        =   5583
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   0
         End
         Begin VB.Label Label12 
            Caption         =   "Invoice No.:"
            Height          =   285
            Left            =   45
            TabIndex        =   25
            Top             =   225
            Width           =   1005
         End
      End
      Begin VB.CommandButton cmdP2Print 
         Caption         =   "&PRINT INVOICE"
         Height          =   600
         Left            =   8910
         TabIndex        =   17
         Top             =   6660
         Width           =   1005
      End
      Begin VB.CommandButton cmdP2Delete 
         Caption         =   "&DELETE INVOICE"
         Height          =   600
         Left            =   9990
         TabIndex        =   18
         Top             =   6660
         Width           =   960
      End
      Begin VB.Frame Frame9 
         Caption         =   "TIME FILTER"
         Height          =   825
         Left            =   5310
         TabIndex        =   19
         Top             =   1125
         Width           =   5730
         Begin VB.OptionButton Option6 
            Caption         =   "ALL"
            Height          =   285
            Left            =   180
            TabIndex        =   11
            Top             =   360
            Width           =   780
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Between"
            Height          =   330
            Left            =   1035
            TabIndex        =   12
            Top             =   360
            Width           =   1005
         End
         Begin MSComCtl2.DTPicker dpP2TimeTo 
            Height          =   330
            Left            =   4050
            TabIndex        =   20
            Top             =   315
            Width           =   1545
            _ExtentX        =   2725
            _ExtentY        =   582
            _Version        =   393216
            Format          =   20643842
            CurrentDate     =   39727
         End
         Begin MSComCtl2.DTPicker dpP2TimeFrom 
            Height          =   330
            Left            =   2070
            TabIndex        =   21
            Top             =   315
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   582
            _Version        =   393216
            Format          =   20643842
            CurrentDate     =   39727
         End
         Begin VB.Label Label14 
            Caption         =   "and"
            Height          =   285
            Left            =   3645
            TabIndex        =   22
            Top             =   360
            Width           =   420
         End
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpenBackend 
         Caption         =   "Open Backend"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuChangeUser 
         Caption         =   "Change User"
         Shortcut        =   ^G
      End
      Begin VB.Menu mnuFind 
         Caption         =   "Find"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuTuitionFees 
         Caption         =   "Tuition Fees"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuReceipts 
         Caption         =   "Receipts"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuReport 
         Caption         =   "Report"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuPrintersetup 
         Caption         =   "Printer setup..."
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuClose 
         Caption         =   "Close backend"
         Shortcut        =   ^L
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuBackend 
      Caption         =   "&Backend"
      Begin VB.Menu mnuCopyTemplate 
         Caption         =   "Copy Template"
      End
      Begin VB.Menu mnuImportTuitionFees 
         Caption         =   "Import Tuition Fees"
      End
      Begin VB.Menu mnuUpdateMasterlist 
         Caption         =   "Update Masterlist"
      End
      Begin VB.Menu mnuResetDatabase 
         Caption         =   "Reset Database"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
      Begin VB.Menu mnuFirelist 
         Caption         =   "Student List"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuInvoiceList 
         Caption         =   "Invoice List"
         Shortcut        =   {F3}
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public bNewInvoice As Boolean
Public nLastInvoiceNumber As Long
Public nInvoiceNumber As Long
Public nStudRefNumber As Long
Public nTuitionRefNumber As Long
Public bNumPicked As Boolean


'dummy vars
Dim frP1InvoiceNumber As Long
Dim frP2Date As Long
Dim frP2Time As Long

Dim sTotal As Single


Private Sub cbP1InvoiceNumber_DropDown()
    SelectInvoiceNumber cbP1InvoiceNumber, False
End Sub

Private Sub cbP1Student_DropDown()
    Dim pvrset As ADODB.Recordset
    
    On Error Resume Next
    
    Set pvrset = cnTuition.Execute("SELECT engname,engsection,studrefnumber from tblMasterList ORDER BY engname")
    
    cbP1Student.Clear
    
    nIndex = 0
    
    With pvrset
        .MoveFirst
        
        While Not .EOF
            cbP1Student.AddItem .Fields("engname") & " (" & Format$(.Fields("studrefnumber"), "00000") & ")", nIndex
            nIndex = nIndex + 1
            .MoveNext
        Wend
        
    End With
    
    Set pvrset = Nothing
End Sub

Private Sub cbP1Student_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cbP1Student_Click
    End If
End Sub

Private Sub cbP2InvoiceNumber_DropDown()
    SelectInvoiceNumber cbP2InvoiceNumber, False
End Sub

Private Sub cmdAdmission_Click()
    PrintVBForm
    Exit Sub
End Sub


Private Sub PrintVBForm()
    Dim pvrset As ADODB.Recordset
    Dim nEngLevelx As Long, nChiLevelx As Long
    Dim nlvSelRefNum As Long
    Dim strSectionString As String
    Dim tlong As Long
    Dim ni As Long
    Dim nrep As VbMsgBoxResult
    
    On Error Resume Next
    
    nrep = MsgBox("Click YES to print all, click NO to print selection only.", vbYesNo, "PRINT ADMISSION")
    
    If nrep = vbNo Then
        If lvP2Invoice.SelectedItem.Index = lvP2Invoice.ListItems.Count Then
            MsgBox "May lahi ka talaga unggoy, no?", vbOKOnly, "ANG DAMING PWEDENG I-CLICK, ITO PA?"
            Exit Sub
        End If
        ni = lvP2Invoice.SelectedItem.Index
    End If
    
    For ni = 1 To lvP2Invoice.ListItems.Count - 1
        
        If nrep = vbNo Then
            If ni <> lvP2Invoice.SelectedItem.Index Then
                GoTo skipme
            End If
        End If
        
        nlvSelRefNum = Val(Me.lvP2Invoice.ListItems(ni).SubItems(6))
        
        Set pvrset = cnTuition.Execute("select invoicenumber,ml.studrefnumber,engname,chiname,engsection,chisection,chilevel from tblMasterlist as ml inner join tblInvoice as inv on ml.studrefnumber=inv.studrefnumber where ml.STUDREFNUMBER = " & nlvSelRefNum)
            
        frmPrintAd.txtlabel(2).Caption = ""
        frmPrintAd.txtlabel(8).Caption = ""
            
        frmPrintAd.txtlabel(1).Caption = strSchoolYear
        frmPrintAd.txtlabel(5).Caption = strSchoolYear
        
        frmPrintAd.txtlabel(0).Caption = Format$(pvrset.Fields("invoicenumber"), "0000")
        frmPrintAd.txtlabel(9).Caption = Format$(pvrset.Fields("invoicenumber"), "0000")
        
        frmPrintAd.txtlabel(4).Caption = pvrset.Fields("chiname")
            
        strSectionString = pvrset.Fields("chisection")
        tlong = Len(strSectionString)
            
        If pvrset.Fields("chilevel") = 1 Then
                frmPrintAd.txtlabel(3).Caption = Mid$(strSectionString, 1, 3)
                frmPrintAd.txtlabel(6).Caption = ""
        Else
                frmPrintAd.txtlabel(3).Caption = Mid$(strSectionString, 1, 2)
                frmPrintAd.txtlabel(6).Caption = Mid$(strSectionString, 3, tlong - 2)
        End If
            
        frmPrintAd.txtlabel(7).Caption = Format$(Now, "mm/dd/yyyy")
        frmPrintAd.txtlabel(12).Caption = Format$(Now, "mm/dd/yyyy")
        
        frmPrintAd.txtlabel(10).Caption = pvrset.Fields("engname")
        frmPrintAd.txtlabel(11).Caption = pvrset.Fields("engsection")

        Printer.Orientation = 2
        Printer.PaperSize = 1
        frmPrintAd.PrintForm
        Printer.EndDoc
skipme:
    Next
    
hell:
    Set pvrset = Nothing
End Sub



'**************************************************************
'G L O B A L  -  M A I N
'**************************************************************

Sub DisableControls()
    Dim c As Object
    
    On Error Resume Next
    
    For Each c In Me.Controls
        c.Enabled = False
    Next
    
    Me.mnuFile.Enabled = True
    Me.mnuOpenBackend.Enabled = True
    Me.mnuAbout.Enabled = True
    mnuBackend.Enabled = True
    mnuCopyTemplate.Enabled = True
End Sub

Sub EnableControls()
    Dim c As Object
    
    On Error Resume Next
    
    For Each c In Me.Controls
        c.Enabled = True
    Next
    
    Me.mnuOpenBackend.Enabled = False
    Me.cmdP2Print.Enabled = False
    Me.cmdP2Delete.Enabled = False
    cmdAdmission.Enabled = False
    Me.cmdP1Save.Enabled = False
    
    If frP1InvoiceNumber = 2 Then
        Me.cbP1InvoiceNumber.Enabled = True
    Else
        Me.cbP1InvoiceNumber.Enabled = False
    End If
    
End Sub




Private Sub Form_Load()
    Set fs = CreateObject("Scripting.FileSystemObject")
    
    DisableControls
    
    mnuOpenBackend_Click
    
    InitDefaultValues
    InitLevelComboBoxes
    InitListViews
    InitInvoiceReportListView
    
    nLastInvoiceNumber = nInvoiceNumber
    
    sbStatbar.Panels(1).Width = 8000

    strPrinter = GetSetting("VBENROLL", "SETTINGS", "TARGETPRINTER", "LPT1")
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If bNumPicked = True Then
         ReturnPickedNumber nInvoiceNumber
         bNumPicked = False
    End If
    
    On Error Resume Next
    'Unload frmPrintAd
End Sub


Sub InitDefaultValues()
    bNumPicked = False

    frP1InvoiceNumber = 3
    strPrinter = "local"
    
    dpP2DateFrom.Value = Now
    dpP2DateTo.Value = Now
    dpP2TimeFrom.Value = #7:00:00 AM#
    dpP2TimeTo.Value = #12:00:00 PM#
    
    
    opSortByInvoice.Value = True
    Option3.Value = True
    Option4.Value = True
    Option6.Value = True
    
End Sub


'**************************************************************
'G L O B A L  E V E N T S
'**************************************************************



'**************************************************************
'P A G E  1  -  I N V O I C E  P A G E
'**************************************************************


Function GetStudentNumber(ByRef nstring As String) As Long
    Dim pt1 As Long
    Dim pt2 As Long
    
    pt1 = InStr(1, nstring, "(", vbTextCompare)
    pt2 = InStr(1, nstring, ")", vbTextCompare)
    
    GetStudentNumber = Val(Mid$(nstring, pt1 + 1, pt2 - pt1 + 1))
End Function

Function GetInvoiceNumber(ByRef nstring As String) As Long
    Dim pt1 As Long
    
    pt1 = InStr(1, nstring, "-", vbTextCompare)
    
    GetInvoiceNumber = Val(Mid$(nstring, 1, pt1 - 2))
End Function



Sub InitLevelComboBoxes()
    Dim pvrset As ADODB.Recordset
    
    'On Error Resume Next
    
    Set pvrset = cnTuition.Execute("SELECT tftext from tblTuitionID order by TFID")
    
    cbP1EnglishLevel.Clear
    
    With pvrset
        .MoveFirst
        
        While Not .EOF
            cbP1EnglishLevel.AddItem .Fields("tftext")
            .MoveNext
        Wend
    
    End With
    
    Set pvrset = Nothing
    
End Sub

Sub InitListViews()

    With lvP1TuitionFees
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "FEE DESCRIPTION", 1500
        .ColumnHeaders.Add 2, , "AMOUNT", 2000, lvwColumnRight
        .ListItems.Clear
    End With
End Sub

Sub ListTuitionFees()
    Dim pvrset As ADODB.Recordset
    
    nTuitionRefNumber = cbP1EnglishLevel.ListIndex + 1
    
    lvP1TuitionFees.ListItems.Clear
    txtP1Total.Caption = ""
    
        
    On Error Resume Next
    
    Set pvrset = cnTuition.Execute("SELECT TFTEXT,TFPRICE from tblTuitionFee WHERE TFREFNUMBER = " & nTuitionRefNumber)
    
    With pvrset
        
        If (.EOF And .BOF) Then
            cmdP1Save.Enabled = False
            Set pvrset = Nothing
            Exit Sub
        End If
        
        .MoveFirst
        
        nIndex = 1
        sTotal = 0
        
        While Not .EOF
            lvP1TuitionFees.ListItems.Add nIndex, , .Fields("TFTEXT")
            lvP1TuitionFees.ListItems(nIndex).SubItems(1) = Format$(.Fields("TFPRICE"), "#,0.00")
            sTotal = sTotal + .Fields("TFPRICE")
            
            nIndex = nIndex + 1
            .MoveNext
        Wend
    End With
    
    txtP1Total.Caption = Format$(sTotal, "#,0.00")
    
    cmdP1Save.Enabled = True
    
    Set pvrset = Nothing
    
End Sub

Function SaveTuitionFee() As Boolean
    Dim tmpstr As String
    
    SaveTuitionFee = True
    
    tmpstr = "INSERT INTO tblInvoice (INVOICENUMBER,STUDREFNUMBER,TFREFNUMBER,AMOUNTDUE,STAMPTIME,TUSER) VALUES (" & _
                      nInvoiceNumber & "," & nStudRefNumber & "," & nTuitionRefNumber & "," & sTotal & "," & EnDate(Now) & "," & EnQuote(strOperator) & ")"
                          
    cnTuition.Execute (tmpstr)
    
End Function

Function InvoiceExists() As Boolean
    Dim pvrset As ADODB.Recordset
    
    Set pvrset = cnTuition.Execute("SELECT invoicenumber from tblInvoice WHERE STUDREFNUMBER = " & nStudRefNumber & " AND TFREFNUMBER=" & nTuitionRefNumber)
    
    With pvrset
        If .EOF And .BOF Then
            InvoiceExists = False
        Else
            InvoiceExists = True
        End If
    End With
    
    Set pvrset = Nothing
    
End Function


Sub NewInvoice()
    bNewInvoice = True
    
    cmdP1Save.Enabled = False
    
    cbP1Student = ""
    cbP1EnglishLevel = ""
    
    lvP1TuitionFees.ListItems.Clear
    txtP1Total.Caption = ""
    txtSaved.Caption = ""
    sTotal = 0
    
    ClearStudentInfo
End Sub


Sub SelectInvoiceNumber(ByRef combobox As Object, ByVal drop As Boolean)
    Dim pvrset As ADODB.Recordset
    
    
    If opSortByInvoice Then
        Set pvrset = cnTuition.Execute("SELECT i.INVOICENUMBER, s.engname FROM tblINVOICE AS i INNER JOIN tblMasterList AS s ON i.studrefnumber=s.studrefnumber ORDER BY InvoiceNumber desc")
    Else
        Set pvrset = cnTuition.Execute("SELECT i.INVOICENUMBER, s.engname FROM tblINVOICE AS i INNER JOIN tblMasterList AS s ON i.studrefnumber=s.studrefnumber ORDER BY s.EngName asc")
    End If
    
    combobox.Clear
    
    With pvrset
        
        If .EOF And .BOF Then
            combobox.Enabled = False
            Set pvrset = Nothing
            Exit Sub
        End If
        
        .MoveFirst
        
        
        nIndex = 0
        
        While Not .EOF
            combobox.AddItem Format$(.Fields("INVOICENUMBER"), "00000") & " - " & .Fields("engname"), nIndex
            nIndex = nIndex + 1
            .MoveNext
        Wend
        
    End With
    
    combobox.Enabled = True
    
    Set pvrset = Nothing

End Sub



'****************************************************************************************
'VB EVENTS
'****************************************************************************************

Private Sub frP1InvoiceNumberSelect()
   
    Select Case frP1InvoiceNumber
        Case 1
            If bNumPicked = True Then
                ReturnPickedNumber nInvoiceNumber
                bNumPicked = False
            End If
            
            cbP1InvoiceNumber.Enabled = False
            nInvoiceNumber = nLastInvoiceNumber
            frInvoice.Caption = "Invoice No.: " & Format$(nInvoiceNumber, "00000")
            
        Case 2
            If bNumPicked = True Then
                ReturnPickedNumber nInvoiceNumber
                bNumPicked = False
            End If
        
            cbP1InvoiceNumber.Enabled = True
            SelectInvoiceNumber cbP1InvoiceNumber, True
        
        Case 3
            If bNumPicked Then
                If MsgBox("Invoice number is available. Pick new one?", vbYesNo + vbInformation, "NEW INVOICE NUMBER") = vbYes Then
                    Do
                        nInvoiceNumber = PickInvoiceNumber
                        If nInvoiceNumber > 0 Then
                            Exit Do
                        End If
                    Loop While True
                End If
            Else
                    Do
                        nInvoiceNumber = PickInvoiceNumber
                        If nInvoiceNumber > 0 Then
                            Exit Do
                        End If
                    Loop While True
            End If
        
            cbP1InvoiceNumber.Enabled = False
            frInvoice.Caption = "Invoice No.: " & Format$(nInvoiceNumber, "00000")
            bNumPicked = True
    End Select
End Sub


Private Sub cbP1EnglishLevel_Click()
    ListTuitionFees
End Sub


Private Sub cbP1InvoiceNumber_Click()
    nInvoiceNumber = CLng(GetInvoiceNumber(cbP1InvoiceNumber.Text))
    frInvoice.Caption = "Invoice No.: " & Format$(nInvoiceNumber, "00000")
End Sub


Private Sub cbP1Student_Click()
    Dim pvrset As ADODB.Recordset
    Dim nEngLevelx As Long, nChiLevelx As Long
    
    nStudRefNumber = GetStudentNumber(cbP1Student.Text)
    
    On Error Resume Next
    
    Set pvrset = cnTuition.Execute("SELECT englevel,chilevel,idnum,engname,chiname,engsection,chisection,sex,isnew from tblMasterList where STUDREFNUMBER = " & nStudRefNumber)
    
    nEngLevelx = pvrset.Fields("englevel")
    nChiLevelx = pvrset.Fields("chilevel")
    
    If nEngLevelx > nChiLevelx Then
        If nEngLevelx < 15 Then
           cbP1EnglishLevel.ListIndex = nEngLevelx - 1
        Else
           cbP1EnglishLevel.ListIndex = nChiLevelx - 1
        End If
    Else
        If nChiLevelx < 15 Then
           cbP1EnglishLevel.ListIndex = nChiLevelx - 1
        Else
           cbP1EnglishLevel.ListIndex = nEngLevelx - 1
        End If
    End If
    
    If pvrset.Fields("isnew") And ((nEngLevelx + nChiLevelx) > 2) Then
        cbP1EnglishLevel.ListIndex = cbP1EnglishLevel.ListIndex + 13
    End If
    
    If nEngLevelx = 14 And nChiLevelx < 14 Then
        cbP1EnglishLevel.ListIndex = 28
        MsgBox "Irregular 4th Year Student.", vbOKOnly + vbExclamation, "NOTE"
    End If
    
    
    ListTuitionFees
    
    ClearStudentInfo
    
    With pvrset
        txtStudID.Caption = .Fields("idnum")
        txtStudName.Caption = .Fields("engname")
        txtSex.Caption = .Fields("sex")
        txtEngClass.Caption = .Fields("engsection")
        txtChiClass.Caption = .Fields("chisection")
        txtChiName.Caption = .Fields("chiname")
    End With
        
    Set pvrset = Nothing
    
End Sub

Sub ClearStudentInfo()
    txtStudID.Caption = ""
    txtStudName.Caption = ""
    txtChiName.Caption = ""
    txtSex.Caption = ""
    txtEngClass.Caption = ""
    txtChiClass.Caption = ""
End Sub


Private Sub cmdP1New_Click()
    NewInvoice
    frP1InvoiceNumberSelect
End Sub


Private Sub cmdP1Save_Click()

    If MsgBox("Confirm save.", vbOKCancel, "SAVE INVOICE") = vbOK Then
        
        If InvoiceExists Then
            MsgBox "INVOICE EXISTS", vbOKOnly, "INVOICE ALREADY EXISTS"
            Exit Sub
        End If
            
        SaveTuitionFee
        
        nLastInvoiceNumber = nInvoiceNumber
        cmdP1Save.Enabled = False
        
        bNumPicked = False
        txtSaved.Caption = "INVOICE CREATED"
        
    End If
    
End Sub








'**************************************************************
'P A G E  2  -  I N V O I C E  R E P O R T
'**************************************************************

Private Sub cmdP2Print_Click()
    frmPrintsetup.Show vbModal, Me
End Sub


Private Sub cmdP2Delete_Click()
        
    If lvP2Invoice.ListItems.Count = 0 Then
        Exit Sub
    End If
        
    If MsgBox("Confirm delete.", vbOKCancel, "DELETE INVOICE") = vbOK Then
        cnTuition.Execute "DELETE from tblInvoice WHERE STUDREFNUMBER = " & CLng(lvP2Invoice.SelectedItem.SubItems(6))
        cmdP2Delete.Enabled = False
        UpdateInvoiceList
    End If
    
End Sub

Sub UpdateInvoiceList()
    Dim tmpstr As String
    Dim sGrandTotal As Single
    Dim pvrset As ADODB.Recordset
    
    tmpstr = "SELECT  iv.invoicenumber,ml.engname,iv.stamptime,iv.amountdue,iv.studrefnumber,tid.tftext,iv.studrefnumber,ml.isnew as newstud " & _
             "from (tblInvoice as iv inner join tblMasterList as ml on iv.studrefnumber=ml.studrefnumber) " & _
             "INNER JOIN tblTuitionID as tID on iv.tfrefnumber=tid.tfid"
    
    tmpstr = tmpstr & " WHERE iv.invoicenumber=" & nPage2InvoiceNumber
    
    If frP2Date = 2 Then
        tmpstr = tmpstr & " AND datevalue(stamptime) between " & EnDate(DateValue(dpP2DateFrom)) & " and " & EnDate(DateValue(dpP2DateTo))
    End If
    
    If frP2Time = 2 Then
        tmpstr = tmpstr & " AND timevalue(stamptime) between " & EnDate(TimeValue(dpP2TimeFrom)) & " and " & EnDate(TimeValue(dpP2TimeTo))
    End If
    
    On Error Resume Next
    
    Set pvrset = cnTuition.Execute(tmpstr)
    
    lvP2Invoice.ListItems.Clear
    
    With pvrset
        If .EOF And .BOF Then
            Set pvrset = Nothing
            cmdP2Print.Enabled = False
            Exit Sub
        End If
        
        .MoveFirst
        
        nIndex = 1
        
        sGrandTotal = 0
        While Not .EOF
            lvP2Invoice.ListItems.Add nIndex, , Format$(.Fields("invoicenumber"), "00000")
            lvP2Invoice.ListItems(nIndex).SubItems(1) = .Fields("engname")
            lvP2Invoice.ListItems(nIndex).SubItems(2) = .Fields("tftext")
            lvP2Invoice.ListItems(nIndex).SubItems(3) = Format$(.Fields("amountdue"), "#,0.00")
            lvP2Invoice.ListItems(nIndex).SubItems(4) = Format$(.Fields("stamptime"), "MM/DD/YY HH:MM ampm")
            lvP2Invoice.ListItems(nIndex).SubItems(5) = .Fields("newstud")
            lvP2Invoice.ListItems(nIndex).SubItems(6) = .Fields("studrefnumber")
            sGrandTotal = sGrandTotal + .Fields("amountdue")
            nIndex = nIndex + 1
            .MoveNext
        Wend
        
        lvP2Invoice.ListItems.Add nIndex, , ""
        lvP2Invoice.ListItems(nIndex).SubItems(2) = "TOTAL:"
        lvP2Invoice.ListItems(nIndex).SubItems(3) = Format$(sGrandTotal, "#,0.00")
        
    End With
    cmdP2Print.Enabled = True

    Set pvrset = Nothing
End Sub


Private Sub cbP2InvoiceNumber_Click()
    If cbP2InvoiceNumber.ListCount = 0 Then
        Exit Sub
    End If
    
    On Error Resume Next

    nPage2InvoiceNumber = GetInvoiceNumber(cbP2InvoiceNumber.Text)
    UpdateInvoiceList
    cmdP2Delete.Enabled = False
End Sub


Sub InitInvoiceReportListView()
    With Me.lvP2Invoice
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Invoice No.", 800
        .ColumnHeaders.Add 2, , "Student", 3000
        .ColumnHeaders.Add 3, , "Fee", 2000
        .ColumnHeaders.Add 4, , "Amount", 1100, lvwColumnRight
        .ColumnHeaders.Add 5, , "Date", 1100
        .ColumnHeaders.Add 6, , "Status", 500
        .ColumnHeaders.Add 7, , "", 0
        .ListItems.Clear
    End With
End Sub



Private Sub lvP2Invoice_Click()
    If lvP2Invoice.ListItems.Count > 0 Then
        cmdP2Delete.Enabled = True
        cmdAdmission.Enabled = True
    End If
End Sub

Private Sub mnuAbout_Click()
    frmAbout.Show vbModal, Me
End Sub

Private Sub mnuChangeUser_Click()
    strOperator = InputBox("SINO KA?", "ALAM NYO NA 'TO", "III")
    Me.Caption = "ENROLLMENT - " & Format$(Now, "dddd mmmm d, yyyy") & " USER: " & strOperator
End Sub


Private Sub mnuClose_Click()
    sbStatbar.Panels(1).Text = ""
    DisableControls
End Sub

Private Sub mnuCopyTemplate_Click()
    frmCreateNewDatabase.Show vbModal, Me
End Sub

Private Sub mnuExit_Click()
    Unload Me
End Sub

Private Sub mnuFind_Click()
    Form2.Show vbModal, Me
End Sub

Private Sub mnuFirelist_Click()
    cbP1EnglishLevel.SetFocus
    SendMsg cbP1Student.hWnd, &H14F, 1, 1
    cbP1Student.SetFocus
End Sub

Private Sub mnuImportTuitionFees_Click()
    cdDialog.FileName = ""
    cdDialog.InitDir = fs.BuildPath(App.Path, "TUITION")
    cdDialog.Filter = "Tuition fees (Excel format)|tuitionfees*.xls"
    cdDialog.ShowOpen
    
    If fs.FileExists(cdDialog.FileName) Then
        If MsgBox("Process " & cdDialog.FileName & "?", vbInformation + vbYesNo, "IMPORT TUITION FEES") = vbYes Then
            ImportTuitionFees cdDialog.FileName
        End If
    End If

End Sub

Private Sub mnuInvoiceList_Click()
    On Error Resume Next
    
    cbP1EnglishLevel.SetFocus
    SendMsg cbP2InvoiceNumber.hWnd, &H14F, 1, 1
    cbP2InvoiceNumber.SetFocus
End Sub

Private Sub mnuOpenBackend_Click()
    If Not InitConnections Then
        Me.sbStatbar.Panels(1).Text = ""
        End
    End If
    
    Me.sbStatbar.Panels(1).Text = "READY Backend: " & SrcDBPath
    
    ReadGlobalString
    EnableControls
    
    mnuChangeUser_Click

End Sub

Private Sub mnuPrintersetup_Click()
    cdDialog.ShowPrinter
    MsgBox cdDialog.PrinterDefault
End Sub

Private Sub mnuReceipts_Click()
    frmReceipts.Show vbModal, Me
End Sub

Private Sub mnuReport_Click()
    frmReport.Show vbModal, Me
End Sub


Private Sub mnuResetDatabase_Click()
    If MsgBox("This will reset the current tuition backend. Invoice counters will be reset and invoice and receipt entries will be cleared. Proceed?", vbExclamation + vbYesNo, "RESET DATABASE") = vbYes Then
        cnTuition.Execute "DELETE * FROM tblInvoice"
        cnTuition.Execute "DELETE * FROM tblPayments"
        cnTuition.Execute "UPDATE tblINPicks SET picked = false"
    End If
End Sub

Private Sub mnuTuitionFees_Click()
    frmBooks.Show vbModal, Me
End Sub

Private Sub mnuUpdateMasterlist_Click()
    cdDialog.FileName = ""
    cdDialog.InitDir = App.Path
    cdDialog.Filter = "Masterlist backend|*masterlist.mdb"
    cdDialog.ShowOpen
    
    If fs.FileExists(cdDialog.FileName) Then
        If MsgBox("Proceed importing master list?", vbInformation + vbYesNo, "IMPORT MASTER LIST") = vbYes Then
            UpdateMasterList cdDialog.FileName
        End If
    End If
End Sub

Private Sub Option1_Click()
    frP1InvoiceNumber = 1
    frP1InvoiceNumberSelect
End Sub

Private Sub Option2_Click()
    frP1InvoiceNumber = 2
    frP1InvoiceNumberSelect
End Sub

Private Sub Option3_Click()
    frP1InvoiceNumber = 3
    frP1InvoiceNumberSelect
End Sub

Private Sub Option4_Click()
    frP2Date = 1
    Me.dpP2DateFrom.Enabled = False
    Me.dpP2DateTo.Enabled = False
End Sub

Private Sub Option5_Click()
    frP2Date = 2
    Me.dpP2DateFrom.Enabled = True
    Me.dpP2DateTo.Enabled = True
    
End Sub

Private Sub Option6_Click()
    frP2Time = 1
    Me.dpP2TimeFrom.Enabled = False
    Me.dpP2TimeTo.Enabled = False

End Sub

Private Sub Option7_Click()
    frP2Time = 2
    Me.dpP2TimeFrom.Enabled = True
    Me.dpP2TimeTo.Enabled = True
    
End Sub
