VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReceipts 
   Caption         =   "RECEIPTS"
   ClientHeight    =   7605
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10845
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   707.442
   ScaleMode       =   0  'User
   ScaleWidth      =   951.316
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "DETAILS"
      Height          =   6630
      Left            =   4455
      TabIndex        =   16
      Top             =   945
      Width           =   6315
      Begin VB.OptionButton opOthers 
         Caption         =   "Others"
         Height          =   285
         Left            =   4500
         TabIndex        =   29
         Top             =   5895
         Width           =   1275
      End
      Begin VB.OptionButton opReserved 
         Caption         =   "Reserved"
         Height          =   240
         Left            =   3015
         TabIndex        =   28
         Top             =   6210
         Width           =   1140
      End
      Begin VB.OptionButton opFullpayment 
         Caption         =   "Full payment"
         Height          =   195
         Left            =   3015
         TabIndex        =   27
         Top             =   5940
         Width           =   1545
      End
      Begin VB.CommandButton cmdP3Delete 
         Caption         =   "DELETE"
         Height          =   510
         Left            =   1440
         TabIndex        =   26
         Top             =   6030
         Width           =   1185
      End
      Begin VB.CommandButton cmdP3Add 
         Caption         =   "ADD"
         Height          =   510
         Left            =   135
         TabIndex        =   25
         Top             =   6030
         Width           =   1185
      End
      Begin VB.TextBox txtP3Remarks 
         Height          =   330
         Left            =   3015
         TabIndex        =   24
         Text            =   "Text2"
         Top             =   5535
         Width           =   3210
      End
      Begin VB.TextBox txtP3Amount 
         Height          =   330
         Left            =   765
         TabIndex        =   22
         Text            =   "Text1"
         Top             =   5535
         Width           =   1365
      End
      Begin MSComctlLib.ListView lvP3Payments 
         Height          =   2265
         Left            =   90
         TabIndex        =   20
         Top             =   3240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   3995
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin MSComctlLib.ListView lvP3Students 
         Height          =   2310
         Left            =   90
         TabIndex        =   18
         Top             =   540
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4075
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.Label Label7 
         Caption         =   "Remarks:"
         Height          =   285
         Left            =   2250
         TabIndex        =   23
         Top             =   5580
         Width           =   780
      End
      Begin VB.Label Label6 
         Caption         =   "Amount:"
         Height          =   240
         Left            =   90
         TabIndex        =   21
         Top             =   5580
         Width           =   645
      End
      Begin VB.Label txtCap 
         Caption         =   "Payment Details"
         Height          =   285
         Left            =   135
         TabIndex        =   19
         Top             =   3015
         Width           =   1230
      End
      Begin VB.Label Label4 
         Caption         =   "Name List"
         Height          =   240
         Left            =   135
         TabIndex        =   17
         Top             =   315
         Width           =   825
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "INVOICE LIST"
      Height          =   6630
      Left            =   90
      TabIndex        =   12
      Top             =   945
      Width           =   4290
      Begin MSComctlLib.ListView lvP3Invoice 
         Height          =   5775
         Left            =   90
         TabIndex        =   15
         Top             =   765
         Width           =   4110
         _ExtentX        =   7250
         _ExtentY        =   10186
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
      Begin VB.CommandButton cmdP3Update 
         Caption         =   "UPDATE"
         Height          =   465
         Left            =   1980
         TabIndex        =   14
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label txtReceiptCount 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   180
         TabIndex        =   13
         Top             =   315
         Width           =   1680
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "TIME FILTER"
      Height          =   735
      Left            =   5445
      TabIndex        =   6
      Top             =   90
      Width           =   5325
      Begin VB.OptionButton opAllTime 
         Caption         =   "ALL"
         Height          =   240
         Left            =   135
         TabIndex        =   9
         Top             =   315
         Width           =   690
      End
      Begin VB.OptionButton opTimeRange 
         Caption         =   "Between"
         Height          =   240
         Left            =   855
         TabIndex        =   7
         Top             =   315
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dpP3TimeFrom 
         Height          =   330
         Left            =   1980
         TabIndex        =   8
         Top             =   270
         Width           =   1365
         _ExtentX        =   2408
         _ExtentY        =   582
         _Version        =   393216
         Format          =   25034754
         CurrentDate     =   39730
      End
      Begin MSComCtl2.DTPicker dpP3TimeTo 
         Height          =   330
         Left            =   3870
         TabIndex        =   10
         Top             =   270
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Format          =   25034754
         CurrentDate     =   39730
      End
      Begin VB.Label Label2 
         Caption         =   "and"
         Height          =   240
         Left            =   3465
         TabIndex        =   11
         Top             =   315
         Width           =   330
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "DATE FILTER"
      Height          =   735
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   5280
      Begin VB.OptionButton opDateRange 
         Caption         =   "Between"
         Height          =   240
         Left            =   855
         TabIndex        =   3
         Top             =   315
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dpP3DateFrom 
         Height          =   330
         Left            =   1980
         TabIndex        =   2
         Top             =   270
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   25034753
         CurrentDate     =   39730
      End
      Begin VB.OptionButton opAllDate 
         Caption         =   "ALL"
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   315
         Width           =   690
      End
      Begin MSComCtl2.DTPicker dpP3DateTo 
         Height          =   330
         Left            =   3825
         TabIndex        =   5
         Top             =   270
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Format          =   25034753
         CurrentDate     =   39730
      End
      Begin VB.Label Label1 
         Caption         =   "and"
         Height          =   240
         Left            =   3420
         TabIndex        =   4
         Top             =   315
         Width           =   330
      End
   End
End
Attribute VB_Name = "frmReceipts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'P A G E  3  -  I N V O I C E  R E P O R T
'**************************************************************

Dim frP3Date As Long
Dim frP3Time As Long
Dim frP3Details As Long

Private Sub Form_Load()
    
    frP3Date = 1
    frP3Time = 1
    frP3Details = 1
    
    dpP3DateFrom.Value = Now
    dpP3DateTo.Value = Now
    dpP3TimeFrom.Value = #7:00:00 AM#
    dpP3TimeTo.Value = #12:00:00 PM#
    
    
    Me.opAllDate = True
    Me.opAllTime = True
    Me.opFullpayment = True

    InitInvoiceReceipts
End Sub



Private Sub cmdP3Update_Click()
    Dim tmpstr As String
    Dim sGrandTotal As Single
    Dim pvrset As adodb.Recordset
    
    tmpstr = "SELECT  iv.invoicenumber,sum(iv.amountdue) as xtotal,count(iv.amountdue) as xcount  from tblInvoice as iv inner join tblMasterList as ml on iv.studrefnumber=ml.studrefnumber "
    
    If frP3Date = 2 Then
        tmpstr = tmpstr & " WHERE datevalue(stamptime) between " & EnDate(DateValue(dpP3DateFrom)) & " and " & EnDate(DateValue(dpP3DateTo))
    End If
    
    If frP3Time = 2 Then
        
        If InStr(1, tmpstr, "WHERE", vbTextCompare) = 0 Then
            tmpstr = tmpstr & " WHERE timevalue(stamptime) between " & EnDate(TimeValue(dpP3TimeFrom)) & " and " & EnDate(TimeValue(dpP3TimeTo))
        Else
            tmpstr = tmpstr & " AND timevalue(stamptime) between " & EnDate(TimeValue(dpP3TimeFrom)) & " and " & EnDate(TimeValue(dpP3TimeTo))
        End If
    End If
    
    tmpstr = tmpstr & " GROUP BY iv.invoicenumber"
    
    Set pvrset = cnTuition.Execute(tmpstr)
    
    lvP3Invoice.ListItems.Clear
    
    With pvrset
        If .EOF And .BOF Then
            Set pvrset = Nothing
            Exit Sub
        End If
        
        .MoveFirst
        
        nIndex = 1
        
        sGrandTotal = 0
        While Not .EOF
            lvP3Invoice.ListItems.Add nIndex, , Format$(.Fields("invoicenumber"), "00000")
            lvP3Invoice.ListItems(nIndex).SubItems(1) = Format$(.Fields("xtotal"), "#,0.00")
            lvP3Invoice.ListItems(nIndex).SubItems(2) = .Fields("xcount")
            nIndex = nIndex + 1
            .MoveNext
        Wend
        
    End With
    
    Me.txtReceiptCount.Caption = "Count: " & lvP3Invoice.ListItems.Count
    
    DisablePaymentControls
    lvP3Students.ListItems.Clear
    
    
    Set pvrset = Nothing
End Sub


Sub InitInvoiceReceipts()
    With Me.lvP3Invoice
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Invoice No.", 1200
        .ColumnHeaders.Add 2, , "Amount", 1200, lvwColumnRight
        .ColumnHeaders.Add 3, , "Count", 500
        .ListItems.Clear
    End With
    
    With Me.lvP3Students
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Name", 2800
        .ColumnHeaders.Add 2, , "Fee", 1800
        .ColumnHeaders.Add 3, , "Total", 1200, lvwColumnRight
        .ColumnHeaders.Add 4, , "Date/Time", 2000
        .ColumnHeaders.Add 5, , "", 0
        .ColumnHeaders.Add 6, , "", 0
        .ListItems.Clear
    End With
    
    With Me.lvP3Payments
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Payment for", 1800
        .ColumnHeaders.Add 2, , "Amount", 1200, lvwColumnRight
        .ColumnHeaders.Add 3, , "Date", 1800
        .ColumnHeaders.Add 4, , "Remarks", 2000
        .ColumnHeaders.Add 5, , "", 0
        .ListItems.Clear
    End With
    
End Sub


Sub DisablePaymentControls()
        lvP3Payments.ListItems.Clear
        txtP3Amount.Enabled = False
        txtP3Remarks.Enabled = False
        cmdP3Delete.Enabled = False
        cmdP3Add.Enabled = False
End Sub


Private Sub lvP3Invoice_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvP3Invoice.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    txtCap.Caption = "Payments"
    
    np4InvoiceNumber = CLng(Item.Text)
    
    FillUpStudents
    DisablePaymentControls
    
End Sub



Sub FillUpStudents()
    Dim pvrset As adodb.Recordset
    Dim tmpstr As String
    
    tmpstr = "SELECT ml.studrefnumber,ml.engname,inv.tfrefnumber,inv.amountdue,tid.tftext,inv.stamptime from" & _
             "(tblInvoice as inv inner join tblMasterList as ml on inv.studrefnumber=ml.studrefnumber) inner join tblTuitionID as tID on tID.tfid=inv.tfrefnumber" & _
             " WHERE inv.InvoiceNumber = " & np4InvoiceNumber
             
    Set pvrset = cnTuition.Execute(tmpstr)
    
    lvP3Students.ListItems.Clear
    
    If pvrset.BOF And pvrset.EOF Then
        Exit Sub
    End If
    
    With pvrset
        
        .MoveFirst
        
        nIndex = 1
        While Not .EOF
            lvP3Students.ListItems.Add nIndex, , .Fields("engname")
            lvP3Students.ListItems(nIndex).SubItems(1) = .Fields("tftext")
            lvP3Students.ListItems(nIndex).SubItems(2) = Format$(.Fields("amountdue"), "#,0.00")
            lvP3Students.ListItems(nIndex).SubItems(3) = Format$(.Fields("stamptime"), "MM-DD-YY hh:mm ampm")
            lvP3Students.ListItems(nIndex).SubItems(4) = .Fields("studrefnumber")
            lvP3Students.ListItems(nIndex).SubItems(5) = .Fields("tfrefnumber")
            nIndex = nIndex + 1
            .MoveNext
        Wend
        
    End With
    
End Sub


Sub FillPaymentList()
    Dim pvrset As adodb.Recordset
    Dim tmpstr As String
    
    tmpstr = "SELECT t.tftext, p.payamount, p.timestamp, p.payremarks, p.payindex" & _
             " FROM tblPayments AS p INNER JOIN tblTuitionID AS t ON p.tfid=t.tfid" & _
             " WHERE STUDREFNUMBER = " & np4StudRefNumber

    Set pvrset = cnTuition.Execute(tmpstr)
    
    lvP3Payments.ListItems.Clear
    
    If pvrset.BOF And pvrset.EOF Then
        Exit Sub
    End If
    
    With pvrset
        
        .MoveFirst
        
        nIndex = 1
        While Not .EOF
            lvP3Payments.ListItems.Add nIndex, , .Fields("tftext")
            lvP3Payments.ListItems(nIndex).SubItems(1) = Format$(.Fields("payamount"), "#,0.00")
            lvP3Payments.ListItems(nIndex).SubItems(2) = Format$(.Fields("timestamp"), "MM-DD-YY hh:mm ampm")
            lvP3Payments.ListItems(nIndex).SubItems(3) = .Fields("payremarks")
            lvP3Payments.ListItems(nIndex).SubItems(4) = .Fields("payindex")
            nIndex = nIndex + 1
            .MoveNext
        Wend
        
    End With
    
End Sub



Private Sub lvP3Payments_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmdP3Delete.Enabled = True
End Sub

Private Sub lvP3Students_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If lvP3Students.ListItems.Count = 0 Then
        Exit Sub
    End If
    
    np4StudRefNumber = CLng(Item.SubItems(4))
    
    txtCap.Caption = Item.Text & " Payments:"
    
    
    frP3Details_Click
    
    FillPaymentList
    
    txtP3Amount.Enabled = True
    txtP3Remarks.Enabled = True
    cmdP3Add.Enabled = True
    frP3Details = 1

End Sub

Private Sub opAllDate_Click()
    frP3Date = 1
    dpP3DateFrom.Enabled = False
    dpP3DateTo.Enabled = False
End Sub

Private Sub opAllTime_Click()
    frP3Time = 1
    dpP3TimeFrom.Enabled = False
    dpP3TimeTo.Enabled = False
End Sub

Private Sub opDateRange_Click()
    frP3Date = 2
    dpP3DateFrom.Enabled = True
    dpP3DateTo.Enabled = True
    
End Sub

Private Sub opFullpayment_Click()
    frP3Details = 1
    frP3Details_Click
End Sub

Private Sub opOthers_Click()
    frP3Details = 3
    frP3Details_Click
End Sub

Private Sub opReserved_Click()
    frP3Details = 2
    frP3Details_Click
End Sub

Private Sub opTimeRange_Click()
    frP3Time = 2
    dpP3TimeFrom.Enabled = True
    dpP3TimeTo.Enabled = True
    
End Sub

Private Sub txtP3Amount_Change()
    
    If IsNumeric(txtP3Amount.Text) Then
        cmdP3Add.Enabled = True
    Else
        cmdP3Add.Enabled = False
    End If
End Sub



Private Sub frP3Details_Click()

    If lvP3Students.ListItems.Count = 0 Then
        Exit Sub
    End If

    Select Case frP3Details
       Case 1
          txtP3Amount = Format$(lvP3Students.SelectedItem.SubItems(2), "#,0.00")
          txtP3Remarks = "Fully paid"
    
       Case 2
          txtP3Amount = Format$(lvP3Students.SelectedItem.SubItems(2) - 1000, "#,0.00")
          txtP3Remarks = "Fully paid (w/ P1,000 reserve fee)"
       
       Case 3
          txtP3Amount = ""
          txtP3Remarks = ""
       
    End Select
End Sub


Private Sub cmdP3Add_Click()
    Dim tmpstr
    
    txtP3Remarks.SetFocus
    
    tmpstr = "INSERT into tblPayments (INVOICENUMBER,STUDREFNUMBER,PAYAMOUNT,PAYREMARKS,[TIMESTAMP],TFID) VALUES " & _
             "(" & _
             np4InvoiceNumber & "," & _
             np4StudRefNumber & "," & _
             CSng(txtP3Amount) & "," & _
             IIf(IsNull(txtP3Remarks), EnQuote("<none>"), EnQuote(txtP3Remarks.Text)) & "," & _
             EnDate(Now) & "," & _
             CLng(lvP3Students.SelectedItem.SubItems(5)) & _
             ")"
             
    cnTuition.Execute tmpstr
    
    MsgBox "Payment added", vbOKOnly
    
    FillPaymentList
    
End Sub


Private Sub cmdP3Delete_Click()
    If MsgBox("Delete this entry?", vbOKCancel + vbExclamation, "CANCEL PAYMENT") = vbOK Then
    
        cnTuition.Execute "DELETE from tblPayments WHERE STUDREFNUMBER = " & np4StudRefNumber & " AND TFID = " & _
        CLng(lvP3Students.SelectedItem.SubItems(5)) & " AND PAYINDEX = " & CLng(lvP3Payments.SelectedItem.SubItems(4))
        
        FillPaymentList

    End If
End Sub



