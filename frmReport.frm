VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmReport 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "REPORT"
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10200
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   10200
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Caption         =   "TIME FILTER"
      Height          =   780
      Left            =   2655
      TabIndex        =   16
      Top             =   945
      Width           =   5235
      Begin MSComCtl2.DTPicker dpP4TimeTo 
         Height          =   330
         Left            =   3600
         TabIndex        =   21
         Top             =   315
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16646146
         CurrentDate     =   39731
      End
      Begin MSComCtl2.DTPicker dpP4TimeFrom 
         Height          =   330
         Left            =   1845
         TabIndex        =   19
         Top             =   315
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16646146
         CurrentDate     =   39731
      End
      Begin VB.OptionButton opP4TimeRange 
         Caption         =   "Between"
         Height          =   330
         Left            =   765
         TabIndex        =   18
         Top             =   315
         Width           =   1140
      End
      Begin VB.OptionButton opP4TimeAll 
         Caption         =   "ALL"
         Height          =   240
         Left            =   90
         TabIndex        =   17
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label2 
         Caption         =   "and"
         Height          =   240
         Left            =   3195
         TabIndex        =   20
         Top             =   360
         Width           =   465
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "DATE FILTER"
      Height          =   780
      Left            =   2655
      TabIndex        =   10
      Top             =   90
      Width           =   5235
      Begin MSComCtl2.DTPicker dpP4DateTo 
         Height          =   330
         Left            =   3600
         TabIndex        =   15
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   39731
      End
      Begin MSComCtl2.DTPicker dpP4DateFrom 
         Height          =   330
         Left            =   1845
         TabIndex        =   13
         Top             =   360
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   582
         _Version        =   393216
         Format          =   16646145
         CurrentDate     =   39731
      End
      Begin VB.OptionButton opP4DateRange 
         Caption         =   "Between"
         Height          =   330
         Left            =   765
         TabIndex        =   12
         Top             =   360
         Width           =   960
      End
      Begin VB.OptionButton opP4DateAll 
         Caption         =   "ALL"
         Height          =   195
         Left            =   90
         TabIndex        =   11
         Top             =   405
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "and"
         Height          =   240
         Left            =   3195
         TabIndex        =   14
         Top             =   405
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "PREVIEW"
      Height          =   5595
      Left            =   45
      TabIndex        =   4
      Top             =   1800
      Width           =   10095
      Begin MSComctlLib.ListView lvP4Report 
         Height          =   5235
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   9234
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.CommandButton cmdP4Export 
      Caption         =   "EXPORT"
      Height          =   555
      Left            =   8190
      TabIndex        =   3
      Top             =   1035
      Width           =   1680
   End
   Begin VB.Frame Frame4 
      Height          =   1680
      Left            =   7965
      TabIndex        =   1
      Top             =   90
      Width           =   2130
      Begin VB.CommandButton cmdP4Update 
         Caption         =   "UPDATE"
         Height          =   555
         Left            =   225
         TabIndex        =   2
         Top             =   270
         Width           =   1680
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "OPTIONS"
      Height          =   1635
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   2535
      Begin VB.OptionButton opSummary 
         Caption         =   "Summary"
         Height          =   240
         Left            =   315
         TabIndex        =   9
         Top             =   1260
         Width           =   1815
      End
      Begin VB.OptionButton opTF 
         Caption         =   "Tuition Fees"
         Height          =   285
         Left            =   315
         TabIndex        =   8
         Top             =   945
         Width           =   1635
      End
      Begin VB.OptionButton opSC 
         Caption         =   "Student Council"
         Height          =   330
         Left            =   315
         TabIndex        =   7
         Top             =   630
         Width           =   1860
      End
      Begin VB.OptionButton opPA 
         Caption         =   "Parents' Association"
         Height          =   330
         Left            =   315
         TabIndex        =   6
         Top             =   315
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim frP4Report As Long
Dim frP4Date As Long
Dim frP4Time As Long


Private Sub Form_Load()
    opPA = True
    opP4DateAll = True
    opP4TimeAll = True
    
    frP4Report = 1
    frP4Date = 1
    frP4Time = 1
    
    dpP4DateFrom.Value = Now
    dpP4DateTo.Value = Now
    Me.dpP4TimeFrom.Value = #7:00:00 AM#
    Me.dpP4TimeTo.Value = #12:00:00 PM#
    
    InitReportList
End Sub


'******************************************************************
'PAGE 4
'******************************************************************

Sub InitReportList()
   With Me.lvP4Report
        .View = lvwReport
        .GridLines = True
        .FullRowSelect = True
        .HideSelection = False
        .LabelEdit = lvwManual
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "Invoice No.", 800
        .ColumnHeaders.Add 2, , "Ref. No.", 800
        .ColumnHeaders.Add 3, , "Name", 3000
        .ColumnHeaders.Add 4, , "Level", 1800
        .ColumnHeaders.Add 5, , "Fees", 1800
        .ColumnHeaders.Add 6, , "Amount", 1000, lvwColumnRight
        .ColumnHeaders.Add 7, , "Paid", 1000
        .ListItems.Clear
        .Sorted = False
    End With
End Sub

Sub FillReportList()
    Dim pvrset As ADODB.Recordset
    Dim tmpstr As String
    Dim strFilter As String
    Dim strFeeString As String
    Dim strFeeAmount As String
    Dim sCountTotal As Double
    
    
    tmpstr = " SELECT invoicenumber, idnum, engname, tlv.level as xleveltext, tid.tftext as xtext, payamount, tf.tftext as xfee, tfprice, timestamp, payremarks, ml.studrefnumber as xrefnum " & _
            " FROM (((tblpayments AS p INNER JOIN tblTuitionID AS tID ON p.tfid=tid.tfid) " & _
            " INNER JOIN tblTuitionFee AS tf ON tf.tfrefnumber=p.tfid) " & _
            " INNER JOIN tblMasterList AS ml ON ml.studrefnumber=p.studrefnumber) " & _
            " inner join tblLevelList as tlv on tlv.levelid = ml.englevel"

    Select Case frP4Report
        Case 1
            tmpstr = tmpstr & " WHERE (tf.tftext like 'PARENTS'' ASSOC.')"
            strFeeString = "xfee"
            strFeeAmount = "tfprice"
        
        Case 2
            tmpstr = tmpstr & " WHERE (tf.tftext like 'STUDENT COUNCIL')"
            strFeeString = "xfee"
            strFeeAmount = "tfprice"
        
        Case 3
            tmpstr = tmpstr & " WHERE (tf.tftext like 'TUITION FEE')"
            strFeeString = "xtext"
            strFeeAmount = "payamount"
        
        Case 4
            tmpstr = tmpstr & " WHERE (tf.tftext like 'TUITION FEE')"
            strFeeString = "xtext"
            strFeeAmount = "payamount"
        
        
    End Select
    
    If frP4Date = 2 Then
        tmpstr = tmpstr & " AND datevalue(timestamp) between " & EnDate(DateValue(dpP4DateFrom)) & " and " & EnDate(DateValue(dpP4DateTo))
    End If
    
    If frP4Time = 2 Then
        tmpstr = tmpstr & " AND timevalue(timestamp) between " & EnDate(TimeValue(dpP4TimeFrom)) & " and " & EnDate(TimeValue(dpP4TimeTo))
    End If
    
    tmpstr = tmpstr & " ORDER BY invoicenumber"
    
    Set pvrset = cnTuition.Execute(tmpstr)
    
    lvP4Report.ListItems.Clear
    
    If pvrset.BOF And pvrset.EOF Then
        Exit Sub
    End If
    
    sCountTotal = 0#
    
    With pvrset
        
        .MoveFirst
        
        nIndex = 1
        
        While Not .EOF
            lvP4Report.ListItems.Add nIndex, , Format$(.Fields("invoicenumber"), "00000")
            lvP4Report.ListItems(nIndex).SubItems(1) = Format$(.Fields("xrefnum"), "00000")
            lvP4Report.ListItems(nIndex).SubItems(2) = .Fields("engname")
            lvP4Report.ListItems(nIndex).SubItems(3) = .Fields("xleveltext")
            lvP4Report.ListItems(nIndex).SubItems(4) = .Fields(strFeeString)
            lvP4Report.ListItems(nIndex).SubItems(5) = Format$(.Fields(strFeeAmount), "#,0.00")
            lvP4Report.ListItems(nIndex).SubItems(6) = Format$(.Fields("timestamp"), "mm/dd/yy hh:mm ampm")
            sCountTotal = sCountTotal + CDbl(lvP4Report.ListItems(nIndex).SubItems(5))
            nIndex = nIndex + 1
            .MoveNext
        Wend
        
    End With
    
    lvP4Report.ListItems.Add nIndex, , ""
    lvP4Report.ListItems(nIndex).SubItems(2) = "COUNT:"
    lvP4Report.ListItems(nIndex).SubItems(3) = lvP4Report.ListItems.Count - 1
    lvP4Report.ListItems(nIndex).SubItems(4) = "TOTAL:"
    lvP4Report.ListItems(nIndex).SubItems(5) = Format$(sCountTotal, "#,0.00")
    

End Sub

Sub InvoiceSummary()
    Dim pvrset As ADODB.Recordset
    Dim tmpstr As String
    Dim strFilter As String
    Dim strFeeString As String
    Dim strFeeAmount As String
    Dim sCountTotal As Double
    
    
    tmpstr = " SELECT invoicenumber, idnum, engname, tlv.level as xleveltext, tid.tftext as xtext, payamount, tf.tftext as xfee, tfprice, timestamp, payremarks, ml.studrefnumber as xrefnum " & _
            " FROM (((tblpayments AS p INNER JOIN tblTuitionID AS tID ON p.tfid=tid.tfid) " & _
            " INNER JOIN tblTuitionFee AS tf ON tf.tfrefnumber=p.tfid) " & _
            " INNER JOIN tblMasterList AS ml ON ml.studrefnumber=p.studrefnumber) " & _
            " inner join tblLevelList as tlv on tlv.levelid = ml.englevel"

    Select Case frP4Report
        Case 1
            tmpstr = tmpstr & " WHERE (tf.tftext like 'PARENTS'' ASSOC.')"
            strFeeString = "xfee"
            strFeeAmount = "tfprice"
        
        Case 2
            tmpstr = tmpstr & " WHERE (tf.tftext like 'STUDENT COUNCIL')"
            strFeeString = "xfee"
            strFeeAmount = "tfprice"
        
        Case 3
            tmpstr = tmpstr & " WHERE (tf.tftext like 'TUITION FEE')"
            strFeeString = "xtext"
            strFeeAmount = "payamount"
        
        Case 4
            tmpstr = tmpstr & " WHERE (tf.tftext like 'TUITION FEE')"
            strFeeString = "xtext"
            strFeeAmount = "payamount"
        
        
    End Select
    
    If frP4Date = 2 Then
        tmpstr = tmpstr & " AND datevalue(timestamp) between " & EnDate(DateValue(dpP4DateFrom)) & " and " & EnDate(DateValue(dpP4DateTo))
    End If
    
    If frP4Time = 2 Then
        tmpstr = tmpstr & " AND timevalue(timestamp) between " & EnDate(TimeValue(dpP4TimeFrom)) & " and " & EnDate(TimeValue(dpP4TimeTo))
    End If
    
    tmpstr = tmpstr & " ORDER BY invoicenumber"
    
    Set pvrset = cnTuition.Execute(tmpstr)
    
    lvP4Report.ListItems.Clear
    
    If pvrset.BOF And pvrset.EOF Then
        Exit Sub
    End If
    
    sCountTotal = 0#
    
    With pvrset
        
        .MoveFirst
        
        nIndex = 1
        
        While Not .EOF
            lvP4Report.ListItems.Add nIndex, , Format$(.Fields("invoicenumber"), "00000")
            lvP4Report.ListItems(nIndex).SubItems(1) = Format$(.Fields("xrefnum"), "00000")
            lvP4Report.ListItems(nIndex).SubItems(2) = .Fields("engname")
            lvP4Report.ListItems(nIndex).SubItems(3) = .Fields("xleveltext")
            lvP4Report.ListItems(nIndex).SubItems(4) = .Fields(strFeeString)
            lvP4Report.ListItems(nIndex).SubItems(5) = Format$(.Fields(strFeeAmount), "#,0.00")
            lvP4Report.ListItems(nIndex).SubItems(6) = Format$(.Fields("timestamp"), "mm/dd/yy hh:mm ampm")
            sCountTotal = sCountTotal + CDbl(lvP4Report.ListItems(nIndex).SubItems(5))
            nIndex = nIndex + 1
            .MoveNext
        Wend
        
    End With
    
    lvP4Report.ListItems.Add nIndex, , ""
    lvP4Report.ListItems(nIndex).SubItems(2) = "COUNT:"
    lvP4Report.ListItems(nIndex).SubItems(3) = lvP4Report.ListItems.Count - 1
    lvP4Report.ListItems(nIndex).SubItems(4) = "TOTAL:"
    lvP4Report.ListItems(nIndex).SubItems(5) = Format$(sCountTotal, "#,0.00")
    
End Sub


Sub ExportToExcel()
    Dim pvrset As ADODB.Recordset
    Dim tmpstr As String
    Dim strFilter As String
    Dim strFeeString As String
    Dim strFeeAmount As String
    Dim sCountTotal As Double
    
    Dim xlObj As Object
    Dim outxl As Object
    Dim oItem As ListItem
    Dim n As Long
    Dim nsheetindex As Long
    
    Set xlObj = CreateObject("Excel.Application")
    
    Set outxl = xlObj.workbooks.Add(App.Path & "\TEMPLATES\Report.xlt")
    
    
    For nsheetindex = 1 To 4
    
        tmpstr = " SELECT invoicenumber, idnum, engname, tlv.level as xleveltext, tid.tftext as xtext, payamount, tf.tftext as xfee, tfprice, timestamp, payremarks, ml.studrefnumber as xrefnum " & _
                " FROM (((tblpayments AS p INNER JOIN tblTuitionID AS tID ON p.tfid=tid.tfid) " & _
                " INNER JOIN tblTuitionFee AS tf ON tf.tfrefnumber=p.tfid) " & _
                " INNER JOIN tblMasterList AS ml ON ml.studrefnumber=p.studrefnumber) " & _
                " inner join tblLevelList as tlv on tlv.levelid = ml.englevel"
    
        Select Case nsheetindex
            Case 1
                tmpstr = tmpstr & " WHERE (tf.tftext like 'TUITION FEE')"
                strFeeString = "xtext"
                strFeeAmount = "payamount"
            
            Case 2
                tmpstr = tmpstr & " WHERE (tf.tftext like 'STUDENT COUNCIL')"
                strFeeString = "xfee"
                strFeeAmount = "tfprice"
            
            Case 3
                tmpstr = tmpstr & " WHERE (tf.tftext like 'PARENTS'' ASSOC.')"
                strFeeString = "xfee"
                strFeeAmount = "tfprice"
            
            Case 4
                tmpstr = tmpstr & " WHERE (tf.tftext like 'TUITION FEE')"
                strFeeString = "xtext"
                strFeeAmount = "payamount"
            
        End Select
        
        If frP4Date = 2 Then
            tmpstr = tmpstr & " AND datevalue(timestamp) between " & EnDate(DateValue(dpP4DateFrom)) & " and " & EnDate(DateValue(dpP4DateTo))
        End If
        
        If frP4Time = 2 Then
            tmpstr = tmpstr & " AND timevalue(timestamp) between " & EnDate(TimeValue(dpP4TimeFrom)) & " and " & EnDate(TimeValue(dpP4TimeTo))
        End If
        
        tmpstr = tmpstr & " ORDER BY invoicenumber"
        
        Set pvrset = cnTuition.Execute(tmpstr)
        
        If pvrset.BOF And pvrset.EOF Then
            Exit Sub
        End If
        
        outxl.sheets(nsheetindex).Cells(6, 1) = "Date from " & Format$(dpP4DateFrom.Value, "mm/dd/yyyy") & " to " & _
                                      Format$(dpP4DateTo.Value, "mm/dd/yyyy") & _
                                      " Time from " & Format$(dpP4TimeFrom.Value, "hh:mm ampm") & " to " & Format$(dpP4TimeTo, "hh:mm ampm")
        
        With pvrset
            
            .MoveFirst
            n = 8
        
            While Not .EOF
                outxl.sheets(nsheetindex).Cells(n, 1) = "'" & n - 7 & "."
                outxl.sheets(nsheetindex).Cells(n, 2) = Format$(.Fields("invoicenumber"), "'00000")
                outxl.sheets(nsheetindex).Cells(n, 3) = .Fields("engname")
                outxl.sheets(nsheetindex).Cells(n, 4) = .Fields("xleveltext")
                outxl.sheets(nsheetindex).Cells(n, 5) = .Fields(strFeeString)
                outxl.sheets(nsheetindex).Cells(n, 6) = Format$(.Fields(strFeeAmount), "#,0.00")
                outxl.sheets(nsheetindex).Cells(n, 7) = Format$(.Fields("timestamp"), "mm/dd/yy hh:mm ampm")
                n = n + 1
                .MoveNext
            Wend
            
        End With
    
    Next
    
    FixSummarySheet outxl.sheets(4)
    
    xlObj.Visible = True
    
End Sub

Sub FixSummarySheet(ByRef xls As Object)
    Dim s1 As String, s2 As String
    Dim nctr As Long
    Dim itemcount As Long
    
    xls.Range("b8:g2500").Sort key1:=xls.Range("d8"), order1:=1, key2:=xls.Range("c8"), order2:=1, header:=2
    
    s1 = "iii"
    nctr = 8
    
    While Len(xls.Cells(nctr, 4)) > 0
        s2 = xls.Cells(nctr, 4)
        
        If StrComp(s1, s2, vbTextCompare) <> 0 Then
            xls.Rows(nctr & ":" & nctr + 1).Insert Shift:=&HFFFFEFE7
            xls.Cells(nctr + 1, 3) = s2
            nctr = nctr + 2
            s1 = s2
            itemcount = 1
        Else
            itemcount = itemcount + 1
        End If
        
        xls.Cells(nctr, 1) = "'" & itemcount & "."
        nctr = nctr + 1
    Wend
    
End Sub


Private Sub cmdP4Update_Click()
    FillReportList
End Sub




Private Sub cmdP4Export_Click()
    ExportToExcel
End Sub


Private Sub lvP4Report_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    lvP4Report.Sorted = True
    
    lvP4Report.SortKey = ColumnHeader.Index - 1
    If lvP4Report.SortOrder = lvwAscending Then
        lvP4Report.SortOrder = lvwDescending
    Else
        lvP4Report.SortOrder = lvwAscending
    End If
    
    lvP4Report.Sorted = False

End Sub

Private Sub opP4DateAll_Click()
    frP4Date = 1
End Sub

Private Sub opP4DateRange_Click()
    frP4Date = 2
End Sub

Private Sub opP4TimeAll_Click()
    frP4Time = 1
End Sub

Private Sub opP4TimeRange_Click()
    frP4Time = 2
End Sub

Private Sub opPA_Click()
    frP4Report = 1
End Sub

Private Sub opSC_Click()
    frP4Report = 2
End Sub

Private Sub opSummary_Click()
    frP4Report = 4
End Sub

Private Sub opTF_Click()
    frP4Report = 3
End Sub
