Attribute VB_Name = "PrintHelper"
Option Explicit

Public Function emp(ByVal v As Long) As String
    If v = 1 Then
        emp = Chr(27) & Chr(69)
    Else
        emp = Chr(27) & Chr(70)
    End If
End Function

Public Function dblwidth(ByVal v As Long) As String
    If v = 1 Then
        dblwidth = Chr(14)
    Else
        dblwidth = Chr(20)
    End If
End Function

Public Function uline(ByVal v As Long) As String
    If v = 1 Then
        uline = Chr(27) & Chr(45) & Chr(1)
    Else
        uline = Chr(27) & Chr(45) & Chr(0)
    End If
End Function

Public Function comp(ByVal v As Long) As String
    If v = 1 Then
        comp = Chr(15)
    Else
        comp = Chr(18)
    End If
End Function


Public Function deflinespace() As String
    deflinespace = Chr(27) & Chr(50)
End Function

Public Sub CancelPrint()

    Open strPrinter For Output As #1

    Print #1, Chr(24)
    
    Close #1
End Sub



Public Function formatcolumn(ByRef s1 As String, ByRef s2 As String, ByVal nwidth As Long) As String
    Dim l1 As Long, l2 As Long
    
        
    If Len(s1) > nwidth Then
        s1 = Mid(s1, 1, nwidth - 2)
    End If
        
    formatcolumn = s1 & String(nwidth - Len(s1) - Len(s2), " ") & s2
    
End Function

Public Function joincolumns(ByRef s1 As String, ByRef s2 As String, ByVal nwidth As Long) As String
    Dim l1 As Long, l2 As Long
    
    If Len(s1) > nwidth Then
        s1 = Mid(s1, 1, nwidth - 2)
    End If
    
    joincolumns = s1 & String(nwidth - Len(s1), " ") & s2
    
End Function



Public Function centertext(ByRef s As String) As String
    centertext = String(Round(((80 - Len(s)) / 2) + 0.5, 0), " ") & s
End Function

Public Function centertextwide(ByRef s As String) As String
    centertextwide = String(Round(((40 - Len(s)) / 2) + 0.5, 0), " ") & s
End Function


Public Sub PrintReceipt(ByVal nIVNumber As Long)
    Dim nctr As Long
    Dim trow As Long
    Dim s1 As Long, s2 As Long
    Dim rcount As Long
    
    Dim pvrset As ADODB.Recordset
    Dim sqlstring As String
    
    Dim sRefNumber As String
    Dim sTransaction As String
    Dim sStudName As String
    Dim sID As String
    Dim sEngClass As String
    Dim sChiClass As String
    Dim sDescription As String
    Dim sPrice As String
    Dim lColString As String
    Dim sTotal As String
    Dim sGrandTotal As Single
    
    On Error Resume Next
    
    Open strPrinter For Output As #1
    
    If Err <> 0 Then
        MsgBox "Printer not found.", vbOKOnly + vbExclamation, "PRINT INVOICE"
        Exit Sub
    End If
    
    Print #1, deflinespace & emp(1) & dblwidth(1) & centertextwide("ZAMBOANGA CHONG HUA HIGH SCHOOL") & dblwidth(0)
    rcount = rcount + 1
    Print #1, centertext(strSchoolYear & " - " & strSemester)
    rcount = rcount + 1
    Print #1, centertext("ENROLLMENT STATEMENT")
    rcount = rcount + 1
    Print #1, formatcolumn("No.: " & Format$(nIVNumber, "00000"), Format$(Now, "MM/DD/YYYY hh:mm ampm") & " by " & strOperator, 80)
    rcount = rcount + 1
    Print #1, String(80, "=")
    rcount = rcount + 1
    Print #1, joincolumns("DETAILS", formatcolumn("DESCRIPTION", "AMOUNT", 35), 45)
    rcount = rcount + 1
    Print #1, String(80, "=")
    rcount = rcount + 1
    
    '*****************
    
    sqlstring = "SELECT iv.invoicenumber AS invoicenum, ml.idnum, ml.engname, ml.engsection, ml.chilevel, tid.tftext as tdesc, tf.tftext as ttext, tf.tfprice, iv.amountdue, iv.stamptime,ml.isnew as xnew, tuser,iv.tfrefnumber,ml.studrefnumber" & _
                " FROM ((tblInvoice AS iv INNER JOIN tblTuitionID AS tID ON iv.tfrefnumber=tid.tfid) INNER JOIN tblTuitionFee AS tf ON tf.tfrefnumber=iv.tfrefnumber) INNER JOIN tblMasterList AS ml ON ml.studrefnumber=iv.studrefnumber" & _
                " WHERE iv.invoicenumber = " & nIVNumber & _
                " ORDER BY ml.englevel,ml.studrefnumber,iv.invoicenumber, listorder"
                
    Set pvrset = cnTuition.Execute(sqlstring)
    
    
    If pvrset.BOF And pvrset.EOF Then
        Close #1
        Set pvrset = Nothing
        Exit Sub
    End If

    pvrset.MoveFirst
    
    s1 = pvrset.Fields("studrefnumber")
    trow = 0
    
    sGrandTotal = 0
    
    While Not pvrset.EOF
        
        s2 = pvrset.Fields("studrefnumber")
        
        If s1 <> s2 Then
            sGrandTotal = sGrandTotal + CSng(sTotal)
            Print #1, joincolumns(String(45, " "), formatcolumn("** TOTAL", "** " & Format$(sTotal, "#,0.00"), 35), 45)
            rcount = rcount + 1
            Print #1, String(80, "-")
            rcount = rcount + 1
            trow = 1
            s1 = s2
        Else
            trow = trow + 1
        End If
        
        sRefNumber = Format$(pvrset.Fields("tfrefnumber"), "0000")
        sTransaction = pvrset.Fields("tdesc")
        sStudName = pvrset.Fields("engname")
        sID = pvrset.Fields("idnum")
        sEngClass = pvrset.Fields("engsection")
        sChiClass = ""
        sTotal = pvrset.Fields("amountdue")
        sDescription = pvrset.Fields("ttext")
        sPrice = Format$(pvrset.Fields("tfprice"), "#,0.00")
        
        Select Case trow
            Case 1
                lColString = sRefNumber & "-" & sTransaction
                
            Case 2
                lColString = sID
                
            Case 3
                lColString = sStudName

            Case 4
                lColString = sEngClass
                
            Case Else
                lColString = String(45, " ")
        End Select
        
        Print #1, joincolumns(lColString, formatcolumn(sDescription, sPrice, 35), 45)
        rcount = rcount + 1
        
        pvrset.MoveNext
    Wend
    
    sGrandTotal = sGrandTotal + CSng(sTotal)
    
    Print #1, joincolumns(String(45, " "), formatcolumn("** TOTAL", "** " & Format$(sTotal, "#,0.00"), 35), 45)
    rcount = rcount + 1
    Print #1, String(80, "-")
    rcount = rcount + 1
    Print #1, joincolumns(String(45, " "), formatcolumn("** GRAND TOTAL", Format$(sGrandTotal, "#,0.00"), 35), 45)
    rcount = rcount + 1
    
'***

    Print #1, Chr(13)
    rcount = rcount + 1
    
    Print #1, deflinespace & "******************************************************************************"
    Print #1, Chr(13)
    Print #1, "Parents' Signature: _________________    TOTAL               _______________"
    Print #1, Chr(13)
    Print #1, "                                         Less: Scholarship   _______________"
    Print #1, Chr(13)
    Print #1, "       Assessed by: _________________    Add: Old Account    _______________"
    Print #1, Chr(13)
    Print #1, "                                         Balance for Payment _______________"
    Print #1, "Cash Received by: _________________"
    Print #1, "                                         Cash Amount         _______________"
    Print #1, Chr(13)
    Print #1, "       Approved by: _________________    Check Amount        _______________"
    Print #1, "                                         Bank/Check No."
    Print #1, "                                         Date of Check"
    Print #1, "                                         TOTAL               _______________"

    rcount = rcount + 16

    While (rcount <= 76)
        Print #1, Chr(13)
        rcount = rcount + 1
    Wend
    
    Close #1
    
    Set pvrset = Nothing
End Sub


