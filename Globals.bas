Attribute VB_Name = "Globals"
Option Explicit

Public cnTuition As ADODB.Connection
Public cnMasterlist As ADODB.Connection
Public nPage2InvoiceNumber As Long


Public strOperator As String
Public strPrinter As String
Public strSchoolYear As String
Public strSemester As String

Public nIndex As Long
Public np4InvoiceNumber As Long
Public np4StudRefNumber As Long
Public np4TFRefNumber As Long

Public fs As Object


Public Declare Function SendMsg Lib "user32.dll" Alias "SendMessageA" (ByVal h As Long, ByVal m As Long, ByVal wp As Long, ByVal lp As Long) As Long

Function InitConnections() As Boolean

    Set cnTuition = CreateObject("ADODB.Connection")
    Set cnMasterlist = CreateObject("ADODB.Connection")
    
        
    If Not OpenUp Then
        InitConnections = False
        Exit Function
    End If
    
    On Error Resume Next
    
    InitConnections = True
    
    cnTuition.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                 "Data Source=" & SrcDBPath
    
    cnTuition.CursorLocation = adUseClient
    cnTuition.Mode = adModeReadWrite
    
    cnTuition.Open
    
    If Err.Number <> 0 Then
        InitConnections = False
    End If
    
End Function


Function PickInvoiceNumber() As Long
    Dim pvrset As ADODB.Recordset
    Dim numpick As Long
    
    On Error Resume Next
    
    Set pvrset = CreateObject("ADODB.RECORDSET")
    
    pvrset.Open "SELECT INUMBER,PICKED from tblINPicks WHERE not picked ORDER BY inumber", cnTuition, adOpenKeyset, adLockPessimistic
    
    cnTuition.BeginTrans
    
    With pvrset
        .MoveFirst
        .Fields("picked") = True
        .Update
        numpick = .Fields("INUMBER")
    End With
    
    If Err.Number <> 0 Then
        PickInvoiceNumber = 0
        cnTuition.RollbackTrans
    Else
        cnTuition.CommitTrans
        PickInvoiceNumber = numpick
    End If
    
    Set pvrset = Nothing
    
End Function

Sub ReturnPickedNumber(ByVal n As Long)
    'On Error Resume Next
    cnTuition.Execute ("UPDATE tblINPicks set picked = false WHERE INUMBER=" & n)
End Sub


Function EnQuote(ByRef s As String) As String
    EnQuote = """" & s & """"
End Function

Function EnDate(ByRef s As String) As String
    EnDate = "#" & s & "#"
End Function


Public Sub ReadGlobalString()
    Dim pvrset As ADODB.Recordset
    
    Set pvrset = cnTuition.Execute("SELECT x from globals")
    
    With pvrset
        .MoveFirst
        strSchoolYear = .Fields("x")
        .MoveNext
        strSemester = .Fields("x")
    End With
    
    Set pvrset = Nothing
End Sub


Public Sub ImportTuitionFees(ByRef srcpath As String)
    Dim xl As Object
    Dim xls As Object
    Dim nctr As Long
    Dim rc As ADODB.Recordset
    
    
    On Error Resume Next
    
    Set xl = GetObject(srcpath)
    
    Set xls = xl.sheets("ENTRY")
    
    If Err.Number <> 0 Then
        MsgBox Err.Description
        Exit Sub
    End If
    
    cnTuition.Execute "DELETE * FROM tblTuitionFee"
    
    Set rc = CreateObject("ADODB.RECORDSET")
    
    rc.Open "tblTuitionFee", cnTuition, adOpenKeyset, adLockOptimistic
    
    nctr = 1
    
    With rc
        While Len(xls.Cells(nctr, 2)) > 0
            .AddNew
            .Fields("TFREFNUMBER") = xls.Cells(nctr, 2)
            .Fields("TFTEXT") = xls.Cells(nctr, 3)
            .Fields("TFPRICE") = xls.Cells(nctr, 4)
            .Update
            nctr = nctr + 1
        Wend
    End With
    
    
    
    MsgBox nctr
    
    rc.Close
    
    Set xls = Nothing
    Set xl = Nothing
    
End Sub

Public Sub UpdateMasterList(ByRef sourcepath As String)
    Dim cnMasterlist As Object
    Dim rsMasterlist As Object
    Dim rstemp As Object
    Dim xsql As String
    Dim n1 As Long, n2 As Long
    Dim ncount As Long
    
    Set cnMasterlist = CreateObject("adodb.connection")
    
    On Error Resume Next
    
    'CLEAR CURRENT CONTENTS OF MASTER LIST TABLE
    
    cnTuition.Execute "DELETE * from tblMasterlist"
    
    cnMasterlist.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                  "Data Source=" & sourcepath
    
    cnMasterlist.CursorLocation = adUseClient
    cnMasterlist.Mode = adModeRead + adModeShareDenyNone
    cnMasterlist.Open
    
    If Err.Number <> 0 Then
        MsgBox Err.Number & vbCrLf & Err.Description
        Exit Sub
    End If

    xsql = "SELECT cl.syrefnumber,ml.studrefnumber, ml.idnum, ml.engname, ml.chiname," & _
                    " (elv.levelprefix & '-' & ecl.classtext) AS engsection," & _
                    " (clv.chileveltext & ccl.chiclasstext) AS chisection," & _
                    " sl.sex2,cl.isnew," & _
                    " cl.englevel, cl.engclass, cl.chilevel, cl.chiclass" & _
                    " FROM " & _
                    " (((((tblMasterlist AS ml LEFT JOIN tblClassifiedlist AS cl ON ml.STUDREFNUMBER=cl.STUDREFNUMBER)" & _
                                                      " LEFT JOIN tblLevelList AS elv ON cl.EngLevel=elv.levelid)" & _
                                                      " LEFT JOIN tblLevelList AS clv ON cl.ChiLevel=clv.levelid)" & _
                                                      " LEFT JOIN tblClasslist AS ecl ON cl.EngClass=ecl.Classid)" & _
                                                      " LEFT JOIN tblClasslist AS ccl ON cl.ChiClass=ccl.Classid)" & _
                                                      " LEFT JOIN tblSexList as sl on ml.studsex = sl.sexid" & _
                                                      " WHERE (cl.syrefnumber=ecl.syid and cl.syrefnumber=ccl.syid) or cl.syrefnumber is null" & _
                                                " ORDER BY ml.engname asc,cl.syrefnumber desc,cl.semester asc"
        
        
     Set rsMasterlist = cnMasterlist.Execute(xsql)
     Set rstemp = CreateObject("adodb.recordset")
     rstemp.Open "tblMasterList", cnTuition, adOpenKeyset, adLockOptimistic
        
     With rsMasterlist
            .MoveFirst
            
            n1 = 1999999999
            
            While Not .EOF
                n2 = .Fields("STUDREFNUMBER")
                
                If n1 <> n2 Then
                    rstemp.AddNew
                    rstemp.Fields("STUDREFNUMBER") = .Fields("STUDREFNUMBER")
                    rstemp.Fields("IDNUM") = .Fields("IDNUM")
                    rstemp.Fields("ENGNAME") = .Fields("ENGNAME")
                    rstemp.Fields("CHINAME") = .Fields("CHINAME")
                    rstemp.Fields("SEX") = .Fields("SEX2")
                    rstemp.Fields("ENGSECTION") = .Fields("engsection")
                    rstemp.Fields("CHISECTION") = .Fields("chisection")
                    rstemp.Fields("ENGLEVEL") = .Fields("englevel")
                    rstemp.Fields("CHILEVEL") = .Fields("chilevel")
                    rstemp.Fields("ENGCLASS") = .Fields("ENGCLASS")
                    rstemp.Fields("CHICLASS") = .Fields("CHICLASS")
                    rstemp.Fields("ISNEW") = .Fields("isnew")
                    rstemp.Update
                    n1 = n2
                    
                    ncount = ncount + 1
                    If Err.Number <> 0 Then
                        MsgBox Err.Description
                    End If
                End If
                
                .MoveNext
            Wend
    End With
        
    MsgBox ncount & " records retrieved", vbOKOnly, "MASTER LIST UPDATE"
        
    rstemp.Close
    cnMasterlist.Close
    
    Set cnMasterlist = Nothing
    Set rstemp = Nothing
End Sub



