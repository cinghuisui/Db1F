Attribute VB_Name = "ModConnection"
Option Explicit
Public rs As ADODB.Recordset
Public cn As ADODB.Connection
Public Const AT As String = "Administrasi Sumur Pantau"

Public Const AppServer    As String = "192.168.56.1" '"192.168.9.9\PKBSQL,1433"
Const AppDB               As String = "MyLASP"
Const AppUser             As String = "sa"
Const AppPwd              As String = "123456"

Function ActiveCn() As String
    ActiveCn = "Provider=SQLOLEDB.1;Persist Security Info=False;User ID=" & AppUser & "; Password=" & AppPwd & ";Initial Catalog=" & AppDB & ";Data Source=" & AppServer
   ' ActiveCn = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=KoperasiAccounting;Data Source=(local)" '(local)" 'K199"
End Function
Public Sub ControlCentreForm(frm As Form, ctrl As Control)
    ctrl.Left = ((frm.Width - ctrl.Width) / 2) - 100
    ctrl.Top = ((frm.Height - ctrl.Height) / 2) - 100
End Sub

Public Sub LoadCentreForm(frm As Form)
    frm.Left = ((MDIProject.Width - frm.Width) / 2) '- 100
    frm.Top = (MDIProject.Height - frm.Height) / 8
End Sub

Public Sub FormSize(Tinggi As Double, Lebar As Double, frm As Form)
    frm.Height = Tinggi
    frm.Width = Lebar
End Sub
 
Public Sub HapusGrid(ocx As VSFlexGrid, baris As Integer)
    Dim i As Long
    ocx.Rows = baris + 1

    For i = 0 To ocx.Cols - 1
        ocx.TextMatrix(baris, i) = ""
    Next i

End Sub
'dari Kian Gie
Public Function DecryptText(strText As String, ByVal strPwd As String)
    Dim i As Integer, c As Integer, strBuff As String, d As Integer
    #If Not CASE_SENSITIVE_PASSWORD Then
        strPwd = UCase$(strPwd)
    #End If
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c - Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            'd = Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    DecryptText = strBuff
End Function
Public Function EncryptText(strText As String, ByVal strPwd As String)
    Dim i       As Integer, c As Integer
    Dim strBuff As String

    #If Not CASE_SENSITIVE_PASSWORD Then
        'Convert password to upper case
        'if not case-sensitive
        strPwd = UCase$(strPwd)
    #End If
    'Encrypt string
    If Len(strPwd) Then
        For i = 1 To Len(strText)
            c = Asc(Mid$(strText, i, 1))
            c = c + Asc(Mid$(strPwd, (i Mod Len(strPwd)) + 1, 1))
            strBuff = strBuff & Chr$(c And &HFF)
        Next i
    Else
        strBuff = strText
    End If
    EncryptText = strBuff
End Function

Function GetDate() As String
    Dim rs As New ADODB.Recordset
    On Error GoTo EH
    Call rs.Open("SELECT CONVERT(varchar, GETDATE(), 103) + ' ' + CONVERT(varchar, GETDATE(), 108) AS Now", ActiveCn)
    rs.MoveFirst
    GetDate = rs!Now
    Set rs = Nothing
    Exit Function
    
EH:
    Set rs = Nothing
    Call ErrMsg(Err)
End Function
Sub ErrMsg(Error As Object)
    Dim msg As String
    Select Case Error.Number
        Case 0
        Case -2147467259: msg = "Connection to database server is broken !"
        Case 13: msg = "Numeric value is not valid !"
        Case 3021: msg = "No data !"
        Case Else: msg = "Error : " & Error.Number & " : " & Error.Description
    End Select
    If Error.Number <> 0 Then Call MsgBox(msg, vbExclamation, AT)
End Sub

Public Sub ConvertToExcel(cmDialog As CommonDialog, fg1 As VSFlexGrid, frm As Form)
On Error GoTo ErrorHandler
Dim mPath As String
cmDialog.Filter = "Microsoft Office Excel |*.xls|"
cmDialog.FilterIndex = 1
cmDialog.ShowSave
If cmDialog.FileName = "" Then Exit Sub
mPath = cmDialog.FileName
With fg1
    frm.MousePointer = MousePointerConstants.vbHourglass
    .SaveGrid mPath, flexFileExcel, SaveExcelSettings.flexXLSaveFixedCells
    frm.MousePointer = MousePointerConstants.vbDefault
End With

ErrorHandler:
If Err.Number <> 0 Then MsgBox Err.Description, vbExclamation, "Informasi"
End Sub

Function SetCondition(OldCond As String, NewCond As String) As String
    SetCondition = IIf((OldCond <> ""), OldCond & " AND " & NewCond, NewCond)
End Function

 Public Sub HLText(ByRef srcText As TextBox)
        srcText.BackColor = &HC0FFFF
End Sub
Public Sub DHLText(ByRef srcText As TextBox)
        srcText.BackColor = &H8000000F
End Sub

Function GetValue(sql As String) As Double
    Dim rs As New ADODB.Recordset
    On Error GoTo EH
    GetValue = 0
    Call rs.Open(sql, ActiveCn, adOpenKeyset)
    'rs.MoveFirst
    GetValue = CStr(rs!GetValue)
    Set rs = Nothing
    Exit Function

EH:
    Set rs = Nothing
    If Err.Number <> 3021 Then MsgBox Err.Description
End Function

Function GetValueString(sql As String) As String
    Dim rs As New ADODB.Recordset
    On Error GoTo EH
    GetValueString = ""
    Call rs.Open(sql, ActiveCn, adOpenKeyset)
    'rs.MoveFirst
    GetValueString = CStr(rs!GetValueString)
    Set rs = Nothing
    Exit Function

EH:
    Set rs = Nothing
    If Err.Number <> 3021 Then MsgBox Err.Description
End Function


Function LoadWilayah(cbo As ComboBoxLB)

    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim sql As String
    Dim i   As Long

    sql = "Select KodeWilayah, Wilayah from tblMstWilayah where KodeWilayah like '" & IIf(MDIProject.GroupUser = "ADM WILAYAH", MDIProject.Wilayah, "") & "%'  order by KodeWilayah Asc"
    rs.Open sql, ActiveCn, adOpenKeyset, adLockReadOnly

    If rs.RecordCount > 0 Then
        cbo.ColumnCount = 2
        cbo.ColumnWidths = "2500;0"
        rs.MoveFirst

        For i = 0 To rs.RecordCount - 1
            cbo.AddItem rs!Wilayah & ";" & rs!KodeWilayah
            rs.MoveNext
        Next i

    End If

End Function
