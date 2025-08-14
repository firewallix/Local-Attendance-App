Attribute VB_Name = "mdlCommon"
Private mString As String
Private rs1 As ADODB.Recordset
Option Explicit

Public Sub RecoverError(Optional rs As ADODB.Recordset, Optional pDispMsg As Boolean = True)
    
    gResume = False
    gResumeNext = False
    
    Select Case Err.Number
        Case 30006        'grid memory allocation fail
            MsgBox "Alokasi memory tampilan data full. " & vbCrLf & _
                   "Jumlah data terlalu banyak." & vbCrLf & _
                   "Sesuaikan kriteria pemilihan data ...", vbInformation
            pDispMsg = False
        Case 3704        'recordset closed
            pDispMsg = False
        Case 3705        'recordset open
            pDispMsg = False
        Case &H80040E4D  'invalid database password
        Case &H80004005  'duplicate key index/referential integrity Access
                pDispMsg = True
        Case -2147217873  'duplicate key index/referential integrity SQL
                pDispMsg = True
        Case &H80040E4E   'row handle referred to deleted row
            pDispMsg = False
        Case &H80040E21   'multiple step error
        Case &H80040E14   'SQL insert/update syntax error
            MsgBox "Simpan data gagal." & vbCrLf & _
                   "Error : " & Err.Description, vbInformation
            pDispMsg = False
        Case 13           'type mismatch
        Case 3021           'BOF or EOF or Deleted
            gResumeNext = True
        Case 94           'invalid use of null
            gResumeNext = True
        Case 32755           'open file dialog canceled
            gResumeNext = True
    End Select
    If pDispMsg Then _
        MsgBox "Error ditemukan." & vbCrLf & _
               "Error : " & Err.Description & vbCrLf & vbCrLf & _
               "Hubungi Administrator sistem.", vbInformation
    Err.Clear
    
End Sub



Public Sub ShowError(ByVal vMessage As String)
    MsgBox vMessage, 48, "Error"
End Sub

Public Sub GFE(mText As TextBox)
    mText.SelStart = 0
    mText.SelLength = Len(mText)
End Sub

Public Function CheckTime(ByVal vTime As Date) As String
    If vTime >= "00:00:00" And vTime <= "05:14:00" Then
        CheckTime = "OUT"
    ElseIf vTime >= "05:15:00" And vTime <= "07:10:00" Then
        CheckTime = "IN"
    ElseIf vTime >= "07:11:00" And vTime <= "13:29:00" Then
        CheckTime = "OUT"
    ElseIf vTime >= "13:30:00" And vTime <= "15:30:00" Then
        CheckTime = "IN"
    ElseIf vTime >= "15:31:00" And vTime <= "16:59:00" Then
        CheckTime = "OUT"
    ElseIf vTime >= "17:00:00" And vTime <= "19:05:00" Then
        CheckTime = "IN"
    ElseIf vTime >= "19:06:00" And vTime <= "20:59:00" Then
        CheckTime = "OUT"
    ElseIf vTime >= "21:00:00" And vTime <= "22:29:00" Then
        CheckTime = "IN"
    ElseIf vTime >= "22:30:00" And vTime <= "23:59:59" Then
        CheckTime = "OUT"
    Else
        CheckTime = "XX"
    End If
End Function

Public Sub Sendkeys(text As Variant, Optional wait As Boolean = False)
   Dim WshShell As Object
   Set WshShell = CreateObject("wscript.shell")
   WshShell.Sendkeys CStr(text), wait
   Set WshShell = Nothing
End Sub
