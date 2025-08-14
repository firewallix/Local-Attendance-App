Attribute VB_Name = "mdlWriteLogs"
Private mString As String
Option Explicit

Public Sub WriteLogs(ByVal vDescription As String)
'202508 EKID Attendance Access Version
    mString = "INSERT INTO TB_TLogs(FDateTime, FDescription, FComputerID) "
    mString = mString & "VALUES('" & Format(Now, gCoDateFormat) & "','" & vDescription & "','"
    mString = mString & gComputerName & "')"
    gConnAccess.Execute mString
End Sub

