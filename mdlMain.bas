Attribute VB_Name = "mdlMain"
'202508
Public gConnAccess As ADODB.Connection
Public gDBAccessName As String
Public gSectionLocation As String
Public gEMPIOnly As String

'202306
'Public declaration
Public gSecurityPost As String
Public gConn As ADODB.Connection
Public gAfterLogin As Boolean
Public gConStatSystem As Boolean
Public gDBName As String, gTrue As String, gFalse As String, mUserID As String, mPassword As String


Public gShiftNow As String
Public gBottomMsg() As String
Public gAutoCut As Boolean, gPrintCoupon As Boolean
Public gBottomNews As String
Public gLineNo As Integer
Public gRptDestination As String
Public gTrxID As Long

'Common for All Programs
Public gUpdateString As String, gCreatedString As String
Public gSetFocus As Boolean


Public gEnableFocus As Boolean, gCanteenPC As Boolean

Public gAppName As String
'Public clsConn As New clsCommon
Public gAppVersion As String

Option Explicit
'March 2020
Public gCoDateFormat As String, gDbDateFormat As String

'Configuration file purpose
Private mFSO As New FileSystemObject
Private mFile1 As File
Private mts As TextStream
Private mFileSpec As String
Private rsData As ADODB.Recordset

'Reading User PC core information
Private Declare Function GetIpAddrTable_API Lib "IpHlpApi" Alias "GetIpAddrTable" (pIPAddrTable As Any, pdwSize As Long, ByVal bOrder As Long) As Long
Public Declare Function GetComputerName Lib "KERNEL32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias _
    "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Dim buffer As String * 512, length As Long
Public gConnFailed As Boolean
Public gOnTrans As Boolean
Public gServerStopped As Boolean
Public gIPAddress As String
Public gWinLogin As String
Public gComputerName As String

'Database purpose
Public gProviderName As String
Public gDataSource As String, gDataSource2 As String, gDataSource3 As String
Public gCatalog As String, gCatalog2 As String, gCatalog3 As String
Public gLocalProvider As String, gLocalDB As String
Public gResume As Boolean, gResumeNext As Boolean
Public gUserIDLCS As String, gPasswordLCS As String
Public gUserIDPMS As String, gPasswordPMS As String
Public gUserIDEMP As String, gPasswordEMP As String

Public gLanguage As String, gLogin As Boolean
Public gShift1Start As String, gShift2Start As String

Public gDATE_PREF As String

Sub Main()
    On Error GoTo ErrHandler
    gCoDateFormat = "dd-MMM-yyyy"
    gDbDateFormat = "yyyy-mm-dd"
    gRptDestination = "Printer"
    gLineNo = 7
    gComputerName = String(50, Chr(0))
    Call GetComputerName(gComputerName, 50)
    gComputerName = Left(gComputerName, 8)
    gIPAddress = GetIPAddress
    If GetUserName(buffer, Len(buffer)) Then
        ' Search the trailing Null character.
        length = InStr(buffer, vbNullChar) - 1
        gWinLogin = Left$(buffer, length)
    Else
        gWinLogin = ""
    End If
    gAppName = "PT. Enkei Indonesia | eAttendance System"
    gAppVersion = App.Major & "." & App.Minor & "." & App.Revision
    gLanguage = "INA"
    gLogin = False
    mFileSpec = App.Path & "\absensi.ini"
    If mFSO.FileExists(mFileSpec) Then
        If DBInfoExist() Then
            Call OpenDB2
            frm2025Main.Show
        Else
            MsgBox "Informasi database dalam file konfigurasi tidak ditemukan atau salah !" & vbCrLf & vbCrLf & _
                    "Hubungi administrator sistem.", vbInformation
            End
        End If
    Else
        MsgBox "File konfiguration tidak ditemukan !" & vbCrLf & vbCrLf & _
                "Hubungi administrator sistem.", vbInformation
        End
    End If
    
    Exit Sub

ErrHandler:
    RecoverError
End Sub

Private Function GetIPAddress()
    Dim IpAddrs
    IpAddrs = GetIpAddrTable
    Dim i As Integer
    For i = LBound(IpAddrs) To UBound(IpAddrs)
         If Len(IpAddrs(i)) > 9 Then GetIPAddress = IpAddrs(i)
    Next
End Function

Private Function DBInfoExist() As Boolean
    'Read connection provider, database path & name
    Dim vText As String
    
    Set mFile1 = mFSO.GetFile(App.Path & "\absensi.ini")
    Set mts = mFile1.OpenAsTextStream(ForReading)
    Do While Not mts.AtEndOfStream
        vText = mts.ReadLine
        If UCase(Left(vText, 10)) = "PROVIDER= " Then
            gProviderName = Trim(Right(vText, Len(vText) - 10))
        End If
        
        If UCase(Left(vText, 13)) = "DATA SOURCE= " Then gDataSource = Trim(Right(vText, Len(vText) - 13))
        If UCase(Left(vText, 17)) = "INITIAL CATALOG= " Then gCatalog = Trim(Right(vText, Len(vText) - 17))
        If UCase(Left(vText, 20)) = "DATA SOURCE LOTCTL= " Then gDataSource2 = Trim(Right(vText, Len(vText) - 20))
        If UCase(Left(vText, 24)) = "INITIAL CATALOG LOTCTL= " Then gCatalog2 = Trim(Right(vText, Len(vText) - 24))
        If UCase(Left(vText, 16)) = "SHIFT1STARTTIME=" Then gShift1Start = Trim(Right(vText, Len(vText) - 16))
        If UCase(Left(vText, 16)) = "SHIFT2STARTTIME=" Then gShift2Start = Trim(Right(vText, Len(vText) - 16))
        If UCase(Left(vText, 16)) = "DATA SOURCE EMP=" Then gDataSource3 = Trim(Right(vText, Len(vText) - 16))
        If UCase(Left(vText, 20)) = "INITIAL CATALOG EMP=" Then gCatalog3 = Trim(Right(vText, Len(vText) - 20))
        If UCase(Left(vText, 10)) = "USERIDPMS=" Then gUserIDPMS = Trim(Right(vText, Len(vText) - 10))
        If UCase(Left(vText, 12)) = "PASSWORDPMS=" Then gPasswordPMS = Trim(Right(vText, Len(vText) - 12))
        If UCase(Left(vText, 10)) = "USERIDEMP=" Then gUserIDEMP = Trim(Right(vText, Len(vText) - 10))
        If UCase(Left(vText, 13)) = "PASSWORDEMP= " Then gPasswordEMP = Trim(Right(vText, Len(vText) - 13))
        If UCase(Left(vText, 8)) = "USERID= " Then gUserIDLCS = Trim(Right(vText, Len(vText) - 8))
        If UCase(Left(vText, 10)) = "PASSWORD= " Then gPasswordLCS = Trim(Right(vText, Len(vText) - 10))
        If UCase(Left(vText, 14)) = "SECURITYPOST= " Then gSecurityPost = Trim(Right(vText, Len(vText) - 14))
        If UCase(Left(vText, 11)) = "EMPI ONLY= " Then gEMPIOnly = Trim(Right(vText, Len(vText) - 11))
'DBAccess = EKIDAtt.accdb
        If UCase(Left(vText, 10)) = "DBACCESS= " Then gDBAccessName = Trim(Right(vText, Len(vText) - 10))
        
        If UCase(Left(vText, 8)) = "USERID= " Then
            gUserIDLCS = Trim(Right(vText, Len(vText) - 8))
        End If
        If UCase(Left(vText, 10)) = "PASSWORD= " Then gPasswordLCS = Trim(Right(vText, Len(vText) - 10))
            
    Loop
    Select Case gProviderName
    Case "SQLOLEDB.1"
        If (gDataSource = "" Or gCatalog = "") Then
            DBInfoExist = False
        Else
            DBInfoExist = True
        End If
    Case Else
        DBInfoExist = False
    End Select
    
    mts.Close
End Function

Private Function GetIpAddrTable()
   Dim Buf(0 To 511) As Byte
   Dim BufSize As Long: BufSize = UBound(Buf) + 1
   Dim rc As Long
   rc = GetIpAddrTable_API(Buf(0), BufSize, 1)
   If rc <> 0 Then Err.Raise vbObjectError, , "GetIpAddrTable failed with return value " & rc
   Dim NrOfEntries As Integer: NrOfEntries = Buf(1) * 256 + Buf(0)
   If NrOfEntries = 0 Then GetIpAddrTable = Array(): Exit Function
   ReDim IpAddrs(0 To NrOfEntries - 1) As String
   Dim i As Integer
   For i = 0 To NrOfEntries - 1
      Dim j As Integer, s As String: s = ""
      For j = 0 To 3: s = s & IIf(j > 0, ".", "") & Buf(4 + i * 24 + j): Next
      IpAddrs(i) = s
      Next
   GetIpAddrTable = IpAddrs
End Function

Public Sub OpenDB2()
Dim vFileDB As String
    
    Set gConnAccess = New ADODB.Connection
    With gConnAccess
        .ConnectionTimeout = 15
        .CursorLocation = adUseClient
        .Provider = "Microsoft.Jet.OLEDB.4.0"
        vFileDB = App.Path & "\" & gDBAccessName
        
        .Properties("Data Source").Value = vFileDB
        gDATE_PREF = "#"
        gFalse = "0"
        gTrue = "1"
        .Open
    End With
    
End Sub

Public Sub OpenDB()

'Opens the named database
'On Error GoTo ErrHandler
  
    Set gConn = New ADODB.Connection
    gConStatSystem = False
    With gConn
        .ConnectionTimeout = 15
        .CursorLocation = adUseClient
        .Provider = gProviderName
        'MS-Access
        If gProviderName = "Microsoft.Jet.OLEDB.4.0" Then
            .Properties("Data Source").Value = gDBName
            .Properties("Jet OLEDB:Database Password") = "asd987654321"
            gDATE_PREF = "#"
            gFalse = "0"
            gTrue = "1"
        Else
            'SQL Server
            .Properties("Data Source").Value = gDataSource
            .Properties("Initial Catalog").Value = gCatalog
            .Properties("Integrated Security").Value = "SSPI"
            .Properties("Persist Security Info").Value = True
            .Properties("User ID").Value = mUserID
            .Properties("Password").Value = mPassword
            gDATE_PREF = "'"
            gFalse = "0"
            gTrue = "1"
        End If
        .Open
    End With
    gConStatSystem = True
    Exit Sub
    
ErrHandler:
   If Err.Number <> 0 Then
     MsgBox Err.Description, , "Error Message"
   End If
   
End Sub

Public Sub CloseDB()
    gConn.Close
    Set gConn = Nothing
    
End Sub
