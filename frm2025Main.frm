VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frm2025Main 
   Caption         =   "Attendance"
   ClientHeight    =   7425
   ClientLeft      =   105
   ClientTop       =   735
   ClientWidth     =   8880
   Icon            =   "frm2025Main.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7425
   ScaleWidth      =   8880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Left            =   480
      Top             =   6480
   End
   Begin MSComctlLib.StatusBar sb1 
      Align           =   2  'Align Bottom
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   6930
      Width           =   8880
      _ExtentX        =   15663
      _ExtentY        =   873
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Date"
            Object.Tag             =   "Date"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Time"
            Object.Tag             =   "Time"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "CompName"
            Object.Tag             =   "CompName"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "Version"
            Object.Tag             =   "Version"
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Left            =   0
      Top             =   6600
   End
   Begin VB.Frame fra1 
      Height          =   7215
      Left            =   2520
      TabIndex        =   2
      Top             =   720
      Width           =   10935
      Begin VB.Frame fraLoading 
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   2520
         Visible         =   0   'False
         Width           =   10695
         Begin MSComctlLib.ProgressBar pb1 
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1320
            Width           =   10455
            _ExtentX        =   18441
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Copying master database ..."
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   30
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   11
            Top             =   360
            Width           =   9975
         End
      End
      Begin VB.CommandButton cmdEntry 
         Appearance      =   0  'Flat
         Caption         =   "v"
         Height          =   375
         Left            =   4200
         TabIndex        =   1
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox txtEntry 
         Height          =   375
         Left            =   2160
         TabIndex        =   0
         Top             =   240
         Width           =   1935
      End
      Begin MSFlexGridLib.MSFlexGrid fgIN 
         Height          =   5535
         Left            =   120
         TabIndex        =   3
         Top             =   1200
         Width           =   8535
         _ExtentX        =   15055
         _ExtentY        =   9763
         _Version        =   393216
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSFlexGridLib.MSFlexGrid fgOUT 
         Height          =   5535
         Left            =   5400
         TabIndex        =   4
         Top             =   1200
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   9763
         _Version        =   393216
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblTime 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   8880
         TabIndex        =   9
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lblCmdOut 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "OUT | F8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   7080
         TabIndex        =   8
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblCmdIN 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         Caption         =   "IN | F2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblMode 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scan sebelum MULAI BEKERJA"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   10575
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PT. ENKEI INDONESIA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2160
      TabIndex        =   13
      Top             =   120
      Width           =   5010
   End
   Begin VB.Image Image1 
      Height          =   585
      Left            =   120
      Picture         =   "frm2025Main.frx":68C6
      Top             =   120
      Width           =   1980
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuIn 
         Caption         =   "&Incoming"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuOutgoing 
         Caption         =   "&Outgoing"
         Shortcut        =   {F8}
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuUpload 
         Caption         =   "&Upload Data"
      End
      Begin VB.Menu mnuDownload 
         Caption         =   "&Download Cards"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
End
Attribute VB_Name = "frm2025Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'202508
Option Explicit
Private mPunchID As String, mName As String, mSection As String, mTime As String, mRFID As String
Private mHourCounter As Integer
Private mPrevRFID As String, mPrevMode As String
Private mString As String
Private rsA As ADODB.Recordset
Private mModeInOut As String
Private mCounter As Long

Private Sub Form_Load()
    Call FormPreferences
    Call SetupFG
    Call CheckMode
End Sub

Private Sub MakeCenterFrame()
Dim vFormSize As Integer, vFraSize As Integer
Dim vDifferenceSize As Integer
Dim vHalfForm As Integer, vHalfFrame As Integer

    vFormSize = Me.Width
    vFraSize = fra1.Width
    
    vHalfForm = vFormSize / 2
    vHalfFrame = vFraSize / 2
    vDifferenceSize = vHalfForm - vHalfFrame
    fra1.Left = Me.Left + vDifferenceSize
    
    vFormSize = Me.Height
    vFraSize = fra1.Height
    vHalfForm = vFormSize / 2
    vHalfFrame = vFraSize / 2
    vDifferenceSize = (vHalfForm - vHalfFrame) / 2
    fra1.Top = vDifferenceSize
End Sub

Private Sub FormPreferences()
    Set rsA = New ADODB.Recordset
    sb1.Panels("Date").text = Format(Date, "dd-MMM-yyyy")
    sb1.Panels("Time").text = Format(Now, "HH:MM:SS")
    sb1.Panels("CompName").text = gComputerName
    sb1.Panels("Version").text = gAppVersion
    Timer1.Interval = 1000
    Timer1.Enabled = True
    Timer2.Interval = 1000
    Timer2.Enabled = True
    
    lblTime.Caption = Format(Now, "HH:MM:SS")
    fgIN.Width = fra1.Width - 240
    fgOUT.Width = fra1.Width - 240
    fgOUT.Left = fgIN.Left
End Sub

Private Sub SelectTime()
    mModeInOut = CheckTime(lblTime.Caption)
    Select Case mModeInOut
    Case "IN"
        Call lblCmdIN_Click
    Case "OUT"
        lblCmdOut_Click
    Case "XX"
    End Select
End Sub

Private Sub SetupFG()
Dim i As Integer
    With fgIN
        .Cols = 30
        For i = 0 To .Cols - 1
            .ColWidth(i) = 0
        Next i
        .TextMatrix(0, 0) = "PIN"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 4) = "Section"
        .TextMatrix(0, 6) = "Time"
        .TextMatrix(0, 8) = "Mode"
        .ColWidth(0) = 1000
        .ColWidth(2) = 4500
        .ColWidth(4) = 2000
        .ColWidth(6) = 1600
        .ColWidth(8) = 1000
        .ColAlignment(0) = 1
        .ColAlignment(6) = 1
    End With
    With fgOUT
        .Cols = 30
        For i = 0 To .Cols - 1
            .ColWidth(i) = 0
        Next i
        .TextMatrix(0, 0) = "PIN"
        .TextMatrix(0, 2) = "Name"
        .TextMatrix(0, 4) = "Section"
        .TextMatrix(0, 6) = "Time"
        .TextMatrix(0, 8) = "Mode"
        .ColWidth(0) = 1000
        .ColWidth(2) = 4500
        .ColWidth(4) = 2000
        .ColWidth(6) = 1600
        .ColWidth(8) = 1000
        .ColAlignment(0) = 1
        .ColAlignment(6) = 1
    End With
End Sub

Private Sub AddToFGIN(ByVal vName As String, ByVal vSection As String, ByVal vTime As String, _
    ByVal vPunchID As String)
    
    Dim vRow As Integer, i As Integer
    With fgIN
        If .Rows > 1 And .TextMatrix(1, 0) <> "" Then
            .Rows = .Rows + 1
        End If
        vRow = .Rows - 1
        .TextMatrix(vRow, 0) = vPunchID
        .TextMatrix(vRow, 2) = vName
        .TextMatrix(vRow, 4) = vSection
        .TextMatrix(vRow, 6) = vTime
        .TextMatrix(vRow, 8) = mModeInOut
        If .Rows < 14 Then
            .TopRow = 1
        ElseIf .Rows > 13 Then
            .TopRow = .Rows - 12
        End If
    End With
End Sub

Private Sub AddToFGOut(ByVal vName As String, ByVal vSection As String, ByVal vTime As String, _
    ByVal vPunchID As String)
    
    Dim vRow As Integer, i As Integer
    With fgOUT
        If .Rows > 1 And .TextMatrix(1, 0) <> "" Then
            .Rows = .Rows + 1
        End If
        vRow = .Rows - 1
        .TextMatrix(vRow, 0) = vPunchID
        .TextMatrix(vRow, 2) = vName
        .TextMatrix(vRow, 4) = vSection
        .TextMatrix(vRow, 6) = vTime
        .TextMatrix(vRow, 8) = mModeInOut
        If .Rows < 16 Then
            .TopRow = 1
        ElseIf .Rows > 15 Then
            .TopRow = .Rows - 14
        End If
    End With
End Sub

Private Sub cmdEntry_GotFocus()
Dim vRFID As String
    vRFID = txtEntry.text
    
    If vRFID = "" Then Exit Sub
    
    If vRFID = mPrevRFID And mModeInOut = mPrevMode Then
        txtEntry.text = ""
        txtEntry.SetFocus
        Exit Sub
    End If
    If ValidRFID(vRFID) Then
        If mModeInOut = "IN" Then
            AddToFGIN mName, mSection, mTime, mPunchID
            WriteData mPunchID, mSection, 1, mPunchID, mName, "", mRFID, mSection
        ElseIf mModeInOut = "OUT" Then
            AddToFGOut mName, mSection, mTime, mPunchID
            WriteData mPunchID, mSection, 2, mPunchID, mName, "", mRFID, mSection
        End If
    Else
        MsgBox "Kartu Absensi tidak valid." & vbCr & "Cek kembali kartu atau hubungi Admin" & vbCr & _
            "Data attendance TIDAK tersimpan", 48
        GFE txtEntry
        txtEntry.SetFocus
        Exit Sub
    End If
    
    txtEntry.text = ""
    txtEntry.SetFocus
    mPrevRFID = vRFID
    mPrevMode = mModeInOut
End Sub

Private Sub WriteData(ByVal vPunchID As String, ByVal vShortDept As String, ByVal vInOut As Integer, _
    ByVal vEmpID As Long, ByVal vName As String, ByVal vCateringID As String, ByVal vRFID As String, _
    ByVal vSection As String)
Dim vText1 As String, vDate As String, vTime As String, vInOutS As String, vHandPunchID As String, vText2 As String
Dim vFinal As String
    vText1 = "31"
    vDate = Format(Date, "yyyymmdd")
    vTime = Format(Now, "HHMM")
    vInOutS = "000" & vInOut
    If Len(vPunchID) = 4 Then
        vHandPunchID = "000000" & vPunchID
    ElseIf Len(vPunchID) = 5 Then
        vHandPunchID = "00000" & vPunchID
    End If
    vText2 = "0001"
    vFinal = vText1 & vDate & vTime & vInOutS & vHandPunchID & vText2
    
    mString = "INSERT INTO TB_TAttendance(FEmployeeID, FPunchID, FDate, FTime, FCompID, FString, FInOut,"
    mString = mString & "FDownloaded, FCreateDate, FCateringID, FRFID, FRemarks, FShortDept, FName)"
    mString = mString & " VALUES(" & vEmpID & ",'" & vPunchID & "','" & Format(Date, "yyyy-MM-dd") & "','"
    mString = mString & Format(Now, "HH:MM:SS") & "','" & gComputerName & "','" & vFinal & "',"
    mString = mString & vInOut & ",0,'" & Format(Now, "yyyy-mm-dd HH:mm:ss") & "','" & vCateringID
    mString = mString & "','" & vRFID & "','','" & vSection & "','" & vName & "')"
    
    gConnAccess.Execute mString
    
    mString = "UPDATE TCOV_MRFID SET FLastCheck = '" & Format(Now, "yyyy/mm/dd hh:mm:ss") & "', FLastPC = '" & gComputerName & "' "
    mString = mString & "WHERE FRFID = '" & vRFID & "' "
    If vInOut = 1 Then
        mString = "UPDATE TB_MRFID SET FLastCheckIN = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "' "
        mString = mString & ", FLastPCIN = '" & gComputerName & "' "
        mString = mString & "WHERE FPunchID = '" & vPunchID & "' "
    Else
        mString = "UPDATE TB_MRFID SET FLastCheckOUT = '" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "'"
        mString = mString & ", FLastPCOUT = '" & gComputerName & "' "
        mString = mString & "WHERE FPunchID = '" & vPunchID & "' "
    End If
    gConnAccess.Execute mString
       
End Sub

Private Function ValidRFID(ByVal vRFID As String) As Boolean
Dim vAllowPIN As Integer
    If Len(vRFID) > 4 Then
        mString = "SELECT * FROM TB_MRFID WHERE FRFID = '" & vRFID & "'"
    ElseIf Len(vRFID) = 4 Then
        mString = "SELECT * FROM TB_MRFID WHERE FPunchID = '" & vRFID & "'"
    End If
    
    rsA.Open mString, gConnAccess, adOpenDynamic, adLockOptimistic
    If rsA.RecordCount > 0 Then
        mPunchID = IIf(IsNull(rsA!FPunchID), "", rsA!FPunchID)
        mName = IIf(IsNull(rsA!FName), "", rsA!FName)
        mSection = IIf(IsNull(rsA!FShortDept), "", rsA!FShortDept)
        mTime = Format(Now, "HH:mm")
        mRFID = IIf(IsNull(rsA!FRFID), "", rsA!FRFID)
        ValidRFID = True
        vAllowPIN = rsA!FAllowPIN
        If Len(vRFID) = 4 And vAllowPIN <> 1 Then ValidRFID = False
    Else
        ValidRFID = False
    End If
    rsA.Close
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Sendkeys "{Tab}"
    End If
End Sub

Private Sub Form_Resize()
    Call MakeCenterFrame
    
End Sub

Private Sub lblCmdIN_Click()
    Dim vTime As String
    Me.BackColor = lblCmdIN.BackColor
    fra1.BackColor = lblCmdIN.BackColor
    mModeInOut = "IN"
    fgIN.Visible = True
    fgOUT.Visible = False
    lblMode.Caption = "Scan sebelum MULAI BEKERJA. Tetap Jaga SAFETY"
End Sub

Private Sub lblCmdOut_Click()
    Dim vTime As String
    Me.BackColor = lblCmdOut.BackColor
    fra1.BackColor = lblCmdOut.BackColor
    mModeInOut = "OUT"
    fgIN.Visible = False
    fgOUT.Visible = True
    lblMode.Caption = "Scan sebelum PULANG. HATI-HATI di Jalan"
End Sub

Private Sub mnuDownload_Click()
    mCounter = 0
    pb1.Value = 0
    Timer2.Enabled = True
    Timer2.Interval = 1000
    txtEntry.SetFocus
    
End Sub

Private Sub mnuExit_Click()
    End
End Sub

Private Sub mnuIn_Click()
    Call lblCmdIN_Click
End Sub

Private Sub mnuOutgoing_Click()
    Call lblCmdOut_Click
End Sub

Private Function CountUploadData() As Long
On Error GoTo ErrHandler
Dim vTotRec As Integer
Dim vPunchID As String, vShortDept As String, vDate As Date, vTime As Date, _
    vCompID As String, vString As String, vInOut As Integer, vName As String, vRFID As String, _
    vCreateDate As Date
Dim vTrxID As Long, vLogDesc As String


    mString = "SELECT * FROM TB_TAttendance WHERE FUploaded = 0"
    With rsA
        .Open mString, gConnAccess, adOpenDynamic, adLockOptimistic
        vTotRec = .RecordCount
        If vTotRec = 0 Then
            .Close
            Exit Function
        End If
        Call OpenDB
        Do While Not .EOF
            vPunchID = !FPunchID
            vShortDept = !FShortDept
            vDate = Format(!FDate, "mm/dd/yyyy")
            vTime = Format(!FTime, "HH:mm:ss")
            vCompID = !FCompID
            vString = !FString
            vInOut = !FInOut
            vTrxID = !FTrxID
            vName = !FName
            
            vCreateDate = !FCreateDate
            mString = "INSERT INTO TCOV_TAttendance (FEmployeeID, FShortDept, FPunchID, "
            mString = mString & "FDate, FTime, FCompID, FString, FInOut, FDownloaded, FCateringID, "
            mString = mString & "FCreateDate, FName, FRemarks, FRFID) "
            mString = mString & "VALUES(" & vPunchID & ",'" & vShortDept & "','" & vPunchID & "'"
            mString = mString & ",'" & vDate & "','" & vTime & "','" & vCompID & "','" & vString
            mString = mString & "'," & vInOut & ",0,'','" & vCreateDate & "','" & vName & "','','"
            mString = mString & vRFID & "')"
            gConn.Execute mString
            
            If vInOut = 1 Then
                mString = "UPDATE TCOV_MRFID SET FLastCheckIN = '" & vDate & "', FLastPCIN = '" & vCompID & "' "
                mString = mString & "WHERE FPunchID = '" & vPunchID & "'"
            ElseIf vInOut = 2 Then
                mString = "UPDATE TCOV_MRFID SET FLastCheckOUT = '" & vDate & "', FLastPCOUT = '" & vCompID & "' "
                mString = mString & "WHERE FPunchID = '" & vPunchID & "'"
            End If
            gConn.Execute mString
            
            mString = "UPDATE TB_TAttendance SET FUploaded = 1 WHERE FTrxID = " & vTrxID
            gConnAccess.Execute mString
            
            .MoveNext
        Loop
        .Close
    End With
    CountUploadData = vTotRec
    vLogDesc = "Uploaded " & vTotRec & " data"
    WriteLogs vLogDesc
    
    Call CloseDB
    Exit Function
ErrHandler:
    MsgBox Err.Description
    Err.Clear
    
End Function

Private Sub mnuUpload_Click()
Dim vTotRec As Long
    vTotRec = CountUploadData
    MsgBox vTotRec & " attendance data uploaded to main database", 64
    
End Sub

Private Sub Timer1_Timer()
Dim vInOut As String, vTotRec As Long
    
    sb1.Panels("Date").text = Format(Date, "dd-MMM-yyyy")
    sb1.Panels("Time").text = Format(Now, "HH:MM:SS")
    lblTime.Caption = Format(Now, "HH:mm:ss")
    mHourCounter = mHourCounter + 1
    If mHourCounter = 3600 Then
        mHourCounter = 0
        Call GetMessages
    End If
    If lblTime.Caption = "00:00:00" Then
        Call ClearFGIN
        Call ClearFGOut
    ElseIf lblTime.Caption = "12:00:00" Then
        Call ClearFGIN
        Call ClearFGOut
    End If
    If Right(lblTime.Caption, 2) = "00" Then Call CheckMode
    
    If lblTime.Caption = "07:10:00" Then
        vTotRec = CountUploadData
    ElseIf lblTime.Caption = "19:30:00" Then
        vTotRec = CountUploadData
    End If
    If Not gSetFocus Then Exit Sub
    
End Sub

Private Sub CheckMode()
    mModeInOut = CheckTime(lblTime.Caption)
    Select Case mModeInOut
        Case "IN"
            Call lblCmdIN_Click
        Case "OUT"
            Call lblCmdOut_Click
    End Select
End Sub

Private Sub GetMessages()
    
End Sub

Private Sub ClearFGIN()
Dim i As Integer
    With fgIN
        .Rows = 2
        For i = 0 To .Cols - 1
            .TextMatrix(1, i) = ""
        Next i
    End With
End Sub

Private Sub ClearFGOut()
Dim i As Integer
    With fgOUT
        .Rows = 2
        For i = 0 To .Cols - 1
            .TextMatrix(1, i) = ""
        Next i
    End With
End Sub

Private Sub CopyDataRFID()
On Error GoTo ErrHandler
Dim vPunchID As String, vName As String, vDept As String, vRFID As String, vAllowPIN As Integer
Dim vTargetID As Long, vResult As String
Dim vTotRec As Long, vRun As Long, vSection As String, vDescription As String

    Screen.MousePointer = vbHourglass
    pb1.Value = 0
    'Call ClearFG
    Call OpenDB
    
    mString = "SELECT  TCOV_MRFID.FAllowPIN, TCOV_MRFID.FMealAllowed, TCOV_MRFID.FRFID, "
    mString = mString & "TCOV_MRFID.FPunchID, TCOV_MRFID.FAllowPIN, THR_MEmployeeNew.FEmployeeName, "
    mString = mString & "THR_MDepartmentNew.FShortName "
    mString = mString & "FROM TCOV_MRFID LEFT OUTER JOIN THR_MEmployeeNew ON THR_MEmployeeNew.FPunchID "
    mString = mString & "= TCOV_MRFID.FPunchID  LEFT OUTER JOIN THR_MDepartmentNew ON "
    mString = mString & "THR_MDepartmentNew.FID = THR_MEmployeeNew.FDeptID "
    mString = mString & "WHERE TCOV_MRFID.FID > 0 "
    mString = mString & "AND THR_MEmployeeNew.FActive = 1"
    If gEMPIOnly = "Y" Then mString = mString & "AND TCOV_MRFID.FShortDept = 'EMPI'"

    rsA.Open mString, gConn, adOpenDynamic, adLockOptimistic
    vTotRec = rsA.RecordCount
    pb1.Max = vTotRec
    If vTotRec > 0 Then
        mString = "DELETE * FROM TB_MRFID"
        gConnAccess.Execute mString
    End If
    Do While Not rsA.EOF
        vRun = vRun + 1
        pb1.Value = vRun
        
        vPunchID = IIf(IsNull(rsA!FPunchID), "", rsA!FPunchID)
        vName = IIf(IsNull(rsA!FEmployeeName), "", rsA!FEmployeeName)
        vDept = IIf(IsNull(rsA!FShortName), "", rsA!FShortName)
        vRFID = IIf(IsNull(rsA!FRFID), "", rsA!FRFID)
        vAllowPIN = IIf(IsNull(rsA!FAllowPIN), 0, rsA!FAllowPIN)
        
        mString = "INSERT INTO TB_MRFID(FPunchID, FName, FShortDept, FRFID, FAllowPIN) "
        mString = mString & "VALUES('" & vPunchID & "', '" & vName & "','" & vDept & "','" & vRFID & "'," & vAllowPIN & ")"
        vResult = "Added"
        gConnAccess.Execute mString
        
        mString = "UPDATE TCOV_MRFID SET FCopied = FCopied + 1 WHERE FPunchID = '" & vPunchID & "'"
        gConn.Execute mString
        rsA.MoveNext
    Loop
    rsA.Close
    Call CloseDB
    Screen.MousePointer = vbNormal
    fraLoading.Visible = False
    vDescription = "Dowloaded " & vTotRec & " data"
    WriteLogs vDescription
    
    Exit Sub
ErrHandler:
    Err.Clear
    fraLoading.Visible = False
    Screen.MousePointer = vbNormal
    
End Sub

Private Sub Timer2_Timer()
    mCounter = mCounter + 1
    fraLoading.Visible = True
    If mCounter = 2 Then
        CopyDataRFID
        Timer2.Enabled = False
    End If

End Sub
