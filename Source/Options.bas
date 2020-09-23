Attribute VB_Name = "Options"
Option Explicit

Public Const FILE_BEGIN = 0
Public Const GENERIC_WRITE = &H40000000
Public Const OPEN_EXISTING = 3

Public timeDiff         As Date
Public nextUpdate       As Date
Public prefFlag         As Boolean
Public servFlag         As Boolean
Public GetConfigOptions As ConfigOptions
Public repairFile       As Boolean

Public Type ConfigOptions
    setEnableNote       As Integer
    setEnableEventLog   As Integer
    setEnableEmail      As Integer
    setEnablePage       As Integer
    setEnableSLog       As Integer
    setEnableALog       As Integer
    setEnableRCMD       As Integer
    setEnableSMTP       As Integer
    setSMTPHost         As String
    setEmailAddress     As String
    setPageAddress      As String
    setInterval         As Integer
    setMinDriveSpace    As Double
    setSMTPHostID       As String
    setSMTPHostPass     As String
End Type

Public Declare Function DeleteFile Lib "kernel32" _
                        Alias "DeleteFileA" _
                        (ByVal lpFileName As String) As Long
                                                
Public Declare Function SetEndOfFile Lib "kernel32" _
                        (ByVal hFile As Long) As Long
                        
Public Declare Function SetFilePointer Lib "kernel32" _
                        (ByVal hFile As Long, _
                        ByVal lDistanceToMove As Long, _
                        lpDistanceToMoveHigh As Long, _
                        ByVal dwMoveMethod As Long) As Long
                        
Public Declare Function CloseHandle Lib "kernel32" _
                        (ByVal hObject As Long) As Long
                        
Public Declare Function CreateFile Lib "kernel32" _
                        Alias "CreateFileA" _
                        (ByVal lpFileName As String, _
                        ByVal dwDesiredAccess As Long, _
                        ByVal dwShareMode As Long, _
                        ByVal lpSecurityAttributes As Long, _
                        ByVal dwCreationDisposition As Long, _
                        ByVal dwFlagsAndAttributes As Long, _
                        ByVal hTemplateFile As Long) As Long
                        
Public Sub SetNotification()
Dim I As Integer
With frmOptions
    Select Case .chkEnable.Value
        Case 1
            .L1.Enabled = True
            .L2.Enabled = True
            .P1.Enabled = True
            .P2.Enabled = True
            .L6.Enabled = True
            .L7.Enabled = True
            .chkEmail.Enabled = True
            .chkPage.Enabled = True
            .txtEmail.Enabled = True
            .txtPage.Enabled = True
            .optOutlook.Enabled = True
            .optSMTP.Enabled = True
            Call SetMailType
            .chkEmail.Value = 1
            .chkPage.Value = 1
        Case 0
            .L1.Enabled = False
            .L2.Enabled = False
            .P1.Enabled = False
            .P2.Enabled = False
            .L6.Enabled = False
            .L7.Enabled = False
            .chkEmail.Enabled = False
            .chkPage.Enabled = False
            .txtEmail.Enabled = False
            .optOutlook.Enabled = False
            .txtPage.Enabled = False
            .optSMTP.Enabled = False
            For I = 0 To 3
                .L5(I).Enabled = False
            Next
            For I = 0 To 2
                .txtHost(I).Enabled = False
            Next
            .chkEmail.Value = 0
            .chkPage.Value = 0
    End Select
    .cmdApply.Enabled = True
End With
End Sub
Public Sub SetMailType()
Dim I  As Integer
With frmOptions
    If .optOutlook.Value = True Then
        For I = 0 To 3
            .L5(I).Enabled = False
        Next
        For I = 0 To 2
            .txtHost(I).Enabled = False
        Next
    Else
        For I = 0 To 3
            .L5(I).Enabled = True
        Next
        For I = 0 To 2
            .txtHost(I).Enabled = True
        Next
    End If
End With
End Sub
Public Sub SetEventLog()
With frmOptions
    Select Case .chkEventLog.Value
        Case 1
            .L3.Enabled = True
            .L4.Enabled = True
            .chkSystem.Enabled = True
            .chkApp.Enabled = True
            .chkApp.Value = 1
            .chkSystem.Value = 1
        Case 0
            .L3.Enabled = False
            .L4.Enabled = False
            .chkSystem.Enabled = False
            .chkApp.Enabled = False
            .chkApp.Value = 0
            .chkSystem.Value = 0
    End Select
    .cmdApply.Enabled = True
End With
End Sub

Public Sub SetOptions()
With frmOptions
    .txtEmail = Trim(GetConfigOptions.setEmailAddress)
    .txtPage = Trim(GetConfigOptions.setPageAddress)
    .cmbInterval = Trim(GetConfigOptions.setInterval)
    .cmbMinSpace = Trim(GetConfigOptions.setMinDriveSpace)
    .txtHost(0) = Trim(GetConfigOptions.setSMTPHost)
    .txtHost(1) = Base64Decode(Trim(GetConfigOptions.setSMTPHostID))
    .txtHost(2) = Base64Decode(Trim(GetConfigOptions.setSMTPHostPass))
    
    If CBool(GetConfigOptions.setEnableSMTP) Then
        .optSMTP.Value = True
        .optOutlook.Value = False
    Else
        .optSMTP.Value = False
        .optOutlook.Value = True
    End If
    
    If CBool(GetConfigOptions.setEnableNote) Then
        .chkEnable.Value = 1
    Else: .chkEnable.Value = 0
    End If
    Call SetNotification
    
    If CBool(GetConfigOptions.setEnableEventLog) Then
        .chkEventLog.Value = 1
    Else: .chkEventLog.Value = 0
    End If
    Call SetEventLog

    If CBool(GetConfigOptions.setEnableEmail) Then
        .chkEmail.Value = 1
    Else: .chkEmail.Value = 0
    End If
    
    If CBool(GetConfigOptions.setEnablePage) Then
        .chkPage.Value = 1
    Else: .chkPage.Value = 0
    End If
    
    If CBool(GetConfigOptions.setEnableSLog) Then
        .chkSystem.Value = 1
    Else: .chkSystem.Value = 0
    End If
    
    If CBool(GetConfigOptions.setEnableALog) Then
        .chkApp.Value = 1
    Else: .chkApp.Value = 0
    End If
    
    If CBool(GetConfigOptions.setEnableRCMD) Then
        .chkRcmd.Value = 1
    Else: .chkRcmd.Value = 0
    End If

End With
End Sub
Public Function ApplySettings() As Boolean
Dim strSQL          As String

With frmOptions

    If Trim(.cmbInterval.Text) > 1440 Or Trim(.cmbInterval.Text) < 1 Then MsgBox "Interval has to be set between 1 and 1440 minutes (24 hours).", _
        vbExclamation, "Error": ApplySettings = False: Exit Function
    
    If .chkPage.Value = 1 And _
        (Trim(.txtPage.Text) = "" Or InStr(Trim(.txtPage.Text), ".") = 0 Or InStr(Trim(.txtPage.Text), "@") = 0) Or _
        .chkEmail.Value = 1 And _
        (Trim(.txtEmail.Text) = "" Or InStr(Trim(.txtEmail.Text), ".") = 0 Or InStr(Trim(.txtEmail.Text), "@") = 0) Then
        MsgBox "You specified an invalid Email or Page Address.", vbExclamation, "Error:"
        ApplySettings = False
        Exit Function
    End If
    
    If .optSMTP.Value And .chkEnable.Value = 1 And Trim(.txtHost(0).Text) = "" Then
        MsgBox "You did not specify an SMTP Host.", vbExclamation, "Error:"
        ApplySettings = False
        Exit Function
    End If
        
    On Error GoTo ErrorHandler

        GetConfigOptions.setEmailAddress = Trim(.txtEmail.Text)
        GetConfigOptions.setEnableALog = .chkApp.Value
        GetConfigOptions.setEnableRCMD = .chkRcmd.Value
        GetConfigOptions.setEnableEmail = .chkEmail.Value
        GetConfigOptions.setEnableEventLog = .chkEventLog.Value
        GetConfigOptions.setEnableNote = .chkEnable.Value
        GetConfigOptions.setEnablePage = .chkPage.Value
        If .optSMTP.Value Then
            GetConfigOptions.setEnableSMTP = 1
        Else: GetConfigOptions.setEnableSMTP = 0
        End If
        GetConfigOptions.setSMTPHost = Trim(.txtHost(0).Text)
        GetConfigOptions.setSMTPHostID = Base64Encode(Trim(.txtHost(1).Text))
        GetConfigOptions.setSMTPHostPass = Base64Encode(Trim(.txtHost(2).Text))
        GetConfigOptions.setEnableSLog = .chkSystem.Value
        GetConfigOptions.setInterval = Trim(.cmbInterval.Text)
        GetConfigOptions.setMinDriveSpace = Trim(.cmbMinSpace.Text)
        GetConfigOptions.setPageAddress = Trim(.txtPage.Text)

        Set objConn = New ADODB.Connection
        objConn.Open strConn

        strSQL = "UPDATE tblConfig SET "
        strSQL = strSQL & "[EmailAddress]='" & GetConfigOptions.setEmailAddress & "',"
        strSQL = strSQL & "[EnableALog]=" & GetConfigOptions.setEnableALog & ","
        strSQL = strSQL & "[EnableSMTP]=" & GetConfigOptions.setEnableSMTP & ","
        strSQL = strSQL & "[EnableRCMD]=" & GetConfigOptions.setEnableRCMD & ","
        strSQL = strSQL & "[EnableEmail]=" & GetConfigOptions.setEnableEmail & ","
        strSQL = strSQL & "[EnableEventLog]=" & GetConfigOptions.setEnableEventLog & ","
        strSQL = strSQL & "[EnableNote]=" & GetConfigOptions.setEnableNote & ","
        strSQL = strSQL & "[EnablePage]=" & GetConfigOptions.setEnablePage & ","
        strSQL = strSQL & "[EnableSLog]=" & GetConfigOptions.setEnableSLog & ","
        strSQL = strSQL & "[interval]=" & GetConfigOptions.setInterval & ","
        strSQL = strSQL & "[MinDriveSpace]=" & GetConfigOptions.setMinDriveSpace & ","
        strSQL = strSQL & "[PageAddress]='" & GetConfigOptions.setPageAddress & "',"
        strSQL = strSQL & "[smtpHost]='" & GetConfigOptions.setSMTPHost & "',"
        strSQL = strSQL & "[SMTPID]='" & GetConfigOptions.setSMTPHostID & "',"
        strSQL = strSQL & "[SMTPPass]='" & GetConfigOptions.setSMTPHostPass & "';"

        objConn.Execute (strSQL)
        objConn.Close
        Set objConn = Nothing
 
End With

nextUpdate = Now + (Trim(GetConfigOptions.setInterval) / 60 / 24)
frmSrvmtr.TB.Buttons.Item(14).Image = 13
frmSrvmtr.TB.Buttons.Item(14).Value = tbrUnpressed
frmSrvmtr.T1.Enabled = True
ApplySettings = True
Exit Function

ErrorHandler:
ApplySettings = False
MsgBox "Error Saving to Database. Error number: " & Err.Number & ". Error description: " & Err.Description, vbCritical, "Error:"

End Function
Public Sub ApplyPref(sRow As Integer)
Dim selServer   As String
Dim selDrive    As String
Dim strSQL      As String
Dim I           As Integer
On Error GoTo ErrorHandler

With frmPref

        selServer = .L1.Text
        selDrive = .L2.Text
        frmSrvmtr.G.TextMatrix(sRow, 9) = Trim(.txtComment.Text)
        frmSrvmtr.G.TextMatrix(sRow, 10) = .chkNoEmail.Value
        frmSrvmtr.G.TextMatrix(sRow, 11) = .chkNoPage.Value
        frmSrvmtr.G.TextMatrix(sRow, 12) = .chkNoSysLog.Value
        frmSrvmtr.G.TextMatrix(sRow, 13) = .chkNoAppLog.Value
        
        Set objConn = New ADODB.Connection
        objConn.Open strConn

        strSQL = "UPDATE tblServer SET "
        strSQL = strSQL & "[comment]='" & sq(Trim(.txtComment.Text)) & "',"
        strSQL = strSQL & "[noEmail]=" & .chkNoEmail.Value & ","
        strSQL = strSQL & "[noPage]=" & .chkNoPage.Value & ","
        strSQL = strSQL & "[noSysLog]=" & .chkNoSysLog.Value & ","
        strSQL = strSQL & "[noAppLog]=" & .chkNoAppLog.Value & " WHERE ([server]='" & selServer & "' AND "
        strSQL = strSQL & "[sdrive]='" & selDrive & "');"
        
        If .setAll.Value = 1 Then
            strSQL = "UPDATE tblServer SET "
            strSQL = strSQL & "[noEmail]=" & .chkNoEmail.Value & ","
            strSQL = strSQL & "[noPage]=" & .chkNoPage.Value & ","
            strSQL = strSQL & "[noSysLog]=" & .chkNoSysLog.Value & ","
            strSQL = strSQL & "[noAppLog]=" & .chkNoAppLog.Value & " WHERE [server]='" & selServer & "';"
            
            For I = 1 To frmSrvmtr.G.Rows - 1
                If frmSrvmtr.G.TextMatrix(I, 0) = selServer Then
                    frmSrvmtr.G.TextMatrix(I, 10) = .chkNoEmail.Value
                    frmSrvmtr.G.TextMatrix(I, 11) = .chkNoPage.Value
                    frmSrvmtr.G.TextMatrix(I, 12) = .chkNoSysLog.Value
                    frmSrvmtr.G.TextMatrix(I, 13) = .chkNoAppLog.Value
                End If
            Next I
            .setAll.Value = 0
        End If

        objConn.Execute (strSQL)
        objConn.Close
        Set objConn = Nothing
        
End With
Exit Sub

ErrorHandler:
MsgBox "Error Saving to Database. Error number: " & Err.Number & ". Error description: " & Err.Description, vbCritical, "Error:"

End Sub
Public Function OpenConfig(repairFile As Boolean)
Dim fileState   As Boolean
Dim RecNum      As Integer
Dim I           As Integer
Dim nextServer  As Variant
Dim strSQL      As String
Dim rs          As Recordset
Dim eAddr       As Variant
Dim minSpace    As Variant
Dim SMTPH       As Variant
Dim SMTPHPass   As Variant
Dim SMTPHID     As Variant
Dim serv        As Variant

On Error GoTo ErrorHandler
    
    Set objConn = New ADODB.Connection
    objConn.Open strConn

    strSQL = "SELECT * FROM tblServer ORDER BY [server],[sdrive];"

    Set rs = objConn.Execute(strSQL)
    I = 1
    If Not rs.EOF Then
        With frmSrvmtr.G
            Do While Not rs.EOF
                If I = 1 Then
                    .TextMatrix(1, 0) = rs(1)
                Else: .AddItem rs(1)
                End If
                serv = rs(8)
                If IsNull(serv) Then serv = Empty
                .TextMatrix(I, 1) = rs(2)
                .TextMatrix(I, 9) = rs(3)
                .TextMatrix(I, 14) = serv
                .TextMatrix(I, 10) = Abs(CInt(rs(4)))
                .TextMatrix(I, 11) = Abs(CInt(rs(5)))
                .TextMatrix(I, 12) = Abs(CInt(rs(6)))
                .TextMatrix(I, 13) = Abs(CInt(rs(7)))
                Call StatusOK(I)
                I = I + 1
                rs.MoveNext
            Loop
        End With
    End If
    rs.Close
    
    strSQL = "SELECT * FROM tblConfig;"
    Set rs = objConn.Execute(strSQL)
    eAddr = rs(1)
    If IsNull(eAddr) Then eAddr = Empty
    minSpace = rs(12)
    If IsNull(minSpace) Then minSpace = Empty
    SMTPH = rs(13)
    If IsNull(SMTPH) Then SMTPH = Empty
    SMTPHID = rs(14)
    If IsNull(SMTPHID) Then SMTPHID = Empty
    SMTPHPass = rs(15)
    If IsNull(SMTPHPass) Then SMTPHPass = Empty
    
    GetConfigOptions.setEmailAddress = eAddr
    GetConfigOptions.setEnableALog = CInt(rs(2))
    GetConfigOptions.setEnableSMTP = CInt(rs(3))
    GetConfigOptions.setEnableRCMD = CInt(rs(4))
    GetConfigOptions.setEnableEmail = CInt(rs(5))
    GetConfigOptions.setEnableEventLog = CInt(rs(6))
    GetConfigOptions.setEnableNote = CInt(rs(7))
    GetConfigOptions.setEnablePage = CInt(rs(8))
    GetConfigOptions.setEnableSLog = CInt(rs(9))
    GetConfigOptions.setInterval = CInt(rs(10))
    GetConfigOptions.setMinDriveSpace = CDbl(rs(11))
    GetConfigOptions.setPageAddress = minSpace
    GetConfigOptions.setSMTPHost = SMTPH
    GetConfigOptions.setSMTPHostID = SMTPHID
    GetConfigOptions.setSMTPHostPass = SMTPHPass
    With frmSrvmtr
        .Width = rs(16)
        .Height = rs(17)
        .Top = rs(18)
        .Left = rs(19)
        .G.Height = rs(20)
        .statusLW.Height = rs(21)
    End With
    rs.Close
    frmSrvmtr.mnuRcmd.Enabled = CBool(GetConfigOptions.setEnableRCMD)
    objConn.Close
    Set objConn = Nothing
    
nextUpdate = Now + (Trim(GetConfigOptions.setInterval) / 60 / 24)
Exit Function

ErrorHandler:
MsgBox "Error Opening Database. Error number: " & Err.Number & ". Error description: " & Err.Description, vbCritical, "Error:"
        
End Function

Public Function BuildPrefList(R As Integer) As Boolean
Dim I           As Integer
Dim C           As Integer
Dim isInList    As Boolean
Dim Server      As String
Dim Drive       As String

    Server = frmSrvmtr.G.TextMatrix(R, 0)
    Drive = frmSrvmtr.G.TextMatrix(R, 1)
    With frmPref
        prefFlag = True
        .L1.AddItem (Server)
        .L1.Text = .L1.List(0)
        .L2.AddItem (Drive)
        .L2.Text = .L2.List(0)
        prefFlag = False
        .txtComment.Text = frmSrvmtr.G.TextMatrix(R, 9)
        '.L3.Caption = frmSrvmtr.G.TextMatrix(R, 9)
        .chkNoEmail.Value = frmSrvmtr.G.TextMatrix(R, 10)
        .chkNoPage.Value = frmSrvmtr.G.TextMatrix(R, 11)
        .chkNoSysLog.Value = frmSrvmtr.G.TextMatrix(R, 12)
        .chkNoAppLog.Value = frmSrvmtr.G.TextMatrix(R, 13)
        .L4.Caption = R
        For I = 1 To frmSrvmtr.G.Rows - 1
            For C = 0 To .L1.ListCount
                If frmSrvmtr.G.TextMatrix(I, 0) = .L1.List(C) Then
                    isInList = True
                    Exit For
                End If
            Next C
            If Not isInList Then .L1.AddItem (frmSrvmtr.G.TextMatrix(I, 0))
            isInList = False
            If frmSrvmtr.G.TextMatrix(I, 0) = Server And _
                Not frmSrvmtr.G.TextMatrix(I, 1) = Drive Then
                .L2.AddItem (frmSrvmtr.G.TextMatrix(I, 1))
            End If
        Next I
    End With
    BuildPrefList = True

End Function


Public Function BuildServiceList(Server As String) As Boolean
Dim I           As Integer
Dim C           As Integer
Dim isInList    As Boolean
Dim strServ     As Variant
Dim strSQL      As String
Dim rs          As Recordset
Dim strArray    As Variant

With frmServ
        servFlag = True
        .L1.Clear
        .L3.Clear
        .L1.AddItem (Server)
        .L1.Text = .L1.List(0)
        servFlag = False
        
        Set objConn = New ADODB.Connection
        objConn.Open strConn
        
        strSQL = "SELECT MonServices FROM tblServer WHERE [server] = '" & Server & "';"
        Set rs = objConn.Execute(strSQL)
        
        If Not rs.EOF Then strServ = rs(0)
        If IsNull(strServ) Then strServ = Empty
        rs.Close
        If Len(strServ) > 0 Then
            strArray = Split(strServ, "|")
            For I = 0 To UBound(strArray)
                .L3.AddItem strArray(I)
            Next I
        End If
        
        objConn.Close
        Set objConn = Nothing
        
        For I = 1 To frmSrvmtr.G.Rows - 1
            For C = 0 To .L1.ListCount
                If frmSrvmtr.G.TextMatrix(I, 0) = .L1.List(C) Then
                    isInList = True
                    Exit For
                End If
            Next C
            If Not isInList Then .L1.AddItem (frmSrvmtr.G.TextMatrix(I, 0))
            isInList = False
        Next I
    Call EnumSystemServices(Server, strServ)
    BuildServiceList = True
    .Lb1(0).Caption = .L2.ListItems.Count
    .Lb1(1).Caption = .L3.ListCount
End With
End Function

Public Function EnumSystemServices(selServer As String, monServ As Variant) As Long
Dim hSCManager      As Long
Dim pntr()          As ENUM_SERVICE_STATUS
Dim cbBuffSize      As Long
Dim cbRequired      As Long
Dim dwReturned      As Long
Dim hEnumResume     As Long
Dim cbBuffer        As Long
Dim success         As Long
Dim I               As Long
Dim J               As Long
Dim sSvcName        As String
Dim sDispName       As String
Dim dwState         As Long
Dim insRet          As ListItem
Dim strArray        As Variant
Dim sItem           As String
Dim isSelected      As Boolean
Dim pingStatus      As String

pingStatus = PingServer(selServer)
If Not Left(pingStatus, 1) = 0 Then
    frmServ.L2.ListItems.Clear
    frmServ.Label5.Visible = True
    Exit Function
End If
If Len(monServ) > 0 Then
    strArray = Split(monServ, "|")
End If
frmServ.Label5.Visible = False

hSCManager = OpenSCManager(selServer, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
If hSCManager <> 0 Then
    success = EnumServicesStatus(hSCManager, SERVICE_WIN32, SERVICE_STATE_ALL, _
                                   ByVal &H0, &H0, cbRequired, dwReturned, hEnumResume)
    If success = 0 And Err.LastDllError = ERROR_MORE_DATA Then
        cbBuffer = (cbRequired \ SIZEOF_SERVICE_STATUS) + 1
        ReDim pntr(0 To cbBuffer)
        cbBuffSize = cbBuffer * SIZEOF_SERVICE_STATUS
        hEnumResume = 0
        If EnumServicesStatus(hSCManager, SERVICE_WIN32, SERVICE_STATE_ALL, pntr(0), _
                               cbBuffSize, cbRequired, dwReturned, hEnumResume) Then
            With frmServ.L2
                .ListItems.Clear
                For I = 0 To dwReturned - 1
                    sDispName = GetStrFromPtrA(ByVal pntr(I).lpDisplayName)
                    sSvcName = GetStrFromPtrA(ByVal pntr(I).lpServiceName)
                    dwState = pntr(I).ServiceStatus.dwCurrentState
                    If dwState = SERVICE_RUNNING Then
                        isSelected = False
                        If IsArray(strArray) Then
                            For J = 0 To UBound(strArray)
                                If LCase(strArray(J)) = LCase(sSvcName) Then isSelected = True
                            Next
                        End If
                        If Not isSelected Then
                            Set insRet = .ListItems.Add(, , sSvcName)
                            insRet.SubItems(1) = sDispName
                        End If
                    End If
                Next
            End With
        End If
    End If
End If
Call CloseServiceHandle(hSCManager)
EnumSystemServices = dwReturned
End Function

Public Function GetStrFromPtrA(ByVal lpszA As Long) As String
   GetStrFromPtrA = String$(lstrlenA(ByVal lpszA), 0)
   Call lstrcpyA(ByVal GetStrFromPtrA, ByVal lpszA)
End Function

Sub ApplyServPref(Server As String)
Dim I           As Integer
Dim strServ     As String

strServ = frmServ.BuildServList()
Set objConn = New ADODB.Connection
objConn.Open strConn
objConn.Execute ("UPDATE tblServer SET [MonServices]='" & strServ & "' WHERE [server]='" & Server & "';")
objConn.Close
Set objConn = Nothing
For I = 1 To frmSrvmtr.G.Rows - 1
    If frmSrvmtr.G.TextMatrix(I, 0) = Server Then
        frmSrvmtr.G.TextMatrix(I, 14) = strServ
    End If
Next I

End Sub

