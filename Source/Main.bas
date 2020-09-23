Attribute VB_Name = "Main"
Option Explicit

Public Const HWND_TOPMOST = -1
Public Const INVALID_HANDLE_VALUE = -1
Public Const MAX_PATH = 260
Public Const SHGFI_DISPLAYNAME = &H200
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
Public Const WM_LBUTTONDOWN = &H201
Public Const MK_LBUTTON = &H1
Public Const SW_SHOWNORMAL = 1
Public Const SW_SHOWMINIMIZED = 2

Public strRespCode      As String
Public strFullResp      As String
Public Timeout          As Integer
Public testMail         As Boolean


Private Type TIME_OF_DAY
    t_elapsedt  As Long
    t_msecs     As Long
    t_hours     As Long
    t_mins      As Long
    t_secs      As Long
    t_hunds     As Long
    t_timezone  As Long
    t_tinterval As Long
    t_day       As Long
    t_month     As Long
    t_year      As Long
    t_weekday   As Long
End Type

Public Type SHFILEINFO
   hIcon            As Long
   iIcon            As Long
   dwAttributes     As Long
   szDisplayName    As String * MAX_PATH
   szTypeName       As String * 80
End Type

Public Type FILETIME
  dwLowDateTime     As Long
  dwHighDateTime    As Long
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Public Declare Function GetVersionEx Lib "kernel32" _
                            Alias "GetVersionExA" _
                            (lpVersionInformation As OSVERSIONINFO) As Long

Declare Function ReleaseCapture Lib "user32" () As Long

Declare Function SendMessage Lib "user32" _
                            Alias "SendMessageA" _
                            (ByVal hwnd As Long, _
                            ByVal wMsg As Long, _
                            ByVal wParam As Long, _
                            lParam As Any) As Long

Public Declare Function SetWindowPos& Lib "user32" _
                            (ByVal hwnd As Long, _
                            ByVal hwndInsertAfter As Long, _
                            ByVal X As Long, _
                            ByVal Y As Long, _
                            ByVal cx As Long, _
                            ByVal cy As Long, _
                            ByVal wFlags As Long)
                            
Public Declare Function SHGetFileInfo Lib "SHELL32" _
                            Alias "SHGetFileInfoA" _
                            (ByVal pszPath As Any, _
                            ByVal dwFileAttributes As Long, _
                            psfi As SHFILEINFO, _
                            ByVal cbFileInfo As Long, _
                            ByVal uFlags As Long) As Long
                            
Public Declare Function ShellExecute Lib "shell32.dll" _
                            Alias "ShellExecuteA" (ByVal hwnd As Long, _
                            ByVal lpOperation As String, _
                            ByVal lpFile As String, _
                            ByVal lpParameters As String, _
                            ByVal lpDirectory As String, _
                            ByVal nShowCmd As Long) As Long
                            
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" _
                            Alias "GetDiskFreeSpaceExA" _
                            (ByVal lpRootPathName As String, _
                            lpFreeBytesAvailableToCaller As Currency, _
                            lpTotalNumberOfBytes As Currency, _
                            lpTotalNumberOfFreeBytes As Currency) As Long
                            
Private Declare Function NetRemoteTOD Lib "netapi32.dll" _
                            (ByVal Server As String, _
                            buffer As Any) As Long

Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" _
                            (pTo As Any, _
                            uFrom As Any, _
                            ByVal lSize As Long)

Private Declare Function NetApiBufferFree Lib "netapi32.dll" _
                            (ByVal Ptr As Long) As Long
Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long

                            
'==== Services ===============

Public Type SERVICE_STATUS
   dwServiceType As Long
   dwCurrentState As Long
   dwControlsAccepted As Long
   dwWin32ExitCode As Long
   dwServiceSpecificExitCode As Long
   dwCheckPoint As Long
   dwWaitHint As Long
End Type

Public Type ENUM_SERVICE_STATUS
    lpServiceName       As Long
    lpDisplayName       As Long
    ServiceStatus       As SERVICE_STATUS
End Type

Public Const SIZEOF_SERVICE_STATUS As Long = 36

Public Const ERROR_MORE_DATA = 234
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Public Const LB_SETTABSTOPS As Long = &H192
Public Const SERVICE_STATE_ALL = &H3
                                     

Public Const SERVICE_KERNEL_DRIVER As Long = &H1
Public Const SERVICE_FILE_SYSTEM_DRIVER As Long = &H2
Public Const SERVICE_ADAPTER As Long = &H4
Public Const SERVICE_RECOGNIZER_DRIVER As Long = &H8
Public Const SERVICE_WIN32_OWN_PROCESS As Long = &H10
Public Const SERVICE_WIN32_SHARE_PROCESS As Long = &H20
Public Const SERVICE_INTERACTIVE_PROCESS As Long = &H100

Public Const SERVICE_WIN32 As Long = SERVICE_WIN32_OWN_PROCESS Or _
                                     SERVICE_WIN32_SHARE_PROCESS
                                     
Public Const SERVICE_DRIVER As Long = SERVICE_KERNEL_DRIVER Or _
                                      SERVICE_FILE_SYSTEM_DRIVER Or _
                                      SERVICE_RECOGNIZER_DRIVER
                                      
Public Const SERVICE_TYPE_ALL As Long = SERVICE_WIN32 Or _
                                        SERVICE_ADAPTER Or _
                                        SERVICE_DRIVER Or _
                                        SERVICE_INTERACTIVE_PROCESS
                                     
Public Const SERVICE_STOPPED As Long = &H1
Public Const SERVICE_START_PENDING As Long = &H2
Public Const SERVICE_STOP_PENDING As Long = &H3
Public Const SERVICE_RUNNING As Long = &H4
Public Const SERVICE_CONTINUE_PENDING As Long = &H5
Public Const SERVICE_PAUSE_PENDING As Long = &H6
Public Const SERVICE_PAUSED As Long = &H7
Public Const SC_MANAGER_CONNECT = &H1
Public Const GENERIC_READ = &H80000000

Public Declare Function OpenSCManager Lib "advapi32" _
   Alias "OpenSCManagerA" _
  (ByVal lpMachineName As String, _
   ByVal lpDatabaseName As String, _
   ByVal dwDesiredAccess As Long) As Long

Public Declare Function EnumServicesStatus Lib "advapi32" _
   Alias "EnumServicesStatusA" _
  (ByVal hSCManager As Long, _
   ByVal dwServiceType As Long, _
   ByVal dwServiceState As Long, _
   lpServices As Any, _
   ByVal cbBufSize As Long, _
   pcbBytesNeeded As Long, _
   lpServicesReturned As Long, _
   lpResumeHandle As Long) As Long
   
Public Declare Function QueryServiceStatus Lib "advapi32.dll" _
      (ByVal hService As Long, _
      lpServiceStatus As SERVICE_STATUS) As Long
   
Public Declare Function OpenService Lib "advapi32.dll" _
      Alias "OpenServiceA" _
      (ByVal hSCManager As Long, _
      ByVal lpServiceName As String, _
      ByVal dwDesiredAccess As Long) As Long
      
Private Declare Function QueryServiceConfig Lib "advapi32.dll" _
      Alias "QueryServiceConfigA" _
      (ByVal hService As Long, _
      lpServiceConfig As Byte, _
      ByVal cbBufSize As Long, _
      pcbBytesNeeded As Long) As Long
      
Public Declare Function CloseServiceHandle Lib "advapi32" _
   (ByVal hSCObject As Long) As Long

Public Declare Function lstrcpyA Lib "kernel32" _
  (ByVal RetVal As String, ByVal Ptr As Long) As Long
                        
Public Declare Function lstrlenA Lib "kernel32" _
  (ByVal Ptr As Any) As Long

                            
'==== Services ===============
                            
Public Const db = "monitor.mdb"
Public Const strConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=.\" & db & ";"
Public objConn As ADODB.Connection
                            

Public Function FileExists(sFile As String) As Boolean
    Dim shfi As SHFILEINFO
    If SHGetFileInfo(ByVal sFile, 0&, shfi, Len(shfi), SHGFI_DISPLAYNAME) Then
        FileExists = True
        Else
        FileExists = False
    End If
End Function
Public Sub StatusOK(I As Integer)
Dim C As Integer
                    'Draws status columns whenever new server is added.
With frmSrvmtr.G
    For C = 4 To 7
        .Col = C
        .Row = I
        Set .CellPicture = frmSrvmtr.IL.ListImages.Item(17).Picture
        .CellPictureAlignment = 4
        .Text = Empty
        .CellBackColor = &HF0F0F0
    Next C
End With
End Sub

Public Sub BuiltTable()
Dim I As Integer
                   'Draws and formats table on Main Form_Load

With frmSrvmtr.G
    .Clear
    .Rows = 2
    .CellAlignment = flexAlignCenterCenter
    .TextMatrix(0, 0) = "Server Name:"
    .ColWidth(0) = 1800
    .TextMatrix(0, 1) = "Drive:"
    .ColWidth(1) = 700
    .TextMatrix(0, 2) = "Available Space (Kb):"
    .ColWidth(2) = 1900
    .TextMatrix(0, 3) = "Total Space (Kb):"
    .ColWidth(3) = 1830
    .TextMatrix(0, 9) = "Comments:"
    .ColWidth(9) = 4000
    For I = 10 To 14
        .ColWidth(I) = 0
    Next I
    For I = 4 To 8
        .Row = 0
        .Col = I
        Set .CellPicture = frmSrvmtr.IL.ListImages.Item(I + 5).Picture
        .CellPictureAlignment = 4
        .ColWidth(I) = 330
        .CellAlignment = flexAlignCenterCenter
        .Row = 1
    Next I
    .Row = 1: .Col = 0
    .CellAlignment = flexAlignLeftCenter
End With
End Sub
Public Sub RunApplication(action As String, Optional Server As String)
Dim OSInf       As OSVERSIONINFO
Dim R           As Integer
Dim openPath    As String
Dim openDrive   As String
Dim openServer  As String
Dim OS          As Long
Dim pingStatus  As String

On Error GoTo ErrorHandler
                            'This sub runs different NT tools selected from menu.
With frmSrvmtr
    R = .G.Row
    openDrive = .G.TextMatrix(R, 1)
    openServer = "\\" & .G.TextMatrix(R, 0)
    openPath = openServer & "\" & openDrive
    
    pingStatus = PingServer(.G.TextMatrix(R, 0))
    If Not Left(pingStatus, 1) = 0 Then MsgBox "Server is down!", vbExclamation, "Monitor": Exit Sub
    
    Select Case action
        Case "cmd"
            Call ShellExecute(.hwnd, "open", Environ$("COMSPEC"), _
            "/c start " & Chr(34) & "Server Monitor Command Prompt. " & openPath & Chr(34) & " " & _
            Environ$("COMSPEC") & " /t:1E /k pushd" & openPath, vbNullString, SW_SHOWNORMAL)
        Case "rcmd"
            Call ShellExecute(.hwnd, "open", "rcmd", openServer, vbNullString, SW_SHOWNORMAL)
        Case "srvmgr"
            Call ShellExecute(.hwnd, "open", "compmgmt.msc", " /computer:" & openServer, vbNullString, SW_SHOWNORMAL)
        Case "services"
            Call ShellExecute(.hwnd, "open", "services.msc", " /computer:" & openServer, vbNullString, SW_SHOWNORMAL)
        Case "winmsd"
            OSInf.dwOSVersionInfoSize = Len(OSInf)
            GetVersionEx OSInf
            OS = OSInf.dwMajorVersion
            If OS > 4 Then openServer = "/computer " & openServer
            Call ShellExecute(.hwnd, "open", "winmsd", openServer, vbNullString, SW_SHOWNORMAL)
        Case "eventvwr"
            Call ShellExecute(.hwnd, "open", "eventvwr", openServer, vbNullString, SW_SHOWNORMAL)
        Case "EventLog"
            Call ShellExecute(.hwnd, "open", "eventvwr", Server, vbNullString, SW_SHOWNORMAL)
        Case "open"
            Call ShellExecute(.hwnd, "open", openPath, vbNullString, vbNullString, SW_SHOWNORMAL)
        Case "openLog"
            Call ShellExecute(.hwnd, "open", "notepad", App.Path & "\SMlog.log", vbNullString, SW_SHOWNORMAL)
        Case "Help"
            Call ShellExecute(.hwnd, "open", App.Path & "\monitor.chm", vbNullString, vbNullString, SW_SHOWNORMAL)
    End Select
End With
Exit Sub
ErrorHandler:
MsgBox "Error opening ''" & action & "'' on Server: ''" & openServer & "''. Error Number: " & Err.Number & _
        " Error Description: " & Err.Description

End Sub

Public Sub SaveList()
Dim Cancel          As Boolean
Dim FileToSave      As String
Dim tableData       As String

'Saves table in a text file or calls "SaveAsExcel" depending on the extention

On Error GoTo ErrorHandler
Cancel = False
With frmSrvmtr.rDialog
    .DialogTitle = "Save List as..."
    .CancelError = True
    .Filter = "Microsoft Excel Workbook (*.xls)|*.xls|Text (Tab Delimited) (*.txt)|*.txt"
    .FileName = "Monitor-List"
    .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt
    .ShowSave
    FileToSave = .FileName
End With

If Not Cancel Then
    If Right(FileToSave, 3) = "xls" Then
        If SaveAsExcel(FileToSave) = True Then
            Exit Sub
        Else: FileToSave = Mid(FileToSave, 1, Len(FileToSave) - 3) & "txt"
        End If
    End If
    With frmSrvmtr.G
        .Row = 0
        .Col = 0
        .ColSel = 8
        .RowSel = .Rows - 1
        tableData = .Clip               'You'll need SP3 and up for Visual Studio for .Clip to work
        Open FileToSave For Output As #3
        tableData = Replace(tableData, vbCr, vbCrLf)
        Print #3, tableData
        Close #3
        .Row = 1: .Col = 1
    End With
End If
Exit Sub

ErrorHandler:
If Err.Number = cdlCancel Then
    Cancel = True
    Resume Next
End If
tableData = ""
End Sub

Public Function SaveAsExcel(saveExcel As String) As Boolean

Dim xlApp       As Excel.Application
Dim C           As Integer
Dim R           As Integer
Dim S           As Series
Dim nextCell    As String
Dim selRange    As Object
Dim XlChart     As Object


frmSrvmtr.Refresh
Set xlApp = New Excel.Application
If Err.Number <> 0 Then GoTo NoExcel
                            
'Excel routine to save file. If Error occured, function returns false to "SaveList" sub
'and file could be saved as a text file.

On Error GoTo ErrorHandler
With xlApp

    .Visible = False
    .Workbooks.Add
    .ActiveWorkbook.Colors(5) = &HCAC7A6
    .ActiveWorkbook.Colors(21) = &HE0E0E0
    .ActiveWorkbook.Colors(26) = &H82D57B
    .Application.DisplayAlerts = False
    .ActiveWorkbook.Worksheets.Add.Name = "Server Monitor Table"

    For C = 0 To 8
        .Worksheets(1).Columns(C + 1).ColumnWidth = 15
        For R = 0 To frmSrvmtr.G.Rows - 1
            nextCell = frmSrvmtr.G.TextMatrix(R, C)
            .Worksheets(1).Cells(R + 1, C + 1).Value = nextCell
            frmSrvmtr.G.Col = C
            frmSrvmtr.G.Row = R
            .Worksheets(1).Cells(R + 1, C + 1).Interior.Color = frmSrvmtr.G.CellBackColor
        Next R
    Next C

    .Worksheets(1).Columns(1).ColumnWidth = 20
    .Worksheets(1).Columns(2).ColumnWidth = 6
    .Worksheets(1).Columns(3).ColumnWidth = 20
    .Worksheets(1).Columns(4).ColumnWidth = 20
    For C = 5 To 8
        .Worksheets(1).Columns(C).ColumnWidth = 6
    Next C
    .Worksheets(1).Columns(9).ColumnWidth = 30
    Set selRange = .ActiveSheet.UsedRange
    selRange.Select
    .Selection.Borders.Weight = xlThin
    .Selection.Font.Bold = True
    
' select correct range and create a chart

    R = .ActiveSheet.UsedRange.Rows.Count
    .Range(.Cells(1, 1), .Cells(R, 4)).Select
    Set XlChart = .Charts.Add
    With XlChart
        .Name = "Space Chart"
        .Type = -4100
        .Move , xlApp.Sheets(2)
        .HasTitle = True
        .ChartTitle.Characters.Text = "Drives Space Chart (Generated by Server Monitor)."
        .ChartArea.Font.Size = 10
        .ChartArea.Font.Color = vbRed
        .ChartArea.Font.Bold = True
        .ChartArea.Fill.ForeColor.SchemeColor = 37
        .PlotArea.Fill.ForeColor.SchemeColor = 44
        .PlotArea.Fill.TwoColorGradient Style:=msoGradientHorizontal, Variant:=1
        .Walls.Fill.ForeColor.SchemeColor = 37
        .Walls.Fill.TwoColorGradient Style:=msoGradientHorizontal, Variant:=1
        .Floor.Fill.ForeColor.SchemeColor = 11
        .Floor.Fill.TwoColorGradient Style:=msoGradientHorizontal, Variant:=1
        For Each S In XlChart.SeriesCollection
            S.ApplyDataLabels Type:=xlDataLabelsShowValue
            S.DataLabels.Font.Size = 6
            S.DataLabels.Font.Color = xlApp.ActiveWorkbook.Colors(27)
            S.DataLabels.Interior.Color = vbRed
            S.DataLabels.Fill.TwoColorGradient Style:=msoGradientHorizontal, Variant:=1
        Next
        With .SeriesCollection(1).Fill
            .ForeColor.SchemeColor = 32
            .BackColor.SchemeColor = 37
            .TwoColorGradient Style:=msoGradientHorizontal, Variant:=1
        End With
        With .SeriesCollection(2).Fill
            .ForeColor.SchemeColor = 50
            .BackColor.SchemeColor = 43
            .TwoColorGradient Style:=msoGradientHorizontal, Variant:=1
        End With
    End With


    .Sheets(1).Activate
    .Range("A1").Select
    .ActiveWorkbook.SaveAs saveExcel
    .ActiveWorkbook.Close (False)
    .Application.Quit
End With

Set xlApp = Nothing
Set selRange = Nothing
Set XlChart = Nothing
SaveAsExcel = True
frmSrvmtr.G.Row = 1: frmSrvmtr.G.Col = 1
Exit Function

ErrorHandler:
If MsgBox("Error Saving In Excel: " & Err.Description & ". Error Number: " & Err.Number & _
    ". Would you like to save it as a ''Tab Delimited'' text file?", vbCritical + vbYesNo, "Error:") = vbYes Then
    Err.Clear
    SaveAsExcel = False
Else: SaveAsExcel = True
End If

xlApp.ActiveWorkbook.Close (False)
xlApp.Application.Quit
Set xlApp = Nothing
Set selRange = Nothing
Set XlChart = Nothing

Exit Function

NoExcel:
If MsgBox("There is no MS Excel application installed on this machine! " & _
    "Would you like to save it as a ''Tab Delimited'' text file?", vbCritical + vbYesNo, "Error:") = vbYes Then
    Err.Clear
    SaveAsExcel = False
Else: SaveAsExcel = True
End If
End Function
Public Function PingServer(Server As String) As String
Dim ipResolve   As String
Dim pingRet     As Long
Dim ECHO        As ICMP_ECHO_REPLY
                                        'Used to ping servers
    Call SocketsInitialize
    ipResolve = GetIPFromHostName(Server)
    If Not ipResolve = Empty Then
        pingRet = Ping(ipResolve, "send this", ECHO)
        If pingRet = 0 Then
            PingServer = 0 & ipResolve
        Else: PingServer = 1 & ipResolve
        End If
    Else: PingServer = 2
    End If
End Function
Public Sub QueryServers(selectedRow As Integer, _
                        singleServer As Boolean, _
                        timerUpdate As Boolean, _
                        isEMail As Boolean, _
                        isPage As Boolean)

Dim R As Integer, C As Integer, totRows As Integer, interval As Integer, _
    setLogTime As Integer, T As Integer
Dim totalSpace As Currency, freeSpace As Currency, freeAvail As Currency
Dim Server As String, Drive As String, pingStatus As String, errMessage As String, rPath As String, _
    strMailMessage As String, strPageMessage As String
Dim serverErr   As Boolean, sysLogErrFlag As Boolean, appLogErrFlag As Boolean, _
    sendEMail As Boolean, sendPage As Boolean, notify As Boolean, mailFlag As Boolean, mUpdate As Boolean
Dim errAdd              As ListItem
Dim minSpace            As Double
Dim remoteTime          As Date
Dim servArray           As Variant
Dim isServiceStopped    As Boolean
Dim serviceNames        As String


frmSrvmtr.MousePointer = vbHourglass
frmSrvmtr.Refresh
With frmSrvmtr.G
    If timerUpdate And (isEMail Or isPage) Then notify = True
    If Not timerUpdate Then mUpdate = True
    frmSrvmtr.sBar.Panels(2).Text = "Query Servers Availability and Status..."
    minSpace = Trim(GetConfigOptions.setMinDriveSpace)
    interval = Trim(GetConfigOptions.setInterval)
    
'Check if single server refresh status requested.

    If singleServer Then
        totRows = selectedRow
    Else: totRows = .Rows - 1
    End If
    
'***************** Start query servers ******************

    For R = selectedRow To totRows
        .Row = R
        Server = .TextMatrix(R, 0): Drive = .TextMatrix(R, 1)
        frmSrvmtr.sBar.Panels(1).Text = Server
        
'Check if Server is up.
       
        pingStatus = PingServer(Server)
        If Not Left(pingStatus, 1) = 0 Then
        
'Server is Down. Paint "Status" columns.

            .Col = 4
            Set .CellPicture = frmSrvmtr.IL.ListImages.Item(18).Picture
            .CellPictureAlignment = 4
            .CellBackColor = &H0&
            .Text = ""
            For C = 5 To 8
                .Col = C
                .Text = "n|a"
                .CellBackColor = &HFF&
                Set .CellPicture = Nothing
            Next C
            
'Add server and error message to the Log Status Window depending on the ping return message.

            Set errAdd = frmSrvmtr.statusLW.ListItems.Add(, , Server, , "Down")
            errAdd.SubItems(1) = Now
            Select Case Left(pingStatus, 1)
                Case 1
                    errMessage = "Server can not be reached. The name was resolved to IP Address: " & Mid(pingStatus, 2)
                    errAdd.SubItems(2) = errMessage
                Case 2
                    errMessage = "Server can not be reached. Unable to resolve name to IP Address."
                    errAdd.SubItems(2) = errMessage
            End Select
            
'Write error to the log file

            Call LogErrors(Server, "All", errMessage)
            
        'I do not want to send message for the same server twice
        'if it's listed with different drives. So...
        
        If Not singleServer And R > 1 Then             'if server is at list second in the table
            For C = 1 To R - 1                                  'then check past rows for the match and
                If .TextMatrix(C, 0) = Server Then              'exit if match found.
                    errAdd.SubItems(3) = "See same server above."
                    GoTo skip
                End If
            Next C
        End If

'See if notification is applied to this server and add message variables to Global Array.
            If notify Then
                If .TextMatrix(R, 10) = 0 And isEMail Then
                    sendEMail = True
                    strMailMessage = strMailMessage & vbCrLf & "Server: " & Server & _
                                     vbCrLf & "Drive: " & "ANY" & _
                                     vbCrLf & "Error: " & errMessage
                Else: sendEMail = False
                End If
                If .TextMatrix(R, 11) = 0 And isPage Then
                    sendPage = True
                    strPageMessage = strPageMessage & vbCrLf & "Server: " & Server & _
                                     vbCrLf & "Drive: " & "ANY" & _
                                     vbCrLf & "Error: " & errMessage
                Else: sendPage = False
                End If
                If sendEMail Or sendPage Then
                    errAdd.SubItems(3) = "Mail Send Pending..."
                    mailFlag = True
                Else: errAdd.SubItems(3) = "Notification does not apply."
                End If
            Else
                errAdd.SubItems(3) = "Notification turned Off."
                sendEMail = False: sendPage = False
            End If
            If mUpdate Then errAdd.SubItems(3) = "Not Sent - Manual Update."
                
'Since server is not pingable there is nothing else I can check for it.
'Raise "serverErr" flag for the status bar update and go to the next server.

            serverErr = True: GoTo skip
        Else

'If you got here, then server is OK so far.
'Repaint "Status" cells just in case it was down at the last update.

            For C = 4 To 8
                .Col = C
                 Set .CellPicture = frmSrvmtr.IL.ListImages.Item(17).Picture
                .CellPictureAlignment = 4
                .Text = Empty
                .CellBackColor = &HF0F0F0
            Next C
        End If
        
'Now lets see if Server's Drive is available.
'If drive is not available start the same routine as with ping.

        rPath = "\\" & Server & "\" & Drive
        If Not FileExists(rPath) Then
        
            .Col = 5
            Set .CellPicture = frmSrvmtr.IL.ListImages.Item(18).Picture
            .CellPictureAlignment = 4
            .CellBackColor = &H0&
            .Text = ""
            .Col = 6
            .Text = "n|a"
            .CellBackColor = &HFF&
            Set .CellPicture = Nothing
           
            
            Set errAdd = frmSrvmtr.statusLW.ListItems.Add(, , Server, , "Freez")
            errAdd.SubItems(1) = Now
            errMessage = "Drive " & Chr(34) & Drive & Chr(34) & _
                         " is not available on the server or server hang."
            errAdd.SubItems(2) = errMessage
            
            Call LogErrors(Server, Drive, errMessage)
            
            If notify Then
                If .TextMatrix(R, 10) = 0 And isEMail Then
                    sendEMail = True
                    strMailMessage = strMailMessage & vbCrLf & "Server: " & Server & _
                                     vbCrLf & "Drive: " & Drive & _
                                     vbCrLf & "Error: " & errMessage
                Else: sendEMail = False
                End If
                If .TextMatrix(R, 11) = 0 And isPage Then
                    sendPage = True
                    strPageMessage = strPageMessage & vbCrLf & "Server: " & Server & _
                                     vbCrLf & "Drive: " & Drive & _
                                     vbCrLf & "Error: " & errMessage
                Else: sendPage = False
                End If
                If sendEMail Or sendPage Then
                    errAdd.SubItems(3) = "Mail Send Pending..."
                    mailFlag = True
                Else: errAdd.SubItems(3) = "Notification does not apply."
                End If
            Else
                errAdd.SubItems(3) = "Notification turned Off."
                sendEMail = False: sendPage = False
            End If
            If mUpdate Then errAdd.SubItems(3) = "Not Sent - Manual Update."
            serverErr = True: GoTo skip
        Else
            For C = 5 To 8
                .Col = C
                Set .CellPicture = frmSrvmtr.IL.ListImages.Item(17).Picture
                .CellPictureAlignment = 4
                .Text = Empty
                .CellBackColor = &HF0F0F0
            Next C
        End If

'Check drive's free space. If less that minimum allowed then add entry
'to the Log file and to the Log Status Window.

        frmSrvmtr.Refresh
        totalSpace = 0: freeSpace = 0
        Call GetDiskFreeSpaceEx(rPath, freeAvail, totalSpace, freeSpace)
            .TextMatrix(R, 3) = Format$(totalSpace * 10, "###,###,###,##0")
            .TextMatrix(R, 2) = Format$(freeSpace * 10, "###,###,###,##0")
            
            If Format$(freeSpace * 10, "###,###,###,##0") < minSpace Then
            
                .Col = 6
                Set .CellPicture = frmSrvmtr.IL.ListImages.Item(18).Picture
                .CellPictureAlignment = 4
                .CellBackColor = &H0&
                .Text = ""
                
                Set errAdd = frmSrvmtr.statusLW.ListItems.Add(, , Server, , "SpaceLow")
                errAdd.SubItems(1) = Now
                errMessage = "Available free space on Drive " & Chr(34) & Drive & Chr(34) & _
                         " is: " & Format$(freeSpace * 10, "###,###,###,##0") & " Kb." & _
                         ", which is below min. allowed of: " & Format$(minSpace, "###,###,###,##0") & " Kb."
                errAdd.SubItems(2) = errMessage
            
                Call LogErrors(Server, Drive, errMessage)
            
                If notify Then
                    If .TextMatrix(R, 10) = 0 And isEMail Then
                        sendEMail = True
                        strMailMessage = strMailMessage & vbCrLf & "Server: " & Server & _
                                         vbCrLf & "Drive: " & Drive & _
                                         vbCrLf & "Error: " & errMessage
                    Else: sendEMail = False
                    End If
                    If .TextMatrix(R, 11) = 0 And isPage Then
                        sendPage = True
                        strPageMessage = strPageMessage & vbCrLf & "Server: " & Server & _
                                         vbCrLf & "Drive: " & Drive & _
                                         vbCrLf & "Error: " & errMessage
                    Else: sendPage = False
                    End If
                    If sendEMail Or sendPage Then
                        errAdd.SubItems(3) = "Mail Send Pending..."
                        mailFlag = True
                    Else: errAdd.SubItems(3) = "Notification does not apply."
                    End If
                Else
                    errAdd.SubItems(3) = "Notification turned Off."
                    sendEMail = False: sendPage = False
                End If
                If mUpdate Then errAdd.SubItems(3) = "Not Sent - Manual Update."
                serverErr = True
            End If
            
            
'Check Event Log for the error type events and monitored services for status.

        'I do not want to access event log for the same server twice
        'if it's listed with different drives. So...
        If Not singleServer And R > 1 Then             'if server is at list second in the table
            For C = 1 To R - 1                         'then check past rows for the match and
                If .TextMatrix(C, 0) = Server Then     'if error raised for the server above then
                    .Col = 7: .Row = C                 'draw error in status and goto next server.
                    If .CellBackColor = &H0& Then
                        .Row = R
                        Set .CellPicture = frmSrvmtr.IL.ListImages.Item(18).Picture
                        .CellPictureAlignment = 4
                        .CellBackColor = &H0&
                        .Text = ""
                    End If
                    GoTo checkServ
                End If
            Next C
        End If
        
        If timerUpdate Then         'If auto. refresh then
            setLogTime = interval   'set time filter to query interval.
        Else: setLogTime = 10       'If manual refresh then pull events for the past 10 min.
        End If
        
        remoteTime = GetServerTime(Server)  'Pulls remote time because event log filter needs to be set
                                            'according to the remote server time. This will ensure
                                            'correct time filter even if server is in a different time zone.
        T = frmSrvmtr.statusLW.ListItems.Count
        If CBool(GetConfigOptions.setEnableEventLog) Then
            If CBool(GetConfigOptions.setEnableSLog) And _
                .TextMatrix(R, 12) = 0 Then
                sysLogErrFlag = ReadEventLog(Server, "Any", "System", _
                                            remoteTime - (setLogTime / 24 / 60))
            End If
            If CBool(GetConfigOptions.setEnableALog) And _
                .TextMatrix(R, 13) = 0 Then
                appLogErrFlag = ReadEventLog(Server, "Any", "Application", _
                                            remoteTime - (setLogTime / 24 / 60))
            End If
            If sysLogErrFlag Or appLogErrFlag Then
                    serverErr = True
                    errMessage = "Server received one or more Error Type Log Events. Please check Event Log"
                    .Col = 7
                    Set .CellPicture = frmSrvmtr.IL.ListImages.Item(18).Picture
                    .CellPictureAlignment = 4
                    .CellBackColor = &H0&
                    .Text = ""

                If notify Then
                    If .TextMatrix(R, 10) = 0 And isEMail Then
                        sendEMail = True
                        strMailMessage = strMailMessage & vbCrLf & "Server: " & Server & _
                                         vbCrLf & "Drive: " & "ANY" & _
                                         vbCrLf & "Error: " & errMessage
                    Else: sendEMail = False
                    End If
                    If .TextMatrix(R, 11) = 0 And isPage Then
                        sendPage = True
                        strPageMessage = strPageMessage & vbCrLf & "Server: " & Server & _
                                         vbCrLf & "Drive: " & "ANY" & _
                                         vbCrLf & "Error: " & errMessage
                    Else: sendPage = False
                    End If
                    
                    If sendEMail Or sendPage Then
                        Call UpdateLogEventsMailStatus(T, "Mail Send Pending...")
                        mailFlag = True
                    Else: Call UpdateLogEventsMailStatus(T, "Notification does not apply.")
                    End If
                Else:
                    sendEMail = False: sendPage = False
                    Call UpdateLogEventsMailStatus(T, "Notification turned Off.")
                End If
                If mUpdate Then Call UpdateLogEventsMailStatus(T, "Not Sent - Manual Update.")
            End If
        End If
checkServ:
        Server = Replace(Server, "\", "")
        
        If Not singleServer And R > 1 Then             'if server is at list second in the table
            For C = 1 To R - 1                         'then check past rows for the match and
                If .TextMatrix(C, 0) = Server Then     'if error raised for the server above then
                    .Col = 8: .Row = C                 'draw error in status and goto next server.
                    If .CellBackColor = &H0& Then
                        .Row = R
                        Set .CellPicture = frmSrvmtr.IL.ListImages.Item(18).Picture
                        .CellPictureAlignment = 4
                        .CellBackColor = &H0&
                        .Text = ""
                    End If
                    GoTo skip
                End If
            Next C
        End If
        
        isServiceStopped = False
        serviceNames = Empty
        If Len(.TextMatrix(R, 14)) > 0 Then
            servArray = Split(.TextMatrix(R, 14), "|")
            For C = 0 To UBound(servArray)
                If Not QueryService(Server, servArray(C)) Then
                   isServiceStopped = True
                   serviceNames = serviceNames & " " & servArray(C)
                End If
            Next
            serviceNames = "[" & Trim(serviceNames) & "]"
            If isServiceStopped Then

                .Col = 8
                Set .CellPicture = frmSrvmtr.IL.ListImages.Item(18).Picture
                .CellPictureAlignment = 4
                .CellBackColor = &H0&
                .Text = ""

                Set errAdd = frmSrvmtr.statusLW.ListItems.Add(, , Server, , "service")
                errAdd.SubItems(1) = Now
                errMessage = "Service " & serviceNames & " has stopped."
                errAdd.SubItems(2) = errMessage

                Call LogErrors(Server, "ANY", errMessage)
                If notify Then
                    errMessage = "Service " & serviceNames & " has stopped."
                    If .TextMatrix(R, 10) = 0 And isEMail Then
                        sendEMail = True
                        strMailMessage = strMailMessage & vbCrLf & "Server: " & Server & _
                                         vbCrLf & "Drive: " & Drive & _
                                         vbCrLf & "Error: " & errMessage
                    Else: sendEMail = False
                    End If
                    If .TextMatrix(R, 11) = 0 And isPage Then
                        sendPage = True
                        strPageMessage = strPageMessage & vbCrLf & "Server: " & Server & _
                                         vbCrLf & "Drive: " & Drive & _
                                         vbCrLf & "Error: " & errMessage
                    Else: sendPage = False
                    End If
                    If sendEMail Or sendPage Then
                        errAdd.SubItems(3) = "Mail Send Pending..."
                        mailFlag = True
                    Else: errAdd.SubItems(3) = "Notification does not apply."
                    End If
                Else
                    errAdd.SubItems(3) = "Notification turned Off."
                    sendEMail = False: sendPage = False
                End If
                If mUpdate Then errAdd.SubItems(3) = "Not Sent - Manual Update."
                serverErr = True
            End If
        End If
skip:
    Next R

'****************** End Query Servers ******************

'Check if mail flag is raised and pass Email or Page message to sendMail sub.

If mailFlag Then
    frmSrvmtr.sBar.Panels(1).Text = ""
    frmSrvmtr.sBar.Panels(2).Text = "Attempt to send Email and/or Page..."
    If Not strMailMessage = "" Then
        If SendMail(strMailMessage, "Email") = True Then
'Change status from pending to sent
            Call UpdateEmailStatus("Email/Page - Success")
        End If
    End If
    If Not strPageMessage = "" Then
        If SendMail(strPageMessage, "Page") = True Then
            Call UpdateEmailStatus("Email/Page - Success")
        End If
    End If
    
End If


'Depending on refresh request type (either manual or automatic),
'timer will start or reset.


    If timerUpdate Then
        nextUpdate = Now + (interval / 60 / 24) 'reset
    Else: nextUpdate = Now + timeDiff + (2 / 60 / 24 / 60) 'pick up time it was stopped at (well... 2 sec. ahead.)
    End If
    frmSrvmtr.T1.Enabled = True

'Display update result at the status bar.

    If serverErr Then
        frmSrvmtr.sBar.Panels(2).Text = "One or more Error(s) were detected during last update."
    Else: frmSrvmtr.sBar.Panels(2).Text = "No Errors were detected during last update."
    End If
    frmSrvmtr.sBar.Panels(1).Text = ""
    frmSrvmtr.MousePointer = vbNormal
End With

End Sub
Public Sub UpdateLogEventsMailStatus(I As Integer, strMessage As String)
Dim R As Integer

With frmSrvmtr.statusLW

    For R = I + 1 To .ListItems.Count
        .ListItems.Item(R).SubItems(3) = strMessage
    Next
        
End With
End Sub
Public Sub UpdateEmailStatus(strMessage As String)
Dim R               As Integer

If testMail Then Exit Sub

On Error Resume Next
With frmSrvmtr.statusLW

    For R = 1 To .ListItems.Count
        If .ListItems.Item(R).SubItems(3) = "Mail Send Pending..." Then
            .ListItems.Item(R).SubItems(3) = strMessage
        End If
    Next
    
End With
End Sub
Public Function SendMail(strMessage As String, action As String) As Boolean
Dim objOutlook      As Object
Dim objMail         As Object
Dim sendTO          As Variant
Dim R               As Integer
Dim I               As Integer
Dim smtpHost        As String
Dim colonPos        As Integer
Dim smtpPort        As Long
Dim rec             As Integer
Dim isAuthenticate  As Boolean

On Error GoTo ErrorHandler

'Select whether Email or Page.

Select Case action
    Case "Email"
    sendTO = Trim(GetConfigOptions.setEmailAddress)
    Case "Page"
    sendTO = Trim(GetConfigOptions.setPageAddress)
End Select


'Select whether use Outlook or SMTP for Email and Page.

Select Case Abs(GetConfigOptions.setEnableSMTP)
    Case 0
'Outlook Obviously
'Easy one
        Set objOutlook = CreateObject("Outlook.Application")
        Set objMail = objOutlook.CreateItem(0)

        With objMail
            .Body = strMessage
            .Subject = "Error Notification form Server Monitor."
            .To = sendTO
            .send
        End With
    Case 1
'SMTP
        sendTO = Split(sendTO, ";")
        With frmSrvmtr
            .Winsock1.Protocol = sckTCPProtocol
            smtpHost = Trim(GetConfigOptions.setSMTPHost)
            colonPos = InStr(smtpHost, ":")
            If colonPos = 0 Then
                .Winsock1.Connect smtpHost, 25
            Else
                smtpPort = CLng(Right$(smtpHost, Len(smtpHost) - colonPos))
                smtpHost = Left$(smtpHost, colonPos - 1)
                .Winsock1.Connect smtpHost, smtpPort
            End If
            If Len(GetConfigOptions.setSMTPHostID) > 0 Then isAuthenticate = True

'Wait for server response.
            If CaptureResponse("220") = False Then SendMail = False: Exit Function
'Say "HELO/EHLO" to server and get response.
            strRespCode = ""
            If isAuthenticate Then
                .Winsock1.SendData "EHLO " & GetConfigOptions.setSMTPHostID & vbCrLf
            Else
                .Winsock1.SendData "HELO " & smtpHost & vbCrLf
            End If
            If CaptureResponse("250") = False Then SendMail = False: Exit Function
'Authenticate
            strRespCode = ""
            If isAuthenticate Then
                .Winsock1.SendData "AUTH LOGIN" & vbCrLf
                If CaptureResponse("334") = False Then SendMail = False: Exit Function
                strRespCode = ""
                .Winsock1.SendData GetConfigOptions.setSMTPHostID & vbCrLf
                If CaptureResponse("334") = False Then SendMail = False: Exit Function
                strRespCode = ""
                .Winsock1.SendData GetConfigOptions.setSMTPHostPass & vbCrLf
                If CaptureResponse("235") = False Then SendMail = False: Exit Function
                strRespCode = ""
            End If
'Send sender Email.
            .Winsock1.SendData "MAIL FROM:" & Trim(sendTO(0)) & vbCrLf
            If CaptureResponse("250") = False Then SendMail = False: Exit Function
'Send recipient Email.
For rec = 0 To UBound(sendTO)
            strRespCode = ""
            .Winsock1.SendData "RCPT TO:" & Trim(sendTO(rec)) & vbCrLf
            If CaptureResponse("250") = False Then SendMail = False: Exit Function
Next rec
'Prepare to send Message.
            strRespCode = ""
            .Winsock1.SendData "DATA" & vbCrLf
            If CaptureResponse("354") = False Then SendMail = False: Exit Function
'Send Subject and Message
            .Winsock1.SendData "From:" & "Server Monitor" & vbCrLf
            .Winsock1.SendData "To:" & "Server Monitor User" & vbCrLf
            .Winsock1.SendData "Subject:" & "Error Notification from Server Monitor." & vbCrLf & vbCrLf
            .Winsock1.SendData strMessage & vbCrLf
            strRespCode = ""
            .Winsock1.SendData "." & vbCrLf
            If CaptureResponse("250") = False Then SendMail = False: Exit Function
            .Winsock1.SendData "QUIT" & vbCrLf
            Call CaptureResponse("221")
            .Winsock1.Close
        End With
End Select

SendMail = True
Exit Function

ErrorHandler:
SendMail = False
Call UpdateEmailStatus("Unable to Mail: " & Err.Description)
Call LogErrors("Mail/Page Error", "Unable to Send", "Error: " & _
              Err.Number & ", Source: " & Err.Source & ", Description: " & Err.Description & " Server Error: " & strFullResp)
Err.Clear
End Function
Public Function CaptureResponse(respCode As String) As Boolean
With frmSrvmtr

    Timeout = 0
    .T2.Enabled = True
    
'Loop for 5 sec. If SMTP host did not send expected response then exit.
'strRespCode is global from Winsock1_DataArrival on form.

    Do
        If Not strRespCode = "" Then
            If strRespCode = respCode Then
                CaptureResponse = True
                .T2.Enabled = False
                Exit Function
            End If
        End If
        DoEvents
        If Timeout > 50 Then
            CaptureResponse = False
            .T2.Enabled = False
            Exit Function
        End If
    Loop
    
End With

End Function
Public Sub LogErrors(Server As String, Drive As String, errMessage As String)
Dim errLog      As String
    On Error Resume Next
    errLog = App.Path & "\SMlog.log"
    Open errLog For Append As #2
    Print #2, Now & " -- " & Server & " -- " & Drive & " -- " & errMessage
    Close #2
End Sub

Public Function GetServerTime(strServerName As String) As String

Dim lngBuffer               As Long
Dim strServer               As String
Dim lngNet32ApiReturnCode   As Long
Dim serverTime              As Date
Dim zoneDiff                As Date
Dim TOD                     As TIME_OF_DAY

On Error Resume Next

strServer = StrConv(strServerName, vbUnicode)
lngNet32ApiReturnCode = NetRemoteTOD(strServer, lngBuffer)
If lngNet32ApiReturnCode = 0 Then
    CopyMem TOD, ByVal lngBuffer, Len(TOD)
    serverTime = DateSerial(70, 1, 1) + (TOD.t_elapsedt / 60 / 60 / 24)
    zoneDiff = (TOD.t_timezone / 60 / 24)
    serverTime = serverTime - zoneDiff
    GetServerTime = serverTime
End If
Call NetApiBufferFree(lngBuffer)

End Function

Public Function sq(inputstr)
    inputstr = Replace(inputstr, "\'", "'")
    sq = Replace(inputstr, "'", "''")
End Function
Public Function QueryService(Server As String, qservice As Variant) As Boolean

Dim hSCM        As Long
Dim hSVC        As Long
Dim pSTATUS     As SERVICE_STATUS
Dim lRet        As Long

hSCM = OpenSCManager(Server, vbNullString, SC_MANAGER_CONNECT)
hSVC = OpenService(hSCM, qservice, GENERIC_READ)
lRet = QueryServiceStatus(hSVC, pSTATUS)
If lRet = 0 Then GoTo CloseHandles
If pSTATUS.dwCurrentState = SERVICE_RUNNING Then QueryService = True

CloseHandles:
CloseServiceHandle (hSVC)
CloseServiceHandle (hSCM)
End Function
