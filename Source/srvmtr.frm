VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmSrvmtr 
   Caption         =   "Server Monitor - v2.0"
   ClientHeight    =   6045
   ClientLeft      =   165
   ClientTop       =   570
   ClientWidth     =   7740
   Icon            =   "srvmtr.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6045
   ScaleWidth      =   7740
   Begin VB.Frame F2 
      BorderStyle     =   0  'None
      Caption         =   "- Drive Space"
      Height          =   375
      Left            =   0
      TabIndex        =   7
      Top             =   5520
      Width           =   8010
      Begin VB.Label Label2 
         Caption         =   "- Services"
         Height          =   255
         Index           =   5
         Left            =   6690
         TabIndex        =   13
         Top             =   60
         Width           =   1245
      End
      Begin VB.Label Label2 
         Caption         =   "- Error Events"
         Height          =   285
         Index           =   4
         Left            =   5280
         TabIndex        =   12
         Top             =   60
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "- Drive Space"
         Height          =   255
         Index           =   3
         Left            =   3780
         TabIndex        =   11
         Top             =   60
         Width           =   1005
      End
      Begin VB.Label Label2 
         Caption         =   "- Drive Status"
         Height          =   255
         Index           =   2
         Left            =   2280
         TabIndex        =   10
         Top             =   60
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "- Network Status"
         Height          =   255
         Index           =   0
         Left            =   540
         TabIndex        =   8
         Top             =   60
         Width           =   1245
      End
      Begin VB.Image I1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   4
         Left            =   6360
         Top             =   30
         Width           =   315
      End
      Begin VB.Image I1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   3
         Left            =   4950
         Top             =   30
         Width           =   315
      End
      Begin VB.Image I1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   2
         Left            =   3450
         Top             =   30
         Width           =   315
      End
      Begin VB.Image I1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   1
         Left            =   1950
         Top             =   30
         Width           =   315
      End
      Begin VB.Image I1 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Index           =   0
         Left            =   210
         Top             =   30
         Width           =   315
      End
   End
   Begin MSComctlLib.Toolbar TB 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "IL"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   14
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "add"
            Description     =   "Add Server"
            Object.ToolTipText     =   "Add Server"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "remove"
            Description     =   "Remove Server"
            Object.ToolTipText     =   "Remove Server"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "clearLog"
            Description     =   "Clear Log Window"
            Object.ToolTipText     =   "Clear Log Window"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "view"
            Description     =   "View Log File"
            Object.ToolTipText     =   "View Log File"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "help"
            Description     =   "Show Help"
            Object.ToolTipText     =   "Show Help"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "exit"
            Description     =   "Exit Program"
            Object.ToolTipText     =   "Exit Program"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "refresh"
            Description     =   "Refresh Servers Status"
            Object.ToolTipText     =   "Refresh Servers List"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "stopTimer"
            Description     =   "Start/Stop Timer"
            Object.ToolTipText     =   "Used to Stop or Start Timer"
            ImageIndex      =   14
            Style           =   1
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList IL 
         Left            =   4890
         Top             =   -120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   18
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":0B3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":0F72
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":10D4
               Key             =   "remove"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":14EB
               Key             =   "viewLog"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":191A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":1D50
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":2188
               Key             =   "help"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":25BF
               Key             =   "refresh"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":29BE
               Key             =   "Down"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":2DDB
               Key             =   "Freez"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":3214
               Key             =   "SpaceLow"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":3659
               Key             =   "StopEvent"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":3A37
               Key             =   "service"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":3E56
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":4285
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":4673
               Key             =   "noAccess"
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":498D
               Key             =   "stup"
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "srvmtr.frx":4C89
               Key             =   "stdown"
            EndProperty
         EndProperty
      End
      Begin VB.Timer T2 
         Enabled         =   0   'False
         Interval        =   100
         Left            =   5490
         Top             =   0
      End
      Begin MSWinsockLib.Winsock Winsock1 
         Left            =   7350
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.Timer T1 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   5880
         Top             =   0
      End
      Begin MSComDlg.CommonDialog rDialog 
         Left            =   6870
         Top             =   -30
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.PictureBox TrayIcon 
         Height          =   225
         Left            =   330
         ScaleHeight     =   165
         ScaleWidth      =   225
         TabIndex        =   6
         Top             =   1290
         Width           =   285
      End
   End
   Begin MSComctlLib.StatusBar sBar 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   5730
      Width           =   7740
      _ExtentX        =   13653
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4057
            MinWidth        =   4057
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   4921
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4093
            MinWidth        =   4093
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame F1 
      BorderStyle     =   0  'None
      Height          =   5050
      Left            =   0
      TabIndex        =   0
      Top             =   450
      Width           =   8010
      Begin MSComctlLib.ListView statusLW 
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   3900
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   1720
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         HotTracking     =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         SmallIcons      =   "IL"
         ForeColor       =   9164914
         BackColor       =   0
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Server"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Date"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Error Message"
            Object.Width           =   17639
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Email/Page Status"
            Object.Width           =   5821
         EndProperty
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid G 
         Height          =   3585
         Left            =   120
         TabIndex        =   3
         Top             =   210
         Width           =   7635
         _ExtentX        =   13467
         _ExtentY        =   6324
         _Version        =   393216
         BackColor       =   15790320
         Cols            =   15
         FixedCols       =   0
         RowHeightMin    =   300
         BackColorFixed  =   14868694
         BackColorSel    =   10050097
         BackColorBkg    =   12630185
         GridColor       =   8224125
         HighLight       =   0
         AllowUserResizing=   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   15
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Status"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   0
         Width           =   645
      End
      Begin VB.Shape S1 
         Height          =   4850
         Left            =   30
         Top             =   120
         Width           =   7810
      End
   End
   Begin VB.Label Label2 
      Caption         =   "- Network Status"
      Height          =   255
      Index           =   1
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   1245
   End
   Begin VB.Menu mnuAction 
      Caption         =   "&Server"
      WindowList      =   -1  'True
      Begin VB.Menu mnuAdd 
         Caption         =   "Add Server"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuRemove 
         Caption         =   "Remove Server"
         Shortcut        =   ^R
      End
      Begin VB.Menu mnuSave 
         Caption         =   "&Save List as..."
      End
      Begin VB.Menu mnuBr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuCTRL 
         Caption         =   "Minimize to Tray"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
         Shortcut        =   ^X
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "&Tools"
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh Server Status"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuBr2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSrvmgr 
         Caption         =   "Computer Manager"
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuServices 
         Caption         =   "Services"
         Shortcut        =   ^V
      End
      Begin VB.Menu mnuEventvwr 
         Caption         =   "Event Viewer"
         Shortcut        =   ^E
      End
      Begin VB.Menu mnuSystem 
         Caption         =   "System Info."
         Shortcut        =   ^Y
      End
      Begin VB.Menu mnuCmd 
         Caption         =   "Command Prompt"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuRcmd 
         Caption         =   "&Remote CMD"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open Drive"
         Shortcut        =   ^D
      End
      Begin VB.Menu mnuBr3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPing 
         Caption         =   "Ping Server"
         Shortcut        =   {F9}
      End
      Begin VB.Menu mnuBr5 
         Caption         =   "-"
      End
      Begin VB.Menu mnuServicesMon 
         Caption         =   "Services Monitor"
      End
      Begin VB.Menu mnuPrefs 
         Caption         =   "Server &Options..."
      End
      Begin VB.Menu mnuOptions 
         Caption         =   "Global Options"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu mnuLog 
      Caption         =   "&Log"
      Begin VB.Menu mnuClear 
         Caption         =   "&Clear Log Window"
      End
      Begin VB.Menu mnuView 
         Caption         =   "&View Log File"
      End
   End
   Begin VB.Menu MnuH 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelp 
         Caption         =   "Help"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuBr4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About Server Monitor"
      End
   End
End
Attribute VB_Name = "frmSrvmtr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private sortTrigger     As Boolean
Private currY           As Integer
Private deltaY          As Integer
Private totalH          As Integer
Private coeffH          As Single
Private mouseDownEvent  As Boolean
Private GridX           As Integer
Private GridY           As Integer


Private Sub Form_Load()
Dim I As Integer
frmSplash.Show
frmSplash.Refresh
DoEvents
Call BuiltTable
Call OpenConfig(False)
totalH = G.Height + statusLW.Height
coeffH = G.Height / statusLW.Height
If Not FileExists(App.Path & "\SMlog.log") Then
    Open App.Path & "\SMlog.log" For Output As #5
    Print #5, "Server Monitor Error Log on \\" & Environ$("COMPUTERNAME")
    Print #5, "======================================"
    Close #5
End If
If G.TextMatrix(1, 0) = "" Then
    T1.Enabled = False
    sBar.Panels(2).Text = "Please add a server to the list."
    sBar.Panels(3).Text = ""
Else:
    timeDiff = Trim(GetConfigOptions.setInterval) / 60 / 24
    Call QueryServers(1, False, False, False, False)
End If
Unload frmSplash
isMin = True
TrayIcon.Picture = frmSrvmtr.Icon
For I = 0 To 4
    I1(I).Picture = IL.ListImages.Item(I + 9).Picture
Next
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If isMin Then
    Set objConn = New ADODB.Connection
    objConn.Open strConn
    objConn.Execute ("UPDATE tblConfig SET " & _
    "[WinWidth]=" & Me.Width & _
    ",[WinHeight]=" & Me.Height & _
    ",[WinTop]=" & Me.Top & _
    ",[WinLeft]=" & Me.Left & _
    ",[GHeight]=" & G.Height & _
    ",[statusLWHeight]=" & statusLW.Height & ";")
    
    objConn.Close
    Set objConn = Nothing
End If
DeleteIcon TrayIcon
Unload frmAddServer
Unload Me
End Sub




Private Sub mnuAbout_Click()
frmAbout.Show
End Sub

Private Sub mnuClear_Click()
If MsgBox("Sure to clear Log Window?", vbInformation + vbYesNo, "Confirm:") = vbYes Then statusLW.ListItems.Clear
End Sub

Private Sub mnuCmd_Click()
Call RunApplication("cmd")
End Sub


Private Sub mnuEventvwr_Click()
Call RunApplication("eventvwr")
End Sub

Private Sub mnuExit_Click()
Form_QueryUnload 1, 1
End Sub

Private Sub mnuHelp_Click()
RunApplication ("Help")
End Sub

Private Sub mnuOpen_Click()
Call RunApplication("open")
End Sub

Private Sub mnuOptions_Click()
frmOptions.Show vbModal
End Sub

Private Sub mnuPing_Click()
Dim R       As Integer
Dim pServer As String
Dim ret     As String

R = G.RowSel
If Not G.TextMatrix(R, 0) = "" Then
    pServer = G.TextMatrix(R, 0)
    ret = PingServer(pServer)
    
    Select Case Left(ret, 1)
        Case 0
            MsgBox "Server ''" & pServer & "'' is responded to Ping. IP Address: " & _
            Mid(ret, 2), vbDefaultButton1, "Success:"
        Case 1
            MsgBox "Server ''" & pServer & "'' is NOT responding to Ping. IP Address: " & _
            Mid(ret, 2), vbExclamation, "Ping Fail:"
        Case 2
            MsgBox "Server ''" & pServer & "''. Can NOT resolve server Name to IP. ", vbCritical, "Ping Fail:"
    End Select
    
End If
End Sub

Private Sub mnuPrefs_Click()
Dim R           As Integer

R = G.RowSel
If Not G.TextMatrix(R, 0) = "" Then
    If Not BuildPrefList(R) Then Exit Sub
    frmPref.Show vbModal
End If
End Sub

Private Sub mnuRcmd_Click()
Call RunApplication("rcmd")
End Sub

Private Sub mnuRefresh_Click()
If G.TextMatrix(1, 0) = "" Then Exit Sub
T1.Enabled = False
sBar.Panels(3).Text = "In Progress..."
Call QueryServers(G.RowSel, True, False, False, False)
End Sub

Private Sub mnuSearch_Click()
Call RunApplication("Search")
End Sub

Private Sub mnuServices_Click()
Call RunApplication("services")
End Sub

Private Sub mnuServicesMon_Click()
Dim R           As Integer
R = G.RowSel
If Not G.TextMatrix(R, 0) = "" Then
    If Not BuildServiceList(G.TextMatrix(R, 0)) Then Exit Sub
    frmServ.Show vbModal
End If
End Sub

Private Sub mnuTools_Click()
If G.TextMatrix(1, 0) = "" Then EnableMenu (False)
End Sub

Private Sub mnuView_Click()
Call RunApplication("openLog")
End Sub



Private Sub statusLW_DblClick()
If statusLW.ListItems.Count > 0 Then
    If statusLW.SelectedItem.SmallIcon = "StopEvent" Then
        Call RunApplication("EventLog", statusLW.SelectedItem.Text)
    End If
End If
End Sub

Private Sub T1_Timer()
Dim T           As String
Dim isEMail     As Boolean
Dim isPage      As Boolean

If TB.Buttons.Item(13).Value = tbrPressed Then T1.Enabled = False

timeDiff = DateDiff("s", Now, nextUpdate) / 60 / 24 / 60

If timeDiff <= #12:00:00 AM# Then
    sBar.Panels(3).Text = "In Progress..."
    T1.Enabled = False
    If CBool(GetConfigOptions.setEnableNote) Then
        isEMail = CBool(GetConfigOptions.setEnableEmail)
        isPage = CBool(GetConfigOptions.setEnablePage)
    Else: isPage = False: isEMail = False
    End If
    Call QueryServers(1, False, True, isEMail, isPage)
End If

T = Format(timeDiff, "hh:nn:ss")
sBar.Panels(3).Text = "Next Update in: " & T
End Sub
Private Sub mnuSave_Click()
Call SaveList
End Sub

Private Sub mnuSrvmgr_Click()
Call RunApplication("srvmgr")
End Sub

Private Sub mnuSystem_Click()
Call RunApplication("winmsd")
End Sub

Private Sub G_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim lastRow As String
Dim firstRow As String

With G
    If .MouseCol = 0 And .MouseRow = 0 Then
        lastRow = .TextMatrix(.Rows - 1, 0)
        firstRow = .TextMatrix(1, 0)
        If sortTrigger Then
            .Sort = flexSortGenericDescending
            sortTrigger = False
        Else
            .Sort = flexSortGenericAscending
            sortTrigger = True
        End If
    End If
End With

End Sub


Private Sub mnuAdd_Click()
frmAddServer.Show vbModal
End Sub

Private Sub mnuRemove_Click()
Dim removeDrive     As String
Dim removeServer    As String
Dim strSQL          As String

With G
    removeDrive = .TextMatrix(.RowSel, 1)
    removeServer = .TextMatrix(.RowSel, 0)
    .Row = .RowSel: .Col = 0
    If .Text = "" Then Exit Sub
    If MsgBox("Remove server [" & removeServer & "\" & removeDrive & "] from the list?", _
            vbExclamation + vbOKCancel, "Remove Server:") = vbCancel Then Exit Sub
    If .Rows = 2 Then
        Call BuiltTable
        T1.Enabled = False
        sBar.Panels(2).Text = "Please add a server to the list."
        sBar.Panels(3).Text = ""
    Else: .RemoveItem (.RowSel)
    End If
    Set objConn = New ADODB.Connection
    objConn.Open strConn
    strSQL = "DELETE FROM tblServer WHERE([sdrive]='" & removeDrive & "' AND [server]='" & removeServer & "');"
    objConn.Execute (strSQL)
    objConn.Close
    Set objConn = Nothing
End With
End Sub



Private Sub TB_ButtonClick(ByVal Button As MSComctlLib.Button)

Select Case Button.Key
    Case "add"
        mnuAdd_Click
    Case "remove"
        mnuRemove_Click
    Case "exit"
        mnuExit_Click
    Case "refresh"
        If G.TextMatrix(1, 0) = "" Then Exit Sub
        T1.Enabled = False
        sBar.Panels(3).Text = "In Progress..."
        Call QueryServers(1, False, False, False, False)
    Case "clearLog"
        mnuClear_Click
    Case "view"
        mnuView_Click
    Case "stopTimer"
        If G.TextMatrix(1, 0) = "" Then Exit Sub
        If Button.Value = tbrPressed Then
            TB.Buttons.Item(14).Image = 15
            T1.Enabled = False
        Else
            TB.Buttons.Item(14).Image = 14
            nextUpdate = Now + timeDiff
            T1.Enabled = True
        End If
    Case "help"
        mnuHelp_Click
End Select

End Sub
Private Sub F1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Y < statusLW.Top And Y > G.Top + G.Height Then
    mouseDownEvent = True
    currY = Y
    deltaY = Y - G.Height
    totalH = G.Height + statusLW.Height
Else: mouseDownEvent = False
End If
End Sub


Private Sub F1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next

If Y < statusLW.Top And Y > G.Top + G.Height Then
    Me.MousePointer = vbCustom
    Me.MouseIcon = IL.ListImages(2).Picture
Else: Me.MousePointer = vbDefault
End If
If mouseDownEvent = True Then
    If statusLW.Height > 100 Or (statusLW.Height <= 100 And Y <= currY) Then
        Err.Clear
        G.Move G.Left, G.Top, G.Width, Y - deltaY
        statusLW.Move statusLW.Left, G.Top + G.Height + 100, statusLW.Width, totalH - G.Height
    End If
    If G.Height < 150 Then
        G.Visible = False
    Else: G.Visible = True
    End If
    If statusLW.Height < 150 Then
        statusLW.Visible = False
    Else: statusLW.Visible = True
    End If
    If G.Height > totalH - 100 Then
        G.Height = totalH - 100
        statusLW.Height = totalH - G.Height
        statusLW.Top = G.Top + G.Height + 100
    End If
End If
End Sub
Private Sub F1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
mouseDownEvent = False
coeffH = G.Height / statusLW.Height
End Sub
Private Sub G_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub
Private Sub statusLW_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.MousePointer = vbDefault
End Sub
Private Sub Form_Resize()
On Error Resume Next
With F1
    .Move .Left, .Top, Me.Width, Me.Height - 1850
End With
With F2
    .Move .Left, F1.Top + F1.Height - 100, Me.Width, .Height
End With
With S1
    .Move .Left, .Top, Me.Width - 200, Me.Height - 2150
End With
With G
    .Move .Left, .Top, Me.Width - 375, ((F1.Height - 490) * coeffH) / (coeffH + 1)
End With
With statusLW
    .Move .Left, G.Height + 310, Me.Width - 375, (F1.Height - G.Height) - 590
End With
If Me.WindowState = 1 And isMin Then NoSysIcon False
BringWindowToTop Me.hwnd
Me.SetFocus
End Sub
Private Sub G_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
With G
    Select Case Button
        Case vbRightButton
            .SetFocus
            .Col = .MouseCol
            .Row = .MouseRow
            If .TextMatrix(.Row, .Col) = "" Or _
                .Row = 0 Then
                EnableMenu (False): PopupMenu mnuTools, vbPopupMenuRightButton, , , mnuRefresh
            Else: EnableMenu (True): PopupMenu mnuTools, vbPopupMenuRightButton, , , mnuRefresh
            End If
        Case vbLeftButton
        GridX = X
        GridY = Y
    End Select
End With
End Sub
Private Sub EnableMenu(showMenu As Boolean)

If showMenu = False Then
    mnuSrvmgr.Enabled = False
    mnuEventvwr.Enabled = False
    mnuSystem.Enabled = False
    mnuCmd.Enabled = False
    mnuRcmd.Enabled = False
    mnuOpen.Enabled = False
    mnuPrefs.Enabled = False
    mnuPing.Enabled = False
    mnuServices.Enabled = False
    mnuServicesMon.Enabled = False
Else
    mnuSrvmgr.Enabled = True
    mnuEventvwr.Enabled = True
    mnuSystem.Enabled = True
    mnuCmd.Enabled = True
    mnuRcmd.Enabled = CBool(GetConfigOptions.setEnableRCMD)
    mnuOpen.Enabled = True
    mnuPrefs.Enabled = True
    mnuPing.Enabled = True
    mnuServices.Enabled = True
    mnuServicesMon.Enabled = True
End If
End Sub
Private Sub statusLW_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    statusLW.SetFocus
    PopupMenu mnuLog, vbPopupMenuRightButton
End If
End Sub
Public Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Winsock1.GetData strRespCode
strFullResp = strFullResp & vbCrLf & strRespCode
strRespCode = Left$(strRespCode, 3)
End Sub
Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
Call UpdateEmailStatus("Unable to Mail: " & Description)
Call LogErrors("Mail/Page Error", "Unable to Send", "Error: " & _
            Number & ", Source: " & Source & ", Description: " & Description)
Winsock1.Close
End Sub
Private Sub T2_Timer()
Timeout = Timeout + 1
End Sub
Public Function ShowProgramInTray()
    Dim result As Long
    INTRAY = True
    NI.cbSize = Len(NI)
    NI.hwnd = TrayIcon.hwnd
    NI.uID = 0
    NI.uID = NI.uID + 1
    NI.uFlags = NIF_MESSAGE Or NIF_ICON Or NIF_TIP
    NI.uCallbackMessage = WM_MOUSEMOVE
    NI.hIcon = frmSrvmtr.Icon
    NI.szTip = "Server Monitor" + Chr$(0)
    result = Shell_NotifyIconA(NIM_ADD, NI)
End Function
Private Sub DeleteIcon(pic As Control)
Dim result As Long
    INTRAY = False
    NI.uID = 0
    NI.uID = NI.uID + 1
    NI.cbSize = Len(NI)
    NI.hwnd = pic.hwnd
    NI.uCallbackMessage = WM_MOUSEMOVE
    result = Shell_NotifyIconA(NIM_DELETE, NI)
End Sub
Public Function NoSysIcon(maxIcon As Boolean)
    Select Case maxIcon
    Case False
        Me.Visible = False
        ShowProgramInTray
        mnuCTRL.Caption = "Expand Monitor"
        isMin = False
    Case Else
        Me.Visible = True
        DeleteIcon TrayIcon
        mnuCTRL.Caption = "Minimize to Tray"
        isMin = True
        Me.WindowState = 0
    End Select
End Function
Private Sub mnuCTRL_Click()
NoSysIcon INTRAY
End Sub
Private Sub TrayIcon_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim result As Long
    Dim msg As Long
    msg = (X And &HFF) * &H100
    Select Case msg
        Case &H3C00
            PopupMenu mnuAction, 2, , , mnuCTRL
        Case &H2D00
            NoSysIcon True
    End Select
End Sub

