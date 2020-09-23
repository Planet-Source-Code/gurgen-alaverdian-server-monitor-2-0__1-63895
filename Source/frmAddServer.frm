VERSION 5.00
Begin VB.Form frmAddServer 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Add Server"
   ClientHeight    =   1095
   ClientLeft      =   7140
   ClientTop       =   7680
   ClientWidth     =   5685
   ControlBox      =   0   'False
   Icon            =   "frmAddServer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   5685
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame AddFrame 
      BorderStyle     =   0  'None
      Caption         =   "Add Server"
      Height          =   1035
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   5700
      Begin VB.TextBox txtComment 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   3
         Top             =   570
         Width           =   4455
      End
      Begin VB.CommandButton Close 
         Caption         =   "&Close"
         Height          =   330
         Left            =   4680
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   180
         Width           =   855
      End
      Begin VB.ComboBox cmbDrive 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmAddServer.frx":0442
         Left            =   2970
         List            =   "frmAddServer.frx":048E
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   765
      End
      Begin VB.TextBox txtServer 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   1080
         TabIndex        =   1
         Top             =   185
         Width           =   1275
      End
      Begin VB.CommandButton AddServer 
         Caption         =   "&Add"
         Height          =   330
         Left            =   3840
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00CAC7A6&
         BackStyle       =   0  'Transparent
         Caption         =   "Comments:"
         Height          =   225
         Left            =   60
         TabIndex        =   8
         Top             =   600
         Width           =   765
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         Height          =   915
         Left            =   0
         Top             =   90
         Width           =   5625
      End
      Begin VB.Label Label2 
         BackColor       =   &H00CAC7A6&
         BackStyle       =   0  'Transparent
         Caption         =   "Drive:"
         Height          =   225
         Left            =   2490
         TabIndex        =   7
         Top             =   240
         Width           =   465
      End
      Begin VB.Label Label1 
         BackColor       =   &H00CAC7A6&
         BackStyle       =   0  'Transparent
         Caption         =   "Server Name:"
         Height          =   225
         Left            =   60
         TabIndex        =   6
         Top             =   240
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmAddServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub AddServer_Click()
Dim strServer   As String
Dim strDrive    As String
Dim R           As Integer
Dim C           As Integer
Dim strSQL      As String
Dim rs          As Recordset
Dim serv        As String
Dim pingStatus  As String

On Error GoTo ErrorHandler
strDrive = cmbDrive.Text
strServer = UCase(Replace(txtServer.Text, "\", ""))

If strServer = "" Then MsgBox "Please type a server name first!", vbExclamation, "Monitor": Exit Sub
pingStatus = PingServer(strServer)
If Not Left(pingStatus, 1) = 0 Then MsgBox "Server Not Found!", vbExclamation, "Monitor": Exit Sub
If Not FileExists("\\" & strServer & "\" & strDrive) Then MsgBox "Server name or drive is invalid!", _
                vbExclamation, "Monitor": Exit Sub
                
With frmSrvmtr.G
    If .Rows >= 2 And Not .TextMatrix(1, 0) = "" Then
        For R = 1 To .Rows - 1
            If strServer = .TextMatrix(R, 0) And strDrive = .TextMatrix(R, 1) Then
                MsgBox "This server\drive is already in the list.", vbInformation, "Duplicate Entry:"
                Exit Sub
            End If
        Next R
    End If

    If Not .TextMatrix(1, 0) = "" Then
        .AddItem strServer
    Else: .TextMatrix(1, 0) = strServer
    End If
    R = .Rows - 1
    .TextMatrix(R, 1) = strDrive
    .TextMatrix(R, 9) = txtComment.Text
    For C = 10 To 13
        .TextMatrix(R, C) = 0
    Next C
    Call StatusOK(R)
    
    Set objConn = New ADODB.Connection
    objConn.Open strConn
    
    Set rs = objConn.Execute("SELECT TOP 1 [MonServices] FROM tblServer WHERE [server]='" & sq(strServer) & "' AND [MonServices]<>'';")
    If Not rs.EOF Then
        serv = rs(0)
        .TextMatrix(R, 14) = serv
    Else: serv = Empty
    End If
    rs.Close
    strSQL = "INSERT INTO tblServer([server],[sdrive],[comment],[noEmail],[noPage],[noSysLog],[noAppLog],[MonServices]) "
    strSQL = strSQL & "VALUES ('" & sq(strServer) & "','" & strDrive & "','" & sq(txtComment.Text) & "',False,False,False,False,'" & serv & "');"
    objConn.Execute (strSQL)

    objConn.Close
    Set objConn = Nothing
    
    frmSrvmtr.sBar.Panels(3).Text = "In Progress..."
    If .Rows = 2 Then timeDiff = Trim(GetConfigOptions.setInterval) / 60 / 24
    Call QueryServers(.Rows - 1, True, False, False, False)
End With
txtServer.Text = ""
txtServer.SetFocus
Exit Sub

ErrorHandler:
MsgBox "Error while adding server. Error number: " & Err.Number & ". Error description: " & Err.Description, vbCritical, "Error:"
    
End Sub
Private Sub Close_Click()
Unload Me
End Sub
Private Sub Form_Load()
cmbDrive.Text = "D$"
End Sub
Private Sub txtComment_KeyPress(KeyAscii As Integer)
If Len(txtComment.Text) > 100 Then
    Beep
    KeyAscii = 0
End If
End Sub
Private Sub txtServer_KeyPress(KeyAscii As Integer)
If Len(txtServer.Text) >= 15 Then
    Beep
    KeyAscii = 0
End If
End Sub

