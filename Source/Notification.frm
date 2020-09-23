VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   6855
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "Notification.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6855
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtHost 
      Enabled         =   0   'False
      Height          =   315
      IMEMode         =   3  'DISABLE
      Index           =   2
      Left            =   2220
      PasswordChar    =   "*"
      TabIndex        =   37
      ToolTipText     =   "Enter Your SMTP Logon Password ONLY if your SMTP Host requires Authentication."
      Top             =   1710
      Width           =   2400
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Notification Method"
      Height          =   705
      Left            =   60
      TabIndex        =   25
      Top             =   5610
      Width           =   4710
      Begin VB.CheckBox chkRcmd 
         Caption         =   "RCMD is available on servers"
         Height          =   285
         Left            =   90
         TabIndex        =   26
         Top             =   240
         Width           =   2865
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "RCMD"
         Height          =   195
         Left            =   150
         TabIndex        =   27
         Top             =   0
         Width           =   615
      End
      Begin VB.Shape Shape3 
         Height          =   525
         Left            =   0
         Top             =   90
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Height          =   330
      Left            =   3570
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6450
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Notification Method"
      Height          =   2895
      Left            =   60
      TabIndex        =   5
      Top             =   90
      Width           =   4710
      Begin VB.TextBox txtHost 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   2160
         TabIndex        =   36
         ToolTipText     =   "Enter Your SMTP Logon ID ONLY if your SMTP Host requires Authentication."
         Top             =   1260
         Width           =   2400
      End
      Begin VB.PictureBox P1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00DBC591&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   2580
         Picture         =   "Notification.frx":08CA
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   33
         Top             =   270
         Width           =   510
      End
      Begin VB.PictureBox P2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00DBC591&
         ForeColor       =   &H80000008&
         Height          =   510
         Left            =   3840
         Picture         =   "Notification.frx":1194
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   32
         Top             =   270
         Width           =   510
      End
      Begin VB.TextBox txtHost 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   2160
         TabIndex        =   30
         ToolTipText     =   "Enter Your SMTP Host. Port 25 will be used. If you're behind proxy, then add SMTP_Host:Port_Number."
         Top             =   900
         Width           =   2400
      End
      Begin VB.OptionButton optSMTP 
         Caption         =   "SMTP"
         Enabled         =   0   'False
         Height          =   225
         Left            =   90
         TabIndex        =   29
         Top             =   960
         Width           =   825
      End
      Begin VB.OptionButton optOutlook 
         Caption         =   "Outlook"
         Enabled         =   0   'False
         Height          =   225
         Left            =   90
         TabIndex        =   28
         Top             =   630
         Value           =   -1  'True
         Width           =   1245
      End
      Begin VB.CheckBox chkEnable 
         Caption         =   "Enable Notification"
         Height          =   285
         Left            =   90
         TabIndex        =   16
         Top             =   240
         Width           =   1695
      End
      Begin VB.TextBox txtEmail 
         Enabled         =   0   'False
         Height          =   315
         Left            =   2160
         TabIndex        =   9
         Top             =   2010
         Width           =   2400
      End
      Begin VB.TextBox txtPage 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2160
         TabIndex        =   8
         ToolTipText     =   "Depending on your Pager service provider the address could be ""Local_Phone_Number.Pin_Number@provider.com"""
         Top             =   2385
         Width           =   2400
      End
      Begin VB.CheckBox chkEmail 
         Caption         =   "Email me"
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   7
         Top             =   2010
         Width           =   945
      End
      Begin VB.CheckBox chkPage 
         Caption         =   "Page me"
         Enabled         =   0   'False
         Height          =   315
         Left            =   90
         TabIndex        =   6
         Top             =   2340
         Width           =   975
      End
      Begin VB.Label L5 
         Caption         =   "SMTP ID and Password are Optional"
         Enabled         =   0   'False
         Height          =   645
         Index           =   3
         Left            =   90
         TabIndex        =   40
         Top             =   1290
         Width           =   1035
      End
      Begin VB.Label L5 
         Caption         =   "Password:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   1290
         TabIndex        =   39
         Top             =   1650
         Width           =   765
      End
      Begin VB.Label L5 
         Caption         =   "SMTP ID:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   38
         Top             =   1290
         Width           =   765
      End
      Begin VB.Label L7 
         Alignment       =   2  'Center
         Caption         =   "Test:"
         Height          =   195
         Left            =   3390
         TabIndex        =   35
         Top             =   420
         Width           =   405
      End
      Begin VB.Label L6 
         Alignment       =   2  'Center
         Caption         =   "Test:"
         Height          =   195
         Left            =   2130
         TabIndex        =   34
         Top             =   420
         Width           =   375
      End
      Begin VB.Label L5 
         Caption         =   "SMTP Host:"
         Enabled         =   0   'False
         Height          =   255
         Index           =   0
         Left            =   1170
         TabIndex        =   31
         Top             =   960
         Width           =   945
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "Notification Method"
         Height          =   195
         Left            =   150
         TabIndex        =   10
         Top             =   0
         Width           =   1515
      End
      Begin VB.Label L1 
         Caption         =   "Address:"
         Enabled         =   0   'False
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         Top             =   2070
         Width           =   765
      End
      Begin VB.Label L2 
         Caption         =   "Pin.Address:"
         Enabled         =   0   'False
         Height          =   240
         Left            =   1170
         TabIndex        =   11
         Top             =   2400
         Width           =   900
      End
      Begin VB.Shape Shape2 
         Height          =   2715
         Left            =   0
         Top             =   90
         Width           =   4695
      End
   End
   Begin VB.Frame Option 
      BorderStyle     =   0  'None
      Caption         =   "Notification Method"
      Height          =   2535
      Left            =   60
      TabIndex        =   0
      Top             =   3030
      Width           =   4710
      Begin VB.CheckBox chkEventLog 
         Caption         =   "Enable Event Log Monitor"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1020
         Width           =   2205
      End
      Begin VB.CheckBox chkApp 
         Caption         =   "Query Application Log"
         Enabled         =   0   'False
         Height          =   255
         Left            =   2580
         TabIndex        =   21
         Top             =   1290
         Width           =   2025
      End
      Begin VB.CheckBox chkSystem 
         Caption         =   "Query System Log"
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         TabIndex        =   20
         Top             =   1290
         Width           =   1605
      End
      Begin VB.ComboBox cmbMinSpace 
         Height          =   315
         ItemData        =   "Notification.frx":1A5E
         Left            =   2580
         List            =   "Notification.frx":1A74
         TabIndex        =   18
         Text            =   "10000"
         ToolTipText     =   "Only Numbers accepted. Drive space in KB."
         Top             =   570
         Width           =   1005
      End
      Begin VB.ComboBox cmbInterval 
         Height          =   315
         ItemData        =   "Notification.frx":1AA6
         Left            =   2580
         List            =   "Notification.frx":1ABF
         TabIndex        =   14
         Text            =   "10"
         ToolTipText     =   "Only numbers accepted. Can not be more than 1440 (24 hours)."
         Top             =   210
         Width           =   1005
      End
      Begin VB.Label L3 
         Caption         =   $"Notification.frx":1AE0
         Enabled         =   0   'False
         Height          =   825
         Left            =   600
         TabIndex        =   23
         Top             =   1650
         Width           =   3915
      End
      Begin VB.Label L4 
         Caption         =   "Note: "
         Enabled         =   0   'False
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   1650
         Width           =   375
      End
      Begin VB.Label Label4 
         Caption         =   "Kb."
         Height          =   240
         Left            =   3720
         TabIndex        =   19
         Top             =   630
         Width           =   420
      End
      Begin VB.Label Label3 
         Caption         =   "Notify if Drive Space is  less than:"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   630
         Width           =   2445
      End
      Begin VB.Label Label2 
         Caption         =   "Minute(s)"
         Height          =   240
         Left            =   3690
         TabIndex        =   15
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label9 
         Caption         =   "Query Servers every:"
         Height          =   285
         Left            =   90
         TabIndex        =   13
         Top             =   270
         Width           =   1500
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "Query Options"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   0
         Width           =   1125
      End
      Begin VB.Shape Shape1 
         Height          =   2415
         Left            =   0
         Top             =   90
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   6450
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   1230
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6450
      Width           =   1185
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_Load()
Call SetOptions
cmdApply.Enabled = False
End Sub

Private Sub cmdApply_Click()
Call ApplySettings
End Sub
Private Sub cmdOK_Click()
If Not cmdApply.Enabled Then Unload Me: Exit Sub
If ApplySettings Then Unload Me
End Sub

Private Sub chkApp_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkEmail_Click()
cmdApply.Enabled = True
If chkEmail.Value = 0 And chkPage.Value = 0 Then chkEnable.Value = 0
Select Case chkEmail.Value
    Case 0
        L1.Enabled = False
        txtEmail.Enabled = False
        L6.Enabled = False
        P1.Enabled = False
    Case 1
        L1.Enabled = True
        txtEmail.Enabled = True
        L6.Enabled = True
        P1.Enabled = True
End Select
End Sub
Private Sub chkPage_Click()
cmdApply.Enabled = True
If chkEmail.Value = 0 And chkPage.Value = 0 Then chkEnable.Value = 0
Select Case chkPage.Value
    Case 0
        L2.Enabled = False
        txtPage.Enabled = False
        L7.Enabled = False
        P2.Enabled = False
    Case 1
        L2.Enabled = True
        txtPage.Enabled = True
        L7.Enabled = True
        P2.Enabled = True
End Select
End Sub

Private Sub chkEnable_Click()
Call SetNotification
End Sub

Private Sub chkEventLog_Click()
Call SetEventLog
End Sub
Private Sub chkRcmd_Click()
frmSrvmtr.mnuRcmd.Enabled = CBool(chkRcmd.Value)
cmdApply.Enabled = True
End Sub


Private Sub chkSystem_Click()
cmdApply.Enabled = True
End Sub

Private Sub optOutlook_Click()
cmdApply.Enabled = True
Call SetMailType
End Sub

Private Sub optSMTP_Click()
cmdApply.Enabled = True
Call SetMailType
End Sub

Private Sub P1_Click()
P1.Appearance = 1
P1.BackColor = &HA6F997
If InStr(Trim(txtEmail.Text), ".") = 0 Or InStr(Trim(txtEmail.Text), "@") = 0 Then
    MsgBox "Email Address is invalid", vbExclamation, "Error:"
    P1.Appearance = 0
    P1.BackColor = &HDBC591
    Exit Sub
End If
Call ApplySettings
testMail = True
If SendMail("Test message from Server Monitor.", "Email") = True Then
    MsgBox "Test Email successfully sent.", vbDefaultButton1, "Success:"
Else: MsgBox "Error sending Page." & vbCrLf & vbCrLf & "SERVER RESPONSE: " & vbCrLf & strFullResp, vbCritical, "Error:"
End If
strFullResp = Empty
testMail = False
frmSrvmtr.Winsock1.Close
P1.Appearance = 0
P1.BackColor = &HDBC591
End Sub

Private Sub P2_Click()
P2.Appearance = 1
P2.BackColor = &HA6F997
If InStr(Trim(txtPage.Text), ".") = 0 Or InStr(Trim(txtPage.Text), "@") = 0 Then
    MsgBox "Pager Address is invalid", vbExclamation, "Error:"
    P2.Appearance = 0
    P2.BackColor = &HDBC591
    Exit Sub
End If
Call ApplySettings
testMail = True
If SendMail("Test message from Server Monitor.", "Page") = True Then
    MsgBox "Test Page successfully sent.", vbDefaultButton1, "Success:"
Else: MsgBox "Error sending Page." & vbCrLf & vbCrLf & "SERVER RESPONSE: " & vbCrLf & strFullResp, vbCritical, "Error:"
End If
testMail = False
frmSrvmtr.Winsock1.Close
strFullResp = Empty
P2.Appearance = 0
P2.BackColor = &HDBC591
End Sub

Private Sub txtEmail_Change()
cmdApply.Enabled = True
End Sub



Private Sub txtHost_Change(Index As Integer)
cmdApply.Enabled = True
End Sub

Private Sub txtPage_Change()
cmdApply.Enabled = True
End Sub
Private Sub cmbInterval_Change()
cmdApply.Enabled = True
End Sub

Private Sub cmbInterval_Click()
cmdApply.Enabled = True
End Sub
Private Sub cmbMinSpace_Change()
cmdApply.Enabled = True
End Sub

Private Sub cmbMinSpace_Click()
cmdApply.Enabled = True
End Sub
Private Sub cmdCancel_Click()
Unload Me
End Sub
Private Sub cmbInterval_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) > 9 And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
Private Sub cmbMinSpace_KeyPress(KeyAscii As Integer)
If Chr(KeyAscii) > 9 And Not KeyAscii = 8 Then KeyAscii = 0
End Sub
