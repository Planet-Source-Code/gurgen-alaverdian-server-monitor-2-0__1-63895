VERSION 5.00
Begin VB.Form frmPref 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Server Options"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4830
   Icon            =   "frmPref.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CheckBox setAll 
      Caption         =   "Set Above Settings to ALL Listed Drives for this Server"
      Height          =   255
      Left            =   150
      TabIndex        =   16
      Top             =   2820
      Width           =   4395
   End
   Begin VB.ComboBox L2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3930
      Style           =   2  'Dropdown List
      TabIndex        =   15
      Top             =   480
      Width           =   735
   End
   Begin VB.ComboBox L1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   480
      Width           =   2295
   End
   Begin VB.TextBox txtComment 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1020
      TabIndex        =   11
      Top             =   900
      Width           =   3645
   End
   Begin VB.CheckBox chkNoAppLog 
      Caption         =   "Do not Check Application Log on this Server"
      Height          =   315
      Left            =   150
      TabIndex        =   9
      Top             =   2400
      Width           =   3525
   End
   Begin VB.CheckBox chkNoSysLog 
      Caption         =   "Do not Check System Log on this Server"
      Height          =   315
      Left            =   150
      TabIndex        =   8
      Top             =   2070
      Width           =   3405
   End
   Begin VB.CheckBox chkNoPage 
      Caption         =   "Do not Page me about this Server"
      Height          =   315
      Left            =   150
      TabIndex        =   7
      Top             =   1740
      Width           =   3015
   End
   Begin VB.CheckBox chkNoEmail 
      Caption         =   "Do not Email me about this Server"
      Height          =   315
      Left            =   150
      TabIndex        =   6
      Top             =   1410
      Width           =   3015
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   1260
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3210
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   2430
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3210
      Width           =   1185
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3210
      Width           =   1185
   End
   Begin VB.Label L4 
      Caption         =   "Label5"
      Height          =   135
      Left            =   2670
      TabIndex        =   13
      Top             =   990
      Width           =   405
   End
   Begin VB.Label L3 
      BackColor       =   &H00CAC7A6&
      Height          =   225
      Left            =   3870
      TabIndex        =   12
      Top             =   960
      Width           =   225
   End
   Begin VB.Label Label4 
      Caption         =   "Comment:"
      Height          =   255
      Left            =   150
      TabIndex        =   10
      Top             =   960
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Drive:"
      Height          =   255
      Left            =   3450
      TabIndex        =   5
      Top             =   510
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "Server:"
      Height          =   255
      Left            =   150
      TabIndex        =   4
      Top             =   510
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Server Preferences"
      Height          =   255
      Left            =   180
      TabIndex        =   3
      Top             =   60
      Width           =   1515
   End
   Begin VB.Shape Shape1 
      Height          =   2925
      Left            =   60
      Top             =   210
      Width           =   4695
   End
End
Attribute VB_Name = "frmPref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkNoAppLog_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkNoEmail_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkNoPage_Click()
cmdApply.Enabled = True
End Sub

Private Sub chkNoSysLog_Click()
cmdApply.Enabled = True
End Sub

Private Sub cmdApply_Click()
Call ApplyPref(L4.Caption)
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
If Not cmdApply.Enabled Then Unload Me: Exit Sub
Call ApplyPref(L4.Caption)
Unload Me
End Sub


Private Sub Form_Load()
cmdApply.Enabled = False
End Sub


Private Sub FindServerRow(serverOnly As Boolean)
Dim nServer As String
Dim nDrive  As String
Dim C       As Integer

nServer = L1.Text
nDrive = L2.Text
L1.Clear
L2.Clear
With frmSrvmtr.G
    For C = 1 To .Rows - 1
        If .TextMatrix(C, 0) = nServer Then
            If Not serverOnly Then
                If .TextMatrix(C, 1) = nDrive Then
                    Call BuildPrefList(C)
                    Exit Sub
                End If
            Else
                Call BuildPrefList(C)
                Exit Sub
            End If
        End If
    Next C
End With

End Sub

Private Sub L1_Click()
If Not prefFlag Then FindServerRow (True)
End Sub
Private Sub L2_Click()
If Not prefFlag Then FindServerRow (False)
End Sub

Private Sub txtComment_KeyPress(KeyAscii As Integer)
cmdApply.Enabled = True
End Sub
