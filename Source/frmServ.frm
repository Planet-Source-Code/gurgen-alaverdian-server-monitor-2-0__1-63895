VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmServ 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Services Monitor"
   ClientHeight    =   4710
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   7530
   Icon            =   "frmServ.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4710
   ScaleWidth      =   7530
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command1 
      Caption         =   "Services (.msc)"
      Height          =   330
      Left            =   90
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4290
      Width           =   1425
   End
   Begin VB.CommandButton C1 
      Caption         =   "<"
      Height          =   465
      Index           =   1
      Left            =   5430
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2670
      Width           =   375
   End
   Begin VB.CommandButton C1 
      Caption         =   ">"
      Height          =   465
      Index           =   0
      Left            =   5430
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   2010
      Width           =   375
   End
   Begin VB.ListBox L3 
      Height          =   2790
      Left            =   5910
      MultiSelect     =   2  'Extended
      TabIndex        =   7
      Top             =   1230
      Width           =   1455
   End
   Begin MSComctlLib.ListView L2 
      Height          =   2805
      Left            =   180
      TabIndex        =   6
      Top             =   1230
      Width           =   5115
      _ExtentX        =   9022
      _ExtentY        =   4948
      View            =   3
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Service Name"
         Object.Width           =   2999
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Display Name"
         Object.Width           =   8819
      EndProperty
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
      Left            =   1050
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.CommandButton cmdApply 
      Caption         =   "&Apply"
      Enabled         =   0   'False
      Height          =   330
      Left            =   6270
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4290
      Width           =   1185
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   330
      Left            =   5070
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4290
      Width           =   1185
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   330
      Left            =   3870
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4290
      Width           =   1185
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Server is Down"
      ForeColor       =   &H000000FF&
      Height          =   195
      Left            =   3510
      TabIndex        =   16
      Top             =   540
      Visible         =   0   'False
      Width           =   1155
   End
   Begin VB.Label Lb1 
      BackStyle       =   0  'Transparent
      Height          =   225
      Index           =   1
      Left            =   6720
      TabIndex        =   15
      Top             =   990
      Width           =   645
   End
   Begin VB.Label Lb1 
      BackStyle       =   0  'Transparent
      Height          =   225
      Index           =   0
      Left            =   2040
      TabIndex        =   14
      Top             =   990
      Width           =   645
   End
   Begin VB.Label Label4 
      BackColor       =   &H00CAC7A6&
      BackStyle       =   0  'Transparent
      Caption         =   "Note: Only running services are displayed in the list."
      Height          =   225
      Left            =   3510
      TabIndex        =   10
      Top             =   300
      Width           =   3915
   End
   Begin VB.Label Label3 
      BackColor       =   &H00CAC7A6&
      BackStyle       =   0  'Transparent
      Caption         =   "Non Watched Services:"
      Height          =   195
      Index           =   1
      Left            =   180
      TabIndex        =   9
      Top             =   990
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackColor       =   &H00CAC7A6&
      BackStyle       =   0  'Transparent
      Caption         =   "Monitored:"
      Height          =   195
      Index           =   0
      Left            =   5910
      TabIndex        =   8
      Top             =   990
      Width           =   825
   End
   Begin VB.Label Label2 
      BackColor       =   &H00CAC7A6&
      BackStyle       =   0  'Transparent
      Caption         =   "Server:"
      Height          =   255
      Left            =   330
      TabIndex        =   5
      Top             =   540
      Width           =   615
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Services Monitor"
      Height          =   255
      Left            =   300
      TabIndex        =   3
      Top             =   60
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      Height          =   3975
      Left            =   60
      Top             =   210
      Width           =   7425
   End
End
Attribute VB_Name = "frmServ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub C1_Click(Index As Integer)
Dim I As Integer
cmdApply.Enabled = True

Select Case Index
    Case 0
        For I = L2.ListItems.Count To 1 Step -1
            If L2.ListItems.Item(I).Selected Then
                L3.AddItem (L2.ListItems.Item(I).Text)
                L2.ListItems.Remove (I)
            End If
        Next
    Case 1
        L3.SetFocus
        For I = L3.ListCount - 1 To 0 Step -1
            If L3.Selected(I) Then L3.RemoveItem (I)
        Next I
        Call EnumSystemServices(L1.Text, BuildServList())
End Select
Lb1(0).Caption = L2.ListItems.Count
Lb1(1).Caption = L3.ListCount
End Sub
Public Function BuildServList() As String
Dim strBuildServ As String
    For I = 0 To L3.ListCount - 1
        strBuildServ = strBuildServ & L3.List(I) & "|"
    Next I
    If Len(strBuildServ) > 0 Then strBuildServ = Left(strBuildServ, Len(strBuildServ) - 1)
    BuildServList = strBuildServ
End Function

Private Sub cmdApply_Click()
Call ApplyServPref(L1.Text)
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()

If Not cmdApply.Enabled Then Unload Me: Exit Sub
Call ApplyServPref(L1.Text)
Unload Me

End Sub

Private Sub Command1_Click()
Dim pingStatus As String
pingStatus = PingServer(L1.Text)
If Not Left(pingStatus, 1) = 0 Then MsgBox "Server is down!", vbExclamation, "Monitor": Exit Sub
Call ShellExecute(Me.hwnd, "open", "services.msc", " /computer:" & L1.Text, vbNullString, SW_SHOWNORMAL)
End Sub

Private Sub Form_Load()
cmdApply.Enabled = False
End Sub

Private Sub L1_Click()
If Not servFlag Then Call BuildServiceList(L1.Text)
End Sub

Private Sub L2_DblClick()
C1_Click 0
End Sub

Private Sub L3_DblClick()
C1_Click 1
End Sub
