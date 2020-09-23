VERSION 5.00
Begin VB.Form frmAbout 
   BackColor       =   &H00643D0D&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2370
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2370
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00643D0D&
      Caption         =   "URL:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBA97D&
      Height          =   225
      Index           =   3
      Left            =   690
      TabIndex        =   2
      Top             =   2040
      Width           =   360
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00643D0D&
      Caption         =   "www.gurgensvbstuff.com"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBA97D&
      Height          =   225
      Index           =   2
      Left            =   1170
      MouseIcon       =   "frmAbout.frx":0000
      MousePointer    =   99  'Custom
      TabIndex        =   1
      Top             =   2040
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00643D0D&
      Caption         =   "Author:  GASoft"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00BBA97D&
      Height          =   225
      Index           =   0
      Left            =   690
      TabIndex        =   0
      Top             =   1740
      Width           =   1260
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1530
      Left            =   555
      Picture         =   "frmAbout.frx":030A
      Top             =   150
      Width           =   3510
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Click()
Unload Me
End Sub
Private Sub Image1_Click()
Unload Me
End Sub
Private Sub Label1_Click(Index As Integer)
Dim strOpen As String
Select Case Index
    Case 1
        strOpen = "mailto:" & Label1(Index).Caption
        ShellExecute Me.hwnd, "open", strOpen, openServer, vbNullString, SW_SHOWNORMAL
    Case 2
        strOpen = "http://www.gurgensvbstuff.com/index.php?ID=04"
        ShellExecute Me.hwnd, "open", strOpen, openServer, vbNullString, SW_SHOWNORMAL
    Case Else
        Unload Me
End Select
End Sub

