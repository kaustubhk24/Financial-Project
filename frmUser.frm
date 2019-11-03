VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form8 
   Caption         =   "Form8"
   ClientHeight    =   4590
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11205
   Icon            =   "frmUser.frx":0000
   LinkTopic       =   "Form8"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   NegotiateMenus  =   0   'False
   Picture         =   "frmUser.frx":2A8B2
   ScaleHeight     =   4590
   ScaleWidth      =   11205
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   4920
      Top             =   240
   End
   Begin ComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   720
      TabIndex        =   7
      Top             =   4200
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   661
      _Version        =   327682
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000018&
      Height          =   3375
      Left            =   5640
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   5175
      Begin VB.CommandButton cmdok 
         BackColor       =   &H8000000D&
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton cmdexit 
         BackColor       =   &H8000000D&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3120
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtUName 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         TabIndex        =   1
         ToolTipText     =   "User Name admin"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtUPass 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   2040
         PasswordChar    =   "*"
         TabIndex        =   2
         ToolTipText     =   "Password is jnawali"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "User Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   1455
      End
   End
   Begin VB.Label Label4 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   10
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label per 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   9
      Top             =   4200
      Width           =   495
   End
   Begin VB.Label complet 
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   720
      TabIndex        =   8
      Top             =   3840
      Width           =   3855
   End
End
Attribute VB_Name = "Form8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public str As String
Private Sub cmdexit_Click()
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim st As String
 Dim vbresult As VbMsgBoxResult
st = "select * FROM tblUser where Name like '" & txtUName.Text & "' and Password like '" & txtUPass.Text & "'"
CustomRecordSetOpen st
If res.EOF = True And res.BOF = True Then
    vbresult = MsgBox("Invalid user or password...", vbCritical)
    txtUName.Text = ""
    txtUPass.Text = ""
    txtUName.SetFocus
    CloseRecordSet
    Exit Sub
End If
    CloseRecordSet
    Form1.Show
    Unload Me
Exit Sub
End Sub

Private Sub Form_Load()
dbconnection
    Timer1.Interval = 50
    Timer1.Enabled = True
    str = "Loading Please Wait..."
   complet.Caption = str
End Sub
Private Sub Timer1_Timer()
ProgressBar1.value = ProgressBar1.value + 1
per.Caption = ProgressBar1.value
If ProgressBar1.value = 20 Then Timer1.Interval = 20
    'complet.Caption = "Loading Forms..."

If ProgressBar1.value = 45 Then Timer1.Interval = 1
    'complet.Caption = "Loading Database Connectivity..."

If ProgressBar1.value = 70 Then Timer1.Interval = 20
    'complet.Caption = "Loading Various Component..."

If ProgressBar1.value = 85 Then Timer1.Interval = 60
    'complet.Caption = "Completing Please Wait..."

If ProgressBar1.value = 100 Then
    complet.Caption = "Completed"
    Frame1.Visible = True
    Timer1.Enabled = False
    ProgressBar1.Visible = False
    per.Visible = False
    complet.Visible = False
    Label4.Visible = False
    txtUName.SetFocus
    Exit Sub
End If

End Sub



Private Sub txtUPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOk_Click
End If
End Sub
