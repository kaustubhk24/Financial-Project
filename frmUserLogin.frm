VERSION 5.00
Begin VB.Form Form9 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form9"
   ClientHeight    =   3135
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6210
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6210
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   2400
      Width           =   1575
   End
   Begin VB.CommandButton cmdOk 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtConfirm 
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
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtNew 
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
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   960
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtCurren 
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
      Left            =   2640
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Confirm Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0C0FF&
      Caption         =   "New Password"
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
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Current Password"
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
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1935
   End
End
Attribute VB_Name = "Form9"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdclose_Click()
Form1.Frame1.Visible = True
Unload Me
End Sub

Private Sub cmdOk_Click()
Dim st As String
Dim temp As String
st = "select name,password from tblUser"

CustomRecordSetOpen st
temp = res!Name
    If (txtNew.Text = txtConfirm.Text) Then
        res!Name = temp
        res!password = txtConfirm.Text
        res.UpdateBatch
        MsgBox "Successfully Edited..."
        txtCurren.Text = ""
        txtNew.Text = ""
        txtConfirm.Text = ""
        txtCurren.SetFocus
        'cmdOk.Visible = False
        cmdclose_Click
    Else
        MsgBox "Sorry Password Mismatch", vbCritical
        txtNew.Text = ""
        txtConfirm.Text = ""
        txtNew.SetFocus
    End If
CloseRecordSet
End Sub

Private Sub Form_Load()
dbconnection
End Sub

Private Sub txtConfirm_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdOk_Click
End If
End Sub

Private Sub txtCurren_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtCurren_LostFocus
End If
End Sub

Private Sub txtCurren_LostFocus()
    Dim st As String
    st = "select password from tblUser" 'where password like '" & txtCurren.Text & "'"
    CustomRecordSetOpen st
    If (txtCurren.Text = res!password) Then
    txtNew.Visible = True
    Label2.Visible = True
    Label3.Visible = True
    txtConfirm.Visible = True
    cmdOk.Visible = True
    txtNew.SetFocus
    Else
        
        MsgBox "sorry wrong password...", vbCritical
        txtCurren.Text = ""
        cmdClose.SetFocus
    End If
        CloseRecordSet
    'End If
    CloseRecordSet
            
End Sub
