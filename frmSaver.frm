VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form6 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form6"
   ClientHeight    =   4230
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6540
   Icon            =   "frmSaver.frx":0000
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4230
   ScaleWidth      =   6540
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000013&
      Caption         =   "Payment"
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
      Left            =   1800
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdclose 
      BackColor       =   &H80000013&
      Caption         =   "Close"
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
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   4920
      Top             =   2160
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   4920
      TabIndex        =   8
      Top             =   1560
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      _Version        =   393216
      Format          =   20709377
      CurrentDate     =   39643
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   7
      Top             =   1560
      Width           =   2055
   End
   Begin VB.TextBox txtinterest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   6
      Top             =   2160
      Width           =   2055
   End
   Begin VB.TextBox txttotal 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   4
      Top             =   2760
      Width           =   2055
   End
   Begin VB.TextBox txtmid 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   960
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Last Visited Days"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   10
      Top             =   1560
      Width           =   2430
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Interest Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   1725
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Amount"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   3
      Top             =   2760
      Width           =   1425
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Member Id"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "SAVER DETAIL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdclose_Click()
Form1.Frame1.Visible = True
Unload Me
End Sub

Private Sub cmdsearch_Click()
Dim a As String
Dim b As String
Dim sqldate As String
a = txtmid.Text
Dim st As String
st = "select amount from tblSaverInterest where mmbr_id=" & a
sqldate = "select date,mmbr_id,amount from tblSaverInterest where mmbr_id=" & a
CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    MsgBox "Sorry Not Found" '& vbCrLf & "hello"
        txtmid.Text = ""
        txtmid.SetFocus
        CloseRecordSet
    Exit Sub
    End If
    With res
        CustomRecordSetOpen sqldate
        Text1.Text = DTPicker1.value - !Date
        b = res!amount
        CloseRecordSet
        st = "select amount,date,mmbr_id from tblSaverInterest where mmbr_id like '" & txtmid.Text & "'"
        CustomRecordSetOpen st
        
        txtinterest.Text = Val((b) * 1.5 * Val((Text1.Text) / 365)) / 100
        txttotal.Text = Val(b) + Val(txtinterest.Text)
        CloseRecordSet
        txtmid.SetFocus
End With

End Sub

Private Sub Command1_Click()
Dim value As String
Dim b As String
Dim st As String
value = InputBox("Amount Payment", "Saver Amount Application", "Enter Paid Amount")
    b = txtmid.Text
    st = "select mmbr_id,amount,date from tblSaverInterest where mmbr_id=" & b
    CustomRecordSetOpen st
    With res
        '!mmbr_id = txtmid.Text
       ' !amount = Val(txttotal.Text) - Val(value)
        '!Date = DTPicker1.value
        MsgBox ("Details submitted")
        .UpdateBatch
    End With
    CloseRecordSet
    
    st = "select * from tblPlSaverInterest"
    CustomRecordSetOpen st
    With res
        .AddNew
        !mmbr_id = txtmid.Text
        !interest = value
        .UpdateBatch
    End With
    CloseRecordSet
    Command1.Enabled = False
    st = "select amount from tblcash "
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !amount = value
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from tblcash "
    CustomRecordSetOpen st
        With res
            !amount = res!amount - Val(value)
            .UpdateBatch
        End With
    CloseRecordSet

End Sub

Private Sub Form_Load()
dbconnection

End Sub

Private Sub Timer1_Timer()
DTPicker1 = Date
End Sub

Private Sub txtmid_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdsearch_Click
End If
End Sub
