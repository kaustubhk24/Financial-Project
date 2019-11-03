VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form3 
   BackColor       =   &H00C0C0FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Member Addition"
   ClientHeight    =   3465
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   7935
   ClipControls    =   0   'False
   Icon            =   "frmMember.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2391.604
   ScaleMode       =   0  'User
   ScaleWidth      =   7451.376
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   7080
      Top             =   2280
   End
   Begin VB.TextBox txtEntryFee 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
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
      Left            =   4920
      TabIndex        =   20
      Top             =   2400
      Width           =   1815
   End
   Begin VB.ComboBox combAccType 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "frmMember.frx":2A8B2
      Left            =   4920
      List            =   "frmMember.frx":2A8BF
      Style           =   2  'Dropdown List
      TabIndex        =   18
      Top             =   1320
      Width           =   1815
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtId 
      BackColor       =   &H00C0E0FF&
      Enabled         =   0   'False
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
      Left            =   1320
      TabIndex        =   8
      Top             =   120
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00C0E0FF&
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
      Left            =   1320
      TabIndex        =   10
      Top             =   720
      Width           =   2055
   End
   Begin VB.OptionButton optMale 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Male"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1320
      TabIndex        =   11
      Top             =   1320
      Width           =   855
   End
   Begin VB.OptionButton optFemale 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Female"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2280
      TabIndex        =   12
      Top             =   1320
      Width           =   1095
   End
   Begin VB.TextBox txtAddress 
      BackColor       =   &H00C0E0FF&
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
      Left            =   1320
      TabIndex        =   13
      Top             =   1800
      Width           =   2055
   End
   Begin VB.TextBox txtPhone 
      BackColor       =   &H00C0E0FF&
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
      Left            =   1320
      TabIndex        =   15
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox txtCell 
      BackColor       =   &H00C0E0FF&
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
      Left            =   4920
      TabIndex        =   16
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox txtAccNo 
      BackColor       =   &H00C0E0FF&
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
      Left            =   4920
      TabIndex        =   19
      Top             =   1800
      Width           =   1815
   End
   Begin MSComCtl2.DTPicker dateDoj 
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CalendarBackColor=   12640511
      Format          =   193986561
      CurrentDate     =   39565
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Entry Fee:"
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
      Left            =   3600
      TabIndex        =   14
      Top             =   2400
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mmbr-Id"
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
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
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
      Left            =   120
      TabIndex        =   7
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Gender:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Address:"
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
      Left            =   120
      TabIndex        =   5
      Top             =   1800
      Width           =   855
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone No:"
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
      Left            =   120
      TabIndex        =   4
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cell No:"
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
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Join:"
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
      Left            =   3600
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Type :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Acc-No:"
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
      Left            =   3600
      TabIndex        =   0
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim st As String
Public str As String

Private Sub cmdAdd_Click()
    st = "member"
    If txtId.Text = "" Or _
        txtAccNo.Text = "" Or _
        txtAddress.Text = "" Or _
        (optFemale.value = False And optMale.value = False) Or _
        txtPhone.Text = "" Or _
        txtCell.Text = "" Or _
        combAccType.Text = "" Or _
        txtEntryFee.Text = "" Then
        MsgBox "Fill all Boxes...", vbInformation
    Else
    CustomRecordSetOpen st
    With res
        .AddNew
        !mmbr_id = txtId.Text
        !Name = txtName.Text
        If optMale.value = True Then !gender = "Male"
        If optFemale.value = True Then !gender = "Female"
        !address = txtAddress.Text
        !phone = txtPhone.Text
        !cell = txtCell.Text
        !join_date = dateDoj.value
        !Type = combAccType.Text
        !acc_no = txtAccNo.Text
        !fee = txtEntryFee.Text
        .UpdateBatch
    End With
    MsgBox "Member Successfully Recorded...", vbInformation
    CloseRecordSet
    Form1.Frame1.Visible = True
    Unload Me
End If
End Sub

Private Sub cmdclose_Click()
     Form1.Frame1.Visible = True
    Unload Me
End Sub
Private Sub combAccType_LostFocus()
    If combAccType.Text = "Saver" Then txtAccNo.Text = "SV-"
    If combAccType.Text = "Share Holder" Then txtAccNo.Text = "SH-"
    If combAccType.Text = "Debtor" Then txtAccNo.Text = "DB-"
End Sub

Private Sub Form_Load()
    dbconnection
    
    blnRecordSetOpen = False
    st = "member"
    If blnRecordSetOpen = True Then
        CloseRecordSet
    Else
        CustomRecordSetOpen st
    End If
    If res.RecordCount = 0 Then
    txtId.Text = "1"
    Else
    res.MoveLast
    txtId.Text = res!mmbr_id + 1
    End If
    CloseRecordSet
End Sub

Private Sub Form_unLoad(cancel As Integer)
    Form1.Frame1.Visible = True
    Unload Me
End Sub

Private Sub Timer1_Timer()
dateDoj = Date
End Sub

Private Sub txtAccNo_LostFocus()
If combAccType.Text = "Saver" Then txtEntryFee.Text = "200"
If combAccType.Text = "Share Holder" Then txtEntryFee.Text = "300"
If combAccType.Text = "Debtor" Then txtEntryFee.Text = "500"
End Sub
