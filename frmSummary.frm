VERSION 5.00
Begin VB.Form Form7 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form7"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   5715
   Icon            =   "frmSummary.frx":0000
   LinkTopic       =   "Form7"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   5715
   Begin VB.CommandButton Command1 
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
      Left            =   2160
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4680
      Width           =   1335
   End
   Begin VB.TextBox txtPL 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2640
      TabIndex        =   13
      Top             =   3960
      Width           =   2415
   End
   Begin VB.TextBox txtBLiab 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2640
      TabIndex        =   12
      Top             =   3360
      Width           =   2415
   End
   Begin VB.TextBox txtBAss 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2640
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox txtCash 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2640
      TabIndex        =   10
      Top             =   2160
      Width           =   2415
   End
   Begin VB.TextBox txtCapital 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2640
      TabIndex        =   9
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox txtMem 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2640
      TabIndex        =   8
      Top             =   960
      Width           =   2415
   End
   Begin VB.TextBox txtEmp 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000009&
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Bank Liabilities"
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
      Left            =   360
      TabIndex        =   6
      Top             =   3480
      Width           =   2160
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Bank Amount"
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
      Left            =   360
      TabIndex        =   5
      Top             =   2880
      Width           =   2040
   End
   Begin VB.Label lblPL 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
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
      Left            =   360
      TabIndex        =   4
      Top             =   4080
      Width           =   1980
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Cash"
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
      Left            =   360
      TabIndex        =   3
      Top             =   2280
      Width           =   1140
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Capital"
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
      Left            =   360
      TabIndex        =   2
      Top             =   1680
      Width           =   1320
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Employee"
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
      Left            =   360
      TabIndex        =   1
      Top             =   480
      Width           =   1635
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0FF&
      Caption         =   "Total Member"
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
      Left            =   360
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Frame1.Visible = True
Unload Me
End Sub


Private Sub Form_Load()
dbconnection
Dim st As String
st = "select count(emp_id) as tot from employee"
CustomRecordSetOpen st
    txtEmp.Text = res!tot
CloseRecordSet
st = "select count(mmbr_id) as tot from member"
CustomRecordSetOpen st
    txtMem.Text = res!tot
CloseRecordSet
st = "select sum(amount) as tot from tblcapital"
CustomRecordSetOpen st
    txtCapital.Text = res!tot
CloseRecordSet

st = "select amount from tblcash"
CustomRecordSetOpen st
    txtCash.Text = res!amount
CloseRecordSet
Dim a As Double
st = "select sum(total) as tot from tblBank where type='Assets'"
CustomRecordSetOpen st
    a = res!tot
CloseRecordSet
st = "select sum(amount) as tot from tblbankdeposit"
CustomRecordSetOpen st
    txtBAss.Text = a + res!tot
CloseRecordSet
a = 0
st = "select sum(total) as tot from tblBank where type='Liabilities'"
CustomRecordSetOpen st
    a = res!tot
CloseRecordSet
st = "select sum(amount) as tot from tblbankwithdraw"
CustomRecordSetOpen st
    txtBLiab.Text = a + res!tot
CloseRecordSet
st = "select * from tblPL"
CustomRecordSetOpen st
    lblpl.Caption = res!particular
    txtpl.Text = res!amount
CloseRecordSet
End Sub
