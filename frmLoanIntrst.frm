VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form4 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Form4"
   ClientHeight    =   3615
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   6945
   Icon            =   "frmLoanIntrst.frx":0000
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3615
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Loan Intrest"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin VB.TextBox txtTotalInterest 
         BackColor       =   &H00FFFFFF&
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
         Left            =   4680
         TabIndex        =   14
         Top             =   2040
         Width           =   1695
      End
      Begin VB.TextBox txtTotalDays 
         BackColor       =   &H00FFFFFF&
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
         Left            =   4680
         TabIndex        =   12
         Top             =   1560
         Width           =   1695
      End
      Begin VB.TextBox txtIRate 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         TabIndex        =   10
         Top             =   1560
         Width           =   1695
      End
      Begin VB.ComboBox cmbloaninterest 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         TabIndex        =   9
         Text            =   "Combo1"
         Top             =   600
         Width           =   1695
      End
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   120
         Top             =   3360
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
         Left            =   3360
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2760
         Width           =   855
      End
      Begin VB.CommandButton cmdPay 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Receipt"
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
         Left            =   1920
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtTotal 
         BackColor       =   &H00FFFFFF&
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
         Left            =   1560
         TabIndex        =   6
         Top             =   2040
         Width           =   1695
      End
      Begin MSComCtl2.DTPicker dateloan 
         Height          =   375
         Left            =   1560
         TabIndex        =   3
         Top             =   1080
         Width           =   1695
         _ExtentX        =   2990
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
         CalendarTitleBackColor=   12640511
         CalendarTrailingForeColor=   255
         Format          =   63569921
         CurrentDate     =   39567
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Total Interest"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   13
         Top             =   2040
         Width           =   1155
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0FF&
         Caption         =   "Total Days"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   3480
         TabIndex        =   11
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Total Amount:"
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
         Top             =   2040
         Width           =   1335
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Interest Rate:"
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
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Date:"
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
         TabIndex        =   2
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Mmbr-Id:"
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
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim st As String
Dim str As String
Public mbr_date As Date



Private Sub cmbloaninterest_Click()
Dim a As String
a = cmbloaninterest.Text
st = "select * from tblLoanInterestReceipt where mmbr_id=" & a
CustomRecordSetOpen st
    'txttemploaninterest.Text = !Date
        txtTotalDays.Text = dateloan.value - res!Date
        
        If (Val(txtTotalDays.Text) <= 10) Then
        txtIRate.Text = "12"
        txttotal.Text = res!amount + (res!amount * 12 * (Val(txtTotalDays.Text) / 365)) / 100
        txtTotalInterest.Text = (Val(txttotal.Text) * 12 * (Val(txtTotalDays.Text) / 365)) / 100
        End If
        If (Val(txtTotalDays.Text) > 10 And (Val(txtTotalDays.Text) <= 20)) Then
        txtIRate.Text = "25"
        txttotal.Text = res!amount + (res!amount * 25 * (Val(txtTotalDays.Text) / 365)) / 100
        txtTotalInterest.Text = (Val(txttotal.Text) * 25 * (Val(txtTotalDays.Text) / 365)) / 100
        End If
        If (Val(txtTotalDays.Text) > 20) Then
        txtIRate.Text = "30"
        txttotal.Text = res!amount + (res!amount * 30 * (Val(txtTotalDays.Text) / 365)) / 100
        txtTotalInterest.Text = (Val(txttotal.Text) * 30 * (Val(txtTotalDays.Text) / 365)) / 100
        End If
CloseRecordSet
End Sub

Private Sub cmdclose_Click()
Form1.Frame1.Visible = True
Unload Me
End Sub

Private Sub cmdPay_Click()
Dim value As String
Dim b As String
value = InputBox("Inteest Receipt", "Interest Application", "Enter Receipt Interest")
    b = cmbloaninterest.Text
    st = "select mmbr_id,amount,date from tblLoanInterestReceipt where mmbr_id=" & b
    CustomRecordSetOpen st
    With res
        !mmbr_id = cmbloaninterest.Text
        !amount = Val(txttotal.Text) - Val(value)
        !Date = dateloan.value
        .UpdateBatch
    End With
    CloseRecordSet
    st = "select * from tblPlLoanInterest"
    CustomRecordSetOpen st
    With res
        .AddNew
        !mmbr_id = cmbloaninterest.Text
        !amount = value
        .UpdateBatch
    End With
    CloseRecordSet
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
            !amount = res!amount + Val(value)
            .UpdateBatch
        End With
    CloseRecordSet

    cmdPay.Enabled = False
End Sub

Private Sub Form_Load()
dbconnection
cmdPay.Enabled = True
'Dim a As String
'Dim b As String
'Dim c As String
st = "select distinct mmbr_id from loan"
    
    cmbloaninterest.Clear
    CustomRecordSetOpen st
    res.MoveFirst
    While Not res.EOF
        cmbloaninterest.AddItem res!mmbr_id
        res.MoveNext
    Wend
    CloseRecordSet
'    st = "select amount,date from tblLoanInterestReceipt"
'    CustomRecordSetOpen st
'    a = dateloan.Value
'    b = Val(a) - res!Date
'    If (Val(b) <= 10) Then
'        c = (res!amount * 12 * (Val(b) / 365)) / 100
'        res!amount = res!amount + Val(c)
'        End If
'        If (Val(b) > 10 And (Val(b) <= 20)) Then
'
'        c = (res!amount * 25 * (Val(b) / 365)) / 100
'        res!amount = res!amount + Val(c)
'        End If
'        If (Val(b) > 20) Then
'
'        txttotal.Text = (res!amount * 30 * (Val(b) / 365)) / 100
'        res!amount = res!amount + Val(c)
'        End If
'CloseRecordSet
End Sub

Private Sub Timer1_Timer()
dateloan = Date
End Sub
