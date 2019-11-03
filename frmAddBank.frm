VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form5 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Add New Bank"
   ClientHeight    =   4695
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   4320
   Icon            =   "frmAddBank.frx":0000
   LinkTopic       =   "Form5"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4695
   ScaleWidth      =   4320
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtAccNo 
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
      Left            =   1920
      TabIndex        =   6
      Top             =   2880
      Width           =   1935
   End
   Begin MSComCtl2.DTPicker date1 
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
      _ExtentX        =   3413
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
      Format          =   193855489
      CurrentDate     =   39567
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
      Left            =   1920
      TabIndex        =   7
      Top             =   3360
      Width           =   1935
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1920
      TabIndex        =   2
      Top             =   960
      Width           =   1935
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0FF&
      Caption         =   "Add New Bank"
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
      Height          =   4455
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.Timer Timer1 
         Interval        =   100
         Left            =   1200
         Top             =   1800
      End
      Begin VB.ComboBox combType 
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
         ItemData        =   "frmAddBank.frx":2A8B2
         Left            =   1800
         List            =   "frmAddBank.frx":2A8BC
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   2280
         Width           =   1935
      End
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H80000018&
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
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   3840
         Width           =   1095
      End
      Begin VB.CommandButton cmdSave 
         BackColor       =   &H80000018&
         Caption         =   "Save"
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
         Left            =   840
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3840
         Width           =   975
      End
      Begin VB.TextBox txtAddress 
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
         Left            =   1800
         TabIndex        =   3
         Top             =   1320
         Width           =   1935
      End
      Begin VB.TextBox txtId 
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
         Left            =   1800
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
      Begin VB.Label Label7 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Type:"
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
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   615
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   15
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label5 
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
         Left            =   240
         TabIndex        =   14
         Top             =   3240
         Width           =   1335
      End
      Begin VB.Label Label4 
         BackColor       =   &H00C0C0FF&
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
         Left            =   240
         TabIndex        =   13
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label Label3 
         BackColor       =   &H00C0C0FF&
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
         Left            =   240
         TabIndex        =   12
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0C0FF&
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
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Bank Id:"
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
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim st As String
Public str As String

Private Sub cmdclose_Click()
    Form1.Frame1.Visible = True
    Unload Me
End Sub

Private Sub cmdSave_Click()
    st = "tblBank"
    If txtName.Text = "" Or _
        txtAddress.Text = "" Or _
        txttotal.Text = "" Or _
        txtAccNo.Text = "" Or _
        combType.Text = "" Then
        MsgBox "Please Fill All Boxes...", vbExclamation
    Else
    CustomRecordSetOpen st
    With res
        .AddNew
        !Name = txtName.Text
        !address = txtAddress.Text
        !Date = date1.value
        !acc_no = txtAccNo.Text
        !Type = combType.Text
        !total = txttotal.Text
        .UpdateBatch
    End With
MsgBox "Successfully Added...", vbInformation
CloseRecordSet
Form1.Frame1.Visible = True
Unload Me
End If

st = "select amount from tblcash "
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        
        !amount = txttotal.Text
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    
   st = "select amount from tblcash "
    CustomRecordSetOpen st
       With res
       If combType.ListIndex = 0 Then
           
            !amount = !amount - Val(txttotal.Text)
            .UpdateBatch
        End If
        If combType.ListIndex = 1 Then
            !amount = !amount + Val(txttotal.Text)
            .UpdateBatch
        End If
        End With
    CloseRecordSet

End Sub

Private Sub Form_Load()
dbconnection

End Sub

Private Sub Timer1_Timer()
date1 = Date
End Sub
