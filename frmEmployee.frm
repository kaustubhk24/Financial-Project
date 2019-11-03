VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form2 
   BackColor       =   &H00C0C0FF&
   Caption         =   "Add Employee"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   8085
   Icon            =   "frmEmployee.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   8085
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog Dlgs 
      Left            =   1080
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   480
      Top             =   3000
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
      TabIndex        =   11
      Top             =   3720
      Width           =   855
   End
   Begin VB.CommandButton cmdAdd 
      BackColor       =   &H00FFC0C0&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1920
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   3720
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2415
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   2415
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   1800
      TabIndex        =   5
      Top             =   2400
      Width           =   2415
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5520
      TabIndex        =   6
      Top             =   120
      Width           =   2415
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
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   5520
      TabIndex        =   7
      Top             =   600
      Width           =   2415
   End
   Begin VB.OptionButton Option1 
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
      Left            =   1800
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.OptionButton Option2 
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
      Left            =   3120
      TabIndex        =   3
      Top             =   1320
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
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
      ItemData        =   "frmEmployee.frx":2A8B2
      Left            =   5520
      List            =   "frmEmployee.frx":2A8CB
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   1800
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   375
      Left            =   5520
      TabIndex        =   8
      Top             =   1200
      Width           =   2415
      _ExtentX        =   4260
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
      Format          =   109182977
      CurrentDate     =   39565
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1800
      Width           =   2415
      _ExtentX        =   4260
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
      Format          =   109182977
      CurrentDate     =   39565
   End
   Begin VB.Label Label11 
      BackColor       =   &H80000018&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Browse"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4560
      TabIndex        =   22
      Top             =   3120
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Photo"
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
      Left            =   4440
      TabIndex        =   21
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   1455
      Left            =   5880
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Emp-Id"
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
      Left            =   360
      TabIndex        =   20
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Left            =   360
      TabIndex        =   19
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
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
      Left            =   360
      TabIndex        =   18
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label Label3 
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
      Left            =   4440
      TabIndex        =   17
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label4 
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
      Left            =   4440
      TabIndex        =   16
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label5 
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
      Left            =   360
      TabIndex        =   15
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Birth"
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
      Left            =   360
      TabIndex        =   14
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Join Date"
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
      Left            =   4440
      TabIndex        =   13
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Department"
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
      Left            =   4440
      TabIndex        =   12
      Top             =   1800
      Width           =   1095
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public str As String
Public st As String
'Private Declare Function GetFullPathName Lib "kernel32" Alias "GetFullPathNameA" _
'    (ByVal lpFileName As String, ByVal nBufferLength As Long, ByVal lpBuffer As String, _
'    ByVal lpFilePart As String) As Long
'    Dim filePath As String * 255
'
'    Public rsf As FileSystemObject
'    Public fn As String
'    Public ft As String
'    Public name1 As String
'   'Public Image1 As String

Private Sub cmdAdd_Click()
    Dim st As String
    Dim vbresult As VbMsgBoxResult
    
    If txtName.Text = "" Or _
        (Option1.value = False And Option2.value = False) Or _
        DTPicker1.value = "" Or _
        DTPicker2.value = "" Or _
        txtAddress.Text = "" Or _
        txtPhone.Text = "" Or _
        txtCell.Text = "" Or _
        Combo1.Text = "" Then
        MsgBox "Plese Fill All the boxes..."
      Else
    st = "employee"
   ' If blnRecordSetOpen = True Then
    '    CloseRecordSet
    'Else
     '   CustomRecordSetOpen st
    'End If
    With res
    .AddNew
    !emp_id = txtId.Text
    !Name = txtName.Text
    If Option1.value = True Then !gender = "Male"
    If Option2.value = True Then !gender = "Female"
    !dob = DTPicker1.value
    !address = txtAddress.Text
    !phone = txtPhone.Text
    !cell = txtCell.Text
    !join_date = DTPicker2.value
    !department = Combo1.Text
    '!photo = ft
    .UpdateBatch
    
    vbresult = MsgBox("Info successfully Recorded...", vbExclamation)
    CloseRecordSet
    End With
    Form1.Frame1.Visible = True
    Unload Me
End If
End Sub

Private Sub cmdclose_Click()
    Form1.Frame1.Visible = True
    Unload Me
End Sub

Private Sub Form_unLoad(cancel As Integer)
    
    Form1.Frame1.Visible = True
    Unload Me
End Sub
Private Sub Form_Load()
   'DTPicker2 = FormatDateTime(Date, 1)
 
    dbconnection
    'Set rsf = New filesystemobject
    blnRecordSetOpen = False
    st = "employee"
    If blnRecordSetOpen = True Then
        CloseRecordSet
    Else
        CustomRecordSetOpen st
    End If
    If res.RecordCount = 0 Then
    txtId.Text = "1"
    Else
    res.MoveLast
    txtId.Text = res!emp_id + 1
    End If
End Sub
'Private Sub Label11_Click()
'Dim retvalue As Long
'Dim fileName As String
'Dim dest As String
'With Dlgs
'        .Filter = "(*.bmp;*.jpg;*.gif;*.dat;*.pcx)| *.bmp;*.jpg;*.gif;*.dat;*.pcx|(*.psd)|*.psd|(*.All files)|*.*"
'        .ShowOpen
'        If .fileName <> "" Then fileName = .fileName
'        retvalue = GetFullPathName(fileName, 255, filePath, 0)
'        If .fileName = "" Then Exit Sub
'        fn = .fileName
'        ft = .FileTitle
'        dest = (App.Path & "\images\")
'        MsgBox dest
'        MsgBox fn
'        MsgBox filePath 'path of the source file<<<----source
'        Set rsf = New FileSystemObject
'        rsf.CopyFile fn, dest
''         rsf.CopyFile fn, dest & txtName.Text & ".jpg"
'
'Image1.Picture = LoadPicture(fn)
'
'End With
'End Sub

Private Sub Timer1_Timer()
DTPicker2 = Date
End Sub
