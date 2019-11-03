VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000009&
   Caption         =   " State Bank of Travancore"
   ClientHeight    =   9105
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   14295
   DrawStyle       =   1  'Dash
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   9105
   ScaleWidth      =   14295
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H000080FF&
      Caption         =   "Exit"
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
      Left            =   7560
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   8640
      Width           =   1215
   End
   Begin VB.CommandButton cmdAbout 
      BackColor       =   &H000080FF&
      Caption         =   "About"
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
      Left            =   5280
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8640
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Financial Management System"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   7575
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   13215
      Begin TabDlg.SSTab SSTab1 
         Height          =   7215
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   13095
         _ExtentX        =   23098
         _ExtentY        =   12726
         _Version        =   393216
         Style           =   1
         Tabs            =   10
         TabsPerRow      =   9
         TabHeight       =   520
         BackColor       =   255
         ForeColor       =   16711680
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Employee"
         TabPicture(0)   =   "frmMain.frx":2A8B2
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "cmdEdit"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdView"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdSave"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "cmdDelete"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "cmdAdd"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "Frame2"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).ControlCount=   6
         TabCaption(1)   =   "Member"
         TabPicture(1)   =   "frmMain.frx":2A8CE
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "Frame3"
         Tab(1).Control(1)=   "cmdCEdit"
         Tab(1).Control(2)=   "cmdCView"
         Tab(1).Control(3)=   "cmdCSave"
         Tab(1).Control(4)=   "cmdCDelete"
         Tab(1).Control(5)=   "cmdCAdd"
         Tab(1).ControlCount=   6
         TabCaption(2)   =   "Transaction"
         TabPicture(2)   =   "frmMain.frx":2A8EA
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "Frame4"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "Expenditure"
         TabPicture(3)   =   "frmMain.frx":2A906
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "flexcaption"
         Tab(3).Control(1)=   "Frame8"
         Tab(3).Control(2)=   "MSFlexGrid1"
         Tab(3).ControlCount=   3
         TabCaption(4)   =   "Income"
         TabPicture(4)   =   "frmMain.frx":2A922
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "flexIncome"
         Tab(4).Control(1)=   "Frame9"
         Tab(4).Control(2)=   "lblIncomeflex"
         Tab(4).ControlCount=   3
         TabCaption(5)   =   "Contra"
         TabPicture(5)   =   "frmMain.frx":2A93E
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "Frame10"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "Day Book"
         TabPicture(6)   =   "frmMain.frx":2A95A
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "DTPicker1"
         Tab(6).Control(1)=   "cmdSearchDaybook"
         Tab(6).Control(2)=   "Timer4"
         Tab(6).Control(3)=   "DTPickerdaybook"
         Tab(6).Control(4)=   "MSFlexGrid2"
         Tab(6).ControlCount=   5
         TabCaption(7)   =   "P/ L Acc"
         TabPicture(7)   =   "frmMain.frx":2A976
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "txtpl"
         Tab(7).Control(1)=   "txtPlInInterest"
         Tab(7).Control(2)=   "txtPlOI"
         Tab(7).Control(3)=   "txtPlSI"
         Tab(7).Control(4)=   "txtPlII"
         Tab(7).Control(5)=   "txtPlDIn"
         Tab(7).Control(6)=   "txtPlExInterest"
         Tab(7).Control(7)=   "txtplpam"
         Tab(7).Control(8)=   "txtpldx"
         Tab(7).Control(9)=   "txtplex"
         Tab(7).Control(10)=   "txttemppl"
         Tab(7).Control(11)=   "Label81"
         Tab(7).Control(12)=   "Label80"
         Tab(7).Control(13)=   "Line5"
         Tab(7).Control(14)=   "Line4"
         Tab(7).Control(15)=   "Line3"
         Tab(7).Control(16)=   "lblpl"
         Tab(7).Control(17)=   "Label79"
         Tab(7).Control(18)=   "Label78"
         Tab(7).Control(19)=   "Label77"
         Tab(7).Control(20)=   "Label64"
         Tab(7).Control(21)=   "Label63"
         Tab(7).Control(22)=   "I"
         Tab(7).Control(23)=   "Label58"
         Tab(7).Control(24)=   "Label46"
         Tab(7).Control(25)=   "Label43"
         Tab(7).Control(26)=   "Line2"
         Tab(7).ControlCount=   27
         TabCaption(8)   =   "Summary          "
         TabPicture(8)   =   "frmMain.frx":2A992
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "Command1"
         Tab(8).ControlCount=   1
         TabCaption(9)   =   "About"
         TabPicture(9)   =   "frmMain.frx":2A9AE
         Tab(9).ControlEnabled=   0   'False
         Tab(9).Control(0)=   "Label82"
         Tab(9).ControlCount=   1
         Begin VB.CommandButton Command1 
            BackColor       =   &H000080FF&
            Caption         =   "Get Summary"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   -69960
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   225
            Top             =   3420
            Width           =   1935
         End
         Begin VB.TextBox txtpl 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   495
            Left            =   -71880
            TabIndex        =   222
            Top             =   6360
            Width           =   2415
         End
         Begin VB.TextBox txtPlInInterest 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   495
            Left            =   -66360
            TabIndex        =   220
            Top             =   5160
            Width           =   2295
         End
         Begin VB.TextBox txtPlOI 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   495
            Left            =   -66360
            TabIndex        =   219
            Top             =   4320
            Width           =   2295
         End
         Begin VB.TextBox txtPlSI 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   495
            Left            =   -66360
            TabIndex        =   218
            Top             =   3480
            Width           =   2295
         End
         Begin VB.TextBox txtPlII 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   495
            Left            =   -66360
            TabIndex        =   217
            Top             =   2640
            Width           =   2295
         End
         Begin VB.TextBox txtPlDIn 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   495
            Left            =   -66360
            TabIndex        =   216
            Top             =   1920
            Width           =   2295
         End
         Begin VB.TextBox txtPlExInterest 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   495
            Left            =   -72000
            TabIndex        =   210
            Top             =   4320
            Width           =   2175
         End
         Begin VB.TextBox txtplpam 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   495
            Left            =   -72000
            TabIndex        =   208
            Top             =   3480
            Width           =   2175
         End
         Begin VB.TextBox txtpldx 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   495
            Left            =   -72000
            TabIndex        =   206
            Top             =   2640
            Width           =   2175
         End
         Begin VB.TextBox txtplex 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
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
            Height          =   495
            Left            =   -72000
            TabIndex        =   204
            Top             =   1920
            Width           =   2175
         End
         Begin VB.TextBox txttemppl 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71280
            TabIndex        =   202
            Top             =   5640
            Visible         =   0   'False
            Width           =   1695
         End
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   375
            Left            =   -67200
            TabIndex        =   201
            Top             =   1080
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
            Format          =   109182977
            CurrentDate     =   39644
         End
         Begin VB.CommandButton cmdSearchDaybook 
            BackColor       =   &H80000018&
            Caption         =   "Search By Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65280
            Style           =   1  'Graphical
            TabIndex        =   200
            Top             =   1080
            Width           =   1815
         End
         Begin VB.Timer Timer4 
            Interval        =   100
            Left            =   -68640
            Top             =   6000
         End
         Begin MSComCtl2.DTPicker DTPickerdaybook 
            Height          =   375
            Left            =   -73440
            TabIndex        =   199
            Top             =   7140
            Visible         =   0   'False
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
            _Version        =   393216
            Format          =   109182977
            CurrentDate     =   39644
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid2 
            Height          =   5295
            Left            =   -74640
            TabIndex        =   198
            Top             =   1560
            Width           =   11295
            _ExtentX        =   19923
            _ExtentY        =   9340
            _Version        =   393216
            ScrollTrack     =   -1  'True
            TextStyle       =   3
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid flexIncome 
            Height          =   3015
            Left            =   -74640
            TabIndex        =   182
            Top             =   3720
            Width           =   9975
            _ExtentX        =   17595
            _ExtentY        =   5318
            _Version        =   393216
            BackColor       =   -2147483633
            TextStyle       =   3
            AllowUserResizing=   3
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
            Height          =   3375
            Left            =   -74640
            TabIndex        =   161
            Top             =   3480
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   5953
            _Version        =   393216
            BackColor       =   -2147483633
            TextStyle       =   3
            AllowUserResizing=   1
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin VB.Frame Frame2 
            Caption         =   "Employee Information"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   3615
            Left            =   600
            TabIndex        =   114
            Top             =   1440
            Width           =   9615
            Begin VB.Timer Timer3 
               Interval        =   100
               Left            =   6120
               Top             =   1440
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
               Left            =   3120
               TabIndex        =   157
               Top             =   360
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
               Left            =   3120
               TabIndex        =   123
               Top             =   1800
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
               Left            =   3120
               TabIndex        =   122
               Top             =   2400
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
               Left            =   3120
               TabIndex        =   121
               Top             =   3000
               Width           =   2415
            End
            Begin VB.OptionButton Option1 
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
               Left            =   3120
               TabIndex        =   120
               Top             =   840
               Width           =   975
            End
            Begin VB.OptionButton Option2 
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
               Left            =   4440
               TabIndex        =   119
               Top             =   840
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
               ItemData        =   "frmMain.frx":2A9CA
               Left            =   7320
               List            =   "frmMain.frx":2A9E3
               Style           =   2  'Dropdown List
               TabIndex        =   116
               Top             =   840
               Width           =   2055
            End
            Begin VB.ListBox List 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00FF0000&
               Height          =   2430
               ItemData        =   "frmMain.frx":2AA27
               Left            =   120
               List            =   "frmMain.frx":2AA29
               TabIndex        =   115
               Top             =   720
               Width           =   1575
            End
            Begin MSComCtl2.DTPicker DTPicker2 
               Height          =   375
               Left            =   7320
               TabIndex        =   117
               Top             =   360
               Width           =   2055
               _ExtentX        =   3625
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
            Begin MSComCtl2.DTPicker DTPickeremp 
               Height          =   375
               Left            =   3120
               TabIndex        =   118
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
               CalendarBackColor=   12640511
               Format          =   109182977
               CurrentDate     =   39565
            End
            Begin VB.Label Label1 
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
               Left            =   1920
               TabIndex        =   132
               Top             =   360
               Width           =   855
            End
            Begin VB.Label Label2 
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
               Left            =   1920
               TabIndex        =   131
               Top             =   1800
               Width           =   855
            End
            Begin VB.Label Label3 
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
               Left            =   1920
               TabIndex        =   130
               Top             =   2400
               Width           =   975
            End
            Begin VB.Label Label4 
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
               Left            =   1920
               TabIndex        =   129
               Top             =   3000
               Width           =   855
            End
            Begin VB.Label Label5 
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
               Left            =   1920
               TabIndex        =   128
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label6 
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
               Left            =   1920
               TabIndex        =   127
               Top             =   1320
               Width           =   1215
            End
            Begin VB.Label Label7 
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
               Left            =   6000
               TabIndex        =   126
               Top             =   360
               Width           =   1095
            End
            Begin VB.Label Label8 
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
               Left            =   6000
               TabIndex        =   125
               Top             =   840
               Width           =   1095
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               Caption         =   "Employee ID"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   480
               TabIndex        =   124
               Top             =   480
               Width           =   1080
            End
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H000080FF&
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
            Left            =   3600
            Style           =   1  'Graphical
            TabIndex        =   113
            Top             =   6120
            Width           =   1215
         End
         Begin VB.CommandButton cmdDelete 
            BackColor       =   &H000080FF&
            Caption         =   "Delete"
            Enabled         =   0   'False
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
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   112
            Top             =   6120
            Width           =   1215
         End
         Begin VB.CommandButton cmdSave 
            BackColor       =   &H000080FF&
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
            Left            =   8280
            Style           =   1  'Graphical
            TabIndex        =   111
            Top             =   6120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdView 
            BackColor       =   &H000080FF&
            Caption         =   "View"
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
            TabIndex        =   110
            Top             =   6120
            Width           =   1215
         End
         Begin VB.CommandButton cmdEdit 
            BackColor       =   &H000080FF&
            Caption         =   "Edit"
            Enabled         =   0   'False
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
            Left            =   6720
            Style           =   1  'Graphical
            TabIndex        =   109
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Frame Frame3 
            Caption         =   "Member Information"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   3495
            Left            =   -74520
            TabIndex        =   90
            Top             =   1440
            Width           =   9615
            Begin VB.ComboBox combCType 
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
               ItemData        =   "frmMain.frx":2AA2B
               Left            =   7080
               List            =   "frmMain.frx":2AA38
               Style           =   2  'Dropdown List
               TabIndex        =   158
               Top             =   840
               Width           =   2295
            End
            Begin VB.TextBox txtCName 
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
               Left            =   3000
               TabIndex        =   99
               Top             =   360
               Width           =   2535
            End
            Begin VB.OptionButton optCMale 
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
               Left            =   3000
               TabIndex        =   98
               Top             =   840
               Width           =   1095
            End
            Begin VB.OptionButton optCFemale 
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
               Left            =   4320
               TabIndex        =   97
               Top             =   840
               Width           =   1215
            End
            Begin VB.TextBox txtCAddress 
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
               Left            =   3000
               TabIndex        =   96
               Top             =   1320
               Width           =   2535
            End
            Begin VB.TextBox txtCPhone 
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
               Left            =   3000
               TabIndex        =   95
               Top             =   1920
               Width           =   2535
            End
            Begin VB.TextBox txtCCell 
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
               Left            =   3000
               TabIndex        =   94
               Top             =   2520
               Width           =   2535
            End
            Begin VB.TextBox txtCAccno 
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
               Left            =   7080
               TabIndex        =   92
               Top             =   1320
               Width           =   2295
            End
            Begin VB.ListBox ListM 
               Appearance      =   0  'Flat
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
               ForeColor       =   &H00FF0000&
               Height          =   2190
               ItemData        =   "frmMain.frx":2AA59
               Left            =   120
               List            =   "frmMain.frx":2AA5B
               TabIndex        =   91
               Top             =   600
               Width           =   1455
            End
            Begin MSComCtl2.DTPicker dateCDoj 
               Height          =   375
               Left            =   7080
               TabIndex        =   93
               Top             =   360
               Width           =   2295
               _ExtentX        =   4048
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
            Begin VB.Label Label9 
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
               Left            =   1920
               TabIndex        =   108
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label10 
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
               Left            =   1920
               TabIndex        =   107
               Top             =   840
               Width           =   735
            End
            Begin VB.Label Label11 
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
               Left            =   1920
               TabIndex        =   106
               Top             =   1320
               Width           =   855
            End
            Begin VB.Label Label12 
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
               Left            =   1920
               TabIndex        =   105
               Top             =   1920
               Width           =   975
            End
            Begin VB.Label Label13 
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
               Left            =   1920
               TabIndex        =   104
               Top             =   2520
               Width           =   975
            End
            Begin VB.Label Label14 
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
               Left            =   5760
               TabIndex        =   103
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label15 
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
               Left            =   5760
               TabIndex        =   102
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label16 
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
               Left            =   5760
               TabIndex        =   101
               Top             =   1320
               Width           =   975
            End
            Begin VB.Label Label60 
               Caption         =   "Member Id:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   360
               TabIndex        =   100
               Top             =   360
               Width           =   975
            End
         End
         Begin VB.CommandButton cmdCEdit 
            BackColor       =   &H000080FF&
            Caption         =   "Edit"
            Enabled         =   0   'False
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
            Left            =   -68040
            Style           =   1  'Graphical
            TabIndex        =   89
            Top             =   6120
            Width           =   1215
         End
         Begin VB.CommandButton cmdCView 
            BackColor       =   &H000080FF&
            Caption         =   "View"
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
            Left            =   -72600
            Style           =   1  'Graphical
            TabIndex        =   88
            Top             =   6120
            Width           =   1215
         End
         Begin VB.CommandButton cmdCSave 
            BackColor       =   &H000080FF&
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
            Left            =   -66600
            Style           =   1  'Graphical
            TabIndex        =   87
            Top             =   6120
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CommandButton cmdCDelete 
            BackColor       =   &H000080FF&
            Caption         =   "Delete"
            Enabled         =   0   'False
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
            Left            =   -69480
            Style           =   1  'Graphical
            TabIndex        =   86
            Top             =   6120
            Width           =   1215
         End
         Begin VB.CommandButton cmdCAdd 
            BackColor       =   &H000080FF&
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
            Left            =   -70920
            Style           =   1  'Graphical
            TabIndex        =   85
            Top             =   6120
            Width           =   1215
         End
         Begin VB.Frame Frame4 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   5895
            Left            =   -74760
            TabIndex        =   62
            Top             =   960
            Width           =   11655
            Begin VB.Frame Frame12 
               Caption         =   "Capital "
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   3015
               Left            =   7920
               TabIndex        =   184
               Top             =   2760
               Width           =   3615
               Begin VB.CommandButton cmdTCapitalSave 
                  BackColor       =   &H000080FF&
                  Caption         =   "Save"
                  BeginProperty Font 
                     Name            =   "Times New Roman"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   495
                  Left            =   1440
                  Style           =   1  'Graphical
                  TabIndex        =   187
                  Top             =   1920
                  Width           =   1215
               End
               Begin VB.TextBox txtTCapitalAmount 
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
                  Left            =   1680
                  TabIndex        =   186
                  Top             =   1080
                  Width           =   1695
               End
               Begin VB.Label Label38 
                  AutoSize        =   -1  'True
                  Caption         =   "Capital Amount:"
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
                  Left            =   240
                  TabIndex        =   185
                  Top             =   1080
                  Width           =   1395
               End
            End
            Begin TabDlg.SSTab SSTab3 
               Height          =   2895
               Left            =   240
               TabIndex        =   133
               Top             =   2880
               Width           =   7455
               _ExtentX        =   13150
               _ExtentY        =   5106
               _Version        =   393216
               Tabs            =   2
               Tab             =   1
               TabsPerRow      =   2
               TabHeight       =   520
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Expenditure"
               TabPicture(0)   =   "frmMain.frx":2AA5D
               Tab(0).ControlEnabled=   0   'False
               Tab(0).Control(0)=   "txtTExId"
               Tab(0).Control(1)=   "Timer1"
               Tab(0).Control(2)=   "combTExBid"
               Tab(0).Control(3)=   "combTExUnderBy"
               Tab(0).Control(4)=   "cmdTExReset"
               Tab(0).Control(5)=   "cmdTExSave"
               Tab(0).Control(6)=   "txtTExCheque"
               Tab(0).Control(7)=   "chkTExCheque"
               Tab(0).Control(8)=   "txtTExCash"
               Tab(0).Control(9)=   "chkTExCash"
               Tab(0).Control(10)=   "txtTExTotal"
               Tab(0).Control(11)=   "txtTExParticular"
               Tab(0).Control(12)=   "dateTEx"
               Tab(0).Control(13)=   "lblBankid"
               Tab(0).Control(14)=   "Label69"
               Tab(0).Control(15)=   "Label68"
               Tab(0).Control(16)=   "Label67"
               Tab(0).Control(17)=   "Label66"
               Tab(0).Control(18)=   "Label65"
               Tab(0).Control(19)=   "Label62"
               Tab(0).ControlCount=   20
               TabCaption(1)   =   "Income"
               TabPicture(1)   =   "frmMain.frx":2AA79
               Tab(1).ControlEnabled=   -1  'True
               Tab(1).Control(0)=   "Label70"
               Tab(1).Control(0).Enabled=   0   'False
               Tab(1).Control(1)=   "Label71"
               Tab(1).Control(1).Enabled=   0   'False
               Tab(1).Control(2)=   "Label72"
               Tab(1).Control(2).Enabled=   0   'False
               Tab(1).Control(3)=   "Label73"
               Tab(1).Control(3).Enabled=   0   'False
               Tab(1).Control(4)=   "Label74"
               Tab(1).Control(4).Enabled=   0   'False
               Tab(1).Control(5)=   "Label75"
               Tab(1).Control(5).Enabled=   0   'False
               Tab(1).Control(6)=   "lblTInBid"
               Tab(1).Control(6).Enabled=   0   'False
               Tab(1).Control(7)=   "txtTInParticular"
               Tab(1).Control(7).Enabled=   0   'False
               Tab(1).Control(8)=   "dateTIn"
               Tab(1).Control(8).Enabled=   0   'False
               Tab(1).Control(9)=   "combTInUnderBy"
               Tab(1).Control(9).Enabled=   0   'False
               Tab(1).Control(10)=   "checkTincash"
               Tab(1).Control(10).Enabled=   0   'False
               Tab(1).Control(11)=   "txtTIncash"
               Tab(1).Control(11).Enabled=   0   'False
               Tab(1).Control(12)=   "CheckTIncheque"
               Tab(1).Control(12).Enabled=   0   'False
               Tab(1).Control(13)=   "txtTIncheque"
               Tab(1).Control(13).Enabled=   0   'False
               Tab(1).Control(14)=   "txtTInTotal"
               Tab(1).Control(14).Enabled=   0   'False
               Tab(1).Control(15)=   "cmdTInSave"
               Tab(1).Control(15).Enabled=   0   'False
               Tab(1).Control(16)=   "cmdTInReset"
               Tab(1).Control(16).Enabled=   0   'False
               Tab(1).Control(17)=   "combTInEid"
               Tab(1).Control(17).Enabled=   0   'False
               Tab(1).Control(18)=   "combTInBid"
               Tab(1).Control(18).Enabled=   0   'False
               Tab(1).Control(19)=   "Timer2"
               Tab(1).Control(19).Enabled=   0   'False
               Tab(1).ControlCount=   20
               Begin VB.ComboBox txtTExId 
                  BackColor       =   &H00C0E0FF&
                  Height          =   420
                  Left            =   -73800
                  Style           =   2  'Dropdown List
                  TabIndex        =   226
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.Timer Timer2 
                  Interval        =   100
                  Left            =   3240
                  Top             =   1440
               End
               Begin VB.Timer Timer1 
                  Interval        =   100
                  Left            =   -71760
                  Top             =   960
               End
               Begin VB.ComboBox combTInBid 
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   6120
                  Style           =   2  'Dropdown List
                  TabIndex        =   175
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.ComboBox combTInEid 
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  Left            =   1200
                  Style           =   2  'Dropdown List
                  TabIndex        =   167
                  Top             =   480
                  Width           =   1935
               End
               Begin VB.ComboBox combTExBid 
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
                  Left            =   -68760
                  Style           =   2  'Dropdown List
                  TabIndex        =   141
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.CommandButton cmdTInReset 
                  BackColor       =   &H000080FF&
                  Caption         =   "Reset"
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
                  Left            =   3600
                  Style           =   1  'Graphical
                  TabIndex        =   179
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.CommandButton cmdTInSave 
                  BackColor       =   &H000080FF&
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
                  Left            =   2400
                  Style           =   1  'Graphical
                  TabIndex        =   178
                  Top             =   2400
                  Width           =   855
               End
               Begin VB.TextBox txtTInTotal 
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1200
                  TabIndex        =   177
                  Top             =   1920
                  Width           =   1935
               End
               Begin VB.TextBox txtTIncheque 
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   5880
                  TabIndex        =   176
                  Top             =   1920
                  Visible         =   0   'False
                  Width           =   1335
               End
               Begin VB.CheckBox CheckTIncheque 
                  Caption         =   "Cheque:"
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
                  Left            =   3720
                  TabIndex        =   173
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.TextBox txtTIncash 
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   2040
                  TabIndex        =   172
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.CheckBox checkTincash 
                  BackColor       =   &H80000004&
                  Caption         =   "Cash"
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
                  Left            =   1200
                  TabIndex        =   171
                  Top             =   1440
                  Width           =   855
               End
               Begin VB.ComboBox combTInUnderBy 
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   315
                  ItemData        =   "frmMain.frx":2AA95
                  Left            =   5400
                  List            =   "frmMain.frx":2AAA5
                  Style           =   2  'Dropdown List
                  TabIndex        =   170
                  Top             =   960
                  Width           =   1815
               End
               Begin MSComCtl2.DTPicker dateTIn 
                  Height          =   375
                  Left            =   5400
                  TabIndex        =   169
                  Top             =   480
                  Width           =   1815
                  _ExtentX        =   3201
                  _ExtentY        =   661
                  _Version        =   393216
                  CalendarBackColor=   12640511
                  Format          =   109182977
                  CurrentDate     =   39568
               End
               Begin VB.TextBox txtTInParticular 
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   1200
                  TabIndex        =   168
                  Top             =   960
                  Width           =   1935
               End
               Begin VB.ComboBox combTExUnderBy 
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
                  ItemData        =   "frmMain.frx":2AAE5
                  Left            =   -69600
                  List            =   "frmMain.frx":2AB0A
                  Style           =   2  'Dropdown List
                  TabIndex        =   137
                  Top             =   960
                  Width           =   1695
               End
               Begin VB.CommandButton cmdTExReset 
                  BackColor       =   &H00FFC0FF&
                  Caption         =   "Reset"
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
                  Left            =   -71160
                  Style           =   1  'Graphical
                  TabIndex        =   145
                  Top             =   2400
                  Width           =   1095
               End
               Begin VB.CommandButton cmdTExSave 
                  BackColor       =   &H00FFC0FF&
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
                  Left            =   -72840
                  Style           =   1  'Graphical
                  TabIndex        =   144
                  Top             =   2400
                  Width           =   975
               End
               Begin VB.TextBox txtTExCheque 
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
                  Left            =   -69000
                  TabIndex        =   142
                  Top             =   1800
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.CheckBox chkTExCheque 
                  Caption         =   "Cheque"
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
                  Left            =   -71280
                  TabIndex        =   140
                  Top             =   1440
                  Width           =   1095
               End
               Begin VB.TextBox txtTExCash 
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
                  Left            =   -72960
                  TabIndex        =   139
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   1095
               End
               Begin VB.CheckBox chkTExCash 
                  Caption         =   "Cash"
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
                  Left            =   -73800
                  TabIndex        =   138
                  Top             =   1440
                  Width           =   855
               End
               Begin VB.TextBox txtTExTotal 
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
                  Left            =   -73800
                  TabIndex        =   143
                  Top             =   1920
                  Width           =   1935
               End
               Begin VB.TextBox txtTExParticular 
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
                  Left            =   -73800
                  TabIndex        =   136
                  Top             =   960
                  Width           =   1935
               End
               Begin MSComCtl2.DTPicker dateTEx 
                  Height          =   375
                  Left            =   -69600
                  TabIndex        =   135
                  Top             =   480
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
                  Format          =   109182977
                  CurrentDate     =   39567
               End
               Begin VB.Label lblTInBid 
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
                  Height          =   255
                  Left            =   5400
                  TabIndex        =   174
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   615
               End
               Begin VB.Label lblBankid 
                  AutoSize        =   -1  'True
                  Caption         =   "Bank-Id"
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
                  Left            =   -69720
                  TabIndex        =   160
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   690
               End
               Begin VB.Label Label75 
                  Caption         =   "Total:"
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
                  TabIndex        =   156
                  Top             =   1920
                  Width           =   615
               End
               Begin VB.Label Label74 
                  Caption         =   "Cr (To):"
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
                  TabIndex        =   155
                  Top             =   1440
                  Width           =   855
               End
               Begin VB.Label Label73 
                  AutoSize        =   -1  'True
                  Caption         =   "Under by:"
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
                  Left            =   3720
                  TabIndex        =   154
                  Top             =   960
                  Width           =   870
               End
               Begin VB.Label Label72 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   3720
                  TabIndex        =   153
                  Top             =   480
                  Width           =   480
               End
               Begin VB.Label Label71 
                  Caption         =   "Particulars:"
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
                  TabIndex        =   152
                  Top             =   960
                  Width           =   855
               End
               Begin VB.Label Label70 
                  Caption         =   "Emp-Id:"
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
                  TabIndex        =   151
                  Top             =   480
                  Width           =   735
               End
               Begin VB.Label Label69 
                  AutoSize        =   -1  'True
                  Caption         =   "Under By:"
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
                  Left            =   -71280
                  TabIndex        =   150
                  Top             =   960
                  Width           =   885
               End
               Begin VB.Label Label68 
                  AutoSize        =   -1  'True
                  Caption         =   "Cr (To):"
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
                  Left            =   -74880
                  TabIndex        =   149
                  Top             =   1440
                  Width           =   660
               End
               Begin VB.Label Label67 
                  AutoSize        =   -1  'True
                  Caption         =   "Total:"
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
                  Left            =   -74880
                  TabIndex        =   148
                  Top             =   1920
                  Width           =   510
               End
               Begin VB.Label Label66 
                  AutoSize        =   -1  'True
                  Caption         =   "Particulars:"
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
                  Left            =   -74880
                  TabIndex        =   147
                  Top             =   960
                  Width           =   990
               End
               Begin VB.Label Label65 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   -71280
                  TabIndex        =   146
                  Top             =   480
                  Width           =   480
               End
               Begin VB.Label Label62 
                  AutoSize        =   -1  'True
                  Caption         =   "Emp-Id:"
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
                  Left            =   -74880
                  TabIndex        =   134
                  Top             =   480
                  Width           =   690
               End
            End
            Begin VB.Frame Frame5 
               Caption         =   "Saving"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   2535
               Left            =   240
               TabIndex        =   77
               Top             =   240
               Width           =   3615
               Begin VB.Timer Timer5 
                  Interval        =   10
                  Left            =   3000
                  Top             =   840
               End
               Begin VB.TextBox txttempsaver 
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   3000
                  TabIndex        =   192
                  Text            =   "Text2"
                  Top             =   1440
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.ComboBox cmbsaving 
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
                  Left            =   1200
                  Style           =   2  'Dropdown List
                  TabIndex        =   191
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.CommandButton cmdSaving 
                  BackColor       =   &H000080FF&
                  Caption         =   "View Report"
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
                  Left            =   1800
                  Style           =   1  'Graphical
                  TabIndex        =   81
                  Top             =   2040
                  Width           =   1455
               End
               Begin VB.TextBox txtSavingAmount 
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
                  Left            =   1200
                  TabIndex        =   79
                  Top             =   1560
                  Width           =   1575
               End
               Begin VB.CommandButton cmdSavingSave 
                  BackColor       =   &H000080FF&
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
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   78
                  Top             =   2040
                  Width           =   1455
               End
               Begin MSComCtl2.DTPicker DTPickerSaving 
                  Height          =   375
                  Left            =   1200
                  TabIndex        =   80
                  Top             =   1080
                  Width           =   1575
                  _ExtentX        =   2778
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
               Begin VB.Label Label17 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   120
                  TabIndex        =   84
                  Top             =   600
                  Width           =   780
               End
               Begin VB.Label Label18 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   120
                  TabIndex        =   83
                  Top             =   1080
                  Width           =   480
               End
               Begin VB.Label Label19 
                  AutoSize        =   -1  'True
                  Caption         =   "Amount"
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
                  Left            =   120
                  TabIndex        =   82
                  Top             =   1560
                  Width           =   675
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Share"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   2535
               Left            =   4080
               TabIndex        =   71
               Top             =   240
               Width           =   3615
               Begin VB.CommandButton cmdShareSave 
                  BackColor       =   &H000080FF&
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
                  Left            =   1320
                  Style           =   1  'Graphical
                  TabIndex        =   228
                  Top             =   2040
                  Width           =   855
               End
               Begin VB.Timer Timer6 
                  Interval        =   10
                  Left            =   3000
                  Top             =   1080
               End
               Begin VB.ComboBox cmbShare 
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
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   193
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.TextBox txtShareAmount 
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
                  MaxLength       =   4
                  TabIndex        =   72
                  Top             =   1560
                  Width           =   1575
               End
               Begin MSComCtl2.DTPicker DTPickerShare 
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   73
                  Top             =   1080
                  Width           =   1575
                  _ExtentX        =   2778
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
               Begin VB.Label Label23 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   120
                  TabIndex        =   76
                  Top             =   600
                  Width           =   780
               End
               Begin VB.Label Label24 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   120
                  TabIndex        =   75
                  Top             =   1080
                  Width           =   480
               End
               Begin VB.Label Label25 
                  AutoSize        =   -1  'True
                  Caption         =   "Amount:"
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
                  Left            =   120
                  TabIndex        =   74
                  Top             =   1560
                  Width           =   720
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "Loan"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   2535
               Left            =   7920
               TabIndex        =   63
               Top             =   240
               Width           =   3615
               Begin VB.Timer Timer7 
                  Interval        =   10
                  Left            =   1200
                  Top             =   2040
               End
               Begin VB.TextBox txttemp 
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   3000
                  TabIndex        =   190
                  Text            =   "Text2"
                  Top             =   1560
                  Visible         =   0   'False
                  Width           =   495
               End
               Begin VB.TextBox Text1 
                  BackColor       =   &H00C0E0FF&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   375
                  Left            =   3000
                  TabIndex        =   189
                  Text            =   "Text1"
                  Top             =   960
                  Visible         =   0   'False
                  Width           =   615
               End
               Begin VB.ComboBox comboloan 
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
                  Left            =   1320
                  Style           =   2  'Dropdown List
                  TabIndex        =   188
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.TextBox txtLoanAmount 
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
                  TabIndex        =   66
                  Top             =   1560
                  Width           =   1575
               End
               Begin VB.CommandButton cmdLoanSave 
                  BackColor       =   &H000080FF&
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
                  Left            =   120
                  Style           =   1  'Graphical
                  TabIndex        =   65
                  Top             =   2040
                  Width           =   855
               End
               Begin VB.CommandButton cmdPayInterest 
                  BackColor       =   &H000080FF&
                  Caption         =   "Pay Interest"
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
                  Left            =   1800
                  MaskColor       =   &H000000FF&
                  Style           =   1  'Graphical
                  TabIndex        =   64
                  Top             =   2040
                  UseMaskColor    =   -1  'True
                  Width           =   1575
               End
               Begin MSComCtl2.DTPicker dateloan 
                  Height          =   375
                  Left            =   1320
                  TabIndex        =   67
                  Top             =   1080
                  Width           =   1575
                  _ExtentX        =   2778
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
               Begin VB.Label Label26 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   120
                  TabIndex        =   70
                  Top             =   600
                  Width           =   780
               End
               Begin VB.Label Label27 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   120
                  TabIndex        =   69
                  Top             =   1080
                  Width           =   480
               End
               Begin VB.Label Label28 
                  AutoSize        =   -1  'True
                  Caption         =   "Amount:"
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
                  Left            =   120
                  TabIndex        =   68
                  Top             =   1560
                  Width           =   720
               End
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "Expenses"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   1815
            Left            =   -74640
            TabIndex        =   54
            Top             =   1080
            Width           =   11055
            Begin VB.TextBox txtTtotal 
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
               Left            =   2880
               TabIndex        =   195
               Top             =   1080
               Width           =   2295
            End
            Begin VB.ComboBox combEx 
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
               ItemData        =   "frmMain.frx":2ABAB
               Left            =   2880
               List            =   "frmMain.frx":2ABAD
               Style           =   2  'Dropdown List
               TabIndex        =   58
               Top             =   480
               Width           =   2295
            End
            Begin VB.TextBox txtExCashAmount 
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
               Left            =   8040
               TabIndex        =   57
               Top             =   360
               Width           =   1935
            End
            Begin VB.TextBox txtExChequeAmount 
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
               Left            =   8040
               TabIndex        =   56
               Top             =   840
               Width           =   1935
            End
            Begin VB.TextBox txtExTotal 
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
               Left            =   8040
               TabIndex        =   55
               Top             =   1320
               Width           =   1935
            End
            Begin VB.Label Label30 
               AutoSize        =   -1  'True
               Caption         =   "Total Expenditure:"
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
               Left            =   1200
               TabIndex        =   194
               Top             =   1080
               Width           =   1620
            End
            Begin VB.Label Label76 
               Caption         =   "Cheque"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   6480
               TabIndex        =   164
               Top             =   840
               Width           =   855
            End
            Begin VB.Label Label44 
               Caption         =   "Cash"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   6480
               TabIndex        =   163
               Top             =   360
               Width           =   615
            End
            Begin VB.Label Label31 
               AutoSize        =   -1  'True
               Caption         =   "Particular:"
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
               Left            =   1200
               TabIndex        =   61
               Top             =   480
               Width           =   885
            End
            Begin VB.Label Label33 
               Caption         =   "To:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   375
               Left            =   5880
               TabIndex        =   60
               Top             =   600
               Width           =   255
            End
            Begin VB.Label Label36 
               Caption         =   "Total:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   6480
               TabIndex        =   59
               Top             =   1320
               Width           =   855
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Income"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   2055
            Left            =   -74640
            TabIndex        =   46
            Top             =   1080
            Width           =   9975
            Begin VB.TextBox txtTIn 
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
               Left            =   2040
               TabIndex        =   197
               Top             =   1200
               Width           =   2415
            End
            Begin VB.TextBox txtIncomeTotal 
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
               Left            =   7080
               TabIndex        =   50
               Top             =   1440
               Width           =   1815
            End
            Begin VB.TextBox txtIncomecheque 
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
               Left            =   7080
               TabIndex        =   49
               Top             =   960
               Width           =   1815
            End
            Begin VB.TextBox txtIncomeCash 
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
               Left            =   7080
               TabIndex        =   48
               Top             =   480
               Width           =   1815
            End
            Begin VB.ComboBox combIncome 
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
               ItemData        =   "frmMain.frx":2ABAF
               Left            =   2040
               List            =   "frmMain.frx":2ABC5
               Style           =   2  'Dropdown List
               TabIndex        =   47
               Top             =   600
               Width           =   2415
            End
            Begin VB.Label Label40 
               AutoSize        =   -1  'True
               Caption         =   "By:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   5040
               TabIndex        =   229
               Top             =   480
               Width           =   285
            End
            Begin VB.Label Label41 
               AutoSize        =   -1  'True
               Caption         =   "Total Income:"
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
               Left            =   720
               TabIndex        =   196
               Top             =   1200
               Width           =   1215
            End
            Begin VB.Label Label35 
               AutoSize        =   -1  'True
               Caption         =   "Cheque"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   240
               Left            =   5760
               TabIndex        =   181
               Top             =   960
               Width           =   810
            End
            Begin VB.Label Label32 
               Caption         =   "Cash"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   5760
               TabIndex        =   180
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label37 
               Caption         =   "Total:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   375
               Left            =   5760
               TabIndex        =   53
               Top             =   1440
               Width           =   615
            End
            Begin VB.Label Label39 
               AutoSize        =   -1  'True
               Caption         =   "By:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Left            =   5040
               TabIndex        =   52
               Top             =   840
               Width           =   285
            End
            Begin VB.Label Label42 
               Caption         =   "Particular:"
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
               Left            =   720
               TabIndex        =   51
               Top             =   600
               Width           =   975
            End
         End
         Begin VB.Frame Frame10 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4815
            Left            =   -74640
            TabIndex        =   5
            Top             =   1200
            Width           =   11175
            Begin VB.Frame Frame11 
               Caption         =   "Add Bank"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H8000000D&
               Height          =   3975
               Left            =   240
               TabIndex        =   29
               Top             =   240
               Width           =   5175
               Begin VB.ComboBox combBankType 
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
                  ItemData        =   "frmMain.frx":2AC13
                  Left            =   2640
                  List            =   "frmMain.frx":2AC1D
                  Style           =   2  'Dropdown List
                  TabIndex        =   41
                  Top             =   2280
                  Width           =   2055
               End
               Begin VB.TextBox txtBankId 
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
                  Left            =   2640
                  TabIndex        =   36
                  Top             =   360
                  Width           =   2055
               End
               Begin VB.TextBox txtBankName 
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
                  Left            =   2640
                  TabIndex        =   37
                  Top             =   840
                  Width           =   2055
               End
               Begin VB.TextBox txtBankAddress 
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
                  Height          =   405
                  Left            =   2640
                  TabIndex        =   39
                  Top             =   1320
                  Width           =   2055
               End
               Begin VB.TextBox txtBankAccNo 
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
                  Left            =   2640
                  TabIndex        =   40
                  Top             =   1800
                  Width           =   2055
               End
               Begin VB.CommandButton cmdBankView 
                  BackColor       =   &H000080FF&
                  Caption         =   "View"
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
                  Left            =   240
                  Style           =   1  'Graphical
                  TabIndex        =   31
                  Top             =   3480
                  Width           =   855
               End
               Begin VB.CommandButton cmdBankDelete 
                  BackColor       =   &H000080FF&
                  Caption         =   "Delete"
                  Enabled         =   0   'False
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
                  Left            =   1560
                  Style           =   1  'Graphical
                  TabIndex        =   33
                  Top             =   3480
                  Width           =   855
               End
               Begin VB.CommandButton cmdBankSave 
                  BackColor       =   &H000080FF&
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
                  Left            =   3840
                  Style           =   1  'Graphical
                  TabIndex        =   0
                  Top             =   3480
                  Visible         =   0   'False
                  Width           =   855
               End
               Begin VB.CommandButton cmdBankEdit 
                  BackColor       =   &H000080FF&
                  Caption         =   "Edit"
                  Enabled         =   0   'False
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
                  Left            =   2760
                  Style           =   1  'Graphical
                  TabIndex        =   34
                  Top             =   3480
                  Width           =   855
               End
               Begin VB.TextBox txtBankTotal 
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
                  Left            =   2640
                  TabIndex        =   42
                  Top             =   2760
                  Width           =   2055
               End
               Begin VB.ListBox ListBank 
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
                  ForeColor       =   &H8000000D&
                  Height          =   2460
                  ItemData        =   "frmMain.frx":2AC36
                  Left            =   360
                  List            =   "frmMain.frx":2AC38
                  TabIndex        =   30
                  Top             =   600
                  Width           =   1335
               End
               Begin VB.Label Label29 
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
                  Left            =   1800
                  TabIndex        =   159
                  Top             =   2280
                  Width           =   735
               End
               Begin VB.Label Label45 
                  Caption         =   "Bank-Id:"
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
                  TabIndex        =   45
                  Top             =   360
                  Width           =   735
               End
               Begin VB.Label Label47 
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
                  Left            =   1800
                  TabIndex        =   44
                  Top             =   840
                  Width           =   855
               End
               Begin VB.Label Label48 
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
                  Left            =   1800
                  TabIndex        =   43
                  Top             =   1320
                  Width           =   855
               End
               Begin VB.Label Label49 
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
                  Left            =   1800
                  TabIndex        =   38
                  Top             =   1800
                  Width           =   855
               End
               Begin VB.Label Label59 
                  Caption         =   "Total:"
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
                  TabIndex        =   35
                  Top             =   2760
                  Width           =   735
               End
               Begin VB.Label Label61 
                  Caption         =   "Bank ID"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   255
                  Left            =   600
                  TabIndex        =   32
                  Top             =   360
                  Width           =   735
               End
            End
            Begin TabDlg.SSTab SSTab2 
               Height          =   3855
               Left            =   6000
               TabIndex        =   6
               Top             =   360
               Width           =   4935
               _ExtentX        =   8705
               _ExtentY        =   6800
               _Version        =   393216
               Tabs            =   2
               TabsPerRow      =   2
               TabHeight       =   520
               BackColor       =   12640511
               ForeColor       =   255
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "MS Sans Serif"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               TabCaption(0)   =   "Deposit"
               TabPicture(0)   =   "frmMain.frx":2AC3A
               Tab(0).ControlEnabled=   -1  'True
               Tab(0).Control(0)=   "Label22"
               Tab(0).Control(0).Enabled=   0   'False
               Tab(0).Control(1)=   "Label57"
               Tab(0).Control(1).Enabled=   0   'False
               Tab(0).Control(2)=   "Label56"
               Tab(0).Control(2).Enabled=   0   'False
               Tab(0).Control(3)=   "Label55"
               Tab(0).Control(3).Enabled=   0   'False
               Tab(0).Control(4)=   "Label54"
               Tab(0).Control(4).Enabled=   0   'False
               Tab(0).Control(5)=   "dateDep"
               Tab(0).Control(5).Enabled=   0   'False
               Tab(0).Control(6)=   "cmdDepReset"
               Tab(0).Control(6).Enabled=   0   'False
               Tab(0).Control(7)=   "cmdDepSave"
               Tab(0).Control(7).Enabled=   0   'False
               Tab(0).Control(8)=   "txtDepAmount"
               Tab(0).Control(8).Enabled=   0   'False
               Tab(0).Control(9)=   "txtDepAccNo"
               Tab(0).Control(9).Enabled=   0   'False
               Tab(0).Control(10)=   "txtDepName"
               Tab(0).Control(10).Enabled=   0   'False
               Tab(0).Control(11)=   "combDepId"
               Tab(0).Control(11).Enabled=   0   'False
               Tab(0).Control(12)=   "Timer9"
               Tab(0).Control(12).Enabled=   0   'False
               Tab(0).ControlCount=   13
               TabCaption(1)   =   "Withdraw"
               TabPicture(1)   =   "frmMain.frx":2AC56
               Tab(1).ControlEnabled=   0   'False
               Tab(1).Control(0)=   "Timer8"
               Tab(1).Control(1)=   "combWithId"
               Tab(1).Control(2)=   "txtWithName"
               Tab(1).Control(3)=   "txtWithAccNo"
               Tab(1).Control(4)=   "txtWithAmount"
               Tab(1).Control(5)=   "cmdWithSave"
               Tab(1).Control(6)=   "cmdWithReset"
               Tab(1).Control(7)=   "dateWith"
               Tab(1).Control(8)=   "Label50"
               Tab(1).Control(9)=   "Label51"
               Tab(1).Control(10)=   "Label52"
               Tab(1).Control(11)=   "Label53"
               Tab(1).Control(12)=   "Label21"
               Tab(1).ControlCount=   13
               Begin VB.Timer Timer9 
                  Interval        =   10
                  Left            =   3960
                  Top             =   1680
               End
               Begin VB.Timer Timer8 
                  Interval        =   10
                  Left            =   -71280
                  Top             =   1440
               End
               Begin VB.ComboBox combWithId 
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
                  Left            =   -73680
                  Style           =   2  'Dropdown List
                  TabIndex        =   166
                  Top             =   600
                  Width           =   2055
               End
               Begin VB.ComboBox combDepId 
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
                  Left            =   1440
                  Style           =   2  'Dropdown List
                  TabIndex        =   165
                  Top             =   600
                  Width           =   2055
               End
               Begin VB.TextBox txtWithName 
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
                  Left            =   -73680
                  TabIndex        =   18
                  Top             =   1080
                  Width           =   2055
               End
               Begin VB.TextBox txtWithAccNo 
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
                  Left            =   -73680
                  TabIndex        =   17
                  Top             =   2040
                  Width           =   2055
               End
               Begin VB.TextBox txtWithAmount 
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
                  Left            =   -73680
                  TabIndex        =   16
                  Top             =   2520
                  Width           =   2055
               End
               Begin VB.CommandButton cmdWithSave 
                  BackColor       =   &H008080FF&
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
                  Left            =   -74640
                  Style           =   1  'Graphical
                  TabIndex        =   15
                  Top             =   3360
                  Width           =   855
               End
               Begin VB.CommandButton cmdWithReset 
                  BackColor       =   &H008080FF&
                  Caption         =   "Reset"
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
                  Left            =   -71880
                  Style           =   1  'Graphical
                  TabIndex        =   14
                  Top             =   3360
                  Width           =   855
               End
               Begin VB.TextBox txtDepName 
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
                  Left            =   1440
                  TabIndex        =   13
                  Top             =   1080
                  Width           =   2055
               End
               Begin VB.TextBox txtDepAccNo 
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
                  Left            =   1440
                  TabIndex        =   12
                  Top             =   2040
                  Width           =   2055
               End
               Begin VB.TextBox txtDepAmount 
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
                  Left            =   1440
                  TabIndex        =   11
                  Top             =   2520
                  Width           =   2055
               End
               Begin VB.CommandButton cmdDepSave 
                  BackColor       =   &H000080FF&
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
                  Left            =   1560
                  Style           =   1  'Graphical
                  TabIndex        =   10
                  Top             =   3360
                  Width           =   975
               End
               Begin VB.CommandButton cmdDepReset 
                  BackColor       =   &H000080FF&
                  Caption         =   "Reset"
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
                  Left            =   2760
                  Style           =   1  'Graphical
                  TabIndex        =   9
                  Top             =   3360
                  Width           =   975
               End
               Begin MSComCtl2.DTPicker dateDep 
                  Height          =   375
                  Left            =   1440
                  TabIndex        =   7
                  Top             =   1560
                  Width           =   2055
                  _ExtentX        =   3625
                  _ExtentY        =   661
                  _Version        =   393216
                  CalendarBackColor=   16777215
                  CalendarTitleBackColor=   16777215
                  CalendarTrailingForeColor=   16777215
                  Format          =   109182977
                  CurrentDate     =   39567
               End
               Begin MSComCtl2.DTPicker dateWith 
                  Height          =   375
                  Left            =   -73680
                  TabIndex        =   8
                  Top             =   1560
                  Width           =   2055
                  _ExtentX        =   3625
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
                  CurrentDate     =   39567
               End
               Begin VB.Label Label50 
                  AutoSize        =   -1  'True
                  Caption         =   "Bank-Id:"
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
                  Left            =   -74640
                  TabIndex        =   28
                  Top             =   600
                  Width           =   735
               End
               Begin VB.Label Label51 
                  AutoSize        =   -1  'True
                  Caption         =   "Name"
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
                  Left            =   -74640
                  TabIndex        =   27
                  Top             =   1080
                  Width           =   555
               End
               Begin VB.Label Label52 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   -74640
                  TabIndex        =   26
                  Top             =   2040
                  Width           =   720
               End
               Begin VB.Label Label53 
                  AutoSize        =   -1  'True
                  Caption         =   "Amount:"
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
                  Left            =   -74640
                  TabIndex        =   25
                  Top             =   2520
                  Width           =   720
               End
               Begin VB.Label Label54 
                  AutoSize        =   -1  'True
                  Caption         =   "Bank-Id:"
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
                  Left            =   240
                  TabIndex        =   24
                  Top             =   600
                  Width           =   735
               End
               Begin VB.Label Label55 
                  AutoSize        =   -1  'True
                  Caption         =   "Name"
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
                  Left            =   240
                  TabIndex        =   23
                  Top             =   1080
                  Width           =   555
               End
               Begin VB.Label Label56 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   240
                  TabIndex        =   22
                  Top             =   2040
                  Width           =   720
               End
               Begin VB.Label Label57 
                  AutoSize        =   -1  'True
                  Caption         =   "Amount:"
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
                  Left            =   240
                  TabIndex        =   21
                  Top             =   2520
                  Width           =   720
               End
               Begin VB.Label Label21 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   -74640
                  TabIndex        =   20
                  Top             =   1560
                  Width           =   480
               End
               Begin VB.Label Label22 
                  AutoSize        =   -1  'True
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
                  Height          =   240
                  Left            =   240
                  TabIndex        =   19
                  Top             =   1560
                  Width           =   480
               End
            End
         End
         Begin VB.Label Label82 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            Caption         =   $"frmMain.frx":2AC72
            ForeColor       =   &H80000008&
            Height          =   1815
            Left            =   -73800
            TabIndex        =   227
            Top             =   2160
            Width           =   8655
         End
         Begin VB.Label Label81 
            Caption         =   "Credit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -65760
            TabIndex        =   224
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Label Label80 
            Caption         =   "Debit"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   -71520
            TabIndex        =   223
            Top             =   1140
            Width           =   1095
         End
         Begin VB.Line Line5 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   -74760
            X2              =   -74760
            Y1              =   1680
            Y2              =   6120
         End
         Begin VB.Line Line4 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   -63480
            X2              =   -63480
            Y1              =   1680
            Y2              =   6120
         End
         Begin VB.Line Line3 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   -74760
            X2              =   -63480
            Y1              =   1680
            Y2              =   1680
         End
         Begin VB.Label lblpl 
            Height          =   495
            Left            =   -74160
            TabIndex        =   221
            Top             =   6360
            Width           =   1935
         End
         Begin VB.Label Label79 
            Caption         =   "Interest"
            Height          =   495
            Left            =   -69120
            TabIndex        =   215
            Top             =   5160
            Width           =   1335
         End
         Begin VB.Label Label78 
            Caption         =   "Other Income"
            Height          =   495
            Left            =   -69120
            TabIndex        =   214
            Top             =   4440
            Width           =   1935
         End
         Begin VB.Label Label77 
            Caption         =   "Sales Income"
            Height          =   495
            Left            =   -69120
            TabIndex        =   213
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label64 
            Caption         =   "Indirect Income"
            Height          =   495
            Left            =   -69120
            TabIndex        =   212
            Top             =   2640
            Width           =   2295
         End
         Begin VB.Label Label63 
            Caption         =   "Direct Income"
            Height          =   495
            Left            =   -69120
            TabIndex        =   211
            Top             =   1920
            Width           =   2295
         End
         Begin VB.Label I 
            Caption         =   "Interest:"
            Height          =   495
            Left            =   -74640
            TabIndex        =   209
            Top             =   4320
            Width           =   1695
         End
         Begin VB.Label Label58 
            Caption         =   "Purchase  Amount"
            Height          =   375
            Left            =   -74640
            TabIndex        =   207
            Top             =   3480
            Width           =   2415
         End
         Begin VB.Label Label46 
            Caption         =   "Direct Expense"
            Height          =   495
            Left            =   -74640
            TabIndex        =   205
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label Label43 
            Caption         =   "Indirect Expense"
            Height          =   375
            Left            =   -74640
            TabIndex        =   203
            Top             =   1920
            Width           =   2295
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            BorderWidth     =   3
            X1              =   -74760
            X2              =   -63480
            Y1              =   6120
            Y2              =   6120
         End
         Begin VB.Label lblIncomeflex 
            BackColor       =   &H8000000D&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   375
            Left            =   -72000
            TabIndex        =   183
            Top             =   3360
            Visible         =   0   'False
            Width           =   3495
            WordWrap        =   -1  'True
         End
         Begin VB.Label flexcaption 
            Alignment       =   2  'Center
            BackColor       =   &H8000000D&
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
            Height          =   375
            Left            =   -71160
            TabIndex        =   162
            Top             =   3120
            Visible         =   0   'False
            Width           =   3975
         End
      End
   End
   Begin VB.Menu f 
      Caption         =   "File"
      Begin VB.Menu cu 
         Caption         =   "Change Password"
      End
   End
   Begin VB.Menu tl 
      Caption         =   "Tools"
      Begin VB.Menu cal 
         Caption         =   "Calculator"
      End
      Begin VB.Menu note 
         Caption         =   "Notepad"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public str As String
Dim cflag As Boolean
Dim mflag As Boolean
Dim flag As Boolean
Dim lstmain1 As ListItem
Dim st As String
Public vbmsgbox As VbMsgBoxResult

Private Sub Check3_Click()

End Sub

Private Sub cal_Click()
Shell ("calc.exe"), vbNormalFocus
End Sub

Private Sub CheckTIncheque_Click()
    If CheckTIncheque.value = 1 Then
        lblTInBid.Visible = True
        combTInBid.Visible = True
        txtTIncheque.Visible = True
    Else
        lblTInBid.Visible = False
        combTInBid.Visible = False
        txtTIncheque.Visible = False
        txtTIncheque.Text = ""
    End If
End Sub

Private Sub chkTExCash_Click()
    If chkTExCash.value = 1 Then
    txtTExCash.Visible = True
    Else
    txtTExCash.Text = ""
    txtTExCash.Visible = False
    End If
End Sub

Private Sub chkTExCheque_Click()
    If chkTExCheque.value = 1 Then
        lblBankid.Visible = True
        txtTExCheque.Visible = True
        combTExBid.Visible = True
    Else
        lblBankid.Visible = False
        txtTExCheque.Visible = False
        combTExBid.Visible = False
        txtTExCheque.Text = ""
        
    End If
End Sub

Private Sub cmdAbout_Click()
    SSTab1.Tab = 9
End Sub

Private Sub cmdAdd_Click()
    Form2.Show
    Frame1.Visible = False
End Sub

Private Sub cmdBankAdd_Click()
    Frame1.Visible = False
    Form5.Show
End Sub

Private Sub cmdBankDelete_Click()
Dim a As Integer
st = ListBank.Text
str = "select * from tblBank where b_id=" & st
CustomRecordSetOpen str
With res
    While Not .EOF
    a = MsgBox("Do you want delete?", vbOKCancel)
    If (a = vbOK) Then
    .Delete
    .MoveNext
    .UpdateBatch
    MsgBox "Deleted Successfully..."
    Else
    Exit Sub
    End If
    Wend
End With
CloseRecordSet
End Sub

Private Sub cmdBankEdit_Click()
    cmdBankAdd.Enabled = False
    cmdBankDelete.Enabled = False
    cmdBankEdit.Enabled = False
    cmdBankSave.Visible = True
End Sub

Private Sub cmdBankSave_Click()
    st = ListBank.Text
    str = "select * from tblBank where b_id=" & st
    CustomRecordSetOpen str
    With res
        !Name = txtBankName.Text
        !address = txtBankAddress.Text
        !acc_no = txtBankAccNo.Text
        !Type = combBankType.Text
        !total = txtBankTotal.Text
        .UpdateBatch
        MsgBox "Successfullt Edited...", vbInformation
    End With
    CloseRecordSet
    cmdBankAdd.Enabled = True
    cmdBankDelete.Enabled = True
    cmdBankEdit.Enabled = True
    cmdBankSave.Visible = False
End Sub

Private Sub cmdBankView_Click()
     If flag = False Then
        flag = True
    st = "tblBank"
    CustomRecordSetOpen st
    With res
        While Not .EOF
        ListBank.AddItem !b_id
        .MoveNext
        Wend
    End With
    CloseRecordSet
    cmdBankDelete.Enabled = True
    cmdBankEdit.Enabled = True
    
    Else
        flag = False
        txtBankAccNo.Text = ""
        txtBankAddress.Text = ""
        txtBankName.Text = ""
        'combBankType.Text = ""
        txtBankTotal.Text = ""
        ListBank.Clear
        cmdBankDelete.Enabled = False
        cmdBankEdit.Enabled = False
    End If
End Sub

Private Sub cmdCAdd_Click()
    Form3.Show
    Frame1.Visible = False
End Sub

Private Sub cmdCDelete_Click()
    st = ListM.Text
    str = "Select name,address,gender,phone,cell,join_date,type,acc_no from member where mmbr_id=" & st
    CustomRecordSetOpen str
    With res
        While Not .EOF
            .Delete
            .MoveNext
            .UpdateBatch
        Wend
    End With
    MsgBox "Member Deleted Successfully..."
    CloseRecordSet
        txtCName.Text = ""
        txtCAddress.Text = ""
        txtCAccno.Text = ""
        optCFemale.value = False
        optCMale.value = False
        txtCPhone.Text = ""
        txtCCell.Text = ""
        'combCType.Text = ""
End Sub

Private Sub cmdCExit_Click()

End Sub

Private Sub cmdCEdit_Click()
cmdCSave.Visible = True
cmdCAdd.Enabled = False
cmdCDelete.Enabled = False
cmdCEdit.Enabled = False

combCType.Enabled = False
txtCAccno.Enabled = False
txtCName.SetFocus
End Sub

Private Sub cmdCSave_Click()
    st = ListM.Text
    'str = "Select name,address,gender,phone,cell,join_date,type,acc_no from member where mmbr_id=" & st
    str = "Select * from member where mmbr_id=" & st
    CustomRecordSetOpen str
    
    With res
        While Not .EOF
            !Name = txtCName.Text
            !address = txtCAddress.Text
            If optCMale.value = True Then !gender = "Male"
            If optCFemale.value = True Then !gender = "Female"
            !phone = txtCPhone.Text
            !cell = txtCCell.Text
            !join_date = dateCDoj.value
            !Type = combCType.Text
            !acc_no = txtCAccno.Text
            .MoveNext
            .UpdateBatch
        Wend
    End With
    CloseRecordSet
    MsgBox "Info Edit Successfully..."
    cmdCSave.Visible = False
    cmdCDelete.Enabled = True
    cmdCEdit.Enabled = True
    cmdCAdd.Enabled = True
    combCType.Enabled = True
    txtCAccno.Enabled = True
End Sub

Private Sub cmdCView_Click()
    cmdCDelete.Enabled = True
    cmdCEdit.Enabled = True
    If mflag = False Then
    mflag = True
    
    
    st = "Select * from member"
    CustomRecordSetOpen st
    With res
        While Not .EOF
        ListM.AddItem !mmbr_id
        .MoveNext
        Wend
    End With
    CloseRecordSet
Else
    mflag = False
    txtCName.Text = ""
    txtCAddress.Text = ""
    txtCAccno.Text = ""
    txtCPhone.Text = ""
    txtCCell.Text = ""
    optCFemale.value = False
    optCMale.value = False
    'combCType.Text = ""
    ListM.Clear
End If
End Sub

Private Sub cmdDelete_Click()
    st = List.Text
    CloseRecordSet
    
    str = "Select emp_id,name,gender,dob,address,phone,cell,join_date,department from employee where emp_id=" & st
    CustomRecordSetOpen str
    With res
    While Not .EOF
        .Delete
        .MoveNext
        .UpdateBatch
        
    Wend
    End With
vbmsgbox = MsgBox("Successfully Deleted", vbInformation)
txtName.Text = ""
txtAddress.Text = ""
Option1.value = False
Option2.value = False
txtPhone.Text = ""
txtCell.Text = ""
'Combo1.Text = ""
CloseRecordSet
End Sub

Private Sub cmdDepReset_Click()
cmdDepSave.Enabled = True
txtDepName.Text = ""
txtDepAccNo.Text = ""
txtDepAmount.Text = ""
End Sub

Private Sub cmdDepSave_Click()
    Dim a As Integer
    a = MsgBox("Are You Sure...", vbOKCancel)
    If (a = vbOK) Then
    st = "Select * from tblbankdeposit"
    CustomRecordSetOpen st
    With res
        .AddNew
        !b_id = combDepId.Text
        !Name = txtDepName.Text
        !Date = dateDep.value
        !acc_no = txtDepAccNo.Text
        !amount = txtDepAmount.Text
        .UpdateBatch
    End With
    CloseRecordSet
    str = combDepId.Text
    st = "Select total from tblBank where b_id=" & str
    CustomRecordSetOpen st
    With res
        !total = !total + Val(txtDepAmount.Text)
        .UpdateBatch
    End With
    MsgBox "Successfully Deposited..."
    CloseRecordSet
    cmdDepSave.Enabled = False
    Else
    txtDepAmount.SetFocus
    Exit Sub
    End If
    
    st = "select amount from tblcash "
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !amount = txtDepAmount.Text
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from tblcash "
    CustomRecordSetOpen st
        With res
            !amount = res!amount - Val(txtDepAmount.Text)
            .UpdateBatch
        End With
    CloseRecordSet

    
End Sub

Private Sub cmdEdit_Click()
cmdSave.Visible = True
cmdDelete.Enabled = False
cmdAdd.Enabled = False
cmdEdit.Enabled = False
txtName.SetFocus
End Sub

Private Sub cmdexit_Click()
    Unload Me
End Sub

Private Sub cmdLoan_Click()
    txtLoanAmount.Enabled = True
    'txtLoanId.Enabled = True
    dateloan.Enabled = True
    'cmdLoanReset.Enabled = True
    cmdLoanSave.Enabled = True
        
End Sub

Private Sub cmdLoanReset_Click()
'txtLoanId.Text = ""
txtLoanAmount.Text = ""
End Sub

Private Sub cmdLoanSave_Click()
Dim a As Integer
Dim b As String
Dim temp As String
a = MsgBox("Are you sure...", vbOKCancel)
If (a = vbOK) Then
    st = "loan"
    CustomRecordSetOpen st
    With res
        .AddNew
        !mmbr_id = comboloan.Text
        !Date = dateloan
        !amount = txtLoanAmount.Text
        .UpdateBatch
    End With
    CloseRecordSet
    txtLoanAmount.Text = ""
    txtLoanAmount.SetFocus
End If
    b = comboloan.Text
    st = "select mmbr_id,amount,date from tblLoanInterestReceipt where mmbr_id=" & b
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !mmbr_id = comboloan.Text
       ' !amount = txtLoanAmount.Text
        !Date = dateloan
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from loan where mmbr_id like '" & b & "'"
    CustomRecordSetOpen st
    txttemp.Text = res!amount
    CloseRecordSet
    st = "select mmbr_id,amount,date from tblLoanInterestReceipt where mmbr_id=" & b
    CustomRecordSetOpen st
        With res
           !mmbr_id = comboloan.Text
            !Date = dateloan
            !amount = Val(txtLoanAmount.Text) + Val(txttemp.Text)
            .UpdateBatch
            MsgBox ("Record saved")
        End With
    CloseRecordSet
    
    st = "select amount from tblcash "
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !amount = txtLoanAmount.Text
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from tblcash "
    CustomRecordSetOpen st
        With res
            !amount = res!amount - Val(txtLoanAmount.Text)
            .UpdateBatch
        End With
    CloseRecordSet

End Sub

Private Sub cmdPayInterest_Click()
Form4.Show
Form1.Frame1.Visible = False
End Sub

Private Sub cmdSave_Click()
st = List.Text
str = "Select * from employee where emp_id=" & st
CustomRecordSetOpen str
With res
    While Not .EOF

    !Name = txtName.Text
    If Option1.value = True Then !gender = "Male"
    If Option2.value = True Then !gender = "Female"
    !dob = DTPickeremp.value
    !address = txtAddress.Text
    !phone = txtPhone.Text
    !cell = txtCell.Text
    !join_date = DTPicker2.value
    !department = Combo1.Text
    .MoveNext
    .UpdateBatch
    Wend
End With
vbmsgbox = MsgBox("Record Edit Successfully...", vbInformation)
cmdSave.Visible = False
cmdAdd.Enabled = True
cmdDelete.Enabled = True
cmdEdit.Enabled = True
CloseRecordSet
End Sub


Private Sub cmdSaving_Click()
Form6.Show
Form1.Frame1.Visible = False
End Sub

Private Sub cmdSavingSave_Click()
Dim a As String
Dim b As String
dbconnection
str = "tblsaver"
CustomRecordSetOpen str
a = MsgBox("Are You Sure!!!", vbOKCancel)
If (a = vbOK) Then
With res
    .AddNew
    !mmbr_id = cmbsaving.Text
    !Date = DTPickerSaving.value
    !amount = txtSavingAmount.Text
    .UpdateBatch
    
    End With
    txtSavingAmount.Text = ""
    cmbsaving.SetFocus
    st = "select amount from tblcash "
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !amount = txtSavingAmount.Text
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from tblcash "
    CustomRecordSetOpen st
        With res
            !amount = res!amount + Val(txtSavingAmount.Text)
            .UpdateBatch
        End With
    CloseRecordSet

    MsgBox "Entry Successfully Recorded..."
Else
    Exit Sub
End If
CloseRecordSet
b = cmbsaving.Text
    st = "select mmbr_id,amount,date from tblSaverInterest where mmbr_id=" & b
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !mmbr_id = cmbsaving.Text
        '!amount = txtSavingAmount.Text
        !Date = DTPickerSaving.value
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from tblsaverInterest where mmbr_id like '" & b & "'"
    CustomRecordSetOpen st
    txttempsaver.Text = res!amount
    CloseRecordSet
    st = "select mmbr_id,amount,date from tblSaverInterest where mmbr_id=" & b
    CustomRecordSetOpen st
        With res
           !mmbr_id = cmbsaving.Text
            !Date = DTPickerSaving.value
            !amount = Val(txtSavingAmount.Text) + Val(txttempsaver.Text)
            .UpdateBatch
        End With
    CloseRecordSet
End Sub

Private Sub cmdShare_Click()
    txtShareAmount.Enabled = True
    'txtShareId.Enabled = True
    DTPickerShare.Enabled = True
    'cmdShareReset.Enabled = True
    cmdShareSave.Enabled = True
End Sub

Private Sub cmdSearchDaybook_Click()
    Dim sum As Integer
    Dim a As Integer
     Dim data As Date
     data = DTPicker1.value
     MSFlexGrid2.Clear
    MSFlexGrid2.Cols = 5
    MSFlexGrid2.TextMatrix(0, 1) = "Date"
    MSFlexGrid2.TextMatrix(0, 2) = "Particular"
    MSFlexGrid2.TextMatrix(0, 3) = "Debit"
    MSFlexGrid2.TextMatrix(0, 4) = "Credit"
    
   
   st = "select date,particular,total from tblExpenseTrans where date like '" & data & "'"
    CustomRecordSetOpen st
    a = 1
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = !particular
        MSFlexGrid2.TextMatrix(a, 3) = !total
        MSFlexGrid2.TextMatrix(a, 4) = ""
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet
st = "select date,name,amount from tblbankdeposit where date like '" & data & "'"
CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = !Name
        MSFlexGrid2.TextMatrix(a, 3) = !amount
        MSFlexGrid2.TextMatrix(a, 4) = ""
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet
 st = "select date,name,amount from tblbankwithdraw where date like '" & data & "'"
CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = !Name
        MSFlexGrid2.TextMatrix(a, 3) = ""
        MSFlexGrid2.TextMatrix(a, 4) = !amount
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet

st = "select date,mmbr_id,amount from loan where date like '" & data & "'"
CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = "Loan Paid To MId  " & "(" & !mmbr_id & ")"
        MSFlexGrid2.TextMatrix(a, 3) = !amount
        MSFlexGrid2.TextMatrix(a, 4) = ""
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet

st = "select Date,Particulars,Total from tblIncomeTrans where Date like '" & data & "'"
    CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = !Particulars
        MSFlexGrid2.TextMatrix(a, 3) = ""
        MSFlexGrid2.TextMatrix(a, 4) = !total
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet

st = "select date,mmbr_id,amount from tblsaver where date like '" & data & "'"
CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = "Save Amount Receipt MId by " & "(" & !mmbr_id & ")"
        MSFlexGrid2.TextMatrix(a, 3) = ""
        MSFlexGrid2.TextMatrix(a, 4) = !amount
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet

st = "select date,mmbr_id,amount from tblShare where date like '" & data & "'"
CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = "Share Receipt MId by " & "(" & !mmbr_id & ")"
        MSFlexGrid2.TextMatrix(a, 3) = ""
        MSFlexGrid2.TextMatrix(a, 4) = !amount
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet

End Sub

Private Sub cmdShareSave_Click()
Dim a As String
Dim b As String
str = "tblShare"
CustomRecordSetOpen str
a = MsgBox("Are You Sure!!!", vbOKCancel)
If (a = vbOK) Then
With res
    .AddNew
    !mmbr_id = cmbShare.Text
    !Date = DTPickerShare.value
    !amount = txtShareAmount.Text
    .UpdateBatch
    End With
    txtShareAmount.Text = ""
    txtShareAmount.SetFocus
    CloseRecordSet
str = "tblcapital"
CustomRecordSetOpen str
With res
    .AddNew
    !particular = "mid" & "(" & cmbShare.Text & ")"
    '!amount = txtShareAmount.Text
    .UpdateBatch
End With

st = "select amount from tblcash "
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !amount = txtShareAmount.Text
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from tblcash "
    CustomRecordSetOpen st
        With res
            !amount = res!amount + Val(txtShareAmount.Text)
            .UpdateBatch
        End With
    CloseRecordSet

    MsgBox "Entry Successfully Recorded..."
Else
    Exit Sub
End If

End Sub

Private Sub cmdTCapitalSave_Click()
Dim a As String
st = "tblcapital"
CustomRecordSetOpen st
a = MsgBox("Are you sure!!!", vbOKCancel)
If (a = vbOK) Then
With res
    .AddNew
    !particular = "Cash"
    !amount = txtTCapitalAmount.Text
    .UpdateBatch
End With
CloseRecordSet

MsgBox "Entry Successfully Recorded..."
txtTCapitalAmount.Text = ""
txtTCapitalAmount.SetFocus
Else
    txtTCapitalAmount.Text = ""
    txtTCapitalAmount.SetFocus
    Exit Sub
End If


    st = "select amount from tblcash "
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !amount = txtTCapitalAmount.Text
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from tblcash "
    CustomRecordSetOpen st
        With res
            !amount = res!amount + Val(txtTCapitalAmount.Text)
            .UpdateBatch
        End With
    CloseRecordSet
End Sub

Private Sub cmdTExReset_Click()
    txtTExId.Text = ""
    txtTExParticular.Text = ""
    chkTExCash.value = 0
    chkTExCheque.value = 0
    txtTExTotal.Text = ""
    'combTExUnderBy.Text = ""
    txtTExId.SetFocus
    
End Sub

Private Sub cmdTExSave_Click()
    Dim a As Integer
    str = "tblExpenseTrans"
    CustomRecordSetOpen str
    a = MsgBox("Are you Sure!!!", vbOKCancel)
    If (a = vbOK) Then
    With res
        .AddNew
        !emp_id = txtTExId.Text
        !Date = dateTEx.value
        !particular = txtTExParticular.Text
        !UnderBy = combTExUnderBy.Text
        !Cash = Val(txtTExCash.Text)
        !Cheque = Val(txtTExCheque.Text)
        !b_id = combTExBid.Text
        !total = txtTExTotal.Text
        .UpdateBatch
    End With
    
    st = "select amount from tblcash "
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !amount = txtTExCash.Text
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from tblcash "
    CustomRecordSetOpen st
        With res
            !amount = res!amount - Val(txtTExCash.Text)
            .UpdateBatch
        End With
    CloseRecordSet
    
    
    MsgBox "Entry Successfully Recorded..."
    
Else
    Exit Sub
End If
CloseRecordSet
If chkTExCheque.value = 1 Then
str = combTExBid.Text
st = "Select total from tblBank where b_id=" & str
CustomRecordSetOpen st
With res
    While Not .EOF
    !total = !total - Val(txtTExCheque.Text)
    .UpdateBatch
    .MoveNext
    Wend
End With
CloseRecordSet
Else
Exit Sub
End If
End Sub


Private Sub cmdTInReset_Click()
'txtTExId.Text = ""
    txtTExParticular.Text = ""
    chkTExCash.value = 0
    chkTExCheque.value = 0
    txtTExTotal.Text = ""
    'combTExUnderBy.Text = ""
    txtTExId.SetFocus
txtTIn.Text = ""
txtTInParticular.Text = ""
checkTincash.value = 0
CheckTIncheque.value = 0
txtTInTotal.Text = ""

End Sub

Private Sub cmdTInSave_Click()
Dim a As Integer
    str = "tblIncomeTrans"
    CustomRecordSetOpen str
    a = MsgBox("Are you Sure!!!", vbOKCancel)
    If (a = vbOK) Then
    With res
        .AddNew
        !emp_id = combTInEid.Text
        !Date = dateTIn.value
        !Particulars = txtTInParticular.Text
        !UnderBy = combTInUnderBy.Text
        !Cash = Val(txtTIncash.Text)
        If CheckTIncheque.value = 1 Then
        !Cheque = Val(txtTIncheque.Text)
        !b_id = combTInBid.Text
        End If
        '!total = txtTInTotal.Text
        .UpdateBatch
    End With
    
    st = "select amount from tblcash "
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !amount = txtTIncash.Text
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from tblcash "
    CustomRecordSetOpen st
        With res
            !amount = res!amount + Val(txtTIncash.Text)
            .UpdateBatch
        End With
    CloseRecordSet

    MsgBox "Entry Successfully Recorded..."
    
Else
    Exit Sub
End If
CloseRecordSet
If CheckTIncheque.value = 1 Then
str = combTInBid.Text
st = "Select total from tblBank where b_id=" & str
CustomRecordSetOpen st
With res
    While Not .EOF
    !total = !total + Val(txtTIncheque.Text)
    .UpdateBatch
    .MoveNext
    Wend
End With
CloseRecordSet
Else
Exit Sub
End If
End Sub

Private Sub cmdView_Click()
cmdDelete.Enabled = True
cmdEdit.Enabled = True
Dim row As Integer
If cflag = False Then
cflag = True
CloseRecordSet
CustomRecordSetOpen "select * from employee"
With res
    .MoveFirst
    While Not .EOF
        List.AddItem !emp_id
    .MoveNext
    Wend
End With
CloseRecordSet
Else
cflag = False
txtName.Text = ""
txtAddress.Text = ""
Option1.value = False
Option2.value = False
txtPhone.Text = ""
txtCell.Text = ""
'Combo1.Text = ""
List.Clear
End If

End Sub

Private Sub cmdWithReset_Click()
txtWithAmount.Text = ""
txtWithName.Text = ""
txtWithAccNo.Text = ""
txtWithAmount.SetFocus
cmdWithSave.Enabled = True
End Sub

Private Sub cmdWithSave_Click()
    Dim a As Integer
    a = MsgBox("Are You Sure...", vbOKCancel)
    If (a = vbOK) Then
    st = "Select * from tblbankwithdraw"
    CustomRecordSetOpen st
    With res
        .AddNew
        !b_id = combWithId.Text
        !Name = txtWithName.Text
        !Date = dateWith.value
        !acc_no = txtWithAccNo.Text
        !amount = txtWithAmount.Text
        .UpdateBatch
    End With
    CloseRecordSet
    str = combWithId.Text
    st = "Select total from tblBank where b_id=" & str
    CustomRecordSetOpen st
    With res
        !total = !total - Val(txtWithAmount.Text)
        .UpdateBatch
    End With
    MsgBox "Successfully Withdrawn..."
    CloseRecordSet
    cmdWithSave.Enabled = False
    Else
    txtWithAmount.SetFocus
    Exit Sub
    End If


st = "select amount from tblcash "
    CustomRecordSetOpen st
    If (res.EOF = True And res.BOF = True) Then
    With res
        .AddNew
        !amount = txtWithAmount.Text
        .UpdateBatch
    End With
    CloseRecordSet
    Exit Sub
    End If
    st = "select amount from tblcash "
    CustomRecordSetOpen st
        With res
            !amount = res!amount + Val(txtWithAmount.Text)
            .UpdateBatch
        End With
    CloseRecordSet

End Sub

Private Sub combDepId_Click()
    str = combDepId.Text
    st = "Select b_id,name,date,acc_no from tblBank where b_id=" & str
    CustomRecordSetOpen st
    With res
        txtDepName = !Name
        txtDepAccNo = !acc_no
        dateDep.value = !Date
    End With
    CloseRecordSet
    txtDepAmount.SetFocus
End Sub

Private Sub combEx_Click()
    flexcaption.Visible = True
    Dim intRow As Integer
    Dim totalparticular As Integer
    Dim a As Integer
    str = combEx.Text
    flexcaption.Caption = str
    st = "select sum(cash)as tot from tblExpenseTrans where underBy like '" & str & "'"
    CustomRecordSetOpen st
    txtExCashAmount.Text = res!tot
    CloseRecordSet
    st = "select sum(cheque)as tot from tblExpenseTrans where underBy like '" & str & "'"
    CustomRecordSetOpen st
    txtExChequeAmount.Text = res!tot
    CloseRecordSet
    txtExTotal.Text = Val(txtExCashAmount.Text) + Val(txtExChequeAmount.Text)
    MSFlexGrid1.Cols = 5
    MSFlexGrid1.TextMatrix(0, 1) = "Date"
    MSFlexGrid1.TextMatrix(0, 2) = "Particular"
    MSFlexGrid1.TextMatrix(0, 3) = "Cash"
    MSFlexGrid1.TextMatrix(0, 4) = "Cheque"
    st = "select date, particular, cash, cheque from tblExpenseTrans where underBy = '" & str & "'"
   ' st = "tblExpenseTrans"
    CustomRecordSetOpen st
    a = 1
    With res
     '   .MoveFirst
        While (Not (.EOF))
        MSFlexGrid1.Rows = a + 1
        MSFlexGrid1.TextMatrix(a, 1) = !Date
        MSFlexGrid1.TextMatrix(a, 2) = !particular
        MSFlexGrid1.TextMatrix(a, 3) = !Cash
        MSFlexGrid1.TextMatrix(a, 4) = !Cheque
        .MoveNext
        a = a + 1
        Wend
           'totalparticular = .RecordCount
    End With
CloseRecordSet
End Sub


Private Sub combIncome_Click()
    lblIncomeflex.Visible = True
    Dim intRow As Integer
    Dim totalparticular As Integer
    Dim a As Integer
    str = combIncome.Text
    lblIncomeflex.Caption = str
    st = "select sum(cash)as tot from tblIncomeTrans where underBy like '" & str & "'"
    CustomRecordSetOpen st
    txtIncomeCash.Text = res!tot
    CloseRecordSet
    st = "select sum(cheque)as tot from tblIncomeTrans where underBy like '" & str & "'"
    CustomRecordSetOpen st
    txtIncomecheque.Text = res!tot
    CloseRecordSet
    txtIncomeTotal.Text = Val(txtIncomeCash.Text) + Val(txtIncomecheque.Text)
    flexIncome.Cols = 5
    flexIncome.TextMatrix(0, 1) = "Date"
    flexIncome.TextMatrix(0, 2) = "Particular"
    flexIncome.TextMatrix(0, 3) = "Cash"
    flexIncome.TextMatrix(0, 4) = "Cheque"
    st = "select date, particulars, cash, cheque from tblIncomeTrans where underBy = '" & str & "'"
    CustomRecordSetOpen st
    a = 1
    With res
     
        While (Not (.EOF))
        flexIncome.Rows = a + 1
        flexIncome.TextMatrix(a, 1) = !Date
        flexIncome.TextMatrix(a, 2) = !Particulars
        flexIncome.TextMatrix(a, 3) = !Cash
        flexIncome.TextMatrix(a, 4) = !Cheque
        .MoveNext
        a = a + 1
        Wend
           'totalparticular = .RecordCount
    End With
CloseRecordSet

End Sub

Private Sub combWithId_Click()
    str = combWithId.Text
    st = "Select b_id,name,date,acc_no from tblBank where b_id=" & str
    CustomRecordSetOpen st
    With res
        txtWithName = !Name
        txtWithAccNo = !acc_no
        dateWith.value = !Date
    End With
    CloseRecordSet
    txtWithAmount.SetFocus

End Sub

Private Sub Command3_Click()
Form4.Show
Form1.Frame1.Visible = False
End Sub

Private Sub Command1_Click()
Frame1.Visible = False
Form7.Show
End Sub

Private Sub cu_Click()
Form1.Frame1.Visible = False
Form9.Show

End Sub

Private Sub flexIncome_Click()
flexIncome.Sort = 1
End Sub

Private Sub Form_Load()
    dbconnection
    SSTab1.Tab = 9
    
End Sub
Private Sub Form_unLoad(cancel As Integer)
    Frame1.Visible = False
End Sub


Private Sub Label82_Click()
Form7.Show
End Sub


Private Sub list_Click()
Dim str As String
str = List.Text
CustomRecordSetOpen "select name,gender,dob,address,phone,cell,join_date,department from employee where emp_id=" & str
With res
txtName.Text = !Name
If !gender = "Male" Then
    Option1.value = True
Else
    Option2.value = True
End If
DTPickeremp.value = !dob
txtAddress.Text = !address
txtPhone.Text = !phone
txtCell.Text = !cell
DTPicker2.value = !join_date
Combo1.Text = !department
End With
CloseRecordSet

End Sub

Private Sub ListBank_Click()
   
     st = ListBank.Text
    str = "Select name,address,date,acc_no,type,total from tblBank where b_id=" & st
    CustomRecordSetOpen str
    With res
        txtBankName.Text = !Name
        txtBankAccNo.Text = !acc_no
        txtBankAddress.Text = !address
        combBankType.Text = !Type
        txtBankTotal.Text = !total
    End With
    CloseRecordSet
    
End Sub

Private Sub ListM_Click()
    st = ListM.Text
    str = "Select name,address,gender,phone,cell,join_date,type,acc_no from member where mmbr_id=" & st
    CustomRecordSetOpen str
    With res
        txtCName.Text = !Name
        txtCAddress.Text = !address
        txtCAccno.Text = !acc_no
        If !gender = "Female" Then optCFemale.value = True
        If !gender = "Male" Then optCMale.value = True
        txtCPhone.Text = !phone
        txtCCell.Text = !cell
        dateCDoj.value = !join_date
        combCType.Text = !Type
    End With
    CloseRecordSet
End Sub

Private Sub MSFlexGrid1_Click()
MSFlexGrid1.Sort = 1
End Sub

Private Sub note_Click()
Shell ("notepad.exe"), vbNormalFocus
End Sub

Private Sub Picture1_Click()

End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    Dim a As Integer
    Dim temp1 As Double
    Dim temp2 As Double
    st = "select sum(amount) as tot from tblPlLoanInterest"
    CustomRecordSetOpen st
    txtPlInInterest.Text = res!tot
    CloseRecordSet
    st = "select sum(interest) as tot from tblPlSaverInterest"
    CustomRecordSetOpen st
    txtPlExInterest.Text = res!tot
    CloseRecordSet
    
    st = "select sum(total) as tot from tblExpenseTrans where underBy like '" & "Indirect Expenses" & "'"
    CustomRecordSetOpen st
    txtplex.Text = res!tot
    CloseRecordSet
      st = "select sum(total) as tot from tblExpenseTrans where underBy like '" & "Direct Expenses" & "'"
    CustomRecordSetOpen st
    txtpldx.Text = res!tot
    CloseRecordSet
      st = "select sum(total) as tot from tblExpenseTrans where underBy like '" & "Purchase" & "'"
    CustomRecordSetOpen st
    txtplpam.Text = res!tot
    CloseRecordSet
    
    Dim myTotal As Double
    Dim flag As Boolean
    Dim ttotal As Double
    myTotal = 0
    flag = False
    
    CustomRecordSetOpen "tblIncomeTrans"
    
    With res
        If res.RecordCount <> 0 Then
            .MoveFirst
            While Not .EOF
            If !UnderBy = "Direct Income" Then
                myTotal = myTotal + !total
                flag = True
            End If
                .MoveNext
            Wend
        Else
            MsgBox "the table is empty"
            CloseRecordSet
            Exit Sub
        End If
        If flag = False Then
            txtPlDIn.Text = "0"
    
        Else
            txtPlDIn.Text = myTotal
        End If
        CloseRecordSet
        End With
    Dim myTotal1 As Double

    myTotal1 = 0
    flag = False
    CustomRecordSetOpen "tblIncomeTrans"
    
    With res
        If res.RecordCount <> 0 Then
            .MoveFirst
            While Not .EOF
            If !UnderBy = "Indirect Income" Then
                myTotal1 = myTotal1 + !total
                flag = True
            End If
                .MoveNext
            Wend
        Else
            MsgBox "the table is empty"
            CloseRecordSet
            Exit Sub
        End If
        If flag = False Then
            txtPlII.Text = "0"
    
        Else
            txtPlII.Text = myTotal1
        End If
        CloseRecordSet
        
    End With
    

    Dim myTotal2 As Double
    'Dim ttotal1 As Double
    myTotal2 = 0
    flag = False
    CustomRecordSetOpen "tblIncomeTrans"
    
    With res
        If res.RecordCount <> 0 Then
            .MoveFirst
            While Not .EOF
            If !UnderBy = "Other Income" Then
                myTotal2 = myTotal2 + !total
                flag = True
            End If
                .MoveNext
            Wend
        Else
            MsgBox "the table is empty"
            CloseRecordSet
            Exit Sub
        End If
        If flag = False Then
            txtPlOI.Text = "0"
    
        Else
            txtPlOI.Text = myTotal
        End If
        CloseRecordSet
        'Exit Sub
    End With
    
    
    Dim myTotal3 As Double
    'Dim ttotal1 As Double
    myTotal3 = 0
    flag = False
    CustomRecordSetOpen "tblIncomeTrans"
    
    With res
        If res.RecordCount <> 0 Then
            .MoveFirst
            While Not .EOF
            If !UnderBy = "Sales Income" Then
                myTotal3 = myTotal3 + !total
                flag = True
            End If
                .MoveNext
            Wend
        Else
            MsgBox "the table is empty"
            CloseRecordSet
            Exit Sub
        End If
        If flag = False Then
            txtPlSI.Text = "0"
    
        Else
            txtPlSI.Text = myTotal3
        End If
        CloseRecordSet
        'Exit Sub
    End With
    
    temp1 = Val(txtplex.Text) + Val(txtpldx.Text) + Val(txtplpam.Text) + Val(txtPlExInterest.Text)
    temp2 = Val(txtPlDIn.Text) + Val(txtPlII.Text) + Val(txtPlInInterest.Text) + Val(txtPlOI.Text) + Val(txtPlSI.Text)
    txttemppl.Text = temp2 - temp1
    If (txttemppl.Text > 0) Then
        lblpl.Caption = "Net Profit"
        txtpl.Text = txttemppl.Text
    Else
        lblpl.Caption = "Net Loss"
        txtpl.Text = txttemppl.Text
    End If
    st = "select sum(cash) as cashtot,sum(cheque) as chequetot from tblExpenseTrans"
    CustomRecordSetOpen st
    txtTtotal.Text = res!cashtot + res!chequetot
    CloseRecordSet
    st = "select sum(cash) as cashtot,sum(cheque) as chequetot from tblIncomeTrans"
    CustomRecordSetOpen st
    txtTIn.Text = res!cashtot + res!chequetot
    CloseRecordSet
    st = "select b_id from tblBank where type='Assets'"
    CustomRecordSetOpen st
    combTExBid.Clear    'Expenditure Bank Id
    combDepId.Clear     'Deposit Bank Id
    combWithId.Clear    'Withdraw Bank Id
    res.MoveFirst
    While Not res.EOF
        combTExBid.AddItem res!b_id
        combDepId.AddItem res!b_id
        combWithId.AddItem res!b_id
        res.MoveNext
    Wend
    CloseRecordSet
    st = "Select emp_id from employee"
    CustomRecordSetOpen st
    combTInEid.Clear
    res.MoveFirst
    While Not res.EOF
        combTInEid.AddItem res!emp_id 'For Income Employee Id
        res.MoveNext
    Wend
    CloseRecordSet
    st = "Select emp_id from employee"
    CustomRecordSetOpen st
    txtTExId.Clear
    res.MoveFirst
    While Not res.EOF
        txtTExId.AddItem res!emp_id 'For Expense Employee Id
        res.MoveNext
    Wend
    CloseRecordSet
    
    'For Expenditure Tab
     st = "Select Distinct underBy from tblExpenseTrans"
    CustomRecordSetOpen st
    combEx.Clear
    While Not res.EOF
        combEx.AddItem res!UnderBy
        res.MoveNext
    Wend
    CloseRecordSet
    'For Income combo groub
    st = "Select Distinct UnderBy from tblIncomeTrans"
    CustomRecordSetOpen st
    combIncome.Clear
    While Not res.EOF
        combIncome.AddItem res!UnderBy
        res.MoveNext
    Wend
    CloseRecordSet

    'for Income Bank Id
    st = "select b_id from tblBank where type='Assets'"
    CustomRecordSetOpen st
    combTInBid.Clear
    res.MoveFirst
    While Not res.EOF
        combTInBid.AddItem res!b_id
        res.MoveNext
    Wend
    CloseRecordSet
    st = "select distinct mmbr_id from tblsaver"
    CustomRecordSetOpen st
    comboloan.Clear
    res.MoveFirst
    While Not res.EOF
        comboloan.AddItem res!mmbr_id
        res.MoveNext
    Wend
    CloseRecordSet
    st = "Select mmbr_id from member"
    CustomRecordSetOpen st
    cmbsaving.Clear
    cmbShare.Clear
    res.MoveFirst
    While Not res.EOF
        cmbsaving.AddItem res!mmbr_id
        cmbShare.AddItem res!mmbr_id
        res.MoveNext
    Wend
    CloseRecordSet
     flex_display
    st = "select * from tblPL"
    CustomRecordSetOpen st
    
        res!particular = lblpl.Caption
        res!amount = txtpl.Text
        res.UpdateBatch
    CloseRecordSet
    
End Sub
Private Sub flex_display()
Dim sum As Integer
Dim a As Integer
     Dim data As Date
     data = DTPicker2.value
    MSFlexGrid2.Cols = 5
    MSFlexGrid2.TextMatrix(0, 1) = "Date"
    MSFlexGrid2.TextMatrix(0, 2) = "Particular"
    MSFlexGrid2.TextMatrix(0, 3) = "Debit"
    MSFlexGrid2.TextMatrix(0, 4) = "Credit"
    
   
   st = "select date,particular,total from tblExpenseTrans where date like '" & data & "'"
    CustomRecordSetOpen st
    a = 1
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = !particular
        MSFlexGrid2.TextMatrix(a, 3) = !total
        MSFlexGrid2.TextMatrix(a, 4) = ""
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet
st = "select date,name,amount from tblbankdeposit where date like '" & data & "'"
CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = !Name
        MSFlexGrid2.TextMatrix(a, 3) = !amount
        MSFlexGrid2.TextMatrix(a, 4) = ""
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet
 st = "select date,name,amount from tblbankwithdraw where date like '" & data & "'"
CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = !Name
        MSFlexGrid2.TextMatrix(a, 3) = ""
        MSFlexGrid2.TextMatrix(a, 4) = !amount
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet

st = "select date,mmbr_id,amount from loan where date like '" & data & "'"
CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = "Loan Paid To MId  " & "(" & !mmbr_id & ")"
        MSFlexGrid2.TextMatrix(a, 3) = !amount
        MSFlexGrid2.TextMatrix(a, 4) = ""
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet

st = "select Date,Particulars,Total from tblIncomeTrans where Date like '" & data & "'"
    CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = !Particulars
        MSFlexGrid2.TextMatrix(a, 3) = ""
        MSFlexGrid2.TextMatrix(a, 4) = !total
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet

st = "select date,mmbr_id,amount from tblsaver where date like '" & data & "'"
CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = "Save Amount Receipt MId by " & "(" & !mmbr_id & ")"
        MSFlexGrid2.TextMatrix(a, 3) = ""
        MSFlexGrid2.TextMatrix(a, 4) = !amount
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet

st = "select date,mmbr_id,amount from tblShare where date like '" & data & "'"
CustomRecordSetOpen st
    a = sum
    With res
         While (Not (.EOF))
        MSFlexGrid2.Rows = a + 1
        MSFlexGrid2.TextMatrix(a, 1) = !Date
        MSFlexGrid2.TextMatrix(a, 2) = "Share Receipt MId by " & "(" & !mmbr_id & ")"
        MSFlexGrid2.TextMatrix(a, 3) = ""
        MSFlexGrid2.TextMatrix(a, 4) = !amount
        .MoveNext
        a = a + 1
        sum = a
        Wend
    End With
CloseRecordSet

End Sub

Private Sub Timer1_Timer()
dateTEx = Date
End Sub

Private Sub Timer2_Timer()
dateTIn = Date
End Sub

Private Sub Timer3_Timer()
DTPicker2 = Date
End Sub

Private Sub Timer4_Timer()
    DTPickerdaybook = Date
End Sub

Private Sub Timer5_Timer()
DTPickerSaving = Date
End Sub

Private Sub Timer6_Timer()
DTPickerShare = Date
End Sub

Private Sub Timer7_Timer()
dateloan = Date
End Sub

Private Sub Timer8_Timer()
dateWith = Date
End Sub

Private Sub Timer9_Timer()
dateDep = Date
End Sub

Private Sub txtLoanAmount_LostFocus()
'Dim a As String
'a = comboloan.Text
st = "select sum(amount) as total from tblsaver where mmbr_id like '" & comboloan.Text & "'"
CustomRecordSetOpen st
    'MsgBox "Hello"
    Text1.Text = res!total
    If (Val(Text1.Text) < 20000) Then
        MsgBox "Sorry You haven't Access", vbInformation
        txtLoanAmount.Text = ""
        comboloan.SetFocus
        CloseRecordSet
        Exit Sub
    Else
        cmdLoanSave.SetFocus

    End If
CloseRecordSet
End Sub


Private Sub txtShareAmount_LostFocus()
    'Dim a As Double
    'a = txtShareAmount.Text
    If (Val(txtShareAmount.Text) < 300) Or (Val(txtShareAmount.Text) > 5000) Then
        MsgBox "Enter amount between 300 and 5000", vbInformation
        txtShareAmount.Text = ""
        'txtShareAmount.SetFocus
    Exit Sub
    End If
        
End Sub

Private Sub txtTExCheque_LostFocus()

    st = combTExBid.Text
    str = "Select total from tblBank where b_id= " & st
    CustomRecordSetOpen str
    If Val(txtTExCheque.Text) <= res!total Then
        MsgBox "Amount is Available...", vbInformation
        txtTExTotal = Val(txtTExCash.Text) + Val(txtTExCheque.Text)
    Else
        MsgBox "Amount is not availble"
        txtTExCheque = ""
        txtTExCheque.SetFocus
    End If
CloseRecordSet
End Sub



Private Sub txtTIncheque_LostFocus()
    txtTExTotal = Val(txtTExCash.Text) + Val(txtTExCheque.Text)
    txtTInTotal.Text = Val(txtTIncash.Text) + Val(txtTIncheque.Text)
End Sub
