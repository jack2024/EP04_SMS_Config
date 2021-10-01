VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmMainConfig 
   Caption         =   "Form1"
   ClientHeight    =   11370
   ClientLeft      =   2715
   ClientTop       =   945
   ClientWidth     =   18210
   Icon            =   "frmMainConfig.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   11370
   ScaleWidth      =   18210
   Begin VB.Frame Frame8 
      Caption         =   "Telephone Number"
      Height          =   975
      Left            =   8640
      TabIndex        =   182
      Top             =   0
      Width           =   1815
      Begin VB.TextBox txtphone 
         Height          =   375
         Left            =   120
         MaxLength       =   15
         TabIndex        =   183
         Text            =   "0123456789"
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Input Delay"
      Height          =   975
      Left            =   16560
      TabIndex        =   138
      Top             =   0
      Width           =   1455
      Begin VB.CommandButton TimeAllChange 
         Caption         =   "ChangeAll"
         Height          =   315
         Left            =   180
         TabIndex        =   181
         Top             =   600
         Width           =   1095
      End
      Begin VB.ComboBox cbbFaultDelayTime 
         Height          =   315
         Left            =   240
         TabIndex        =   139
         Text            =   "Combo1"
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "Both(Non)"
      Height          =   375
      Left            =   4320
      TabIndex        =   137
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command10 
      Caption         =   "Both(All)"
      Height          =   375
      Left            =   4320
      TabIndex        =   136
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame6 
      Caption         =   "RGA(All)"
      Height          =   975
      Left            =   5280
      TabIndex        =   132
      Top             =   0
      Width           =   855
      Begin VB.OptionButton Option3 
         Caption         =   "A"
         Height          =   195
         Left            =   120
         TabIndex        =   135
         Top             =   720
         Width           =   495
      End
      Begin VB.OptionButton Gall 
         Caption         =   "G"
         Height          =   255
         Left            =   120
         TabIndex        =   134
         Top             =   480
         Width           =   375
      End
      Begin VB.OptionButton Rall 
         Caption         =   "R"
         Height          =   255
         Left            =   120
         TabIndex        =   133
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Alarm (All)"
      Height          =   375
      Left            =   240
      TabIndex        =   131
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Indicator (All)"
      Height          =   375
      Left            =   240
      TabIndex        =   130
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "NO (All)"
      Height          =   375
      Left            =   1260
      TabIndex        =   129
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "NC (All)"
      Height          =   375
      Left            =   1260
      TabIndex        =   128
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Manual (All)"
      Height          =   375
      Left            =   2280
      TabIndex        =   127
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Auto (All)"
      Height          =   375
      Left            =   2280
      TabIndex        =   126
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Buzzer (All)"
      Height          =   375
      Left            =   3300
      TabIndex        =   125
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Bell (All)"
      Height          =   375
      Left            =   3300
      TabIndex        =   124
      Top             =   600
      Width           =   975
   End
   Begin VB.Frame Frame4 
      Caption         =   "Modbus Setting"
      Height          =   975
      Left            =   14040
      TabIndex        =   14
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         Height          =   495
         Left            =   1680
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TextAddrNew 
         Height          =   375
         Left            =   1200
         TabIndex        =   16
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.TextBox TextAddrNow 
         Height          =   375
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   15
         Text            =   "1"
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label2 
         Caption         =   "Current     Addr."
         Height          =   495
         Left            =   600
         TabIndex        =   18
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "No.Of Point."
      Height          =   975
      Left            =   6120
      TabIndex        =   12
      Top             =   0
      Width           =   1215
      Begin VB.OptionButton OptNoInput20 
         Caption         =   "20"
         Height          =   255
         Left            =   600
         TabIndex        =   123
         Top             =   520
         Width           =   495
      End
      Begin VB.OptionButton OptNoInput10 
         Caption         =   "10"
         Height          =   255
         Left            =   120
         TabIndex        =   122
         Top             =   520
         Width           =   495
      End
      Begin VB.OptionButton OptNoInput8 
         Caption         =   "8"
         Height          =   255
         Left            =   120
         TabIndex        =   121
         Top             =   240
         Width           =   495
      End
      Begin VB.OptionButton OptNoInput16 
         Caption         =   "16"
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Flash Timing"
      Height          =   975
      Left            =   10440
      TabIndex        =   2
      Top             =   0
      Width           =   3615
      Begin VB.OptionButton OptflashingRate1000 
         Caption         =   "1000ms"
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OptflashingRate875 
         Caption         =   "875ms"
         Height          =   375
         Left            =   1800
         TabIndex        =   10
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OptflashingRate750 
         Caption         =   "750ms"
         Height          =   375
         Left            =   960
         TabIndex        =   9
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OptflashingRate625 
         Caption         =   "625ms"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   855
      End
      Begin VB.OptionButton OptflashingRate500 
         Caption         =   "500ms"
         Height          =   375
         Left            =   2640
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptflashingRate375 
         Caption         =   "375ms"
         Height          =   375
         Left            =   1800
         TabIndex        =   6
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptflashingRate250 
         Caption         =   "250ms"
         Height          =   375
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   855
      End
      Begin VB.OptionButton OptflashingRate125 
         Caption         =   "125ms"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Auto Ack."
      Height          =   975
      Left            =   7200
      TabIndex        =   0
      Top             =   0
      Width           =   1455
      Begin VB.ComboBox cbbAutoAckTime 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Text            =   "Combo1"
         Top             =   590
         Width           =   735
      End
      Begin VB.CheckBox chkAutoAck 
         Caption         =   "AutoAck."
         Height          =   285
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   1000
      End
      Begin VB.Label Label1 
         Caption         =   "Sec."
         Height          =   255
         Left            =   1005
         TabIndex        =   184
         Top             =   600
         Width           =   360
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10215
      Left            =   240
      TabIndex        =   19
      Top             =   1080
      Width           =   17895
      _ExtentX        =   31565
      _ExtentY        =   18018
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "INPUT 1-16"
      TabPicture(0)   =   "frmMainConfig.frx":2072
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame5"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "VScroll2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.VScrollBar VScroll2 
         Height          =   9735
         LargeChange     =   960
         Left            =   17280
         Max             =   960
         SmallChange     =   960
         TabIndex        =   140
         Top             =   360
         Width           =   495
      End
      Begin VB.Frame Frame5 
         Height          =   9855
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Width           =   17175
         Begin VB.Frame FrameT1 
            Caption         =   "Input Delay"
            Height          =   975
            Left            =   7440
            TabIndex        =   179
            Top             =   120
            Width           =   1095
            Begin VB.ComboBox ComboT1 
               Height          =   315
               Left            =   120
               TabIndex        =   180
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT20 
            Height          =   975
            Left            =   15960
            TabIndex        =   177
            Top             =   8760
            Width           =   1095
            Begin VB.ComboBox ComboT20 
               Height          =   315
               Left            =   120
               TabIndex        =   178
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT19 
            Height          =   975
            Left            =   15960
            TabIndex        =   175
            Top             =   7800
            Width           =   1095
            Begin VB.ComboBox ComboT19 
               Height          =   315
               Left            =   120
               TabIndex        =   176
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT18 
            Height          =   975
            Left            =   15960
            TabIndex        =   173
            Top             =   6840
            Width           =   1095
            Begin VB.ComboBox ComboT18 
               Height          =   315
               Left            =   120
               TabIndex        =   174
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT17 
            Height          =   975
            Left            =   15960
            TabIndex        =   171
            Top             =   5880
            Width           =   1095
            Begin VB.ComboBox ComboT17 
               Height          =   315
               Left            =   120
               TabIndex        =   172
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT16 
            Height          =   975
            Left            =   15960
            TabIndex        =   169
            Top             =   4920
            Width           =   1095
            Begin VB.ComboBox ComboT16 
               Height          =   315
               Left            =   120
               TabIndex        =   170
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT15 
            Height          =   975
            Left            =   15960
            TabIndex        =   167
            Top             =   3960
            Width           =   1095
            Begin VB.ComboBox ComboT15 
               Height          =   315
               Left            =   120
               TabIndex        =   168
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT14 
            Height          =   975
            Left            =   15960
            TabIndex        =   165
            Top             =   3000
            Width           =   1095
            Begin VB.ComboBox ComboT14 
               Height          =   315
               Left            =   120
               TabIndex        =   166
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT13 
            Height          =   975
            Left            =   15960
            TabIndex        =   163
            Top             =   2040
            Width           =   1095
            Begin VB.ComboBox ComboT13 
               Height          =   315
               Left            =   120
               TabIndex        =   164
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT12 
            Height          =   975
            Left            =   15960
            TabIndex        =   161
            Top             =   1080
            Width           =   1095
            Begin VB.ComboBox ComboT12 
               Height          =   315
               Left            =   120
               TabIndex        =   162
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT11 
            Caption         =   "Input Delay"
            Height          =   975
            Left            =   15960
            TabIndex        =   159
            Top             =   120
            Width           =   1095
            Begin VB.ComboBox ComboT11 
               Height          =   315
               Left            =   120
               TabIndex        =   160
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT10 
            Height          =   975
            Left            =   7440
            TabIndex        =   157
            Top             =   8760
            Width           =   1095
            Begin VB.ComboBox ComboT10 
               Height          =   315
               Left            =   120
               TabIndex        =   158
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT9 
            Height          =   975
            Left            =   7440
            TabIndex        =   155
            Top             =   7800
            Width           =   1095
            Begin VB.ComboBox ComboT9 
               Height          =   315
               Left            =   120
               TabIndex        =   156
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT8 
            Height          =   975
            Left            =   7440
            TabIndex        =   153
            Top             =   6840
            Width           =   1095
            Begin VB.ComboBox ComboT8 
               Height          =   315
               Left            =   120
               TabIndex        =   154
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT7 
            Height          =   975
            Left            =   7440
            TabIndex        =   151
            Top             =   5880
            Width           =   1095
            Begin VB.ComboBox ComboT7 
               Height          =   315
               Left            =   120
               TabIndex        =   152
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT6 
            Height          =   975
            Left            =   7440
            TabIndex        =   149
            Top             =   4920
            Width           =   1095
            Begin VB.ComboBox ComboT6 
               Height          =   315
               Left            =   120
               TabIndex        =   150
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT5 
            Height          =   975
            Left            =   7440
            TabIndex        =   147
            Top             =   3960
            Width           =   1095
            Begin VB.ComboBox ComboT5 
               Height          =   315
               Left            =   120
               TabIndex        =   148
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT4 
            Height          =   975
            Left            =   7440
            TabIndex        =   145
            Top             =   3000
            Width           =   1095
            Begin VB.ComboBox ComboT4 
               Height          =   315
               Left            =   120
               TabIndex        =   146
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT3 
            Height          =   975
            Left            =   7440
            TabIndex        =   143
            Top             =   2040
            Width           =   1095
            Begin VB.ComboBox ComboT3 
               Height          =   315
               Left            =   120
               TabIndex        =   144
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameT2 
            Height          =   975
            Left            =   7440
            TabIndex        =   141
            Top             =   1080
            Width           =   1095
            Begin VB.ComboBox ComboT2 
               Height          =   315
               Left            =   120
               TabIndex        =   142
               Text            =   "Combo1"
               Top             =   360
               Width           =   855
            End
         End
         Begin VB.Frame FrameRGA20 
            Height          =   975
            Left            =   14280
            TabIndex        =   117
            Top             =   8760
            Width           =   1695
            Begin VB.OptionButton A20 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   120
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G20 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   119
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R20 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   118
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA19 
            Height          =   975
            Left            =   14280
            TabIndex        =   113
            Top             =   7800
            Width           =   1695
            Begin VB.OptionButton A19 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   116
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G19 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   115
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R19 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   114
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA18 
            Height          =   975
            Left            =   14280
            TabIndex        =   109
            Top             =   6840
            Width           =   1695
            Begin VB.OptionButton A18 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   112
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G18 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   111
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R18 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   110
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA17 
            Height          =   975
            Left            =   14280
            TabIndex        =   105
            Top             =   5880
            Width           =   1695
            Begin VB.OptionButton A17 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   108
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G17 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   107
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R17 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   106
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA16 
            Height          =   975
            Left            =   14280
            TabIndex        =   101
            Top             =   4920
            Width           =   1695
            Begin VB.OptionButton A16 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   104
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G16 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   103
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R16 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   102
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA15 
            Height          =   975
            Left            =   14280
            TabIndex        =   97
            Top             =   3960
            Width           =   1695
            Begin VB.OptionButton A15 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   100
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G15 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   99
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R15 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   98
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA14 
            Height          =   975
            Left            =   14280
            TabIndex        =   93
            Top             =   3000
            Width           =   1695
            Begin VB.OptionButton A14 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   96
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G14 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   95
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R14 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   94
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA13 
            Height          =   975
            Left            =   14280
            TabIndex        =   89
            Top             =   2040
            Width           =   1695
            Begin VB.OptionButton A13 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   92
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G13 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   91
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R13 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   90
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA12 
            Height          =   975
            Left            =   14280
            TabIndex        =   85
            Top             =   1080
            Width           =   1695
            Begin VB.OptionButton R12 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   88
               Top             =   300
               Width           =   495
            End
            Begin VB.OptionButton G12 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   87
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton A12 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   86
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA11 
            Height          =   975
            Left            =   14280
            TabIndex        =   81
            Top             =   120
            Width           =   1695
            Begin VB.OptionButton R11 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   84
               Top             =   300
               Width           =   495
            End
            Begin VB.OptionButton G11 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   83
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton A11 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   82
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA10 
            Height          =   975
            Left            =   5760
            TabIndex        =   66
            Top             =   8760
            Width           =   1695
            Begin VB.OptionButton A10 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   69
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G10 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   68
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R10 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   67
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA9 
            Height          =   975
            Left            =   5760
            TabIndex        =   61
            Top             =   7800
            Width           =   1695
            Begin VB.OptionButton R9 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   64
               Top             =   300
               Width           =   495
            End
            Begin VB.OptionButton G9 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   63
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton A9 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   62
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Timer Timer1 
            Interval        =   1000
            Left            =   480
            Top             =   0
         End
         Begin VB.Frame FrameRGA1 
            Height          =   975
            Left            =   5760
            TabIndex        =   49
            Top             =   120
            Width           =   1695
            Begin VB.OptionButton R1 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   52
               Top             =   300
               Width           =   495
            End
            Begin VB.OptionButton G1 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   51
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton A1 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   50
               Top             =   360
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA2 
            Height          =   975
            Left            =   5760
            TabIndex        =   45
            Top             =   1080
            Width           =   1695
            Begin VB.OptionButton A2 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   48
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G2 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   47
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R2 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   46
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA3 
            Height          =   975
            Left            =   5760
            TabIndex        =   41
            Top             =   2040
            Width           =   1695
            Begin VB.OptionButton A3 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   44
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G3 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   43
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R3 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   42
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA4 
            Height          =   975
            Left            =   5760
            TabIndex        =   37
            Top             =   3000
            Width           =   1695
            Begin VB.OptionButton A4 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   40
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G4 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   39
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R4 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   38
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA5 
            Height          =   975
            Left            =   5760
            TabIndex        =   33
            Top             =   3960
            Width           =   1695
            Begin VB.OptionButton A5 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   36
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G5 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   35
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R5 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   34
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA6 
            Height          =   975
            Left            =   5760
            TabIndex        =   29
            Top             =   4920
            Width           =   1695
            Begin VB.OptionButton A6 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   32
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G6 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   31
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R6 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   30
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA7 
            Height          =   975
            Left            =   5760
            TabIndex        =   25
            Top             =   5880
            Width           =   1695
            Begin VB.OptionButton A7 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   28
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G7 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   27
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R7 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   26
               Top             =   300
               Width           =   495
            End
         End
         Begin VB.Frame FrameRGA8 
            Height          =   975
            Left            =   5760
            TabIndex        =   21
            Top             =   6840
            Width           =   1695
            Begin VB.OptionButton A8 
               Caption         =   "A"
               Height          =   375
               Left            =   1080
               TabIndex        =   24
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton G8 
               Caption         =   "G"
               Height          =   375
               Left            =   600
               TabIndex        =   23
               Top             =   360
               Width           =   495
            End
            Begin VB.OptionButton R8 
               Caption         =   "R"
               Height          =   495
               Left            =   120
               TabIndex        =   22
               Top             =   300
               Width           =   495
            End
         End
         Begin MSCommLib.MSComm MSComm1 
            Left            =   -120
            Top             =   0
            _ExtentX        =   1005
            _ExtentY        =   1005
            _Version        =   393216
            DTREnable       =   -1  'True
         End
         Begin ESPAN04.InputProperties InputP8 
            Height          =   975
            Left            =   120
            TabIndex        =   53
            Top             =   6840
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP7 
            Height          =   1095
            Left            =   120
            TabIndex        =   54
            Top             =   5880
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1931
         End
         Begin ESPAN04.InputProperties InputP6 
            Height          =   1095
            Left            =   120
            TabIndex        =   55
            Top             =   4920
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   1931
         End
         Begin ESPAN04.InputProperties InputP5 
            Height          =   975
            Left            =   120
            TabIndex        =   56
            Top             =   3960
            Width           =   5775
            _ExtentX        =   10186
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP4 
            Height          =   1095
            Left            =   120
            TabIndex        =   57
            Top             =   3000
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1931
         End
         Begin ESPAN04.InputProperties InputP3 
            Height          =   975
            Left            =   120
            TabIndex        =   58
            Top             =   2040
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP2 
            Height          =   1095
            Left            =   120
            TabIndex        =   59
            Top             =   1080
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1931
         End
         Begin ESPAN04.InputProperties InputP1 
            Height          =   975
            Left            =   120
            TabIndex        =   60
            Top             =   120
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP9 
            Height          =   975
            Left            =   120
            TabIndex        =   65
            Top             =   7800
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP10 
            Height          =   975
            Left            =   120
            TabIndex        =   70
            Top             =   8760
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP14 
            Height          =   975
            Left            =   8640
            TabIndex        =   71
            Top             =   3000
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP13 
            Height          =   975
            Left            =   8640
            TabIndex        =   72
            Top             =   2040
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP12 
            Height          =   1095
            Left            =   8640
            TabIndex        =   73
            Top             =   1080
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1931
         End
         Begin ESPAN04.InputProperties InputP11 
            Height          =   975
            Left            =   8640
            TabIndex        =   74
            Top             =   120
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP18 
            Height          =   975
            Left            =   8640
            TabIndex        =   75
            Top             =   6840
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP17 
            Height          =   975
            Left            =   8640
            TabIndex        =   76
            Top             =   5880
            Width           =   5895
            _ExtentX        =   10398
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP16 
            Height          =   1095
            Left            =   8640
            TabIndex        =   77
            Top             =   4920
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1931
         End
         Begin ESPAN04.InputProperties InputP15 
            Height          =   975
            Left            =   8640
            TabIndex        =   78
            Top             =   3960
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP19 
            Height          =   975
            Left            =   8640
            TabIndex        =   79
            Top             =   7800
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1720
         End
         Begin ESPAN04.InputProperties InputP20 
            Height          =   975
            Left            =   8640
            TabIndex        =   80
            Top             =   8760
            Width           =   5655
            _ExtentX        =   9975
            _ExtentY        =   1720
         End
      End
   End
   Begin VB.Line Line1 
      X1              =   6120
      X2              =   7320
      Y1              =   5040
      Y2              =   5520
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuComSetting 
      Caption         =   "Communication Setting"
   End
   Begin VB.Menu mnuDownloadUpload 
      Caption         =   "Read/Write Config"
      Begin VB.Menu mnuDownload 
         Caption         =   "Write Config"
      End
      Begin VB.Menu mnuUpload 
         Caption         =   "Read Config"
      End
   End
   Begin VB.Menu CreateLabelmenu 
      Caption         =   "CreateLabel"
   End
   Begin VB.Menu FaultNamemenu 
      Caption         =   "Faultname"
   End
End
Attribute VB_Name = "frmMainConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Dim CRCTable(0 To 511) As Byte
Dim CRC_Low, CRC_High As Byte

Dim InputType1_8 As Byte
Dim InputType9_16 As Byte
Dim InputType17_24 As Byte
Dim InputType25_32 As Byte
Dim InputType33_40 As Byte
Dim InputType41_48 As Byte
Dim InputType49_56 As Byte
Dim InputType57_64 As Byte

Dim FaultType1_8 As Byte
Dim FaultType9_16 As Byte
Dim FaultType17_24 As Byte
Dim FaultType25_32 As Byte
Dim FaultType33_40 As Byte
Dim FaultType41_48 As Byte
Dim FaultType49_56 As Byte
Dim FaultType57_64 As Byte

Dim OutputType1_8 As Byte
Dim OutputType9_16 As Byte
Dim OutputType17_24 As Byte
Dim OutputType25_32 As Byte
Dim OutputType33_40 As Byte
Dim OutputType41_48 As Byte
Dim OutputType49_56 As Byte
Dim OutputType57_64 As Byte

Dim Fault_Indicator1_8 As Byte
Dim Fault_Indicator9_16 As Byte
Dim Fault_Indicator17_24 As Byte
Dim Fault_Indicator25_32 As Byte
Dim Fault_Indicator33_40 As Byte
Dim Fault_Indicator41_48 As Byte
Dim Fault_Indicator49_56 As Byte
Dim Fault_Indicator57_64 As Byte

Dim OutputBoth1_8 As Byte
Dim OutputBoth9_16 As Byte
Dim OutputBoth17_24 As Byte
Dim OutputBoth25_32 As Byte
Dim OutputBoth33_40 As Byte
Dim OutputBoth41_48 As Byte
Dim OutputBoth49_56 As Byte
Dim OutputBoth57_64 As Byte

Dim AutoAckStatus As Byte
Dim AutoAckTime As Byte
Dim FlashRate As Byte
Dim NoOfInput As Byte
Dim MasterSlave As Byte
Dim FaultDelayTime As Byte

Dim DelayTime1 As Byte
Dim DelayTime2 As Byte
Dim DelayTime3 As Byte
Dim DelayTime4 As Byte
Dim DelayTime5 As Byte
Dim DelayTime6 As Byte
Dim DelayTime7 As Byte
Dim DelayTime8 As Byte
Dim DelayTime9 As Byte
Dim DelayTime10 As Byte
Dim DelayTime11 As Byte
Dim DelayTime12 As Byte
Dim DelayTime13 As Byte
Dim DelayTime14 As Byte
Dim DelayTime15 As Byte
Dim DelayTime16 As Byte
Dim DelayTime17 As Byte
Dim DelayTime18 As Byte
Dim DelayTime19 As Byte
Dim DelayTime20 As Byte

Dim Red1_8 As Byte
Dim Red9_10 As Byte
Dim Red11_18 As Byte
Dim Red19_20 As Byte

Dim Green1_8 As Byte
Dim Green9_10 As Byte
Dim Green11_18 As Byte
Dim Green19_20 As Byte

Dim Temp_ColourRED As Byte
Dim Temp_ColourGREEN As Byte


Dim addr As Byte

Dim Buffer As String

Public phonenum As String


Dim DatBuff(80) As Byte
Public CommState As String
Dim Timeout As Integer
Dim CommBuff(4) As Byte



Private Sub FaultNamemenu_Click()
    SMSsetting.Show
End Sub





Private Sub TimeAllChange_Click()
    ComboT1.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT2.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT3.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT4.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT5.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT6.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT7.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT8.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT9.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT10.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT11.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT12.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT13.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT14.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT15.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT16.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT17.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT18.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT19.ListIndex = cbbFaultDelayTime.ListIndex
    ComboT20.ListIndex = cbbFaultDelayTime.ListIndex
End Sub

Private Sub chkAutoAck_Click()
    If chkAutoAck.Value = 1 Then
        cbbAutoAckTime.Enabled = True
        'Label1.Enabled = True
        cbbAutoAckTime.ListIndex = 0
    Else
        cbbAutoAckTime.Enabled = False
        'Label1.Enabled = False
        cbbAutoAckTime.ListIndex = -1
    End If
    
End Sub

Private Sub Command1_Click()
Dim Data As String
    If Not MSComm1.PortOpen Then
                MSComm1.CommPort = GetSetting("ESPAN-01", "Setting", "Comm Port", "1")
                MSComm1.Settings = "9600,N,8,1"
                MSComm1.PortOpen = True
                MSComm1.RThreshold = 1
            End If
        
            Data = ""
            Data = Chr(Val(TextAddrNow.Text)) + Chr(5) + Chr(0) + Chr(100) + Chr(0) + Chr(Val(TextAddrNew.Text))
            CRC_16 Data, 6
            Data = Data + Chr(CRC_High) + Chr(CRC_Low)
            MSComm1.InputLen = 0
            MSComm1.Output = Data
            TextAddrNow.Text = TextAddrNew.Text
            Buffer = ""
            Do While MSComm1.OutBufferCount > 0
            Loop
            MsgBox "DOWNLOAD COMPLETED", vbInformation, "ESPAN-01"
End Sub

Private Sub CreateLabelmenu_Click()
    Form20Point.Show
End Sub

Private Sub Form_Load()

Dim i As Integer
    
    InputP1.Caption = "INPUT 1"
    InputP2.Caption = "INPUT 2"
    InputP3.Caption = "INPUT 3"
    InputP4.Caption = "INPUT 4"
    InputP5.Caption = "INPUT 5"
    InputP6.Caption = "INPUT 6"
    InputP7.Caption = "INPUT 7"
    InputP8.Caption = "INPUT 8"
    InputP9.Caption = "INPUT 9"
    InputP10.Caption = "INPUT 10"
    InputP11.Caption = "INPUT 11"
    InputP12.Caption = "INPUT 12"
    InputP13.Caption = "INPUT 13"
    InputP14.Caption = "INPUT 14"
    InputP15.Caption = "INPUT 15"
    InputP16.Caption = "INPUT 16"
    
    InputP17.Caption = "INPUT 17"
    InputP18.Caption = "INPUT 18"
    InputP19.Caption = "INPUT 19"
    InputP20.Caption = "INPUT 20"
    
    
    
    
    InputP9.Visible = True
    InputP10.Visible = True
    InputP11.Visible = True
    InputP12.Visible = True
    InputP13.Visible = True
    InputP14.Visible = True
    InputP15.Visible = True
    InputP16.Visible = True
    InputP17.Visible = True
    InputP18.Visible = True
    InputP19.Visible = True
    InputP20.Visible = True
    
    FrameRGA9.Visible = True
    FrameRGA10.Visible = True
    FrameRGA11.Visible = True
    FrameRGA12.Visible = True
    FrameRGA13.Visible = True
    FrameRGA14.Visible = True
    FrameRGA15.Visible = True
    FrameRGA16.Visible = True
    FrameRGA17.Visible = True
    FrameRGA18.Visible = True
    FrameRGA19.Visible = True
    FrameRGA20.Visible = True
    
    Rall.Value = True
    R1.Value = True
    R2.Value = True
    R3.Value = True
    R4.Value = True
    R5.Value = True
    R6.Value = True
    R7.Value = True
    R8.Value = True
    R9.Value = True
    R10.Value = True
    R11.Value = True
    R12.Value = True
    R13.Value = True
    R14.Value = True
    R15.Value = True
    R16.Value = True
    R17.Value = True
    R18.Value = True
    R19.Value = True
    R20.Value = True
    
    OptflashingRate250.Value = True
    'OptSyncMaster.Value = True
      
    Me.Caption = "ESPAN-04"
    
    OptNoInput8.Enabled = True
    OptNoInput16.Enabled = True
    
    
    
    'Shape1.BackColor = vbGreen
    CRCTable(0) = &H0
    CRCTable(1) = &HC1
    CRCTable(2) = &H81
    CRCTable(3) = &H40
    CRCTable(4) = &H1
    CRCTable(5) = &HC0
    CRCTable(6) = &H80
    CRCTable(7) = &H41
    CRCTable(8) = &H1
    CRCTable(9) = &HC0
    CRCTable(10) = &H80
    CRCTable(11) = &H41
    CRCTable(12) = &H0
    CRCTable(13) = &HC1
    CRCTable(14) = &H81
    CRCTable(15) = &H40
    CRCTable(16) = &H1
    CRCTable(17) = &HC0
    CRCTable(18) = &H80
    CRCTable(19) = &H41
    CRCTable(20) = &H0
    CRCTable(21) = &HC1
    CRCTable(22) = &H81
    CRCTable(23) = &H40
    CRCTable(24) = &H0
    CRCTable(25) = &HC1
    CRCTable(26) = &H81
    CRCTable(27) = &H40
    CRCTable(28) = &H1
    CRCTable(29) = &HC0
    CRCTable(30) = &H80
    CRCTable(31) = &H41
    CRCTable(32) = &H1
    CRCTable(33) = &HC0
    CRCTable(34) = &H80
    CRCTable(35) = &H41
    CRCTable(36) = &H0
    CRCTable(37) = &HC1
    CRCTable(38) = &H81
    CRCTable(39) = &H40
    CRCTable(40) = &H0
    CRCTable(41) = &HC1
    CRCTable(42) = &H81
    CRCTable(43) = &H40
    CRCTable(44) = &H1
    CRCTable(45) = &HC0
    CRCTable(46) = &H80
    CRCTable(47) = &H41
    CRCTable(48) = &H0
    CRCTable(49) = &HC1
    CRCTable(50) = &H81
    CRCTable(51) = &H40
    CRCTable(52) = &H1
    CRCTable(53) = &HC0
    CRCTable(54) = &H80
    CRCTable(55) = &H41
    CRCTable(56) = &H1
    CRCTable(57) = &HC0
    CRCTable(58) = &H80
    CRCTable(59) = &H41
    CRCTable(60) = &H0
    CRCTable(61) = &HC1
    CRCTable(62) = &H81
    CRCTable(63) = &H40
    CRCTable(64) = &H1
    CRCTable(65) = &HC0
    CRCTable(66) = &H80
    CRCTable(67) = &H41
    CRCTable(68) = &H0
    CRCTable(69) = &HC1
    CRCTable(70) = &H81
    CRCTable(71) = &H40
    CRCTable(72) = &H0
    CRCTable(73) = &HC1
    CRCTable(74) = &H81
    CRCTable(75) = &H40
    CRCTable(76) = &H1
    CRCTable(77) = &HC0
    CRCTable(78) = &H80
    CRCTable(79) = &H41
    CRCTable(80) = &H0
    CRCTable(81) = &HC1
    CRCTable(82) = &H81
    CRCTable(83) = &H40
    CRCTable(84) = &H1
    CRCTable(85) = &HC0
    CRCTable(86) = &H80
    CRCTable(87) = &H41
    CRCTable(88) = &H1
    CRCTable(89) = &HC0
    CRCTable(90) = &H80
    CRCTable(91) = &H41
    CRCTable(92) = &H0
    CRCTable(93) = &HC1
    CRCTable(94) = &H81
    CRCTable(95) = &H40
    CRCTable(96) = &H0
    CRCTable(97) = &HC1
    CRCTable(98) = &H81
    CRCTable(99) = &H40
    CRCTable(100) = &H1
    CRCTable(101) = &HC0
    CRCTable(102) = &H80
    CRCTable(103) = &H41
    CRCTable(104) = &H1
    CRCTable(105) = &HC0
    CRCTable(106) = &H80
    CRCTable(107) = &H41
    CRCTable(108) = &H0
    CRCTable(109) = &HC1
    CRCTable(110) = &H81
    CRCTable(111) = &H40
    CRCTable(112) = &H1
    CRCTable(113) = &HC0
    CRCTable(114) = &H80
    CRCTable(115) = &H41
    CRCTable(116) = &H0
    CRCTable(117) = &HC1
    CRCTable(118) = &H81
    CRCTable(119) = &H40
    CRCTable(120) = &H0
    CRCTable(121) = &HC1
    CRCTable(122) = &H81
    CRCTable(123) = &H40
    CRCTable(124) = &H1
    CRCTable(125) = &HC0
    CRCTable(126) = &H80
    CRCTable(127) = &H41
    CRCTable(128) = &H1
    CRCTable(129) = &HC0
    CRCTable(130) = &H80
    CRCTable(131) = &H41
    CRCTable(132) = &H0
    CRCTable(133) = &HC1
    CRCTable(134) = &H81
    CRCTable(135) = &H40
    CRCTable(136) = &H0
    CRCTable(137) = &HC1
    CRCTable(138) = &H81
    CRCTable(139) = &H40
    CRCTable(140) = &H1
    CRCTable(141) = &HC0
    CRCTable(142) = &H80
    CRCTable(143) = &H41
    CRCTable(144) = &H0
    CRCTable(145) = &HC1
    CRCTable(146) = &H81
    CRCTable(147) = &H40
    CRCTable(148) = &H1
    CRCTable(149) = &HC0
    CRCTable(150) = &H80
    CRCTable(151) = &H41
    CRCTable(152) = &H1
    CRCTable(153) = &HC0
    CRCTable(154) = &H80
    CRCTable(155) = &H41
    CRCTable(156) = &H0
    CRCTable(157) = &HC1
    CRCTable(158) = &H81
    CRCTable(159) = &H40
    CRCTable(160) = &H0
    CRCTable(161) = &HC1
    CRCTable(162) = &H81
    CRCTable(163) = &H40
    CRCTable(164) = &H1
    CRCTable(165) = &HC0
    CRCTable(166) = &H80
    CRCTable(167) = &H41
    CRCTable(168) = &H1
    CRCTable(169) = &HC0
    CRCTable(170) = &H80
    CRCTable(171) = &H41
    CRCTable(172) = &H0
    CRCTable(173) = &HC1
    CRCTable(174) = &H81
    CRCTable(175) = &H40
    CRCTable(176) = &H1
    CRCTable(177) = &HC0
    CRCTable(178) = &H80
    CRCTable(179) = &H41
    CRCTable(180) = &H0
    CRCTable(181) = &HC1
    CRCTable(182) = &H81
    CRCTable(183) = &H40
    CRCTable(184) = &H0
    CRCTable(185) = &HC1
    CRCTable(186) = &H81
    CRCTable(187) = &H40
    CRCTable(188) = &H1
    CRCTable(189) = &HC0
    CRCTable(190) = &H80
    CRCTable(191) = &H41
    CRCTable(192) = &H0
    CRCTable(193) = &HC1
    CRCTable(194) = &H81
    CRCTable(195) = &H40
    CRCTable(196) = &H1
    CRCTable(197) = &HC0
    CRCTable(198) = &H80
    CRCTable(199) = &H41
    CRCTable(200) = &H1
    CRCTable(201) = &HC0
    CRCTable(202) = &H80
    CRCTable(203) = &H41
    CRCTable(204) = &H0
    CRCTable(205) = &HC1
    CRCTable(206) = &H81
    CRCTable(207) = &H40
    CRCTable(208) = &H1
    CRCTable(209) = &HC0
    CRCTable(210) = &H80
    CRCTable(211) = &H41
    CRCTable(212) = &H0
    CRCTable(213) = &HC1
    CRCTable(214) = &H81
    CRCTable(215) = &H40
    CRCTable(216) = &H0
    CRCTable(217) = &HC1
    CRCTable(218) = &H81
    CRCTable(219) = &H40
    CRCTable(220) = &H1
    CRCTable(221) = &HC0
    CRCTable(222) = &H80
    CRCTable(223) = &H41
    CRCTable(224) = &H1
    CRCTable(225) = &HC0
    CRCTable(226) = &H80
    CRCTable(227) = &H41
    CRCTable(228) = &H0
    CRCTable(229) = &HC1
    CRCTable(230) = &H81
    CRCTable(231) = &H40
    CRCTable(232) = &H0
    CRCTable(233) = &HC1
    CRCTable(234) = &H81
    CRCTable(235) = &H40
    CRCTable(236) = &H1
    CRCTable(237) = &HC0
    CRCTable(238) = &H80
    CRCTable(239) = &H41
    CRCTable(240) = &H0
    CRCTable(241) = &HC1
    CRCTable(242) = &H81
    CRCTable(243) = &H40
    CRCTable(244) = &H1
    CRCTable(245) = &HC0
    CRCTable(246) = &H80
    CRCTable(247) = &H41
    CRCTable(248) = &H1
    CRCTable(249) = &HC0
    CRCTable(250) = &H80
    CRCTable(251) = &H41
    CRCTable(252) = &H0
    CRCTable(253) = &HC1
    CRCTable(254) = &H81
    CRCTable(255) = &H40
    CRCTable(256) = &H0
    CRCTable(257) = &HC0
    CRCTable(258) = &HC1
    CRCTable(259) = &H1
    CRCTable(260) = &HC3
    CRCTable(261) = &H3
    CRCTable(262) = &H2
    CRCTable(263) = &HC2
    CRCTable(264) = &HC6
    CRCTable(265) = &H6
    CRCTable(266) = &H7
    CRCTable(267) = &HC7
    CRCTable(268) = &H5
    CRCTable(269) = &HC5
    CRCTable(270) = &HC4
    CRCTable(271) = &H4
    CRCTable(272) = &HCC
    CRCTable(273) = &HC
    CRCTable(274) = &HD
    CRCTable(275) = &HCD
    CRCTable(276) = &HF
    CRCTable(277) = &HCF
    CRCTable(278) = &HCE
    CRCTable(279) = &HE
    CRCTable(280) = &HA
    CRCTable(281) = &HCA
    CRCTable(282) = &HCB
    CRCTable(283) = &HB
    CRCTable(284) = &HC9
    CRCTable(285) = &H9
    CRCTable(286) = &H8
    CRCTable(287) = &HC8
    CRCTable(288) = &HD8
    CRCTable(289) = &H18
    CRCTable(290) = &H19
    CRCTable(291) = &HD9
    CRCTable(292) = &H1B
    CRCTable(293) = &HDB
    CRCTable(294) = &HDA
    CRCTable(295) = &H1A
    CRCTable(296) = &H1E
    CRCTable(297) = &HDE
    CRCTable(298) = &HDF
    CRCTable(299) = &H1F
    CRCTable(300) = &HDD
    CRCTable(301) = &H1D
    CRCTable(302) = &H1C
    CRCTable(303) = &HDC
    CRCTable(304) = &H14
    CRCTable(305) = &HD4
    CRCTable(306) = &HD5
    CRCTable(307) = &H15
    CRCTable(308) = &HD7
    CRCTable(309) = &H17
    CRCTable(310) = &H16
    CRCTable(311) = &HD6
    CRCTable(312) = &HD2
    CRCTable(313) = &H12
    CRCTable(314) = &H13
    CRCTable(315) = &HD3
    CRCTable(316) = &H11
    CRCTable(317) = &HD1
    CRCTable(318) = &HD0
    CRCTable(319) = &H10
    CRCTable(320) = &HF0
    CRCTable(321) = &H30
    CRCTable(322) = &H31
    CRCTable(323) = &HF1
    CRCTable(324) = &H33
    CRCTable(325) = &HF3
    CRCTable(326) = &HF2
    CRCTable(327) = &H32
    CRCTable(328) = &H36
    CRCTable(329) = &HF6
    CRCTable(330) = &HF7
    CRCTable(331) = &H37
    CRCTable(332) = &HF5
    CRCTable(333) = &H35
    CRCTable(334) = &H34
    CRCTable(335) = &HF4
    CRCTable(336) = &H3C
    CRCTable(337) = &HFC
    CRCTable(338) = &HFD
    CRCTable(339) = &H3D
    CRCTable(340) = &HFF
    CRCTable(341) = &H3F
    CRCTable(342) = &H3E
    CRCTable(343) = &HFE
    CRCTable(344) = &HFA
    CRCTable(345) = &H3A
    CRCTable(346) = &H3B
    CRCTable(347) = &HFB
    CRCTable(348) = &H39
    CRCTable(349) = &HF9
    CRCTable(350) = &HF8
    CRCTable(351) = &H38
    CRCTable(352) = &H28
    CRCTable(353) = &HE8
    CRCTable(354) = &HE9
    CRCTable(355) = &H29
    CRCTable(356) = &HEB
    CRCTable(357) = &H2B
    CRCTable(358) = &H2A
    CRCTable(359) = &HEA
    CRCTable(360) = &HEE
    CRCTable(361) = &H2E
    CRCTable(362) = &H2F
    CRCTable(363) = &HEF
    CRCTable(364) = &H2D
    CRCTable(365) = &HED
    CRCTable(366) = &HEC
    CRCTable(367) = &H2C
    CRCTable(368) = &HE4
    CRCTable(369) = &H24
    CRCTable(370) = &H25
    CRCTable(371) = &HE5
    CRCTable(372) = &H27
    CRCTable(373) = &HE7
    CRCTable(374) = &HE6
    CRCTable(375) = &H26
    CRCTable(376) = &H22
    CRCTable(377) = &HE2
    CRCTable(378) = &HE3
    CRCTable(379) = &H23
    CRCTable(380) = &HE1
    CRCTable(381) = &H21
    CRCTable(382) = &H20
    CRCTable(383) = &HE0
    CRCTable(384) = &HA0
    CRCTable(385) = &H60
    CRCTable(386) = &H61
    CRCTable(387) = &HA1
    CRCTable(388) = &H63
    CRCTable(389) = &HA3
    CRCTable(390) = &HA2
    CRCTable(391) = &H62
    CRCTable(392) = &H66
    CRCTable(393) = &HA6
    CRCTable(394) = &HA7
    CRCTable(395) = &H67
    CRCTable(396) = &HA5
    CRCTable(397) = &H65
    CRCTable(398) = &H64
    CRCTable(399) = &HA4
    CRCTable(400) = &H6C
    CRCTable(401) = &HAC
    CRCTable(402) = &HAD
    CRCTable(403) = &H6D
    CRCTable(404) = &HAF
    CRCTable(405) = &H6F
    CRCTable(406) = &H6E
    CRCTable(407) = &HAE
    CRCTable(408) = &HAA
    CRCTable(409) = &H6A
    CRCTable(410) = &H6B
    CRCTable(411) = &HAB
    CRCTable(412) = &H69
    CRCTable(413) = &HA9
    CRCTable(414) = &HA8
    CRCTable(415) = &H68
    CRCTable(416) = &H78
    CRCTable(417) = &HB8
    CRCTable(418) = &HB9
    CRCTable(419) = &H79
    CRCTable(420) = &HBB
    CRCTable(421) = &H7B
    CRCTable(422) = &H7A
    CRCTable(423) = &HBA
    CRCTable(424) = &HBE
    CRCTable(425) = &H7E
    CRCTable(426) = &H7F
    CRCTable(427) = &HBF
    CRCTable(428) = &H7D
    CRCTable(429) = &HBD
    CRCTable(430) = &HBC
    CRCTable(431) = &H7C
    CRCTable(432) = &HB4
    CRCTable(433) = &H74
    CRCTable(434) = &H75
    CRCTable(435) = &HB5
    CRCTable(436) = &H77
    CRCTable(437) = &HB7
    CRCTable(438) = &HB6
    CRCTable(439) = &H76
    CRCTable(440) = &H72
    CRCTable(441) = &HB2
    CRCTable(442) = &HB3
    CRCTable(443) = &H73
    CRCTable(444) = &HB1
    CRCTable(445) = &H71
    CRCTable(446) = &H70
    CRCTable(447) = &HB0
    CRCTable(448) = &H50
    CRCTable(449) = &H90
    CRCTable(450) = &H91
    CRCTable(451) = &H51
    CRCTable(452) = &H93
    CRCTable(453) = &H53
    CRCTable(454) = &H52
    CRCTable(455) = &H92
    CRCTable(456) = &H96
    CRCTable(457) = &H56
    CRCTable(458) = &H57
    CRCTable(459) = &H97
    CRCTable(460) = &H55
    CRCTable(461) = &H95
    CRCTable(462) = &H94
    CRCTable(463) = &H54
    CRCTable(464) = &H9C
    CRCTable(465) = &H5C
    CRCTable(466) = &H5D
    CRCTable(467) = &H9D
    CRCTable(468) = &H5F
    CRCTable(469) = &H9F
    CRCTable(470) = &H9E
    CRCTable(471) = &H5E
    CRCTable(472) = &H5A
    CRCTable(473) = &H9A
    CRCTable(474) = &H9B
    CRCTable(475) = &H5B
    CRCTable(476) = &H99
    CRCTable(477) = &H59
    CRCTable(478) = &H58
    CRCTable(479) = &H98
    CRCTable(480) = &H88
    CRCTable(481) = &H48
    CRCTable(482) = &H49
    CRCTable(483) = &H89
    CRCTable(484) = &H4B
    CRCTable(485) = &H8B
    CRCTable(486) = &H8A
    CRCTable(487) = &H4A
    CRCTable(488) = &H4E
    CRCTable(489) = &H8E
    CRCTable(490) = &H8F
    CRCTable(491) = &H4F
    CRCTable(492) = &H8D
    CRCTable(493) = &H4D
    CRCTable(494) = &H4C
    CRCTable(495) = &H8C
    CRCTable(496) = &H44
    CRCTable(497) = &H84
    CRCTable(498) = &H85
    CRCTable(499) = &H45
    CRCTable(500) = &H87
    CRCTable(501) = &H47
    CRCTable(502) = &H46
    CRCTable(503) = &H86
    CRCTable(504) = &H82
    CRCTable(505) = &H42
    CRCTable(506) = &H43
    CRCTable(507) = &H83
    CRCTable(508) = &H41
    CRCTable(509) = &H81
    CRCTable(510) = &H80
    CRCTable(511) = &H40

    cbbAutoAckTime.Enabled = False
    
    'InputP1.OutputBoth = "BOTH"
       
    For i = 0 To 239 Step 1
        cbbAutoAckTime.List(i) = i + 1
    Next i
    
    'For i = 0 To 10 Step 1
    'cbbFaultDelayTime.List(i) = (i + 0.5)
        'cbbFaultDelayTime.ListIndex = 0
    'Next i
    'cbbFaultDelayTime
    
    cbbFaultDelayTime.List(0) = 0
    cbbFaultDelayTime.List(1) = 0.2
    cbbFaultDelayTime.List(2) = 0.4
    cbbFaultDelayTime.List(3) = 0.6
    cbbFaultDelayTime.List(4) = 0.8
    cbbFaultDelayTime.List(5) = 1
    cbbFaultDelayTime.List(6) = 1.2
    cbbFaultDelayTime.List(7) = 1.4
    cbbFaultDelayTime.List(8) = 1.6
    cbbFaultDelayTime.List(9) = 1.8
    cbbFaultDelayTime.List(10) = 2
    cbbFaultDelayTime.List(11) = 2.2
    cbbFaultDelayTime.List(12) = 2.4
      
    cbbFaultDelayTime.ListIndex = 0
    '*************************************
    ComboT1.List(0) = 0
    ComboT1.List(1) = 0.2
    ComboT1.List(2) = 0.4
    ComboT1.List(3) = 0.6
    ComboT1.List(4) = 0.8
    ComboT1.List(5) = 1
    ComboT1.List(6) = 1.2
    ComboT1.List(7) = 1.4
    ComboT1.List(8) = 1.6
    ComboT1.List(9) = 1.8
    ComboT1.List(10) = 2
    ComboT1.List(11) = 2.2
    ComboT1.List(12) = 2.4
      
    ComboT1.ListIndex = 0
    '*************************************
    ComboT2.List(0) = 0
    ComboT2.List(1) = 0.2
    ComboT2.List(2) = 0.4
    ComboT2.List(3) = 0.6
    ComboT2.List(4) = 0.8
    ComboT2.List(5) = 1
    ComboT2.List(6) = 1.2
    ComboT2.List(7) = 1.4
    ComboT2.List(8) = 1.6
    ComboT2.List(9) = 1.8
    ComboT2.List(10) = 2
    ComboT2.List(11) = 2.2
    ComboT2.List(12) = 2.4
      
    ComboT2.ListIndex = 0
    '*************************************
    ComboT3.List(0) = 0
    ComboT3.List(1) = 0.2
    ComboT3.List(2) = 0.4
    ComboT3.List(3) = 0.6
    ComboT3.List(4) = 0.8
    ComboT3.List(5) = 1
    ComboT3.List(6) = 1.2
    ComboT3.List(7) = 1.4
    ComboT3.List(8) = 1.6
    ComboT3.List(9) = 1.8
    ComboT3.List(10) = 2
    ComboT3.List(11) = 2.2
    ComboT3.List(12) = 2.4
      
    ComboT3.ListIndex = 0
    '*************************************
    ComboT4.List(0) = 0
    ComboT4.List(1) = 0.2
    ComboT4.List(2) = 0.4
    ComboT4.List(3) = 0.6
    ComboT4.List(4) = 0.8
    ComboT4.List(5) = 1
    ComboT4.List(6) = 1.2
    ComboT4.List(7) = 1.4
    ComboT4.List(8) = 1.6
    ComboT4.List(9) = 1.8
    ComboT4.List(10) = 2
    ComboT4.List(11) = 2.2
    ComboT4.List(12) = 2.4
      
    ComboT4.ListIndex = 0
    '*************************************
    ComboT5.List(0) = 0
    ComboT5.List(1) = 0.2
    ComboT5.List(2) = 0.4
    ComboT5.List(3) = 0.6
    ComboT5.List(4) = 0.8
    ComboT5.List(5) = 1
    ComboT5.List(6) = 1.2
    ComboT5.List(7) = 1.4
    ComboT5.List(8) = 1.6
    ComboT5.List(9) = 1.8
    ComboT5.List(10) = 2
    ComboT5.List(11) = 2.2
    ComboT5.List(12) = 2.4
      
    ComboT5.ListIndex = 0
    '*************************************
    ComboT6.List(0) = 0
    ComboT6.List(1) = 0.2
    ComboT6.List(2) = 0.4
    ComboT6.List(3) = 0.6
    ComboT6.List(4) = 0.8
    ComboT6.List(5) = 1
    ComboT6.List(6) = 1.2
    ComboT6.List(7) = 1.4
    ComboT6.List(8) = 1.6
    ComboT6.List(9) = 1.8
    ComboT6.List(10) = 2
    ComboT6.List(11) = 2.2
    ComboT6.List(12) = 2.4
      
    ComboT6.ListIndex = 0
    '*************************************
    ComboT7.List(0) = 0
    ComboT7.List(1) = 0.2
    ComboT7.List(2) = 0.4
    ComboT7.List(3) = 0.6
    ComboT7.List(4) = 0.8
    ComboT7.List(5) = 1
    ComboT7.List(6) = 1.2
    ComboT7.List(7) = 1.4
    ComboT7.List(8) = 1.6
    ComboT7.List(9) = 1.8
    ComboT7.List(10) = 2
    ComboT7.List(11) = 2.2
    ComboT7.List(12) = 2.4
      
    ComboT7.ListIndex = 0
    '*************************************
    ComboT8.List(0) = 0
    ComboT8.List(1) = 0.2
    ComboT8.List(2) = 0.4
    ComboT8.List(3) = 0.6
    ComboT8.List(4) = 0.8
    ComboT8.List(5) = 1
    ComboT8.List(6) = 1.2
    ComboT8.List(7) = 1.4
    ComboT8.List(8) = 1.6
    ComboT8.List(9) = 1.8
    ComboT8.List(10) = 2
    ComboT8.List(11) = 2.2
    ComboT8.List(12) = 2.4
      
    ComboT8.ListIndex = 0
    '*************************************
    ComboT9.List(0) = 0
    ComboT9.List(1) = 0.2
    ComboT9.List(2) = 0.4
    ComboT9.List(3) = 0.6
    ComboT9.List(4) = 0.8
    ComboT9.List(5) = 1
    ComboT9.List(6) = 1.2
    ComboT9.List(7) = 1.4
    ComboT9.List(8) = 1.6
    ComboT9.List(9) = 1.8
    ComboT9.List(10) = 2
    ComboT9.List(11) = 2.2
    ComboT9.List(12) = 2.4
      
    ComboT9.ListIndex = 0
    '*************************************
    ComboT10.List(0) = 0
    ComboT10.List(1) = 0.2
    ComboT10.List(2) = 0.4
    ComboT10.List(3) = 0.6
    ComboT10.List(4) = 0.8
    ComboT10.List(5) = 1
    ComboT10.List(6) = 1.2
    ComboT10.List(7) = 1.4
    ComboT10.List(8) = 1.6
    ComboT10.List(9) = 1.8
    ComboT10.List(10) = 2
    ComboT10.List(11) = 2.2
    ComboT10.List(12) = 2.4
      
    ComboT10.ListIndex = 0
    '*************************************
    
    '*************************************
    ComboT11.List(0) = 0
    ComboT11.List(1) = 0.2
    ComboT11.List(2) = 0.4
    ComboT11.List(3) = 0.6
    ComboT11.List(4) = 0.8
    ComboT11.List(5) = 1
    ComboT11.List(6) = 1.2
    ComboT11.List(7) = 1.4
    ComboT11.List(8) = 1.6
    ComboT11.List(9) = 1.8
    ComboT11.List(10) = 2
    ComboT11.List(11) = 2.2
    ComboT11.List(12) = 2.4
      
    ComboT11.ListIndex = 0
    '*************************************
    ComboT12.List(0) = 0
    ComboT12.List(1) = 0.2
    ComboT12.List(2) = 0.4
    ComboT12.List(3) = 0.6
    ComboT12.List(4) = 0.8
    ComboT12.List(5) = 1
    ComboT12.List(6) = 1.2
    ComboT12.List(7) = 1.4
    ComboT12.List(8) = 1.6
    ComboT12.List(9) = 1.8
    ComboT12.List(10) = 2
    ComboT12.List(11) = 2.2
    ComboT12.List(12) = 2.4
      
    ComboT12.ListIndex = 0
    '*************************************
    ComboT13.List(0) = 0
    ComboT13.List(1) = 0.2
    ComboT13.List(2) = 0.4
    ComboT13.List(3) = 0.6
    ComboT13.List(4) = 0.8
    ComboT13.List(5) = 1
    ComboT13.List(6) = 1.2
    ComboT13.List(7) = 1.4
    ComboT13.List(8) = 1.6
    ComboT13.List(9) = 1.8
    ComboT13.List(10) = 2
    ComboT13.List(11) = 2.2
    ComboT13.List(12) = 2.4
      
    ComboT13.ListIndex = 0
    '*************************************
    ComboT14.List(0) = 0
    ComboT14.List(1) = 0.2
    ComboT14.List(2) = 0.4
    ComboT14.List(3) = 0.6
    ComboT14.List(4) = 0.8
    ComboT14.List(5) = 1
    ComboT14.List(6) = 1.2
    ComboT14.List(7) = 1.4
    ComboT14.List(8) = 1.6
    ComboT14.List(9) = 1.8
    ComboT14.List(10) = 2
    ComboT14.List(11) = 2.2
    ComboT14.List(12) = 2.4
      
    ComboT14.ListIndex = 0
    '*************************************
    ComboT15.List(0) = 0
    ComboT15.List(1) = 0.2
    ComboT15.List(2) = 0.4
    ComboT15.List(3) = 0.6
    ComboT15.List(4) = 0.8
    ComboT15.List(5) = 1
    ComboT15.List(6) = 1.2
    ComboT15.List(7) = 1.4
    ComboT15.List(8) = 1.6
    ComboT15.List(9) = 1.8
    ComboT15.List(10) = 2
    ComboT15.List(11) = 2.2
    ComboT15.List(12) = 2.4
      
    ComboT15.ListIndex = 0
    '*************************************
    ComboT16.List(0) = 0
    ComboT16.List(1) = 0.2
    ComboT16.List(2) = 0.4
    ComboT16.List(3) = 0.6
    ComboT16.List(4) = 0.8
    ComboT16.List(5) = 1
    ComboT16.List(6) = 1.2
    ComboT16.List(7) = 1.4
    ComboT16.List(8) = 1.6
    ComboT16.List(9) = 1.8
    ComboT16.List(10) = 2
    ComboT16.List(11) = 2.2
    ComboT16.List(12) = 2.4
      
    ComboT16.ListIndex = 0
    '*************************************
    ComboT17.List(0) = 0
    ComboT17.List(1) = 0.2
    ComboT17.List(2) = 0.4
    ComboT17.List(3) = 0.6
    ComboT17.List(4) = 0.8
    ComboT17.List(5) = 1
    ComboT17.List(6) = 1.2
    ComboT17.List(7) = 1.4
    ComboT17.List(8) = 1.6
    ComboT17.List(9) = 1.8
    ComboT17.List(10) = 2
    ComboT17.List(11) = 2.2
    ComboT17.List(12) = 2.4
      
    ComboT17.ListIndex = 0
    '*************************************
    ComboT18.List(0) = 0
    ComboT18.List(1) = 0.2
    ComboT18.List(2) = 0.4
    ComboT18.List(3) = 0.6
    ComboT18.List(4) = 0.8
    ComboT18.List(5) = 1
    ComboT18.List(6) = 1.2
    ComboT18.List(7) = 1.4
    ComboT18.List(8) = 1.6
    ComboT18.List(9) = 1.8
    ComboT18.List(10) = 2
    ComboT18.List(11) = 2.2
    ComboT18.List(12) = 2.4
      
    ComboT18.ListIndex = 0
    '*************************************
    ComboT19.List(0) = 0
    ComboT19.List(1) = 0.2
    ComboT19.List(2) = 0.4
    ComboT19.List(3) = 0.6
    ComboT19.List(4) = 0.8
    ComboT19.List(5) = 1
    ComboT19.List(6) = 1.2
    ComboT19.List(7) = 1.4
    ComboT19.List(8) = 1.6
    ComboT19.List(9) = 1.8
    ComboT19.List(10) = 2
    ComboT19.List(11) = 2.2
    ComboT19.List(12) = 2.4
      
    ComboT19.ListIndex = 0
    '*************************************
    ComboT20.List(0) = 0
    ComboT20.List(1) = 0.2
    ComboT20.List(2) = 0.4
    ComboT20.List(3) = 0.6
    ComboT20.List(4) = 0.8
    ComboT20.List(5) = 1
    ComboT20.List(6) = 1.2
    ComboT20.List(7) = 1.4
    ComboT20.List(8) = 1.6
    ComboT20.List(9) = 1.8
    ComboT20.List(10) = 2
    ComboT20.List(11) = 2.2
    ComboT20.List(12) = 2.4
      
    ComboT20.ListIndex = 0
    '*************************************
    
    OptNoInput8.Value = True
        
    SSTab1.TabEnabled(0) = True
    'SSTab1.Caption = "INPUT 1-20"
    SSTab1.Caption = "INPUT 1-8"
    
    InputP1.Visible = True
    InputP2.Visible = True
    InputP3.Visible = True
    InputP4.Visible = True
    InputP5.Visible = True
    InputP6.Visible = True
    InputP7.Visible = True
    InputP8.Visible = True
    InputP9.Visible = False
    InputP10.Visible = False
    InputP11.Visible = False
    InputP12.Visible = False
    InputP13.Visible = False
    InputP14.Visible = False
    InputP15.Visible = False
    InputP16.Visible = False
    InputP17.Visible = False
    InputP18.Visible = False
    InputP19.Visible = False
    InputP20.Visible = False
    
    FrameRGA1.Visible = True
    FrameRGA2.Visible = True
    FrameRGA3.Visible = True
    FrameRGA4.Visible = True
    FrameRGA5.Visible = True
    FrameRGA6.Visible = True
    FrameRGA7.Visible = True
    FrameRGA8.Visible = True
    FrameRGA9.Visible = False
    FrameRGA10.Visible = False
    FrameRGA11.Visible = False
    FrameRGA12.Visible = False
    FrameRGA13.Visible = False
    FrameRGA14.Visible = False
    FrameRGA15.Visible = False
    FrameRGA16.Visible = False
    FrameRGA17.Visible = False
    FrameRGA18.Visible = False
    FrameRGA19.Visible = False
    FrameRGA20.Visible = False
    
    FrameT1.Visible = True
    FrameT2.Visible = True
    FrameT3.Visible = True
    FrameT4.Visible = True
    FrameT5.Visible = True
    FrameT6.Visible = True
    FrameT7.Visible = True
    FrameT8.Visible = True
    FrameT9.Visible = False
    FrameT10.Visible = False
    FrameT11.Visible = False
    FrameT12.Visible = False
    FrameT13.Visible = False
    FrameT14.Visible = False
    FrameT15.Visible = False
    FrameT16.Visible = False
    FrameT17.Visible = False
    FrameT18.Visible = False
    FrameT19.Visible = False
    FrameT20.Visible = False
    
    InputP1.OutputBoth = "BOTH"
    InputP2.OutputBoth = "BOTH"
    InputP3.OutputBoth = "BOTH"
    InputP4.OutputBoth = "BOTH"
    InputP5.OutputBoth = "BOTH"
    InputP6.OutputBoth = "BOTH"
    InputP7.OutputBoth = "BOTH"
    InputP8.OutputBoth = "BOTH"
    InputP9.OutputBoth = "BOTH"
    InputP10.OutputBoth = "BOTH"
    InputP11.OutputBoth = "BOTH"
    InputP12.OutputBoth = "BOTH"
    InputP13.OutputBoth = "BOTH"
    InputP14.OutputBoth = "BOTH"
    InputP15.OutputBoth = "BOTH"
    InputP16.OutputBoth = "BOTH"
    InputP17.OutputBoth = "BOTH"
    InputP18.OutputBoth = "BOTH"
    InputP19.OutputBoth = "BOTH"
    InputP20.OutputBoth = "BOTH"
    
    
    CommState = "IDLE"
    Timeout = 0
    
    'Me.au = True
    
End Sub

Private Sub CRC_16(ByVal Data As String, Length As Integer)
Dim i As Integer
Dim Index As Byte

CRC_Low = &HFF
CRC_High = &HFF

For i = 1 To Length
Index = CRC_High Xor Asc(Mid(Data, i, 1))
CRC_High = CRC_Low Xor CRCTable(Index)
CRC_Low = CRCTable(Index + 256)
Next i

End Sub

Private Sub mnuComSetting_Click()
    frmSetup.Show
    'frmSetup2_2.Show
End Sub

Private Sub MSComm1_OnComm()
'Dim Buffer As Variant
Dim i As Integer


    Select Case MSComm1.CommEvent
        Case comEvReceive
        'Buffer = ""
        Sleep (500)                         'Wait for receive completed data lenght
            Buffer = Buffer + MSComm1.Input
            
            'Text2.Text = ""
            'Text2.Text = Buffer
    End Select
    '/////////////////////////////////////////////////////
    If CommState = "DOWNLOAD" And Len(Buffer) = 4 Then
    
        For i = 0 To 3
            CommBuff(i) = Asc(Mid(Buffer, i + 1, 1))
        Next i
        
        If CommBuff(0) = TextAddrNow.Text And (CommBuff(1) = 33 Or CommBuff(1) = 34) Then
            CommState = "IDLE"
            Timeout = 0
            MsgBox "DOWNLOAD COMPLETED", vbInformation, "ESPAN-04"
        End If
    End If
    '////////////////////////////////////////////////////////////////
    
    If CommState = "UPLOAD" And (Len(Buffer) = 54) Then
        'Command1.Caption = "Search"
        'Search_Addr = False
        'Timer1.Enabled = False
        'Timer2.Enabled = False
        'Poll_Counter = 0
        
        addr = Asc(Mid(Buffer, 1, 1)) 'keep  addr now
        
        For i = 4 To 53
            DatBuff(i - 3) = Asc(Mid(Buffer, i, 1))
            'DecodeValue
        Next i
        DecodeValue

        'Shape1.BackColor = vbGreen
        Buffer = ""
        
        Timer1.Enabled = False
        CommState = "IDLE"
        Timeout = 0
        MsgBox "UPLOAD COMPLETED", vbInformation, "ESPAN-04"
        
    Else
    Buffer = ""
    End If
        
   
End Sub

Private Sub mnuUpload_Click()
On Error GoTo Error
Dim Data As String
    'Shape1.BackColor = vbRed
    If Not MSComm1.PortOpen Then
        MSComm1.CommPort = GetSetting("ESPAN-01", "Setting", "Comm Port", "1")
        'MSComm1.CommPort = 1
        MSComm1.Settings = "9600,N,8,1"
        MSComm1.PortOpen = True
        MSComm1.RThreshold = 1
    End If

    Data = ""
    'Data = Chr(Val(TextAddrNow.Text)) + Chr(32)
    Data = Chr(170) + Chr(32) '170(AA)is any Address 32 is (20) Read Setting
    CRC_16 Data, 2
    Data = Data + Chr(CRC_High) + Chr(CRC_Low)
    MSComm1.InputLen = 0
    MSComm1.Output = Data
    Buffer = ""
    Do While MSComm1.OutBufferCount > 0
    Loop
    'MSComm1.Output = "243"
    'Sleep (100)
    
    CommState = "UPLOAD"
    Timeout = 0
    Timer1.Enabled = True
    
    GoTo AAA
Error:
    MsgBox ("Invalid Port Number.")
    frmSetup.Show
AAA:
End Sub

Private Sub mnuDownload_Click()

Dim SumX As Byte
Dim SumXtemp As Double
Dim SumXInputType1_8 As Byte
Dim SumXInputType9_16 As Byte
Dim SumXFaultType1_8 As Byte
Dim SumXFaultType9_16 As Byte
Dim SumXOutputType1_8 As Byte
Dim SumXOutputType9_16 As Byte
Dim SumXFlashRate As Byte
Dim SumXIndicator1_8 As Byte
Dim SumXIndicator9_16 As Byte

Dim TxBuff() As Byte
Dim i As Integer
    
Dim Data As String
Dim mymassage As String
Dim Error As Boolean

    'Shape1.BackColor = vbRed
    Error = CalculateValue      'check valid data
    If Error Then GoTo AAA
    
    If CommState = "IDLE" Then
        On Error GoTo Error
        If Not MSComm1.PortOpen Then
            MSComm1.CommPort = GetSetting("ESPAN-01", "Setting", "Comm Port", "1")
            'MSComm1.CommPort = 1
            MSComm1.Settings = "9600,N,8,1"
            MSComm1.PortOpen = True
            MSComm1.RThreshold = 1
        End If
    
        Data = ""
        mymassage = ""
        'Data = Chr(Val(Text1.Text)) + Chr(33) + Chr(8) + Chr(InputType1_8) + Chr(InputType9_16) + Chr(FaultType1_8) + Chr(FaultType9_16) + Chr(OutputType1_8) + Chr(OutputType9_16) + Chr(FlashRate) + Chr(Val(Text3.Text))
    
        'Data = Chr(Val(TextAddrNow.Text)) + Chr(33) + Chr(59) '59 are data range(62) -3
        
        mymassage = Chr(InputType1_8) + Chr(InputType9_16) + Chr(InputType17_24)
        mymassage = mymassage + Chr(FaultType1_8) + Chr(FaultType9_16) + Chr(FaultType17_24)
        mymassage = mymassage + Chr(OutputType1_8) + Chr(OutputType9_16) + Chr(OutputType17_24)
        mymassage = mymassage + Chr(OutputBoth1_8) + Chr(OutputBoth9_16) + Chr(OutputBoth17_24)
        mymassage = mymassage + Chr(Fault_Indicator1_8) + Chr(Fault_Indicator9_16) + Chr(Fault_Indicator17_24)
        mymassage = mymassage + Chr(Red1_8) + Chr(Red9_10) + Chr(Red11_18) + Chr(Red19_20) + Chr(Green1_8) + Chr(Green9_10) + Chr(Green11_18) + Chr(Green19_20)
        mymassage = mymassage + Chr(AutoAckStatus) + Chr(AutoAckTime) + Chr(FlashRate) + Chr(NoOfInput) + Chr(FaultDelayTime) + Chr(Val(TextAddrNow.Text))
        mymassage = mymassage + Chr(DelayTime1) + Chr(DelayTime2) + Chr(DelayTime3) + Chr(DelayTime4) + Chr(DelayTime5) + Chr(DelayTime6) + Chr(DelayTime7) + Chr(DelayTime8) + Chr(DelayTime9) + Chr(DelayTime10)
        mymassage = mymassage + Chr(DelayTime11) + Chr(DelayTime12) + Chr(DelayTime13) + Chr(DelayTime14) + Chr(DelayTime15) + Chr(DelayTime16) + Chr(DelayTime17) + Chr(DelayTime18) + Chr(DelayTime19) + Chr(DelayTime20) ' total49
        
        mymassage = mymassage + txtphone.Text + Chr(13) '10 + 1 charrector
        

        'CRC_16 Data, 62
        Dim massagelen As Integer
        massagelen = Len(mymassage)
                                            '0x21 WRITE SETTING
        Data = Chr(Val(TextAddrNow.Text)) + Chr(33) + Chr(massagelen) '59 are data range(62) -3
            
        Data = Data + mymassage
        
        massagelen = Len(Data)
        
        CRC_16 Data, massagelen

        Data = Data + Chr(CRC_High) + Chr(CRC_Low)
        MSComm1.InputLen = 0
        MSComm1.Output = Data
        Buffer = ""
        Do While MSComm1.OutBufferCount > 0
        Loop
        'Bit_Setting = True
        'Timer2.Enabled = True
        CommState = "DOWNLOAD"
        Timer1.Enabled = True
    End If
    
    GoTo AAA
Error:
    MsgBox ("Invalid Port Number.")
    frmSetup.Show
AAA:



End Sub


Private Function CalculateValue() As Boolean
    InputType1_8 = 0
    InputType9_16 = 0
    InputType17_24 = 0
    
    FaultType1_8 = 0
    FaultType9_16 = 0
    FaultType17_24 = 0
    
    
    OutputType1_8 = 0
    OutputType9_16 = 0
    OutputType17_24 = 0
    
    
    OutputBoth1_8 = 0
    OutputBoth9_16 = 0
    OutputBoth17_24 = 0
    
    
    Fault_Indicator1_8 = 0
    Fault_Indicator9_16 = 0
    Fault_Indicator17_24 = 0
    
    Red1_8 = 0
    Red9_10 = 0
    Red11_18 = 0
    Red19_20 = 0

    Green1_8 = 0
    Green9_10 = 0
    Green11_18 = 0
    Green19_20 = 0


    '//////////////////// LED Colour Config /////////////
    '///////////////////////////////////////////////////
    
    If R1.Value = True Then
        Red1_8 = Red1_8 + 1
    ElseIf G1.Value = True Then
        Green1_8 = Green1_8 + 1
    ElseIf A1.Value = True Then
            Red1_8 = Red1_8 + 1
            Green1_8 = Green1_8 + 1
        'End If
    End If
    '///////////////////
    If R2.Value = True Then
        Red1_8 = Red1_8 + 2
    ElseIf G2.Value = True Then
        Green1_8 = Green1_8 + 2
    ElseIf A2.Value = True Then
            Red1_8 = Red1_8 + 2
            Green1_8 = Green1_8 + 2
        'End If
    End If
    '///////////////////
    If R3.Value = True Then
        Red1_8 = Red1_8 + 4
    ElseIf G3.Value = True Then
        Green1_8 = Green1_8 + 4
    ElseIf A3.Value = True Then
            Red1_8 = Red1_8 + 4
            Green1_8 = Green1_8 + 4
        'End If
    End If
    '///////////////////
    If R4.Value = True Then
        Red1_8 = Red1_8 + 8
    ElseIf G4.Value = True Then
        Green1_8 = Green1_8 + 8
    ElseIf A4.Value = True Then
            Red1_8 = Red1_8 + 8
            Green1_8 = Green1_8 + 8
        'End If
    End If
    '///////////////////
    If R5.Value = True Then
        Red1_8 = Red1_8 + 16
    ElseIf G5.Value = True Then
        Green1_8 = Green1_8 + 16
    ElseIf A5.Value = True Then
            Red1_8 = Red1_8 + 16
            Green1_8 = Green1_8 + 16
        'End If
    End If
    '///////////////////
    If R6.Value = True Then
        Red1_8 = Red1_8 + 32
    ElseIf G6.Value = True Then
        Green1_8 = Green1_8 + 32
    ElseIf A6.Value = True Then
            Red1_8 = Red1_8 + 32
            Green1_8 = Green1_8 + 32
        'End If
    End If
    '///////////////////
    If R7.Value = True Then
        Red1_8 = Red1_8 + 64
    ElseIf G7.Value = True Then
        Green1_8 = Green1_8 + 64
    ElseIf A7.Value = True Then
            Red1_8 = Red1_8 + 64
            Green1_8 = Green1_8 + 64
        'End If
    End If
    '///////////////////
    If R8.Value = True Then
        Red1_8 = Red1_8 + 128
    ElseIf G8.Value = True Then
        Green1_8 = Green1_8 + 128
    ElseIf A8.Value = True Then
            Red1_8 = Red1_8 + 128
            Green1_8 = Green1_8 + 128
        'End If
    End If
    '///////////////////
    If R9.Value = True Then
        Red9_10 = Red9_10 + 1
    ElseIf G9.Value = True Then
        Green9_10 = Green9_10 + 1
    Else
        If A9.Value = True Then
            Red9_10 = Red9_10 + 1
            Green9_10 = Green9_10 + 1
        End If
    End If
    '///////////////////
    If R10.Value = True Then
        Red9_10 = Red9_10 + 2
    ElseIf G10.Value = True Then
        Green9_10 = Green9_10 + 2
    Else
        If A10.Value = True Then
            Red9_10 = Red9_10 + 2
            Green9_10 = Green9_10 + 2
        End If
    End If
    '///////////////////
    If R11.Value = True Then
        Red11_18 = Red11_18 + 1
    ElseIf G11.Value = True Then
        Green11_18 = Green11_18 + 1
    Else
        If A11.Value = True Then
            Red11_18 = Red11_18 + 1
            Green11_18 = Green11_18 + 1
        End If
    End If
    '///////////////////
    If R12.Value = True Then
        Red11_18 = Red11_18 + 2
    ElseIf G12.Value = True Then
        Green11_18 = Green11_18 + 2
    Else
        If A12.Value = True Then
            Red11_18 = Red11_18 + 2
            Green11_18 = Green11_18 + 2
        End If
    End If
    '///////////////////
    If R13.Value = True Then
        Red11_18 = Red11_18 + 4
    ElseIf G13.Value = True Then
        Green11_18 = Green11_18 + 4
    Else
        If A13.Value = True Then
            Red11_18 = Red11_18 + 4
            Green11_18 = Green11_18 + 4
        End If
    End If
    '///////////////////
    If R14.Value = True Then
        Red11_18 = Red11_18 + 8
    ElseIf G14.Value = True Then
        Green11_18 = Green11_18 + 8
    Else
        If A14.Value = True Then
            Red11_18 = Red11_18 + 8
            Green11_18 = Green11_18 + 8
        End If
    End If
    '///////////////////
    If R15.Value = True Then
        Red11_18 = Red11_18 + 16
    ElseIf G15.Value = True Then
        Green11_18 = Green11_18 + 16
    Else
        If A15.Value = True Then
            Red11_18 = Red11_18 + 16
            Green11_18 = Green11_18 + 16
        End If
    End If
    '///////////////////
    If R16.Value = True Then
        Red11_18 = Red11_18 + 32
    ElseIf G16.Value = True Then
        Green11_18 = Green11_18 + 32
    Else
        If A16.Value = True Then
            Red11_18 = Red11_18 + 32
            Green11_18 = Green11_18 + 32
        End If
    End If
    '///////////////////
    If R17.Value = True Then
        Red11_18 = Red11_18 + 64
    ElseIf G17.Value = True Then
        Green11_18 = Green11_18 + 64
    Else
        If A17.Value = True Then
            Red11_18 = Red11_18 + 64
            Green11_18 = Green11_18 + 64
        End If
    End If
    '///////////////////
    If R18.Value = True Then
        Red11_18 = Red11_18 + 128
    ElseIf G18.Value = True Then
        Green11_18 = Green11_18 + 128
    Else
        If A18.Value = True Then
            Red11_18 = Red11_18 + 128
            Green11_18 = Green11_18 + 128
        End If
    End If
    '///////////////////
    If R19.Value = True Then
        Red19_20 = Red19_20 + 1
    ElseIf G19.Value = True Then
        Green19_20 = Green19_20 + 1
    Else
        If A19.Value = True Then
            Red19_20 = Red19_20 + 1
            Green19_20 = Green19_20 + 1
        End If
    End If
    '///////////////////
    If R20.Value = True Then
        Red19_20 = Red19_20 + 2
    ElseIf G20.Value = True Then
        Green19_20 = Green19_20 + 2
    Else
        If A20.Value = True Then
            Red19_20 = Red19_20 + 2
            Green19_20 = Green19_20 + 2
        End If
    End If
    '///////////////////
    
    
       
    '////////////// Data byte 1-8 ////////////////////////////////
    
    If InputP1.InputType = "NO" Then InputType1_8 = InputType1_8 + 1    ' Input Type 1-8
    If InputP2.InputType = "NO" Then InputType1_8 = InputType1_8 + 2
    If InputP3.InputType = "NO" Then InputType1_8 = InputType1_8 + 4
    If InputP4.InputType = "NO" Then InputType1_8 = InputType1_8 + 8
    If InputP5.InputType = "NO" Then InputType1_8 = InputType1_8 + 16
    If InputP6.InputType = "NO" Then InputType1_8 = InputType1_8 + 32
    If InputP7.InputType = "NO" Then InputType1_8 = InputType1_8 + 64
    If InputP8.InputType = "NO" Then InputType1_8 = InputType1_8 + 128
    
    If InputP9.InputType = "NO" Then InputType9_16 = InputType9_16 + 1  ' Input Type 9-16
    If InputP10.InputType = "NO" Then InputType9_16 = InputType9_16 + 2
    If InputP11.InputType = "NO" Then InputType9_16 = InputType9_16 + 4
    If InputP12.InputType = "NO" Then InputType9_16 = InputType9_16 + 8
    If InputP13.InputType = "NO" Then InputType9_16 = InputType9_16 + 16
    If InputP14.InputType = "NO" Then InputType9_16 = InputType9_16 + 32
    If InputP15.InputType = "NO" Then InputType9_16 = InputType9_16 + 64
    If InputP16.InputType = "NO" Then InputType9_16 = InputType9_16 + 128
    
    If InputP17.InputType = "NO" Then InputType17_24 = InputType17_24 + 1  ' Input Type 17-24
    If InputP18.InputType = "NO" Then InputType17_24 = InputType17_24 + 2
    If InputP19.InputType = "NO" Then InputType17_24 = InputType17_24 + 4
    If InputP20.InputType = "NO" Then InputType17_24 = InputType17_24 + 8
        
    ''////////////////////// Data byte 9-16 ///////////////////////////////////
    
    If InputP1.FaultType = "MANUAL" Then FaultType1_8 = FaultType1_8 + 1    ' Fault Type 1-8
    If InputP2.FaultType = "MANUAL" Then FaultType1_8 = FaultType1_8 + 2
    If InputP3.FaultType = "MANUAL" Then FaultType1_8 = FaultType1_8 + 4
    If InputP4.FaultType = "MANUAL" Then FaultType1_8 = FaultType1_8 + 8
    If InputP5.FaultType = "MANUAL" Then FaultType1_8 = FaultType1_8 + 16
    If InputP6.FaultType = "MANUAL" Then FaultType1_8 = FaultType1_8 + 32
    If InputP7.FaultType = "MANUAL" Then FaultType1_8 = FaultType1_8 + 64
    If InputP8.FaultType = "MANUAL" Then FaultType1_8 = FaultType1_8 + 128
    
    If InputP9.FaultType = "MANUAL" Then FaultType9_16 = FaultType9_16 + 1  ' Fault Type 9-16
    If InputP10.FaultType = "MANUAL" Then FaultType9_16 = FaultType9_16 + 2
    If InputP11.FaultType = "MANUAL" Then FaultType9_16 = FaultType9_16 + 4
    If InputP12.FaultType = "MANUAL" Then FaultType9_16 = FaultType9_16 + 8
    If InputP13.FaultType = "MANUAL" Then FaultType9_16 = FaultType9_16 + 16
    If InputP14.FaultType = "MANUAL" Then FaultType9_16 = FaultType9_16 + 32
    If InputP15.FaultType = "MANUAL" Then FaultType9_16 = FaultType9_16 + 64
    If InputP16.FaultType = "MANUAL" Then FaultType9_16 = FaultType9_16 + 128
    
    If InputP17.FaultType = "MANUAL" Then FaultType17_24 = FaultType17_24 + 1  ' Fault Type 17-24
    If InputP18.FaultType = "MANUAL" Then FaultType17_24 = FaultType17_24 + 2
    If InputP19.FaultType = "MANUAL" Then FaultType17_24 = FaultType17_24 + 4
    If InputP20.FaultType = "MANUAL" Then FaultType17_24 = FaultType17_24 + 8
        
    ''////////////////////// Data byte 17-24 ///////////////////////////////////
    
    If InputP1.OutputType = "BUZZER" Then OutputType1_8 = OutputType1_8 + 1    ' Output Type 1-8
    If InputP2.OutputType = "BUZZER" Then OutputType1_8 = OutputType1_8 + 2
    If InputP3.OutputType = "BUZZER" Then OutputType1_8 = OutputType1_8 + 4
    If InputP4.OutputType = "BUZZER" Then OutputType1_8 = OutputType1_8 + 8
    If InputP5.OutputType = "BUZZER" Then OutputType1_8 = OutputType1_8 + 16
    If InputP6.OutputType = "BUZZER" Then OutputType1_8 = OutputType1_8 + 32
    If InputP7.OutputType = "BUZZER" Then OutputType1_8 = OutputType1_8 + 64
    If InputP8.OutputType = "BUZZER" Then OutputType1_8 = OutputType1_8 + 128
    
    If InputP9.OutputType = "BUZZER" Then OutputType9_16 = OutputType9_16 + 1  ' Output Type 9-16
    If InputP10.OutputType = "BUZZER" Then OutputType9_16 = OutputType9_16 + 2
    If InputP11.OutputType = "BUZZER" Then OutputType9_16 = OutputType9_16 + 4
    If InputP12.OutputType = "BUZZER" Then OutputType9_16 = OutputType9_16 + 8
    If InputP13.OutputType = "BUZZER" Then OutputType9_16 = OutputType9_16 + 16
    If InputP14.OutputType = "BUZZER" Then OutputType9_16 = OutputType9_16 + 32
    If InputP15.OutputType = "BUZZER" Then OutputType9_16 = OutputType9_16 + 64
    If InputP16.OutputType = "BUZZER" Then OutputType9_16 = OutputType9_16 + 128
    
    If InputP17.OutputType = "BUZZER" Then OutputType17_24 = OutputType17_24 + 1  ' Output Type 17-24
    If InputP18.OutputType = "BUZZER" Then OutputType17_24 = OutputType17_24 + 2
    If InputP19.OutputType = "BUZZER" Then OutputType17_24 = OutputType17_24 + 4
    If InputP20.OutputType = "BUZZER" Then OutputType17_24 = OutputType17_24 + 8
      
    ''////////////////////// Data byte 25-32 ///////////////////////////////////
    
    If InputP1.OutputBoth = "SINGLE" Then OutputBoth1_8 = OutputBoth1_8 + 1    ' Output Both 1-8
    If InputP2.OutputBoth = "SINGLE" Then OutputBoth1_8 = OutputBoth1_8 + 2
    If InputP3.OutputBoth = "SINGLE" Then OutputBoth1_8 = OutputBoth1_8 + 4
    If InputP4.OutputBoth = "SINGLE" Then OutputBoth1_8 = OutputBoth1_8 + 8
    If InputP5.OutputBoth = "SINGLE" Then OutputBoth1_8 = OutputBoth1_8 + 16
    If InputP6.OutputBoth = "SINGLE" Then OutputBoth1_8 = OutputBoth1_8 + 32
    If InputP7.OutputBoth = "SINGLE" Then OutputBoth1_8 = OutputBoth1_8 + 64
    If InputP8.OutputBoth = "SINGLE" Then OutputBoth1_8 = OutputBoth1_8 + 128
    
    If InputP9.OutputBoth = "SINGLE" Then OutputBoth9_16 = OutputBoth9_16 + 1  ' Output Both 9-16
    If InputP10.OutputBoth = "SINGLE" Then OutputBoth9_16 = OutputBoth9_16 + 2
    If InputP11.OutputBoth = "SINGLE" Then OutputBoth9_16 = OutputBoth9_16 + 4
    If InputP12.OutputBoth = "SINGLE" Then OutputBoth9_16 = OutputBoth9_16 + 8
    If InputP13.OutputBoth = "SINGLE" Then OutputBoth9_16 = OutputBoth9_16 + 16
    If InputP14.OutputBoth = "SINGLE" Then OutputBoth9_16 = OutputBoth9_16 + 32
    If InputP15.OutputBoth = "SINGLE" Then OutputBoth9_16 = OutputBoth9_16 + 64
    If InputP16.OutputBoth = "SINGLE" Then OutputBoth9_16 = OutputBoth9_16 + 128
    
    If InputP17.OutputBoth = "SINGLE" Then OutputBoth17_24 = OutputBoth17_24 + 1  ' Output Both 17-24
    If InputP18.OutputBoth = "SINGLE" Then OutputBoth17_24 = OutputBoth17_24 + 2
    If InputP19.OutputBoth = "SINGLE" Then OutputBoth17_24 = OutputBoth17_24 + 4
    If InputP20.OutputBoth = "SINGLE" Then OutputBoth17_24 = OutputBoth17_24 + 8
    
 
    
    ''////////////////////// Data byte 33-40 ///////////////////////////////////
    
    If InputP1.Fault_Indicator = "FAULT" Then Fault_Indicator1_8 = Fault_Indicator1_8 + 1    ' Fault / Indicator 1-8
    If InputP2.Fault_Indicator = "FAULT" Then Fault_Indicator1_8 = Fault_Indicator1_8 + 2
    If InputP3.Fault_Indicator = "FAULT" Then Fault_Indicator1_8 = Fault_Indicator1_8 + 4
    If InputP4.Fault_Indicator = "FAULT" Then Fault_Indicator1_8 = Fault_Indicator1_8 + 8
    If InputP5.Fault_Indicator = "FAULT" Then Fault_Indicator1_8 = Fault_Indicator1_8 + 16
    If InputP6.Fault_Indicator = "FAULT" Then Fault_Indicator1_8 = Fault_Indicator1_8 + 32
    If InputP7.Fault_Indicator = "FAULT" Then Fault_Indicator1_8 = Fault_Indicator1_8 + 64
    If InputP8.Fault_Indicator = "FAULT" Then Fault_Indicator1_8 = Fault_Indicator1_8 + 128
    
    If InputP9.Fault_Indicator = "FAULT" Then Fault_Indicator9_16 = Fault_Indicator9_16 + 1  ' Fault / Indicator 9-16
    If InputP10.Fault_Indicator = "FAULT" Then Fault_Indicator9_16 = Fault_Indicator9_16 + 2
    If InputP11.Fault_Indicator = "FAULT" Then Fault_Indicator9_16 = Fault_Indicator9_16 + 4
    If InputP12.Fault_Indicator = "FAULT" Then Fault_Indicator9_16 = Fault_Indicator9_16 + 8
    If InputP13.Fault_Indicator = "FAULT" Then Fault_Indicator9_16 = Fault_Indicator9_16 + 16
    If InputP14.Fault_Indicator = "FAULT" Then Fault_Indicator9_16 = Fault_Indicator9_16 + 32
    If InputP15.Fault_Indicator = "FAULT" Then Fault_Indicator9_16 = Fault_Indicator9_16 + 64
    If InputP16.Fault_Indicator = "FAULT" Then Fault_Indicator9_16 = Fault_Indicator9_16 + 128
    
    If InputP17.Fault_Indicator = "FAULT" Then Fault_Indicator17_24 = Fault_Indicator17_24 + 1  ' Fault / Indicator 17-24
    If InputP18.Fault_Indicator = "FAULT" Then Fault_Indicator17_24 = Fault_Indicator17_24 + 2
    If InputP19.Fault_Indicator = "FAULT" Then Fault_Indicator17_24 = Fault_Indicator17_24 + 4
    If InputP20.Fault_Indicator = "FAULT" Then Fault_Indicator17_24 = Fault_Indicator17_24 + 8
    
    ''////////////////////// Data byte 41 ///////////////////////////////////
    
    If chkAutoAck.Value = 1 Then
        AutoAckStatus = 15 '// 0x0F
    Else
        AutoAckStatus = 240 '//0xF0
    End If
    
    ''////////////////////// Data byte 42 ///////////////////////////////////
    
    If cbbAutoAckTime.ListIndex < 0 And chkAutoAck.Value = 1 Then
        MsgBox "Please Select Time Delay", vbOKOnly
        CalculateValue = True       ' data not valid
    ElseIf cbbAutoAckTime.ListIndex = 0 Then
        AutoAckTime = 1
        CalculateValue = False
    ElseIf chkAutoAck.Value = 0 Then
        AutoAckTime = 0
        CalculateValue = False
    Else
        AutoAckTime = cbbAutoAckTime.ListIndex + 1
        CalculateValue = False
    End If
    
    ''////////////////////// Data byte 43 ///////////////////////////////////
    
    If OptflashingRate125.Value = True Then FlashRate = 12
    If OptflashingRate250.Value = True Then FlashRate = 25
    If OptflashingRate375.Value = True Then FlashRate = 37
    If OptflashingRate500.Value = True Then FlashRate = 50
    If OptflashingRate625.Value = True Then FlashRate = 62
    If OptflashingRate750.Value = True Then FlashRate = 75
    If OptflashingRate875.Value = True Then FlashRate = 87
    If OptflashingRate1000.Value = True Then FlashRate = 100
    
    ''////////////////////// Data byte 44 ///////////////////////////////////
    
    If OptNoInput8.Value = True Then NoOfInput = 8 '//0x0F
    If OptNoInput10.Value = True Then NoOfInput = 10
    If OptNoInput16.Value = True Then NoOfInput = 16
    If OptNoInput20.Value = True Then NoOfInput = 20
      
    ''////////////////////// Data byte 45 ///////////////////////////////////
    
    FaultDelayTime = cbbFaultDelayTime.ListIndex
    
    ''////////////////////// Data byte 46 ///////////////////////////////////
    DelayTime1 = ComboT1.ListIndex
    ''////////////////////// Data byte 47 ///////////////////////////////////
    DelayTime2 = ComboT2.ListIndex
    ''////////////////////// Data byte 48 ///////////////////////////////////
    DelayTime3 = ComboT3.ListIndex
    ''////////////////////// Data byte 49 ///////////////////////////////////
    DelayTime4 = ComboT4.ListIndex
    ''////////////////////// Data byte 50 ///////////////////////////////////
    DelayTime5 = ComboT5.ListIndex
    ''////////////////////// Data byte 51 ///////////////////////////////////
    DelayTime6 = ComboT6.ListIndex
    ''////////////////////// Data byte 52 ///////////////////////////////////
    DelayTime7 = ComboT7.ListIndex
    ''////////////////////// Data byte 53 ///////////////////////////////////
    DelayTime8 = ComboT8.ListIndex
    ''////////////////////// Data byte 54 ///////////////////////////////////
    DelayTime9 = ComboT9.ListIndex
    ''////////////////////// Data byte 55 ///////////////////////////////////
    DelayTime10 = ComboT10.ListIndex
    
    ''////////////////////// Data byte 56 ///////////////////////////////////
    DelayTime11 = ComboT11.ListIndex
    ''////////////////////// Data byte 57 ///////////////////////////////////
    DelayTime12 = ComboT12.ListIndex
    ''////////////////////// Data byte 58 ///////////////////////////////////
    DelayTime13 = ComboT13.ListIndex
    ''////////////////////// Data byte 59 ///////////////////////////////////
    DelayTime14 = ComboT14.ListIndex
    ''////////////////////// Data byte 60 ///////////////////////////////////
    DelayTime15 = ComboT15.ListIndex
    ''////////////////////// Data byte 61 ///////////////////////////////////
    DelayTime16 = ComboT16.ListIndex
    ''////////////////////// Data byte 62 ///////////////////////////////////
    DelayTime17 = ComboT17.ListIndex
    ''////////////////////// Data byte 63 ///////////////////////////////////
    DelayTime18 = ComboT18.ListIndex
    ''////////////////////// Data byte 64 ///////////////////////////////////
    DelayTime19 = ComboT19.ListIndex
    ''////////////////////// Data byte 65 ///////////////////////////////////
    DelayTime20 = ComboT20.ListIndex
    
    'If OptSyncMaster.Value = True Then
        'MasterSlave = 15 '// 0x0F
    'Else
        'MasterSlave = 240 '//0xF0
    'End If
   
End Function

Private Sub DecodeValue()

'////////////////////////////////////////////////////////////////////////////
    '//////////////////InputType1_8/////////////////
    TextAddrNow.Text = addr
    
    If (DatBuff(1) And 1) = 1 Then
        InputP1.InputType = "NO"
    Else
        InputP1.InputType = "NC"
    End If
    
    If (DatBuff(1) And 2) = 2 Then
        InputP2.InputType = "NO"
    Else
        InputP2.InputType = "NC"
    End If
  
    If (DatBuff(1) And 4) = 4 Then
        InputP3.InputType = "NO"
    Else
        InputP3.InputType = "NC"
    End If
    
    If (DatBuff(1) And 8) = 8 Then
        InputP4.InputType = "NO"
    Else
        InputP4.InputType = "NC"
    End If
    
    If (DatBuff(1) And 16) = 16 Then
        InputP5.InputType = "NO"
    Else
        InputP5.InputType = "NC"
    End If
    
    If (DatBuff(1) And 32) = 32 Then
        InputP6.InputType = "NO"
    Else
        InputP6.InputType = "NC"
    End If
    
    If (DatBuff(1) And 64) = 64 Then
        InputP7.InputType = "NO"
    Else
        InputP7.InputType = "NC"
    End If
    
    If (DatBuff(1) And 128) = 128 Then
        InputP8.InputType = "NO"
    Else
        InputP8.InputType = "NC"
    End If
    
    '//////////////////InputType9_16/////////////////
    If (DatBuff(2) And 1) = 1 Then
        InputP9.InputType = "NO"
    Else
        InputP9.InputType = "NC"
    End If
    
    If (DatBuff(2) And 2) = 2 Then
        InputP10.InputType = "NO"
    Else
        InputP10.InputType = "NC"
    End If
  
    If (DatBuff(2) And 4) = 4 Then
        InputP11.InputType = "NO"
    Else
        InputP11.InputType = "NC"
    End If
    
    If (DatBuff(2) And 8) = 8 Then
        InputP12.InputType = "NO"
    Else
        InputP12.InputType = "NC"
    End If
    
    If (DatBuff(2) And 16) = 16 Then
        InputP13.InputType = "NO"
    Else
        InputP13.InputType = "NC"
    End If
    
    If (DatBuff(2) And 32) = 32 Then
        InputP14.InputType = "NO"
    Else
        InputP14.InputType = "NC"
    End If
    
    If (DatBuff(2) And 64) = 64 Then
        InputP15.InputType = "NO"
    Else
        InputP15.InputType = "NC"
    End If
    
    If (DatBuff(2) And 128) = 128 Then
        InputP16.InputType = "NO"
    Else
        InputP16.InputType = "NC"
    End If
    
    '//////////////////InputType17_24/////////////////
    If (DatBuff(3) And 1) = 1 Then
        InputP17.InputType = "NO"
    Else
        InputP17.InputType = "NC"
    End If
    
    If (DatBuff(3) And 2) = 2 Then
        InputP18.InputType = "NO"
    Else
        InputP18.InputType = "NC"
    End If
  
    If (DatBuff(3) And 4) = 4 Then
        InputP19.InputType = "NO"
    Else
        InputP19.InputType = "NC"
    End If
    
    If (DatBuff(3) And 8) = 8 Then
        InputP20.InputType = "NO"
    Else
        InputP20.InputType = "NC"
    End If
    
   
    

  '////////////////////////////////////////////////////////////////////////////
    '//////////////////FaultType1_8/////////////////
    If (DatBuff(4) And 1) = 1 Then
        InputP1.FaultType = "MANUAL"
    Else
        InputP1.FaultType = "AUTO"
    End If
    
    If (DatBuff(4) And 2) = 2 Then
        InputP2.FaultType = "MANUAL"
    Else
        InputP2.FaultType = "AUTO"
    End If
  
    If (DatBuff(4) And 4) = 4 Then
        InputP3.FaultType = "MANUAL"
    Else
        InputP3.FaultType = "AUTO"
    End If
    
    If (DatBuff(4) And 8) = 8 Then
        InputP4.FaultType = "MANUAL"
    Else
        InputP4.FaultType = "AUTO"
    End If
    
    If (DatBuff(4) And 16) = 16 Then
        InputP5.FaultType = "MANUAL"
    Else
        InputP5.FaultType = "AUTO"
    End If
    
    If (DatBuff(4) And 32) = 32 Then
        InputP6.FaultType = "MANUAL"
    Else
        InputP6.FaultType = "AUTO"
    End If
    
    If (DatBuff(4) And 64) = 64 Then
        InputP7.FaultType = "MANUAL"
    Else
        InputP7.FaultType = "AUTO"
    End If
    
    If (DatBuff(4) And 128) = 128 Then
        InputP8.FaultType = "MANUAL"
    Else
        InputP8.FaultType = "AUTO"
    End If
    
    '//////////////////faultType9_16/////////////////
    If (DatBuff(5) And 1) = 1 Then
        InputP9.FaultType = "MANUAL"
    Else
        InputP9.FaultType = "AUTO"
    End If
    
    If (DatBuff(5) And 2) = 2 Then
        InputP10.FaultType = "MANUAL"
    Else
        InputP10.FaultType = "AUTO"
    End If
  
    If (DatBuff(5) And 4) = 4 Then
        InputP11.FaultType = "MANUAL"
    Else
        InputP11.FaultType = "AUTO"
    End If
    
    If (DatBuff(5) And 8) = 8 Then
        InputP12.FaultType = "MANUAL"
    Else
        InputP12.FaultType = "AUTO"
    End If
    
    If (DatBuff(5) And 16) = 16 Then
        InputP13.FaultType = "MANUAL"
    Else
        InputP13.FaultType = "AUTO"
    End If
    
    If (DatBuff(5) And 32) = 32 Then
        InputP14.FaultType = "MANUAL"
    Else
        InputP14.FaultType = "AUTO"
    End If
    
    If (DatBuff(5) And 64) = 64 Then
        InputP15.FaultType = "MANUAL"
    Else
        InputP15.FaultType = "AUTO"
    End If
    
    If (DatBuff(5) And 128) = 128 Then
        InputP16.FaultType = "MANUAL"
    Else
        InputP16.FaultType = "AUTO"
    End If
    
    '//////////////////faultType17_24/////////////////
    If (DatBuff(6) And 1) = 1 Then
        InputP17.FaultType = "MANUAL"
    Else
        InputP17.FaultType = "AUTO"
    End If
    
    If (DatBuff(6) And 2) = 2 Then
        InputP18.FaultType = "MANUAL"
    Else
        InputP18.FaultType = "AUTO"
    End If
  
    If (DatBuff(6) And 4) = 4 Then
        InputP19.FaultType = "MANUAL"
    Else
        InputP19.FaultType = "AUTO"
    End If
    
    If (DatBuff(6) And 8) = 8 Then
        InputP20.FaultType = "MANUAL"
    Else
        InputP20.FaultType = "AUTO"
    End If
    

  
  '////////////////////////////////////////////////////////////////////////////
    '//////////////////OutputType1_8/////////////////
    If (DatBuff(7) And 1) = 1 Then
        InputP1.OutputType = "BUZZER"
    Else
        InputP1.OutputType = "BELL"
    End If
    
    If (DatBuff(7) And 2) = 2 Then
        InputP2.OutputType = "BUZZER"
    Else
        InputP2.OutputType = "BELL"
    End If
  
    If (DatBuff(7) And 4) = 4 Then
        InputP3.OutputType = "BUZZER"
    Else
        InputP3.OutputType = "BELL"
    End If
    
    If (DatBuff(7) And 8) = 8 Then
        InputP4.OutputType = "BUZZER"
    Else
        InputP4.OutputType = "BELL"
    End If
    
    If (DatBuff(7) And 16) = 16 Then
        InputP5.OutputType = "BUZZER"
    Else
        InputP5.OutputType = "BELL"
    End If
    
    If (DatBuff(7) And 32) = 32 Then
        InputP6.OutputType = "BUZZER"
    Else
        InputP6.OutputType = "BELL"
    End If
    
    If (DatBuff(7) And 64) = 64 Then
        InputP7.OutputType = "BUZZER"
    Else
        InputP7.OutputType = "BELL"
    End If
    
    If (DatBuff(7) And 128) = 128 Then
        InputP8.OutputType = "BUZZER"
    Else
        InputP8.OutputType = "BELL"
    End If
    
    '//////////////////OutputType9_16/////////////////
    If (DatBuff(8) And 1) = 1 Then
        InputP9.OutputType = "BUZZER"
    Else
        InputP9.OutputType = "BELL"
    End If
    
    If (DatBuff(8) And 2) = 2 Then
        InputP10.OutputType = "BUZZER"
    Else
        InputP10.OutputType = "BELL"
    End If
  
    If (DatBuff(8) And 4) = 4 Then
        InputP11.OutputType = "BUZZER"
    Else
        InputP11.OutputType = "BELL"
    End If
    
    If (DatBuff(8) And 8) = 8 Then
        InputP12.OutputType = "BUZZER"
    Else
        InputP12.OutputType = "BELL"
    End If
    
    If (DatBuff(8) And 16) = 16 Then
        InputP13.OutputType = "BUZZER"
    Else
        InputP13.OutputType = "BELL"
    End If
    
    If (DatBuff(8) And 32) = 32 Then
        InputP14.OutputType = "BUZZER"
    Else
        InputP14.OutputType = "BELL"
    End If
    
    If (DatBuff(8) And 64) = 64 Then
        InputP15.OutputType = "BUZZER"
    Else
        InputP15.OutputType = "BELL"
    End If
    
    If (DatBuff(8) And 128) = 128 Then
        InputP16.OutputType = "BUZZER"
    Else
        InputP16.OutputType = "BELL"
    End If
    
    '//////////////////OutputType17_24/////////////////
    If (DatBuff(9) And 1) = 1 Then
        InputP17.OutputType = "BUZZER"
    Else
        InputP17.OutputType = "BELL"
    End If
    
    If (DatBuff(9) And 2) = 2 Then
        InputP18.OutputType = "BUZZER"
    Else
        InputP18.OutputType = "BELL"
    End If
  
    If (DatBuff(9) And 4) = 4 Then
        InputP19.OutputType = "BUZZER"
    Else
        InputP19.OutputType = "BELL"
    End If
    
    If (DatBuff(9) And 8) = 8 Then
        InputP20.OutputType = "BUZZER"
    Else
        InputP20.OutputType = "BELL"
    End If
    


'////////////////////////////////////////////////////////////////////////////
    '//////////////////OutputBoth1_8/////////////////
    If (DatBuff(10) And 1) = 1 Then
        InputP1.OutputBoth = "SINGLE"
    Else
        InputP1.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(10) And 2) = 2 Then
        InputP2.OutputBoth = "SINGLE"
    Else
        InputP2.OutputBoth = "BOTH"
    End If
  
    If (DatBuff(10) And 4) = 4 Then
        InputP3.OutputBoth = "SINGLE"
    Else
        InputP3.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(10) And 8) = 8 Then
        InputP4.OutputBoth = "SINGLE"
    Else
        InputP4.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(10) And 16) = 16 Then
        InputP5.OutputBoth = "SINGLE"
    Else
        InputP5.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(10) And 32) = 32 Then
        InputP6.OutputBoth = "SINGLE"
    Else
        InputP6.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(10) And 64) = 64 Then
        InputP7.OutputBoth = "SINGLE"
    Else
        InputP7.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(10) And 128) = 128 Then
        InputP8.OutputBoth = "SINGLE"
    Else
        InputP8.OutputBoth = "BOTH"
    End If
    
    '//////////////////OutputBoth9_16/////////////////
    If (DatBuff(11) And 1) = 1 Then
        InputP9.OutputBoth = "SINGLE"
    Else
        InputP9.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(11) And 2) = 2 Then
        InputP10.OutputBoth = "SINGLE"
    Else
        InputP10.OutputBoth = "BOTH"
    End If
  
    If (DatBuff(11) And 4) = 4 Then
        InputP11.OutputBoth = "SINGLE"
    Else
        InputP11.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(11) And 8) = 8 Then
        InputP12.OutputBoth = "SINGLE"
    Else
        InputP12.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(11) And 16) = 16 Then
        InputP13.OutputBoth = "SINGLE"
    Else
        InputP13.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(11) And 32) = 32 Then
        InputP14.OutputBoth = "SINGLE"
    Else
        InputP14.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(11) And 64) = 64 Then
        InputP15.OutputBoth = "SINGLE"
    Else
        InputP15.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(11) And 128) = 128 Then
        InputP16.OutputBoth = "SINGLE"
    Else
        InputP16.OutputBoth = "BOTH"
    End If
    
    '//////////////////OutputBoth17_24/////////////////
    If (DatBuff(12) And 1) = 1 Then
        InputP17.OutputBoth = "SINGLE"
    Else
        InputP17.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(12) And 2) = 2 Then
        InputP18.OutputBoth = "SINGLE"
    Else
        InputP18.OutputBoth = "BOTH"
    End If
  
    If (DatBuff(12) And 4) = 4 Then
        InputP19.OutputBoth = "SINGLE"
    Else
        InputP19.OutputBoth = "BOTH"
    End If
    
    If (DatBuff(12) And 8) = 8 Then
        InputP20.OutputBoth = "SINGLE"
    Else
        InputP20.OutputBoth = "BOTH"
    End If
    

    
'////////////////////////////////////////////////////////////////////////////
    '//////////////////Fault_Indicator1_8/////////////////
    If (DatBuff(13) And 1) = 1 Then
        InputP1.Fault_Indicator = "FAULT"
    Else
        InputP1.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(13) And 2) = 2 Then
        InputP2.Fault_Indicator = "FAULT"
    Else
        InputP2.Fault_Indicator = "INDICATOR"
    End If
  
    If (DatBuff(13) And 4) = 4 Then
        InputP3.Fault_Indicator = "FAULT"
    Else
        InputP3.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(13) And 8) = 8 Then
        InputP4.Fault_Indicator = "FAULT"
    Else
        InputP4.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(13) And 16) = 16 Then
        InputP5.Fault_Indicator = "FAULT"
    Else
        InputP5.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(13) And 32) = 32 Then
        InputP6.Fault_Indicator = "FAULT"
    Else
        InputP6.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(13) And 64) = 64 Then
        InputP7.Fault_Indicator = "FAULT"
    Else
        InputP7.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(13) And 128) = 128 Then
        InputP8.Fault_Indicator = "FAULT"
    Else
        InputP8.Fault_Indicator = "INDICATOR"
    End If
    
    '//////////////////Fault_Indicator9_16/////////////////
    If (DatBuff(14) And 1) = 1 Then
        InputP9.Fault_Indicator = "FAULT"
    Else
        InputP9.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(14) And 2) = 2 Then
        InputP10.Fault_Indicator = "FAULT"
    Else
        InputP10.Fault_Indicator = "INDICATOR"
    End If
  
    If (DatBuff(14) And 4) = 4 Then
        InputP11.Fault_Indicator = "FAULT"
    Else
        InputP11.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(14) And 8) = 8 Then
        InputP12.Fault_Indicator = "FAULT"
    Else
        InputP12.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(14) And 16) = 16 Then
        InputP13.Fault_Indicator = "FAULT"
    Else
        InputP13.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(14) And 32) = 32 Then
        InputP14.Fault_Indicator = "FAULT"
    Else
        InputP14.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(14) And 64) = 64 Then
        InputP15.Fault_Indicator = "FAULT"
    Else
        InputP15.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(14) And 128) = 128 Then
        InputP16.Fault_Indicator = "FAULT"
    Else
        InputP16.Fault_Indicator = "INDICATOR"
    End If
    
    '//////////////////Fault_Indicator17_24/////////////////
    If (DatBuff(15) And 1) = 1 Then
        InputP17.Fault_Indicator = "FAULT"
    Else
        InputP17.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(15) And 2) = 2 Then
        InputP18.Fault_Indicator = "FAULT"
    Else
        InputP18.Fault_Indicator = "INDICATOR"
    End If
  
    If (DatBuff(15) And 4) = 4 Then
        InputP19.Fault_Indicator = "FAULT"
    Else
        InputP19.Fault_Indicator = "INDICATOR"
    End If
    
    If (DatBuff(15) And 8) = 8 Then
        InputP20.Fault_Indicator = "FAULT"
    Else
        InputP20.Fault_Indicator = "INDICATOR"
    End If
    
    '////////////////////////////////////////////////////////////////////////////
    '//////////////////LED Colour1_8/////////////////
    '/////////////////////////////////////////
    '///////////////////////////////////////////////////////////////
    If ((DatBuff(16) And 1) = 1) And ((DatBuff(20) And 1) = 0) Then
       R1.Value = True
    ElseIf ((DatBuff(16) And 1) = 0) And ((DatBuff(20) And 1) = 1) Then
        G1.Value = True
    ElseIf ((DatBuff(16) And 1) = 1) And ((DatBuff(20) And 1) = 1) Then
        A1.Value = True
        
    End If
    'MsgBox ("END1.")
    '/////////////////////////////////////////
    If (DatBuff(16) And 2) = 2 And (DatBuff(20) And 2) = 0 Then
        R2.Value = True
    ElseIf (DatBuff(16) And 2) = 0 And (DatBuff(20) And 2) = 2 Then
        G2.Value = True
    ElseIf (DatBuff(16) And 2) = 2 And (DatBuff(20) And 2) = 2 Then
        A2.Value = True
    End If
    
    '/////////////////////////////////////////
    If (DatBuff(16) And 4) = 4 And (DatBuff(20) And 4) = 0 Then
        R3.Value = True
    ElseIf (DatBuff(16) And 4) = 0 And (DatBuff(20) And 4) = 4 Then
        G3.Value = True
    ElseIf (DatBuff(16) And 4) = 4 And (DatBuff(20) And 4) = 4 Then
            A3.Value = True
    End If
    '/////////////////////////////////////////
    If (DatBuff(16) And 8) = 8 And (DatBuff(20) And 8) = 0 Then
        R4.Value = True
    ElseIf (DatBuff(16) And 8) = 0 And (DatBuff(20) And 8) = 8 Then
        G4.Value = True
    ElseIf (DatBuff(16) And 8) = 8 And (DatBuff(20) And 8) = 8 Then
            A4.Value = True
        'End If
    End If
    '/////////////////////////////////////////
    If (DatBuff(16) And 16) = 16 And (DatBuff(20) And 16) = 0 Then
        R5.Value = True
    ElseIf (DatBuff(16) And 16) = 0 And (DatBuff(20) And 16) = 16 Then
        G5.Value = True
    ElseIf (DatBuff(16) And 16) = 16 And (DatBuff(20) And 16) = 16 Then
            A5.Value = True
        'End If
    End If
    '/////////////////////////////////////////
    If (DatBuff(16) And 32) = 32 And (DatBuff(20) And 32) = 0 Then
        R6.Value = True
    ElseIf (DatBuff(16) And 32) = 0 And (DatBuff(20) And 32) = 32 Then
        G6.Value = True
    ElseIf (DatBuff(16) And 32) = 32 And (DatBuff(20) And 32) = 32 Then
            A6.Value = True
        'End If
    End If
    '/////////////////////////////////////////
    If (DatBuff(16) And 64) = 64 And (DatBuff(20) And 64) = 0 Then
        R7.Value = True
    ElseIf (DatBuff(16) And 64) = 0 And (DatBuff(20) And 64) = 64 Then
        G7.Value = True
    ElseIf (DatBuff(16) And 64) = 64 And (DatBuff(20) And 64) = 64 Then
            A7.Value = True
        'End If
    End If
    '/////////////////////////////////////////
    If (DatBuff(16) And 128) = 128 And (DatBuff(20) And 128) = 0 Then
        R8.Value = True
    ElseIf (DatBuff(16) And 128) = 0 And (DatBuff(20) And 128) = 128 Then
        G8.Value = True
    ElseIf (DatBuff(16) And 128) = 128 And (DatBuff(20) And 128) = 128 Then
            A8.Value = True
        'End If
    End If
    '/////////////////////////////////////////
    If (DatBuff(17) And 1) = 1 And (DatBuff(21) And 1) = 0 Then
        R9.Value = True
    ElseIf (DatBuff(17) And 1) = 0 And (DatBuff(21) And 1) = 1 Then
        G9.Value = True
    Else
        If (DatBuff(17) And 1) = 1 And (DatBuff(21) And 1) = 1 Then
            A9.Value = True
        End If
    End If
    '/////////////////////////////////////////
    If (DatBuff(17) And 2) = 2 And (DatBuff(21) And 2) = 0 Then
        R10.Value = True
    ElseIf (DatBuff(17) And 2) = 0 And (DatBuff(21) And 2) = 2 Then
        G10.Value = True
    Else
        If (DatBuff(17) And 2) = 2 And (DatBuff(21) And 2) = 2 Then
            A10.Value = True
        End If
    End If
    '/////////////////////////////////////////
    If (DatBuff(18) And 1) = 1 And (DatBuff(22) And 1) = 0 Then
        R11.Value = True
    ElseIf (DatBuff(18) And 1) = 0 And (DatBuff(22) And 1) = 1 Then
        G11.Value = True
    Else
        If (DatBuff(18) And 1) = 1 And (DatBuff(22) And 1) = 1 Then
            A11.Value = True
        End If
    End If
    '/////////////////////////////////////////
     If (DatBuff(18) And 2) = 2 And (DatBuff(22) And 2) = 0 Then
        R12.Value = True
    ElseIf (DatBuff(18) And 2) = 0 And (DatBuff(22) And 2) = 2 Then
        G12.Value = True
    Else
        If (DatBuff(18) And 2) = 2 And (DatBuff(22) And 2) = 2 Then
            A12.Value = True
        End If
    End If
    '/////////////////////////////////////////
     If (DatBuff(18) And 4) = 4 And (DatBuff(22) And 4) = 0 Then
        R13.Value = True
    ElseIf (DatBuff(18) And 4) = 0 And (DatBuff(22) And 4) = 4 Then
        G13.Value = True
    Else
        If (DatBuff(18) And 4) = 4 And (DatBuff(22) And 4) = 4 Then
            A13.Value = True
        End If
    End If
    '/////////////////////////////////////////
     If (DatBuff(18) And 8) = 8 And (DatBuff(22) And 8) = 0 Then
        R14.Value = True
    ElseIf (DatBuff(18) And 8) = 0 And (DatBuff(22) And 8) = 8 Then
        G14.Value = True
    Else
        If (DatBuff(18) And 8) = 8 And (DatBuff(22) And 8) = 8 Then
            A14.Value = True
        End If
    End If
    '/////////////////////////////////////////
     If (DatBuff(18) And 16) = 16 And (DatBuff(22) And 16) = 0 Then
        R15.Value = True
    ElseIf (DatBuff(18) And 16) = 0 And (DatBuff(22) And 16) = 16 Then
        G15.Value = True
    Else
        If (DatBuff(18) And 16) = 16 And (DatBuff(22) And 16) = 16 Then
            A15.Value = True
        End If
    End If
    '/////////////////////////////////////////
     If (DatBuff(18) And 32) = 32 And (DatBuff(22) And 32) = 0 Then
        R16.Value = True
    ElseIf (DatBuff(18) And 32) = 0 And (DatBuff(22) And 32) = 32 Then
        G16.Value = True
    Else
        If (DatBuff(18) And 32) = 32 And (DatBuff(22) And 32) = 32 Then
            A16.Value = True
        End If
    End If
    '/////////////////////////////////////////
     If (DatBuff(18) And 64) = 64 And (DatBuff(22) And 64) = 0 Then
        R17.Value = True
    ElseIf (DatBuff(18) And 64) = 0 And (DatBuff(22) And 64) = 64 Then
        G17.Value = True
    Else
        If (DatBuff(18) And 64) = 64 And (DatBuff(22) And 64) = 64 Then
            A17.Value = True
        End If
    End If
    '/////////////////////////////////////////
     If (DatBuff(18) And 128) = 128 And (DatBuff(22) And 128) = 0 Then
        R18.Value = True
    ElseIf (DatBuff(18) And 128) = 0 And (DatBuff(22) And 128) = 128 Then
        G18.Value = True
    Else
        If (DatBuff(18) And 128) = 128 And (DatBuff(22) And 128) = 128 Then
            A18.Value = True
        End If
    End If
    '/////////////////////////////////////////
     If (DatBuff(19) And 1) = 1 And (DatBuff(23) And 1) = 0 Then
        R19.Value = True
    ElseIf (DatBuff(19) And 1) = 0 And (DatBuff(23) And 1) = 1 Then
        G19.Value = True
    Else
        If (DatBuff(19) And 1) = 1 And (DatBuff(23) And 1) = 1 Then
            A19.Value = True
        End If
    End If
    '/////////////////////////////////////////
    If (DatBuff(19) And 2) = 2 And (DatBuff(23) And 2) = 0 Then
        R20.Value = True
    ElseIf (DatBuff(19) And 2) = 0 And (DatBuff(23) And 2) = 2 Then
        G20.Value = True
    Else
        If (DatBuff(19) And 2) = 2 And (DatBuff(23) And 2) = 2 Then
            A20.Value = True
        End If
    End If
    '/////////////////////////////////////////
    
   
  '/////////////////////////////////////////////////////////////////////////////////////////////
  '////////////////////Auto Acknoledge Status //////////////////////////////
    If DatBuff(24) = 15 Then
        chkAutoAck.Value = 1
    ElseIf DatBuff(24) = 240 Then
        chkAutoAck.Value = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    
    '/////////////////////////////////////////////////////////////////////////////////////////////
  '////////////////////Auto Acknowledge Time //////////////////////////////
    If DatBuff(25) = 0 Then
        chkAutoAck.Value = 0
        chkAutoAck_Click
        cbbAutoAckTime.ListIndex = -1
    ElseIf DatBuff(25) > 239 Then
        cbbAutoAckTime.ListIndex = -1
    Else
        cbbAutoAckTime.ListIndex = DatBuff(42) - 1
    End If
    '//////////////////////////////////////////////////////////////////////////
    
  '/////////////////////////////////////////////////////////////////////////////////////////////
  '/////////////////////////////////////////////////////////////////////////////////////////////
  '////////////////////Flashing Rate //////////////////////////////
    If DatBuff(26) = 12 Then
        OptflashingRate125.Value = True
    ElseIf DatBuff(26) = 25 Then
        OptflashingRate250.Value = True
    ElseIf DatBuff(26) = 37 Then
        OptflashingRate375.Value = True
    ElseIf DatBuff(26) = 50 Then
        OptflashingRate500.Value = True
    ElseIf DatBuff(26) = 62 Then
        OptflashingRate625.Value = True
    ElseIf DatBuff(26) = 75 Then
        OptflashingRate750.Value = True
    ElseIf DatBuff(26) = 87 Then
        OptflashingRate875.Value = True
    ElseIf DatBuff(26) = 100 Then
        OptflashingRate1000.Value = True
    End If
    '//////////////////////////////////////////////////////////////////////////
    
    '/////////////////////////////////////////////////////////////////////////////////////////////
  '//////////////////// No of Input //////////////////////////////
  
    If DatBuff(27) = 8 Then
        OptNoInput8.Value = True
    ElseIf DatBuff(27) = 10 Then
        OptNoInput10.Value = True
    ElseIf DatBuff(27) = 16 Then
        OptNoInput16.Value = True
    ElseIf DatBuff(27) = 20 Then
        OptNoInput20.Value = True
    'ElseIf DatBuff(28) = 48 Then
       ' OptNoInput48.Value = True
    'ElseIf DatBuff(28) = 56 Then
        'OptNoInput56.Value = True
    'ElseIf DatBuff(28) = 64 Then
        'OptNoInput64.Value = True
    End If
    '//////////////////////////////////////////////////////////////////////////
    
    '/////////////////////////////////////////////////////////////////////////////////////////////
  
    '/////////////////////////FaultDelayTime/////////////////////////////////////////////////
    If DatBuff(28) = 0 Then
        cbbFaultDelayTime.ListIndex = 0
    ElseIf DatBuff(28) = 1 Then
        cbbFaultDelayTime.ListIndex = 1
    ElseIf DatBuff(28) = 2 Then
        cbbFaultDelayTime.ListIndex = 2
    ElseIf DatBuff(28) = 3 Then
        cbbFaultDelayTime.ListIndex = 3
    ElseIf DatBuff(28) = 4 Then
        cbbFaultDelayTime.ListIndex = 4
    ElseIf DatBuff(28) = 5 Then
        cbbFaultDelayTime.ListIndex = 5
    ElseIf DatBuff(28) = 6 Then
        cbbFaultDelayTime.ListIndex = 6
    ElseIf DatBuff(28) = 7 Then
        cbbFaultDelayTime.ListIndex = 7
    ElseIf DatBuff(28) = 8 Then
        cbbFaultDelayTime.ListIndex = 8
    ElseIf DatBuff(28) = 9 Then
        cbbFaultDelayTime.ListIndex = 9
    ElseIf DatBuff(28) = 10 Then
        cbbFaultDelayTime.ListIndex = 10
    Else
        cbbFaultDelayTime.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    
    '//////////////////// Master / Slave //////////////////////////////
    'If DatBuff(29) = 15 Then
        'OptSyncMaster.Value = True
    'ElseIf DatBuff(29) = 240 Then
        'OptSyncSlave.Value = True

    'End If
    '//////////////////////////////////////////////////
    
    
    If DatBuff(30) = 0 Then
        ComboT1.ListIndex = 0
    ElseIf DatBuff(30) = 1 Then
        ComboT1.ListIndex = 1
    ElseIf DatBuff(30) = 2 Then
        ComboT1.ListIndex = 2
    ElseIf DatBuff(30) = 3 Then
        ComboT1.ListIndex = 3
    ElseIf DatBuff(30) = 4 Then
        ComboT1.ListIndex = 4
    ElseIf DatBuff(30) = 5 Then
        ComboT1.ListIndex = 5
    ElseIf DatBuff(30) = 6 Then
        ComboT1.ListIndex = 6
    ElseIf DatBuff(30) = 7 Then
        ComboT1.ListIndex = 7
    ElseIf DatBuff(30) = 8 Then
        ComboT1.ListIndex = 8
    ElseIf DatBuff(30) = 9 Then
        ComboT1.ListIndex = 9
    ElseIf DatBuff(30) = 10 Then
        ComboT1.ListIndex = 10
    ElseIf DatBuff(30) = 11 Then
        ComboT1.ListIndex = 11
    ElseIf DatBuff(30) = 12 Then
        ComboT1.ListIndex = 12
    Else
        ComboT1.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(31) = 0 Then
        ComboT2.ListIndex = 0
    ElseIf DatBuff(31) = 1 Then
        ComboT2.ListIndex = 1
    ElseIf DatBuff(31) = 2 Then
        ComboT2.ListIndex = 2
    ElseIf DatBuff(31) = 3 Then
        ComboT2.ListIndex = 3
    ElseIf DatBuff(31) = 4 Then
        ComboT2.ListIndex = 4
    ElseIf DatBuff(31) = 5 Then
        ComboT2.ListIndex = 5
    ElseIf DatBuff(31) = 6 Then
        ComboT2.ListIndex = 6
    ElseIf DatBuff(31) = 7 Then
        ComboT2.ListIndex = 7
    ElseIf DatBuff(31) = 8 Then
        ComboT2.ListIndex = 8
    ElseIf DatBuff(31) = 9 Then
        ComboT2.ListIndex = 9
    ElseIf DatBuff(31) = 10 Then
        ComboT2.ListIndex = 10
    ElseIf DatBuff(31) = 11 Then
        ComboT2.ListIndex = 11
    ElseIf DatBuff(31) = 12 Then
        ComboT2.ListIndex = 12
    Else
        ComboT2.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(32) = 0 Then
        ComboT3.ListIndex = 0
    ElseIf DatBuff(32) = 1 Then
        ComboT3.ListIndex = 1
    ElseIf DatBuff(32) = 2 Then
        ComboT3.ListIndex = 2
    ElseIf DatBuff(32) = 3 Then
        ComboT3.ListIndex = 3
    ElseIf DatBuff(32) = 4 Then
        ComboT3.ListIndex = 4
    ElseIf DatBuff(32) = 5 Then
        ComboT3.ListIndex = 5
    ElseIf DatBuff(32) = 6 Then
        ComboT3.ListIndex = 6
    ElseIf DatBuff(32) = 7 Then
        ComboT3.ListIndex = 7
    ElseIf DatBuff(32) = 8 Then
        ComboT3.ListIndex = 8
    ElseIf DatBuff(32) = 9 Then
        ComboT3.ListIndex = 9
    ElseIf DatBuff(32) = 10 Then
        ComboT3.ListIndex = 10
    ElseIf DatBuff(32) = 11 Then
        ComboT3.ListIndex = 11
    ElseIf DatBuff(32) = 12 Then
        ComboT3.ListIndex = 12
    Else
        ComboT3.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(33) = 0 Then
        ComboT4.ListIndex = 0
    ElseIf DatBuff(33) = 1 Then
        ComboT4.ListIndex = 1
    ElseIf DatBuff(33) = 2 Then
        ComboT4.ListIndex = 2
    ElseIf DatBuff(33) = 3 Then
        ComboT4.ListIndex = 3
    ElseIf DatBuff(33) = 4 Then
        ComboT4.ListIndex = 4
    ElseIf DatBuff(33) = 5 Then
        ComboT4.ListIndex = 5
    ElseIf DatBuff(33) = 6 Then
        ComboT4.ListIndex = 6
    ElseIf DatBuff(33) = 7 Then
        ComboT4.ListIndex = 7
    ElseIf DatBuff(33) = 8 Then
        ComboT4.ListIndex = 8
    ElseIf DatBuff(33) = 9 Then
        ComboT4.ListIndex = 9
    ElseIf DatBuff(33) = 10 Then
        ComboT4.ListIndex = 10
    ElseIf DatBuff(33) = 11 Then
        ComboT4.ListIndex = 11
    ElseIf DatBuff(33) = 12 Then
        ComboT4.ListIndex = 12
    Else
        ComboT4.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(34) = 0 Then
        ComboT5.ListIndex = 0
    ElseIf DatBuff(34) = 1 Then
        ComboT5.ListIndex = 1
    ElseIf DatBuff(34) = 2 Then
        ComboT5.ListIndex = 2
    ElseIf DatBuff(34) = 3 Then
        ComboT5.ListIndex = 3
    ElseIf DatBuff(34) = 4 Then
        ComboT5.ListIndex = 4
    ElseIf DatBuff(34) = 5 Then
        ComboT5.ListIndex = 5
    ElseIf DatBuff(34) = 6 Then
        ComboT5.ListIndex = 6
    ElseIf DatBuff(34) = 7 Then
        ComboT5.ListIndex = 7
    ElseIf DatBuff(34) = 8 Then
        ComboT5.ListIndex = 8
    ElseIf DatBuff(34) = 9 Then
        ComboT5.ListIndex = 9
    ElseIf DatBuff(34) = 10 Then
        ComboT5.ListIndex = 10
    ElseIf DatBuff(34) = 11 Then
        ComboT5.ListIndex = 11
    ElseIf DatBuff(34) = 12 Then
        ComboT5.ListIndex = 12
    Else
        ComboT5.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(35) = 0 Then
        ComboT6.ListIndex = 0
    ElseIf DatBuff(35) = 1 Then
        ComboT6.ListIndex = 1
    ElseIf DatBuff(35) = 2 Then
        ComboT6.ListIndex = 2
    ElseIf DatBuff(35) = 3 Then
        ComboT6.ListIndex = 3
    ElseIf DatBuff(35) = 4 Then
        ComboT6.ListIndex = 4
    ElseIf DatBuff(35) = 5 Then
        ComboT6.ListIndex = 5
    ElseIf DatBuff(35) = 6 Then
        ComboT6.ListIndex = 6
    ElseIf DatBuff(35) = 7 Then
        ComboT6.ListIndex = 7
    ElseIf DatBuff(35) = 8 Then
        ComboT6.ListIndex = 8
    ElseIf DatBuff(35) = 9 Then
        ComboT6.ListIndex = 9
    ElseIf DatBuff(35) = 10 Then
        ComboT6.ListIndex = 10
    ElseIf DatBuff(35) = 11 Then
        ComboT6.ListIndex = 11
    ElseIf DatBuff(35) = 12 Then
        ComboT6.ListIndex = 12
    Else
        ComboT6.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(36) = 0 Then
        ComboT7.ListIndex = 0
    ElseIf DatBuff(36) = 1 Then
        ComboT7.ListIndex = 1
    ElseIf DatBuff(36) = 2 Then
        ComboT7.ListIndex = 2
    ElseIf DatBuff(36) = 3 Then
        ComboT7.ListIndex = 3
    ElseIf DatBuff(36) = 4 Then
        ComboT7.ListIndex = 4
    ElseIf DatBuff(36) = 5 Then
        ComboT7.ListIndex = 5
    ElseIf DatBuff(36) = 6 Then
        ComboT7.ListIndex = 6
    ElseIf DatBuff(36) = 7 Then
        ComboT7.ListIndex = 7
    ElseIf DatBuff(36) = 8 Then
        ComboT7.ListIndex = 8
    ElseIf DatBuff(36) = 9 Then
        ComboT7.ListIndex = 9
    ElseIf DatBuff(36) = 10 Then
        ComboT7.ListIndex = 10
    ElseIf DatBuff(36) = 11 Then
        ComboT7.ListIndex = 11
    ElseIf DatBuff(36) = 12 Then
        ComboT7.ListIndex = 12
    Else
        ComboT7.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(37) = 0 Then
        ComboT8.ListIndex = 0
    ElseIf DatBuff(37) = 1 Then
        ComboT8.ListIndex = 1
    ElseIf DatBuff(37) = 2 Then
        ComboT8.ListIndex = 2
    ElseIf DatBuff(37) = 3 Then
        ComboT8.ListIndex = 3
    ElseIf DatBuff(37) = 4 Then
        ComboT8.ListIndex = 4
    ElseIf DatBuff(37) = 5 Then
        ComboT8.ListIndex = 5
    ElseIf DatBuff(37) = 6 Then
        ComboT8.ListIndex = 6
    ElseIf DatBuff(37) = 7 Then
        ComboT8.ListIndex = 7
    ElseIf DatBuff(37) = 8 Then
        ComboT8.ListIndex = 8
    ElseIf DatBuff(37) = 9 Then
        ComboT8.ListIndex = 9
    ElseIf DatBuff(37) = 10 Then
        ComboT8.ListIndex = 10
    ElseIf DatBuff(37) = 11 Then
        ComboT8.ListIndex = 11
    ElseIf DatBuff(37) = 12 Then
        ComboT8.ListIndex = 12
    Else
        ComboT8.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(38) = 0 Then
        ComboT9.ListIndex = 0
    ElseIf DatBuff(38) = 1 Then
        ComboT9.ListIndex = 1
    ElseIf DatBuff(38) = 2 Then
        ComboT9.ListIndex = 2
    ElseIf DatBuff(38) = 3 Then
        ComboT9.ListIndex = 3
    ElseIf DatBuff(38) = 4 Then
        ComboT9.ListIndex = 4
    ElseIf DatBuff(38) = 5 Then
        ComboT9.ListIndex = 5
    ElseIf DatBuff(38) = 6 Then
        ComboT9.ListIndex = 6
    ElseIf DatBuff(38) = 7 Then
        ComboT9.ListIndex = 7
    ElseIf DatBuff(38) = 8 Then
        ComboT9.ListIndex = 8
    ElseIf DatBuff(38) = 9 Then
        ComboT9.ListIndex = 9
    ElseIf DatBuff(38) = 10 Then
        ComboT9.ListIndex = 10
    ElseIf DatBuff(38) = 11 Then
        ComboT9.ListIndex = 11
    ElseIf DatBuff(38) = 12 Then
        ComboT9.ListIndex = 12
    Else
        ComboT9.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(39) = 0 Then
        ComboT10.ListIndex = 0
    ElseIf DatBuff(39) = 1 Then
        ComboT10.ListIndex = 1
    ElseIf DatBuff(39) = 2 Then
        ComboT10.ListIndex = 2
    ElseIf DatBuff(39) = 3 Then
        ComboT10.ListIndex = 3
    ElseIf DatBuff(39) = 4 Then
        ComboT10.ListIndex = 4
    ElseIf DatBuff(39) = 5 Then
        ComboT10.ListIndex = 5
    ElseIf DatBuff(39) = 6 Then
        ComboT10.ListIndex = 6
    ElseIf DatBuff(39) = 7 Then
        ComboT10.ListIndex = 7
    ElseIf DatBuff(39) = 8 Then
        ComboT10.ListIndex = 8
    ElseIf DatBuff(39) = 9 Then
        ComboT10.ListIndex = 9
    ElseIf DatBuff(39) = 10 Then
        ComboT10.ListIndex = 10
    ElseIf DatBuff(39) = 11 Then
        ComboT10.ListIndex = 11
    ElseIf DatBuff(39) = 12 Then
        ComboT10.ListIndex = 12
    Else
        ComboT10.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(40) = 0 Then
        ComboT11.ListIndex = 0
    ElseIf DatBuff(40) = 1 Then
        ComboT11.ListIndex = 1
    ElseIf DatBuff(40) = 2 Then
        ComboT11.ListIndex = 2
    ElseIf DatBuff(40) = 3 Then
        ComboT11.ListIndex = 3
    ElseIf DatBuff(40) = 4 Then
        ComboT11.ListIndex = 4
    ElseIf DatBuff(40) = 5 Then
        ComboT11.ListIndex = 5
    ElseIf DatBuff(40) = 6 Then
        ComboT11.ListIndex = 6
    ElseIf DatBuff(40) = 7 Then
        ComboT11.ListIndex = 7
    ElseIf DatBuff(40) = 8 Then
        ComboT11.ListIndex = 8
    ElseIf DatBuff(40) = 9 Then
        ComboT11.ListIndex = 9
    ElseIf DatBuff(40) = 10 Then
        ComboT11.ListIndex = 10
    ElseIf DatBuff(40) = 11 Then
        ComboT11.ListIndex = 11
    ElseIf DatBuff(40) = 12 Then
        ComboT11.ListIndex = 12
    Else
        ComboT11.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(41) = 0 Then
        ComboT12.ListIndex = 0
    ElseIf DatBuff(41) = 1 Then
        ComboT12.ListIndex = 1
    ElseIf DatBuff(41) = 2 Then
        ComboT12.ListIndex = 2
    ElseIf DatBuff(41) = 3 Then
        ComboT12.ListIndex = 3
    ElseIf DatBuff(41) = 4 Then
        ComboT12.ListIndex = 4
    ElseIf DatBuff(41) = 5 Then
        ComboT12.ListIndex = 5
    ElseIf DatBuff(41) = 6 Then
        ComboT12.ListIndex = 6
    ElseIf DatBuff(41) = 7 Then
        ComboT12.ListIndex = 7
    ElseIf DatBuff(41) = 8 Then
        ComboT12.ListIndex = 8
    ElseIf DatBuff(41) = 9 Then
        ComboT12.ListIndex = 9
    ElseIf DatBuff(41) = 10 Then
        ComboT12.ListIndex = 10
    ElseIf DatBuff(41) = 11 Then
        ComboT12.ListIndex = 11
    ElseIf DatBuff(41) = 12 Then
        ComboT12.ListIndex = 12
    Else
        ComboT12.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(42) = 0 Then
        ComboT13.ListIndex = 0
    ElseIf DatBuff(42) = 1 Then
        ComboT13.ListIndex = 1
    ElseIf DatBuff(42) = 2 Then
        ComboT13.ListIndex = 2
    ElseIf DatBuff(42) = 3 Then
        ComboT13.ListIndex = 3
    ElseIf DatBuff(42) = 4 Then
        ComboT13.ListIndex = 4
    ElseIf DatBuff(42) = 5 Then
        ComboT13.ListIndex = 5
    ElseIf DatBuff(42) = 6 Then
        ComboT13.ListIndex = 6
    ElseIf DatBuff(42) = 7 Then
        ComboT13.ListIndex = 7
    ElseIf DatBuff(42) = 8 Then
        ComboT13.ListIndex = 8
    ElseIf DatBuff(42) = 9 Then
        ComboT13.ListIndex = 9
    ElseIf DatBuff(42) = 10 Then
        ComboT13.ListIndex = 10
    ElseIf DatBuff(42) = 11 Then
        ComboT13.ListIndex = 11
    ElseIf DatBuff(42) = 12 Then
        ComboT13.ListIndex = 12
    Else
        ComboT13.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(43) = 0 Then
        ComboT14.ListIndex = 0
    ElseIf DatBuff(43) = 1 Then
        ComboT14.ListIndex = 1
    ElseIf DatBuff(43) = 2 Then
        ComboT14.ListIndex = 2
    ElseIf DatBuff(43) = 3 Then
        ComboT14.ListIndex = 3
    ElseIf DatBuff(43) = 4 Then
        ComboT14.ListIndex = 4
    ElseIf DatBuff(43) = 5 Then
        ComboT14.ListIndex = 5
    ElseIf DatBuff(43) = 6 Then
        ComboT14.ListIndex = 6
    ElseIf DatBuff(43) = 7 Then
        ComboT14.ListIndex = 7
    ElseIf DatBuff(43) = 8 Then
        ComboT14.ListIndex = 8
    ElseIf DatBuff(43) = 9 Then
        ComboT14.ListIndex = 9
    ElseIf DatBuff(43) = 10 Then
        ComboT14.ListIndex = 10
    ElseIf DatBuff(43) = 11 Then
        ComboT14.ListIndex = 11
    ElseIf DatBuff(43) = 12 Then
        ComboT14.ListIndex = 12
    Else
        ComboT14.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(44) = 0 Then
        ComboT15.ListIndex = 0
    ElseIf DatBuff(44) = 1 Then
        ComboT15.ListIndex = 1
    ElseIf DatBuff(44) = 2 Then
        ComboT15.ListIndex = 2
    ElseIf DatBuff(44) = 3 Then
        ComboT15.ListIndex = 3
    ElseIf DatBuff(44) = 4 Then
        ComboT15.ListIndex = 4
    ElseIf DatBuff(44) = 5 Then
        ComboT15.ListIndex = 5
    ElseIf DatBuff(44) = 6 Then
        ComboT15.ListIndex = 6
    ElseIf DatBuff(44) = 7 Then
        ComboT15.ListIndex = 7
    ElseIf DatBuff(44) = 8 Then
        ComboT15.ListIndex = 8
    ElseIf DatBuff(44) = 9 Then
        ComboT15.ListIndex = 9
    ElseIf DatBuff(44) = 10 Then
        ComboT15.ListIndex = 10
    ElseIf DatBuff(44) = 11 Then
        ComboT15.ListIndex = 11
    ElseIf DatBuff(44) = 12 Then
        ComboT15.ListIndex = 12
    Else
        ComboT15.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(45) = 0 Then
        ComboT16.ListIndex = 0
    ElseIf DatBuff(45) = 1 Then
        ComboT16.ListIndex = 1
    ElseIf DatBuff(45) = 2 Then
        ComboT16.ListIndex = 2
    ElseIf DatBuff(45) = 3 Then
        ComboT16.ListIndex = 3
    ElseIf DatBuff(45) = 4 Then
        ComboT16.ListIndex = 4
    ElseIf DatBuff(45) = 5 Then
        ComboT16.ListIndex = 5
    ElseIf DatBuff(45) = 6 Then
        ComboT16.ListIndex = 6
    ElseIf DatBuff(45) = 7 Then
        ComboT16.ListIndex = 7
    ElseIf DatBuff(45) = 8 Then
        ComboT16.ListIndex = 8
    ElseIf DatBuff(45) = 9 Then
        ComboT16.ListIndex = 9
    ElseIf DatBuff(45) = 10 Then
        ComboT16.ListIndex = 10
    ElseIf DatBuff(45) = 11 Then
        ComboT16.ListIndex = 11
    ElseIf DatBuff(45) = 12 Then
        ComboT16.ListIndex = 12
    Else
        ComboT16.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(46) = 0 Then
        ComboT17.ListIndex = 0
    ElseIf DatBuff(46) = 1 Then
        ComboT17.ListIndex = 1
    ElseIf DatBuff(46) = 2 Then
        ComboT17.ListIndex = 2
    ElseIf DatBuff(46) = 3 Then
        ComboT17.ListIndex = 3
    ElseIf DatBuff(46) = 4 Then
        ComboT17.ListIndex = 4
    ElseIf DatBuff(46) = 5 Then
        ComboT17.ListIndex = 5
    ElseIf DatBuff(46) = 6 Then
        ComboT17.ListIndex = 6
    ElseIf DatBuff(46) = 7 Then
        ComboT17.ListIndex = 7
    ElseIf DatBuff(46) = 8 Then
        ComboT17.ListIndex = 8
    ElseIf DatBuff(46) = 9 Then
        ComboT17.ListIndex = 9
    ElseIf DatBuff(46) = 10 Then
        ComboT17.ListIndex = 10
    ElseIf DatBuff(46) = 11 Then
        ComboT17.ListIndex = 11
    ElseIf DatBuff(46) = 12 Then
        ComboT17.ListIndex = 12
    Else
        ComboT17.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(47) = 0 Then
        ComboT18.ListIndex = 0
    ElseIf DatBuff(47) = 1 Then
        ComboT18.ListIndex = 1
    ElseIf DatBuff(47) = 2 Then
        ComboT18.ListIndex = 2
    ElseIf DatBuff(47) = 3 Then
        ComboT18.ListIndex = 3
    ElseIf DatBuff(47) = 4 Then
        ComboT18.ListIndex = 4
    ElseIf DatBuff(47) = 5 Then
        ComboT18.ListIndex = 5
    ElseIf DatBuff(47) = 6 Then
        ComboT18.ListIndex = 6
    ElseIf DatBuff(47) = 7 Then
        ComboT18.ListIndex = 7
    ElseIf DatBuff(47) = 8 Then
        ComboT18.ListIndex = 8
    ElseIf DatBuff(47) = 9 Then
        ComboT18.ListIndex = 9
    ElseIf DatBuff(47) = 10 Then
        ComboT18.ListIndex = 10
    ElseIf DatBuff(47) = 11 Then
        ComboT18.ListIndex = 11
    ElseIf DatBuff(47) = 12 Then
        ComboT18.ListIndex = 12
    Else
        ComboT18.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(48) = 0 Then
        ComboT19.ListIndex = 0
    ElseIf DatBuff(48) = 1 Then
        ComboT19.ListIndex = 1
    ElseIf DatBuff(48) = 2 Then
        ComboT19.ListIndex = 2
    ElseIf DatBuff(48) = 3 Then
        ComboT19.ListIndex = 3
    ElseIf DatBuff(48) = 4 Then
        ComboT19.ListIndex = 4
    ElseIf DatBuff(48) = 5 Then
        ComboT19.ListIndex = 5
    ElseIf DatBuff(48) = 6 Then
        ComboT19.ListIndex = 6
    ElseIf DatBuff(48) = 7 Then
        ComboT19.ListIndex = 7
    ElseIf DatBuff(48) = 8 Then
        ComboT19.ListIndex = 8
    ElseIf DatBuff(48) = 9 Then
        ComboT19.ListIndex = 9
    ElseIf DatBuff(48) = 10 Then
        ComboT19.ListIndex = 10
    ElseIf DatBuff(48) = 11 Then
        ComboT19.ListIndex = 11
    ElseIf DatBuff(48) = 12 Then
        ComboT19.ListIndex = 12
    Else
        ComboT19.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    If DatBuff(49) = 0 Then
        ComboT20.ListIndex = 0
    ElseIf DatBuff(49) = 1 Then
        ComboT20.ListIndex = 1
    ElseIf DatBuff(49) = 2 Then
        ComboT20.ListIndex = 2
    ElseIf DatBuff(49) = 3 Then
        ComboT20.ListIndex = 3
    ElseIf DatBuff(49) = 4 Then
        ComboT20.ListIndex = 4
    ElseIf DatBuff(49) = 5 Then
        ComboT20.ListIndex = 5
    ElseIf DatBuff(49) = 6 Then
        ComboT20.ListIndex = 6
    ElseIf DatBuff(49) = 7 Then
        ComboT20.ListIndex = 7
    ElseIf DatBuff(49) = 8 Then
        ComboT20.ListIndex = 8
    ElseIf DatBuff(49) = 9 Then
        ComboT20.ListIndex = 9
    ElseIf DatBuff(49) = 10 Then
        ComboT20.ListIndex = 10
    ElseIf DatBuff(49) = 11 Then
        ComboT20.ListIndex = 11
    ElseIf DatBuff(49) = 12 Then
        ComboT20.ListIndex = 12
    Else
        ComboT20.ListIndex = 0
    End If
    '//////////////////////////////////////////////////////////////////////////
    
    
  
End Sub



Private Sub OptNoInput8_Click()
    InputP1.Visible = True
    InputP2.Visible = True
    InputP3.Visible = True
    InputP4.Visible = True
    InputP5.Visible = True
    InputP6.Visible = True
    InputP7.Visible = True
    InputP8.Visible = True
    InputP9.Visible = False
    InputP10.Visible = False
    InputP11.Visible = False
    InputP12.Visible = False
    InputP13.Visible = False
    InputP14.Visible = False
    InputP15.Visible = False
    InputP16.Visible = False
    InputP17.Visible = False
    InputP18.Visible = False
    InputP19.Visible = False
    InputP20.Visible = False
    
    FrameRGA1.Visible = True
    FrameRGA2.Visible = True
    FrameRGA3.Visible = True
    FrameRGA4.Visible = True
    FrameRGA5.Visible = True
    FrameRGA6.Visible = True
    FrameRGA7.Visible = True
    FrameRGA8.Visible = True
    FrameRGA9.Visible = False
    FrameRGA10.Visible = False
    FrameRGA11.Visible = False
    FrameRGA12.Visible = False
    FrameRGA13.Visible = False
    FrameRGA14.Visible = False
    FrameRGA15.Visible = False
    FrameRGA16.Visible = False
    FrameRGA17.Visible = False
    FrameRGA18.Visible = False
    FrameRGA19.Visible = False
    FrameRGA20.Visible = False
    
    FrameT1.Visible = True
    FrameT2.Visible = True
    FrameT3.Visible = True
    FrameT4.Visible = True
    FrameT5.Visible = True
    FrameT6.Visible = True
    FrameT7.Visible = True
    FrameT8.Visible = True
    FrameT9.Visible = False
    FrameT10.Visible = False
    FrameT11.Visible = False
    FrameT12.Visible = False
    FrameT13.Visible = False
    FrameT14.Visible = False
    FrameT15.Visible = False
    FrameT16.Visible = False
    FrameT17.Visible = False
    FrameT18.Visible = False
    FrameT19.Visible = False
    FrameT20.Visible = False
    
    SSTab1.Caption = "INPUT 1-8"

End Sub
Private Sub OptNoInput10_Click()
    InputP1.Visible = True
    InputP2.Visible = True
    InputP3.Visible = True
    InputP4.Visible = True
    InputP5.Visible = True
    InputP6.Visible = True
    InputP7.Visible = True
    InputP8.Visible = True
    InputP9.Visible = True
    InputP10.Visible = True
    InputP11.Visible = False
    InputP12.Visible = False
    InputP13.Visible = False
    InputP14.Visible = False
    InputP15.Visible = False
    InputP16.Visible = False
    InputP17.Visible = False
    InputP18.Visible = False
    InputP19.Visible = False
    InputP20.Visible = False
    
    FrameRGA1.Visible = True
    FrameRGA2.Visible = True
    FrameRGA3.Visible = True
    FrameRGA4.Visible = True
    FrameRGA5.Visible = True
    FrameRGA6.Visible = True
    FrameRGA7.Visible = True
    FrameRGA8.Visible = True
    FrameRGA9.Visible = True
    FrameRGA10.Visible = True
    FrameRGA11.Visible = False
    FrameRGA12.Visible = False
    FrameRGA13.Visible = False
    FrameRGA14.Visible = False
    FrameRGA15.Visible = False
    FrameRGA16.Visible = False
    FrameRGA17.Visible = False
    FrameRGA18.Visible = False
    FrameRGA19.Visible = False
    FrameRGA20.Visible = False
    
    FrameT1.Visible = True
    FrameT2.Visible = True
    FrameT3.Visible = True
    FrameT4.Visible = True
    FrameT5.Visible = True
    FrameT6.Visible = True
    FrameT7.Visible = True
    FrameT8.Visible = True
    FrameT9.Visible = True
    FrameT10.Visible = True
    FrameT11.Visible = False
    FrameT12.Visible = False
    FrameT13.Visible = False
    FrameT14.Visible = False
    FrameT15.Visible = False
    FrameT16.Visible = False
    FrameT17.Visible = False
    FrameT18.Visible = False
    FrameT19.Visible = False
    FrameT20.Visible = False
    
    SSTab1.Caption = "INPUT 1-10"

End Sub
Private Sub OptNoInput16_Click()
    InputP1.Visible = True
    InputP2.Visible = True
    InputP3.Visible = True
    InputP4.Visible = True
    InputP5.Visible = True
    InputP6.Visible = True
    InputP7.Visible = True
    InputP8.Visible = True
    InputP9.Visible = True
    InputP10.Visible = True
    InputP11.Visible = True
    InputP12.Visible = True
    InputP13.Visible = True
    InputP14.Visible = True
    InputP15.Visible = True
    InputP16.Visible = True
    InputP17.Visible = False
    InputP18.Visible = False
    InputP19.Visible = False
    InputP20.Visible = False
    
    FrameRGA1.Visible = True
    FrameRGA2.Visible = True
    FrameRGA3.Visible = True
    FrameRGA4.Visible = True
    FrameRGA5.Visible = True
    FrameRGA6.Visible = True
    FrameRGA7.Visible = True
    FrameRGA8.Visible = True
    FrameRGA9.Visible = True
    FrameRGA10.Visible = True
    FrameRGA11.Visible = True
    FrameRGA12.Visible = True
    FrameRGA13.Visible = True
    FrameRGA14.Visible = True
    FrameRGA15.Visible = True
    FrameRGA16.Visible = True
    FrameRGA17.Visible = False
    FrameRGA18.Visible = False
    FrameRGA19.Visible = False
    FrameRGA20.Visible = False
    
    FrameT1.Visible = True
    FrameT2.Visible = True
    FrameT3.Visible = True
    FrameT4.Visible = True
    FrameT5.Visible = True
    FrameT6.Visible = True
    FrameT7.Visible = True
    FrameT8.Visible = True
    FrameT9.Visible = True
    FrameT10.Visible = True
    FrameT11.Visible = True
    FrameT12.Visible = True
    FrameT13.Visible = True
    FrameT14.Visible = True
    FrameT15.Visible = True
    FrameT16.Visible = True
    FrameT17.Visible = False
    FrameT18.Visible = False
    FrameT19.Visible = False
    FrameT20.Visible = False
    
    SSTab1.Caption = "INPUT 1-16"
    
End Sub
Private Sub OptNoInput20_Click()
    InputP1.Visible = True
    InputP2.Visible = True
    InputP3.Visible = True
    InputP4.Visible = True
    InputP5.Visible = True
    InputP6.Visible = True
    InputP7.Visible = True
    InputP8.Visible = True
    InputP9.Visible = True
    InputP10.Visible = True
    InputP11.Visible = True
    InputP12.Visible = True
    InputP13.Visible = True
    InputP14.Visible = True
    InputP15.Visible = True
    InputP16.Visible = True
    InputP17.Visible = True
    InputP18.Visible = True
    InputP19.Visible = True
    InputP20.Visible = True
    
    FrameRGA1.Visible = True
    FrameRGA2.Visible = True
    FrameRGA3.Visible = True
    FrameRGA4.Visible = True
    FrameRGA5.Visible = True
    FrameRGA6.Visible = True
    FrameRGA7.Visible = True
    FrameRGA8.Visible = True
    FrameRGA9.Visible = True
    FrameRGA10.Visible = True
    FrameRGA11.Visible = True
    FrameRGA12.Visible = True
    FrameRGA13.Visible = True
    FrameRGA14.Visible = True
    FrameRGA15.Visible = True
    FrameRGA16.Visible = True
    FrameRGA17.Visible = True
    FrameRGA18.Visible = True
    FrameRGA19.Visible = True
    FrameRGA20.Visible = True
    
    FrameT1.Visible = True
    FrameT2.Visible = True
    FrameT3.Visible = True
    FrameT4.Visible = True
    FrameT5.Visible = True
    FrameT6.Visible = True
    FrameT7.Visible = True
    FrameT8.Visible = True
    FrameT9.Visible = True
    FrameT10.Visible = True
    FrameT11.Visible = True
    FrameT12.Visible = True
    FrameT13.Visible = True
    FrameT14.Visible = True
    FrameT15.Visible = True
    FrameT16.Visible = True
    FrameT17.Visible = True
    FrameT18.Visible = True
    FrameT19.Visible = True
    FrameT20.Visible = True
    
    SSTab1.Caption = "INPUT 1-20"
    
End Sub

Private Sub Command8_Click()
    InputP1.Fault_Indicator = "INDICATOR"
    InputP2.Fault_Indicator = "INDICATOR"
    InputP3.Fault_Indicator = "INDICATOR"
    InputP4.Fault_Indicator = "INDICATOR"
    InputP5.Fault_Indicator = "INDICATOR"
    InputP6.Fault_Indicator = "INDICATOR"
    InputP7.Fault_Indicator = "INDICATOR"
    InputP8.Fault_Indicator = "INDICATOR"
    InputP9.Fault_Indicator = "INDICATOR"
    InputP10.Fault_Indicator = "INDICATOR"
    InputP11.Fault_Indicator = "INDICATOR"
    InputP12.Fault_Indicator = "INDICATOR"
    InputP13.Fault_Indicator = "INDICATOR"
    InputP14.Fault_Indicator = "INDICATOR"
    InputP15.Fault_Indicator = "INDICATOR"
    InputP16.Fault_Indicator = "INDICATOR"
    InputP17.Fault_Indicator = "INDICATOR"
    InputP18.Fault_Indicator = "INDICATOR"
    InputP19.Fault_Indicator = "INDICATOR"
    InputP20.Fault_Indicator = "INDICATOR"

End Sub
Private Sub Command9_Click()
    InputP1.Fault_Indicator = "FAULT"
    InputP2.Fault_Indicator = "FAULT"
    InputP3.Fault_Indicator = "FAULT"
    InputP4.Fault_Indicator = "FAULT"
    InputP5.Fault_Indicator = "FAULT"
    InputP6.Fault_Indicator = "FAULT"
    InputP7.Fault_Indicator = "FAULT"
    InputP8.Fault_Indicator = "FAULT"
    InputP9.Fault_Indicator = "FAULT"
    InputP10.Fault_Indicator = "FAULT"
    InputP11.Fault_Indicator = "FAULT"
    InputP12.Fault_Indicator = "FAULT"
    InputP13.Fault_Indicator = "FAULT"
    InputP14.Fault_Indicator = "FAULT"
    InputP15.Fault_Indicator = "FAULT"
    InputP16.Fault_Indicator = "FAULT"
    InputP17.Fault_Indicator = "FAULT"
    InputP18.Fault_Indicator = "FAULT"
    InputP19.Fault_Indicator = "FAULT"
    InputP20.Fault_Indicator = "FAULT"
    
End Sub
Private Sub Command2_Click()
    InputP1.InputType = "NO"
    InputP2.InputType = "NO"
    InputP3.InputType = "NO"
    InputP4.InputType = "NO"
    InputP5.InputType = "NO"
    InputP6.InputType = "NO"
    InputP7.InputType = "NO"
    InputP8.InputType = "NO"
    InputP9.InputType = "NO"
    InputP10.InputType = "NO"
    InputP11.InputType = "NO"
    InputP12.InputType = "NO"
    InputP13.InputType = "NO"
    InputP14.InputType = "NO"
    InputP15.InputType = "NO"
    InputP16.InputType = "NO"
    InputP17.InputType = "NO"
    InputP18.InputType = "NO"
    InputP19.InputType = "NO"
    InputP20.InputType = "NO"
    
End Sub
Private Sub Command3_Click()
    InputP1.InputType = "NC"
    InputP2.InputType = "NC"
    InputP3.InputType = "NC"
    InputP4.InputType = "NC"
    InputP5.InputType = "NC"
    InputP6.InputType = "NC"
    InputP7.InputType = "NC"
    InputP8.InputType = "NC"
    InputP9.InputType = "NC"
    InputP10.InputType = "NC"
    InputP11.InputType = "NC"
    InputP12.InputType = "NC"
    InputP13.InputType = "NC"
    InputP14.InputType = "NC"
    InputP15.InputType = "NC"
    InputP16.InputType = "NC"
    InputP17.InputType = "NC"
    InputP18.InputType = "NC"
    InputP19.InputType = "NC"
    InputP20.InputType = "NC"
End Sub
Private Sub Command4_Click()
    InputP1.FaultType = "MANUAL"
    InputP2.FaultType = "MANUAL"
    InputP3.FaultType = "MANUAL"
    InputP4.FaultType = "MANUAL"
    InputP5.FaultType = "MANUAL"
    InputP6.FaultType = "MANUAL"
    InputP7.FaultType = "MANUAL"
    InputP8.FaultType = "MANUAL"
    InputP9.FaultType = "MANUAL"
    InputP10.FaultType = "MANUAL"
    InputP11.FaultType = "MANUAL"
    InputP12.FaultType = "MANUAL"
    InputP13.FaultType = "MANUAL"
    InputP14.FaultType = "MANUAL"
    InputP15.FaultType = "MANUAL"
    InputP16.FaultType = "MANUAL"
    InputP17.FaultType = "MANUAL"
    InputP18.FaultType = "MANUAL"
    InputP19.FaultType = "MANUAL"
    InputP20.FaultType = "MANUAL"
End Sub
Private Sub Command5_Click()
    InputP1.FaultType = "AUTO"
    InputP2.FaultType = "AUTO"
    InputP3.FaultType = "AUTO"
    InputP4.FaultType = "AUTO"
    InputP5.FaultType = "AUTO"
    InputP6.FaultType = "AUTO"
    InputP7.FaultType = "AUTO"
    InputP8.FaultType = "AUTO"
    InputP9.FaultType = "AUTO"
    InputP10.FaultType = "AUTO"
    InputP11.FaultType = "AUTO"
    InputP12.FaultType = "AUTO"
    InputP13.FaultType = "AUTO"
    InputP14.FaultType = "AUTO"
    InputP15.FaultType = "AUTO"
    InputP16.FaultType = "AUTO"
    InputP17.FaultType = "AUTO"
    InputP18.FaultType = "AUTO"
    InputP19.FaultType = "AUTO"
    InputP20.FaultType = "AUTO"
End Sub
Private Sub Command6_Click()
    InputP1.OutputType = "BUZZER"
    InputP2.OutputType = "BUZZER"
    InputP3.OutputType = "BUZZER"
    InputP4.OutputType = "BUZZER"
    InputP5.OutputType = "BUZZER"
    InputP6.OutputType = "BUZZER"
    InputP7.OutputType = "BUZZER"
    InputP8.OutputType = "BUZZER"
    InputP9.OutputType = "BUZZER"
    InputP10.OutputType = "BUZZER"
    InputP11.OutputType = "BUZZER"
    InputP12.OutputType = "BUZZER"
    InputP13.OutputType = "BUZZER"
    InputP14.OutputType = "BUZZER"
    InputP15.OutputType = "BUZZER"
    InputP16.OutputType = "BUZZER"
    InputP17.OutputType = "BUZZER"
    InputP18.OutputType = "BUZZER"
    InputP19.OutputType = "BUZZER"
    InputP20.OutputType = "BUZZER"
End Sub
Private Sub Command7_Click()
    InputP1.OutputType = "BELL"
    InputP2.OutputType = "BELL"
    InputP3.OutputType = "BELL"
    InputP4.OutputType = "BELL"
    InputP5.OutputType = "BELL"
    InputP6.OutputType = "BELL"
    InputP7.OutputType = "BELL"
    InputP8.OutputType = "BELL"
    InputP9.OutputType = "BELL"
    InputP10.OutputType = "BELL"
    InputP11.OutputType = "BELL"
    InputP12.OutputType = "BELL"
    InputP13.OutputType = "BELL"
    InputP14.OutputType = "BELL"
    InputP15.OutputType = "BELL"
    InputP16.OutputType = "BELL"
    InputP17.OutputType = "BELL"
    InputP18.OutputType = "BELL"
    InputP19.OutputType = "BELL"
    InputP20.OutputType = "BELL"
End Sub

Private Sub Rall_Click()
    R1.Value = True
    R2.Value = True
    R3.Value = True
    R4.Value = True
    R5.Value = True
    R6.Value = True
    R7.Value = True
    R8.Value = True
    R9.Value = True
    R10.Value = True
    R11.Value = True
    R12.Value = True
    R13.Value = True
    R14.Value = True
    R15.Value = True
    R16.Value = True
    R17.Value = True
    R18.Value = True
    R19.Value = True
    R20.Value = True
End Sub
Private Sub Gall_Click()
    G1.Value = True
    G2.Value = True
    G3.Value = True
    G4.Value = True
    G5.Value = True
    G6.Value = True
    G7.Value = True
    G8.Value = True
    G9.Value = True
    G10.Value = True
    G11.Value = True
    G12.Value = True
    G13.Value = True
    G14.Value = True
    G15.Value = True
    G16.Value = True
    G17.Value = True
    G18.Value = True
    G19.Value = True
    G20.Value = True
End Sub
Private Sub Option3_Click()
    A1.Value = True
    A2.Value = True
    A3.Value = True
    A4.Value = True
    A5.Value = True
    A6.Value = True
    A7.Value = True
    A8.Value = True
    A9.Value = True
    A10.Value = True
    A11.Value = True
    A12.Value = True
    A13.Value = True
    A14.Value = True
    A15.Value = True
    A16.Value = True
    A17.Value = True
    A18.Value = True
    A19.Value = True
    A20.Value = True
End Sub
Private Sub Command10_Click()
    InputP1.OutputBoth = "BOTH"
    InputP2.OutputBoth = "BOTH"
    InputP3.OutputBoth = "BOTH"
    InputP4.OutputBoth = "BOTH"
    InputP5.OutputBoth = "BOTH"
    InputP6.OutputBoth = "BOTH"
    InputP7.OutputBoth = "BOTH"
    InputP8.OutputBoth = "BOTH"
    InputP9.OutputBoth = "BOTH"
    InputP10.OutputBoth = "BOTH"
    InputP11.OutputBoth = "BOTH"
    InputP12.OutputBoth = "BOTH"
    InputP13.OutputBoth = "BOTH"
    InputP14.OutputBoth = "BOTH"
    InputP15.OutputBoth = "BOTH"
    InputP16.OutputBoth = "BOTH"
    InputP17.OutputBoth = "BOTH"
    InputP18.OutputBoth = "BOTH"
    InputP19.OutputBoth = "BOTH"
    InputP20.OutputBoth = "BOTH"
End Sub
Private Sub Command11_Click()
    InputP1.OutputBoth = "SINGLE"
    InputP2.OutputBoth = "SINGLE"
    InputP3.OutputBoth = "SINGLE"
    InputP4.OutputBoth = "SINGLE"
    InputP5.OutputBoth = "SINGLE"
    InputP6.OutputBoth = "SINGLE"
    InputP7.OutputBoth = "SINGLE"
    InputP8.OutputBoth = "SINGLE"
    InputP9.OutputBoth = "SINGLE"
    InputP10.OutputBoth = "SINGLE"
    InputP11.OutputBoth = "SINGLE"
    InputP12.OutputBoth = "SINGLE"
    InputP13.OutputBoth = "SINGLE"
    InputP14.OutputBoth = "SINGLE"
    InputP15.OutputBoth = "SINGLE"
    InputP16.OutputBoth = "SINGLE"
    InputP17.OutputBoth = "SINGLE"
    InputP18.OutputBoth = "SINGLE"
    InputP19.OutputBoth = "SINGLE"
    InputP20.OutputBoth = "SINGLE"
End Sub



Private Sub Timer1_Timer()

    If CommState = "DOWNLOAD" Then
        Timeout = Timeout + 1
        If Timeout = 3 Then
            MsgBox "DOWNLOAD FAIL.", vbCritical, "ESPAN-04"
            CommState = "IDLE"
            Timeout = 0
            Timer1.Enabled = False
        End If
        
    End If
    
    If CommState = "UPLOAD" Then
        Timeout = Timeout + 1
        If Timeout = 4 Then
            MsgBox "UPLOAD FAIL.", vbCritical, "ESPAN-04"
            CommState = "IDLE"
            Timeout = 0
            Timer1.Enabled = False
        End If
    End If
    
    
End Sub



Private Sub VScroll2_Change()
    InputP1.Top = 120 - VScroll2.Value
    InputP2.Top = 1080 - VScroll2.Value
    InputP3.Top = 2040 - VScroll2.Value
    InputP4.Top = 3000 - VScroll2.Value
    InputP5.Top = 3960 - VScroll2.Value
    InputP6.Top = 4920 - VScroll2.Value
    InputP7.Top = 5880 - VScroll2.Value
    InputP8.Top = 6840 - VScroll2.Value
    InputP9.Top = 7800 - VScroll2.Value
    InputP10.Top = 8760 - VScroll2.Value
    
    InputP11.Top = 120 - VScroll2.Value
    InputP12.Top = 1080 - VScroll2.Value
    InputP13.Top = 2040 - VScroll2.Value
    InputP14.Top = 3000 - VScroll2.Value
    InputP15.Top = 3960 - VScroll2.Value
    InputP16.Top = 4920 - VScroll2.Value
    InputP17.Top = 5880 - VScroll2.Value
    InputP18.Top = 6840 - VScroll2.Value
    InputP19.Top = 7800 - VScroll2.Value
    InputP20.Top = 8760 - VScroll2.Value
    
    FrameRGA1.Top = 120 - VScroll2.Value
    FrameRGA2.Top = 1080 - VScroll2.Value
    FrameRGA3.Top = 2040 - VScroll2.Value
    FrameRGA4.Top = 3000 - VScroll2.Value
    FrameRGA5.Top = 3960 - VScroll2.Value
    FrameRGA6.Top = 4920 - VScroll2.Value
    FrameRGA7.Top = 5880 - VScroll2.Value
    FrameRGA8.Top = 6840 - VScroll2.Value
    FrameRGA9.Top = 7800 - VScroll2.Value
    FrameRGA10.Top = 8760 - VScroll2.Value
    
    FrameRGA11.Top = 120 - VScroll2.Value
    FrameRGA12.Top = 1080 - VScroll2.Value
    FrameRGA13.Top = 2040 - VScroll2.Value
    FrameRGA14.Top = 3000 - VScroll2.Value
    FrameRGA15.Top = 3960 - VScroll2.Value
    FrameRGA16.Top = 4920 - VScroll2.Value
    FrameRGA17.Top = 5880 - VScroll2.Value
    FrameRGA18.Top = 6840 - VScroll2.Value
    FrameRGA19.Top = 7800 - VScroll2.Value
    FrameRGA20.Top = 8760 - VScroll2.Value
    
    FrameT1.Top = 120 - VScroll2.Value
    FrameT2.Top = 1080 - VScroll2.Value
    FrameT3.Top = 2040 - VScroll2.Value
    FrameT4.Top = 3000 - VScroll2.Value
    FrameT5.Top = 3960 - VScroll2.Value
    FrameT6.Top = 4920 - VScroll2.Value
    FrameT7.Top = 5880 - VScroll2.Value
    FrameT8.Top = 6840 - VScroll2.Value
    FrameT9.Top = 7800 - VScroll2.Value
    FrameT10.Top = 8760 - VScroll2.Value
    
    FrameT11.Top = 120 - VScroll2.Value
    FrameT12.Top = 1080 - VScroll2.Value
    FrameT13.Top = 2040 - VScroll2.Value
    FrameT14.Top = 3000 - VScroll2.Value
    FrameT15.Top = 3960 - VScroll2.Value
    FrameT16.Top = 4920 - VScroll2.Value
    FrameT17.Top = 5880 - VScroll2.Value
    FrameT18.Top = 6840 - VScroll2.Value
    FrameT19.Top = 7800 - VScroll2.Value
    FrameT20.Top = 8760 - VScroll2.Value
    
End Sub
