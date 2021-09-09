VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form8Point 
   Caption         =   "Form1"
   ClientHeight    =   5175
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   5175
   ScaleWidth      =   6030
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   3975
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.TextBox Text1 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   16
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
      Begin VB.TextBox Text2 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   16
         TabIndex        =   7
         Top             =   720
         Width           =   2295
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   16
         TabIndex        =   6
         Top             =   1200
         Width           =   2295
      End
      Begin VB.TextBox Text4 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   16
         TabIndex        =   5
         Top             =   1680
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   16
         TabIndex        =   4
         Top             =   2160
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   16
         TabIndex        =   3
         Top             =   2640
         Width           =   2295
      End
      Begin VB.TextBox Text7 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   16
         TabIndex        =   2
         Top             =   3120
         Width           =   2295
      End
      Begin VB.TextBox Text8 
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         MaxLength       =   16
         TabIndex        =   1
         Top             =   3600
         Width           =   2295
      End
      Begin VB.Label Label1 
         Caption         =   "INPUT1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "INPUT2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "INPUT3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "INPUT4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "INPUT5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2160
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "INPUT6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   2640
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "INPUT7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   3120
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "INPUT8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000015&
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Width           =   735
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   720
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu Printmenu 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "Form8Point"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
