VERSION 5.00
Begin VB.Form FormPrint20preview 
   BackColor       =   &H00FFFFFF&
   Caption         =   "ESPAN-04"
   ClientHeight    =   13125
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11355
   Icon            =   "FormPrint20preview.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   13125
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   3360
      Top             =   12600
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   10335
      Left            =   240
      TabIndex        =   0
      Top             =   2040
      Width           =   10335
      Begin VB.Line Line42 
         BorderColor     =   &H80000005&
         BorderWidth     =   6
         X1              =   7440
         X2              =   8160
         Y1              =   600
         Y2              =   840
      End
      Begin VB.Line Line41 
         BorderColor     =   &H80000005&
         BorderWidth     =   6
         X1              =   4920
         X2              =   6360
         Y1              =   1080
         Y2              =   600
      End
      Begin VB.Line Line40 
         BorderColor     =   &H80000005&
         BorderWidth     =   25
         X1              =   7560
         X2              =   9120
         Y1              =   480
         Y2              =   960
      End
      Begin VB.Line Line39 
         BorderColor     =   &H80000005&
         BorderWidth     =   25
         X1              =   4920
         X2              =   6240
         Y1              =   840
         Y2              =   480
      End
      Begin VB.Line Line38 
         BorderColor     =   &H80000005&
         BorderWidth     =   5
         X1              =   4320
         X2              =   4680
         Y1              =   600
         Y2              =   720
      End
      Begin VB.Line Line37 
         BorderColor     =   &H80000005&
         BorderWidth     =   7
         X1              =   3360
         X2              =   4800
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line Line36 
         BorderColor     =   &H80000005&
         BorderWidth     =   25
         X1              =   3480
         X2              =   4920
         Y1              =   480
         Y2              =   960
      End
      Begin VB.Line Line35 
         BorderColor     =   &H80000005&
         BorderWidth     =   7
         X1              =   480
         X2              =   2280
         Y1              =   1200
         Y2              =   600
      End
      Begin VB.Line Line34 
         BorderColor     =   &H80000005&
         BorderWidth     =   30
         X1              =   360
         X2              =   2040
         Y1              =   960
         Y2              =   480
      End
      Begin VB.Line Line33 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   7320
         X2              =   8160
         Y1              =   5640
         Y2              =   5880
      End
      Begin VB.Line Line32 
         BorderColor     =   &H80000005&
         BorderWidth     =   25
         X1              =   9120
         X2              =   7560
         Y1              =   6000
         Y2              =   5520
      End
      Begin VB.Line Line31 
         BorderColor     =   &H80000005&
         BorderWidth     =   6
         X1              =   4800
         X2              =   6360
         Y1              =   6120
         Y2              =   5640
      End
      Begin VB.Line Line30 
         BorderColor     =   &H80000005&
         BorderWidth     =   20
         X1              =   4920
         X2              =   6240
         Y1              =   5880
         Y2              =   5520
      End
      Begin VB.Line Line29 
         BorderColor     =   &H80000005&
         BorderWidth     =   6
         X1              =   3360
         X2              =   4920
         Y1              =   5640
         Y2              =   6120
      End
      Begin VB.Line Line28 
         BorderColor     =   &H80000005&
         BorderWidth     =   25
         X1              =   3480
         X2              =   4800
         Y1              =   5520
         Y2              =   5880
      End
      Begin VB.Line Line27 
         BorderColor     =   &H80000005&
         BorderWidth     =   25
         X1              =   960
         X2              =   2160
         Y1              =   5880
         Y2              =   5520
      End
      Begin VB.Line Line26 
         BorderColor     =   &H80000005&
         BorderWidth     =   10
         X1              =   720
         X2              =   2400
         Y1              =   6120
         Y2              =   5640
      End
      Begin VB.Shape Shape12 
         BackColor       =   &H80000008&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   6555
         Top             =   5400
         Width           =   735
      End
      Begin VB.Shape Shape11 
         BackColor       =   &H80000008&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   2520
         Top             =   5400
         Width           =   735
      End
      Begin VB.Line Line25 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   8880
         X2              =   7200
         Y1              =   6120
         Y2              =   5640
      End
      Begin VB.Line Line24 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   4920
         X2              =   6600
         Y1              =   6120
         Y2              =   5640
      End
      Begin VB.Line Line23 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   4800
         X2              =   3120
         Y1              =   6120
         Y2              =   5640
      End
      Begin VB.Line Line22 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   840
         X2              =   2640
         Y1              =   6120
         Y2              =   5640
      End
      Begin VB.Shape Shape10 
         BackColor       =   &H80000008&
         BackStyle       =   1  'Opaque
         Height          =   495
         Left            =   6480
         Top             =   360
         Width           =   855
      End
      Begin VB.Shape Shape9 
         BackColor       =   &H80000008&
         BackStyle       =   1  'Opaque
         Height          =   375
         Left            =   2400
         Top             =   360
         Width           =   855
      End
      Begin VB.Line Line21 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   9240
         X2              =   7200
         Y1              =   1200
         Y2              =   600
      End
      Begin VB.Line Line19 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   5040
         X2              =   6600
         Y1              =   1080
         Y2              =   600
      End
      Begin VB.Line Line18 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   3120
         X2              =   5040
         Y1              =   600
         Y2              =   1080
      End
      Begin VB.Line Line20 
         BorderColor     =   &H80000005&
         BorderWidth     =   3
         X1              =   480
         X2              =   2520
         Y1              =   1200
         Y2              =   600
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   2520
         Top             =   5400
         Width           =   735
      End
      Begin VB.Shape Shape8 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   420
         Left            =   5160
         Top             =   5685
         Width           =   3495
      End
      Begin VB.Shape Shape7 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   420
         Left            =   1080
         Top             =   5685
         Width           =   3495
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   420
         Left            =   5160
         Top             =   660
         Width           =   3495
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   420
         Left            =   1080
         Top             =   660
         Width           =   3495
      End
      Begin VB.Label Label11 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HHHHHHHHHHHHHHHH"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   80
         Top             =   1230
         Width           =   1620
      End
      Begin VB.Label Label13 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   79
         Top             =   1905
         Width           =   1620
      End
      Begin VB.Label Label14 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   78
         Top             =   2250
         Width           =   1620
      End
      Begin VB.Label Label15 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   77
         Top             =   2595
         Width           =   1620
      End
      Begin VB.Label Label17 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   76
         Top             =   3270
         Width           =   1620
      End
      Begin VB.Label Label18 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   75
         Top             =   3615
         Width           =   1620
      End
      Begin VB.Label Label8 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   74
         Top             =   3615
         Width           =   1620
      End
      Begin VB.Label Label7 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   73
         Top             =   3270
         Width           =   1620
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   72
         Top             =   2940
         Width           =   1620
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   71
         Top             =   2595
         Width           =   1620
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   70
         Top             =   2250
         Width           =   1620
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   69
         Top             =   1905
         Width           =   1620
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "AAAAAAAAAAAAAAAAAA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   68
         Top             =   1230
         Width           =   1620
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         X1              =   1095
         X2              =   4575
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         X1              =   1080
         X2              =   1080
         Y1              =   660
         Y2              =   4619
      End
      Begin VB.Label Label19 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   67
         Top             =   3960
         Width           =   1620
      End
      Begin VB.Label Label16 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   66
         Top             =   2940
         Width           =   1620
      End
      Begin VB.Label Label9 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   65
         Top             =   3960
         Width           =   1620
      End
      Begin VB.Label Label20 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   64
         Top             =   4305
         Width           =   1620
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   63
         Top             =   1545
         Width           =   1620
      End
      Begin VB.Label Label12 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   62
         Top             =   1575
         Width           =   1620
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         X1              =   4575
         X2              =   4575
         Y1              =   660
         Y2              =   4619
      End
      Begin VB.Line Line4 
         BorderColor     =   &H8000000C&
         X1              =   1095
         X2              =   4575
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   2475
         Top             =   360
         Width           =   740
      End
      Begin VB.Label Label10 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   61
         Top             =   4305
         Width           =   1620
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   6555
         Top             =   360
         Width           =   740
      End
      Begin VB.Line Line5 
         BorderColor     =   &H8000000C&
         X1              =   5175
         X2              =   8655
         Y1              =   660
         Y2              =   660
      End
      Begin VB.Line Line6 
         BorderColor     =   &H8000000C&
         X1              =   5175
         X2              =   5175
         Y1              =   660
         Y2              =   4619
      End
      Begin VB.Line Line7 
         BorderColor     =   &H8000000C&
         X1              =   5175
         X2              =   8655
         Y1              =   4620
         Y2              =   4620
      End
      Begin VB.Line Line8 
         BorderColor     =   &H8000000C&
         X1              =   8655
         X2              =   8655
         Y1              =   660
         Y2              =   4619
      End
      Begin VB.Label Label21 
         BackColor       =   &H00FFFFFF&
         Caption         =   "AAAAAAAAAAAAAAAAAA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   60
         Top             =   1230
         Width           =   1620
      End
      Begin VB.Label Label22 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   59
         Top             =   1575
         Width           =   1620
      End
      Begin VB.Label Label23 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   58
         Top             =   1905
         Width           =   1620
      End
      Begin VB.Label Label24 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   57
         Top             =   2250
         Width           =   1620
      End
      Begin VB.Label Label25 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   56
         Top             =   2595
         Width           =   1620
      End
      Begin VB.Label Label26 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   55
         Top             =   2940
         Width           =   1620
      End
      Begin VB.Label Label27 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   54
         Top             =   3270
         Width           =   1620
      End
      Begin VB.Label Label28 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   53
         Top             =   3615
         Width           =   1620
      End
      Begin VB.Label Label29 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   52
         Top             =   3960
         Width           =   1620
      End
      Begin VB.Label Label30 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   51
         Top             =   4305
         Width           =   1620
      End
      Begin VB.Label Label31 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HHHHHHHHHHHHHHHH"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   50
         Top             =   1230
         Width           =   1620
      End
      Begin VB.Label Label32 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   49
         Top             =   1575
         Width           =   1620
      End
      Begin VB.Label Label33 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   48
         Top             =   1905
         Width           =   1620
      End
      Begin VB.Label Label34 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   47
         Top             =   2250
         Width           =   1620
      End
      Begin VB.Label Label35 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   46
         Top             =   2595
         Width           =   1620
      End
      Begin VB.Label Label36 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   45
         Top             =   2940
         Width           =   1620
      End
      Begin VB.Label Label37 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   44
         Top             =   3270
         Width           =   1620
      End
      Begin VB.Label Label38 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   43
         Top             =   3615
         Width           =   1620
      End
      Begin VB.Label Label39 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   42
         Top             =   3960
         Width           =   1620
      End
      Begin VB.Label Label40 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   41
         Top             =   4305
         Width           =   1620
      End
      Begin VB.Line Line9 
         BorderColor     =   &H8000000C&
         X1              =   1080
         X2              =   4560
         Y1              =   5685
         Y2              =   5685
      End
      Begin VB.Line Line10 
         BorderColor     =   &H8000000C&
         X1              =   1080
         X2              =   1080
         Y1              =   5685
         Y2              =   9644
      End
      Begin VB.Line Line11 
         BorderColor     =   &H8000000C&
         X1              =   1080
         X2              =   4560
         Y1              =   9645
         Y2              =   9645
      End
      Begin VB.Line Line12 
         BorderColor     =   &H8000000C&
         X1              =   4575
         X2              =   4575
         Y1              =   5685
         Y2              =   9644
      End
      Begin VB.Label Label41 
         BackColor       =   &H00FFFFFF&
         Caption         =   "AAAAAAAAAAAAAAAAAA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   40
         Top             =   6255
         Width           =   1620
      End
      Begin VB.Label Label42 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   39
         Top             =   6600
         Width           =   1620
      End
      Begin VB.Label Label43 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   38
         Top             =   6930
         Width           =   1620
      End
      Begin VB.Label Label44 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   37
         Top             =   7275
         Width           =   1620
      End
      Begin VB.Label Label45 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   36
         Top             =   7620
         Width           =   1620
      End
      Begin VB.Label Label46 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   35
         Top             =   7965
         Width           =   1620
      End
      Begin VB.Label Label47 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   34
         Top             =   8295
         Width           =   1620
      End
      Begin VB.Label Label48 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   33
         Top             =   8640
         Width           =   1620
      End
      Begin VB.Label Label49 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   32
         Top             =   8985
         Width           =   1620
      End
      Begin VB.Label Label50 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1140
         TabIndex        =   31
         Top             =   9330
         Width           =   1620
      End
      Begin VB.Label Label51 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HHHHHHHHHHHHHHHH"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   30
         Top             =   6255
         Width           =   1620
      End
      Begin VB.Label Label52 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   29
         Top             =   6600
         Width           =   1620
      End
      Begin VB.Label Label53 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   28
         Top             =   6930
         Width           =   1620
      End
      Begin VB.Label Label54 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   27
         Top             =   7275
         Width           =   1620
      End
      Begin VB.Label Label55 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   26
         Top             =   7620
         Width           =   1620
      End
      Begin VB.Label Label56 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   25
         Top             =   7965
         Width           =   1620
      End
      Begin VB.Label Label57 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   24
         Top             =   8295
         Width           =   1620
      End
      Begin VB.Label Label58 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   23
         Top             =   8640
         Width           =   1620
      End
      Begin VB.Label Label59 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   22
         Top             =   8985
         Width           =   1620
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   6555
         Top             =   5385
         Width           =   740
      End
      Begin VB.Line Line13 
         BorderColor     =   &H8000000C&
         X1              =   5175
         X2              =   8655
         Y1              =   5685
         Y2              =   5685
      End
      Begin VB.Line Line14 
         BorderColor     =   &H8000000C&
         X1              =   5160
         X2              =   5160
         Y1              =   5685
         Y2              =   9645
      End
      Begin VB.Line Line15 
         BorderColor     =   &H8000000C&
         X1              =   5175
         X2              =   8655
         Y1              =   9645
         Y2              =   9645
      End
      Begin VB.Line Line16 
         BorderColor     =   &H8000000C&
         X1              =   8655
         X2              =   8655
         Y1              =   5685
         Y2              =   9644
      End
      Begin VB.Label Label60 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   2910
         TabIndex        =   21
         Top             =   9330
         Width           =   1620
      End
      Begin VB.Label Label61 
         BackColor       =   &H00FFFFFF&
         Caption         =   "AAAAAAAAAAAAAAAAAA"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   20
         Top             =   6255
         Width           =   1620
      End
      Begin VB.Label Label62 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   19
         Top             =   6600
         Width           =   1620
      End
      Begin VB.Label Label63 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   18
         Top             =   6930
         Width           =   1620
      End
      Begin VB.Label Label64 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   17
         Top             =   7275
         Width           =   1620
      End
      Begin VB.Label Label65 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   16
         Top             =   7620
         Width           =   1620
      End
      Begin VB.Label Label66 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   15
         Top             =   7965
         Width           =   1620
      End
      Begin VB.Label Label67 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   14
         Top             =   8295
         Width           =   1620
      End
      Begin VB.Label Label68 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   13
         Top             =   8640
         Width           =   1620
      End
      Begin VB.Label Label69 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   12
         Top             =   8985
         Width           =   1620
      End
      Begin VB.Label Label70 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5235
         TabIndex        =   11
         Top             =   9330
         Width           =   1620
      End
      Begin VB.Label Label71 
         BackColor       =   &H00FFFFFF&
         Caption         =   "HHHHHHHHHHHHHHHH"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   10
         Top             =   6255
         Width           =   1620
      End
      Begin VB.Label Label72 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   9
         Top             =   6600
         Width           =   1620
      End
      Begin VB.Label Label73 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   8
         Top             =   6930
         Width           =   1620
      End
      Begin VB.Label Label74 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   7
         Top             =   7275
         Width           =   1620
      End
      Begin VB.Label Label75 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   6
         Top             =   7620
         Width           =   1620
      End
      Begin VB.Label Label76 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   5
         Top             =   7965
         Width           =   1620
      End
      Begin VB.Label Label77 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   4
         Top             =   8295
         Width           =   1620
      End
      Begin VB.Label Label78 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   3
         Top             =   8640
         Width           =   1620
      End
      Begin VB.Label Label79 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   2
         Top             =   8985
         Width           =   1620
      End
      Begin VB.Label Label80 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   7005
         TabIndex        =   1
         Top             =   9330
         Width           =   1620
      End
   End
   Begin VB.Line Line17 
      BorderColor     =   &H8000000C&
      X1              =   0
      X2              =   0
      Y1              =   0
      Y2              =   3959
   End
   Begin VB.Menu Printmenu 
      Caption         =   "Print"
   End
End
Attribute VB_Name = "FormPrint20preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public fontsizetemp As Byte





Private Sub Form_Load()
    Frame1.Top = 120
    Frame1.Left = 120
    Me.Height = 11130
    Timer1.Enabled = True
    'Label1_1.Caption = Label1.Caption
    'Text1.Visible = False
End Sub

Private Sub Label1_Click()
    fontsizetemp = 1
    frmResizeFont.Show
    'Label1.FontSize = fontsizetemp
End Sub


Private Sub Label10_Click()
    fontsizetemp = 10
    frmResizeFont.Show
End Sub

Private Sub Label11_Click()
    fontsizetemp = 11
    frmResizeFont.Show
End Sub

Private Sub Label12_Click()
    fontsizetemp = 12
    frmResizeFont.Show
End Sub

Private Sub Label13_Click()
    fontsizetemp = 13
    frmResizeFont.Show
End Sub

Private Sub Label14_Click()
    fontsizetemp = 14
    frmResizeFont.Show
End Sub

Private Sub Label15_Click()
    fontsizetemp = 15
    frmResizeFont.Show
End Sub

Private Sub Label16_Click()
    fontsizetemp = 16
    frmResizeFont.Show
End Sub

Private Sub Label17_Click()
    fontsizetemp = 17
    frmResizeFont.Show
End Sub

Private Sub Label18_Click()
    fontsizetemp = 18
    frmResizeFont.Show
End Sub

Private Sub Label19_Click()
    fontsizetemp = 19
    frmResizeFont.Show
End Sub

Private Sub Label2_Click()
    fontsizetemp = 2
    frmResizeFont.Show
    
End Sub

Private Sub Label20_Click()
    fontsizetemp = 20
    frmResizeFont.Show
End Sub

Private Sub Label21_Click()
    fontsizetemp = 21
    frmResizeFont.Show
End Sub

Private Sub Label22_Click()
    fontsizetemp = 22
    frmResizeFont.Show
End Sub

Private Sub Label23_Click()
    fontsizetemp = 23
    frmResizeFont.Show
End Sub

Private Sub Label24_Click()
    fontsizetemp = 24
    frmResizeFont.Show
End Sub

Private Sub Label25_Click()
    fontsizetemp = 25
    frmResizeFont.Show
End Sub

Private Sub Label26_Click()
    fontsizetemp = 26
    frmResizeFont.Show
End Sub

Private Sub Label27_Click()
    fontsizetemp = 27
    frmResizeFont.Show
End Sub

Private Sub Label28_Click()
    fontsizetemp = 28
    frmResizeFont.Show
End Sub

Private Sub Label29_Click()
    fontsizetemp = 29
    frmResizeFont.Show
End Sub

Private Sub Label3_Click()
    fontsizetemp = 3
    frmResizeFont.Show
End Sub

Private Sub Label30_Click()
    fontsizetemp = 30
    frmResizeFont.Show
End Sub

Private Sub Label31_Click()
    fontsizetemp = 31
    frmResizeFont.Show
End Sub

Private Sub Label32_Click()
    fontsizetemp = 32
    frmResizeFont.Show
End Sub

Private Sub Label33_Click()
    fontsizetemp = 33
    frmResizeFont.Show
End Sub

Private Sub Label34_Click()
    fontsizetemp = 34
    frmResizeFont.Show
End Sub

Private Sub Label35_Click()
    fontsizetemp = 35
    frmResizeFont.Show
End Sub

Private Sub Label36_Click()
    fontsizetemp = 36
    frmResizeFont.Show
End Sub

Private Sub Label37_Click()
    fontsizetemp = 37
    frmResizeFont.Show
End Sub

Private Sub Label38_Click()
    fontsizetemp = 38
    frmResizeFont.Show
End Sub

Private Sub Label39_Click()
    fontsizetemp = 39
    frmResizeFont.Show
End Sub

Private Sub Label4_Click()
    fontsizetemp = 4
    frmResizeFont.Show
End Sub

Private Sub Label40_Click()
    fontsizetemp = 40
    frmResizeFont.Show
End Sub

Private Sub Label41_Click()
    fontsizetemp = 41
    frmResizeFont.Show
End Sub

Private Sub Label42_Click()
    fontsizetemp = 42
    frmResizeFont.Show
End Sub

Private Sub Label43_Click()
    fontsizetemp = 43
    frmResizeFont.Show
End Sub

Private Sub Label44_Click()
    fontsizetemp = 44
    frmResizeFont.Show
End Sub

Private Sub Label45_Click()
    fontsizetemp = 45
    frmResizeFont.Show
End Sub

Private Sub Label46_Click()
    fontsizetemp = 46
    frmResizeFont.Show
End Sub

Private Sub Label47_Click()
    fontsizetemp = 47
    frmResizeFont.Show
End Sub

Private Sub Label48_Click()
    fontsizetemp = 48
    frmResizeFont.Show
End Sub

Private Sub Label49_Click()
    fontsizetemp = 49
    frmResizeFont.Show
End Sub

Private Sub Label5_Click()
    fontsizetemp = 5
    frmResizeFont.Show
End Sub

Private Sub Label50_Click()
    fontsizetemp = 50
    frmResizeFont.Show
End Sub

Private Sub Label51_Click()
    fontsizetemp = 51
    frmResizeFont.Show
End Sub

Private Sub Label52_Click()
    fontsizetemp = 52
    frmResizeFont.Show
End Sub

Private Sub Label53_Click()
    fontsizetemp = 53
    frmResizeFont.Show
End Sub

Private Sub Label54_Click()
    fontsizetemp = 54
    frmResizeFont.Show
End Sub

Private Sub Label55_Click()
    fontsizetemp = 55
    frmResizeFont.Show
End Sub

Private Sub Label56_Click()
    fontsizetemp = 56
    frmResizeFont.Show
End Sub

Private Sub Label57_Click()
    fontsizetemp = 57
    frmResizeFont.Show
End Sub

Private Sub Label58_Click()
    fontsizetemp = 58
    frmResizeFont.Show
End Sub

Private Sub Label59_Click()
    fontsizetemp = 59
    frmResizeFont.Show
End Sub

Private Sub Label6_Click()
    fontsizetemp = 6
    frmResizeFont.Show
End Sub

Private Sub Label60_Click()
    fontsizetemp = 60
    frmResizeFont.Show
End Sub

Private Sub Label61_Click()
    fontsizetemp = 61
    frmResizeFont.Show
End Sub

Private Sub Label62_Click()
    fontsizetemp = 62
    frmResizeFont.Show
End Sub

Private Sub Label63_Click()
    fontsizetemp = 63
    frmResizeFont.Show
End Sub

Private Sub Label64_Click()
    fontsizetemp = 64
    frmResizeFont.Show
End Sub

Private Sub Label65_Click()
    fontsizetemp = 65
    frmResizeFont.Show
End Sub

Private Sub Label66_Click()
    fontsizetemp = 66
    frmResizeFont.Show
End Sub

Private Sub Label67_Click()
    fontsizetemp = 67
    frmResizeFont.Show
End Sub

Private Sub Label68_Click()
    fontsizetemp = 68
    frmResizeFont.Show
End Sub

Private Sub Label69_Click()
    fontsizetemp = 69
    frmResizeFont.Show
End Sub

Private Sub Label7_Click()
    fontsizetemp = 7
    frmResizeFont.Show
End Sub

Private Sub Label70_Click()
    fontsizetemp = 70
    frmResizeFont.Show
End Sub

Private Sub Label71_Click()
    fontsizetemp = 71
    frmResizeFont.Show
End Sub

Private Sub Label72_Click()
    fontsizetemp = 72
    frmResizeFont.Show
End Sub

Private Sub Label73_Click()
    fontsizetemp = 73
    frmResizeFont.Show
End Sub

Private Sub Label74_Click()
    fontsizetemp = 74
    frmResizeFont.Show
End Sub

Private Sub Label75_Click()
    fontsizetemp = 75
    frmResizeFont.Show
End Sub

Private Sub Label76_Click()
    fontsizetemp = 76
    frmResizeFont.Show
End Sub

Private Sub Label77_Click()
    fontsizetemp = 77
    frmResizeFont.Show
End Sub

Private Sub Label78_Click()
    fontsizetemp = 78
    frmResizeFont.Show
End Sub

Private Sub Label79_Click()
    fontsizetemp = 79
    frmResizeFont.Show
End Sub

Private Sub Label8_Click()
    fontsizetemp = 8
    frmResizeFont.Show
End Sub

Private Sub Label80_Click()
    fontsizetemp = 80
    frmResizeFont.Show
End Sub

Private Sub Label9_Click()
    fontsizetemp = 9
    frmResizeFont.Show
End Sub

Private Sub Printmenu_Click()
    
    CommonDialog1.ShowPrinter
    Frame1.Top = 2880
    Frame1.Left = 1100
    Me.Height = 14000
    Me.PrintForm
    Unload Me
End Sub

Private Sub Text1_Change()
    Text1.FontSize = 5
End Sub

Private Sub Timer1_Timer()
    Timer1.Enabled = False
    MsgBox ("Click at text label when you want to change textsize or any parameter.")
End Sub
