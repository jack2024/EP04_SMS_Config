VERSION 5.00
Begin VB.Form frmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2895
   Icon            =   "frmSetup.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   2895
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Communication Setup"
      Height          =   2775
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   2895
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmSetup.frx":599A
         Left            =   1320
         List            =   "frmSetup.frx":59B0
         TabIndex        =   9
         Text            =   "Com1"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Text            =   "1"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Text            =   "None"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "8"
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   1
         Text            =   "9600"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Comm Port :"
         BeginProperty Font 
            Name            =   "Leelawadee UI"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Stop Bits :"
         BeginProperty Font 
            Name            =   "Leelawadee UI"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label lblParity 
         Alignment       =   1  'Right Justify
         Caption         =   "Parity :"
         BeginProperty Font 
            Name            =   "Leelawadee UI"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label lblDataBits 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Bits :"
         BeginProperty Font 
            Name            =   "Leelawadee UI"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Baud Rate :"
         BeginProperty Font 
            Name            =   "Leelawadee UI"
            Size            =   9.75
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub Command1_Click()
    
    Me.Hide
    
    If Not frmMainConfig.MSComm1.PortOpen Then
        frmMainConfig.MSComm1.CommPort = Mid(frmSetup.Combo3.Text, 4, 1)
    End If
    
    SaveSetting "ESPAN-01", "Setting", "Comm Port", Mid(frmSetup.Combo3.Text, 4, 1)
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Combo3.AddItem "COM2", 0
    
    Combo3.Text = "Com" & GetSetting("ESPAN-01", "Setting", "Comm Port", "1")
End Sub


