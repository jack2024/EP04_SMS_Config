VERSION 5.00
Begin VB.UserControl InputProperties 
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   ScaleHeight     =   1080
   ScaleWidth      =   5760
   Begin VB.Frame Frame7 
      Height          =   975
      Left            =   120
      TabIndex        =   10
      Top             =   0
      Width           =   1335
      Begin VB.CheckBox ChkIndicator 
         Caption         =   "Indicator"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblInput 
         Caption         =   "INPUT 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   222
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame4 
      Height          =   975
      Left            =   3840
      TabIndex        =   6
      Top             =   0
      Width           =   1815
      Begin VB.OptionButton OptBell 
         Caption         =   "Bell"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   735
      End
      Begin VB.OptionButton OptBuzzer 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Buzzer"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.CheckBox ChkBoth 
         Caption         =   "Both"
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
   End
   Begin VB.Frame Frame6 
      Height          =   975
      Left            =   1320
      TabIndex        =   3
      Top             =   0
      Width           =   1335
      Begin VB.OptionButton OptNc 
         Caption         =   "NC"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   975
      End
      Begin VB.OptionButton OptNo 
         BackColor       =   &H00C0C0C0&
         Caption         =   "NO"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
   End
   Begin VB.Frame Frame3 
      Height          =   975
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   1335
      Begin VB.OptionButton OptManual 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Manual"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   855
      End
      Begin VB.OptionButton OptAuto 
         Caption         =   "Auto"
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "InputProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim InputTypeReg As String
Dim FaultTypeReg As String
Dim OutputTypeReg As String
Dim OutputBothReg As String
Dim Fault_IndicatorReg As String

'Public Event UserClick()




Public Property Get InputType() As Variant
    InputType = InputTypeReg
End Property

Public Property Let InputType(ByVal vNewValue As Variant)
    If vNewValue = 1 Or vNewValue = "NO" Or vNewValue = "no" Then
        OptNo.Value = True
        OptNc.Value = False
        InputTypeReg = "NO"
    ElseIf vNewValue = 0 Or vNewValue = "NC" Or vNewValue = "nc" Then
        OptNo.Value = False
        OptNc.Value = True
        InputTypeReg = "NC"
    End If
    
    PropertyChanged "InputType"
End Property


Public Property Get Caption() As Variant
    Caption = UserControl.lblInput.Caption
End Property

Public Property Let Caption(ByVal vNewValue As Variant)
    UserControl.lblInput.Caption = vNewValue
    'PropertyChanged "Caption"
End Property

Public Property Get FaultType() As Variant
    FaultType = FaultTypeReg
End Property

Public Property Let FaultType(ByVal vNewValue As Variant)
    If vNewValue = 1 Or vNewValue = "MANUAL" Or vNewValue = "manual" Then
        OptManual.Value = True
        OptAuto.Value = False
        FaultTypeReg = "MANUAL"
    ElseIf vNewValue = 0 Or vNewValue = "AUTO" Or vNewValue = "auto" Then
        OptManual.Value = False
        OptAuto.Value = True
        FaultTypeReg = "AUTO"
    End If
End Property

Public Property Get OutputType() As Variant
    OutputType = OutputTypeReg
End Property

Public Property Let OutputType(ByVal vNewValue As Variant)
    If vNewValue = 1 Or vNewValue = "BUZZER" Or vNewValue = "buzzer" Then
        OptBuzzer.Value = True
        OptBell.Value = False
        OutputTypeReg = "BUZZER"
    ElseIf vNewValue = 0 Or vNewValue = "BELL" Or vNewValue = "bell" Then
        OptBuzzer.Value = False
        OptBell.Value = True
        OutputTypeReg = "BELL"
    End If
End Property

Public Property Get OutputBoth() As Variant
    OutputBoth = OutputBothReg
End Property

Public Property Let OutputBoth(ByVal vNewValue As Variant)
    If vNewValue = 1 Or vNewValue = "BOTH" Or vNewValue = "both" Then
        ChkBoth.Value = 1
        OutputBothReg = "BOTH"
        OptBuzzer.Enabled = False
        OptBell.Enabled = False
        'OptBuzzer.Value = True

        
    ElseIf vNewValue = 0 Or vNewValue = "SINGLE" Or vNewValue = "singgle" Then
        ChkBoth.Value = 0
        OptBuzzer.Enabled = True
        OptBell.Enabled = True
        'OptBuzzer.Value = True
        OutputBothReg = "SINGLE"
    End If
End Property

Public Property Get Fault_Indicator() As Variant
    Fault_Indicator = Fault_IndicatorReg
End Property

Public Property Let Fault_Indicator(ByVal vNewValue As Variant)
    If vNewValue = 1 Or vNewValue = "INDICATOR" Or vNewValue = "indicator" Then
        ChkIndicator.Value = 1
        OptManual.Enabled = False
        OptAuto.Enabled = False
        ChkBoth.Enabled = False
        OptBuzzer.Enabled = False
        OptBell.Enabled = False
        Fault_IndicatorReg = "INDICATOR"
        
    ElseIf vNewValue = 0 Or vNewValue = "FAULT" Or vNewValue = "fault" Then
        ChkIndicator.Value = 0
        OptManual.Enabled = True
        OptAuto.Enabled = True
        ChkBoth.Enabled = True
        If ChkBoth.Value = 0 Then
            OptBuzzer.Enabled = True
            OptBell.Enabled = True
        End If
        Fault_IndicatorReg = "FAULT"
        
    End If
End Property

Private Sub ChkBoth_Click()
    
    If ChkBoth.Value = 1 Then
        OutputBothReg = "BOTH"
        OptBuzzer.Enabled = False
        OptBell.Enabled = False
        OptBuzzer.Value = True
    Else
        OutputBothReg = "SINGLE"
        OptBuzzer.Enabled = True
        OptBell.Enabled = True
        OptBuzzer.Value = True
        
    End If
End Sub

Private Sub ChkIndicator_Click()
    If ChkIndicator = 1 Then
        Fault_IndicatorReg = "INDICATOR"
        FaultTypeReg = "MANUAL"
        OptManual.Enabled = False
        OptAuto.Enabled = False
        OptManual.Value = True
        
        OutputBothReg = "SINGLE"
        ChkBoth.Enabled = False
        ChkBoth.Value = 0

        OutputTypeReg = "BUZZER"
        OptBuzzer.Enabled = False
        OptBell.Enabled = False
        OptBuzzer.Value = True

    Else
        Fault_IndicatorReg = "FAULT"
        FaultTypeReg = "MANUAL"
        OptManual.Enabled = True
        OptAuto.Enabled = True
        OptManual.Value = True
        
        OutputBothReg = "SINGLE"
        ChkBoth.Enabled = True
        ChkBoth.Value = 0
        
        OutputTypeReg = "BUZZER"
        OptBuzzer.Enabled = True
        OptBell.Enabled = True
        OptBuzzer.Value = True
        
    End If
End Sub

Private Sub OptAuto_Click()
    FaultTypeReg = "AUTO"
    OptAuto.BackColor = &HC0C0C0
    OptManual.BackColor = &H8000000F
End Sub

Private Sub OptBell_Click()
    OutputTypeReg = "BELL"
    OptBell.BackColor = &HC0C0C0
    OptBuzzer.BackColor = &H8000000F
End Sub

Private Sub OptBuzzer_Click()
    OutputTypeReg = "BUZZER"
    OptBuzzer.BackColor = &HC0C0C0
    OptBell.BackColor = &H8000000F
End Sub

Private Sub OptManual_Click()
    FaultTypeReg = "MANUAL"
    OptManual.BackColor = &HC0C0C0
    OptAuto.BackColor = &H8000000F
End Sub

Private Sub OptNc_Click()
    InputTypeReg = "NC"
    OptNc.BackColor = &HC0C0C0
    OptNo.BackColor = &H8000000F
End Sub

Private Sub OptNo_Click()
    InputTypeReg = "NO"
    OptNo.BackColor = &HC0C0C0
    OptNc.BackColor = &H8000000F
End Sub

Private Sub UserControl_Initialize()
    
    ChkIndicator.Value = 0
    ChkBoth.Value = 0
    OptNo.Value = True
    OptManual.Value = True
    OptBuzzer.Value = True
    ChkIndicator.Value = 0
    ChkBoth.Value = 0

    InputType = "NO"
    FaultType = "MANUAL"
    OutputType = "BUZZER"
    OutputBoth = "SINGLE"
    Fault_IndicatorReg = "FAULT"
End Sub

