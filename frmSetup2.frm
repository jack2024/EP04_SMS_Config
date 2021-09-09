VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "Mscomm32.ocx"
Begin VB.Form frmSetup2_2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Setup"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3480
   Icon            =   "frmSetup2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3480
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   2400
      Top             =   4440
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   2880
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
      InBufferSize    =   20
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1560
      TabIndex        =   12
      Top             =   4440
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   375
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Communication Setup"
      Height          =   4215
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   3375
      Begin VB.TextBox Text5 
         Height          =   285
         Left            =   120
         TabIndex        =   17
         Text            =   "Text5"
         Top             =   1560
         Visible         =   0   'False
         Width           =   2775
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   16
         Text            =   "1"
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command4 
         Caption         =   "Change Address"
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Width           =   1815
      End
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   14.25
            Charset         =   222
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   1920
         TabIndex        =   14
         Text            =   "1"
         Top             =   960
         Width           =   1215
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Search"
         Height          =   495
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1815
      End
      Begin VB.ComboBox Combo3 
         Height          =   315
         ItemData        =   "frmSetup2.frx":599A
         Left            =   1320
         List            =   "frmSetup2.frx":59B0
         TabIndex        =   9
         Text            =   "Com1"
         Top             =   3840
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1320
         TabIndex        =   5
         Text            =   "1"
         Top             =   3360
         Width           =   1335
      End
      Begin VB.ComboBox Combo2 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSetup2.frx":59D8
         Left            =   1320
         List            =   "frmSetup2.frx":59E5
         TabIndex        =   4
         Text            =   "None"
         Top             =   2880
         Width           =   1335
      End
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1320
         TabIndex        =   3
         Text            =   "8"
         Top             =   2400
         Width           =   1335
      End
      Begin VB.ComboBox Combo1 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmSetup2.frx":59FA
         Left            =   1320
         List            =   "frmSetup2.frx":5A07
         TabIndex        =   1
         Text            =   "9600"
         Top             =   1920
         Width           =   1335
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Comm Port :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Stop Bits :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label lblParity 
         Alignment       =   1  'Right Justify
         Caption         =   "Parity :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Label lblDataBits 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Bits :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Baud Rate :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   1920
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmSetup2_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)


Dim CRCTable(0 To 511) As Byte
Dim CRC_Low, CRC_High As Byte

Dim Buffer As String     'store Data recieved from ESPAN-03 Board
Dim Buffer2 As String
Dim DatBuff(20) As Byte      '1-2 Input type,3-4 Fault type,5-6 Output type,7 Flashing Rate, 8-9 Indicator
Dim Temp_text As String
Dim Poll_Counter As Byte
Dim Search_Addr As Boolean
Dim Address As Integer


Private Sub Command1_Click()
    Me.Hide
    
    
    'If Not frmMainV1_0.MSComm1.PortOpen Then
        'frmMainV1_0.MSComm1.CommPort = Mid(frmSetup.Combo3.Text, 4, 1)
   ' ElseIf frmMainV1_1.MSComm1.PortOpen Then
       ' frmMainV1_1.MSComm1.CommPort = Mid(frmSetup.Combo3.Text, 4, 1)
    'ElseIf frmMainV1_2.MSComm1.PortOpen Then
       ' frmMainV1_2.MSComm1.CommPort = Mid(frmSetup.Combo3.Text, 4, 1)
   ' ElseIf frmMainV2_0.MSComm1.PortOpen Then
       ' frmMainV2_0.MSComm1.CommPort = Mid(frmSetup.Combo3.Text, 4, 1)
   ' ElseIf frmMainV2_2.MSComm1.PortOpen Then
        frmMainConfig.MSComm1.CommPort = Mid(frmSetup.Combo3.Text, 4, 1)
    'End If
    
    SaveSetting "ESPAN-03", "Setting", "Comm Port", Mid(frmSetup2_2.Combo3.Text, 4, 1)
    Unload Me
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Command3_Click()
    
    If frmMainConfig.Timer1.Enabled = True Then frmMainConfig.Timer1.Enabled = False
    If frmMainConfig.MSComm1.PortOpen = True Then frmMainConfig.MSComm1.PortOpen = False
    
    If Command3.Caption = "Search" Then
        Command3.Caption = "Stop"
        'Shape1.BackColor = vbRed
        Address = 1
        Search_Addr = True
        Timer1.Enabled = True
    ElseIf Command3.Caption = "Stop" Then
        Command3.Caption = "Search"
        Timer1.Enabled = False
    End If
End Sub

Private Sub Command4_Click()
Dim Data As String
    If Not MSComm1.PortOpen Then
                MSComm1.CommPort = GetSetting("ESPAN-03", "Setting", "Comm Port", "1")
                MSComm1.Settings = "9600,N,8,1"
                MSComm1.PortOpen = True
                MSComm1.RThreshold = 1
            End If
        
            Data = ""
            Data = Chr(Val(Text3.Text)) + Chr(5) + Chr(0) + Chr(100) + Chr(0) + Chr(Val(Text4.Text))
            CRC_16 Data, 6
            Data = Data + Chr(CRC_High) + Chr(CRC_Low)
            MSComm1.InputLen = 0
            MSComm1.Output = Data
            Text3.Text = Text4.Text
            Buffer = ""
            Do While MSComm1.OutBufferCount > 0
            Loop
End Sub

Private Sub Form_Load()
   ' If frmMainV2_2.MSComm1.PortOpen = True Then frmMainV2_2.MSComm1.PortOpen = False
    Combo3.Text = "Com" & GetSetting("ESPAN-03", "Setting", "Comm Port", "1")
    
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
    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    Unload Me
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

Private Sub MSComm1_OnComm()
Dim i As Integer
Dim aa As String

    Select Case MSComm1.CommEvent
        Case comEvReceive
        'Buffer = ""
        Sleep (200)                         'Wait for receive completed data lenght
            Buffer = Buffer + MSComm1.Input
            
            Text5.Text = ""
            Text5.Text = Buffer
    End Select
    
    aa = Asc(Mid(Buffer, 2, 1))
    If (Asc(Mid(Buffer, 2, 1)) = 20) Then            'get "2" and "6"   Send text function
        If (Len(Buffer) = 18) Then
    'If (Len(Buffer) = 20) Then
            Timer1.Enabled = False
            Command3.Caption = "Search"
            Search_Addr = False
        'Timer2.Enabled = False
            Poll_Counter = 0
            'frmMainConfig.Text1.Text = Text3.Text
        
        'For i = 4 To 12
        '    DatBuff(i - 3) = Asc(Mid(Buffer, i, 1))
            'DecodeValue
        'Next i
        'DecodeValue
        'Shape1.BackColor = vbGreen
            Buffer = ""
        End If
    Else
    Buffer = ""
    End If
End Sub

Private Sub Timer1_Timer()
    Dim Data As String

            If Search_Addr = False Then
                Timer1.Enabled = False
                GoTo AAA
            End If
            
            If Not MSComm1.PortOpen Then
                MSComm1.CommPort = GetSetting("ESPAN-03", "Setting", "Comm Port", "1")
                MSComm1.Settings = "9600,N,8,1"
                MSComm1.PortOpen = True
                MSComm1.RThreshold = 1
            End If
        
            Text3.Text = Address
            Text4.Text = Address
            Data = ""
            Data = Chr(Val(Text3.Text)) + Chr(20) + Chr(1) + Chr(1)
            CRC_16 Data, 4
            Data = Data + Chr(CRC_High) + Chr(CRC_Low)
            MSComm1.InputLen = 0
            MSComm1.Output = Data
            Buffer = ""
            Do While MSComm1.OutBufferCount > 0
            Loop
            Address = Address + 1
            If Address > 247 Then
                Timer1.Enabled = False
                MsgBox "Device not found"
                Command3.Caption = "Search"
            End If
AAA:
End Sub


