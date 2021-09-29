VERSION 5.00
Begin VB.Form SMSsetting 
   Caption         =   "FaultName Setting"
   ClientHeight    =   8520
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9660
   Icon            =   "SMSsetting.frx":0000
   ScaleHeight     =   8520
   ScaleWidth      =   9660
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Height          =   8400
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   9400
      Begin VB.TextBox Text19 
         Height          =   375
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   41
         Text            =   "Fault_19"
         Top             =   6120
         Width           =   3375
      End
      Begin VB.TextBox Text11 
         Height          =   375
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   40
         Text            =   "Fault_11"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Text20 
         Height          =   375
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   35
         Text            =   "Fault_20"
         Top             =   6840
         Width           =   3375
      End
      Begin VB.TextBox Text18 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   34
         Text            =   "Fault_18"
         Top             =   6840
         Width           =   3375
      End
      Begin VB.TextBox Text17 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   33
         Text            =   "Fault_17"
         Top             =   6120
         Width           =   3375
      End
      Begin VB.TextBox Text10 
         Height          =   375
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   30
         Text            =   "Fault_10"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text16 
         Height          =   375
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   23
         Text            =   "Fault_16"
         Top             =   5400
         Width           =   3375
      End
      Begin VB.TextBox Text15 
         Height          =   375
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   22
         Text            =   "Fault_15"
         Top             =   4680
         Width           =   3375
      End
      Begin VB.TextBox Text14 
         Height          =   375
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   21
         Text            =   "Fault_14"
         Top             =   3960
         Width           =   3375
      End
      Begin VB.TextBox Text13 
         Height          =   375
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   20
         Text            =   "Fault_13"
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   19
         Text            =   "Fault_12"
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox Text9 
         Height          =   375
         Left            =   5700
         MaxLength       =   30
         TabIndex        =   17
         Text            =   "Fault_9"
         Top             =   360
         Width           =   3375
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   9
         Text            =   "Fault_8"
         Top             =   5400
         Width           =   3375
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   8
         Text            =   "Fault_7"
         Top             =   4680
         Width           =   3375
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   7
         Text            =   "Fault_6"
         Top             =   3960
         Width           =   3375
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   6
         Text            =   "Fault_5"
         Top             =   3240
         Width           =   3375
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   5
         Text            =   "Fault_4"
         Top             =   2520
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   4
         Text            =   "Fault_3"
         Top             =   1800
         Width           =   3375
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   3
         Text            =   "Fault_2"
         Top             =   1080
         Width           =   3375
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   960
         MaxLength       =   30
         TabIndex        =   2
         Text            =   "Fault_1"
         Top             =   360
         Width           =   3375
      End
      Begin VB.CommandButton SAVE_btn 
         Caption         =   "SAVE"
         Height          =   615
         Left            =   4000
         TabIndex        =   1
         Top             =   7600
         Width           =   2055
      End
      Begin VB.Label Label20 
         Caption         =   "INPUT20"
         Height          =   255
         Left            =   4800
         TabIndex        =   39
         Top             =   6900
         Width           =   735
      End
      Begin VB.Label Label19 
         Caption         =   "INPUT19"
         Height          =   255
         Left            =   4800
         TabIndex        =   38
         Top             =   6180
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "INPUT18"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   6900
         Width           =   735
      End
      Begin VB.Label Label17 
         Caption         =   "INPUT17"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   6180
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "INPUT10"
         Height          =   255
         Left            =   4800
         TabIndex        =   32
         Top             =   1195
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "INPUT2"
         Height          =   255
         Left            =   200
         TabIndex        =   31
         Top             =   1195
         Width           =   735
      End
      Begin VB.Label Label16 
         Caption         =   "INPUT16"
         Height          =   255
         Left            =   4800
         TabIndex        =   29
         Top             =   5490
         Width           =   735
      End
      Begin VB.Label Label15 
         Caption         =   "INPUT15"
         Height          =   255
         Left            =   4800
         TabIndex        =   28
         Top             =   4770
         Width           =   735
      End
      Begin VB.Label Label14 
         Caption         =   "INPUT14"
         Height          =   255
         Left            =   4800
         TabIndex        =   27
         Top             =   4055
         Width           =   735
      End
      Begin VB.Label Label13 
         Caption         =   "INPUT13"
         Height          =   255
         Left            =   4800
         TabIndex        =   26
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label12 
         Caption         =   "INPUT12"
         Height          =   255
         Left            =   4800
         TabIndex        =   25
         Top             =   2625
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "INPUT11"
         Height          =   255
         Left            =   4800
         TabIndex        =   24
         Top             =   1910
         Width           =   735
      End
      Begin VB.Label Label9 
         Caption         =   "INPUT9"
         Height          =   255
         Left            =   4800
         TabIndex        =   18
         Top             =   420
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "INPUT8"
         Height          =   375
         Left            =   200
         TabIndex        =   16
         Top             =   5490
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "INPUT7"
         Height          =   255
         Left            =   200
         TabIndex        =   15
         Top             =   4770
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "INPUT6"
         Height          =   255
         Left            =   200
         TabIndex        =   14
         Top             =   4055
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "INPUT5"
         Height          =   255
         Left            =   200
         TabIndex        =   13
         Top             =   3360
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "INPUT4"
         Height          =   255
         Left            =   200
         TabIndex        =   12
         Top             =   2625
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "INPUT3"
         Height          =   255
         Left            =   200
         TabIndex        =   11
         Top             =   1910
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "INPUT1"
         Height          =   255
         Left            =   200
         TabIndex        =   10
         Top             =   420
         Width           =   735
      End
   End
End
Attribute VB_Name = "SMSsetting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim CRCTable(0 To 511) As Byte
Dim CRC_Low, CRC_High As Byte

Private Sub Form_Load()
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

    
    If frmMainConfig.OptNoInput8.Value = True Then
       SMSsetting.Height = 7770
       SMSsetting.Width = 4900
       SAVE_btn.Top = 6100
       SAVE_btn.Left = 1300
       Frame1.Height = 7000
       Frame1.Width = 4500
       Text17.Visible = False
       Text18.Visible = False
       Label17.Visible = False
       Label18.Visible = False
    ElseIf frmMainConfig.OptNoInput10.Value = True Then
       SMSsetting.Height = 9100
       SMSsetting.Width = 4900
       SAVE_btn.Top = 7600
       SAVE_btn.Left = 1300
       Frame1.Height = 8400
       Frame1.Width = 4500
       Text17.Visible = False
       Text18.Visible = False
       Label17.Visible = False
       Label18.Visible = False
       Label9.Left = 120
       Label10.Left = 120
       Text9.Left = 960
       Text10.Left = 960
       
       Label9.Top = 6180
       Label10.Top = 6900
       Text9.Top = 6120
       Text10.Top = 6840
       
       
    ElseIf frmMainConfig.OptNoInput16.Value = True Then
       SMSsetting.Height = 7770
       SMSsetting.Width = 9900
       SAVE_btn.Top = 6100
       SAVE_btn.Left = 4000
       Frame1.Height = 7000
       Frame1.Width = 9400
       Text17.Visible = False
       Text18.Visible = False
       Label17.Visible = False
       Label18.Visible = False
       Text19.Visible = False
       Text20.Visible = False
       Label19.Visible = False
       Label20.Visible = False
    ElseIf frmMainConfig.OptNoInput20.Value = True Then
       SMSsetting.Height = 9100
       SMSsetting.Width = 9900
       SAVE_btn.Top = 7600
       SAVE_btn.Left = 4000
       Frame1.Height = 8400
       Frame1.Width = 9400
       
       Label9.Top = 6180
       Label10.Top = 6900
       Text9.Top = 6120
       Text10.Top = 6840
       
       Label9.Left = 200
       Label10.Left = 120
       Text9.Left = 960
       Text10.Left = 960
       
       Label11.Top = 420
       Label12.Top = 1195
       Label13.Top = 1910
       Label14.Top = 2625
       Label15.Top = 3360
       Label16.Top = 4055
       Label17.Top = 4770
       Label18.Top = 5490
       
       Label11.Left = 4800
       Label12.Left = 4800
       Label13.Left = 4800
       Label14.Left = 4800
       Label15.Left = 4800
       Label16.Left = 4800
       Label17.Left = 4800
       Label18.Left = 4800
       
       Text11.Left = 5700
       Text12.Left = 5700
       Text13.Left = 5700
       Text14.Left = 5700
       Text15.Left = 5700
       Text16.Left = 5700
       Text17.Left = 5700
       Text18.Left = 5700
 
       Text11.Top = 360
       Text12.Top = 1080
       Text13.Top = 1800
       Text14.Top = 2520
       Text15.Top = 3240
       Text16.Top = 3960
       Text17.Top = 4680
       Text18.Top = 5400
                     
    End If
        
    
    
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

Private Sub SAVE_btn_Click()
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
        'Error = CalculateValue      'check valid data
        'If Error Then GoTo AAA
        
        If frmMainConfig.CommState = "IDLE" Then
            On Error GoTo Error
            If Not frmMainConfig.MSComm1.PortOpen Then
                frmMainConfig.MSComm1.CommPort = GetSetting("ESPAN-01", "Setting", "Comm Port", "1")
                'MSComm1.CommPort = 1
                frmMainConfig.MSComm1.Settings = "9600,N,8,1"
                frmMainConfig.MSComm1.PortOpen = True
                frmMainConfig.MSComm1.RThreshold = 1
            End If
        
            Data = ""
            mymassage = "" 'Clear Buffer
                                    
            If frmMainConfig.OptNoInput8.Value = True Then
                mymassage = mymassage + Text1.Text + Chr(13)
                mymassage = mymassage + Text2.Text + Chr(13)
                mymassage = mymassage + Text3.Text + Chr(13)
                mymassage = mymassage + Text4.Text + Chr(13)
                mymassage = mymassage + Text5.Text + Chr(13)
                mymassage = mymassage + Text6.Text + Chr(13)
                mymassage = mymassage + Text7.Text + Chr(13)
                mymassage = mymassage + Text8.Text + Chr(13)
               
            ElseIf frmMainConfig.OptNoInput10.Value = True Then
                mymassage = mymassage + Text1.Text + Chr(13)
                mymassage = mymassage + Text2.Text + Chr(13)
                mymassage = mymassage + Text3.Text + Chr(13)
                mymassage = mymassage + Text4.Text + Chr(13)
                mymassage = mymassage + Text5.Text + Chr(13)
                mymassage = mymassage + Text6.Text + Chr(13)
                mymassage = mymassage + Text7.Text + Chr(13)
                mymassage = mymassage + Text8.Text + Chr(13)
                mymassage = mymassage + Text9.Text + Chr(13)
                mymassage = mymassage + Text10.Text + Chr(13)
               
            ElseIf frmMainConfig.OptNoInput16.Value = True Then
                mymassage = mymassage + Text1.Text + Chr(13)
                mymassage = mymassage + Text2.Text + Chr(13)
                mymassage = mymassage + Text3.Text + Chr(13)
                mymassage = mymassage + Text4.Text + Chr(13)
                mymassage = mymassage + Text5.Text + Chr(13)
                mymassage = mymassage + Text6.Text + Chr(13)
                mymassage = mymassage + Text7.Text + Chr(13)
                mymassage = mymassage + Text8.Text + Chr(13)
                mymassage = mymassage + Text9.Text + Chr(13)
                mymassage = mymassage + Text10.Text + Chr(13)
                mymassage = mymassage + Text11.Text + Chr(13)
                mymassage = mymassage + Text12.Text + Chr(13)
                mymassage = mymassage + Text13.Text + Chr(13)
                mymassage = mymassage + Text14.Text + Chr(13)
                mymassage = mymassage + Text15.Text + Chr(13)
                mymassage = mymassage + Text16.Text + Chr(13)
               
            ElseIf frmMainConfig.OptNoInput20.Value = True Then
               mymassage = mymassage + Text1.Text + Chr(13)
                mymassage = mymassage + Text2.Text + Chr(13)
                mymassage = mymassage + Text3.Text + Chr(13)
                mymassage = mymassage + Text4.Text + Chr(13)
                mymassage = mymassage + Text5.Text + Chr(13)
                mymassage = mymassage + Text6.Text + Chr(13)
                mymassage = mymassage + Text7.Text + Chr(13)
                mymassage = mymassage + Text8.Text + Chr(13)
                mymassage = mymassage + Text9.Text + Chr(13)
                mymassage = mymassage + Text10.Text + Chr(13)
                mymassage = mymassage + Text11.Text + Chr(13)
                mymassage = mymassage + Text12.Text + Chr(13)
                mymassage = mymassage + Text13.Text + Chr(13)
                mymassage = mymassage + Text14.Text + Chr(13)
                mymassage = mymassage + Text15.Text + Chr(13)
                mymassage = mymassage + Text16.Text + Chr(13)
                mymassage = mymassage + Text17.Text + Chr(13)
                mymassage = mymassage + Text18.Text + Chr(13)
                mymassage = mymassage + Text19.Text + Chr(13)
                mymassage = mymassage + Text20.Text + Chr(13)
            End If
                                               
            Dim massagelen As Integer
            Dim Hi_maassagelen As Byte
            Dim Lo_maassagelen As Byte
            Dim testlen As Integer
            massagelen = Len(mymassage)
            If massagelen < 100 Then
                Hi_maassagelen = 0
                Lo_maassagelen = massagelen Mod 100
            Else
                Hi_maassagelen = massagelen \ 100
                Lo_maassagelen = massagelen Mod 100
            End If
            
            'testlen = (Hi_maassagelen * 100) + Lo_maassagelen ' test data
            Data = ""
                                                             '0x22  WRITE FAULTNAME
            'Data = Chr(Val(frmMainConfig.TextAddrNow.Text)) + Chr(34) + Chr(massagelen) '59 are data range(62) -3
            
            Data = Chr(Val(frmMainConfig.TextAddrNow.Text)) + Chr(34) + Chr(Hi_maassagelen) + Chr(Lo_maassagelen)
            
            Data = Data + mymassage
            massagelen = Len(Data)
            
            CRC_16 Data, massagelen
        
            Data = Data + Chr(CRC_High) + Chr(CRC_Low)
            frmMainConfig.MSComm1.InputLen = 0
            frmMainConfig.MSComm1.Output = Data
            Buffer = ""
            Do While frmMainConfig.MSComm1.OutBufferCount > 0
            Loop
            'Bit_Setting = True
            'Timer2.Enabled = True
            frmMainConfig.CommState = "DOWNLOAD"
            frmMainConfig.Timer1.Enabled = True
        End If
        
        GoTo AAA
Error:
        MsgBox ("Invalid Port Number.")
        frmSetup.Show
AAA:



End Sub


