VERSION 5.00
Begin VB.Form frmResizeFont 
   Caption         =   "ResizeFont"
   ClientHeight    =   3990
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6705
   Icon            =   "Formresizefon2t.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3990
   ScaleWidth      =   6705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Setting"
      Height          =   3135
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   6015
      Begin VB.CommandButton Command2 
         Caption         =   "OK"
         Height          =   375
         Left            =   3000
         TabIndex        =   7
         Top             =   2400
         Width           =   1455
      End
      Begin VB.TextBox Text2 
         Height          =   405
         Left            =   1200
         TabIndex        =   6
         Text            =   "Text2"
         Top             =   840
         Width           =   4575
      End
      Begin VB.TextBox Text1 
         Height          =   405
         Left            =   1200
         TabIndex        =   5
         Text            =   "Text1"
         Top             =   285
         Width           =   4575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "2Line"
         Height          =   255
         Left            =   480
         TabIndex        =   4
         Top             =   915
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "1Line"
         Height          =   255
         Left            =   480
         TabIndex        =   3
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Change"
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   2400
         Width           =   1455
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         Left            =   2400
         TabIndex        =   1
         Text            =   "Combo1"
         Top             =   1680
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmResizeFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
    If FormPrint20preview.fontsizetemp = 1 Then
        FormPrint20preview.Label1.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label1.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 2 Then
        FormPrint20preview.Label2.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label2.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 3 Then
        FormPrint20preview.Label3.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label3.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 4 Then
        FormPrint20preview.Label4.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label4.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 5 Then
        FormPrint20preview.Label5.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label5.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 6 Then
        FormPrint20preview.Label6.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label6.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 7 Then
        FormPrint20preview.Label7.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label7.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 8 Then
        FormPrint20preview.Label8.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label8.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 9 Then
        FormPrint20preview.Label9.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label9.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 10 Then
        FormPrint20preview.Label10.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label10.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 11 Then
        FormPrint20preview.Label11.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label11.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 12 Then
        FormPrint20preview.Label12.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label12.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 13 Then
        FormPrint20preview.Label13.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label13.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 14 Then
        FormPrint20preview.Label14.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label14.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 15 Then
        FormPrint20preview.Label15.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label15.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 16 Then
        FormPrint20preview.Label16.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label16.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 17 Then
        FormPrint20preview.Label17.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label17.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 18 Then
        FormPrint20preview.Label18.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label18.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 19 Then
        FormPrint20preview.Label19.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label19.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 20 Then
        FormPrint20preview.Label20.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label20.Caption = Text1.Text + vbCrLf + Text2.Text
        
    ElseIf FormPrint20preview.fontsizetemp = 21 Then
        FormPrint20preview.Label21.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label21.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 22 Then
        FormPrint20preview.Label22.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label22.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 23 Then
        FormPrint20preview.Label23.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label23.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 24 Then
        FormPrint20preview.Label24.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label24.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 25 Then
        FormPrint20preview.Label25.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label25.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 26 Then
        FormPrint20preview.Label26.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label26.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 27 Then
        FormPrint20preview.Label27.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label27.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 28 Then
        FormPrint20preview.Label28.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label28.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 29 Then
        FormPrint20preview.Label29.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label29.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 30 Then
        FormPrint20preview.Label30.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label30.Caption = Text1.Text + vbCrLf + Text2.Text
        
    ElseIf FormPrint20preview.fontsizetemp = 31 Then
        FormPrint20preview.Label31.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label31.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 32 Then
        FormPrint20preview.Label32.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label32.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 33 Then
        FormPrint20preview.Label33.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label33.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 34 Then
        FormPrint20preview.Label34.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label34.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 35 Then
        FormPrint20preview.Label35.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label35.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 36 Then
        FormPrint20preview.Label36.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label36.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 37 Then
        FormPrint20preview.Label37.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label37.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 38 Then
        FormPrint20preview.Label38.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label38.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 39 Then
        FormPrint20preview.Label39.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label39.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 40 Then
        FormPrint20preview.Label40.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label40.Caption = Text1.Text + vbCrLf + Text2.Text
        
    ElseIf FormPrint20preview.fontsizetemp = 41 Then
        FormPrint20preview.Label41.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label41.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 42 Then
        FormPrint20preview.Label42.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label42.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 43 Then
        FormPrint20preview.Label43.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label43.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 44 Then
        FormPrint20preview.Label44.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label44.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 45 Then
        FormPrint20preview.Label45.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label45.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 46 Then
        FormPrint20preview.Label46.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label46.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 47 Then
        FormPrint20preview.Label47.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label47.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 48 Then
        FormPrint20preview.Label48.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label48.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 49 Then
        FormPrint20preview.Label49.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label49.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 50 Then
        FormPrint20preview.Label50.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label50.Caption = Text1.Text + vbCrLf + Text2.Text
        
    ElseIf FormPrint20preview.fontsizetemp = 51 Then
        FormPrint20preview.Label51.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label51.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 52 Then
        FormPrint20preview.Label52.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label52.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 53 Then
        FormPrint20preview.Label53.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label53.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 54 Then
        FormPrint20preview.Label54.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label54.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 55 Then
        FormPrint20preview.Label55.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label55.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 56 Then
        FormPrint20preview.Label56.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label56.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 57 Then
        FormPrint20preview.Label57.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label57.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 58 Then
        FormPrint20preview.Label58.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label58.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 59 Then
        FormPrint20preview.Label59.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label59.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 60 Then
        FormPrint20preview.Label60.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label60.Caption = Text1.Text + vbCrLf + Text2.Text
        
    ElseIf FormPrint20preview.fontsizetemp = 61 Then
        FormPrint20preview.Label61.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label61.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 62 Then
        FormPrint20preview.Label62.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label62.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 63 Then
        FormPrint20preview.Label63.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label63.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 64 Then
        FormPrint20preview.Label64.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label64.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 65 Then
        FormPrint20preview.Label65.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label65.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 66 Then
        FormPrint20preview.Label66.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label66.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 67 Then
        FormPrint20preview.Label67.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label67.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 68 Then
        FormPrint20preview.Label68.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label68.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 69 Then
        FormPrint20preview.Label69.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label69.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 70 Then
        FormPrint20preview.Label70.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label70.Caption = Text1.Text + vbCrLf + Text2.Text
        
    ElseIf FormPrint20preview.fontsizetemp = 71 Then
        FormPrint20preview.Label71.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label71.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 72 Then
        FormPrint20preview.Label72.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label72.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 73 Then
        FormPrint20preview.Label73.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label73.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 74 Then
        FormPrint20preview.Label74.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label74.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 75 Then
        FormPrint20preview.Label75.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label75.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 76 Then
        FormPrint20preview.Label76.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label76.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 77 Then
        FormPrint20preview.Label77.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label77.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 78 Then
        FormPrint20preview.Label78.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label78.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 79 Then
        FormPrint20preview.Label79.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label79.Caption = Text1.Text + vbCrLf + Text2.Text
    ElseIf FormPrint20preview.fontsizetemp = 80 Then
        FormPrint20preview.Label80.FontSize = Combo1.ListIndex + 5
        FormPrint20preview.Label80.Caption = Text1.Text + vbCrLf + Text2.Text
        
    End If
    
    'FormPrint20preview.Label1.FontSize = Combo1.ListIndex + 5
    'FormPrint20preview.fontsizetemp = Combo1.ListIndex + 5
End Sub





Private Sub Form_Load()
'Combo1.AddItem
    For i = 0 To 10 Step 1
        Combo1.List(i) = i + 5
    Next i
    Combo1.ListIndex = 3
    Option1.Value = True
    If Option1.Value = True Then
        Text2.Visible = False
        
    End If
    
    If FormPrint20preview.fontsizetemp = 1 Then
        Text1.Text = FormPrint20preview.Label1.Caption
    ElseIf FormPrint20preview.fontsizetemp = 2 Then
        Text1.Text = FormPrint20preview.Label2.Caption
    ElseIf FormPrint20preview.fontsizetemp = 3 Then
        Text1.Text = FormPrint20preview.Label3.Caption
    ElseIf FormPrint20preview.fontsizetemp = 4 Then
        Text1.Text = FormPrint20preview.Label4.Caption
    ElseIf FormPrint20preview.fontsizetemp = 5 Then
        Text1.Text = FormPrint20preview.Label5.Caption
    ElseIf FormPrint20preview.fontsizetemp = 6 Then
        Text1.Text = FormPrint20preview.Label6.Caption
    ElseIf FormPrint20preview.fontsizetemp = 7 Then
        Text1.Text = FormPrint20preview.Label7.Caption
    ElseIf FormPrint20preview.fontsizetemp = 8 Then
        Text1.Text = FormPrint20preview.Label8.Caption
    ElseIf FormPrint20preview.fontsizetemp = 9 Then
        Text1.Text = FormPrint20preview.Label9.Caption
    ElseIf FormPrint20preview.fontsizetemp = 10 Then
        Text1.Text = FormPrint20preview.Label10.Caption
    ElseIf FormPrint20preview.fontsizetemp = 11 Then
        Text1.Text = FormPrint20preview.Label11.Caption
    ElseIf FormPrint20preview.fontsizetemp = 12 Then
        Text1.Text = FormPrint20preview.Label12.Caption
    ElseIf FormPrint20preview.fontsizetemp = 13 Then
        Text1.Text = FormPrint20preview.Label13.Caption
    ElseIf FormPrint20preview.fontsizetemp = 14 Then
        Text1.Text = FormPrint20preview.Label14.Caption
    ElseIf FormPrint20preview.fontsizetemp = 15 Then
        Text1.Text = FormPrint20preview.Label15.Caption
    ElseIf FormPrint20preview.fontsizetemp = 16 Then
        Text1.Text = FormPrint20preview.Label16.Caption
    ElseIf FormPrint20preview.fontsizetemp = 17 Then
        Text1.Text = FormPrint20preview.Label17.Caption
    ElseIf FormPrint20preview.fontsizetemp = 18 Then
        Text1.Text = FormPrint20preview.Label18.Caption
    ElseIf FormPrint20preview.fontsizetemp = 19 Then
        Text1.Text = FormPrint20preview.Label19.Caption
    ElseIf FormPrint20preview.fontsizetemp = 20 Then
        Text1.Text = FormPrint20preview.Label20.Caption
        
    ElseIf FormPrint20preview.fontsizetemp = 21 Then
        Text1.Text = FormPrint20preview.Label21.Caption
    ElseIf FormPrint20preview.fontsizetemp = 22 Then
        Text1.Text = FormPrint20preview.Label22.Caption
    ElseIf FormPrint20preview.fontsizetemp = 23 Then
        Text1.Text = FormPrint20preview.Label23.Caption
    ElseIf FormPrint20preview.fontsizetemp = 24 Then
        Text1.Text = FormPrint20preview.Label24.Caption
    ElseIf FormPrint20preview.fontsizetemp = 25 Then
        Text1.Text = FormPrint20preview.Label25.Caption
    ElseIf FormPrint20preview.fontsizetemp = 26 Then
        Text1.Text = FormPrint20preview.Label26.Caption
    ElseIf FormPrint20preview.fontsizetemp = 27 Then
        Text1.Text = FormPrint20preview.Label27.Caption
    ElseIf FormPrint20preview.fontsizetemp = 28 Then
        Text1.Text = FormPrint20preview.Label28.Caption
    ElseIf FormPrint20preview.fontsizetemp = 29 Then
        Text1.Text = FormPrint20preview.Label29.Caption
    ElseIf FormPrint20preview.fontsizetemp = 30 Then
        Text1.Text = FormPrint20preview.Label30.Caption
        
    ElseIf FormPrint20preview.fontsizetemp = 31 Then
        Text1.Text = FormPrint20preview.Label31.Caption
    ElseIf FormPrint20preview.fontsizetemp = 32 Then
        Text1.Text = FormPrint20preview.Label32.Caption
    ElseIf FormPrint20preview.fontsizetemp = 33 Then
        Text1.Text = FormPrint20preview.Label33.Caption
    ElseIf FormPrint20preview.fontsizetemp = 34 Then
        Text1.Text = FormPrint20preview.Label34.Caption
    ElseIf FormPrint20preview.fontsizetemp = 35 Then
        Text1.Text = FormPrint20preview.Label35.Caption
    ElseIf FormPrint20preview.fontsizetemp = 36 Then
        Text1.Text = FormPrint20preview.Label36.Caption
    ElseIf FormPrint20preview.fontsizetemp = 37 Then
        Text1.Text = FormPrint20preview.Label37.Caption
    ElseIf FormPrint20preview.fontsizetemp = 38 Then
        Text1.Text = FormPrint20preview.Label38.Caption
    ElseIf FormPrint20preview.fontsizetemp = 39 Then
        Text1.Text = FormPrint20preview.Label39.Caption
    ElseIf FormPrint20preview.fontsizetemp = 40 Then
       Text1.Text = FormPrint20preview.Label40.Caption
        
    ElseIf FormPrint20preview.fontsizetemp = 41 Then
        Text1.Text = FormPrint20preview.Label41.Caption
    ElseIf FormPrint20preview.fontsizetemp = 42 Then
        Text1.Text = FormPrint20preview.Label42.Caption
    ElseIf FormPrint20preview.fontsizetemp = 43 Then
        Text1.Text = FormPrint20preview.Label43.Caption
    ElseIf FormPrint20preview.fontsizetemp = 44 Then
        Text1.Text = FormPrint20preview.Label44.Caption
    ElseIf FormPrint20preview.fontsizetemp = 45 Then
        Text1.Text = FormPrint20preview.Label45.Caption
    ElseIf FormPrint20preview.fontsizetemp = 46 Then
        Text1.Text = FormPrint20preview.Label46.Caption
    ElseIf FormPrint20preview.fontsizetemp = 47 Then
        Text1.Text = FormPrint20preview.Label47.Caption
    ElseIf FormPrint20preview.fontsizetemp = 48 Then
        Text1.Text = FormPrint20preview.Label48.Caption
    ElseIf FormPrint20preview.fontsizetemp = 49 Then
        Text1.Text = FormPrint20preview.Label49.Caption
    ElseIf FormPrint20preview.fontsizetemp = 50 Then
        Text1.Text = FormPrint20preview.Label50.Caption
        
    ElseIf FormPrint20preview.fontsizetemp = 51 Then
        Text1.Text = FormPrint20preview.Label51.Caption
    ElseIf FormPrint20preview.fontsizetemp = 52 Then
        Text1.Text = FormPrint20preview.Label52.Caption
    ElseIf FormPrint20preview.fontsizetemp = 53 Then
        Text1.Text = FormPrint20preview.Label53.Caption
    ElseIf FormPrint20preview.fontsizetemp = 54 Then
        Text1.Text = FormPrint20preview.Label54.Caption
    ElseIf FormPrint20preview.fontsizetemp = 55 Then
        Text1.Text = FormPrint20preview.Label55.Caption
    ElseIf FormPrint20preview.fontsizetemp = 56 Then
        Text1.Text = FormPrint20preview.Label56.Caption
    ElseIf FormPrint20preview.fontsizetemp = 57 Then
        Text1.Text = FormPrint20preview.Label57.Caption
    ElseIf FormPrint20preview.fontsizetemp = 58 Then
        Text1.Text = FormPrint20preview.Label58.Caption
    ElseIf FormPrint20preview.fontsizetemp = 59 Then
        Text1.Text = FormPrint20preview.Label59.Caption
    ElseIf FormPrint20preview.fontsizetemp = 60 Then
        Text1.Text = FormPrint20preview.Label60.Caption
        
    ElseIf FormPrint20preview.fontsizetemp = 61 Then
        Text1.Text = FormPrint20preview.Label61.Caption
    ElseIf FormPrint20preview.fontsizetemp = 62 Then
        Text1.Text = FormPrint20preview.Label62.Caption
    ElseIf FormPrint20preview.fontsizetemp = 63 Then
        Text1.Text = FormPrint20preview.Label63.Caption
    ElseIf FormPrint20preview.fontsizetemp = 64 Then
        Text1.Text = FormPrint20preview.Label64.Caption
    ElseIf FormPrint20preview.fontsizetemp = 65 Then
        Text1.Text = FormPrint20preview.Label65.Caption
    ElseIf FormPrint20preview.fontsizetemp = 66 Then
        Text1.Text = FormPrint20preview.Label66.Caption
    ElseIf FormPrint20preview.fontsizetemp = 67 Then
        Text1.Text = FormPrint20preview.Label67.Caption
    ElseIf FormPrint20preview.fontsizetemp = 68 Then
        Text1.Text = FormPrint20preview.Label68.Caption
    ElseIf FormPrint20preview.fontsizetemp = 69 Then
        Text1.Text = FormPrint20preview.Label69.Caption
    ElseIf FormPrint20preview.fontsizetemp = 70 Then
        Text1.Text = FormPrint20preview.Label70.Caption
        
    ElseIf FormPrint20preview.fontsizetemp = 71 Then
        Text1.Text = FormPrint20preview.Label71.Caption
    ElseIf FormPrint20preview.fontsizetemp = 72 Then
        Text1.Text = FormPrint20preview.Label72.Caption
    ElseIf FormPrint20preview.fontsizetemp = 73 Then
        Text1.Text = FormPrint20preview.Label73.Caption
    ElseIf FormPrint20preview.fontsizetemp = 74 Then
        Text1.Text = FormPrint20preview.Label74.Caption
    ElseIf FormPrint20preview.fontsizetemp = 75 Then
        Text1.Text = FormPrint20preview.Label75.Caption
    ElseIf FormPrint20preview.fontsizetemp = 76 Then
        Text1.Text = FormPrint20preview.Label76.Caption
    ElseIf FormPrint20preview.fontsizetemp = 77 Then
        Text1.Text = FormPrint20preview.Label77.Caption
    ElseIf FormPrint20preview.fontsizetemp = 78 Then
        Text1.Text = FormPrint20preview.Label78.Caption
    ElseIf FormPrint20preview.fontsizetemp = 79 Then
        Text1.Text = FormPrint20preview.Label79.Caption
    ElseIf FormPrint20preview.fontsizetemp = 80 Then
        Text1.Text = FormPrint20preview.Label80.Caption
    End If
    
    Text2.Text = ""
    
End Sub

Private Sub Option1_Click()
    If Option1.Value = True Then
        Text2.Visible = False
        Combo1.ListIndex = 3
        Combo1.Enabled = True
    End If
End Sub

Private Sub Option2_Click()
    If Option2.Value = True Then
        Text2.Visible = True
        Combo1.ListIndex = 0
        Combo1.Enabled = False
    End If
End Sub
Private Sub Command2_Click()
    Unload Me
End Sub

