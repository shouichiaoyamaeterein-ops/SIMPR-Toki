VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CisYM_6.0.ocx"
Begin VB.Form COK0510 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "買掛状況照会"
   ClientHeight    =   11010
   ClientLeft      =   1680
   ClientTop       =   4245
   ClientWidth     =   15375
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11010
   ScaleWidth      =   15375
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   5040
      Top             =   5715
      Visible         =   0   'False
      Width           =   5160
      _ExtentX        =   9102
      _ExtentY        =   1640
      BackColor       =   16761024
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderWidth     =   5
      Begin Cis3D_v60.CIS3D Message 
         Height          =   525
         Left            =   240
         Top             =   210
         Width           =   4680
         _ExtentX        =   8255
         _ExtentY        =   926
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "【抽出中】しばらくお待ち下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   2
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   8985
      Left            =   285
      ScaleHeight     =   8925
      ScaleWidth      =   14730
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1845
      Width           =   14790
      Begin vsFlexLib.vsFlexArray VSGrid1 
         DragIcon        =   "COK0510.frx":0000
         Height          =   8670
         Left            =   195
         TabIndex        =   5
         Top             =   150
         Width           =   14415
         _Version        =   196608
         _ExtentX        =   25426
         _ExtentY        =   15293
         _StockProps     =   228
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "NO     |契約先                    |自社工区|車型 |品番                          |予測金額          |実績金額          |"
         Rows            =   26
         Cols            =   8
         BackColor       =   12648447
         BackColorBkg    =   12632256
         HighLight       =   0
         RowHeightMin    =   330
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1290
      Left            =   270
      ScaleHeight     =   1230
      ScaleWidth      =   11985
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   465
      Width           =   12045
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   1140
         Left            =   6645
         Top             =   45
         Width           =   3390
         _ExtentX        =   5980
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "品番"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   135
            TabIndex        =   3
            Top             =   510
            Width           =   3210
            _ExtentX        =   5662
            _ExtentY        =   661
            cWildeCard      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   1140
         Left            =   5805
         Top             =   45
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "車型"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Syakei 
            Height          =   375
            Left            =   90
            TabIndex        =   2
            Top             =   510
            Width           =   645
            _ExtentX        =   1138
            _ExtentY        =   661
            cWildeCard      =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXX"
            MaxLength       =   4
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   1140
         Left            =   2040
         Top             =   45
         Width           =   3765
         _ExtentX        =   6641
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "契約先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   150
            TabIndex        =   1
            Top             =   510
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
         Begin Cis3D_v60.CIS3D H1lb_Torcd 
            Height          =   360
            Left            =   1140
            Top             =   510
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1140
         Index           =   0
         Left            =   75
         Top             =   45
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "対象年月"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   30
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   240
            TabIndex        =   0
            Top             =   495
            Width           =   1455
            _ExtentX        =   2328
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1140
         Left            =   10035
         Top             =   45
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   2011
         ForeColor       =   16711680
         Caption         =   "自社工区"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   6
         cPositionY      =   60
         Begin VB.ComboBox H1Comb_Jisya 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   285
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   4
            Top             =   510
            Width           =   1395
         End
      End
   End
   Begin Cis3D_v60.CIS3D UC_3D31 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "　　　【 買掛状況照会 】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   -2147483643
      cFont3DColor2   =   8421504
      cBoderWidth     =   3
      cAlingnment     =   1
      Begin VB.PictureBox Back_Over 
         Height          =   315
         Left            =   10245
         ScaleHeight     =   255
         ScaleWidth      =   1590
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   45
         Width           =   1650
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "明細"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   1
            Left            =   45
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   0
            Width           =   705
         End
         Begin VB.OptionButton Op_Overlap 
            BackColor       =   &H00C0C0C0&
            Caption         =   "集計"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   9
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   255
            Index           =   0
            Left            =   780
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   0
            Value           =   -1  'True
            Width           =   765
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D32 
         Height          =   255
         Left            =   7260
         Top             =   45
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   450
         BackColor       =   14737632
         Caption         =   "抽出件数[              ]"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cBoderStyle     =   2
         cPositionX      =   50
         Begin Cis3D_v60.CIS3D lb_Kensu 
            Height          =   285
            Left            =   1095
            Top             =   15
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   503
            BackColor       =   14737632
            ForeColor       =   0
            Caption         =   "##,### / ##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   14737632
            cBoderColor2    =   14737632
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   2
         End
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14415
         Top             =   90
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   450
         BackColor       =   14737632
         ForeColor       =   16711680
         Caption         =   "SIMPR-A"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   14737632
         cBoderColor2    =   14737632
         cFont3DColor1   =   8454143
         cFont3DColor2   =   16777215
         cFont3D         =   1
         cAlingnment     =   7
         cBoderStyle     =   2
      End
      Begin Cis3D_v60.CIS3D Head_Tanto 
         Height          =   315
         Left            =   11925
         Top             =   30
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   556
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cBoderStyle     =   1
         cPositionX      =   50
      End
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Left            =   360
      TabIndex        =   6
      Text            =   "D"
      Top             =   585
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Left            =   630
      Top             =   585
   End
   Begin VB.Timer Timer2 
      Left            =   1350
      Top             =   585
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   1290
      Left            =   12435
      ScaleHeight     =   1230
      ScaleWidth      =   2580
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   465
      Width           =   2640
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   1725
         Top             =   615
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   1
         BtnCaption      =   "終了"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16711680
         BtnShadowColor  =   16777215
         PFCaption       =   "( F12 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   12648447
         PFShadowColor   =   0
         PFBackColor     =   16711680
      End
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   885
         Top             =   615
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   2
         BtnCaption      =   "入力"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   128
         BtnShadowColor  =   16777215
         PFCaption       =   "(Enter)"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   16777215
         PFShadowColor   =   0
         PFBackColor     =   128
      End
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   45
         Top             =   615
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   3
         BtnCaption      =   "取消"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   32768
         BtnShadowColor  =   16777215
         PFCaption       =   "( Esc )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   65535
         PFShadowColor   =   0
         PFBackColor     =   49152
      End
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   45
         Top             =   30
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   11
         BtnCaption      =   "検索"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   33023
         BtnShadowColor  =   16777215
         PFCaption       =   "( F7 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   12648447
         PFShadowColor   =   0
         PFBackColor     =   33023
      End
      Begin CisBtn_60.CisBtn PB_DTL 
         Height          =   585
         Left            =   885
         Top             =   30
         Visible         =   0   'False
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "詳細"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   192
         BtnShadowColor  =   16777215
         PFCaption       =   "( F8 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   12648447
         PFShadowColor   =   0
         PFBackColor     =   192
      End
   End
End
Attribute VB_Name = "COK0510"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  買掛状況照会
'**       フォームID    :  COK0510
'**       処理概要      :
'**
'**       作  成  日    :  2008/06/18 By CIS
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim wObj            As Object       ' ｵﾌﾞｼﾞｪｸﾄﾜｰｸｴﾘｱ
    Dim Sl_Where        As String       ' SQL条件
    Dim Ttl_Kensu       As Integer      ' 表示件数
    Dim mDispMax        As Integer      ' ｸﾞﾘｯﾄﾞ全行数退避
        
    Dim CisComb1        As Cis_Comb.CisCombo
    
    Dim CisVsGrid       As New CisVsGrid3
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   # 二 重 起 動 防 止  #
'   #-------------------#
    If Not Dupli_Start Then End
'   #-------------------#
'   # 初 期 内 容 取 得  #
'   #-------------------#
    If Not IniGet Then End
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VSGrid1
        .NoSpaceCol = 1
        .NumberSet = No
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 7
        .InitGet = False
        .Init
    End With
    
    Call GridInit
    
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
    mDispMax = CisVsGrid.DispMax
    lb_Kensu = ""

    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
'
'   +------------------+
'   +   ｺﾝﾎﾞﾎﾞｯｸｽｾｯﾄ   +
'   +------------------+
    Set CisComb1 = New CisCombo
    With CisComb1
        Set .Connect = CisDB
        Set .CombObj = H1Comb_Jisya
        .Control = "名称マスタ"
        .ComboName = "自社工区"
        .Ess = False
        .IniPath = gIniExe
        .SetSpace = 50
        .ComboInit
    End With
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    Op_Overlap(0).Value = True
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    '********* システム日付
    iH1_YM = CisFun.Mid2(Format(Now(), "yyyyMMdd"), 1, 6)
    
    Call CisComb1.ComboSet

    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    If Not FormAct Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function
    
    Select Case Key_Code
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
        Case vbKeyF8:       If PB_DTL.Visible Then Call PB_DTL_Click    '【詳細】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
        Case vbKeyReturn
                            If PB_Ent.Visible Then                      '【入力】
                                Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                Call PB_ENT_Click
                                Dummy.Enabled = False
                            End If
        Case Else:          Exit Function
    End Select
    Key_Acc = True

End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
   Call CisDB.DBDISConnect
   Call CisFrm.UnLoadDisp
   End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
   Call ReturnPress
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_CAN_Click()
   Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        Call HeadClear
        iH1_YM.SetFocus
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B1" Then
        Call DispTotal(False)
        lb_Kensu = ""
        CisVsGrid.DispMax = mDispMax
        VSGrid1.Rows = mDispMax
        CisVsGrid.Clear
        Call GridInit
        ProcHB = "H1"
        Call DispChange(ProcHB)
        iH1_YM.SetFocus
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検索   (F7)             +
'+----------------------------+
Private Sub PB_Look_Click()
    If ProcHB = "H1" Then
       If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
       If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    End If
End Sub
'+-------------------+
'+     詳     細     +
'+-------------------+
Private Sub PB_DTL_Click()
    RV_YM = iH1_YM
    RV_KTorcd = iH1_Torcd
    Call CisComb1.CodeGet
    RV_Code = CisComb1.Code
    Rv_Call3 = "X"
    
    COK0520.Show vbModal
    Unload COK0520
    Set COK0520 = Nothing
    
End Sub
'+----------------------------+
'+      検         索         +
'+----------------------------+
Private Sub Look_Torcd()

    RV_TorKb = 5    ' 契約先
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
       iH1_Torcd = RV_TorcdK
       Call iH1_Torcd_LostFocus
    End If

End Sub
'+----------------------------+
'+       品番検索            +
'+----------------------------+
Private Sub Look_Hinbn()

    RV_Left = 0
    RV_Top = 0
    RVI_Hinmoku = 1
    
    CKK0030.Show vbModal
    Unload CKK0030
    Set CKK0030 = Nothing

    If RV_Rtn Then
       iH1_Hinbn = RV_Hinbn
    End If
End Sub
'+----------------------------------------+
'+  iH1_Torcd ( GotFocus )   契約先
'+----------------------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
'+----------------------------------------+
'+  iH1_Torcd ( LostFocus )  契約先
'+----------------------------------------+
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
       PB_Look.Visible = False
       PB_Look.Tag = ""
    End If
    If TorNmGet(iH1_Torcd, 5) Then
        H1lb_Torcd = TRM_RName
        If Trim(H1lb_Torcd) = "" Then
           H1lb_Torcd = TRM_Name
        End If
    Else
        H1lb_Torcd = ""
    End If
End Sub
'+-----------------------------+
'+  iH1_Hinbn
'+-----------------------------+
Private Sub iH1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+----------------------------------------+
'+      小計 表示･非表示 切替              +
'+----------------------------------------+
Private Sub Op_Overlap_Click(Index As Integer)
    Call CisFrm.MousePT(11)

    If ProcHB = "B1" Then
        CisVsGrid.DispMax = mDispMax
        VSGrid1.Rows = mDispMax
        CisVsGrid.Clear
        Call GridInit
        Call BodySet
    Else
        If Op_Overlap(0).Value Then
            Call DispTotal(True)
        Else
            Call DispTotal(False)
        End If
    End If

    On Error Resume Next
    If VSGrid1.Enabled Then
        VSGrid1.SetFocus
    Else
        iH1_YM.SetFocus
    End If
    On Error GoTo 0
    
    Call CisFrm.MousePT(1)
    
End Sub
'****************************
'*                          *
'*      HEAD/BODY 処理      *
'*                          *
'****************************
'****************************
'*      入力切替 処理        *
'****************************
Private Function DispChange(Pro As String)
    Dim H1Mode      As Boolean
    Dim H1Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    H1Mode = False
    H1Color = gPLostSel
    B1Mode = False
    B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1"
          H1Mode = True
          H1Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color

'    PB_CAN.Visible = Not H1Mode
    PB_Ent.Visible = H1Mode
    PB_DTL.Visible = B1Mode
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
   Call CisFrm.MousePT(11)
   If ProcHB = "H1" Then
      ProcHB = "B1"
      If Not Head1Chk Then
          ProcHB = "H1"
          GoTo ReturnPress_Ed
      End If
      Call DispChange(ProcHB)
      VSGrid1.SetFocus
      GoTo ReturnPress_Ed
   End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1初期化
'****************************
Private Sub HeadClear()
    For Each wObj In Me.Controls
        If wObj.Name Like "iH*" Or _
           wObj.Name Like "H1lb_*" Then wObj = ""
        If wObj.Name Like "Ch_*" Then wObj.Value = 0
    Next wObj
    Op_Overlap(0).Value = True
End Sub

'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
 
' 印刷年月
    If Trim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "     対象年月を入力して下さい      "
            .MB_MSG(3) = "   "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If
    If Not IsDate(iH1_YM.cYear & "/" & iH1_YM.cMonth & "/01") Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "        対象年月             "
            .MB_MSG(4) = "   正しい年月にて入力して下さい        "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
    End If

    If Not BodySet Then
       With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください        "
            .MB_Button = Error
            .MB_Title = "抽出処理"
            Call .MBOX
       End With
       PNL_Wait.Visible = False
       Exit Function
    End If
    Head1Chk = True
End Function
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'+==========================================================+
'+      グリッド処理                                        +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear ("NO")
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容)
        Call .FixedSet(中中, 右中, なし, 7, 0, "№")
        Call .FixedSet(中中, 左中, あり, 27, 1, "契約先")
        Call .FixedSet(中中, 左中, あり, 12, 2, "自社工区")
        Call .FixedSet(中中, 左中, あり, 6, 3, "車型")
        Call .FixedSet(中中, 左中, あり, 24, 4, "品番")
        Call .FixedSet(中中, 右中, なし, 18, 5, "予測金額")
        Call .FixedSet(中中, 右中, なし, 18, 6, "実績金額")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 7, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
        .InitDisp
    End With
End Sub
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet() As Boolean
    Dim ID          As Long
    Dim ID2         As Long
    Dim wChk        As Boolean
    Dim GridCol()   As Long
    Dim wCnt        As Integer
    Dim wStr        As String
    Dim XStr        As String

    BodySet = False

    On Error GoTo BodySet_Err:
    
    Sl_Where = "": gSL_Select = ""
' 表示条件 セット
    With CisFun
        .Where_CreateKB = NewDocument
        
        .Where_Create AlphaMe, "SUBSTRING(NY.入荷日, 1, 6)", iH1_YM
        ' 契約先
        .Where_Create AlphaMe, "HTM.契約先", iH1_Torcd
        ' 車型
        .Where_Create AlphaMe, "HN.車型", iH1_Syakei
        ' 品番
        .Where_Create AlphaMe, "NY.品番", iH1_Hinbn
        ' 自社工区
        Call CisComb1.CodeGet
        .Where_Create AlphaMe, "SY1.値", CisComb1.Code
    End With

    gSL_Select = ""
    gSL_Select = gSL_Select & " SELECT CASE WHEN ISNULL(TR.略称,'') <> ''" & vbCrLf
    gSL_Select = gSL_Select & "             THEN HTM.契約先 + ':' + TR.略称" & vbCrLf
    gSL_Select = gSL_Select & "             ELSE HTM.契約先 END 契約先名," & vbCrLf
    gSL_Select = gSL_Select & "        ISNULL(SY1.値名称, '') 自社工区名, NY.受入," & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(HN.表示品番,'') = '' THEN NY.品番 Else HN.表示品番 End 品番," & vbCrLf
    gSL_Select = gSL_Select & "        ISNULL(HN.車型, '') 車型," & vbCrLf
    gSL_Select = gSL_Select & "        SUM("
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KTN1.単価, 0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR.計算単位, 0) = 0 THEN   ROUND(NT.数量 * ISNULL(KTN1.単価, 0), 0, 1)" & vbCrLf   ' 切捨て
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 1 THEN CEILING(NT.数量 * ISNULL(KTN1.単価, 0))" & vbCrLf   ' 切上げ
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 2 THEN   ROUND(NT.数量 * ISNULL(KTN1.単価, 0), 0)" & vbCrLf   ' 四捨五入
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 3 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CASE WHEN ISNULL(TR.計算位1, 0) = 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KTN1.単価, 1, 1), 0, 1     )" & vbCrLf   '１：切捨て　  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KTN1.単価 * 10) / 10 , 0, 1)" & vbCrLf   '１：切捨て 　 ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KTN1.単価, 1)   , 0, 1     )" & vbCrLf   '１：切捨て 　 ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 1 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN CEILING(  ROUND(NT.数量 * KTN1.単価, 1, 1)           )" & vbCrLf   '１：切上げ  　２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN CEILING(CEILING(NT.数量 * KTN1.単価 * 10) / 10       )" & vbCrLf   '１：切上げ  　２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN CEILING(  ROUND(NT.数量 * KTN1.単価, 1)              )" & vbCrLf   '１：切上げ  　２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 2 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KTN1.単価, 1, 1), 0        )" & vbCrLf   '１：四捨五入  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KTN1.単価 * 10 ) / 10, 0   )" & vbCrLf   '１：四捨五入  ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KTN1.単価, 1   ), 0        )" & vbCrLf   '１：四捨五入  ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                  End" & vbCrLf
    gSL_Select = gSL_Select & "             END" & vbCrLf
    gSL_Select = gSL_Select & "             WHEN ISNULL(KTN2.単価, 0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR.計算単位, 0) = 0 THEN   ROUND(NT.数量 * ISNULL(KTN2.単価, 0), 0, 1)" & vbCrLf   ' 切捨て
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 1 THEN CEILING(NT.数量 * ISNULL(KTN2.単価, 0))" & vbCrLf   ' 切上げ
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 2 THEN   ROUND(NT.数量 * ISNULL(KTN2.単価, 0), 0)" & vbCrLf   ' 四捨五入
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 3 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CASE WHEN ISNULL(TR.計算位1, 0) = 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KTN2.単価, 1, 1), 0, 1     )" & vbCrLf   '１：切捨て　  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KTN2.単価 * 10) / 10 , 0, 1)" & vbCrLf   '１：切捨て 　 ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KTN2.単価, 1)   , 0, 1     )" & vbCrLf   '１：切捨て 　 ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 1 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN CEILING(  ROUND(NT.数量 * KTN2.単価, 1, 1)           )" & vbCrLf   '１：切上げ  　２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN CEILING(CEILING(NT.数量 * KTN2.単価 * 10) / 10       )" & vbCrLf   '１：切上げ  　２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN CEILING(  ROUND(NT.数量 * KTN2.単価, 1)              )" & vbCrLf   '１：切上げ  　２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 2 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KTN2.単価, 1, 1), 0        )" & vbCrLf   '１：四捨五入  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KTN2.単価 * 10 ) / 10, 0   )" & vbCrLf   '１：四捨五入  ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KTN2.単価, 1   ), 0        )" & vbCrLf   '１：四捨五入  ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                  End" & vbCrLf
    gSL_Select = gSL_Select & "             END" & vbCrLf
    gSL_Select = gSL_Select & "             WHEN ISNULL(KTN3.単価, 0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR.計算単位, 0) = 0 THEN   ROUND(NT.数量 * ISNULL(KTN3.単価, 0), 0, 1)" & vbCrLf   ' 切捨て
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 1 THEN CEILING(NT.数量 * ISNULL(KTN3.単価, 0))" & vbCrLf   ' 切上げ
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 2 THEN   ROUND(NT.数量 * ISNULL(KTN3.単価, 0), 0)" & vbCrLf   ' 四捨五入
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 3 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CASE WHEN ISNULL(TR.計算位1, 0) = 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KTN3.単価, 1, 1), 0, 1     )" & vbCrLf   '１：切捨て　  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KTN3.単価 * 10) / 10 , 0, 1)" & vbCrLf   '１：切捨て 　 ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KTN3.単価, 1)   , 0, 1     )" & vbCrLf   '１：切捨て 　 ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 1 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN CEILING(  ROUND(NT.数量 * KTN3.単価, 1, 1)           )" & vbCrLf   '１：切上げ  　２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN CEILING(CEILING(NT.数量 * KTN3.単価 * 10) / 10       )" & vbCrLf   '１：切上げ  　２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN CEILING(  ROUND(NT.数量 * KTN3.単価, 1)              )" & vbCrLf   '１：切上げ  　２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 2 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KTN3.単価, 1, 1), 0        )" & vbCrLf   '１：四捨五入  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KTN3.単価 * 10 ) / 10, 0   )" & vbCrLf   '１：四捨五入  ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KTN3.単価, 1   ), 0        )" & vbCrLf   '１：四捨五入  ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                  End" & vbCrLf
    gSL_Select = gSL_Select & "             END" & vbCrLf
    gSL_Select = gSL_Select & "             WHEN ISNULL(KTN4.単価, 0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR.計算単位, 0) = 0 THEN   ROUND(NT.数量 * ISNULL(KTN4.単価, 0), 0, 1)" & vbCrLf   ' 切捨て
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 1 THEN CEILING(NT.数量 * ISNULL(KTN4.単価, 0))" & vbCrLf   ' 切上げ
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 2 THEN   ROUND(NT.数量 * ISNULL(KTN4.単価, 0), 0)" & vbCrLf   ' 四捨五入
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 3 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CASE WHEN ISNULL(TR.計算位1, 0) = 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KTN4.単価, 1, 1), 0, 1     )" & vbCrLf   '１：切捨て　  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KTN4.単価 * 10) / 10 , 0, 1)" & vbCrLf   '１：切捨て 　 ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KTN4.単価, 1)   , 0, 1     )" & vbCrLf   '１：切捨て 　 ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 1 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN CEILING(  ROUND(NT.数量 * KTN4.単価, 1, 1)           )" & vbCrLf   '１：切上げ  　２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN CEILING(CEILING(NT.数量 * KTN4.単価 * 10) / 10       )" & vbCrLf   '１：切上げ  　２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN CEILING(  ROUND(NT.数量 * KTN4.単価, 1)              )" & vbCrLf   '１：切上げ  　２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 2 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KTN4.単価, 1, 1), 0        )" & vbCrLf   '１：四捨五入  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KTN4.単価 * 10 ) / 10, 0   )" & vbCrLf   '１：四捨五入  ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KTN4.単価, 1   ), 0        )" & vbCrLf   '１：四捨五入  ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                  End" & vbCrLf
    gSL_Select = gSL_Select & "             END" & vbCrLf
    gSL_Select = gSL_Select & "        ELSE 0 END) 予測金額," & vbCrLf
    '******* 入荷 *******
    gSL_Select = gSL_Select & "        SUM("
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(NT.納期, NTM.納期) <> NTM.納期 THEN 0 ELSE" & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT.単価, 0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR.計算単位, 0) = 0 THEN   ROUND(NY.入荷数 * ISNULL(KT.単価, 0), 0, 1)" & vbCrLf   ' 切捨て
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 1 THEN CEILING(NY.入荷数 * ISNULL(KT.単価, 0))" & vbCrLf   ' 切上げ
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 2 THEN   ROUND(NY.入荷数 * ISNULL(KT.単価, 0), 0)" & vbCrLf   ' 四捨五入
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 3 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CASE WHEN ISNULL(TR.計算位1, 0) = 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NY.入荷数 * KT.単価, 1, 1), 0, 1     )" & vbCrLf   '１：切捨て　  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NY.入荷数 * KT.単価 * 10) / 10 , 0, 1)" & vbCrLf   '１：切捨て 　 ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NY.入荷数 * KT.単価, 1)   , 0, 1     )" & vbCrLf   '１：切捨て 　 ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 1 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN CEILING(  ROUND(NY.入荷数 * KT.単価, 1, 1)           )" & vbCrLf   '１：切上げ  　２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN CEILING(CEILING(NY.入荷数 * KT.単価 * 10) / 10       )" & vbCrLf   '１：切上げ  　２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN CEILING(  ROUND(NY.入荷数 * KT.単価, 1)              )" & vbCrLf   '１：切上げ  　２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 2 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NY.入荷数 * KT.単価, 1, 1), 0        )" & vbCrLf   '１：四捨五入  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NY.入荷数 * KT.単価 * 10 ) / 10, 0   )" & vbCrLf   '１：四捨五入  ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NY.入荷数 * KT.単価, 1   ), 0        )" & vbCrLf   '１：四捨五入  ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                  End" & vbCrLf
    gSL_Select = gSL_Select & "             END" & vbCrLf
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT2.単価, 0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR.計算単位, 0) = 0 THEN   ROUND(NY.入荷数 * ISNULL(KT2.単価, 0), 0, 1)" & vbCrLf   ' 切捨て
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 1 THEN CEILING(NY.入荷数 * ISNULL(KT2.単価, 0))" & vbCrLf   ' 切上げ
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 2 THEN   ROUND(NY.入荷数 * ISNULL(KT2.単価, 0), 0)" & vbCrLf   ' 四捨五入
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 3 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CASE WHEN ISNULL(TR.計算位1, 0) = 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NY.入荷数 * KT2.単価, 1, 1), 0, 1     )" & vbCrLf   '１：切捨て　  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NY.入荷数 * KT2.単価 * 10) / 10 , 0, 1)" & vbCrLf   '１：切捨て 　 ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NY.入荷数 * KT2.単価, 1)   , 0, 1     )" & vbCrLf   '１：切捨て 　 ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 1 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN CEILING(  ROUND(NY.入荷数 * KT2.単価, 1, 1)           )" & vbCrLf   '１：切上げ  　２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN CEILING(CEILING(NY.入荷数 * KT2.単価 * 10) / 10       )" & vbCrLf   '１：切上げ  　２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN CEILING(  ROUND(NY.入荷数 * KT2.単価, 1)              )" & vbCrLf   '１：切上げ  　２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 2 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NY.入荷数 * KT2.単価, 1, 1), 0        )" & vbCrLf   '１：四捨五入  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NY.入荷数 * KT2.単価 * 10 ) / 10, 0   )" & vbCrLf   '１：四捨五入  ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NY.入荷数 * KT2.単価, 1   ), 0        )" & vbCrLf   '１：四捨五入  ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                  End" & vbCrLf
    gSL_Select = gSL_Select & "             END" & vbCrLf
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT3.単価, 0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR.計算単位, 0) = 0 THEN   ROUND(NY.入荷数 * ISNULL(KT3.単価, 0), 0, 1)" & vbCrLf   ' 切捨て
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 1 THEN CEILING(NY.入荷数 * ISNULL(KT3.単価, 0))" & vbCrLf   ' 切上げ
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 2 THEN   ROUND(NY.入荷数 * ISNULL(KT3.単価, 0), 0)" & vbCrLf   ' 四捨五入
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 3 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CASE WHEN ISNULL(TR.計算位1, 0) = 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NY.入荷数 * KT3.単価, 1, 1), 0, 1     )" & vbCrLf   '１：切捨て　  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NY.入荷数 * KT3.単価 * 10) / 10 , 0, 1)" & vbCrLf   '１：切捨て 　 ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NY.入荷数 * KT3.単価, 1)   , 0, 1     )" & vbCrLf   '１：切捨て 　 ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 1 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN CEILING(  ROUND(NY.入荷数 * KT3.単価, 1, 1)           )" & vbCrLf   '１：切上げ  　２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN CEILING(CEILING(NY.入荷数 * KT3.単価 * 10) / 10       )" & vbCrLf   '１：切上げ  　２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN CEILING(  ROUND(NY.入荷数 * KT3.単価, 1)              )" & vbCrLf   '１：切上げ  　２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 2 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NY.入荷数 * KT3.単価, 1, 1), 0        )" & vbCrLf   '１：四捨五入  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NY.入荷数 * KT3.単価 * 10 ) / 10, 0   )" & vbCrLf   '１：四捨五入  ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NY.入荷数 * KT3.単価, 1   ), 0        )" & vbCrLf   '１：四捨五入  ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                  End" & vbCrLf
    gSL_Select = gSL_Select & "             END" & vbCrLf
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR.計算単位, 0) = 0 THEN   ROUND(NY.入荷数 * ISNULL(KT4.単価, 0), 0, 1)" & vbCrLf   ' 切捨て
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 1 THEN CEILING(NY.入荷数 * ISNULL(KT4.単価, 0))" & vbCrLf   ' 切上げ
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 2 THEN   ROUND(NY.入荷数 * ISNULL(KT4.単価, 0), 0)" & vbCrLf   ' 四捨五入
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 3 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CASE WHEN ISNULL(TR.計算位1, 0) = 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NY.入荷数 * KT4.単価, 1, 1), 0, 1     )" & vbCrLf   '１：切捨て　  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NY.入荷数 * KT4.単価 * 10) / 10 , 0, 1)" & vbCrLf   '１：切捨て 　 ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NY.入荷数 * KT4.単価, 1)   , 0, 1     )" & vbCrLf   '１：切捨て 　 ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 1 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN CEILING(  ROUND(NY.入荷数 * KT4.単価, 1, 1)           )" & vbCrLf   '１：切上げ  　２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN CEILING(CEILING(NY.入荷数 * KT4.単価 * 10) / 10       )" & vbCrLf   '１：切上げ  　２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN CEILING(  ROUND(NY.入荷数 * KT4.単価, 1)              )" & vbCrLf   '１：切上げ  　２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 2 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NY.入荷数 * KT4.単価, 1, 1), 0        )" & vbCrLf   '１：四捨五入  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NY.入荷数 * KT4.単価 * 10 ) / 10, 0   )" & vbCrLf   '１：四捨五入  ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NY.入荷数 * KT4.単価, 1   ), 0        )" & vbCrLf   '１：四捨五入  ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                  End" & vbCrLf
    gSL_Select = gSL_Select & "             END" & vbCrLf
    gSL_Select = gSL_Select & "        ELSE 0 END END) 実績金額" & vbCrLf
    gSL_Select = gSL_Select & "   FROM 入荷テーブル NY" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HN" & vbCrLf
    gSL_Select = gSL_Select & "             ON HN.品番 = NY.品番" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HTM" & vbCrLf
    gSL_Select = gSL_Select & "             ON HTM.品番 = NY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.取引先 = NY.仕入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.受入 = NY.受入" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.品目 = 1" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR" & vbCrLf
    gSL_Select = gSL_Select & "             ON TR.取引先CD = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "            AND TR.取引先区分 = 1" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY1" & vbCrLf
    gSL_Select = gSL_Select & "             ON SY1.区分名称 = '自社工区'" & vbCrLf
    gSL_Select = gSL_Select & "            AND SY1.英数字1 = NY.受入" & vbCrLf
    gSL_Select = gSL_Select & "            AND SY1.区分タイプ = ''" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN (SELECT MAX(発注管理NO) 発注管理NO," & vbCrLf
    gSL_Select = gSL_Select & "                                品番, 仕入先, 受入"
    gSL_Select = gSL_Select & "                           FROM 入荷テーブル" & vbCrLf
    gSL_Select = gSL_Select & "                          GROUP BY 品番, 仕入先, 受入) NY2" & vbCrLf
    gSL_Select = gSL_Select & "             ON NY2.発注管理NO = NY.発注管理NO" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 発注内示テーブル NT" & vbCrLf
    gSL_Select = gSL_Select & "             ON SUBSTRING(NT.納期, 1, 6) = SUBSTRING(NY.入荷日, 1, 6)" & vbCrLf
    gSL_Select = gSL_Select & "            AND NT.品番 = NY2.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND NT.手配先 = NY2.仕入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND NT.受入 = NY2.受入" & vbCrLf
    gSL_Select = gSL_Select & "    LEFT OUTER JOIN (SELECT MAX(納期) 納期, 内示年月, 品番, 手配先, 受入" & vbCrLf
    gSL_Select = gSL_Select & "                       From 発注内示テーブル" & vbCrLf
    gSL_Select = gSL_Select & "                      GROUP BY 内示年月, 品番, 手配先, 受入) NTM" & vbCrLf
    gSL_Select = gSL_Select & "         ON NTM.内示年月 = SUBSTRING(NY.入荷日, 1, 6)" & vbCrLf
    gSL_Select = gSL_Select & "        AND NTM.品番 = NY2.品番" & vbCrLf
    gSL_Select = gSL_Select & "        AND NTM.手配先 = NY2.仕入先" & vbCrLf
    gSL_Select = gSL_Select & "        AND NTM.受入 = NY2.受入" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT" & vbCrLf
    gSL_Select = gSL_Select & "             ON KT.品番 = NY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT.取引先 = NY.仕入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT.契約先 = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT.履歴NO = (SELECT   MAX(KTS.履歴NO) FROM 購入単価マスタ KTS" & vbCrLf
    gSL_Select = gSL_Select & "                              WHERE   KTS.品番 = NY.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS.取引先 = NY.仕入先" & vbCrLf
    'gSL_Select = gSL_Select & "                                AND   KTS.契約先 = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS.適用開始日付 < NY.入荷日" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS.適用開始日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS.適用開始日付 = '')" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS.適用終了日付 > NY.入荷日" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS.適用終了日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS.適用終了日付 = ''))" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT2" & vbCrLf
    gSL_Select = gSL_Select & "             ON KT2.品番 = NY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT2.取引先 = NY.仕入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT2.契約先 = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT2.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT2.履歴NO = (SELECT   MAX(KTS2.履歴NO) FROM 購入単価マスタ KTS2" & vbCrLf
    gSL_Select = gSL_Select & "                              WHERE   KTS2.品番 = NY.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS2.取引先 = NY.仕入先" & vbCrLf
    'gSL_Select = gSL_Select & "                                AND   KTS2.契約先 = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS2.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS2.適用開始日付 < NY.入荷日" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS2.適用開始日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS2.適用開始日付 = '')" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS2.適用終了日付 > NY.入荷日" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS2.適用終了日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS2.適用終了日付 = ''))" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT3" & vbCrLf
    gSL_Select = gSL_Select & "             ON KT3.品番 = NY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT3.取引先 = NY.仕入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT3.契約先 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT3.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT3.履歴NO = (SELECT   MAX(KTS3.履歴NO) FROM 購入単価マスタ KTS3" & vbCrLf
    gSL_Select = gSL_Select & "                              WHERE   KTS3.品番 = NY.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS3.取引先 = NY.仕入先" & vbCrLf
    'gSL_Select = gSL_Select & "                                AND   KTS3.契約先 = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS3.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS3.適用開始日付 < NY.入荷日" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS3.適用開始日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS3.適用開始日付 = '')" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS3.適用終了日付 > NY.入荷日" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS3.適用終了日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS3.適用終了日付 = ''))" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT4" & vbCrLf
    gSL_Select = gSL_Select & "             ON KT4.品番 = NY.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT4.取引先 = NY.仕入先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT4.契約先 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT4.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT4.履歴NO = (SELECT   MAX(KTS4.履歴NO) FROM 購入単価マスタ KTS4" & vbCrLf
    gSL_Select = gSL_Select & "                              WHERE   KTS4.品番 = NY.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS4.取引先 = NY.仕入先" & vbCrLf
    'gSL_Select = gSL_Select & "                                AND   KTS4.契約先 = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS4.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS4.適用開始日付 < NY.入荷日" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS4.適用開始日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS4.適用開始日付 = '')" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS4.適用終了日付 > NY.入荷日" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS4.適用終了日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS4.適用終了日付 = ''))" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KTN1" & vbCrLf
    gSL_Select = gSL_Select & "             ON KTN1.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN1.取引先 = NT.手配先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN1.契約先 = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN1.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN1.履歴NO = (SELECT   MAX(KTSN1.履歴NO) FROM 購入単価マスタ KTSN1" & vbCrLf
    gSL_Select = gSL_Select & "                              WHERE   KTSN1.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN1.取引先 = NT.手配先" & vbCrLf
    'gSL_Select = gSL_Select & "                                AND   KTSN1.契約先 = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN1.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTSN1.適用開始日付 < NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN1.適用開始日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTSN1.適用開始日付 = '')" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTSN1.適用終了日付 > NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN1.適用終了日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTSN1.適用終了日付 = ''))" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KTN2" & vbCrLf
    gSL_Select = gSL_Select & "             ON KTN2.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN2.取引先 = NT.手配先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN2.契約先 = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN2.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN2.履歴NO = (SELECT   MAX(KTSN2.履歴NO) FROM 購入単価マスタ KTSN2" & vbCrLf
    gSL_Select = gSL_Select & "                              WHERE   KTSN2.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN2.取引先 = NT.手配先" & vbCrLf
    'gSL_Select = gSL_Select & "                                AND   KTSN2.契約先 = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN2.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTSN2.適用開始日付 < NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN2.適用開始日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTSN2.適用開始日付 = '')" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTSN2.適用終了日付 > NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN2.適用終了日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTSN2.適用終了日付 = ''))" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KTN3" & vbCrLf
    gSL_Select = gSL_Select & "             ON KTN3.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN3.取引先 = NT.手配先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN3.契約先 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN3.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN3.履歴NO = (SELECT   MAX(KTSN3.履歴NO) FROM 購入単価マスタ KTSN3" & vbCrLf
    gSL_Select = gSL_Select & "                              WHERE   KTSN3.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN3.取引先 = NT.手配先" & vbCrLf
    'gSL_Select = gSL_Select & "                                AND   KTSN3.契約先 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN3.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTSN3.適用開始日付 < NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN3.適用開始日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTSN3.適用開始日付 = '')" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTSN3.適用終了日付 > NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN3.適用終了日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTSN3.適用終了日付 = ''))" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KTN4" & vbCrLf
    gSL_Select = gSL_Select & "             ON KTN4.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN4.取引先 = NT.手配先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN4.契約先 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN4.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "            AND KTN4.履歴NO = (SELECT   MAX(KTSN4.履歴NO) FROM 購入単価マスタ KTSN4" & vbCrLf
    gSL_Select = gSL_Select & "                              WHERE   KTSN4.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN4.取引先 = NT.手配先" & vbCrLf
    'gSL_Select = gSL_Select & "                                AND   KTSN4.契約先 = ''" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN4.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTSN4.適用開始日付 < NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN4.適用開始日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTSN4.適用開始日付 = '')" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTSN4.適用終了日付 > NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTSN4.適用終了日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTSN4.適用終了日付 = ''))" & vbCrLf
' 抽出文　作成
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    
    gSL_Select = gSL_Select & "  GROUP BY TR.略称, HTM.契約先, SY1.値名称, NY.受入," & vbCrLf
    gSL_Select = gSL_Select & "           NY.品番, HN.表示品番, HN.車型, KT.単価,"
    gSL_Select = gSL_Select & "           TR.計算単位, KT2.単価, TR.計算位1, TR.計算位2" & vbCrLf
    
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 INSERT START

    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create AlphaMe, "SUBSTRING(NT.納期, 1, 6)", iH1_YM
         ' 契約先
         .Where_Create AlphaMe, "HTM.契約先", iH1_Torcd
         ' 車型
         .Where_Create AlphaMe, "HN.車型", iH1_Syakei
         ' 品番
         .Where_Create AlphaMe, "NT.品番", iH1_Hinbn
         ' 自社工区
         Call CisComb1.CodeGet
         .Where_Create AlphaMe, "SY1.値", CisComb1.Code
         ' 発注管理NO
         Call CisComb1.CodeGet
         .Where_Create Numeric, "ISNULL(発注管理NO, 0)", 0, , Create
    End With
    
    gSL_Select = gSL_Select & "  UNION"
    gSL_Select = gSL_Select & " SELECT CASE WHEN ISNULL(TR.略称,'') <> ''" & vbCrLf
    gSL_Select = gSL_Select & "             THEN HTM.契約先 + ':' + TR.略称" & vbCrLf
    gSL_Select = gSL_Select & "             ELSE HTM.契約先 END 契約先名," & vbCrLf
    gSL_Select = gSL_Select & "        ISNULL(SY1.値名称, '') 自社工区名, NT.受入," & vbCrLf
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(HN.表示品番,'') = '' THEN NT.品番 Else HN.表示品番 End 品番," & vbCrLf
    gSL_Select = gSL_Select & "        ISNULL(HN.車型, '') 車型," & vbCrLf
    gSL_Select = gSL_Select & "        SUM("
    gSL_Select = gSL_Select & "        CASE WHEN ISNULL(KT3.単価, 0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR.計算単位, 0) = 0 THEN   ROUND(NT.数量 * ISNULL(KT3.単価, 0), 0, 1)" & vbCrLf   ' 切捨て
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 1 THEN CEILING(NT.数量 * ISNULL(KT3.単価, 0))" & vbCrLf   ' 切上げ
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 2 THEN   ROUND(NT.数量 * ISNULL(KT3.単価, 0), 0)" & vbCrLf   ' 四捨五入
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 3 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CASE WHEN ISNULL(TR.計算位1, 0) = 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KT3.単価, 1, 1), 0, 1     )" & vbCrLf   '１：切捨て　  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KT3.単価 * 10) / 10 , 0, 1)" & vbCrLf   '１：切捨て 　 ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KT3.単価, 1)   , 0, 1     )" & vbCrLf   '１：切捨て 　 ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 1 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN CEILING(  ROUND(NT.数量 * KT3.単価, 1, 1)           )" & vbCrLf   '１：切上げ  　２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN CEILING(CEILING(NT.数量 * KT3.単価 * 10) / 10       )" & vbCrLf   '１：切上げ  　２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN CEILING(  ROUND(NT.数量 * KT3.単価, 1)              )" & vbCrLf   '１：切上げ  　２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 2 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KT3.単価, 1, 1), 0        )" & vbCrLf   '１：四捨五入  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KT3.単価 * 10 ) / 10, 0   )" & vbCrLf   '１：四捨五入  ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KT3.単価, 1   ), 0        )" & vbCrLf   '１：四捨五入  ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                  End" & vbCrLf
    gSL_Select = gSL_Select & "             END" & vbCrLf
    gSL_Select = gSL_Select & "             WHEN ISNULL(KT4.単価, 0) <> 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "             CASE WHEN ISNULL(TR.計算単位, 0) = 0 THEN   ROUND(NT.数量 * ISNULL(KT4.単価, 0), 0, 1)" & vbCrLf   ' 切捨て
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 1 THEN CEILING(NT.数量 * ISNULL(KT4.単価, 0))" & vbCrLf   ' 切上げ
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 2 THEN   ROUND(NT.数量 * ISNULL(KT4.単価, 0), 0)" & vbCrLf   ' 四捨五入
    gSL_Select = gSL_Select & "                  WHEN ISNULL(TR.計算単位, 0) = 3 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                  CASE WHEN ISNULL(TR.計算位1, 0) = 0 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KT4.単価, 1, 1), 0, 1     )" & vbCrLf   '１：切捨て　  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KT4.単価 * 10) / 10 , 0, 1)" & vbCrLf   '１：切捨て 　 ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KT4.単価, 1)   , 0, 1     )" & vbCrLf   '１：切捨て 　 ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 1 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN CEILING(  ROUND(NT.数量 * KT4.単価, 1, 1)           )" & vbCrLf   '１：切上げ  　２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN CEILING(CEILING(NT.数量 * KT4.単価 * 10) / 10       )" & vbCrLf   '１：切上げ  　２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN CEILING(  ROUND(NT.数量 * KT4.単価, 1)              )" & vbCrLf   '１：切上げ  　２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                       WHEN ISNULL(TR.計算位1, 0) = 2 THEN" & vbCrLf
    gSL_Select = gSL_Select & "                       CASE WHEN ISNULL(TR.計算位2, 0) = 0 THEN   ROUND(  ROUND(NT.数量 * KT4.単価, 1, 1), 0        )" & vbCrLf   '１：四捨五入  ２：切捨て
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 1 THEN   ROUND(CEILING(NT.数量 * KT4.単価 * 10 ) / 10, 0   )" & vbCrLf   '１：四捨五入  ２：切上げ
    gSL_Select = gSL_Select & "                            WHEN ISNULL(TR.計算位2, 0) = 2 THEN   ROUND(  ROUND(NT.数量 * KT4.単価, 1   ), 0        )" & vbCrLf   '１：四捨五入  ２：四捨五入
    gSL_Select = gSL_Select & "                       End" & vbCrLf
    gSL_Select = gSL_Select & "                  End" & vbCrLf
    gSL_Select = gSL_Select & "             END" & vbCrLf
    gSL_Select = gSL_Select & "        ELSE 0 END) 予測金額, 0 実績金額"
    gSL_Select = gSL_Select & "   FROM 発注内示テーブル NT"
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 入荷テーブル NY" & vbCrLf
    gSL_Select = gSL_Select & "             ON NY.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND NY.仕入先 = NT.手配先" & vbCrLf
    gSL_Select = gSL_Select & "            AND NY.受入 = NT.受入" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番マスタ HN" & vbCrLf
    gSL_Select = gSL_Select & "             ON HN.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 品番取引先マスタ HTM" & vbCrLf
    gSL_Select = gSL_Select & "             ON HTM.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.取引先 = NT.手配先" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.受入 = NT.受入" & vbCrLf
    gSL_Select = gSL_Select & "            AND HTM.品目 = 1" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 取引先マスタ TR" & vbCrLf
    gSL_Select = gSL_Select & "             ON TR.取引先CD = HTM.契約先" & vbCrLf
    gSL_Select = gSL_Select & "            AND TR.取引先区分 = 0" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 名称マスタ SY1" & vbCrLf
    gSL_Select = gSL_Select & "             ON SY1.区分名称 = '自社工区'" & vbCrLf
    gSL_Select = gSL_Select & "            AND SY1.英数字1 = NT.受入" & vbCrLf
    gSL_Select = gSL_Select & "            AND SY1.区分タイプ = ''" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT3" & vbCrLf
    gSL_Select = gSL_Select & "             ON KT3.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT3.取引先 = NT.手配先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT3.仮単価区分 = ''" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT3.履歴NO = (SELECT   MAX(KTS3.履歴NO) FROM 購入単価マスタ KTS3" & vbCrLf
    gSL_Select = gSL_Select & "                              WHERE   KTS3.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS3.取引先 = NT.手配先" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS3.適用開始日付 < NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS3.適用開始日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS3.適用開始日付 = '')" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS3.適用終了日付 > NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS3.適用終了日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS3.適用終了日付 = ''))" & vbCrLf
    gSL_Select = gSL_Select & "        LEFT OUTER JOIN 購入単価マスタ KT4" & vbCrLf
    gSL_Select = gSL_Select & "             ON KT4.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT4.取引先 = NT.手配先" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT4.仮単価区分 = '1'" & vbCrLf
    gSL_Select = gSL_Select & "            AND KT4.履歴NO = (SELECT   MAX(KTS4.履歴NO) FROM 購入単価マスタ KTS4" & vbCrLf
    gSL_Select = gSL_Select & "                              WHERE   KTS4.品番 = NT.品番" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS4.取引先 = NT.手配先" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS4.適用開始日付 < NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS4.適用開始日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS4.適用開始日付 = '')" & vbCrLf
    gSL_Select = gSL_Select & "                                AND ((KTS4.適用終了日付 > NT.納期" & vbCrLf
    gSL_Select = gSL_Select & "                                AND   KTS4.適用終了日付 <> '')" & vbCrLf
    gSL_Select = gSL_Select & "                                 OR   KTS4.適用終了日付 = ''))" & vbCrLf
' 抽出文　作成
    gSL_Select = gSL_Select & CisFun.Where_Phrase
    gSL_Select = gSL_Select & ""
    
    gSL_Select = gSL_Select & "  GROUP BY TR.略称, HTM.契約先, SY1.値名称, NT.受入," & vbCrLf
    gSL_Select = gSL_Select & "           NT.品番, HN.表示品番, HN.車型"
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/06/12 INSERT END
    
    gSL_Select = gSL_Select & " ORDER BY 契約先名, 受入, 車型, 品番, 予測金額, 実績金額" & vbCrLf
    
    CisFun.ClipPut gSL_Select, False
    
    If Not NYTRead(gSL_Select) Then GoTo BodySet_ED1

    Ttl_Kensu = CisDB.RecordCount
    lb_Kensu.Caption = Format(Ttl_Kensu, "#,###,##0")
    Back_Over.Enabled = False    ' 分類 入力禁止

    PNL_Wait.Visible = True      ' 抽出中ﾒｯｾｰｼﾞ表示
    DoEvents

' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VSGrid1
        .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(ID)
        Next ID
    End With
    ID = 1
    VSGrid1.Rows = CisDB.RecordCount + 1

    Do Until Not NYT_RDSTS
        With VSGrid1
            If Op_Overlap(1).Value Then
                .TextMatrix(ID, GridCol(0)) = CisFun.RSetFld(ID, 3, "###")
            End If
            If Trim(NYT.契約先名) = "" Then
                .TextMatrix(ID, GridCol(1)) = "    "
            Else
                .TextMatrix(ID, GridCol(1)) = NYT.契約先名
            End If
            If Trim(NYT.自社工区名) = "" Then
                .TextMatrix(ID, GridCol(2)) = "    "
            Else
                .TextMatrix(ID, GridCol(2)) = NYT.自社工区名
            End If
            If Trim(NYT.車型) = "" Then
                .TextMatrix(ID, GridCol(3)) = "    "
            Else
                .TextMatrix(ID, GridCol(3)) = NYT.車型
            End If
            If Trim(NYT.品番) = "" Then
                .TextMatrix(ID, GridCol(4)) = "    "
            Else
                .TextMatrix(ID, GridCol(4)) = NYT.品番
            End If
            If CisFun.Val2(NYT.予測金額) = 0 Then
                .TextMatrix(ID, GridCol(5)) = ""
            Else
                .TextMatrix(ID, GridCol(5)) = Format(NYT.予測金額, "#,###")
            End If
            If CisFun.Val2(NYT.実績金額) = 0 Then
                .TextMatrix(ID, GridCol(6)) = ""
            Else
                .TextMatrix(ID, GridCol(6)) = Format(NYT.実績金額, "#,###")
            End If
            
            ID = ID + 1
        End With
        Call NYTReadNext
    Loop
    Call NYTClose

    CisVsGrid.DispMax = VSGrid1.Rows

BodySet_ED:

    Call GridInit
    If Op_Overlap(0).Value Then
        Call DispTotal(True)
    Else
        Call DispTotal(False)
    End If
    PNL_Wait.Visible = False
    With VSGrid1
         .Col = 1: .ColSel = .Cols - 1
         .Row = 1: .Redraw = True
    End With
    BodySet = True

BodySet_ED1:
    Back_Over.Enabled = True    ' 分類 入力禁止 解除
    PNL_Wait.Visible = False   ' 抽出中ﾒｯｾｰｼﾞ非表示
    Exit Function
BodySet_Err:
     CisFun.ErrorBox ("BodySet")
End Function
'+----------------------------------------+
'+      数量・金額 小計設定                 +
'+----------------------------------------+
Private Sub DispTotal(Disp As Boolean)
    Static FlexSt   As Integer
    Static FixedCol As Long
    
    With VSGrid1
        Call CisVsGrid.VSSort(True)
        
        .SubtotalPosition = flexSTAbove
        FlexSt = flexSTClear
        If Not Disp Or ProcHB = "H1" Then
            .OutlineBar = flexOutlineBarNone
            .Outline -1
            .Subtotal flexSTClear
        Else
            .Subtotal flexSTClear
            FlexSt = flexSTSum
        End If
        
        ' 数量・型費・印額　合計
        For gLong = 5 To 6
            FixedCol = CisVsGrid.FixedGet(gLong)
            .Subtotal FlexSt, 4, FixedCol, ",", &HC0FFC0, , False, " %s", 1
            .Subtotal FlexSt, 3, FixedCol, ",", &HC0C0FF, , False, " %s", 1
            .Subtotal FlexSt, 2, FixedCol, ",", &HC0FFC0, , False, " %s", 1
            .Subtotal FlexSt, 1, FixedCol, ",", &HC0C0FF, , False, " %s", 1
            .Subtotal FlexSt, -1, FixedCol, ",", &H80000008, &HFFFFFF, False, "総計"
        Next gLong

        If Not Disp Or ProcHB = "H1" Then
            .Outline -1
        Else
            .OutlineBar = flexOutlineBarComplete
            .Outline 2
        End If
    End With

'
'---------------------------------------------< 2004.11.09 MNT START >
'
'    If Not Disp Then
'        CisVsGrid.DispMax = mDispMax
'        VSGrid1.Rows = mDispMax
'    End If
'
'---------------------------------------------< 2004.11.09 MNT  END  >
'
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
' Button : 1 = 左ﾎﾞﾀﾝ ,2 = 右ﾎﾞﾀﾝ
    With VSGrid1
        .Tag = ""
        If Button <> 1 Then Exit Sub
    ' 固定行以外は行わない
        If .MouseRow <> 0 Then Exit Sub
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
    ' 数量・金額は行わない
        If .MouseCol >= 5 Then Exit Sub
    ' 移動元列をタグへ退避
        .Tag = Str(.MouseCol)
        .Drag 1
    End With
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)

    With VSGrid1
    ' 固定列は行わない
        If .MouseCol = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
    ' 数量・金額より前のみ移動可能
        If .MouseCol >= 5 Then Exit Sub
' ｸﾞﾘｯﾄﾞ表示 禁止
        Call CisFrm.MousePT(11)

        .Redraw = False
' 列移動
        .ColPosition(CisFun.Val2(.Tag)) = .MouseCol
' 列分類
        .Col = 1
        .ColSel = .Cols - 1
        '.Sort = GridSortAsc
' 合計表示・非表示
        If Op_Overlap(0).Value Then
            Call DispTotal(True)
            If .MouseCol > 0 And .MouseCol < 5 Then
                VSGrid1.Outline .MouseCol
            End If
        Else
            Call DispTotal(False)
        End If
' ｸﾞﾘｯﾄﾞ表示 解除
        .Redraw = True

       Call CisFrm.MousePT(1)

    End With
End Sub


