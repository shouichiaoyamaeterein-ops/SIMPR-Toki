VERSION 5.00
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Begin VB.Form CSJ0200 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "製造指示入力"
   ClientHeight    =   11190
   ClientLeft      =   120
   ClientTop       =   3480
   ClientWidth     =   15360
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   Begin VB.PictureBox DebugPicture 
      Height          =   300
      Left            =   10680
      ScaleHeight     =   240
      ScaleWidth      =   840
      TabIndex        =   17
      Top             =   960
      Visible         =   0   'False
      Width           =   900
      Begin VB.CheckBox DebugCheck1 
         Caption         =   "Debug"
         Height          =   180
         Left            =   30
         TabIndex        =   18
         Top             =   30
         Width           =   975
      End
   End
   Begin Cis3D_v60.CIS3D MessageGet 
      Height          =   870
      Left            =   4065
      Top             =   5070
      Visible         =   0   'False
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   1535
      BackColor       =   16744576
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      Begin Cis3D_v60.CIS3D LBL_Message 
         Height          =   720
         Left            =   75
         Top             =   75
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   1270
         BackColor       =   16777152
         Caption         =   "【 抽 出 中 】 しばらくお待ち下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HG創英角ﾎﾟｯﾌﾟ体"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin Cis3D_v60.CIS3D Message 
      Height          =   870
      Left            =   4065
      Top             =   5145
      Visible         =   0   'False
      Width           =   6870
      _ExtentX        =   12118
      _ExtentY        =   1535
      BackColor       =   16744576
      Caption         =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   720
         Left            =   75
         Top             =   75
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   1270
         BackColor       =   16777152
         Caption         =   "【 処 理 中 】 しばらくお待ち下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "HG創英角ﾎﾟｯﾌﾟ体"
            Size            =   18
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   345
      Index           =   1
      Left            =   7695
      ScaleHeight     =   345
      ScaleWidth      =   3900
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   30
      Visible         =   0   'False
      Width           =   3900
      Begin Cis3D_v60.CIS3D LBL_D_Kensu 
         Height          =   345
         Left            =   0
         Top             =   0
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "表示件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   20
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   285
            Left            =   990
            Top             =   30
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            BackColor       =   12648384
            ForeColor       =   0
            Caption         =   "##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D7 
         Height          =   345
         Left            =   1935
         Top             =   0
         Width           =   1965
         _ExtentX        =   3466
         _ExtentY        =   609
         BackColor       =   16777152
         ForeColor       =   0
         Caption         =   "変更件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   20
         Begin Cis3D_v60.CIS3D iT1_Mntsu 
            Height          =   285
            Left            =   990
            Top             =   45
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   503
            BackColor       =   16777152
            ForeColor       =   0
            Caption         =   "##,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -30
         End
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   825
      Left            =   11685
      ScaleHeight     =   765
      ScaleWidth      =   3540
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   960
      Width           =   3600
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   60
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   915
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
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
      Begin CisBtn_60.CisBtn PB_ENT 
         Height          =   585
         Left            =   1770
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
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
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   2625
         Top             =   90
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
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
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   7860
      Left            =   90
      ScaleHeight     =   7800
      ScaleWidth      =   15135
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2745
      Width           =   15195
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   7725
         Left            =   45
         Top             =   30
         Width           =   15060
         _ExtentX        =   26564
         _ExtentY        =   13626
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   9
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         Begin vsFlexLib.vsFlexArray VsGrid 
            Height          =   7665
            Left            =   30
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   30
            Width           =   15015
            _Version        =   196608
            _ExtentX        =   26485
            _ExtentY        =   13520
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
            FormatString    =   $"CSJ0200.frx":0000
            Rows            =   23
            Cols            =   33
            BackColor       =   12648384
            BackColorFixed  =   14737632
            BackColorSel    =   12648384
            ForeColorSel    =   -2147483640
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            GridLines       =   2
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Input_Area 
               Height          =   345
               Left            =   4890
               Top             =   330
               Width           =   1785
               _ExtentX        =   3149
               _ExtentY        =   609
               BackColor       =   4210752
               Caption         =   ""
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ Ｐゴシック"
                  Size            =   9
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               Begin CisText_V60.CisText iB1_Yusen 
                  Height          =   375
                  Left            =   0
                  TabIndex        =   8
                  Top             =   0
                  Width           =   600
                  _ExtentX        =   1058
                  _ExtentY        =   661
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "XXX"
                  MaxLength       =   3
                  cDataReplace    =   1
               End
               Begin CisText_V60.CisText iB1_Kakou 
                  Height          =   375
                  Left            =   585
                  TabIndex        =   9
                  Top             =   0
                  Width           =   1185
                  _ExtentX        =   2090
                  _ExtentY        =   661
                  cFormat         =   "#,###,###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "#,###,###"
                  MaxLength       =   9
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "#######"
                  cILength        =   7
                  cAlignment      =   1
               End
            End
         End
         Begin VB.TextBox B_TabDummy 
            Height          =   330
            Left            =   225
            TabIndex        =   10
            Text            =   "Text1"
            Top             =   120
            Width           =   240
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   5985
      Top             =   15
   End
   Begin VB.TextBox Dummy 
      Height          =   270
      IMEMode         =   3  'ｵﾌ固定
      Left            =   0
      TabIndex        =   12
      Text            =   "Dummy"
      Top             =   0
      Width           =   210
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00FF0000&
      Height          =   1260
      Left            =   90
      ScaleHeight     =   1200
      ScaleWidth      =   9300
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   900
      Width           =   9360
      Begin Cis3D_v60.CIS3D CIS3D16 
         Height          =   540
         Left            =   2640
         Top             =   60
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   953
         BackColor       =   16761087
         ForeColor       =   16711680
         Caption         =   "部署"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   900
            TabIndex        =   1
            Top             =   90
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   661
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1890
            Top             =   90
            Width           =   2490
            _ExtentX        =   4392
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
      Begin Cis3D_v60.CIS3D LBL_B_Kubun 
         Height          =   1080
         Left            =   7140
         Top             =   60
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   1905
         ForeColor       =   16711680
         Caption         =   "部品区分"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
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
         Begin VB.OptionButton OpH1_HinKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "仕掛品"
            Height          =   225
            Index           =   3
            Left            =   1110
            TabIndex        =   7
            Top             =   750
            Width           =   870
         End
         Begin VB.OptionButton OpH1_HinKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "購入品"
            Height          =   225
            Index           =   2
            Left            =   120
            TabIndex        =   6
            Top             =   750
            Width           =   870
         End
         Begin VB.OptionButton OpH1_HinKb 
            BackColor       =   &H00C0C0C0&
            Caption         =   "完成品"
            Height          =   225
            Index           =   1
            Left            =   1110
            TabIndex        =   5
            Top             =   420
            Width           =   870
         End
         Begin VB.OptionButton OpH1_HinKb 
            BackColor       =   &H00C0FFC0&
            Caption         =   "全て"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   420
            Value           =   -1  'True
            Width           =   870
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   540
         Left            =   60
         Top             =   60
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   953
         ForeColor       =   16711680
         Caption         =   "優先順位"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin CisText_V60.CisText iH1_Yusen 
            Height          =   375
            Left            =   1200
            TabIndex        =   0
            Top             =   90
            Width           =   480
            _ExtentX        =   847
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
            Text            =   "XXX"
            MaxLength       =   3
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   540
         Left            =   60
         Top             =   600
         Width           =   2580
         _ExtentX        =   4551
         _ExtentY        =   953
         ForeColor       =   16711680
         Caption         =   "機械"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin CisText_V60.CisText iH1_Kikai 
            Height          =   375
            Left            =   690
            TabIndex        =   2
            Top             =   90
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "XXX"
            MaxLength       =   3
         End
         Begin Cis3D_v60.CIS3D H1lb_Kikai 
            Height          =   360
            Left            =   1200
            Top             =   90
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1"
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
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   540
         Left            =   2640
         Top             =   600
         Width           =   4500
         _ExtentX        =   7938
         _ExtentY        =   953
         ForeColor       =   16711680
         Caption         =   "得意先"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin CisText_V60.CisText iH1_Tokui 
            Height          =   375
            Left            =   900
            TabIndex        =   3
            Top             =   90
            Width           =   990
            _ExtentX        =   1746
            _ExtentY        =   661
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
         Begin Cis3D_v60.CIS3D H1lb_Tokui 
            Height          =   360
            Left            =   1890
            Top             =   90
            Width           =   2490
            _ExtentX        =   4392
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
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【製造指示入力  】"
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
      cPositionX      =   1200
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14220
         Top             =   90
         Width           =   1005
         _ExtentX        =   1773
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
         Height          =   285
         Left            =   11640
         Top             =   60
         Width           =   2490
         _ExtentX        =   4392
         _ExtentY        =   503
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderWidth     =   3
         cAlingnment     =   7
      End
   End
End
Attribute VB_Name = "CSJ0200"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   製造指示入力
'**       フォームID    :   CSJ0200 (Based CXM2160)
'**       処理概要      :
'**
'**       作  成  日    :   2010/01/26  By CIS
'**       変  更  日    :   2010/07/05  By CIS
'**       変  更  日    :   2011/01/28  By CIS
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid               As New CisVsGrid3
    
    Dim ActObj                  As Object
    Dim ErrObj                  As Object
    Dim ActRow                  As Integer
    Dim SaveSYmd                As String
    Dim SaveEYmd                As String
    Dim SaveRow                 As Integer
    
    Dim KamiKI                  As Integer
    Dim SimoKI                  As Integer
    Dim wEtcTan                 As Currency
    
    Dim mIAreaLeft              As Long
    Dim DebugFlg                As Boolean      ' Debug
Private Sub DebugCheck1_Click()
    If ProcHB <> "H1" Then Exit Sub
    If DebugCheck1.Value = 0 Then DebugFlg = False Else DebugFlg = True
    Call GridInit
End Sub

'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    DebugFlg = False                ' Debug
    mIAreaLeft = Input_Area.Left
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
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 15     ' 21    ' 2010.01.26 upd
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
    
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    iH1_Torcd = "200"                           '2010/07/05
    Call iH1_Torcd_LostFocus                    '2010/07/05

    If FormAct Then Exit Sub
    FormAct = Not FormAct
   
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》                 +
'+-------------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+-------------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》         +
'+-------------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    
    Select Case Key_Code
           Case vbKeyF7:        If PB_LOOK.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                      '【入力】
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select
    
    Key_Acc = True
End Function

Private Sub OpH1_HinKb_Click(INDEX As Integer)
    Dim wCnt    As Integer
    For wCnt = 0 To 3
        OpH1_HinKb(wCnt).BackColor = LBL_B_Kubun.BackColor
    Next
    OpH1_HinKb(INDEX).BackColor = LBL_D_Kensu.BackColor
End Sub

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'    Call CisDB.DBDISConnect                            ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
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
   Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    
    If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_Yusen.SetFocus
      GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_LOOK.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_LOOK.Tag = "iH1_Kikai" Then Call Look_Kikai: Exit Sub
    If PB_LOOK.Tag = "iH1_Tokui" Then Call Look_Tokui: Exit Sub
End Sub

'+-----------------------------+
'+  iH1_Yusen
'+-----------------------------+
Private Sub iH1_Yusen_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = False
End Sub
Private Sub iH1_Yusen_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
End Sub

'+------------------------------+
'+       部署(取引先)検索       +
'+------------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 1                '   発注
    RV_TorKb2 = 1               '   社内

    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RV_Torcd
       Call iH1_Torcd_LostFocus
    End If
End Sub
'+------------------------------+
'+  iH1_Torcd                   +
'+------------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    H1lb_Tornm = ""
''''If TorNmGet(iH1_Torcd, 3, 1) Then
''''   H1lb_Tornm = TRM_RName
''''End If
    If Trim(iH1_Torcd) <> "" Then
        gSL_Select = "SELECT * FROM 取引先マスタ"
        gSL_Select = gSL_Select & " WHERE 取引先CD   = '" & Trim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   AND 取引先区分 = 1"
        gSL_Select = gSL_Select & "   AND 社内区分   = 1"
        If TRMRead(gSL_Select) Then
            H1lb_Tornm = TRM.正式名称
        Else
            H1lb_Tornm = ""
        End If
    End If
End Sub
'+------------------------------+
'+       機械検索               +
'+------------------------------+
Private Sub Look_Kikai()
    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "機械"
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
       iH1_Kikai = RV_Code          ' 値
       H1lb_Kikai = RV_Kubun        ' 値名称
    End If
End Sub
'+------------------------------+
'+  iH1_Kikai                   +
'+------------------------------+
Private Sub iH1_Kikai_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Kikai_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    H1lb_Kikai = ""
    If Trim(iH1_Kikai) <> "" Then
        gSL_Select = "SELECT * FROM 名称マスタ"
        gSL_Select = gSL_Select & " WHERE 区分名称 = '機械'"
        gSL_Select = gSL_Select & "   AND 値       = '" & Trim(iH1_Kikai) & "'"
        If SYMRead(gSL_Select) Then
            H1lb_Kikai = SYM.値名称
        Else
            H1lb_Kikai = ""
        End If
    End If
End Sub

'+------------------------------+
'+       得意先検索             +
'+------------------------------+
Private Sub Look_Tokui()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 3
    RV_Call = ""

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Tokui = RV_TorcdK
        H1lb_Tokui = RV_TorRName
    End If
End Sub
'+------------------------------+
'+  iH1_Tokui                   +
'+------------------------------+
Private Sub iH1_Tokui_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Tokui_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
    H1lb_Tokui = ""
    If TorNmGet(iH1_Tokui, 3, 1) Then
       H1lb_Tokui = TRM_RName
    End If
End Sub

'+-----------------------------+
'+  iB1_Yusen
'+-----------------------------+
Private Sub iB1_Yusen_GotFocus()
    Set ActObj = iB1_Yusen
    B_TabDummy.Tag = "E"
End Sub
Private Sub iB1_Yusen_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Yusen_LostFocus()
    If CisFun.Val2(iB1_Yusen) <> 0 Then
'        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 5) = iB1_Yusen
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 4) = iB1_Yusen
    Else
'        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 5) = ""
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 4) = ""
    End If
End Sub
'+-----------------------------+
'+  iB1_Kakou
'+-----------------------------+
Private Sub iB1_Kakou_GotFocus()
    Set ActObj = iB1_Kakou
End Sub
Private Sub iB1_Kakou_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Kakou_LostFocus()
    If RTrim(iB1_Kakou) <> "" Then
'        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = iB1_Kakou
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 5) = iB1_Kakou
    Else
'        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 6) = ""
        VsGrid.TextMatrix(CisFun.Val2(Input_Area.Tag), 5) = ""
    End If
End Sub
'+-----------------------------+
'+  B_TabDummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    If B_TabDummy.Tag = "S" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyUp)
    End If
    If B_TabDummy.Tag = "E" Then
        Call TextToVsGrid(Val(Input_Area.Tag))
        Call Grid_RowChange(vbKeyDown)
    End If
    SendKeys "{Tab}"
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   Call CisFrm.TimeSet(1)
End Sub
'*******************************
'*      処理区分名点滅制御      *
'*******************************
Private Sub Timer2_Timer()
   If K_Sykbnm.Caption <> "" Then
      K_Sykbnm.Caption = ""
   Else
      K_Sykbnm.Caption = SyoriNM
   End If
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
    PB_LOOK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    
    Call GridTextEnabled(B1Mode)
        
' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
    DebugCheck1.Enabled = H1Mode    ' Debug
    Input_Area.Left = mIAreaLeft
'
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
                .Enabled = True
            End If
            If .Name Like "Chk" & wbc_Enm & "_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid.Clear
    End If
    If HeadBody = "B1" Then
        CisVsGrid.Clear
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        VsGrid.Row = 1
        Call VsGridToText(VsGrid.Row)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        iH1_Torcd.SetFocus
        GoTo ReturnPress_Ed
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
    
    '=============================== 2010/07/05 Start
    If RTrim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     担当部署を入力して下さい        "
            .MB_Title = ""
            .MB_Button = OK
            .MBOX
        End With
        iH1_Torcd.SetFocus
        Exit Function
    End If
    '=============================== 2010/07/05 End
    
    If Not DBInput("INV") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     該当データがありません            "
            .MB_Title = ""
            .MB_Button = OK
            .MBOX
        End With
        iH1_Yusen.SetFocus
        Exit Function
    End If

    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call GridTextEnabled(False)
    
    '== グリッドデータ移動以外、特に処理なし
''ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
'    With CisFun
'        .MB_Lines = 4
'        .MB_Title = ""
'        .MB_Button = Error
'    End With
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    VsGrid.Row = gCnt1
    Call VsGridToText(VsGrid.Row)
    Call GridTextEnabled(True)
    ErrObj.SetFocus
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    Dim wStr1   As String
    Dim wStr2   As String
    Dim wCnt1   As Integer
    
    DBInput = False
    On Error GoTo DBInput_Err:
    Call HeadBodyClear("B1")
    
    wStr1 = "": wStr2 = " WHERE "
    If Trim(iH1_Yusen) <> "" Then
        wStr1 = wStr1 & wStr2 & "HKJ.優先順位 = '" & Trim(iH1_Yusen) & "'"
        wStr2 = "   AND "
    End If
    If Trim(iH1_Torcd) <> "" Then
        wStr1 = wStr1 & wStr2 & "HKJ.取引先   = '" & Trim(iH1_Torcd) & "'"
        wStr2 = "   AND "
    End If
    If Trim(iH1_Kikai) <> "" Then
        wStr1 = wStr1 & wStr2 & "HKJ.機械     = '" & Trim(iH1_Kikai) & "'"
        wStr2 = "   AND "
    End If
    If Trim(iH1_Tokui) <> "" Then
        wStr1 = wStr1 & wStr2 & "HM1.得意先   = '" & Trim(iH1_Tokui) & "'"
        wStr2 = "   AND "
    End If
    If OpH1_HinKb(0).Value Then
        ''''    If Not DebugFlg Then
        ''''wStr1 = wStr1 & wStr2 & "(HT1.部品区分 = 0 OR HT1.部品区分 = 1 OR HT1.部品区分 = 2)"
        ''''wStr2 = "   AND "
        ''''    End If
    Else
        For wCnt1 = 1 To 3
            If OpH1_HinKb(wCnt1).Value = True Then
                wStr1 = wStr1 & wStr2 & "ISNULL(HT1.部品区分, 0) =  " & Trim(Str(wCnt1 - 1))
                wStr2 = "   AND "
                Exit For
            End If
        Next
    End If
    wStr1 = wStr1 & wStr2 & "TM1.プレススポット区分 = 0": wStr2 = "   AND "
    wStr1 = wStr1 & wStr2 & "HKJ.日報出力 = 0": wStr2 = "   AND "
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT ISNULL(HKJ.優先順位,'') AS 優先順位, ISNULL(HKJ.加工数,0) AS 加工数"
    gSL_Select = gSL_Select & "      ,HKJ.取引先, TM1.略称 AS 取引先名"
    gSL_Select = gSL_Select & "      ,HKJ.機械, ISNULL(MM1.値名称,'') AS 機械名, HKJ.機械NO, ISNULL(MM2.値名称,'') AS 機械NO名"
    gSL_Select = gSL_Select & "      ,HM1.表示品番, HM1.得意先, TM2.略称 AS 得意先名, ZH1.材料管理番号, (ZM1.材質 + ZM1.表示寸法) 材料名"
    gSL_Select = gSL_Select & "      ,HKJ.備考, CASE WHEN ISNULL(HT1.部品区分,0) = 0 THEN 0 ELSE HT1.部品区分 END 部品区分"
    gSL_Select = gSL_Select & "      ,HKJ.品番, HKJ.受入, HKJ.品目, HKJ.工順, TM1.プレススポット区分 AS 表示工程数"
    gSL_Select = gSL_Select & " FROM  品番工順マスタ HKJ"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TM1"
    gSL_Select = gSL_Select & "        ON  TM1.取引先CD = HKJ.取引先"
    gSL_Select = gSL_Select & "        AND TM1.取引先区分 = 1 AND TM1.社内区分 = 1"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ MM1"
    gSL_Select = gSL_Select & "        ON  MM1.区分名称 = '機械'   AND MM1.値 = HKJ.機械"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 名称マスタ MM2"
    gSL_Select = gSL_Select & "        ON  MM2.区分名称 = '機械NO' AND MM2.値 = HKJ.機械NO"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番マスタ HM1"
    gSL_Select = gSL_Select & "        ON  HM1.品番 = HKJ.品番"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 取引先マスタ TM2"
    gSL_Select = gSL_Select & "        ON  TM2.取引先CD = HM1.得意先"
    gSL_Select = gSL_Select & "        AND TM2.取引先区分 = 0"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT 使用品番, MAX(材料管理番号) AS 材料管理番号"
    gSL_Select = gSL_Select & "                         FROM  材料品番マスタ"
    gSL_Select = gSL_Select & "                         GROUP BY 使用品番)ZH1"
    gSL_Select = gSL_Select & "        ON  ZH1.使用品番 = HKJ.品番"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 材料マスタ ZM1"
    gSL_Select = gSL_Select & "        ON  ZM1.材料管理番号 = ZH1.材料管理番号"
''''gSL_Select = gSL_Select & "       LEFT OUTER JOIN 品番取引先マスタ HT1"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT 品番, 取引先, 受入, ISNULL(部品区分,0) AS 部品区分 FROM 品番取引先マスタ) HT1"
    gSL_Select = gSL_Select & "        ON  HT1.品番   = HKJ.品番"
    gSL_Select = gSL_Select & "        AND HT1.取引先 = HKJ.取引先"
    gSL_Select = gSL_Select & "        AND HT1.受入   = HKJ.受入"
    gSL_Select = gSL_Select & wStr1
''''gSL_Select = gSL_Select & " GROUP BY HKJ.優先順位,HKJ.加工数"
''''gSL_Select = gSL_Select & "         ,HKJ.取引先, TM1.略称"
''''gSL_Select = gSL_Select & "         ,HKJ.機械, MM1.値名称, HKJ.機械NO, MM2.値名称"
''''gSL_Select = gSL_Select & "         ,HM1.表示品番, HM1.得意先, TM2.略称, ZH1.材料管理番号, ZM1.材料名"
''''gSL_Select = gSL_Select & "         ,HKJ.備考, HT1.部品区分"
''''gSL_Select = gSL_Select & "         ,HKJ.品番, HKJ.受入, HKJ.品目, HKJ.工順, TM1.プレススポット区分"
'================================== 2011/01/28 Start 品番順に変更
''    gSL_Select = gSL_Select & " ORDER BY CASE WHEN ISNULL(HKJ.優先順位,'')='' THEN 'ZZZZ' ELSE HKJ.優先順位 END"
''    gSL_Select = gSL_Select & "         ,HKJ.取引先, HKJ.機械, HKJ.機械NO, HKJ.品番, HM1.得意先"
    gSL_Select = gSL_Select & " ORDER BY CASE WHEN ISNULL(HKJ.優先順位,'')='' THEN 'ZZZZ' ELSE HKJ.優先順位 END"
    gSL_Select = gSL_Select & "         ,HKJ.取引先, HKJ.品番, HKJ.機械, HKJ.機械NO,  HM1.得意先"
'================================== 2011/01/28 End

    If Not HKTRead(gSL_Select, 0, 1) Then GoTo DBInput_Ed
    
    On Error GoTo 0
    MessageGet.Visible = True: DoEvents
    Call GridSet
    MessageGet.Visible = False: DoEvents
    DBInput = True
    
DBInput_Ed:
    Exit Function
DBInput_Err:
    CisFun.ErrorBox
End Function

'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    更新処理          +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim RirekiNo  As Integer

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     更新処理を行います            "
        .MB_Title = "更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Call VsGridToText(Val(Input_Area.Tag))
            Call GridTextEnabled(True)
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    Message.Visible = True
    DoEvents
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

'   マスタ項目移送      品番=19/取引先=16/受入=21/品目=27/工順=22/優先順位=05/加工数=06
    For gIndex = 1 To VsGrid.Rows - 1
        If RTrim(VsGrid.TextMatrix(gIndex, CisVsGrid.SelectCol)) <> "" Then
            gSL_Select = "SELECT * FROM 品番工順マスタ"
            gSL_Select = gSL_Select & " WHERE 品番   = '" & Trim(RTrim(VsGrid.TextMatrix(gIndex, 19))) & "'"
            gSL_Select = gSL_Select & "   AND 取引先 = '" & Trim(RTrim(VsGrid.TextMatrix(gIndex, 16))) & "'"
            gSL_Select = gSL_Select & "   AND 受入   = '" & Trim(RTrim(VsGrid.TextMatrix(gIndex, 21))) & "'"
            gSL_Select = gSL_Select & "   AND 品目   =  " & Trim(RTrim(VsGrid.TextMatrix(gIndex, 27)))
            gSL_Select = gSL_Select & "   AND 工順   =  " & Trim(RTrim(VsGrid.TextMatrix(gIndex, 22)))
            Call HKTRead(gSL_Select, 1)
            With HKT
               '===== 優先順位 =====
               '.優先順位 = Trim(RTrim(VsGrid.TextMatrix(gIndex, 5)))               '2010/07/05
               .優先順位 = Trim(RTrim(VsGrid.TextMatrix(gIndex, 4)))                '2010/07/05
               '===== 加工数 =======
               '.加工数 = Trim(Str(CisFun.Val2(VsGrid.TextMatrix(gIndex, 6))))      '2010/07/05
               .加工数 = Trim(Str(CisFun.Val2(VsGrid.TextMatrix(gIndex, 5))))       '2010/07/05
               '===== 更新者 =======
               .更新者 = gTanto
            End With
       ' ( 更新処理 )
            Call HKTUpdate
        End If
    Next gIndex
'
' ( 更新処理 )
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    Message.Visible = False
    DoEvents
    DBPut = True
    Exit Function
DBPut_Err:
    Message.Visible = False
    DoEvents
    CisFun.ErrorBox
    End
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(左中, 左中, あり, 3, 0, "№")
'        Call .FixedSet(左中, 左中, あり, 23, 1, "部署")                    '2010/07/05
        Call .FixedSet(左中, 左中, あり, 0, 1, "部署")                      '2010/07/05
        Call .FixedSet(左中, 左中, あり, 10, 2, "機械")
'        Call .FixedSet(左中, 左中, あり, 10, 3, "機械№")                  '2010/07/05
        Call .FixedSet(左中, 左中, あり, 5, 3, "機№")                    '2010/07/05
'        Call .FixedSet(左中, 左中, あり, 26, 4, "品番")                    '2010/07/05
        Call .FixedSet(左中, 左中, あり, 20, 4, "品番")                     '2010/07/05
        Call .FixedSet(左中, 左中, あり, 4, 5, "優先")
'        Call .FixedSet(左中, 右中, あり, 8, 6, "加工数")
        Call .FixedSet(左中, 右中, あり, 9, 6, "加工数")
        Call .FixedSet(左中, 左中, あり, 13, 7, "得意先")
        Call .FixedSet(左中, 左中, あり, 8, 8, "SD(親)")
'        Call .FixedSet(左中, 左中, あり, 26, 9, "材料")                    '2010/07/05
        Call .FixedSet(左中, 左中, あり, 30, 9, "材料")                     '2010/07/05
        Call .FixedSet(左中, 右中, あり, 10, 10, "材料在庫量")
        Call .FixedSet(左中, 右中, あり, 10, 11, "内示当月")
        Call .FixedSet(左中, 右中, あり, 10, 12, "内示翌月")
        Call .FixedSet(左中, 右中, あり, 10, 13, "内示翌々月")
        Call .FixedSet(左中, 左中, あり, 12, 14, "プレス順")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 10, 15, "Select")
        Call .FixedSet(中中, 左中, なし, 10, 16, "x取引先")
        Call .FixedSet(中中, 左中, なし, 10, 17, "x機械")
        Call .FixedSet(中中, 左中, なし, 10, 18, "x機械№")
        Call .FixedSet(中中, 左中, なし, 10, 19, "x品番")
        Call .FixedSet(中中, 左中, なし, 10, 20, "x得意先")
        Call .FixedSet(中中, 左中, なし, 10, 21, "x受入")
        Call .FixedSet(中中, 左中, なし, 10, 22, "x工順")
        Call .FixedSet(中中, 左中, なし, 10, 23, "x部品区分")
        Call .FixedSet(中中, 左中, なし, 10, 24, "x材料管理番号")
        Call .FixedSet(中中, 左中, なし, 10, 25, "x優先順位")
        Call .FixedSet(中中, 左中, なし, 10, 26, "x加工数")
        Call .FixedSet(中中, 左中, なし, 10, 27, "x品目")
        Call .FixedSet(中中, 左中, なし, 10, 28, "x修正")
        Call .FixedSet(中中, 左中, なし, 10, 29, "xプス区分")
        
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
'        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14")
        Call .PatanSet(0, "0,2,3,4,5,6,7,9,11,12,13,14")
        
        If DebugFlg Then Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15,16,17,18,19,20,21,22,23,24,25,26,27,28,29")  ' Debug 用
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
    Dim wkSQL       As String
    
' 抽出処理 初期ｾｯﾄ
    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 > .Rows Then
            .Rows = Ttl_Kensu + 1
        End If
    End With
    
    ID = 0
    Do Until Not HKT_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With VsGrid                                         ''' 記憶域は16以降
            .TextMatrix(ID, GridCol(0)) = ID
            '============= 部署 ================
            .TextMatrix(ID, GridCol(1)) = Trim(HKT.取引先)
            If Trim(HKT.取引先名) <> "" Then .TextMatrix(ID, GridCol(1)) = _
                                             .TextMatrix(ID, GridCol(1)) & ":" & Trim(HKT.取引先名)
            .TextMatrix(ID, GridCol(16)) = Trim(HKT.取引先)

            '============== 機械 ===============
            .TextMatrix(ID, GridCol(2)) = Trim(HKT.機械)
            If Trim(HKT.機械名) <> "" Then .TextMatrix(ID, GridCol(2)) = _
                                           .TextMatrix(ID, GridCol(2)) & ":" & Trim(HKT.機械名)
            .TextMatrix(ID, GridCol(17)) = Trim(HKT.機械)

            '============== 機械№ =============
            .TextMatrix(ID, GridCol(3)) = Trim(HKT.機械NO)
            If Trim(HKT.機械NO名) <> "" Then .TextMatrix(ID, GridCol(3)) = _
                                             .TextMatrix(ID, GridCol(3)) & ":" & Trim(HKT.機械NO名)
            .TextMatrix(ID, GridCol(18)) = Trim(HKT.機械NO)

            '============== 品番 ===============
            .TextMatrix(ID, GridCol(4)) = Trim(HKT.表示品番)
            .TextMatrix(ID, GridCol(19)) = Trim(HKT.品番)

            '============== 優先 ===============
            .TextMatrix(ID, GridCol(5)) = Trim(HKT.優先順位)
            .TextMatrix(ID, GridCol(25)) = Trim(HKT.優先順位)

            '============== 加工数 =============
            .TextMatrix(ID, GridCol(6)) = Format(HKT.加工数, "###,###")
            .TextMatrix(ID, GridCol(26)) = Format(HKT.加工数, "###,###")

            '============== 得意先 =============
            .TextMatrix(ID, GridCol(7)) = Trim(HKT.得意先)
            If Trim(HKT.得意先名) <> "" Then .TextMatrix(ID, GridCol(7)) = _
                                             .TextMatrix(ID, GridCol(7)) & ":" & Trim(HKT.得意先名)
            .TextMatrix(ID, GridCol(20)) = Trim(HKT.得意先)
            
''''        '============== SD =================
''''        .TextMatrix(ID, GridCol(8)) = Trim(HKT.SD)
            
            '============== 材料 ===============
            .TextMatrix(ID, GridCol(9)) = Trim(HKT.材料名)
            '============== 材料在庫量 =========
            .TextMatrix(ID, GridCol(10)) = HKT.材料在庫量

            '============== 内示当月 ===========
            .TextMatrix(ID, GridCol(11)) = HKT.内示当月
            '============== 内示翌月 ===========
            .TextMatrix(ID, GridCol(12)) = HKT.内示翌月
            '============== 内示翌々月 =========
            .TextMatrix(ID, GridCol(13)) = HKT.内示翌々月
            
            '============== プレス順(備考) =====
            .TextMatrix(ID, GridCol(14)) = Trim(HKT.備考)
        
            '============== 受入 ===============
            .TextMatrix(ID, GridCol(21)) = Trim(HKT.受入)
            '============== 工順 ===============
            .TextMatrix(ID, GridCol(22)) = Trim(HKT.工順)
            '============== 部品区分 ===========
            .TextMatrix(ID, GridCol(23)) = Trim(HKT.部品区分)
            '============== 材料管理番号 =======
            .TextMatrix(ID, GridCol(24)) = Trim(HKT.材料管理番号)
            '============== 品目 ===============
            .TextMatrix(ID, GridCol(27)) = Trim(HKT.品目)
            
            '============== プレススポット区分 ===============
            .TextMatrix(ID, GridCol(29)) = Trim(HKT.表示工程数)
        
            '============== SD(親品番) =========
'=================================== 2010/07/05 Start
'            Call GetSD(ID)                      ' 構成問合せ
'
'            wkSQL = "SELECT MAX(KTW.親品番) AS 親品番, HTM.背番号 AS 表示品番"
'            wkSQL = wkSQL & " FROM  構成問合せワーク KTW"
'            wkSQL = wkSQL & "       LEFT OUTER JOIN (SELECT 品番, MAX(背番号) AS 背番号"
'            wkSQL = wkSQL & "                         From 品番取引先マスタ"
'            wkSQL = wkSQL & "                         Where 品目 = 0"
'            wkSQL = wkSQL & "                         GROUP BY 品番) HTM"
'            wkSQL = wkSQL & "        ON  HTM.品番 = KTW.親品番"
'            wkSQL = wkSQL & " WHERE 処理端末 = HOST_NAME()"
'            wkSQL = wkSQL & "   AND 子品番   = '" & Trim(.TextMatrix(ID, 19)) & "'"
'            wkSQL = wkSQL & " GROUP BY HTM.背番号"
'            If Not KOLRead(wkSQL) Then .TextMatrix(ID, GridCol(8)) = "" _
'                                  Else .TextMatrix(ID, GridCol(8)) = KOL.表示品番
            .TextMatrix(ID, GridCol(8)) = ""
'=================================== 2010/07/05 End
            '============== 材料在庫数 =========
            'Call GetZaikoSu(ID)                                            '2010/07/05
            '.TextMatrix(ID, 10) = Format(HKT.材料在庫量, "#,###0")         '2010/07/05
            .TextMatrix(ID, GridCol(10)) = ""                                        '2010/07/05
            '============== 内示数 =============
            Call GetNaijiSu(ID)
            .TextMatrix(ID, GridCol(11)) = Format(HKT.内示当月, "#,###0")
            .TextMatrix(ID, GridCol(12)) = Format(HKT.内示翌月, "#,###0")
            .TextMatrix(ID, GridCol(13)) = Format(HKT.内示翌々月, "#,###0")
                                  
        End With
        
        Call HKTReadNext(1)
    Loop
    
    With VsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    Call HKTClose(1)
    
End Sub

'---------------------------------------------------'
'       親品番抽出                                  '
'---------------------------------------------------'
Private Function GetSD(INDEX As Integer) As Boolean
    With CisDB
        .SQL = "DELETE FROM 構成問合せワーク WHERE 処理端末 = HOST_NAME()"
        .DBExec
        
        .SQL = "構成問合せ情報抽出"
        .StoadoCount = 4
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = Input用
        .ParaNo = 4: .ParaIO = Input用
    
        .ParaNo = 1: .ParaValue = Trim(VsGrid.TextMatrix(INDEX, 19))
        .ParaNo = 2: .ParaValue = 0     ' 品番／材料　：品番        ' CisFun.Val2(H1Op_Hinkb(0).Tag)
        .ParaNo = 3: .ParaValue = 1     ' 正　／逆    ：正          ' CisFun.Val2(H1Op_Tenkb(0).Tag)
        .ParaNo = 4: .ParaValue = 1     ' 直前／製品　：製品        ' CisFun.Val2(H1Op_Oyakb(0).Tag)
    
        If Not .DBStored Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     構成マスタ情報の抽出に失敗しました。     "
                .MB_Title = "抽出処理"
                .MB_Button = Error
                If .MBOX Then Exit Function
            End With
        End If
    End With
End Function

'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    Input_Area.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Area.Tag))
    gInt = VsGrid.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > VsGrid.Rows - 1 Then
            GoTo Grid_RowChange_End
        Else
            If RTrim(VsGrid.TextMatrix(gInt + 1, 1)) = "" Then
                GoTo Grid_RowChange_End
            End If
        End If
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid.DispMax
        End If
    End If
    'PageDown
    If KeyCode = vbKeyPageDown Then
        If gInt + CisVsGrid.DispMax > VsGrid.Rows - 1 Then
            gInt = VsGrid.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    VsGrid.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText(VsGrid.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With VsGrid
        '====== 優先順位 ======
        If Trim(iB1_Yusen) <> "" Then
'            .TextMatrix(tRow, 5) = Trim(iB1_Yusen)         '2010/07/05
            .TextMatrix(tRow, 4) = Trim(iB1_Yusen)          '2010/07/05
        Else
'            .TextMatrix(tRow, 5) = ""                      '2010/07/05
            .TextMatrix(tRow, 4) = ""                       '2010/07/05
        End If
        '====== 加工数 ========
        If CisFun.Val2(iB1_Kakou) <> 0 Then
'            .TextMatrix(tRow, 6) = Format(CisFun.Val2(iB1_Kakou), "###,##0")               '2010/07/05
            .TextMatrix(tRow, 5) = Format(CisFun.Val2(iB1_Kakou), "###,##0")                '2010/07/05
        Else
'            .TextMatrix(tRow, 6) = ""                      '2010/07/05
            .TextMatrix(tRow, 5) = ""                       '2010/07/05
        End If
    End With
    SaveRow = VsGrid.Row
    Call UpdateCheck(tRow)
    VsGrid.Row = SaveRow
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With VsGrid
        Input_Area.Tag = tRow
        Input_Area.Move Input_Area.Left, VsGrid.CellTop, Input_Area.Width, VsGrid.CellHeight
        '== 優先順位
'        If Trim(.TextMatrix(tRow, 5)) <> "" Then
        If Trim(.TextMatrix(tRow, 4)) <> "" Then
'            iB1_Yusen = Trim(.TextMatrix(tRow, 5))
            iB1_Yusen = Trim(.TextMatrix(tRow, 4))
        Else
            iB1_Yusen = ""
        End If
        '== 加工数
'        If CisFun.Val2(.TextMatrix(tRow, 6)) <> 0 Then                                         '2010/07/05
        If CisFun.Val2(.TextMatrix(tRow, 5)) <> 0 Then                                          '2010/07/05
'            iB1_Kakou = Format(CisFun.Val2(.TextMatrix(tRow, 6)), "###,##0")                   '2010/07/05
            iB1_Kakou = Format(CisFun.Val2(.TextMatrix(tRow, 5)), "###,##0")                    '2010/07/05
        Else
            iB1_Kakou = ""
        End If
    End With
End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    Call TextToVsGrid(Val(Input_Area.Tag))
    Call VsGridToText(VsGrid.Row)
End Sub
Private Sub VSGrid_Click()
    Call TextToVsGrid(Val(Input_Area.Tag))
    If VsGrid.TextMatrix(VsGrid.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(VsGrid.Row)
    Call vsGrid_Scroll
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
Private Sub VSGrid_DblClick()
    If RTrim(VsGrid.TextMatrix(VsGrid.Row, 1)) = "" Then Exit Sub
    If VsGrid.Col = 1 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     " & VsGrid.Row & " 行目をクリアしますか？      "
            .MB_Title = ""
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Sub
        End With
        iB1_Yusen = "": iB1_Kakou = ""
        Call TextToVsGrid(Val(Input_Area.Tag))
    End If
End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
    If ProcHB <> "B1" Then Exit Sub
    With VsGrid
'        Input_Area.Left = .ColPos(5) - 30              '2010/07/05
        Input_Area.Left = .ColPos(4) - 30               '2010/07/05
        Input_Area.Top = .RowPos(.Row)
        If Input_Area.Left < .ColWidth(0) - 30 Or Input_Area.Top < .RowHeightMin Then
            Input_Area.Visible = False
        Else
            Input_Area.Visible = True
        End If
       'DoEvents
    End With
End Sub

'+---------------------------------------------+
'+      変更チェック
'+---------------------------------------------+
Private Sub UpdateCheck(sRow As Integer)
    VsGrid.Row = sRow
'================== '2010/07/05
'    If Trim(VsGrid.TextMatrix(VsGrid.Row, 5)) <> Trim(VsGrid.TextMatrix(VsGrid.Row, 25)) Or
'       CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 6)) <> CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 26))
    If Trim(VsGrid.TextMatrix(VsGrid.Row, 4)) <> Trim(VsGrid.TextMatrix(VsGrid.Row, 25)) Or _
       CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 5)) <> CisFun.Val2(VsGrid.TextMatrix(VsGrid.Row, 26)) _
    Then
       If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) = "" Then
'           VsGrid.TextMatrix(VsGrid.Row, 28) = "修正"                          '2010/07/05
           VsGrid.TextMatrix(VsGrid.Row, 28) = "修正"                           '2010/07/05
            iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) + 1, "###")
           Call CisVsGrid.vsColor(個別, &HFFFFC0)
       End If
    Else
       If RTrim(VsGrid.TextMatrix(VsGrid.Row, CisVsGrid.SelectCol)) <> "" Then
'           VsGrid.TextMatrix(VsGrid.Row, 28) = "*"                             '2010/07/05
           VsGrid.TextMatrix(VsGrid.Row, 28) = "*"                              '2010/07/05
            iT1_Mntsu = Format(CisFun.Val2(iT1_Mntsu) - 1, "###")
           Call CisVsGrid.vsColor(消去個別)
       End If
    End If
End Sub

'****************************************
'*          内示数算出　　              *
'****************************************   ' Based PSA0080
Private Function GetNaijiSu(INDEX As Integer) As Boolean
    Static wsBDFlg  As Boolean
    Dim YM0 As String
    Dim YM1 As String
    Dim YM2 As String
    
    GetNaijiSu = False
'
    YM0 = Format(Now(), "yyyymm")
        
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add RTrim(YM0), M, 1, ResultDate: YM1 = .Date_Result
        .Date_Add RTrim(YM0), M, 2, ResultDate: YM2 = .Date_Result
    End With

' 発注内示テーブル検索
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT HN0.内示年月, HN0.品番, HN0.手配先, HN0.受入"
    gSL_Select = gSL_Select & "      ,ISNULL(HN1.数量,0) AS 当月内示数"
    gSL_Select = gSL_Select & "      ,ISNULL(HN2.数量,0) AS 翌月内示数"
    gSL_Select = gSL_Select & "      ,ISNULL(HN3.数量,0) AS 翌々月内示数"
    gSL_Select = gSL_Select & " FROM  発注内示テーブル HN0"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(H1.数量) AS 数量, H1.内示年月, H1.品番, H1.手配先, H1.受入"
    gSL_Select = gSL_Select & "                         FROM  発注内示テーブル H1"
    gSL_Select = gSL_Select & "                         WHERE H1.内示年月 = '" & Trim(YM0) & "'"
    gSL_Select = gSL_Select & "                           AND H1.計画年月 = '" & Trim(YM0) & "'"
    gSL_Select = gSL_Select & "                         GROUP BY H1.内示年月, H1.計画年月, H1.品番, H1.手配先, H1.受入) HN1"
    gSL_Select = gSL_Select & "        ON  HN1.内示年月 = HN0.内示年月"
    gSL_Select = gSL_Select & "        AND HN1.品番     = HN0.品番"
    gSL_Select = gSL_Select & "        AND HN1.手配先   = HN0.手配先"
    gSL_Select = gSL_Select & "        AND HN1.受入     = HN0.受入"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(H2.数量) AS 数量, H2.内示年月, H2.品番, H2.手配先, H2.受入"
    gSL_Select = gSL_Select & "                         FROM  発注内示テーブル H2"
    gSL_Select = gSL_Select & "                         WHERE H2.内示年月 = '" & Trim(YM0) & "'"
    gSL_Select = gSL_Select & "                           AND H2.計画年月 = '" & Trim(YM1) & "'"
    gSL_Select = gSL_Select & "                         GROUP BY H2.内示年月, H2.計画年月, H2.品番, H2.手配先, H2.受入) HN2"
    gSL_Select = gSL_Select & "        ON  HN2.内示年月 = HN0.内示年月"
    gSL_Select = gSL_Select & "        AND HN2.品番     = HN0.品番"
    gSL_Select = gSL_Select & "        AND HN2.手配先   = HN0.手配先"
    gSL_Select = gSL_Select & "        AND HN2.受入     = HN0.受入"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUM(H3.数量) AS 数量, H3.内示年月, H3.品番, H3.手配先, H3.受入"
    gSL_Select = gSL_Select & "                         FROM  発注内示テーブル H3"
    gSL_Select = gSL_Select & "                         WHERE H3.内示年月 = '" & Trim(YM0) & "'"
    gSL_Select = gSL_Select & "                           AND H3.計画年月 = '" & Trim(YM2) & "'"
    gSL_Select = gSL_Select & "                         GROUP BY H3.内示年月, H3.計画年月, H3.品番, H3.手配先, H3.受入) HN3"
    gSL_Select = gSL_Select & "        ON  HN3.内示年月 = HN0.内示年月"
    gSL_Select = gSL_Select & "        AND HN3.品番     = HN0.品番"
    gSL_Select = gSL_Select & "        AND HN3.手配先   = HN0.手配先"
    gSL_Select = gSL_Select & "        AND HN3.受入     = HN0.受入"
    gSL_Select = gSL_Select & " WHERE HN0.内示年月 = '" & YM0 & "'"
    gSL_Select = gSL_Select & "   AND HN0.手配先   = '" & Trim(VsGrid.TextMatrix(INDEX, 16)) & "'"
    gSL_Select = gSL_Select & "   AND HN0.品番     = '" & Trim(VsGrid.TextMatrix(INDEX, 19)) & "'"
    gSL_Select = gSL_Select & "   AND HN0.受入     = '" & Trim(VsGrid.TextMatrix(INDEX, 21)) & "'"
    
    gSL_Select = gSL_Select & " GROUP BY HN0.内示年月, HN0.品番, HN0.手配先, HN0.受入, HN1.数量, HN2.数量, HN3.数量"
    
    HKT.内示当月 = 0: HKT.内示翌月 = 0: HKT.内示翌々月 = 0
    If Not HNJRead(gSL_Select) Then
       Exit Function
    End If
    HKT.内示当月 = HNJ.当月内示数: HKT.内示翌月 = HNJ.翌月内示数: HKT.内示翌々月 = HNJ.翌々月内示数
        
    Call HNJClose
    GetNaijiSu = True
End Function
'****************************************
'*          材料在庫量算出              *
'****************************************   ' Based PXZ0020
Private Function GetZaikoSu(INDEX As Integer) As Boolean
    GetZaikoSu = False
    Dim wKanriNo    As Integer: wKanriNo = 0
    Dim wZaikoSu    As Long:    wZaikoSu = 0
    Dim wMasterSu   As Long:    wMasterSu = 0
    Dim YM0         As String:  YM0 = Format(Now(), "yyyymm")

    If CisFun.Val2(VsGrid.TextMatrix(INDEX, 24)) = 0 Then
        VsGrid.TextMatrix(INDEX, 10) = 0
        GetZaikoSu = True
        Exit Function
    End If
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT * FROM 材料在庫マスタ"
    gSL_Select = gSL_Select & " WHERE 材料管理番号 = " & Trim(VsGrid.TextMatrix(INDEX, 24))
    gSL_Select = gSL_Select & "   AND 年月度 = '" & Trim(YM0) & "'"
    
    If Not ZZMRead(gSL_Select) Then
        wMasterSu = 0
        '---在庫マスタ：無
        gSL_Select = ""
        gSL_Select = gSL_Select & ""
        gSL_Select = gSL_Select & "SELECT SUBSTRING(SN.入荷日,1,6) AS 入荷日, SN.材質, SN.板厚, SN.幅, SN.長さ"
        gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) + ISNULL(JK1.実績数,0) - ISNULL(ZS1.使用量,0) - ISNULL(JK2.実績数,0) AS 入荷数"
        gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) AS 入荷数入, ISNULL(JK1.実績数,0) AS 実績数入"
        gSL_Select = gSL_Select & "      ,ISNULL(JK2.実績数,0) AS 実績数出, ISNULL(ZS1.使用量,0) AS 使用量出"
        gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(S1.入荷日,1,6) AS 入荷日, S1.材質, S1.板厚, S1.幅, S1.長さ, SUM(S1.入荷数) AS 入荷数"
        gSL_Select = gSL_Select & "                         FROM  材料入荷テーブル S1"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(S1.入荷日,1,6), S1.材質, S1.板厚, S1.幅, S1.長さ) SN1"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(SN1.入荷日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND SN1.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND SN1.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND SN1.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND SN1.長さ     = SN.長さ"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J1.実績日,1,6) AS 実績日, J1.材質, J1.板厚, J1.幅, J1.長さ, SUM(J1.実績数) AS 実績数"
        gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J1"
        gSL_Select = gSL_Select & "                         WHERE J1.実績区分 = 1"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J1.実績日,1,6), J1.材質, J1.板厚, J1.幅, J1.長さ) JK1"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND JK1.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND JK1.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND JK1.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND JK1.長さ     = SN.長さ"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J2.実績日,1,6) AS 実績日, J2.材質, J2.板厚, J2.幅, J2.長さ, SUM(J2.実績数) AS 実績数"
        gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J2"
        gSL_Select = gSL_Select & "                         WHERE J2.実績区分 <> 1 AND J2.入力種類 <> 'A'"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J2.実績日,1,6), J2.材質, J2.板厚, J2.幅, J2.長さ) JK2"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND JK2.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND JK2.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND JK2.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND JK2.長さ     = SN.長さ"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(Z1.作業日,1,6) AS 作業日, Z1.材質, Z1.板厚, Z1.幅, Z1.長さ, SUM(Z1.使用量) AS 使用量"
        gSL_Select = gSL_Select & "                         FROM  材料使用量テーブル Z1"
        gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(z1.作業日,1,6), Z1.材質, Z1.板厚, Z1.幅, Z1.長さ) ZS1"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(ZS1.作業日,1,6) = SUBSTRING(SN.入荷日,1,6)"
        gSL_Select = gSL_Select & "        AND ZS1.材質     = SN.材質"
        gSL_Select = gSL_Select & "        AND ZS1.板厚     = SN.板厚"
        gSL_Select = gSL_Select & "        AND ZS1.幅       = SN.幅"
        gSL_Select = gSL_Select & "        AND ZS1.長さ     = SN.長さ"
        gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(YM0) & "'"
        gSL_Select = gSL_Select & "   AND SN.材料管理番号 = " & Trim(VsGrid.TextMatrix(INDEX, 24))
''''''''gSL_Select = gSL_Select & "   AND SN.材質 = '" & Trim(iH1_Zaist) & "'"
''''''''gSL_Select = gSL_Select & "   AND SN.板厚 =  " & Trim(iH1_Itatu) & " "
''''''''gSL_Select = gSL_Select & "   AND SN.幅   =  " & Trim(iH1_Width) & " "
''''''''gSL_Select = gSL_Select & "   AND SN.長さ =  " & Trim(iH1_Long) & " "
        gSL_Select = gSL_Select & " GROUP BY SUBSTRING(SN.入荷日,1,6), SN.材質,   SN.板厚,   SN.幅, SN.長さ"
        gSL_Select = gSL_Select & "         ,SN1.入荷数, JK1.実績数  , JK2.実績数,ZS1.使用量"
        gSL_Select = gSL_Select & " ORDER BY SUBSTRING(SN.入荷日,1,6)"
    Else
        '---在庫マスタ：有
        wKanriNo = ZZM.材料管理番号
        wMasterSu = ZZM.在庫数
        If Trim(ZZM.実棚日) = "" Then
            '---実棚日：空白
            gSL_Select = ""
            gSL_Select = gSL_Select & ""
            gSL_Select = gSL_Select & "SELECT SUBSTRING(SN.入荷日,1,6) AS 入荷日, SN.材質, SN.板厚, SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) + ISNULL(JK1.実績数,0) - ISNULL(ZS1.使用量,0) - ISNULL(JK2.実績数,0) AS 入荷数"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) AS 入荷数入, ISNULL(JK1.実績数,0) AS 実績数入"
            gSL_Select = gSL_Select & "      ,ISNULL(JK2.実績数,0) AS 実績数出, ISNULL(ZS1.使用量,0) AS 使用量出"
            gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(S1.入荷日,1,6) AS 入荷日, S1.材質, S1.板厚, S1.幅, S1.長さ, SUM(S1.入荷数) AS 入荷数"
            gSL_Select = gSL_Select & "                         FROM  材料入荷テーブル S1"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(S1.入荷日,1,6), S1.材質, S1.板厚, S1.幅, S1.長さ) SN1"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(SN1.入荷日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND SN1.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND SN1.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND SN1.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND SN1.長さ     = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J1.実績日,1,6) AS 実績日, J1.材質, J1.板厚, J1.幅, J1.長さ, SUM(J1.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J1"
            gSL_Select = gSL_Select & "                         WHERE J1.実績区分 = 1"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J1.実績日,1,6), J1.材質, J1.板厚, J1.幅, J1.長さ) JK1"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND JK1.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND JK1.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND JK1.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND JK1.長さ     = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J2.実績日,1,6) AS 実績日, J2.材質, J2.板厚, J2.幅, J2.長さ, SUM(J2.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J2"
            gSL_Select = gSL_Select & "                         WHERE J2.実績区分 <> 1 AND J2.入力種類 <> 'A'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J2.実績日,1,6), J2.材質, J2.板厚, J2.幅, J2.長さ) JK2"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND JK2.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND JK2.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND JK2.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND JK2.長さ     = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(Z1.作業日,1,6) AS 作業日, Z1.材質, Z1.板厚, Z1.幅, Z1.長さ, SUM(Z1.使用量) AS 使用量"
            gSL_Select = gSL_Select & "                         FROM  材料使用量テーブル Z1"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(z1.作業日,1,6), Z1.材質, Z1.板厚, Z1.幅, Z1.長さ) ZS1"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(ZS1.作業日,1,6) = SUBSTRING(SN.入荷日,1,6)"
            gSL_Select = gSL_Select & "        AND ZS1.材質     = SN.材質"
            gSL_Select = gSL_Select & "        AND ZS1.板厚     = SN.板厚"
            gSL_Select = gSL_Select & "        AND ZS1.幅       = SN.幅"
            gSL_Select = gSL_Select & "        AND ZS1.長さ     = SN.長さ"
            gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(YM0) & "'"
            gSL_Select = gSL_Select & "   AND SN.材料管理番号 = " & Trim(Str(wKanriNo))
            gSL_Select = gSL_Select & " GROUP BY SUBSTRING(SN.入荷日,1,6), SN.材質,   SN.板厚,   SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "         ,SN1.入荷数, JK1.実績数  , JK2.実績数,ZS1.使用量"
            gSL_Select = gSL_Select & " ORDER BY SUBSTRING(SN.入荷日,1,6)"
        Else
            '---実棚日：入力済
            gSL_Select = ""
            gSL_Select = gSL_Select & ""
            gSL_Select = gSL_Select & "SELECT SUBSTRING(SN.入荷日,1,6) AS 入荷日, SN.材質, SN.板厚, SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) + ISNULL(JK1.実績数,0) - ISNULL(ZS1.使用量,0) - ISNULL(JK2.実績数,0) AS 入荷数"
            gSL_Select = gSL_Select & "      ,ISNULL(SN1.入荷数,0) AS 入荷数入, ISNULL(JK1.実績数,0) AS 実績数入"
            gSL_Select = gSL_Select & "      ,ISNULL(JK2.実績数,0) AS 実績数出, ISNULL(ZS1.使用量,0) AS 使用量出"
            gSL_Select = gSL_Select & " FROM 材料入荷テーブル SN"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(S1.入荷日,1,6) AS 入荷日, S1.材質, S1.板厚, S1.幅, S1.長さ, SUM(S1.入荷数) AS 入荷数"
            gSL_Select = gSL_Select & "                         FROM  材料入荷テーブル S1"
            gSL_Select = gSL_Select & "                         WHERE S1.入荷日 > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(S1.入荷日,1,6), S1.材質, S1.板厚, S1.幅, S1.長さ) SN1"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(SN1.入荷日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  SN1.材質 = SN.材質 AND SN1.板厚 = SN.板厚 AND SN1.幅 = SN.幅 AND SN1.長さ = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J1.実績日,1,6) AS 実績日, J1.材質, J1.板厚, J1.幅, J1.長さ, SUM(J1.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J1"
            gSL_Select = gSL_Select & "                         WHERE J1.実績区分 = 1 AND J1.実績日 > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J1.実績日,1,6), J1.材質, J1.板厚, J1.幅, J1.長さ) JK1"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(JK1.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  JK1.材質 = SN.材質 AND JK1.板厚 = SN.板厚 AND JK1.幅 = SN.幅 AND JK1.長さ = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(J2.実績日,1,6) AS 実績日, J2.材質, J2.板厚, J2.幅, J2.長さ, SUM(J2.実績数) AS 実績数"
            gSL_Select = gSL_Select & "                         FROM  材料実績管理テーブル J2"
            gSL_Select = gSL_Select & "                         WHERE J2.実績区分 <> 1 AND J2.入力種類 <> 'A' AND J2.実績日 > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(J2.実績日,1,6), J2.材質, J2.板厚, J2.幅, J2.長さ) JK2"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(JK2.実績日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  JK2.材質 = SN.材質 AND JK2.板厚 = SN.板厚 AND JK2.幅 = SN.幅 AND JK2.長さ = SN.長さ"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN (SELECT SUBSTRING(Z1.作業日,1,6) AS 作業日, Z1.材質, Z1.板厚, Z1.幅, Z1.長さ, SUM(Z1.使用量) AS 使用量"
            gSL_Select = gSL_Select & "                         FROM  材料使用量テーブル Z1"
            gSL_Select = gSL_Select & "                         WHERE Z1.作業日   > '" & Trim(ZZM.実棚日) & "'"
            gSL_Select = gSL_Select & "                         GROUP BY SUBSTRING(z1.作業日,1,6), Z1.材質, Z1.板厚, Z1.幅, Z1.長さ) ZS1"
          ''gSL_Select = gSL_Select & "      --ON  SUBSTRING(ZS1.作業日,1,6) = SUBSTRING(SN.入荷日,1,6)"
          ''gSL_Select = gSL_Select & "      --AND"
            gSL_Select = gSL_Select & "        ON  ZS1.材質 = SN.材質 AND ZS1.板厚 = SN.板厚 AND ZS1.幅 = SN.幅 AND ZS1.長さ = SN.長さ"
            gSL_Select = gSL_Select & " WHERE SUBSTRING(SN.入荷日,1,6) >= '" & Trim(YM0) & "' AND SN.材料管理番号 = " & Trim(Str(ZZM.材料管理番号))
            gSL_Select = gSL_Select & " GROUP BY SUBSTRING(SN.入荷日,1,6), SN.材質,   SN.板厚,   SN.幅, SN.長さ"
            gSL_Select = gSL_Select & "         ,SN1.入荷数, JK1.実績数  , JK2.実績数,ZS1.使用量"
            gSL_Select = gSL_Select & " ORDER BY SUBSTRING(SN.入荷日,1,6)"
        End If
    End If
    
    If ZYTRead(gSL_Select) Then
        Do While ZYT_RDSTS
            wZaikoSu = wZaikoSu + ZYT.入荷数
            Call ZYTReadNext
        Loop
    End If
    
    wZaikoSu = wMasterSu + wZaikoSu
    HKT.材料在庫量 = wZaikoSu
  ''VsGrid.TextMatrix(Index, 10) = Format(wZaikoSu, "#,###0")

    GetZaikoSu = True
End Function
