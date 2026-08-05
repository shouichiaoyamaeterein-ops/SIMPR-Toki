VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{BA5D7819-2B67-11D4-806A-00000EA57F9E}#3.0#0"; "CISYMD~1.OCX"
Begin VB.Form CXZ0040 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "棚卸数量入力（材料）"
   ClientHeight    =   11190
   ClientLeft      =   1650
   ClientTop       =   1665
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   12
      Charset         =   128
      Weight          =   400
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
   ScaleHeight     =   11190
   ScaleWidth      =   15360
   WindowState     =   2  '最大化
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   810
      Left            =   10800
      ScaleHeight     =   750
      ScaleWidth      =   3630
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   2250
      Width           =   3690
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   75
         Top             =   75
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
         Left            =   975
         Top             =   75
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
         Left            =   1845
         Top             =   75
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
         Left            =   2715
         Top             =   75
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
   Begin VB.PictureBox H2_Area1 
      Height          =   960
      Left            =   870
      ScaleHeight     =   900
      ScaleWidth      =   8400
      TabIndex        =   10
      Top             =   2085
      Width           =   8460
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   870
         Index           =   0
         Left            =   45
         Top             =   15
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "棚卸ﾌﾟﾚｼｰﾄ№"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iH2_No 
            Height          =   360
            Left            =   345
            TabIndex        =   5
            Top             =   450
            Width           =   840
            _ExtentX        =   1482
            _ExtentY        =   635
            cFormat         =   "######"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "######"
            MaxLength       =   6
            cDataType       =   1
            cDataReplace    =   1
            cFbComma        =   0
            cFdAutoFormat   =   1
            cGFormat        =   "######"
            cILength        =   5
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   870
         Left            =   1635
         Top             =   15
         Width           =   3315
         _ExtentX        =   5847
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "材     質"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iH2_Zaishitsu 
            Height          =   375
            Left            =   60
            TabIndex        =   6
            Top             =   420
            Width           =   3255
            _ExtentX        =   5741
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
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   870
         Left            =   4965
         Top             =   15
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "板 厚"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iH2_Itaatsu 
            Height          =   375
            Left            =   60
            TabIndex        =   7
            Top             =   420
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   661
            cFormat         =   "###0.00"
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
            Text            =   "###0.00"
            MaxLength       =   7
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "###0.00"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   870
         Left            =   6105
         Top             =   15
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "幅"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iH2_Haba 
            Height          =   375
            Left            =   60
            TabIndex        =   8
            Top             =   420
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   661
            cFormat         =   "###0.00"
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
            Text            =   "###0.00"
            MaxLength       =   7
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "###0.00"
            cILength        =   4
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   870
         Left            =   7245
         Top             =   15
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   1535
         ForeColor       =   16711680
         Caption         =   "長 さ"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisText_V60.CisText iH2_Nagasa 
            Height          =   375
            Left            =   60
            TabIndex        =   9
            Top             =   420
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   661
            cFormat         =   "###0.00"
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
            Text            =   "###0.00"
            MaxLength       =   7
            cDataType       =   2
            cDataReplace    =   1
            cFbComma        =   0
            cFcDecimal      =   2
            cFdAutoFormat   =   1
            cGFormat        =   "###0.00"
            cILength        =   4
         End
      End
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Left            =   -30
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   420
      Width           =   855
   End
   Begin VB.PictureBox B1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7170
      Left            =   870
      ScaleHeight     =   7110
      ScaleWidth      =   13515
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   3285
      Width           =   13575
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   7005
         Left            =   60
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   45
         Width           =   13365
         _Version        =   196608
         _ExtentX        =   23574
         _ExtentY        =   12356
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "№ |ﾌﾟﾚｼｰﾄ№|材質                |板厚    |幅     |長さ    |置場      |理論在庫数|単価    |棚卸数量  |"
         Rows            =   21
         Cols            =   12
         BackColor       =   12648384
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         BackColorAlternate=   12648384
         Begin CisText_V60.CisText iB1_TSuryo 
            Height          =   330
            Left            =   7980
            TabIndex        =   13
            Top             =   315
            Width           =   1320
            _ExtentX        =   2328
            _ExtentY        =   582
            cFormat         =   "##,###,##0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "##,###,###"
            MaxLength       =   10
            cDataType       =   1
            cDataReplace    =   1
            cGFormat        =   "#######0"
            cILength        =   8
         End
         Begin Cis3D_v60.CIS3D Disp_Msg 
            Height          =   1005
            Left            =   2520
            Top             =   1890
            Visible         =   0   'False
            Width           =   7905
            _ExtentX        =   13944
            _ExtentY        =   1773
            BackColor       =   4210752
            Caption         =   ""
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            Begin Cis3D_v60.CIS3D Msg 
               Height          =   885
               Left            =   75
               Top             =   60
               Width           =   7755
               _ExtentX        =   13679
               _ExtentY        =   1561
               BackColor       =   12648447
               ForeColor       =   8388608
               Caption         =   "更新中　しばらくお待ち下さい"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ 明朝"
                  Size            =   21.75
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
      End
   End
   Begin VB.PictureBox H1_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   990
      Left            =   870
      ScaleHeight     =   930
      ScaleWidth      =   5085
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   960
      Width           =   5145
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   885
         Index           =   5
         Left            =   45
         Top             =   30
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "日付(実棚日)"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         Begin CisYMD_V60.CisYMD iH1_YMD 
            Height          =   390
            Left            =   225
            TabIndex        =   3
            Top             =   390
            Width           =   1935
            _ExtentX        =   3069
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
            Object.Height          =   390
            Object.Width           =   1935
            cSize           =   -1  'True
            cChkResult      =   0   'False
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   885
         Index           =   10
         Left            =   2475
         Top             =   30
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "担 当 者"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Tanto 
            Height          =   360
            Left            =   150
            TabIndex        =   4
            Top             =   390
            Width           =   495
            _ExtentX        =   873
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
            Text            =   "XXX"
            MaxLength       =   3
         End
         Begin Cis3D_v60.CIS3D H1lb_Tanto 
            Height          =   345
            Left            =   660
            Top             =   390
            Width           =   1725
            _ExtentX        =   3043
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXXXXXX3"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   795
   End
   Begin VB.Timer Timer2 
      Left            =   690
      Top             =   795
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
      Caption         =   "【 棚卸数量入力 (材料)  】"
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
         Left            =   11670
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
   End
End
Attribute VB_Name = "CXZ0040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   棚卸数量入力（材料）
'**       フォームID    :   CXZ0040
'**       処理概要      :
'**
'**       作  成  日    :   2010/01/14  By CIS
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替

    Dim CisVsGrid       As New CisVsGrid3
    
    Dim ActObj          As Object
    Dim ActRow          As Integer
    
    Dim wsDate          As String
    Dim SCol()          As Long
    
    Dim wDate           As String
    Dim wEndDay         As Byte
    
    Dim SaveRow         As Integer
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #------------------#
'   # 二 重 起 動 防 止 #
'   #------------------#
    If Not Dupli_Start Then End
'   #------------------#
'   # 初 期 内 容 取 得 #
'   #------------------#
    If Not IniGet Then End
'   #------------------#
'   #  項 目 初 期 化   #
'   #------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
    Call CisFrm.InitFld
    
    If gTantoName = "" Then
       Head_Tanto.Visible = False
    Else
       Head_Tanto.Visible = True
       Head_Tanto = gTantoName
    End If
    
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
    
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = VsGrid1
        .NoSpaceCol = 0
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    wsDate = Format(Now(), "yyyymmdd")
    iH1_YMD = Mid(wsDate, 1, 8)
    
    FormAct = Not FormAct
'
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
    Static W_IVENT As String
    W_IVENT = "Down"
    Key_Acc = False
    Select Case Key_Code
           Case vbKeyF7:       If PB_LOOK.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click    '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click    '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                      '【入力】
                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
                                  Dummy.Enabled = False
                               End If
           Case Else:          Exit Function
    End Select
    Key_Acc = True
End Function
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    '
    Cancel = 1
    Call PB_END_Click
End Sub
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'+---------------------+
'+   ﾃﾞｰﾀﾍﾞｰｽ切断
'+---------------------+
    Call CisDB.DBDISConnect
'+---------------------+
'+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
'+---------------------+
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
    
    If ProcHB = "B1" Then
        ProcHB = "H2"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        iH2_No.SetFocus
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "H2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("H")
        iH1_YMD.SetFocus
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_LOOK.Tag = "iH1_Tanto" Then Call Look_Tanto: Exit Sub
    If PB_LOOK.Tag = "iH2_Zaishitsu" Then Call Look_Zairyo: Exit Sub
    If PB_LOOK.Tag = "iH2_Itaatsu" Then Call Look_Zairyo: Exit Sub
    If PB_LOOK.Tag = "iH2_Haba" Then Call Look_Zairyo: Exit Sub
    If PB_LOOK.Tag = "iH2_Nagasa" Then Call Look_Zairyo: Exit Sub
End Sub
'+----------------------------------------------------------------------------+
'+      担当者(名称マスタ) 検索
'+----------------------------------------------------------------------------+
Private Sub Look_Tanto()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "担当管理"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    If RV_Rtn Then
        iH1_Tanto = RV_Code
        Call iH1_Tanto_LostFocus
    End If
End Sub
'+----------------------------+
'+      材料マスタ検索        +
'+----------------------------+
Private Sub Look_Zairyo()

    RV_Left = 0:    RV_Top = 0
    'RV_TorKb = 7: RV_Call = "C"
    RV_Zaist = iH2_Zaishitsu
    RV_Call = ""
    
    CKK0070.Show vbModal
    Unload CKK0070
    Set CKK0070 = Nothing

    If RV_Rtn Then
       iH2_Zaishitsu = RTrim(RV_Zaist)
       iH2_Itaatsu = RTrim(RV_Itatu)
       iH2_Haba = RTrim(RV_Width)
       iH2_Nagasa = RTrim(RV_Long)
    End If

End Sub
'+-----------------------------+
'+  iH1_YM
'+-----------------------------+
Private Sub iH1_YMD_GotFocus()
    PB_LOOK.Visible = False
    DoEvents
End Sub
'+-----------------------------+
'+  iH1_Tanto
'+-----------------------------+
Private Sub iH1_Tanto_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH1_Tanto_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If

    H1lb_Tanto = ""
    If MeisyoGet("担当管理", RTrim(iH1_Tanto)) Then
        H1lb_Tanto = RTrim(SYM_Meisyo)
    End If
End Sub
'+-----------------------------+
'+  iH2_Zaishitsu
'+-----------------------------+
Private Sub iH2_Zaishitsu_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH2_Zaishitsu_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH2_Itaatsu
'+-----------------------------+
Private Sub iH2_Itaatsu_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH2_Itaatsu_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH2_Haba
'+-----------------------------+
Private Sub iH2_Haba_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH2_Haba_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH2_Nagasa
'+-----------------------------+
Private Sub iH2_Nagasa_GotFocus()
    PB_LOOK.Tag = ActiveControl.Name
    PB_LOOK.Visible = True
End Sub
Private Sub iH2_Nagasa_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_LOOK.Tag = ""
        PB_LOOK.Visible = False
    End If
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
   CisFrm.TimeSet CisFun.INI_FileID, 110
   
   'Debug.Print ActiveControl.Name
   
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
    Dim H2Mode      As Boolean
    Dim H2Color     As Long
    Dim B1Mode      As Boolean
    Dim B1Color     As Long
    H1Mode = False
    H1Color = gPLostSel
    H2Mode = False
    H2Color = gPLostSel
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
       Case "H2"
          H2Mode = True
          H2Color = gPGotSel
       Case "B1"
          B1Mode = True
          B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H2_Area1.Enabled = H2Mode
    H2_Area1.BackColor = H2Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_LOOK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    
    Call GridTextEnabled(B1Mode)

' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
'
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*2" Else wbc_Enm = Trim(ClrType) & "2"
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
'                .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
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
            
        DoEvents
        ProcHB = "H2"
        Call DispChange(ProcHB)
        iH2_No.SetFocus
        DoEvents
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "H2" Then
        If Not Head2Chk Then GoTo ReturnPress_Ed

        DoEvents
        ProcHB = "B1"
        Call DispChange(ProcHB)
        VsGrid1.Row = 1
        Call VsGridToText(VsGrid1.Row)
        DoEvents
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
        If Not DBPut Then GoTo ReturnPress_Ed

        Call HeadBodyClear("B")
        ProcHB = "H2"
        Call DispChange(ProcHB)
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
    
' 日付（実棚日）
    If RTrim(iH1_YMD) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付(実棚日)を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_YMD.SetFocus
       Exit Function
    End If
    
    If iH1_YMD.cYear = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付(実棚日) - 年 を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_YMD.SetFocus
       Exit Function
    End If
    
    If iH1_YMD.cMonth = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付(実棚日) - 月 を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_YMD.SetFocus
       Exit Function
    End If
    
    If iH1_YMD.cDay = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    日付(実棚日) - 日 を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_YMD.SetFocus
       Exit Function
    End If
    
    If iH1_YMD.cChkResult = False Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    入力された日付(実棚日)は日付として正しくありません。       "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_YMD.SetFocus
       Exit Function
    End If
    
' 担当者
    If RTrim(iH1_Tanto) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    担当者を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Tanto.SetFocus
       Exit Function
    End If

    If RTrim(iH1_Tanto) <> "" Then
        'H1lb_Tanto = ""
        If Not MeisyoGet("担当管理", RTrim(iH1_Tanto)) Then
             With CisFun
                 .MB_Lines = 5
                 .MB_MSG(2) = "    担 当 者  未 登 録                "
                 .MB_MSG(4) = "      【 名称マスタ 】                "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iH1_Tanto.SetFocus
            Exit Function
        End If
    End If
    
    Head1Chk = True
End Function
'****************************
'*      HEAD2ﾁｪｯｸ処理
'****************************
Private Function Head2Chk() As Boolean
    Head2Chk = False
        
' 対象データ存在チェック
    Call DB_Select
    If Not ZTPRead(gSL_Select) Then
         With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "     該当データがありません。                 "
             .MB_Title = ""
             .MB_Button = Error
             .MBOX
         End With
        iH2_No.SetFocus
        Exit Function
    End If
    
' BODY部 セット
    Call BodySet("")
    
    Head2Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet(SetType As String)
    Static wsBDFlg  As Boolean

'
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add RTrim(iH1_YMD), M, -12, ResultDate
        gStr = .Date_Result
    End With


    Call DB_Select
    
    If ZTPRead(gSL_Select) Then
       Call GridSet
    End If

    Call ZTPClose
    
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()

    Dim WorkNengetsu As String
    Dim WorkZaiKanri As Integer
    Dim WCounter As Integer

    DBPut = False

    Call TextToVsGrid(Val(iB1_TSuryo.Tag))

    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     更新処理を行います      "
        .MB_Title = "更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    '更新中メッセージ表示
    Disp_Msg.Visible = True
    DoEvents

    WorkNengetsu = ""
    WorkZaiKanri = 0

    For gIndex = 1 To VsGrid1.Rows - 1

        If RTrim(VsGrid1.TextMatrix(gIndex, 6)) <> "" Then

            '材料棚卸プレシート発行テーブル読込
            gSL_Select = ""
            gSL_Select = "             Select ZP.*,ZZ.材質,ZZ.寸法区分,ZZ.板厚,ZZ.幅,ZZ.長さ,ZZ.SIZE From 材料棚卸プレシート発行テーブル ZP" & Chr(13)
            gSL_Select = gSL_Select & "    Left Outer Join 材料在庫マスタ ZZ" & Chr(13)
            gSL_Select = gSL_Select & "    On ZZ.材料管理番号 = ZP.材料管理番号" & Chr(13)
            gSL_Select = gSL_Select & "   And ZZ.年月度 = '" & Mid(iH1_YMD, 1, 6) & "'" & Chr(13)
            gSL_Select = gSL_Select & " Where 発行年月  = '" & Mid(iH1_YMD, 1, 6) & "'" & Chr(13)
            gSL_Select = gSL_Select & "   And 伝票NO    = '" & CisFun.Val2(Mid(VsGrid1.TextMatrix(gIndex, 1), 1, 6)) & "'" & Chr(13)
            gSL_Select = gSL_Select & "   And 伝票行NO  = '" & CisFun.Val2(Mid(VsGrid1.TextMatrix(gIndex, 1), 8, 2)) & "'" & Chr(13)
            If Not ZTPRead(gSL_Select) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_Title = "更新処理"
                    .MB_MSG(2) = "     材料棚卸プレシート発行テーブル     "
                    .MB_MSG(3) = "               読込エラー               "
                    .MB_Button = Error
                    If Not .MBOX Then Exit Function
                End With
            Else
                '材料在庫マスタ 削除
                gSL_Select = ""
                gSL_Select = "DELETE 材料在庫マスタ " & Chr(13)
                gSL_Select = gSL_Select & " WHERE 年月度         = '" & RTrim(ZTP.発行年月) & "'" & Chr(13)
                gSL_Select = gSL_Select & "   And 材料管理番号   = '" & RTrim(ZTP.材料管理番号) & "'" & Chr(13)
                With CisDB
                    .SQL = gSL_Select
                    If Not .DBExec Then Exit Function
                End With
                
                '材料在庫内訳テーブル作成
                gSL_Select = ""
                gSL_Select = "Select * From 材料在庫内訳テーブル " & Chr(13)
                gSL_Select = gSL_Select & " Where 年月度 = '" & RTrim(ZTP.発行年月) & "'" & Chr(13)
                gSL_Select = gSL_Select & "   And 材料管理番号   = '" & RTrim(ZTP.材料管理番号) & "'" & Chr(13)
                gSL_Select = gSL_Select & "   And 置場 = '" & RTrim(ZTP.在庫場所1) & RTrim(ZTP.在庫場所2) & RTrim(ZTP.棚番) & "'" & Chr(13)
                If ZZURead(gSL_Select) Then
                    '材料在庫内訳テーブル更新
                    gSL_Select = ""
                    gSL_Select = "Update 材料在庫内訳テーブル SET " & Chr(13)
                    gSL_Select = gSL_Select & " 材質     = '" & RTrim(ZTP.材質) & "'" & Chr(13)
                    gSL_Select = gSL_Select & ",寸法区分 = " & ZTP.寸法区分 & Chr(13)
                    gSL_Select = gSL_Select & ",板厚     = " & ZTP.板厚 & Chr(13)
                    gSL_Select = gSL_Select & ",幅       = " & ZTP.幅 & Chr(13)
                    gSL_Select = gSL_Select & ",長さ     = " & ZTP.長さ & Chr(13)
                    gSL_Select = gSL_Select & ",在庫数   = " & RTrim(VsGrid1.TextMatrix(gIndex, 6)) & Chr(13)
                    gSL_Select = gSL_Select & ",実棚日   = '" & RTrim(iH1_YMD) & "'" & Chr(13)
                    gSL_Select = gSL_Select & " Where 年月度 = '" & RTrim(ZTP.発行年月) & "'" & Chr(13)
                    gSL_Select = gSL_Select & "   And 材料管理番号   = '" & RTrim(ZTP.材料管理番号) & "'" & Chr(13)
                    gSL_Select = gSL_Select & "   And 置場 = '" & RTrim(ZTP.在庫場所1) & RTrim(ZTP.在庫場所2) & RTrim(ZTP.棚番) & "'" & Chr(13)
                    With CisDB
                        .SQL = gSL_Select
                        If Not .DBExec Then Exit Function
                    End With
                Else
                    '材料在庫内訳テーブル登録
                    Call ItemsClearZZU
                    With ZZU
                        .年月度 = Mid(iH1_YMD, 1, 6)
                        .材料管理番号 = RTrim(ZTP.材料管理番号)
                        .置場 = RTrim(ZTP.在庫場所1) & RTrim(ZTP.在庫場所2) & RTrim(ZTP.棚番)
                        .材質 = RTrim(ZTP.材質)
                        .寸法区分 = ZTP.寸法区分
                        .板厚 = ZTP.板厚
                        .幅 = ZTP.幅
                        .長さ = ZTP.長さ
                        .在庫数 = RTrim(VsGrid1.TextMatrix(gIndex, 6))
                        .実棚日 = RTrim(iH1_YMD)
                    End With
                    Call ZZUInsert
                End If
                
                '材料在庫マスタ作成
                gSL_Select = ""
                gSL_Select = "Select Distinct Z1.年月度,Z1.材料管理番号,Z1.材質,Z1.寸法区分,Z1.板厚,Z1.幅,Z1.長さ,Z2.在庫数" & Chr(13)
                gSL_Select = gSL_Select & " from 材料在庫内訳テーブル Z1" & Chr(13)
                gSL_Select = gSL_Select & " Left Outer Join (select 材料管理番号,SUM(在庫数) 在庫数" & Chr(13)
                gSL_Select = gSL_Select & "                    from 材料在庫内訳テーブル" & Chr(13)
                gSL_Select = gSL_Select & "                   Where 年月度 = '" & Mid(iH1_YMD, 1, 6) & "'" & Chr(13)
                gSL_Select = gSL_Select & "                   Group by 材料管理番号" & Chr(13)
                gSL_Select = gSL_Select & "                 ) Z2" & Chr(13)
                gSL_Select = gSL_Select & "   On Z2.材料管理番号 = Z1.材料管理番号" & Chr(13)
                gSL_Select = gSL_Select & " Where Z1.年月度 = '" & Mid(iH1_YMD, 1, 6) & "'" & Chr(13)
                gSL_Select = gSL_Select & "   And Z1.材料管理番号 = '" & RTrim(ZTP.材料管理番号) & "'" & Chr(13)
                If ZZURead(gSL_Select) Then
                    Call ItemsClearZZM
                    With ZZM
                        .年月度 = Mid(iH1_YMD, 1, 6)
                        .材料管理番号 = ZZU.材料管理番号
                        .材質 = RTrim(ZZU.材質)
                        .寸法区分 = ZZU.寸法区分
                        .板厚 = ZZU.板厚
                        .幅 = ZZU.幅
                        .長さ = ZZU.長さ
                        .在庫数 = ZZU.在庫数
                        .実棚日 = RTrim(iH1_YMD)
                        .SIZE = ZTP.SIZE
                        .作成区分 = 1
                    End With
                    ZZM.作成者 = RTrim(H1lb_Tanto)
                    Call ZZMInsert
                End If

                '材料在庫履歴マスタ更新
                gSL_Select = ""
                gSL_Select = "Update 材料在庫履歴マスタ SET " & Chr(13)
                gSL_Select = gSL_Select & " 履歴区分 = 1 " & Chr(13)
                gSL_Select = gSL_Select & " WHERE 年月度       = '" & Mid(iH1_YMD, 1, 6) & "'" & Chr(13)
                gSL_Select = gSL_Select & "   AND 材料管理番号 = '" & RTrim(ZTP.材料管理番号) & "'" & Chr(13)
                gSL_Select = gSL_Select & "   AND 実棚日       = '" & RTrim(iH1_YMD) & "'" & Chr(13)
                With CisDB
                    .SQL = gSL_Select
                    If Not .DBExec Then Exit Function
                End With

                '材料在庫履歴マスタ作成
                Call ItemsClearZZR
                With ZZR
                    .年月度 = Mid(iH1_YMD, 1, 6)
                    .材料管理番号 = ZZU.材料管理番号
                    .材質 = RTrim(ZZU.材質)
                    .寸法区分 = ZZU.寸法区分
                    .板厚 = ZZU.板厚
                    .幅 = ZZU.幅
                    .長さ = ZZU.長さ
                    .入力日付 = Now()
                    .入力担当者 = RTrim(iH1_Tanto)
                    .在庫数 = ZZU.在庫数
                    .実棚日 = RTrim(iH1_YMD)
                    .変更事由 = ""
                    .作成区分 = 1
                End With
                ZZR.作成者 = RTrim(H1lb_Tanto)
                If WorkNengetsu = Mid(iH1_YMD, 1, 6) And WorkZaiKanri = ZZU.材料管理番号 Then
                    WCounter = WCounter + 1
                Else
                    WCounter = 0
                End If
                
                WorkNengetsu = Mid(iH1_YMD, 1, 6)
                WorkZaiKanri = ZZU.材料管理番号
                
                Call ZZRInsert2(, WCounter)

                '材料棚卸プレシート発行テーブル更新
                gSL_Select = ""
                gSL_Select = "UPDATE 材料棚卸プレシート発行テーブル SET " & Chr(13)
                gSL_Select = gSL_Select & " 棚卸入力済フラグ = 1 " & Chr(13)
                gSL_Select = gSL_Select & " Where 発行年月 = '" & Mid(iH1_YMD, 1, 6) & "'" & Chr(13)
                gSL_Select = gSL_Select & "   And 伝票NO   = '" & CisFun.Val2(Mid(VsGrid1.TextMatrix(gIndex, 1), 1, 6)) & "'" & Chr(13)
                gSL_Select = gSL_Select & "   And 伝票行NO = '" & CisFun.Val2(Mid(VsGrid1.TextMatrix(gIndex, 1), 8, 2)) & "'" & Chr(13)
                With CisDB
                    .SQL = gSL_Select
                    If Not .DBExec Then Exit Function
                End With

            End If

        End If

    Next gIndex
    
    '更新中メッセージ消去
    Disp_Msg.Visible = False
    DoEvents
    
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    DBPut = True
    Exit Function
DBPut_Err:
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
    ReDim SCol(11)
    
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "NO")
        Call .FixedSet(中中, 中中, あり, 9, 1, "ﾌﾟﾚｼｰﾄ№")
        Call .FixedSet(中中, 左中, あり, 20, 2, "材質")
        Call .FixedSet(中中, 右中, あり, 10, 3, "板厚")
        Call .FixedSet(中中, 右中, あり, 10, 4, "幅")
        Call .FixedSet(中中, 右中, あり, 10, 5, "長さ")
        Call .FixedSet(中中, 右中, あり, 10, 6, "棚卸数量")
        Call .FixedSet(中中, 左中, あり, 11, 7, "置場")
        Call .FixedSet(中中, 右中, あり, 10, 8, "理論在庫数")
        Call .FixedSet(中中, 右中, あり, 8, 9, "単価")

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 10, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9")
        .InitDisp
    
        SCol(1) = .FixedGet(1)  ' ﾌﾟﾚｼｰﾄ№
        SCol(2) = .FixedGet(2)  ' 部署
        SCol(3) = .FixedGet(3)  ' 機械
        SCol(4) = .FixedGet(4)  ' 機械№
        SCol(5) = .FixedGet(5)  ' 品番
        SCol(6) = .FixedGet(6)  ' 棚卸数量
        SCol(7) = .FixedGet(7)  ' 置場
        SCol(8) = .FixedGet(8)  ' 理論在庫数
        SCol(9) = .FixedGet(9)  ' 単価
    End With
'    KCol = 8   ' 非表示項目開始ｶﾗﾑ
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(CisFun.Val2(ID))
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    
    ID = 0
    
    Do Until Not ZTP_RDSTS
        With VsGrid1
            ID = ID + 1
            ' №
            .TextMatrix(ID, 0) = Format(ID, "###")
            ' プレシート№
            .TextMatrix(ID, SCol(1)) = RTrim(ZTP.プレシートNO)
            ' 材質
            .TextMatrix(ID, SCol(2)) = RTrim(ZTP.材質)
            ' 板厚
            .TextMatrix(ID, SCol(3)) = Format(ZTP.板厚, "#,##0.#0")
            ' 幅
            .TextMatrix(ID, SCol(4)) = Format(ZTP.幅, "#,##0.#0")
            ' 長さ
            .TextMatrix(ID, SCol(5)) = Format(ZTP.長さ, "#,##0.#0")
            ' 置場
            .TextMatrix(ID, SCol(7)) = RTrim(ZTP.置場)
            ' 理論在庫数
            If ZTP.理論在庫数 = 0 Then
                .TextMatrix(ID, SCol(8)) = ""
            Else
                .TextMatrix(ID, SCol(8)) = Format(RTrim(ZTP.理論在庫数), "###,###,###,###")
            End If
            ' 単価
            With CisDB
                .SQL = "材料購入単価決定"
                .StoadoCount = 9
                .ParaNo = 0:    .ParaIO = Return用
                .ParaNo = 1:    .ParaIO = Input用
                .ParaNo = 2:    .ParaIO = Input用
                .ParaNo = 3:    .ParaIO = Input用
                .ParaNo = 4:    .ParaIO = Input用
                .ParaNo = 5:    .ParaIO = OutPut用
                .ParaNo = 6:    .ParaIO = OutPut用
                .ParaNo = 7:    .ParaIO = OutPut用
                .ParaNo = 8:    .ParaIO = OutPut用
                .ParaNo = 9:    .ParaIO = OutPut用
            
                .ParaNo = 1:    .ParaValue = ZTP.材料管理番号               ' 材料管理番号
                .ParaNo = 2:    .ParaValue = RTrim(ZTP.取引先)              ' 取引先
                .ParaNo = 3:    .ParaValue = ""                             ' 契約先
                .ParaNo = 4:    .ParaValue = RTrim(iH1_YMD)                 ' 単価決定日付
               
                .DBStored
            
               .ParaNo = 0
               If .ParaValue <> 0 Then
                   CisFun.MB_Lines = 5
                   CisFun.MB_MSG(2) = "    材料購入単価決定　異常！      "
                   CisFun.MB_MSG(4) = "    Error No = " & .ParaValue
                   CisFun.MB_Button = Error
                   CisFun.MBOX
                   End
               End If
               
                .ParaNo = 5
                If .ParaValue = 0 Then
                    VsGrid1.TextMatrix(ID, SCol(9)) = ""
                Else
                    VsGrid1.TextMatrix(ID, SCol(9)) = Format(RTrim(.ParaValue), "#,###.#0")
                End If
                
               
            End With
        
        End With
        Call ZTPReadNext
    Loop
    
    With VsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'******************************************'
'*****     グリッド内容セットSQL      *****
'******************************************'
Private Sub DB_Select()

    gSL_Select = ""
    gSL_Select = "             Select Right('000000' + Convert(varchar, ZP.伝票NO), 6) + '-' + " & Chr(13)
    gSL_Select = gSL_Select & "       Right('00' + Convert(varchar, ZP.伝票行NO), 2) プレシートNO," & Chr(13)
    gSL_Select = gSL_Select & "       IsNull(ZM.材質,'') 材質,IsNull(ZM.板厚,0) 板厚,IsNull(ZM.幅,0) 幅,IsNull(ZM.長さ,0) 長さ," & Chr(13)
    gSL_Select = gSL_Select & "       Case When IsNull(ZP.在庫場所1,'') = '' then '' " & Chr(13)
    gSL_Select = gSL_Select & "            Else ZP.在庫場所1 + '-' + ZP.在庫場所2 + '-' + SUBSTRING(ZP.棚番,1,2) + '-' + SUBSTRING(ZP.棚番,3,2) End 置場," & Chr(13)
    gSL_Select = gSL_Select & "       Isnull(ZP.在庫数,0) 理論在庫数,MT.取引先,ZP.材料管理番号" & Chr(13)
    gSL_Select = gSL_Select & "  from 材料棚卸プレシート発行テーブル ZP" & Chr(13)
    gSL_Select = gSL_Select & "  Left Outer Join 材料マスタ ZM" & Chr(13)
    gSL_Select = gSL_Select & "    On ZM.材料管理番号 = ZP.材料管理番号" & Chr(13)
    gSL_Select = gSL_Select & "  Left Outer Join (Select wk.材料管理番号,substring(wk.最小取引先,1,7) 取引先 " & Chr(13)
    gSL_Select = gSL_Select & "                     from (Select 材料管理番号,Min(Left(取引先 + '       ',7) + 受入) 最小取引先 from 材料取引先マスタ" & Chr(13)
    gSL_Select = gSL_Select & "                            Group by 材料管理番号) WK" & Chr(13)
    gSL_Select = gSL_Select & "                   ) MT"
    gSL_Select = gSL_Select & "    On MT.材料管理番号 = ZP.材料管理番号" & Chr(13)
    
    
    gSL_Select = gSL_Select & " Where ZP.発行年月 = '" & iH1_YMD.cYear & iH1_YMD.cMonth & "' " & Chr(13)
    gSL_Select = gSL_Select & "   And ZP.棚卸入力済フラグ = 0 " & Chr(13)
    
    
    '棚卸ﾌﾟﾚｼｰﾄ№
    If RTrim(iH2_No) <> "" Then
        gSL_Select = gSL_Select & "  And ZP.伝票NO = " & iH2_No & ""
    End If
    '材質
    If RTrim(iH2_Zaishitsu) <> "" Then
        gSL_Select = gSL_Select & "  And ZM.材質 = '" & iH2_Zaishitsu & "'"
    End If
    '板厚
    If RTrim(iH2_Itaatsu) <> "" Then
        gSL_Select = gSL_Select & "  And ZM.板厚 = '" & iH2_Itaatsu & "'"
    End If
    '幅
    If RTrim(iH2_Haba) <> "" Then
        gSL_Select = gSL_Select & "  And ZM.幅 = '" & iH2_Haba & "'"
    End If
    '長さ
    If RTrim(iH2_Nagasa) <> "" Then
        gSL_Select = gSL_Select & "  And ZM.長さ = '" & iH2_Nagasa & "'"
    End If

End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    iB1_TSuryo.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub

'テキストからグリッドに退避
    Call TextToVsGrid(Val(iB1_TSuryo.Tag))
    gInt = VsGrid1.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
    
        If gInt + 1 > VsGrid1.Rows - 1 Then
                GoTo Grid_RowChange_End
        Else
            If VsGrid1.TextMatrix(VsGrid1.Row + 1, 0) = "" Then GoTo Grid_RowChange_End
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
    
        If CisDB.RecordCount < CisVsGrid.DispMax - 1 Then
            gInt = CisDB.RecordCount
            VsGrid1.Row = gInt
            Call VsGridToText(VsGrid1.Row)
            Dummy.Enabled = True
            Dummy.SetFocus
            iB1_TSuryo.SetFocus
            Dummy.Enabled = False
            GoTo Grid_RowChange_End
        End If
    
        If gInt + CisVsGrid.DispMax > VsGrid1.Rows - 1 Then
            gInt = VsGrid1.Rows - 1
        Else
            gInt = gInt + CisVsGrid.DispMax
        End If
    End If
    
    VsGrid1.Row = gInt
    
'グリッドからテキストに表示
    Call VsGridToText(VsGrid1.Row)
    Dummy.Enabled = True
    Dummy.SetFocus
    iB1_TSuryo.SetFocus
    Dummy.Enabled = False
    
Grid_RowChange_End:
    KeyCode = 0
'    Me.Refresh
    DoEvents
End Sub
Private Sub iB1_TSuryo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    
    With VsGrid1
    
    '棚卸数量
        If iB1_TSuryo = "" Then
            .TextMatrix(tRow, 6) = ""
        Else
            .TextMatrix(tRow, 6) = Format(CisFun.Val2(iB1_TSuryo), "#,##0")
        End If
        
        If RTrim(.TextMatrix(tRow, 6)) = "" Then
            If RTrim(.TextMatrix(tRow, 10)) = "X" Then
                VsGrid1.Row = tRow
                Call CisVsGrid.vsColor(消去個別)
            End If
        Else
            If RTrim(.TextMatrix(tRow, 10)) = "X" Then
            Else
                VsGrid1.Row = tRow
                Call CisVsGrid.vsColor(個別, &HC0E0FF)
            End If
        End If

    End With
    
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With VsGrid1
        iB1_TSuryo.Tag = tRow
        iB1_TSuryo.Move iB1_TSuryo.Left, VsGrid1.CellTop, iB1_TSuryo.Width, VsGrid1.CellHeight
        '棚卸数量
        If RTrim(.TextMatrix(tRow, 6)) = "" Then
            iB1_TSuryo = ""
        Else
            iB1_TSuryo = CisFun.Val2(.TextMatrix(tRow, 6))
        End If
        DoEvents
    End With
    
End Sub
Private Sub VSGrid1_Click()
    Call TextToVsGrid(Val(iB1_TSuryo.Tag))
    If VsGrid1.TextMatrix(VsGrid1.Row, 0) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(VsGrid1.Row)
    On Error Resume Next
    iB1_TSuryo.SetFocus
    SendKeys "{LEFT}"
    On Error GoTo 0
End Sub
