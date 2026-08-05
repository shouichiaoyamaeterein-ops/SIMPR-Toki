VERSION 5.00
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form COS0430 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "受給差異明細メンテナンス"
   ClientHeight    =   11145
   ClientLeft      =   15
   ClientTop       =   360
   ClientWidth     =   15300
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11145
   ScaleWidth      =   15300
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'なし
      Height          =   390
      Left            =   5940
      ScaleHeight     =   390
      ScaleWidth      =   5865
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   30
      Width           =   5865
      Begin Cis3D_v60.CIS3D Sentaku 
         Height          =   345
         Left            =   30
         Top             =   30
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BackColor       =   14737632
         Caption         =   "選択         件"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_SKensu 
            Height          =   285
            Left            =   600
            Top             =   30
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
            BackColor       =   16761024
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   2
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -20
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   345
         Left            =   1965
         Top             =   30
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BackColor       =   14737632
         Caption         =   "表示         件"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   285
            Left            =   600
            Top             =   30
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
            BackColor       =   16761024
            ForeColor       =   0
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   2
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -20
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   345
         Left            =   3900
         Top             =   30
         Width           =   1935
         _ExtentX        =   3413
         _ExtentY        =   609
         BackColor       =   16761087
         Caption         =   "削除         件"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D lb_DelCnt 
            Height          =   285
            Left            =   600
            Top             =   30
            Width           =   945
            _ExtentX        =   1667
            _ExtentY        =   503
            BackColor       =   16761024
            ForeColor       =   0
            Caption         =   "999,999"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   8421504
            cBoderColor2    =   -2147483643
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   2
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -20
         End
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   1320
      Left            =   11520
      ScaleHeight     =   1260
      ScaleWidth      =   3540
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   510
      Width           =   3600
      Begin CisBtn_60.CisBtn PB_Mei 
         Height          =   585
         Left            =   45
         Top             =   630
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
         BtnForeColor    =   0
         BtnShadowColor  =   16777215
         PFCaption       =   "( F1 )"
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
         PFBackColor     =   0
      End
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   2610
         Top             =   30
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
         Left            =   1770
         Top             =   30
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
         Left            =   915
         Top             =   30
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
      Begin CisBtn_60.CisBtn PB_PDEL 
         Height          =   585
         Left            =   1785
         Top             =   630
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
         BtnCaption      =   "完削除"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   255
         BtnShadowColor  =   16777215
         PFCaption       =   "( F3 )"
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
         PFBackColor     =   255
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   8955
      Left            =   135
      ScaleHeight     =   8895
      ScaleWidth      =   14925
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2010
      Width           =   14985
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   8760
         Left            =   60
         Top             =   90
         Width           =   14745
         _ExtentX        =   26009
         _ExtentY        =   15452
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
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin vsFlexLib.vsFlexArray VSGRID1 
            DragIcon        =   "COS0430.frx":0000
            Height          =   8640
            Left            =   60
            TabIndex        =   6
            Top             =   30
            Width           =   14655
            _Version        =   196608
            _ExtentX        =   25850
            _ExtentY        =   15240
            _StockProps     =   228
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   $"COS0430.frx":0442
            Rows            =   26
            Cols            =   24
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   330
            SelectionMode   =   1
            AllowUserResizing=   1
            MouseIcon       =   "COS0430.frx":050D
            BackColorAlternate=   12648447
            Begin Cis3D_v60.CIS3D S1_Area1 
               Height          =   4335
               Left            =   3660
               Top             =   690
               Width           =   10305
               _ExtentX        =   18177
               _ExtentY        =   7646
               BackColor       =   0
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
               Begin Cis3D_v60.CIS3D CIS3D1 
                  Height          =   405
                  Left            =   180
                  Top             =   2400
                  Width           =   10005
                  _ExtentX        =   17648
                  _ExtentY        =   714
                  BackColor       =   65535
                  Caption         =   "【　購　入　明　細　】"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   15.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cAlingnment     =   7
               End
               Begin vsFlexLib.vsFlexArray VsGrid2 
                  Height          =   1410
                  Left            =   180
                  TabIndex        =   12
                  TabStop         =   0   'False
                  Top             =   2790
                  Width           =   9960
                  _Version        =   196608
                  _ExtentX        =   17568
                  _ExtentY        =   2487
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
                  FormatString    =   "No|日付      | 数量|納入先|受入|売上年月|売上処理日|"
                  Rows            =   4
                  Cols            =   8
                  BackColor       =   16777152
                  BackColorBkg    =   14737632
                  FocusRect       =   0
                  HighLight       =   0
                  ScrollBars      =   2
                  RowHeightMin    =   330
                  SelectionMode   =   1
                  BackColorAlternate=   16777152
               End
               Begin Cis3D_v60.CIS3D CIS3D4 
                  Height          =   4170
                  Left            =   90
                  Top             =   90
                  Width           =   10125
                  _ExtentX        =   17859
                  _ExtentY        =   7355
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
                  Begin vsFlexLib.vsFlexArray VsGrid3 
                     Height          =   1485
                     Left            =   4470
                     TabIndex        =   17
                     TabStop         =   0   'False
                     Top             =   30
                     Width           =   5580
                     _Version        =   196608
                     _ExtentX        =   9842
                     _ExtentY        =   2619
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
                     FormatString    =   "   |数量      | 単価| 金額|"
                     Rows            =   4
                     Cols            =   5
                     BackColor       =   16777152
                     BackColorBkg    =   14737632
                     FocusRect       =   0
                     HighLight       =   0
                     ScrollBars      =   0
                     RowHeightMin    =   330
                     SelectionMode   =   1
                     BackColorAlternate=   16777152
                  End
                  Begin Cis3D_v60.CIS3D CIS3D7 
                     Height          =   375
                     Left            =   60
                     Top             =   1530
                     Width           =   10005
                     _ExtentX        =   17648
                     _ExtentY        =   661
                     BackColor       =   0
                     ForeColor       =   65535
                     Caption         =   "【 買 掛 金 情 報 】"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   15.75
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cFont3DColor1   =   16777215
                     cFont3DColor2   =   8421504
                     cFont3D         =   2
                     cAlingnment     =   7
                  End
                  Begin VB.Label B2lb_Tei 
                     BackStyle       =   0  '透明
                     Caption         =   "XX"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   9510
                     TabIndex        =   31
                     Top             =   2010
                     Width           =   240
                  End
                  Begin VB.Shape Shape17 
                     BackColor       =   &H00C0FFFF&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   9150
                     Shape           =   4  '丸みのある長方形
                     Top             =   1920
                     Width           =   915
                  End
                  Begin VB.Label Label9 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "訂正"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   8580
                     TabIndex        =   30
                     Top             =   2010
                     Width           =   480
                  End
                  Begin VB.Label B2lb_Ukeir 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "XXXX"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   7860
                     TabIndex        =   29
                     Top             =   2010
                     Width           =   480
                  End
                  Begin VB.Shape Shape15 
                     BackColor       =   &H00C0FFFF&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   7740
                     Shape           =   4  '丸みのある長方形
                     Top             =   1920
                     Width           =   735
                  End
                  Begin VB.Label Label8 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "受入"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   7170
                     TabIndex        =   28
                     Top             =   2010
                     Width           =   480
                  End
                  Begin VB.Label B2lb_Ymd 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "XXXXXXXXX11"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   5700
                     TabIndex        =   27
                     Top             =   2010
                     Width           =   1320
                  End
                  Begin VB.Shape Shape13 
                     BackColor       =   &H00C0FFFF&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   5550
                     Shape           =   4  '丸みのある長方形
                     Top             =   1920
                     Width           =   1545
                  End
                  Begin VB.Label Label7 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "日  付"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   4650
                     TabIndex        =   26
                     Top             =   2010
                     Width           =   720
                  End
                  Begin VB.Label B2lb_Shiki 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "XXXXXXXXXXXXXXXXXXXXX"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   1320
                     TabIndex        =   25
                     Top             =   2010
                     Width           =   3030
                  End
                  Begin VB.Label Label5 
                     BackStyle       =   0  '透明
                     Caption         =   "識別"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   330
                     TabIndex        =   24
                     Top             =   2010
                     Width           =   555
                  End
                  Begin VB.Label Label6 
                     BackStyle       =   0  '透明
                     Caption         =   "契約先"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   240
                     TabIndex        =   23
                     Top             =   1230
                     Width           =   735
                  End
                  Begin VB.Label B2lb_SYM 
                     BackStyle       =   0  '透明
                     Caption         =   "XXXX/XX"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   3540
                     TabIndex        =   22
                     Top             =   870
                     Width           =   885
                  End
                  Begin VB.Label Label2 
                     BackStyle       =   0  '透明
                     Caption         =   "照合年月"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   2490
                     TabIndex        =   21
                     Top             =   870
                     Width           =   960
                  End
                  Begin VB.Label Label4 
                     BackStyle       =   0  '透明
                     Caption         =   "納入№"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   240
                     TabIndex        =   20
                     Top             =   870
                     Width           =   735
                  End
                  Begin VB.Label B2lb_Nouno 
                     BackStyle       =   0  '透明
                     Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   1230
                     TabIndex        =   19
                     Top             =   870
                     Width           =   1185
                  End
                  Begin VB.Label B2lb_Hinbn 
                     BackStyle       =   0  '透明
                     Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXXX"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   1260
                     TabIndex        =   18
                     Top             =   510
                     Width           =   3105
                  End
                  Begin VB.Label Label1 
                     BackStyle       =   0  '透明
                     Caption         =   "照合結果"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   150
                     TabIndex        =   16
                     Top             =   120
                     Width           =   1080
                  End
                  Begin VB.Label B2lb_Syogo 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  '透明
                     Caption         =   "XXXXXXXXXXXXXXXXXXXXX"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   240
                     Left            =   1245
                     TabIndex        =   15
                     Top             =   120
                     Width           =   2520
                  End
                  Begin VB.Label Label3 
                     BackStyle       =   0  '透明
                     Caption         =   "品  番"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   240
                     TabIndex        =   14
                     Top             =   510
                     Width           =   735
                  End
                  Begin VB.Label B2lb_Torcd 
                     BackStyle       =   0  '透明
                     Caption         =   "XXXXX：ＸＸＸＸＸＸＸＸＸＸＸ"
                     BeginProperty Font 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   12
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   225
                     Left            =   1230
                     TabIndex        =   13
                     Top             =   1230
                     Width           =   3855
                  End
                  Begin VB.Shape Shape1 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   60
                     Shape           =   4  '丸みのある長方形
                     Top             =   45
                     Width           =   1125
                  End
                  Begin VB.Shape Shape2 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   1155
                     Shape           =   4  '丸みのある長方形
                     Top             =   45
                     Width           =   3300
                  End
                  Begin VB.Shape Shape3 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   60
                     Shape           =   4  '丸みのある長方形
                     Top             =   420
                     Width           =   1125
                  End
                  Begin VB.Shape Shape4 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   1170
                     Shape           =   4  '丸みのある長方形
                     Top             =   420
                     Width           =   3285
                  End
                  Begin VB.Shape Shape6 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   1170
                     Shape           =   4  '丸みのある長方形
                     Top             =   780
                     Width           =   1275
                  End
                  Begin VB.Shape Shape5 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   60
                     Shape           =   4  '丸みのある長方形
                     Top             =   780
                     Width           =   1125
                  End
                  Begin VB.Shape Shape8 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   2340
                     Shape           =   4  '丸みのある長方形
                     Top             =   780
                     Width           =   1125
                  End
                  Begin VB.Shape Shape7 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   3450
                     Shape           =   4  '丸みのある長方形
                     Top             =   780
                     Width           =   1005
                  End
                  Begin VB.Shape Shape10 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   1170
                     Shape           =   4  '丸みのある長方形
                     Top             =   1140
                     Width           =   3285
                  End
                  Begin VB.Shape Shape9 
                     BackColor       =   &H00C0FFC0&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   60
                     Shape           =   4  '丸みのある長方形
                     Top             =   1140
                     Width           =   1125
                  End
                  Begin VB.Shape Shape12 
                     BackColor       =   &H00C0FFFF&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   60
                     Shape           =   4  '丸みのある長方形
                     Top             =   1920
                     Width           =   1125
                  End
                  Begin VB.Shape Shape11 
                     BackColor       =   &H00C0FFFF&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   1170
                     Shape           =   4  '丸みのある長方形
                     Top             =   1920
                     Width           =   3285
                  End
                  Begin VB.Shape Shape14 
                     BackColor       =   &H00C0FFFF&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   4440
                     Shape           =   4  '丸みのある長方形
                     Top             =   1920
                     Width           =   1125
                  End
                  Begin VB.Shape Shape16 
                     BackColor       =   &H00C0FFFF&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   7080
                     Shape           =   4  '丸みのある長方形
                     Top             =   1920
                     Width           =   675
                  End
                  Begin VB.Shape Shape18 
                     BackColor       =   &H00C0FFFF&
                     BackStyle       =   1  '不透明
                     Height          =   390
                     Left            =   8460
                     Shape           =   4  '丸みのある長方形
                     Top             =   1920
                     Width           =   705
                  End
               End
            End
            Begin Cis3D_v60.CIS3D PNL_Wait 
               Height          =   930
               Left            =   5160
               Top             =   2370
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
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   1335
      Left            =   165
      ScaleHeight     =   1327.042
      ScaleMode       =   0  'ﾕｰｻﾞｰ
      ScaleWidth      =   11205
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   510
      Width           =   11265
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   1170
         Left            =   4185
         Top             =   60
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   2064
         ForeColor       =   16711680
         Caption         =   "品   番"
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
         cPositionY      =   110
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   60
            TabIndex        =   3
            Top             =   540
            Width           =   3195
            _ExtentX        =   5636
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
      Begin Cis3D_v60.CIS3D CIS3D6 
         Height          =   1170
         Left            =   60
         Top             =   60
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   2064
         ForeColor       =   16711680
         Caption         =   "照合結果"
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
         cPositionY      =   110
         Begin VB.ComboBox CombH1_Syogo 
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
            ItemData        =   "COS0430.frx":095F
            Left            =   300
            List            =   "COS0430.frx":0966
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   1
            Top             =   525
            Width           =   2175
         End
         Begin CisText_V60.CisText iH1_Syogo 
            Height          =   375
            Left            =   60
            TabIndex        =   0
            Top             =   525
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   661
            cFormat         =   "#"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "X"
            MaxLength       =   1
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   1170
         Left            =   7455
         Top             =   60
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   2064
         ForeColor       =   16711680
         Caption         =   "納入番号"
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
         cPositionY      =   110
         Begin CisText_V60.CisText iH1_Nouno 
            Height          =   375
            Left            =   150
            TabIndex        =   4
            Top             =   540
            Width           =   765
            _ExtentX        =   1349
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
            Text            =   "XXXXX"
            MaxLength       =   5
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D3 
         Height          =   1170
         Left            =   2595
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   2064
         ForeColor       =   16711680
         Caption         =   "照合年月"
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
         cPositionY      =   110
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   390
            Left            =   90
            TabIndex        =   2
            Top             =   525
            Width           =   1410
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
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   1170
         Left            =   8580
         Top             =   60
         Width           =   2565
         _ExtentX        =   4524
         _ExtentY        =   2064
         ForeColor       =   16711680
         Caption         =   "契 約 先"
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
         cPositionY      =   110
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   375
            Left            =   990
            Top             =   555
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   661
            BackColor       =   16777152
            Caption         =   "XXXXXXXXXX"
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   45
            TabIndex        =   5
            Top             =   555
            Width           =   975
            _ExtentX        =   1720
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
            Text            =   "XXXXXXX"
            MaxLength       =   7
         End
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   630
      Top             =   630
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   60
      TabIndex        =   8
      Text            =   "D"
      Top             =   600
      Width           =   255
   End
   Begin Cis3D_v60.CIS3D UC_3D31 
      Height          =   435
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   767
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "　　　【 受給差異明細メンテナンス 】"
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
End
Attribute VB_Name = "COS0430"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  受給差異明細メンテナンス
'**       フォームID    :  COS0430
'**       処理概要      :
'**
'**       作  成  日    :  2006/09/01  By CIS
'**       変  更  日    :  2009/06/10  By CIS  受給材料対応
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim EndYM           As String       ' 最終照合年月
    Dim SelCnt          As Long
    
    Const Msg_Delete    As String = "削除"
    
    Dim mCmbObj         As CisCombo    'ComboProcクラス定義
    Dim CisVsGrid       As New CisVsGrid3
    Dim CisVsGridT      As New CisVsGrid3
    Dim CisVsGridU      As New CisVsGrid3
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
        Set .GridObj = VSGRID1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .SelectCol = 23
        .SortAsc = flexSortUseColSort
        Set .SelectDispObj = B1lb_SKensu
        .SelectDispFmt = "#,###"
        .InitGet = False
        .Init
    End With
    
    Call GridInit
    '
    With CisVsGridT
        Set .GridObj = VsGrid3
        .NoSpaceCol = 1
        .NumberSet = No
        .FixedHeight = 420
        .PatanMax = 1
        .InitGet = False
        .Init
    End With
    
    Call GridInitT
    
    With CisVsGridU
        Set .GridObj = VsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 350
        .PatanMax = 1
        .InitGet = False
        .Init
    End With
    
    Call GridInitU
    
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
    
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
'
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

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
' 照合年月　初期表示
    gSL_Select = "SELECT MAX(照合年月) 照合年月 FROM 受給差異明細テーブル "
    If JSTRead(gSL_Select, 1) Then
       iH1_YM = JST.照合年月
    End If

''ｺﾝﾎﾞﾎﾞｯｸｽ設定
    Call Get_Code
''+---------------------+
''+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
''+---------------------+
    Set mCmbObj = New CisCombo
    With mCmbObj
        Set .Connect = CisDB
        Set .CombObj = CombH1_Syogo         'ｺﾝﾎﾞﾎﾞｯｸｽｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .IniPath = gIniExe
        .Control = "名称マスタ"
        .ComboName = "照合結果受給"
        .SetSpace = 50                  '空白文字数
        .Widht = 0                      'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ幅(省略時、変更なし)
        .Ess = False                    '空白行(True:必須入力用,False:任意入力用)
        Call .ComboInit
        .ComboSet
    End With
'
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 ＫｅｙＤｏｗｎイベント 》               +
'+----------------------------------------------------------+
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If Key_Acc(KeyCode, "Down", Shift) Then KeyCode = 0
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Cancel = 1
    Call PB_END_Click
End Sub
'+----------------------------------------------------------+
'+      フォーム《 ファンクションキー用アクセラレータ 》       +
'+----------------------------------------------------------+
Private Function Key_Acc(Key_Code As Integer, IVENT As String, Shift As Integer)
    Key_Acc = False
    If Not FormAct Then Exit Function
    If Not PNL_PF.Enabled Then Exit Function
    If UCase(IVENT) <> "DOWN" Then Exit Function
    
    Select Case Key_Code
        Case vbKeyF1:       If PB_Mei.Visible Then Call PB_Mei_Click        '【明細】
        Case vbKeyF3:       If PB_PDEL.Visible Then Call PB_PDEL_Click      '【完全削除】
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      '【検索】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
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
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
'
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
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        GoTo PB_CAN_Ed
    Else
        Call HeadBodyClear("H")
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    明細情報表示(F1)        +
'+----------------------------+
Private Sub PB_Mei_Click()
    If S1_Area1.Visible Then
        S1_Area1.Visible = False
    Else
        S1_Area1.Visible = True
        Call GridSet2
    End If
End Sub
'+----------------------------+
'+    完全削除  キ ー(F3)     +
'+----------------------------+
Private Sub PB_PDEL_Click()
    If ProcHB = "B1" Then
       Call CisFrm.MousePT(11)
       Call DBPut(1)
       Call CisFrm.MousePT(1)
       Call PB_CAN_Click
    End If
End Sub
'+----------------------------+
'+    検索   (F7)             +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Torcd_Look: Exit Sub
End Sub
'+---------------------+
'+    契約先検索        +
'+---------------------+
Private Sub Torcd_Look()

    RV_Left = 0
    RV_Top = 0

    RVI_Kubun = "照合対象受給契約先"
    
    CKK0010.Show vbModal
    Unload CKK0010
    Set CKK0010 = Nothing

    iH1_Torcd.SetFocus

    If RV_Rtn Then
       iH1_Torcd = RV_Code
       H1lb_Tornm = RV_Kubun
    End If

End Sub
'+----------------------------------------+
'+  iH1_Syogo ( LostFocus )  照合結果
'+----------------------------------------+
Private Sub iH1_Syogo_LostFocus()
    With mCmbObj
        .Code = iH1_Syogo
        Call .TextGet
    End With
End Sub
'+---------------------+
'+    CombH1_SYOGO     +
'+---------------------+
Private Sub CombH1_Syogo_GotFocus()
    CombH1_Syogo.BackColor = gIGotSel
End Sub
Private Sub CombH1_Syogo_LostFocus()
    CombH1_Syogo.BackColor = gILostSel
End Sub
Private Sub CombH1_Syogo_Click()
    With mCmbObj
        Call .CodeGet
        iH1_Syogo = .Code
    End With
End Sub
'+---------------------+
'+    iH1_Torcd        +
'+---------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Visible = True
    PB_Look.Tag = ActiveControl.Name
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
       PB_Look.Visible = False
       PB_Look.Tag = ""
    End If
    
    Call TorGet(iH1_Torcd)
End Sub
Private Function TorGet(Torcd As String) As Boolean
    TorGet = True
    H1lb_Tornm = ""
    If RTrim(Torcd) = "" Then Exit Function
    gSL_Select = "SELECT 値名称,英数字1 "
    gSL_Select = gSL_Select & " FROM 名称マスタ"
    gSL_Select = gSL_Select & " WHERE 区分名称 = '照合対象受給契約先'"
    gSL_Select = gSL_Select & "   And 値 = '" & RTrim(Torcd) & "'"
    If SYMRead(gSL_Select, 1) Then
        H1lb_Tornm = SYM.値名称
    Else
        TorGet = False
    End If

End Function

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
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
    Dummy.Enabled = True
    Dummy.SetFocus
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

    PB_END.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_Mei.Visible = B1Mode
    PB_PDEL.Visible = B1Mode
    S1_Area1.Visible = False
    
    If ProcHB = "H1" Then
       B1lb_Kensu = ""
       B1lb_SKensu = ""
    End If
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
       If Not Head1Chk Then GoTo ReturnPress_Ed
       ProcHB = "B1"
       Call DispChange(ProcHB)
       VSGRID1.SetFocus
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
       If Not DBPut(0) Then GoTo ReturnPress_Ed
       Call PB_CAN_Click
    End If
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
    ' 照合年月
    If Trim(iH1_YM) <> "" Then
        If Not iH1_YM.cChkResult Then
           With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    正しい年月にて入力して下さい。   "
                .MB_Button = Error
                Call .MBOX
           End With
           iH1_YM.SetFocus
           Exit Function
        End If
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
                '.Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
        Call CisVsGrid.Clear
    End If
End Sub
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
    Dim Ttl_Kensu   As Integer
    Dim SL_SelectW      As String       ' SQL SELECT(Where)条件文
 
' 抽出処理 初期ｾｯﾄ
    
    PNL_PF.Enabled = False      ' ﾌｧﾝｸｼｮﾝ使用禁止
    PNL_Wait.Visible = True      ' 抽出中ﾒｯｾｰｼﾞ表示
    B1lb_Kensu = ""
    SL_SelectW = ""
    DoEvents
    
    gSL_Select = "SELECT JS.*,HN.表示品番,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(JS.照合赤黒,0) = 0 THEN '' ELSE '赤' END 赤黒名,"
    gSL_Select = gSL_Select & " ISNULL(SY.値名称,'') 照合結果名, "
    gSL_Select = gSL_Select & " ISNULL(SY2.値名称,'') 契約先名, "
    gSL_Select = gSL_Select & " ISNULL(SY3.値名称,'') 識別子名 "
    gSL_Select = gSL_Select & " FROM 受給差異明細テーブル JS "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "      ON HN.品番 = JS.品番 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "      ON  SY.区分名称 = '照合結果受給'"
    gSL_Select = gSL_Select & "      AND SY.区分タイプ = '' "
    gSL_Select = gSL_Select & "      AND SY.値 = JS.照合結果 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY2"
    gSL_Select = gSL_Select & "      ON  SY2.区分名称 = '照合対象受給契約先'"
    gSL_Select = gSL_Select & "      AND SY2.区分タイプ = '' "
    gSL_Select = gSL_Select & "      AND SY2.値 = JS.契約先 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY3"
    gSL_Select = gSL_Select & "      ON  SY3.区分名称 = 'TEXT買掛金'"
    gSL_Select = gSL_Select & "      AND SY3.区分タイプ = '' "
    gSL_Select = gSL_Select & "      AND SY3.値 = JS.契約先 + JS.FCD + JS.識別子 "
    
' 表示条件 セット
    With CisFun
        .Where_CreateKB = NewDocument
        ' 削除区分
        .Where_Create AlphaMe, "JS.削除区分", "", , Create
        ' 照合結果
        .Where_Create AlphaMe, "JS.照合結果", iH1_Syogo
        ' 照合年月
        .Where_Create AlphaMe, "JS.照合年月", iH1_YM
        ' 契約先
        .Where_Create AlphaMe, "JS.契約先", iH1_Torcd
        ' 品番
        .Where_Create AlphaMe, "JS.品番", iH1_Hinbn
        ' 納入番号
        .Where_Create AlphaMe, "JS.納入番号", iH1_Nouno
        
        SL_SelectW = .Where_Phrase
    End With
    
    gSL_Select = gSL_Select & SL_SelectW & " ORDER BY JS.照合結果,JS.品番,JS.納入日,JS.納入番号 "

    If Not JSTRead(gSL_Select) Then GoTo BodySet_ED
    
    Call GridSet
    BodySet = True
BodySet_ED:
    Call JSTClose
    PNL_PF.Enabled = True      ' ﾌｧﾝｸｼｮﾝ使用禁止 解除
    PNL_Wait.Visible = False   ' 抽出中ﾒｯｾｰｼﾞ解除
    Exit Function
BodySet_Err:
    CisFun.ErrorBox
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
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 4, 0, "No")
        Call .FixedSet(中中, 左中, あり, 18, 1, "照合結果", 英字)
        Call .FixedSet(中中, 左中, あり, 25, 2, "品番", 英字)
        Call .FixedSet(中中, 中中, あり, 11, 3, "納入日", 英字)
        Call .FixedSet(中中, 左中, あり, 6, 4, "納入№", 英字)
        Call .FixedSet(中中, 右中, あり, 0, 5, "契約先単価", 数字)
        Call .FixedSet(中中, 右中, あり, 0, 6, "契約先数量", 数字)
        Call .FixedSet(中中, 右中, あり, 0, 7, "契約先金額", 数字)
        Call .FixedSet(中中, 右中, あり, 0, 8, "自社単価", 数字)
        Call .FixedSet(中中, 右中, あり, 0, 9, "自社数量", 数字)
        Call .FixedSet(中中, 右中, あり, 0, 10, "自社金額", 数字)
        Call .FixedSet(中中, 中中, あり, 0, 11, "照合年月", 数字)
        Call .FixedSet(中中, 左中, あり, 30, 12, "識 別", 英字)
        Call .FixedSet(中中, 左中, あり, 0, 13, "照合品番K")
        Call .FixedSet(中中, 左中, あり, 0, 14, "照合年月K")
        Call .FixedSet(中中, 左中, あり, 0, 15, "照合納番K")
        Call .FixedSet(中中, 左中, あり, 0, 16, "照合管理K")
        Call .FixedSet(中中, 左中, あり, 0, 17, "財務管理NO")
        Call .FixedSet(中中, 左中, あり, 0, 18, "契約先")
        Call .FixedSet(中中, 左中, あり, 30, 19, "契約先名", 英字)
        
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 20, "Select")

        Call .FixedSet(中中, 中中, あり, 4, 21, "赤黒", 英字)
        Call .FixedSet(中中, 中中, あり, 0, 22, "赤黒K", 数字)
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,19,21")
'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
    Dim wMaxGyo     As Long
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    wMaxGyo = 5000
    If Ttl_Kensu > wMaxGyo Then
       Ttl_Kensu = wMaxGyo
       With CisFun
            .MB_Lines = 5
            .MB_MSG(1) = "    最大表示行を超えました（" & StrConv(Format(wMaxGyo, "#,##0"), vbWide) & ")      "
            .MB_MSG(3) = "        表示可能行数を表示します。              "
            .MB_Title = "最大表示確認"
            .MB_Button = OK
            .MBOX
       End With
    End If
    B1lb_Kensu = Format(Ttl_Kensu, "#,##0")
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VSGRID1
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGrid.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
            .Rows = CisVsGrid.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    With VSGRID1
        Do Until Not JST_RDSTS
           ID = ID + 1
           If ID > wMaxGyo Then Exit Do
           '照合結果
           If Trim(JST.照合結果名) = "" Then
               .TextMatrix(ID, GridCol(1)) = JST.照合結果
           Else
               .TextMatrix(ID, GridCol(1)) = JST.照合結果 & "：" & JST.照合結果名
           End If
           ' 品番
           If Trim(JST.表示品番) = "" Then
               .TextMatrix(ID, GridCol(2)) = JST.品番
           Else
               .TextMatrix(ID, GridCol(2)) = JST.表示品番
           End If
           ' 納入日
           If Len(JST.納入日) < 7 Then
              gStr = Mid(JST.納入日, 1, 4) & "/" & Mid(JST.納入日, 5, 2)
           Else
              With CisFun
                  .DateE_BefVal = JST.納入日
                  .DateEdit UserDeli, "/"
                  gStr = .DateE_AfVal
              End With
           End If
           .TextMatrix(ID, GridCol(3)) = gStr
           ' 納入番号
           .TextMatrix(ID, GridCol(4)) = JST.納入番号
           
           ' 契約先情報
           If JST.財務管理番号 <> 0 Then
              .TextMatrix(ID, GridCol(5)) = Format(JST.契約先単価, "#,##0.00")
              .TextMatrix(ID, GridCol(6)) = Format(JST.契約先数量, "#,##0")
              .TextMatrix(ID, GridCol(7)) = Format(JST.契約先金額, "#,##0")
           End If
           ' 自社情報
           If Trim(JST.照合日付) <> "" Then
              .TextMatrix(ID, GridCol(8)) = Format(JST.自社単価, "#,##0.00")
              .TextMatrix(ID, GridCol(9)) = Format(JST.自社数量, "#,##0")
              .TextMatrix(ID, GridCol(10)) = Format(JST.自社金額, "#,##0")
           End If
           ' 照合年月
            .TextMatrix(ID, GridCol(11)) = Mid(JST.照合年月, 1, 4) & "/" & Mid(JST.照合年月, 5, 2)

           ' 識別
           If Trim(JST.識別子名) = "" Then
               .TextMatrix(ID, GridCol(12)) = JST.識別子
           Else
               .TextMatrix(ID, GridCol(12)) = JST.識別子 & "：" & JST.識別子名
           End If
           ' 契約先
           If Trim(JST.契約先名) = "" Then
               .TextMatrix(ID, GridCol(19)) = JST.契約先
           Else
               .TextMatrix(ID, GridCol(19)) = JST.契約先 & "：" & JST.契約先名
           End If
           ' 更新用　情報
           .TextMatrix(ID, GridCol(13)) = JST.照合品番
           .TextMatrix(ID, GridCol(14)) = JST.照合日付
           .TextMatrix(ID, GridCol(15)) = JST.照合納番
           .TextMatrix(ID, GridCol(16)) = JST.照合管理
           .TextMatrix(ID, GridCol(17)) = JST.財務管理番号
           .TextMatrix(ID, GridCol(18)) = JST.契約先
        
           .TextMatrix(ID, GridCol(21)) = JST.赤黒名
           .TextMatrix(ID, GridCol(22)) = JST.照合赤黒
           
           Call JSTReadNext
        Loop
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
    Call CisVsGrid.VSSort
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGRID1_Click()
    If S1_Area1.Visible Then S1_Area1.Visible = False
    Call CisVsGrid.vsColor(, &HFFC0FF)
End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    Call VSGRID1_Click
End Sub
Private Sub vsGrid1_RowColChange()
    Call GridSet2
End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With VSGRID1
        .Tag = ""
    
        If Button <> vbLeftButton Then Exit Sub
        If .MouseRow <> 0 Or .MouseCol = 0 Then Exit Sub
        
        If S1_Area1.Visible Then S1_Area1.Visible = False

    
        .Tag = Str(.MouseCol)
        .Drag vbBeginDrag
    End With
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)
    With VSGRID1
        If .MouseCol = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
    
        .Redraw = False
    
        .ColPosition(CisFun.Val2(.Tag)) = .MouseCol
        .Col = 1: .ColSel = .Cols - 1
    
        .Redraw = True
    
        CisVsGrid.VSSort
    End With
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut(SKB As Byte)
    DBPut = False

    If CisVsGrid.SelectCount = 0 Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "        削除する行を選択して下さい            "
            .MB_Title = "削除行選択"
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If
'  ( 確認メッセージ )
    If SKB = 0 Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "         【 差異明細削除処理確認 】            "
            .MB_MSG(4) = "              開始しますか？                   "
            .MB_Title = "削除確認"
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Function
        End With
    Else
       With CisFun
            .MB_Lines = 5
            .MB_MSG(1) = "        【 差異明細完全削除処理確認 】         "
            .MB_MSG(2) = "   　　 差異データを『完全削除』します         "
            .MB_MSG(4) = "              開始しますか？                   "
            .MB_Title = "削除確認"
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Function
        End With
    End If
'
    PNL_PF.Enabled = False      ' ﾌｧﾝｸｼｮﾝ使用禁止
    On Error GoTo DBPut_Err
    VSGRID1.SetFocus
'
    Call CisDB.DBTran(TransBegin)
'
    With VSGRID1
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, CisVsGrid.SelectCol)) = "X" Then
               '<< 削除処理  >>
               If Not Data_Delete(gLong, SKB) Then Exit Function

            End If
        Next gLong
    End With
    
    Call CisDB.DBTran(TransCommit)

    DBPut = True
    PNL_PF.Enabled = True      ' ﾌｧﾝｸｼｮﾝ使用可
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'*****************************************************************************
'* データ削除更新（DelKB 0:通常削除 1:完全削除）
'*****************************************************************************
Private Function Data_Delete(ID As Long, DelKB As Byte)
    Data_Delete = False
    
    Dim wYM     As String
    Dim wSyogo  As String
    Dim wTorcd  As String
    Dim wHinbn  As String
    Dim wNouno  As String
    Dim wYMD    As String
    Dim wNYMD   As String
    Dim wKanri  As String
    Dim wZaimu  As Long
    Dim wAKKbn  As Byte
        
    On Error GoTo Data_Delete_Err
    
    With VSGRID1
        wYM = Replace(.TextMatrix(ID, CisVsGrid.FixedGet(11)), "/", "")
        wSyogo = Mid(.TextMatrix(ID, CisVsGrid.FixedGet(1)), 1, 1)
        wTorcd = Trim(.TextMatrix(ID, CisVsGrid.FixedGet(18)))
        wYMD = Trim(.TextMatrix(ID, CisVsGrid.FixedGet(14)))
        wHinbn = Trim(.TextMatrix(ID, CisVsGrid.FixedGet(13)))
        wNouno = Trim(.TextMatrix(ID, CisVsGrid.FixedGet(15)))
        wKanri = Trim(.TextMatrix(ID, CisVsGrid.FixedGet(16)))
        wZaimu = CisFun.Val2(.TextMatrix(ID, CisVsGrid.FixedGet(17)))
        wNYMD = Replace(.TextMatrix(ID, CisVsGrid.FixedGet(3)), "/", "")
        wAKKbn = CisFun.Val2(.TextMatrix(ID, CisVsGrid.FixedGet(22)))
    End With
    
    If Len(wNYMD) > 6 Then
       wNYMD = Mid(wNYMD, 1, 6)
    End If
    
    With CisDB
        ' 受給差異明細テーブル 更新
        .SQL = "UPDATE 受給差異明細テーブル SET "
        .SQL = .SQL & " 削除区分 = 'D',"
        .SQL = .SQL & " 更新日 = GetDate(),"
        .SQL = .SQL & " 更新端末 = Host_Name()"
        .SQL = .SQL & " WHERE 照合年月 = '" & wYM & "'"
        .SQL = .SQL & "   AND 照合結果 = '" & wSyogo & "'"
        .SQL = .SQL & "   AND 契約先 = '" & wTorcd & "'"
        .SQL = .SQL & "   AND 照合日付 = '" & wYMD & "'"
        .SQL = .SQL & "   AND 照合品番 = '" & wHinbn & "'"
        .SQL = .SQL & "   AND 照合納番 = '" & wNouno & "'"
        .SQL = .SQL & "   AND 照合管理 = '" & wKanri & "'"
        .SQL = .SQL & "   AND 財務管理番号 = " & wZaimu
        .SQL = .SQL & "   AND 照合赤黒 = " & wAKKbn

        .DBExec
    
        ' 購入明細テーブル 更新
        If Trim(wYMD) <> "" Then
           .SQL = "Update 購入明細テーブル Set"
           If DelKB = 0 Then
              .SQL = .SQL & " 照合年月 = '" & wYM & "',"
           Else
              .SQL = .SQL & " 照合年月 = '" & Mid(wYMD, 1, 6) & "',"
           End If
           .SQL = .SQL & " 照合処理日 = CONVERT(VARCHAR(8),GETDATE(),112),"
           .SQL = .SQL & " 照合結果 = SUBSTRING(照合結果,1,1) + 'X',"
           .SQL = .SQL & " 更新日 = GETDATE(),"
           .SQL = .SQL & " 更新端末 = Host_Name()"
           '.SQL = .SQL & " FROM 購入明細テーブル (INDEX = IX_受給照合)"
           .SQL = .SQL & " FROM 購入明細テーブル"
           .SQL = .SQL & " Where 契約先 = '" & wTorcd & "'"
           
           .SQL = .SQL & "   AND 照合日付 = '" & wYMD & "'"
           .SQL = .SQL & "   AND 照合品番 = '" & wHinbn & "'"
           .SQL = .SQL & "   AND 照合納番 = '" & wNouno & "'"
           .SQL = .SQL & "   AND 照合管理 = '" & wKanri & "'"
           .SQL = .SQL & "   AND 照合赤黒 = " & wAKKbn
        
           .DBExec
        
           '======================================================== 2009/06/10 Start
           .SQL = "Update 材料購入明細テーブル Set"
           If DelKB = 0 Then
              .SQL = .SQL & " 照合年月 = '" & wYM & "',"
           Else
              .SQL = .SQL & " 照合年月 = '" & Mid(wYMD, 1, 6) & "',"
           End If
           .SQL = .SQL & " 照合処理日 = CONVERT(VARCHAR(8),GETDATE(),112),"
           .SQL = .SQL & " 照合結果 = SUBSTRING(照合結果,1,1) + 'X',"
           .SQL = .SQL & " 更新日 = GETDATE(),"
           .SQL = .SQL & " 更新端末 = Host_Name()"
           .SQL = .SQL & " FROM 材料購入明細テーブル"
           .SQL = .SQL & " Where 契約先 = '" & wTorcd & "'"
           .SQL = .SQL & "   AND 照合日付 = '" & wYMD & "'"
           .SQL = .SQL & "   AND 照合受給品番 = '" & wHinbn & "'"
           .SQL = .SQL & "   AND 照合納番 = '" & wNouno & "'"
           .SQL = .SQL & "   AND 照合管理 = '" & wKanri & "'"
           .SQL = .SQL & "   AND 照合赤黒 = " & wAKKbn
        
           .DBExec
           '======================================================== 2009/06/10 End
        End If
        ' TEXT買掛金テーブル
        If wZaimu <> 0 Then
            .SQL = " UPDATE TEXT買掛金テーブル SET "
            If DelKB = 0 Then
               .SQL = .SQL & " 照合年月 = '" & wYM & "',"
            Else
               .SQL = .SQL & " 照合年月 = '" & wNYMD & "',"
            End If
            .SQL = .SQL & " 照合結果 = SUBSTRING(照合結果,1,1) + 'X',"
            .SQL = .SQL & " 更新日 = GETDATE(),"
            .SQL = .SQL & " 更新端末 = Host_Name()"
            .SQL = .SQL & " WHERE 財務管理番号 = " & wZaimu
                    
            .DBExec
        End If
    End With
    
    lb_DelCnt.Caption = Format(CisFun.Val2(lb_DelCnt) + 1, "#,##0")
    DoEvents

    Data_Delete = True
    Exit Function
Data_Delete_Err:
    CisFun.ErrorBox ("削除処理")
    End
End Function
'+==========================================================+
'+      グリッド処理                                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInitT()
    With CisVsGridT
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 中中, あり, 10, 0, "          ")
        Call .FixedSet(中中, 右中, あり, 12, 1, "単価")
        Call .FixedSet(中中, 右中, あり, 9, 2, "数量")
        Call .FixedSet(中中, 右中, あり, 12, 3, "金額")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
        .InitDisp
    End With
End Sub
'+==========================================================+
'+      グリッド処理(支給情報)                              +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInitU()
    With CisVsGridU
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 0, 0, "No")
        Call .FixedSet(中中, 中中, あり, 11, 1, "日付")
        Call .FixedSet(中中, 右中, あり, 10, 2, "数量")
        Call .FixedSet(中中, 左中, あり, 31, 3, "仕入先")
        Call .FixedSet(中中, 中中, あり, 2, 4, "受")
        Call .FixedSet(中中, 中中, あり, 8, 5, "実績年月")
        Call .FixedSet(中中, 中中, あり, 11, 6, "実績処理日")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
        .InitDisp
    End With
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSetU()
    Dim ID          As Long
    Dim GridCol()   As Long
    Dim Ttl_Kensu   As Long
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With VsGrid2
       .Redraw = False
        ReDim GridCol(.Cols)
        For ID = 0 To .Cols - 1
            GridCol(ID) = CisVsGridU.FixedGet(ID)
        Next ID
        If Ttl_Kensu + 1 < CisVsGridU.DispMax Then
            .Rows = CisVsGridU.DispMax
        Else
            .Rows = Ttl_Kensu + 1
        End If
    End With
    ID = 0
    Do Until Not KMT_RDSTS
        With VsGrid2
            ID = ID + 1
           .TextMatrix(ID, GridCol(0)) = ID
            With CisFun
                 .DateE_BefVal = KMT.検収日
                 .DateEdit UserDeli, "/"
                 gStr = .DateE_AfVal
            End With
           .TextMatrix(ID, GridCol(1)) = gStr
           .TextMatrix(ID, GridCol(2)) = Format(KMT.数量, "#,###")
           If Trim(KMT.仕入先名) = "" Then
               .TextMatrix(ID, GridCol(3)) = KMT.仕入先
           Else
               .TextMatrix(ID, GridCol(3)) = KMT.仕入先 & "：" & KMT.仕入先名
           End If
           .TextMatrix(ID, GridCol(4)) = KMT.受入
           .TextMatrix(ID, GridCol(5)) = Mid(KMT.実績年月, 1, 4) & "/" & Mid(KMT.実績年月, 5, 2)
            With CisFun
                 .DateE_BefVal = KMT.実績処理日
                 .DateEdit UserDeli, "/"
                 gStr = .DateE_AfVal
            End With
           .TextMatrix(ID, GridCol(6)) = gStr
        
        
        End With
        Call KMTReadNext
    Loop
    With VsGrid2
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'+------------------------------------+
'+      明細内容表示
'+------------------------------------+
Private Sub GridSet2()
    Dim wCuren  As Currency
    Dim wYM     As String
    Dim wSyogo  As String
    Dim wTorcd  As String
    Dim wHinbn  As String
    Dim wNouno  As String
    Dim wYMD    As String
    Dim wKanri  As String
    Dim wZaimu  As Long
    Dim wAKKbn  As Byte
    
    If Not FormAct Or ProcHB <> "B1" Then Exit Sub
    
    If Not S1_Area1.Visible Then Exit Sub
    
    Call CisVsGridT.Clear
    Call CisVsGridU.Clear
    Call HeadBodyClear("B2")
    ' グリッド情報より表示
    With VSGRID1
        If .Row = 0 Then Exit Sub
        
        B2lb_Syogo = .TextMatrix(.Row, CisVsGrid.FixedGet(1))
        B2lb_Hinbn = .TextMatrix(.Row, CisVsGrid.FixedGet(2))
        B2lb_Nouno = .TextMatrix(.Row, CisVsGrid.FixedGet(4))
        B2lb_SYM = .TextMatrix(.Row, CisVsGrid.FixedGet(11))
        B2lb_Torcd = .TextMatrix(.Row, CisVsGrid.FixedGet(19))
        
        B2lb_Shiki = .TextMatrix(.Row, CisVsGrid.FixedGet(12))
        ' 数量・単価・金額
        VsGrid3.TextMatrix(1, 0) = "契約先情報"
        VsGrid3.TextMatrix(2, 0) = "自社情報"
        VsGrid3.TextMatrix(3, 0) = "差"
        VsGrid3.TextMatrix(1, 1) = .TextMatrix(.Row, CisVsGrid.FixedGet(5))
        VsGrid3.TextMatrix(1, 2) = .TextMatrix(.Row, CisVsGrid.FixedGet(6))
        VsGrid3.TextMatrix(1, 3) = .TextMatrix(.Row, CisVsGrid.FixedGet(7))
        VsGrid3.TextMatrix(2, 1) = .TextMatrix(.Row, CisVsGrid.FixedGet(8))
        VsGrid3.TextMatrix(2, 2) = .TextMatrix(.Row, CisVsGrid.FixedGet(9))
        VsGrid3.TextMatrix(2, 3) = .TextMatrix(.Row, CisVsGrid.FixedGet(10))
    End With

    For gInt = 1 To 3
        With VsGrid3
            wCuren = CisFun.Val2(.TextMatrix(1, gInt)) _
                                 - CisFun.Val2(.TextMatrix(2, gInt))
            If gInt = 1 And wCuren <> 0 Then
               .TextMatrix(3, gInt) = Format(wCuren, "#,##0.00")
            Else
               .TextMatrix(3, gInt) = Format(wCuren, "#,###")
            End If
        End With
    Next gInt
    
    With VSGRID1
        wYM = Replace(.TextMatrix(.Row, CisVsGrid.FixedGet(11)), "/", "")
        wSyogo = Mid(.TextMatrix(.Row, CisVsGrid.FixedGet(1)), 1, 1)
        wTorcd = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(18)))
        wYMD = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(14)))
        wHinbn = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(13)))
        wNouno = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(15)))
        wKanri = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(16)))
        wZaimu = CisFun.Val2(.TextMatrix(.Row, CisVsGrid.FixedGet(17)))
        wAKKbn = CisFun.Val2(.TextMatrix(.Row, CisVsGrid.FixedGet(22)))
    End With
    
    ' 購入明細テーブルより表示
    gSL_Select = "Select KM.*,"
    gSL_Select = gSL_Select & " ISNULL(TR.略称,'') 仕入先名 "
    gSL_Select = gSL_Select & " From 購入明細テーブル KM "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "      ON  TR.取引先CD = KM.仕入先"
    gSL_Select = gSL_Select & "      AND TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & " Where KM.契約先 = '" & wTorcd & "'"
    gSL_Select = gSL_Select & "   AND KM.照合日付 = '" & wYMD & "'"
    gSL_Select = gSL_Select & "   AND KM.照合品番 = '" & wHinbn & "'"
    gSL_Select = gSL_Select & "   AND KM.照合納番 = '" & wNouno & "'"
    gSL_Select = gSL_Select & "   AND KM.照合管理 = '" & wKanri & "'"
    gSL_Select = gSL_Select & "   AND KM.照合赤黒 = " & wAKKbn
    If KMTRead(gSL_Select) Then
       Call GridSetU
    End If
    Call KMTClose
    ' TEXT買掛金テーブルより表示
    gSL_Select = "Select * "
    gSL_Select = gSL_Select & " From TEXT買掛金テーブル "
    gSL_Select = gSL_Select & " Where 財務管理番号 = " & wZaimu
    If TKKRead(gSL_Select, 1) Then
         With CisFun
              .DateE_BefVal = TKK.検収日
              .DateEdit UserDeli, "/"
              gStr = .DateE_AfVal
         End With
        B2lb_Ymd = gStr
        B2lb_Ukeir = TKK.受入
        B2lb_Tei = TKK.訂正理由
    End If
    
End Sub
'*------------------------------------*
'*   照合結果 登録確認(区分名称)      *
'*------------------------------------*
Private Sub Get_Code()
        
' 名称マスタ検索　確認
    Call ItemsClearSYMCreate
    With SYMCreate
        .区分名称 = "照合結果受給"
        .区分桁数 = 1
        .値 = "B"
        .値名称 = "自社ﾃﾞｰﾀなし"
        .制御区分 = "X"
        .作成者 = gTanto
    End With
    If Not SYMInitCreate(True) Then
        With SYMCreate
            .値 = "C"
            .値名称 = "契約先未検収"
        End With
        Call SYMInitCreate
        With SYMCreate
            .値 = "D"
            .値名称 = "数量の相違"
        End With
        Call SYMInitCreate
        With SYMCreate
            .値 = "E"
            .値名称 = "単価の相違"
        End With
        Call SYMInitCreate
        With SYMCreate
            .値 = "F"
            .値名称 = "単価･数量の相違"
        End With
        Call SYMInitCreate
        With SYMCreate
            .値 = "G"
            .値名称 = "日付の相違"
        End With
        Call SYMInitCreate
    End If
End Sub

