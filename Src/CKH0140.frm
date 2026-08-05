VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CKH0140 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん増分一括発行"
   ClientHeight    =   10980
   ClientLeft      =   60
   ClientTop       =   1485
   ClientWidth     =   15315
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10980
   ScaleWidth      =   15315
   WindowState     =   2  '最大化
   Begin Cis3D_v60.CIS3D CIS3D4 
      Height          =   390
      Left            =   13050
      Top             =   1515
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   688
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
      cBoderColor1    =   8421504
      cBoderColor2    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   345
         Left            =   30
         Top             =   30
         Width           =   1920
         _ExtentX        =   3387
         _ExtentY        =   609
         Caption         =   "表示件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   50
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   285
            Left            =   990
            Top             =   30
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   503
            BackColor       =   16777152
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -50
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   8985
      Left            =   300
      ScaleHeight     =   8925
      ScaleWidth      =   14655
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1935
      Width           =   14715
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   8820
         Left            =   60
         Top             =   60
         Width           =   14550
         _ExtentX        =   25665
         _ExtentY        =   15558
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
         Begin vsFlexLib.vsFlexArray vsGrid1 
            Height          =   8685
            Left            =   90
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   60
            Width           =   14355
            _Version        =   196608
            _ExtentX        =   25321
            _ExtentY        =   15319
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
            FormatString    =   "№  |仕入先              |受|品番                          |背番号|収容数|最終№|前月|当月|追加|発行|"
            Rows            =   26
            Cols            =   20
            BackColor       =   12648384
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   0
            RowHeightMin    =   330
            GridLines       =   2
            SelectionMode   =   1
            Begin CisText_V60.CisText iB1_SetSu 
               Height          =   375
               Left            =   11220
               TabIndex        =   9
               Top             =   330
               Visible         =   0   'False
               Width           =   600
               _ExtentX        =   1058
               _ExtentY        =   661
               cFormat         =   "###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###"
               MaxLength       =   3
               cDataType       =   1
               cDataReplace    =   1
               cFbComma        =   0
               cFdAutoFormat   =   1
               cGFormat        =   "###"
               cILength        =   2
            End
         End
      End
   End
   Begin VB.PictureBox H1_Area1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1260
      Left            =   315
      ScaleHeight     =   1200
      ScaleWidth      =   10665
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   600
      Width           =   10725
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1110
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "発行連番"
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
         cPositionX      =   30
         cPositionY      =   30
         Begin VB.OptionButton OpH1_ZGKbn 
            BackColor       =   &H00C0FFC0&
            Caption         =   "再発行"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   0
            Left            =   90
            TabIndex        =   2
            Top             =   375
            Value           =   -1  'True
            Width           =   1125
         End
         Begin VB.OptionButton OpH1_ZGKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "追　加"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Index           =   1
            Left            =   90
            TabIndex        =   1
            Top             =   735
            Width           =   1155
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   3
         Left            =   5040
         Top             =   60
         Width           =   660
         _ExtentX        =   1164
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "受入"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   150
            TabIndex        =   4
            Top             =   480
            Width           =   375
            _ExtentX        =   661
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
            Text            =   "XX"
            MaxLength       =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   2
         Left            =   1380
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "仕  入  先"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   90
            TabIndex        =   3
            Top             =   480
            Width           =   1005
            _ExtentX        =   1773
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1080
            Top             =   480
            Width           =   2460
            _ExtentX        =   4339
            _ExtentY        =   635
            BackColor       =   16777152
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   4
         Left            =   5700
         Top             =   60
         Width           =   3330
         _ExtentX        =   5874
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "品　　番"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   90
            TabIndex        =   5
            Top             =   480
            Width           =   3195
            _ExtentX        =   5636
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
            Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXX"
            MaxLength       =   25
            IMEMode         =   2
            cDataReplace    =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   1125
         Index           =   5
         Left            =   9030
         Top             =   60
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   1984
         ForeColor       =   16711680
         Caption         =   "年 月"
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
         cPositionX      =   30
         cPositionY      =   60
         Begin Cis3D_v60.CIS3D CIS3D3 
            Height          =   360
            Left            =   180
            Top             =   450
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "9999年99月"
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
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   720
      Left            =   11520
      ScaleHeight     =   660
      ScaleWidth      =   3435
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   555
      Width           =   3495
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   900
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   1740
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
         Left            =   2580
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
      Begin CisBtn_60.CisBtn PB_Look 
         Height          =   585
         Left            =   45
         Top             =   30
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
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
   End
   Begin Cis3D_v60.CIS3D UC_3D11 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 かんばん増分一括発行 】"
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
         Left            =   14385
         Top             =   60
         Width           =   915
         _ExtentX        =   1614
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
         Height          =   300
         Left            =   11835
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   529
         BackColor       =   14737632
         Caption         =   "XXXXXXXXX1XXXXXXXXX2"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   0
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
   Begin VB.Timer Timer2 
      Left            =   705
      Top             =   -120
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   -150
   End
   Begin VB.TextBox Dummy 
      Height          =   375
      IMEMode         =   3  'ｵﾌ固定
      Left            =   -15
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   -105
      Width           =   345
   End
End
Attribute VB_Name = "CKH0140"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :  かんばん増減一括メンテナンス
'**       フォームID    :  CKH0120
'**       処理概要      :
'**
'**       作  成  日    :  2004/02/03
'**       変  更  日    :  2004/06/22  By CIS - 背番号＆品番の半角カナ入力対応
'**       変  更  日    :  2004/09/24   検索PFキー
'**       変  更  日    :  2004/10/04  By CIS - バグ対応（縦スクロールが表示しない！）
'**       変  更  日    :  2004/10/19  By CIS - PageUp/Downによる縦スクロール
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim CisVsGrid1              As New CisVsGrid3
'   増減区分の背景色
    Dim mColor_On               As Long
    Dim mColor_Off              As Long
'   グリッドの背景色
    Dim mColorOrg               As Long
    Dim mColorChg               As Long
'   行
    Dim NowRow                  As Long
    Dim NewRow                  As Long
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
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
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

'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 11
        .NumberSet = NoSpace
        .EditCheckCol = 11
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 18
        .InitGet = False
        .Init
    End With
    Call GridInit1

'年月初期表示
    Call GetNJProc_YM(3)
    iH1_Ym = gn_HNYM
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
'+----------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》              +
'+----------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

'   背景色の設定（増減区分）
    mColor_On = OpH1_ZGKbn(0).BackColor
    mColor_Off = OpH1_ZGKbn(1).BackColor
'   背景色の設定（グリッド）
    mColorOrg = vsGrid1.BackColor
    mColorChg = &HC0E0FF

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

    Select Case Key_Code

'< 2004.09.24 MNT START >
'
'        Case vbKeyF7:       If Dummy.Tag <> "" Then Call PB_Look_Click      '【検索】
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      '【検索】
'
'< 2004.09.24 MNT  END  >

        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          '【入力】
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
    If ProcHB = "B1" Then
       ProcHB = "H1"
       Call DispChange(ProcHB)
       Call HeadBodyClear("B")
       GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
End Sub
'+----------------------------+
'+      取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9            ' [発注]手配先/契約先/社内検索

    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+----------------------------+
'+      取引先受入検索        +
'+----------------------------+
Private Sub Look_Ukeir()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9            ' [発注]手配先/契約先/社内検索
    RV_Torcd = iH1_Torcd
    RV_Kojyo = ""

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+----------------------------+
'+      品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RVI_Torcd = RTrim(iH1_Torcd)
    RVI_Ukeir = RTrim(iH1_Ukeir)
    RVI_Hinmoku = 1

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        iH1_Ukeir = RV_Ukeir
        H1lb_Tornm = ""
        If TorNmGet(iH1_Torcd, 9) Then
           H1lb_Tornm = TRM_RName
        End If
    End If
End Sub
'+-----------------------------+
'+  OpH1_ZGKbn
'+-----------------------------+
Private Sub OpH1_ZGKbn_Click(Index As Integer)
    If Index = 0 Then
        OpH1_ZGKbn(1).BackColor = mColor_Off
        OpH1_ZGKbn(0).BackColor = mColor_On
        iB1_BinSu.Enabled = True
    Else
        OpH1_ZGKbn(0).BackColor = mColor_Off
        OpH1_ZGKbn(1).BackColor = mColor_On
        iB1_BinSu.Enabled = False
    End If
End Sub
'+-----------------------------+
'+  iH1_Torcd
'+-----------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    H1lb_Tornm = ""
    If TorNmGet(iH1_Torcd, 9) Then
       H1lb_Tornm = TRM_RName
    End If
End Sub
'+-----------------------------+
'+  iH1_Ukeir
'+-----------------------------+
Private Sub iH1_Ukeir_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
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

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
       Case "H1": H1Mode = True: H1Color = gPGotSel
       Case "B1": B1Mode = True: B1Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color

    PB_Look.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode

' ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
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
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = ""
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
        Call CisVsGrid1.Clear
        iB1_SetSu.Visible = False
        iB1_BinSu.Visible = False
    End If
End Sub
'*********************************
'*          入力処理(Enter)      *
'*********************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = "B1"
        Call DispChange(ProcHB)
        Call CisVsGrid1.EditVisible("C")        ' ｸﾞﾘｯﾄﾞ内容表示後、ﾃｷｽﾄ入力
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
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

'   取引先
    If Trim(iH1_Torcd) <> "" Then
        If Not TorNmGet(iH1_Torcd, 9) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    取引先　未登録！（取引先マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Torcd.SetFocus
           Exit Function
        End If
'       受入
        If Trim(iH1_Ukeir) <> "" Then
            gSL_Select = ""
            gSL_Select = gSL_Select & "select *"
            gSL_Select = gSL_Select & "  from 取引先受入マスタ"
            gSL_Select = gSL_Select & " where 取引先CD = '" & RTrim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "   and 取引先区分 = 1"
            gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH1_Ukeir) & "'"
            If Not TUMRead(gSL_Select, 1) Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    受入　未登録！（取引先受入マスタ）            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
               iH1_Ukeir.SetFocus
               Exit Function
            End If
        End If
    End If
'   品番
    If Trim(iH1_Hinbn) <> "" Then
'*---------------------------------------------<< 2004/06/22 Insert Start >>
    '   >> 半角ﾁｪｯｸ（品番）
        If CisFun.Len2(RTrim(iH1_Hinbn)) <> Len(RTrim(iH1_Hinbn)) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    品番は半角にて入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Hinbn.SetFocus
            Exit Function
        End If
'*---------------------------------------------<< 2004/06/22 Insert End >>
        gSL_Select = ""
        gSL_Select = gSL_Select & "select *"
        gSL_Select = gSL_Select & "  from 品番マスタ"
        gSL_Select = gSL_Select & " where 品番 = '" & RTrim(iH1_Hinbn) & "'"
        If Not HNMRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    品番　未登録！（品番マスタ）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
           iH1_Hinbn.SetFocus
           Exit Function
        End If
    End If
'   年月
    If Trim(iH1_Ym) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    年月を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iH1_Ym.SetFocus
        Exit Function
    Else
        If Not iH1_Ym.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    年月を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iH1_Ym.SetFocus
            Exit Function
        End If
    End If

    Call SqlSelectSet
    If Not DBInput("INV") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "    該当データが存在しません｡            "
            .MB_MSG(3) = "    再度抽出条件を入力してください       "
            .MB_Title = "抽出処理"
            .MB_Button = Error
            .MBOX
        End With
        Exit Function
    End If

    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False

    Call CisVsGrid1.EditEnable
    Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ

    Dim ID          As Long
    With vsGrid1
        For ID = 1 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For
        '   設定枚数
            If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(9))) = 0 Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    " & ID & "行目 ： 設定枚数を入力して下さい            "
                    .MB_Title = ""
                    .MB_Button = Error
                    .MBOX
                End With
                .Row = ID
                Call CisVsGrid1.EditEnable
                Call CisVsGrid1.EditVisible
                Exit Function
            End If
        Next ID
    End With

    Body1Chk = True
End Function
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示       *
'*                                    *
'**************************************
Private Function DBInput(InputKey As String) As Boolean
    DBInput = False
    On Error GoTo DBInput_Err:
    
    If GridSet1 Then
        DBInput = True
    End If

DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'****************************
'*      SQL文ｾｯﾄ処理
'****************************
Private Sub SqlSelectSet(Optional InputKey As String = "")
    Dim wYmT        As String
    Dim wYmY        As String
'
    With CisFun
        wYmT = iH1_Ym
        .Date_RsFormat = fYM
        .Date_Add iH1_Ym, M, 1
        wYmY = .Date_Result
    End With
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  kzt.*"
    gSL_Select = gSL_Select & ", isnull(trm.略称, '') 取引先名"
    gSL_Select = gSL_Select & ", isnull(hnm.表示品番, '') 表示品番"
    gSL_Select = gSL_Select & ", isnull(htm.背番号, '') 背番号"
    gSL_Select = gSL_Select & ", case when isnull(htm.変更日, '') = '' then 収容数1"
    gSL_Select = gSL_Select & "       when htm.変更日 > '" & Format(Date, "yyyymmdd") & "'"
    gSL_Select = gSL_Select & "            then htm.収容数1"
    gSL_Select = gSL_Select & "            else htm.収容数2"
    gSL_Select = gSL_Select & "   end 収容数"
    gSL_Select = gSL_Select & ", isnull(hnx.当月内示数, 0) 当月内示数"
    gSL_Select = gSL_Select & ", isnull(hnx.翌月内示数, 0) 翌月内示数"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  かんばん増減テーブル kzt"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = kzt.取引先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
    gSL_Select = gSL_Select & "    on hnm.品番 = kzt.品番"
    gSL_Select = gSL_Select & "  left outer join 品番取引先マスタ htm"
    gSL_Select = gSL_Select & "    on htm.品番 = kzt.品番"
    gSL_Select = gSL_Select & "   and htm.取引先 = kzt.取引先"
    gSL_Select = gSL_Select & "   and htm.受入 = kzt.受入"
    gSL_Select = gSL_Select & "  left outer join"
    gSL_Select = gSL_Select & " ("
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  hnj.品番"
    gSL_Select = gSL_Select & ", hnj.手配先"
    gSL_Select = gSL_Select & ", hnj.受入"
    gSL_Select = gSL_Select & ", sum(case when substring(hnj.納期, 1, 6) = '" & wYmT & "' then hnj.数量 else 0 end) 当月内示数"
    gSL_Select = gSL_Select & ", sum(case when substring(hnj.納期, 1, 6) = '" & wYmY & "' then hnj.数量 else 0 end) 翌月内示数"
    gSL_Select = gSL_Select & " from 発注内示テーブル hnj"
    gSL_Select = gSL_Select & " where hnj.内示年月 = '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & " group by"
    gSL_Select = gSL_Select & "  hnj.品番"
    gSL_Select = gSL_Select & ", hnj.手配先"
    gSL_Select = gSL_Select & ", hnj.受入"
    gSL_Select = gSL_Select & ") hnx"
    gSL_Select = gSL_Select & "    on hnx.品番 = kzt.品番"
    gSL_Select = gSL_Select & "   and hnx.手配先 = kzt.取引先"
    gSL_Select = gSL_Select & "   and hnx.受入 = kzt.受入"
    gSL_Select = gSL_Select & " where kzt.設定内容区分 = 0"
'   増減区分
    If OpH1_ZGKbn(0) Then gSL_Select = gSL_Select & "   and kzt.増減区分 = 0"
    If OpH1_ZGKbn(1) Then gSL_Select = gSL_Select & "   and kzt.増減区分 = 1"
'   取引先
    If RTrim(iH1_Torcd) <> "" Then
        gSL_Select = gSL_Select & "   and kzt.取引先 = '" & RTrim(iH1_Torcd) & "'"
    End If
'   受入
    If RTrim(iH1_Ukeir) <> "" Then
        gSL_Select = gSL_Select & "   and kzt.受入 = '" & RTrim(iH1_Ukeir) & "'"
    End If
'   品番
    If RTrim(iH1_Hinbn) <> "" Then
        gSL_Select = gSL_Select & "   and kzt.品番 = '" & RTrim(iH1_Hinbn) & "'"
    End If
'   年月
    gSL_Select = gSL_Select & "   and substring(kzt.開始日, 1, 6) <= '" & iH1_Ym & "'"
    gSL_Select = gSL_Select & "   and case when kzt.終了日 = ''"
    gSL_Select = gSL_Select & "            then '999912'"
    gSL_Select = gSL_Select & "            else substring(kzt.終了日, 1, 6)"
    gSL_Select = gSL_Select & "        end >= '" & iH1_Ym & "'"

    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  kzt.品番"
    gSL_Select = gSL_Select & ", kzt.取引先"
    gSL_Select = gSL_Select & ", kzt.受入"
    gSL_Select = gSL_Select & ", kzt.開始日"
'
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
    DBPut = False

    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "    更新処理を行います。            "
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Title = "更新処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 更新処理 )
    Dim ID          As Long
    With vsGrid1
        For ID = 1 To .Rows - 1
            If Trim(.TextMatrix(ID, CisVsGrid1.FixedGet(0))) = "" Then Exit For
            If CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(9))) <> CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(13))) Or _
               CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(10))) <> CisFun.Val2(.TextMatrix(ID, CisVsGrid1.FixedGet(14))) Then
                Call DBUpdate_KZT(ID)
            End If
        Next ID
    End With

    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      かんばん増減テーブル更新
'+------------------------------------+
Public Function DBUpdate_KZT(SetRow As Long, Optional DBNo As Integer) As Boolean
    DBUpdate_KZT = False
    Call ItemsClearKZT
    With KZT
        .取引先 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(11)))
        .受入 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(2)))
        .品番 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(16)))
        .開始日 = RTrim(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(12)))
        .設定枚数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(9)))
        .便当枚数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(10)))
        .更新者 = gTanto
        .累積枚数 = CisFun.Val2(vsGrid1.TextMatrix(SetRow, CisVsGrid1.FixedGet(15)))

        gSL_Select = ""
        gSL_Select = gSL_Select & "update かんばん増減テーブル"
        gSL_Select = gSL_Select & "   set 設定枚数 = " & .設定枚数 & ""
        gSL_Select = gSL_Select & "     , 便当枚数 = " & .便当枚数 & ""
        If .設定枚数 > .累積枚数 Then
            gSL_Select = gSL_Select & "     , 完了日 = ''"
        End If
        gSL_Select = gSL_Select & "     , 更新日 = getdate()"
        gSL_Select = gSL_Select & "     , 更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & "     , 更新端末 = substring(host_name(),1,20)"
        gSL_Select = gSL_Select & " where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   and 開始日 = '" & RTrim(.開始日) & "'"
    End With
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBUpdate_KZT = True
End Function



'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - 流動日付
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 4, 0, "№")
        Call .FixedSet(中中, 左中, あり, 20, 1, "仕 入 先")
        Call .FixedSet(中中, 左中, あり, 2, 2, "受")
        Call .FixedSet(中中, 左中, あり, 30, 3, "品　 番")
        Call .FixedSet(中中, 左中, あり, 6, 4, "背番号")
        Call .FixedSet(中中, 右中, あり, 6, 5, "収容数")
        Call .FixedSet(中中, 右中, あり, 10, 6, "当月内示数")
        Call .FixedSet(中中, 右中, あり, 10, 7, "翌月内示数")
        Call .FixedSet(中中, 中中, あり, 12, 8, "設定期間")
        Call .FixedSet(中中, 右中, あり, 4, 9, "設定")
        Call .FixedSet(中中, 右中, あり, 4, 10, "便当")
        Call .FixedSet(中中, 左中, あり, 0, 11, "X仕入先")
        Call .FixedSet(中中, 左中, あり, 0, 12, "X開始日")
        Call .FixedSet(中中, 右中, あり, 0, 13, "X設定枚数")
        Call .FixedSet(中中, 右中, あり, 0, 14, "X便当枚数")
        Call .FixedSet(中中, 右中, あり, 0, 15, "X累積枚数")
        Call .FixedSet(中中, 左中, あり, 0, 16, "X品番")
        Call .FixedSet(中中, 左中, なし, 0, 17, "NotDisp")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 18, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10")
'   編集内容ｾｯﾄ
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO,非表示ｶﾗﾑ,編集ﾌｫｰﾏｯﾄ,日付編集)
        Call .EditSet(iB1_SetSu, 9)
        Call .EditSet(iB1_BinSu, 10, 17)
        .InitDisp
    End With
End Sub
'***********************************************'
'*****      グリッド内容セット - 流動日付
'***********************************************'
Private Function GridSet1() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet1 = False

    If KZTRead(gSL_Select, , 1) Then

'       抽出処理 初期ｾｯﾄ
        B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid1
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid1.DispMax Then
                .Rows = CisVsGrid1.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
        End With
        ID = 0
        Do Until Not KZT_RDSTS
            ID = ID + 1

            Call KztInfoToGrid1(ID)

            Call KZTReadNext(1)
        Loop
        With vsGrid1
             .Row = 1: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With
        NowRow = vsGrid1.Row

        GridSet1 = True
    End If
    Call KZTClose
End Function
'+----------------------------------------------+
'+  かんばん増減テーブルよりグリッドに転送
'+----------------------------------------------+
Private Sub KztInfoToGrid1(SetRow As Long)
    With vsGrid1
        'No
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(0)) = SetRow
        '仕入先
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = RTrim(KZT.取引先名)
        '受入
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = RTrim(KZT.受入)
        '品番
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = RTrim(KZT.表示品番)
        '背番号
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = RTrim(KZT.背番号)
        '収容数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(5)) = Format(KZT.収容数, "#,###")
        '当月内示数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(6)) = Format(KZT.当月内示数, "#,###")
        '翌月内示数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = Format(KZT.翌月内示数, "#,###")
        '設定機関
        If Trim(KZT.終了日) = "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(8)) = CisFun.Mid2(Format(KZT.開始日, "0000/00/00"), 6, 5) & Space(7)
        Else
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(8)) = CisFun.Mid2(Format(KZT.開始日, "0000/00/00"), 6, 5) _
                                                 & "～" & CisFun.Mid2(Format(KZT.終了日, "0000/00/00"), 6, 5)
        End If
        '設定枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(9)) = Format(KZT.設定枚数, "###")
        '便当枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(10)) = Format(KZT.便当枚数, "###")
        '退避
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(11)) = KZT.取引先
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(12)) = KZT.開始日
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(13)) = KZT.設定枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(14)) = KZT.便当枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(15)) = KZT.累積枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(16)) = KZT.品番
        '表示制御
        Select Case OpH1_ZGKbn(0)
            Case True:  .TextMatrix(SetRow, CisVsGrid1.FixedGet(17)) = ""
            Case False: .TextMatrix(SetRow, CisVsGrid1.FixedGet(17)) = "NotDisp"
        End Select
    End With
End Sub
'+-------------------------------+
'+      VSGrid1_Click
'+-------------------------------+
Private Sub VSGrid1_Click()
    If ProcHB = "B1" Then
        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
        Call CisVsGrid1.EditEnable
        Call CisVsGrid1.EditVisible
        If Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(0))) <> "" Then
            NowRow = vsGrid1.Row
        End If
    End If
End Sub
'+---------------------------------------------------------+
'+          グリッド編集エリア
'+---------------------------------------------------------+
'+-------------------------------+
'+      iB1_SetSu
'+-------------------------------+
Private Sub iB1_SetSu_KeyDown(KeyCode As Integer, Shift As Integer)
    
'---------------------------------------------------------------< 2004.10.19 MNT START >
    Call BD_InKeyCheck(ActiveControl, KeyCode, Shift)
'    If KeyCode = vbKeyUp Or _
'       KeyCode = vbKeyDown Then
'        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
'        vsGrid1.SetFocus
'        Call CisVsGrid1.KeyEdit(iB1_SetSu, KeyCode, Shift)
'        NowRow = vsGrid1.Row
'    End If
'---------------------------------------------------------------< 2004.10.19 MNT  END  >

End Sub
'+-------------------------------+
'+      iB1_BinSu
'+-------------------------------+
Private Sub iB1_BinSu_KeyDown(KeyCode As Integer, Shift As Integer)
    
'---------------------------------------------------------------< 2004.10.19 MNT START >
    Call BD_InKeyCheck(ActiveControl, KeyCode, Shift)
'    If KeyCode = vbKeyUp Or _
'       KeyCode = vbKeyDown Then
'        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
'        vsGrid1.SetFocus
'        Call CisVsGrid1.KeyEdit(iB1_BinSu, KeyCode, Shift)
'        NowRow = vsGrid1.Row
'    End If
'---------------------------------------------------------------< 2004.10.19 MNT  END  >

End Sub
'+--------------------------------------+
'+      ボディ入力項目キーチェック
'+--------------------------------------+
Private Function BD_InKeyCheck(TObj As Object, KeyCode As Integer, Shift As Integer)

'---------------------------------------------------------------< 2004.10.19 ADD START >
'< PageUp・PageDown >
    If KeyCode = vbKeyPageDown Or _
       KeyCode = vbKeyPageUp Then
        
        If vsGrid1.Rows <= CisVsGrid1.DispMax Then Exit Function
        
        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
        
        If KeyCode = vbKeyPageDown Then
            
            '< PageDown >
            If vsGrid1.TopRow + 25 <= vsGrid1.Rows - 1 Then
                vsGrid1.TopRow = vsGrid1.TopRow + 25
                
                If vsGrid1.TopRow + 24 <= vsGrid1.Rows - 1 Then
                    vsGrid1.Row = vsGrid1.TopRow + 24
                Else
                    vsGrid1.Row = vsGrid1.Rows - 1
                End If
            Else
                vsGrid1.TopRow = vsGrid1.Rows - 1 - 24
            End If
        Else
            
            '< PageUp >
            If vsGrid1.TopRow - 25 > 0 Then
                vsGrid1.TopRow = vsGrid1.TopRow - 25
            Else
                vsGrid1.TopRow = 1
            End If
            vsGrid1.Row = vsGrid1.TopRow
        End If
        
        vsGrid1.SetFocus
        Call CisVsGrid1.EditEnable
        Call CisVsGrid1.EditVisible
        NowRow = vsGrid1.Row
        TObj.SetFocus
        Exit Function
    End If
    
'< ↑・↓ >
    If KeyCode = vbKeyUp Or _
       KeyCode = vbKeyDown Then
        Call ChgColorGrid1(NowRow)      ' 内容変更ﾁｪｯｸ
        vsGrid1.SetFocus
        Call CisVsGrid1.KeyEdit(TObj, KeyCode, Shift)
        NowRow = vsGrid1.Row
    End If

'---------------------------------------------------------------< 2004.10.19 ADD  END  >

End Function
'+-------------------------------+
'+      背景色の再設定
'+-------------------------------+
Private Sub ChgColorGrid1(SetRow As Long)
    With vsGrid1
        .Redraw = False
        NewRow = .Row
        .Row = SetRow
        If CisFun.Val2(iB1_SetSu) <> CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(13))) Or _
           CisFun.Val2(iB1_BinSu) <> CisFun.Val2(.TextMatrix(SetRow, CisVsGrid1.FixedGet(14))) Then
            For gInt = 1 To 10
                .Col = gInt
                .CellBackColor = mColorChg
            Next gInt
        Else
            For gInt = 1 To 10
                .Col = gInt
                .CellBackColor = mColorOrg
            Next gInt
        End If
        .Row = NewRow
        .Redraw = True
    End With
End Sub
