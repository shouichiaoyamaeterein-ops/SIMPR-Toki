VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "Cis3D_6.0.ocx"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CisBtn_6.0.ocx"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CisText_6.0.ocx"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CisYM_6.0.ocx"
Begin VB.Form CXZ0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "在庫マスタ(部品)"
   ClientHeight    =   11190
   ClientLeft      =   90
   ClientTop       =   1500
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
   Begin VB.PictureBox PNL_Disp 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1560
      ScaleHeight     =   435
      ScaleWidth      =   9420
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2640
      Width           =   9480
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   450
         Left            =   0
         Top             =   0
         Width           =   2070
         _ExtentX        =   3651
         _ExtentY        =   794
         BackColor       =   12632319
         ForeColor       =   0
         Caption         =   "機械："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         cPositionY      =   15
         Begin Cis3D_v60.CIS3D H1lb_Kikai 
            Height          =   360
            Left            =   750
            Top             =   60
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   635
            BackColor       =   12632319
            Caption         =   "XXXXXXXXX1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632319
            cBoderColor2    =   12632319
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   450
         Left            =   2070
         Top             =   0
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   794
         BackColor       =   12632319
         ForeColor       =   0
         Caption         =   "機械№："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         cPositionY      =   15
         Begin Cis3D_v60.CIS3D H1lb_KikNo 
            Height          =   360
            Left            =   1020
            Top             =   60
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   635
            BackColor       =   12632319
            Caption         =   "XXXXXXXXX1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632319
            cBoderColor2    =   12632319
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   450
         Left            =   4440
         Top             =   0
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   794
         BackColor       =   12632319
         ForeColor       =   0
         Caption         =   "置場："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         cPositionY      =   15
         Begin Cis3D_v60.CIS3D H1lb_Okiba 
            Height          =   360
            Left            =   750
            Top             =   60
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   635
            BackColor       =   12632319
            Caption         =   "XXXXXXXXX1"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632319
            cBoderColor2    =   12632319
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   450
         Left            =   6480
         Top             =   0
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   794
         BackColor       =   12632319
         ForeColor       =   0
         Caption         =   "理論在庫数："
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         cPositionY      =   15
         Begin Cis3D_v60.CIS3D H1lb_Zaiko 
            Height          =   360
            Left            =   1500
            Top             =   60
            Width           =   1380
            _ExtentX        =   2434
            _ExtentY        =   635
            BackColor       =   12632319
            Caption         =   "###,###,##0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cBoderColor1    =   12632319
            cBoderColor2    =   12632319
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   2
            cPositionX      =   30
         End
      End
   End
   Begin VB.PictureBox B2_Area1 
      Height          =   6690
      Left            =   1680
      ScaleHeight     =   6630
      ScaleWidth      =   11175
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   4230
      Width           =   11235
      Begin vsFlexLib.vsFlexArray VsGrid1 
         Height          =   6630
         Left            =   0
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   0
         Width           =   11175
         _Version        =   196608
         _ExtentX        =   19711
         _ExtentY        =   11695
         _StockProps     =   228
         Appearance      =   1
         ConvInfo        =   1418783674
         FormatString    =   "№ |9999/99  |999,999,999|実棚 |XXXXXXXXX1XXXXXXXXX1XXXXXXXXX1XXXXXXXXX1|担当|9999/99/99 "
         Rows            =   21
         Cols            =   12
         BackColor       =   12648447
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   330
         SelectionMode   =   1
         BackColorAlternate=   12648447
      End
   End
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
      Height          =   765
      Left            =   7650
      ScaleHeight     =   705
      ScaleWidth      =   5295
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1020
      Width           =   5355
      Begin CisBtn_60.CisBtn PB_LOOK 
         Height          =   585
         Left            =   1800
         Top             =   60
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
         Left            =   2670
         Top             =   60
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
         Left            =   3540
         Top             =   60
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
         Left            =   4410
         Top             =   60
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
      Begin CisBtn_60.CisBtn PB_BACK 
         Height          =   585
         Left            =   60
         Top             =   60
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
         BtnType         =   6
         BtnCaption      =   "後退"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8421376
         BtnShadowColor  =   16777215
         PFCaption       =   "( Up )"
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
         PFBackColor     =   8421376
      End
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   930
         Top             =   60
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
         BtnType         =   5
         BtnCaption      =   "前進"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8421376
         BtnShadowColor  =   16777215
         PFCaption       =   "( Down )"
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
         PFBackColor     =   8421376
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
      TabIndex        =   8
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
      Height          =   7800
      Left            =   1560
      ScaleHeight     =   7740
      ScaleWidth      =   11385
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3210
      Width           =   11445
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   8055
         Left            =   90
         Top             =   60
         Width           =   11235
         _ExtentX        =   19817
         _ExtentY        =   14208
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
         Begin Cis3D_v60.CIS3D BLB_Date_B 
            Height          =   885
            Left            =   4500
            Top             =   30
            Width           =   1185
            _ExtentX        =   2090
            _ExtentY        =   1561
            ForeColor       =   16711680
            Caption         =   "実棚日"
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
            Begin CisText_V60.CisText iB1_Day 
               Height          =   375
               Left            =   180
               TabIndex        =   4
               Top             =   390
               Width           =   435
               _ExtentX        =   767
               _ExtentY        =   661
               cFormat         =   "##"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "##"
               MaxLength       =   2
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "##"
               cILength        =   2
            End
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   285
               Left            =   660
               Top             =   480
               Width           =   345
               _ExtentX        =   609
               _ExtentY        =   503
               ForeColor       =   0
               Caption         =   "日"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   12632256
               cBoderColor2    =   12632256
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   7
               cBoderStyle     =   2
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D9 
            Height          =   885
            Index           =   10
            Left            =   30
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
            Begin CisText_V60.CisText iB1_Tanto 
               Height          =   360
               Left            =   150
               TabIndex        =   2
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
            Begin Cis3D_v60.CIS3D B1lb_Tanto 
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
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   885
            Index           =   1
            Left            =   2610
            Top             =   30
            Width           =   1875
            _ExtentX        =   3307
            _ExtentY        =   1561
            ForeColor       =   16711680
            Caption         =   "在 庫 数"
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
            Begin CisText_V60.CisText iB1_Zaiko 
               Height          =   375
               Left            =   180
               TabIndex        =   3
               Top             =   390
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   661
               cFormat         =   "###,###,###"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   12
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Text            =   "###,###,###"
               MaxLength       =   11
               cDataType       =   1
               cDataReplace    =   1
               cFdAutoFormat   =   1
               cGFormat        =   "#########"
               cILength        =   9
            End
         End
         Begin Cis3D_v60.CIS3D UC_3D15 
            Height          =   885
            Index           =   2
            Left            =   5700
            Top             =   30
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   1561
            ForeColor       =   16711680
            Caption         =   "変　更　事　由"
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
            Begin CisText_V60.CisText iB1_Change 
               Height          =   375
               Left            =   240
               TabIndex        =   5
               Top             =   420
               Width           =   5055
               _ExtentX        =   8916
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
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   40
               IMEMode         =   1
               cDataReplace    =   1
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
      Height          =   1950
      Left            =   1560
      ScaleHeight     =   1890
      ScaleWidth      =   5115
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   570
      Width           =   5175
      Begin Cis3D_v60.CIS3D UC_3D15 
         Height          =   885
         Index           =   5
         Left            =   90
         Top             =   60
         Width           =   1560
         _ExtentX        =   2752
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "年月度"
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
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   360
            Left            =   90
            TabIndex        =   0
            Top             =   420
            Width           =   1380
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
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   885
         Left            =   1650
         Top             =   60
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1561
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
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   375
            Left            =   90
            TabIndex        =   1
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
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   885
         Left            =   90
         Top             =   945
         Width           =   3675
         _ExtentX        =   6482
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "部　　　署"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   375
            Left            =   90
            TabIndex        =   12
            Top             =   420
            Width           =   960
            _ExtentX        =   1693
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
            IMEMode         =   2
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1050
            Top             =   420
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   885
         Left            =   3765
         Top             =   945
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1561
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
         cPositionX      =   15
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   375
            Left            =   150
            TabIndex        =   13
            Top             =   420
            Width           =   330
            _ExtentX        =   582
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
            Text            =   "XX"
            MaxLength       =   2
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   885
         Left            =   4395
         Top             =   945
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   1561
         ForeColor       =   16711680
         Caption         =   "工順"
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
         cPositionX      =   15
         cPositionY      =   60
         Begin CisText_V60.CisText iH1_Kojun 
            Height          =   375
            Left            =   150
            TabIndex        =   14
            Top             =   420
            Width           =   330
            _ExtentX        =   582
            _ExtentY        =   661
            cFormat         =   "##"
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
            Text            =   "##"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   270
      Top             =   420
   End
   Begin VB.Timer Timer2 
      Left            =   690
      Top             =   420
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
      Caption         =   "【 在庫マスタ メンテ(部品)  】"
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
Attribute VB_Name = "CXZ0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   在庫マスタメンテナンス（部品）
'**       フォームID    :   CXZ0010
'**       処理概要      :
'**
'**       作  成  日    :   2003/12/01  By CIS SIMPR-A
'**       変  更  日    :   2004/04/14  バグ修正
'**       変  更  日    :   2004/06/28  前進＆後退キーの処理内容が逆になっていたのを修正
'**       変  更  日    :   2004/10/06  実棚日の入力条件をなくす
'**       　  　  　                    ＆翌月在庫データが存在時、メッセージ表示（Yes or No）
'**       　  　  　                    ＆バグ修正（実棚日が当日以降でもエラーにならない！）
'**       変  更  日    :   2010/01/14  条件追加：部署/受入/工数
'**                                     表示追加：機械/機械№/置場/理論在庫数
'**       変  更  日    :   2010/07/05  変更理由入力条件をなくす
'**       変  更  日    :   2010/11/11  他画面パラメーター追加
'**       変  更  日    :   2019/12/17  棚卸ﾌﾟﾚｼｰﾄ条件中止
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
    
    Dim wSplit()        As String               '2010/11/11
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
        .SelectCol = 9
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
    iH1_YM = Mid(wsDate, 1, 6)
    
'======================================================= 2010/11/11 Start
    If RTrim$(Command) <> "" Then
       wSplit = Split(Command, ",")
       If UBound(wSplit) = 3 Then
          iH1_YM = wSplit(0)
          iH1_Hinbn = wSplit(1)
          iH1_Torcd = wSplit(2)
          Call iH1_Torcd_LostFocus
                   
          gSL_Select = "SELECT 品番,取引先,受入,工順 FROM 品番工順マスタ "
          gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"
          gSL_Select = gSL_Select & " AND   取引先 = '" & RTrim(iH1_Torcd) & "'"
          If HKTRead(gSL_Select, 1) Then
             If CisDB.RecordCount = 1 Then
                iH1_Ukeir = HKT.受入
                iH1_Kojun = Format(HKT.工順, "##")
             End If
          End If
       End If
    End If
'======================================================= 2010/11/11 End
    
    
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
           Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
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
   If ProcHB = "B1" Or ProcHB = "B2" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      iH1_YM.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
'---------------------------------------------------------------------------------------------------' 2010.01.15 del end
''''    gSL_Select = " SELECT 年月度,品番"
''''    gSL_Select = gSL_Select & ",取引先, 受入, 工順"                         ' 2010.01.14 add
''''    gSL_Select = gSL_Select & " FROM 在庫マスタ "
''''    gSL_Select = gSL_Select & " WHERE 年月度 = '" & RTrim(iH1_YM) & "'"
'''''   gSL_Select = gSL_Select & "   AND 品番 < '" & RTrim(iH1_Hinbn) & "'"    ' 2010.01.15 del
'''''---------------------------------------------------------------------------' 2010.01.15 add start
''''    gSL_Select = gSL_Select & "   AND (SUBSTRING(品番   + SPACE(25),1,25) +"
''''    gSL_Select = gSL_Select & "        SUBSTRING(取引先 + SPACE( 7),1, 7) +"
''''    gSL_Select = gSL_Select & "        SUBSTRING(受入   + SPACE( 2),1, 2) +"
''''    gSL_Select = gSL_Select & "        CONVERT(VARCHAR,工順)  <  "
''''    gSL_Select = gSL_Select & "'" & Mid(Trim(iH1_Hinbn) & Space(25), 1, 25) _
''''                                  & Mid(Trim(iH1_Torcd) & Space(7), 1, 7) _
''''                                  & Mid(Trim(iH1_Ukeir) & Space(2), 1, 2) _
''''                                  & Trim(iH1_Kojun) & "')"
'''''---------------------------------------------------------------------------' 2010.01.15 add end
''''    gSL_Select = gSL_Select & " GROUP BY 年月度,品番 "
''''    gSL_Select = gSL_Select & ",取引先, 受入, 工順"                         ' 2010.01.14 add
''''    gSL_Select = gSL_Select & " ORDER BY 年月度,品番 Desc "
''''    gSL_Select = gSL_Select & ",取引先 DESC, 受入 DESC, 工順 DESC"          ' 2010.01.14 add
'---------------------------------------------------------------------------------------------------' 2010.01.15 del end
'---------------------------------------------------------------------------------------------------' 2010.01.15 add start
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT ZK.年月度, HK.品番, HK.取引先, HK.受入, HK.工順"
    gSL_Select = gSL_Select & " FROM  品番工順マスタ HK"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 在庫マスタ ZK"
    gSL_Select = gSL_Select & "        ON  ZK.年月度 = '" & Trim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "        AND HK.品番   = ZK.品番"
    gSL_Select = gSL_Select & "        AND HK.取引先 = ZK.取引先"
    gSL_Select = gSL_Select & "        AND HK.受入   = ZK.受入"
    gSL_Select = gSL_Select & "        AND HK.工順   = ZK.工順"
    gSL_Select = gSL_Select & " WHERE (SUBSTRING(HK.品番   + SPACE(25),1,25) +"
    gSL_Select = gSL_Select & "        SUBSTRING(HK.取引先 + SPACE( 7),1, 7) +"
    gSL_Select = gSL_Select & "        SUBSTRING(HK.受入   + SPACE( 2),1, 2) +"
    gSL_Select = gSL_Select & "        CONVERT(VARCHAR,HK.工順)"
    gSL_Select = gSL_Select & "     < '" & Mid(Trim(iH1_Hinbn) & Space(25), 1, 25) _
                                         & Mid(Trim(iH1_Torcd) & Space(7), 1, 7) _
                                         & Mid(Trim(iH1_Ukeir) & Space(2), 1, 2) _
                                         & Trim(iH1_Kojun) & "')"
    'gSL_Select = gSL_Select & "   AND HK.棚卸プレシート発行 = 0"                               '2019/12/17 ｺﾒﾝﾄ
    gSL_Select = gSL_Select & " GROUP BY HK.品番 ,HK.取引先, HK.受入, HK.工順, ZK.年月度"
    gSL_Select = gSL_Select & " ORDER BY HK.品番 DESC ,HK.取引先 DESC, HK.受入 DESC, HK.工順 DESC"
'---------------------------------------------------------------------------------------------------' 2010.01.15 add end
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
'---------------------------------------------------------------------------------------------------' 2010.01.15 del start
''''    gSL_Select = " SELECT 年月度,品番"
''''    gSL_Select = gSL_Select & ",取引先, 受入, 工順"                         ' 2010.01.14 add
''''    gSL_Select = gSL_Select & " FROM 在庫マスタ "
''''    gSL_Select = gSL_Select & " WHERE 年月度 = '" & RTrim(iH1_YM) & "'"
'''''   gSL_Select = gSL_Select & "   AND 品番 > '" & RTrim(iH1_Hinbn) & "'"    ' 2010.01.15 del
'''''---------------------------------------------------------------------------' 2010.01.15 add start
''''    gSL_Select = gSL_Select & "   AND (SUBSTRING(品番   + SPACE(25),1,25) +"
''''    gSL_Select = gSL_Select & "        SUBSTRING(取引先 + SPACE( 7),1, 7) +"
''''    gSL_Select = gSL_Select & "        SUBSTRING(受入   + SPACE( 2),1, 2) +"
''''    gSL_Select = gSL_Select & "        CONVERT(VARCHAR,工順)  >  "
''''    gSL_Select = gSL_Select & "'" & Mid(Trim(iH1_Hinbn) & Space(25), 1, 25) _
''''                                  & Mid(Trim(iH1_Torcd) & Space(7), 1, 7) _
''''                                  & Mid(Trim(iH1_Ukeir) & Space(2), 1, 2) _
''''                                  & Trim(iH1_Kojun) & "')"
'''''---------------------------------------------------------------------------' 2010.01.15 add end
''''    gSL_Select = gSL_Select & " GROUP BY 年月度,品番 "
''''    gSL_Select = gSL_Select & ",取引先, 受入, 工順"                         ' 2010.01.14 add
''''    gSL_Select = gSL_Select & " ORDER BY 年月度,品番"
''''    gSL_Select = gSL_Select & ",取引先, 受入, 工順"                         ' 2010.01.14 add
'---------------------------------------------------------------------------------------------------' 2010.01.15 del end
'---------------------------------------------------------------------------------------------------' 2010.01.15 add start
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT ZK.年月度, HK.品番, HK.取引先, HK.受入, HK.工順"
    gSL_Select = gSL_Select & " FROM  品番工順マスタ HK"
    gSL_Select = gSL_Select & "       LEFT OUTER JOIN 在庫マスタ ZK"
    gSL_Select = gSL_Select & "        ON  ZK.年月度 = '" & Trim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "        AND HK.品番   = ZK.品番"
    gSL_Select = gSL_Select & "        AND HK.取引先 = ZK.取引先"
    gSL_Select = gSL_Select & "        AND HK.受入   = ZK.受入"
    gSL_Select = gSL_Select & "        AND HK.工順   = ZK.工順"
    gSL_Select = gSL_Select & " WHERE (SUBSTRING(HK.品番   + SPACE(25),1,25) +"
    gSL_Select = gSL_Select & "        SUBSTRING(HK.取引先 + SPACE( 7),1, 7) +"
    gSL_Select = gSL_Select & "        SUBSTRING(HK.受入   + SPACE( 2),1, 2) +"
    gSL_Select = gSL_Select & "        CONVERT(VARCHAR,HK.工順)"
    gSL_Select = gSL_Select & "     > '" & Mid(Trim(iH1_Hinbn) & Space(25), 1, 25) _
                                         & Mid(Trim(iH1_Torcd) & Space(7), 1, 7) _
                                         & Mid(Trim(iH1_Ukeir) & Space(2), 1, 2) _
                                         & Trim(iH1_Kojun) & "')"
    'gSL_Select = gSL_Select & "   AND HK.棚卸プレシート発行 = 0"       2019/12/17ｺﾒﾝﾄ
    gSL_Select = gSL_Select & " GROUP BY HK.品番 ,HK.取引先, HK.受入, HK.工順, ZK.年月度"
    gSL_Select = gSL_Select & " ORDER BY HK.品番 ,HK.取引先, HK.受入, HK.工順"
'---------------------------------------------------------------------------------------------------' 2010.01.15 add end
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Hinbn.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iB1_Tanto" Then Call Look_Tanto: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub     ' 2010.01.14 add
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir: Exit Sub     ' 2010.01.14 add
   'If PB_Look.Tag = "iH1_Kojun" Then Call Look_Kojun: Exit Sub     ' 2010.01.14 add    ' DEL
End Sub
'+----------------------------+
'+       品番 検索              +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
        
    RVI_Torcd = Trim(iH1_Torcd)
    RVI_Hinmoku = 1
    RVI_Tehai = 9   ' 2
    RVI_SK = 9      '
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK: Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'-----------------------------------------------------------------------------------------------------------' 2010.01.14 add start
'+---------------------+
'+    取引先検索        +
'+---------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 1                '   発注
    RV_TorKb2 = 0               '   社内非限定

    CKK0020.Show vbModal
    Unload CKK0020
    Set CKK0020 = Nothing

    If RV_Rtn Then
       iH1_Torcd = RV_Torcd
       Call iH1_Torcd_LostFocus
    End If
End Sub
'+---------------------+
'+    受入検索　        +
'+---------------------+
Private Sub Look_Ukeir()
    RV_Left = 0
    RV_Top = 0
  ''RV_TorKb = 3            ' 受注ならこっち
    RV_TorKb = 9            ' 発注ならこっち
    RV_Torcd = iH1_Torcd
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+---------------------+
'+    工順検索　        +
'+---------------------+
Private Sub Look_Kojun()
    RV_Left = 0
    RV_Top = 0

    RV_Hinbn = Trim(iH1_Hinbn)
    RV_Hinnm = Trim(H1lb_Tornm)
    RVI_Torcd = Trim(iH1_Torcd)
    RVI_Ukeir = Trim(iH1_Ukeir)
    RVI_Kubun = "工順"
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT MM1.値名称 AS 機械名, MM2.値名称 AS 機械NO名, HK.*"
    gSL_Select = gSL_Select & " FROM 品番工順マスタ HK"
    gSL_Select = gSL_Select & "      LEFT OUTER JOIN 名称マスタ MM1"
    gSL_Select = gSL_Select & "       ON  MM1.区分名称 = '機械'"
    gSL_Select = gSL_Select & "       AND MM1.値       = HK.機械"
    gSL_Select = gSL_Select & "      LEFT OUTER JOIN 名称マスタ MM2"
    gSL_Select = gSL_Select & "       ON  MM2.区分名称 = '機械NO'"
    gSL_Select = gSL_Select & "       AND MM2.値       = HK.機械NO"
    gSL_Select = gSL_Select & " WHERE HK.品番   = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND HK.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND HK.受入   = '" & RTrim(iH1_Ukeir) & "'"
    'gSL_Select = gSL_Select & "   AND HK.棚卸プレシート発行 = 0"               '2019/12/17ｺﾒﾝﾄ
    If Not HKTRead(gSL_Select) Then Exit Sub
    
    If CisDB.RecordCount = 1 Then       ' 工順が１つしかなければその値を表示する
        iH1_Kojun = Trim(Str(HKT.工順))
        Exit Sub
    End If
    
    CKJ0016.Show vbModal
    Unload CKJ0016
    Set CKJ0016 = Nothing

    If RV_Rtn Then
       iH1_Kojun = CStr(Rv_CallRtn)     ' 工順
    End If
End Sub
'-----------------------------------------------------------------------------------------------------------' 2010.01.14 add end

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
        iB1_Tanto = RV_Code
        Call iB1_Tanto_LostFocus
    End If
End Sub
'+-----------------------------+
'+  iH1_YM
'+-----------------------------+
Private Sub iH1_YM_GotFocus()
    PB_BACK.Visible = True  ' False ' 2010.01.15 upd
    PB_NEXT.Visible = True  ' False ' 2010.01.15 upd
    PB_Look.Visible = False
' *ADD* 2004/04/14
    DoEvents
End Sub
Private Sub iH1_YM_LostFocus()
' *CHG* 2004/04/14
'    PB_BACK.Visible = True
'    PB_NEXT.Visible = True
'    If ProcHB = "H1" Then          ' 2010.01.15 del
'        PB_BACK.Visible = True     ' 2010.01.15 del
'        PB_NEXT.Visible = True     ' 2010.01.15 del
'    End If                         ' 2010.01.15 del
End Sub
'+-----------------------------+
'+  iH1_Hinbn
'+-----------------------------+
Private Sub iH1_Hinbn_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
    PB_BACK.Visible = True          ' 2010.01.15 add
    PB_NEXT.Visible = True          ' 2010.01.15 add
End Sub
Private Sub iH1_Hinbn_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub

'-----------------------------------------------------------------------------------------------------------' 2010.01.14 add start
'+-----------------------------+
'+  iH1_Torcd
'+-----------------------------+
Private Sub iH1_Torcd_GotFocus()
    PB_BACK.Visible = True
    PB_NEXT.Visible = True
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Torcd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
    If RTrim(iH1_Torcd) <> "" Then
        If TorNmGet(iH1_Torcd, 0) Then
            H1lb_Tornm = TRM_RName
        Else
            H1lb_Tornm = ""
        End If
    Else
        H1lb_Tornm = ""
    End If
''''If ProcHB = "H1" Then
''''    PB_BACK.Visible = True
''''    PB_NEXT.Visible = True
''''End If
End Sub
'+-----------------------------+
'+  iH1_Ukeir
'+-----------------------------+
Private Sub iH1_Ukeir_GotFocus()
    PB_BACK.Visible = True
    PB_NEXT.Visible = True
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Ukeir_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
''''If ProcHB = "H1" Then
''''    PB_BACK.Visible = True
''''    PB_NEXT.Visible = True
''''End If
End Sub
'+-----------------------------+
'+  iH1_Kojun
'+-----------------------------+
Private Sub iH1_Kojun_GotFocus()
    PB_BACK.Visible = True
    PB_NEXT.Visible = True
''''PB_Look.Tag = ActiveControl.Name
''''PB_Look.Visible = True
End Sub
Private Sub iH1_Kojun_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
''''If ProcHB = "H1" Then
''''    PB_BACK.Visible = True
''''    PB_NEXT.Visible = True
''''End If
End Sub
'-----------------------------------------------------------------------------------------------------------' 2010.01.14 add end

'+-----------------------------+
'+  iB1_Tanto
'+-----------------------------+
Private Sub iB1_Tanto_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_Tanto_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If

    B1lb_Tanto = ""
    If MeisyoGet("担当管理", RTrim(iB1_Tanto)) Then
        B1lb_Tanto = RTrim(SYM_Meisyo)
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
    Dim B2Mode      As Boolean
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
    
    PB_Look.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    
'*--------------------------------------------------*<< 2004/10/06 >> Delete Start
' 実棚日
'   If Pro = "B1" Then
'       If RTrim(iH1_YM) <> Mid(wsDate, 1, 6) Then
'           BLB_Date_B.Enabled = False
'       Else
'           BLB_Date_B.Enabled = True
'       End If
'   End If
'*--------------------------------------------------*<< 2004/10/06 >> Delete End
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
        ProcHB = "B1"
        Call DispChange(ProcHB)
        DoEvents
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed

        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call iH1_Hinbn.SetFocus     ' 2010.01.15 add
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
    
' 年月度
    If RTrim(iH1_YM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    年月度を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_YM.SetFocus
       Exit Function
    End If
' 品番
    If RTrim(iH1_Hinbn) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    品番を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Hinbn.SetFocus
       Exit Function
    Else
        gSL_Select = "SELECT * FROM 品番マスタ"
''''    gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(iH1_Hinbn) & "'"                    ' 2010.01.15 del
        gSL_Select = gSL_Select & " WHERE 品番 = '" & RTrim(Replace(iH1_Hinbn, "-", "")) & "'"  ' 2010.01.15 add
        If Not HNMRead(gSL_Select, 1) Then
            With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    品番マスタ  未登録            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
            End With
            iH1_Hinbn.SetFocus
            Exit Function
        End If
    End If
'-----------------------------------------------------------------------------------------------------------' 2010.01.14 add start
' 部署
    If RTrim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    部署を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    Else
        If Trim(H1lb_Tornm) = "" Then
            With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    部署  未登録               "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
            End With
            iH1_Torcd.SetFocus
            Exit Function
        End If
    End If
' 受入
    If RTrim(iH1_Ukeir) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    受入を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Ukeir.SetFocus
       Exit Function
    Else
        '---受入登録チェックが必要ならばここに記載する
    End If
' 工順
    If RTrim(iH1_Kojun) = "" Then
        Call Look_Kojun
    End If
    If RTrim(iH1_Kojun) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    工順を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Kojun.SetFocus
       Exit Function
    Else
''''        gSL_Select = ""
''''        gSL_Select = gSL_Select & "SELECT MM1.値名称 AS 機械名, MM2.値名称 AS 機械NO名, HK.*"
''''        gSL_Select = gSL_Select & " FROM 品番工順マスタ HK"
''''        gSL_Select = gSL_Select & "      LEFT OUTER JOIN 名称マスタ MM1"
''''        gSL_Select = gSL_Select & "       ON  MM1.区分名称 = '機械'"
''''        gSL_Select = gSL_Select & "       AND MM1.値       = HK.機械"
''''        gSL_Select = gSL_Select & "      LEFT OUTER JOIN 名称マスタ MM2"
''''        gSL_Select = gSL_Select & "       ON  MM2.区分名称 = '機械NO'"
''''        gSL_Select = gSL_Select & "       AND MM2.値       = HK.機械NO"
''''        gSL_Select = gSL_Select & " WHERE HK.品番   = '" & RTrim(iH1_Hinbn) & "'"
''''        gSL_Select = gSL_Select & "   AND HK.取引先 = '" & RTrim(iH1_Torcd) & "'"
''''        gSL_Select = gSL_Select & "   AND HK.受入   = '" & RTrim(iH1_Ukeir) & "'"
''''        gSL_Select = gSL_Select & "   AND HK.工順   = '" & RTrim(iH1_Kojun) & "'"
''''        If Not HKTRead(gSL_Select, 1) Then
        If Not DispMeisai() Then
            With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    品番工順マスタ  未登録            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
            End With
            Call iH1_Hinbn.SetFocus     ' iH1_Kojun.SetFocus
            Exit Function
        End If
        
        
'        If HKT.棚卸プレシート発行 <> 0 Then                                 '          '2019/12/17ｺﾒﾝﾄ
'            With CisFun
'                 .MB_Lines = 4
'                 .MB_MSG(2) = "    在庫区分=有ではありません            "
'                 .MB_Title = ""
'                 .MB_Button = Error
'                 .MBOX
'            End With
'            Call iH1_Hinbn.SetFocus     ' iH1_Kojun.SetFocus
'            Exit Function
'        End If
    
    
    End If
''''    With HKT
''''        H1lb_Kikai = .機械名
''''        H1lb_KikNo = .機械NO名
''''        H1lb_Okiba = .置場
''''    End With
    Call GetZaikoSu
'-----------------------------------------------------------------------------------------------------------' 2010.01.14 add end
    
'*--------------------------------------------------*<< 2004/10/06 >> Insert Start
'   >> 翌月在庫データの存在ﾁｪｯｸ
    gSL_Select = "select * from 在庫マスタ"
    gSL_Select = gSL_Select & " where 年月度 > '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 受入   = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND 工順   = '" & RTrim(iH1_Kojun) & "'"
    If ZKMRead(gSL_Select, 1) Then
        With CisFun
           .MB_Lines = 4
           .MB_MSG(1) = "    " & CisFun.Mid2(ZKM.年月度, 1, 4) & "年" & CisFun.Mid2(ZKM.年月度, 5, 2) & "月度のデータが在庫マスタに存在します。            "
           .MB_MSG(3) = "          処理を続行しますか？            "
           .MB_Title = "処理確認"
           .MB_Button = Yes_No
           If Not .MBOX Then Call iH1_Hinbn.SetFocus: Exit Function
        End With
    End If
'*--------------------------------------------------*<< 2004/10/06 >> Insert End

' BODY部 セット
    Call BodySet("")
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
' 担当者
    If RTrim(iB1_Tanto) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    担当者を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_Tanto.SetFocus
       Exit Function
    Else
        If Not MeisyoGet("担当管理", RTrim(iB1_Tanto)) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    担 当 者  未 登 録                "
                .MB_MSG(4) = "     【 名称マスタ 】                "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_Tanto.SetFocus
            Exit Function
        End If
    End If

' 実棚日
    wDate = RTrim(iH1_YM) & "01"
    With CisFun
        .Date_RsFormat = fDD
        .Date_End wDate
        wEndDay = .Val2(.Date_Result)
    End With
    If CisFun.Val2(iB1_Day) > wEndDay Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    正しい実棚日を入力して下さい        "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Day.SetFocus
        Exit Function
    End If
'*--------------------------------------------------*<< 2004/10/06 >> Update Start
'   If CisFun.Val2(iB1_Day) > Format(Now(), "dd") Then
'       With CisFun
'           .MB_Lines = 4
'           .MB_MSG(2) = "    正しい実棚日を入力して下さい        "
'           .MB_Title = ""
'           .MB_Button = Error
'           .MBOX
'       End With
'       iB1_Day.SetFocus
'       Exit Function
'   End If
    If (iH1_YM & Format(CisFun.Val2(iB1_Day), "@@")) > Format(Now(), "yyyymmdd") Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    実棚日は今日迄の日で入力して下さい        "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        iB1_Day.SetFocus
        Exit Function
    End If
'*--------------------------------------------------*<< 2004/10/06 >> Update End

' 変更事由
'======================== 2010/07/05
'    If RTrim(iB1_Change) = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    変更事由を入力して下さい        "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'        End With
'        iB1_Change.SetFocus
'        Exit Function
'    End If
            
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
    Call HeadBodyClear("B")
    
    If Not ZKMRead(gSL_Select, 1) Then
        If InputKey = "Back" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     先頭レコードです            "
                .MB_Title = "【後退】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
        If InputKey = "Next" Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "     最終レコードです            "
                .MB_Title = "【前進】"
                .MB_Button = OK
                .MBOX
            End With
            GoTo DBInput_Ed
        End If
    Else
''''    iH1_YM = RTrim(ZKM.年月度)      ' 2010.01.15 del
        iH1_Hinbn = RTrim(ZKM.品番)
        iH1_Torcd = RTrim(ZKM.取引先)   ' 2010.01.14 add
        iH1_Ukeir = RTrim(ZKM.受入)     ' 2010.01.14 add
        iH1_Kojun = Trim(Str(ZKM.工順)) ' 2010.01.14 add
        Call BodySet("")
        Call iH1_Torcd_LostFocus        ' 2010.01.14 add
        DBInput = True
    End If
DBInput_Ed:
    Exit Function
DBInput_Err:
     CisFun.ErrorBox
End Function
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Sub BodySet(SetType As String)
    Static wsBDFlg  As Boolean

'
    With CisFun
        .Date_RsFormat = fYM
        .Date_Add RTrim(iH1_YM), M, -12, ResultDate
        gStr = .Date_Result
    End With

' 在庫履歴マスタ 情報
    gSL_Select = "Select * from 在庫履歴マスタ "
    gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(iH1_Hinbn) & "'"
'---------------------------------------------------------------------------' 2010.01.15 add start
    gSL_Select = gSL_Select & "   AND 取引先 = '" & Trim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 受入   = '" & Trim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND 工順   =  " & Trim(iH1_Kojun) & " "
'---------------------------------------------------------------------------' 2010.01.15 add end
    gSL_Select = gSL_Select & "   AND 年月度 BETWEEN '" & RTrim(gStr) & "' AND '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & " ORDER BY 年月度 DESC,入力日付 DESC "
    If ZKRRead(gSL_Select) Then
       Call GridSet
    End If
    Call DispMeisai     ' 2010.01.15 add
    Call ZKRClose
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
    Dim AddFlg  As Boolean
    Dim wZAIKO  As Long
    Dim wJDate      As String

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

'   在庫マスタ 削除
    gSL_Select = "DELETE 在庫マスタ "
    gSL_Select = gSL_Select & " WHERE 年月度 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   And 品番   = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(iH1_Torcd) & "'"      ' 2010.01.15 add
    gSL_Select = gSL_Select & "   AND 受入   = '" & RTrim(iH1_Ukeir) & "'"      ' 2010.01.15 add
    gSL_Select = gSL_Select & "   AND 工順   =  " & CisFun.Val2(iH1_Kojun)      ' 2010.01.15 add
    With CisDB
        .ConnectNo = 1
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    
' 在庫マスタ 作成
    Call ItemsClearZKM
    With ZKM
        .年月度 = RTrim(iH1_YM)
        .品番 = RTrim(iH1_Hinbn)
        .取引先 = RTrim(iH1_Torcd)      ' 2010.01.15 add
        .受入 = RTrim(iH1_Ukeir)        ' 2010.01.15 add
        .工順 = CisFun.Val2(iH1_Kojun)  ' 2010.01.15 add
        .在庫数 = CisFun.Val2(iB1_Zaiko)
        wJDate = ""
        If CisFun.Val2(iB1_Day) <> 0 Then
            wJDate = RTrim(iH1_YM) & Format(iB1_Day, "00")
        End If
        .実棚日 = RTrim(wJDate)
        .変更事由 = RTrim(iB1_Change)
        .作成区分 = 1
    End With
    ZKM.作成者 = gTanto
    Call ZKMInsert

' 履歴区分更新
    gSL_Select = " UPDATE 在庫履歴マスタ SET "
    gSL_Select = gSL_Select & " 履歴区分 = 1 "
    gSL_Select = gSL_Select & " WHERE 年月度 = '" & RTrim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   AND 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & RTrim(iH1_Torcd) & "'"      ' 2010.01.15 add
    gSL_Select = gSL_Select & "   AND 受入   = '" & RTrim(iH1_Ukeir) & "'"      ' 2010.01.15 add
    gSL_Select = gSL_Select & "   AND 工順   =  " & CisFun.Val2(iH1_Kojun)      ' 2010.01.15 add
    gSL_Select = gSL_Select & "   AND 実棚日 = '" & RTrim(wJDate) & "'"
    With CisDB
        .ConnectNo = 1
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With


' 在庫履歴マスタ 作成
    Call ItemsClearZKR
    With ZKR
        .年月度 = RTrim(iH1_YM)
        .品番 = RTrim(iH1_Hinbn)
        .取引先 = RTrim(iH1_Torcd)      ' 2010.01.15 add
        .受入 = RTrim(iH1_Ukeir)        ' 2010.01.15 add
        .工順 = CisFun.Val2(iH1_Kojun)  ' 2010.01.15 add
        .入力日付 = Now()
        .入力担当者 = RTrim(iB1_Tanto)
        .在庫数 = CisFun.Val2(iB1_Zaiko)
        If CisFun.Val2(iB1_Day) <> 0 Then
            .実棚日 = RTrim(iH1_YM) & Format(iB1_Day, "00")
        End If
        .変更事由 = RTrim(iB1_Change)
        .作成区分 = 1
    End With
    ZKR.作成者 = gTanto
    Call ZKRInsert
    
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
    Call DispMeisai                     ' 2010.01.15 add
    
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
    ReDim SCol(10)
    
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "NO")
        Call .FixedSet(中中, 中中, あり, 4, 1, "区分")
        Call .FixedSet(中中, 中中, あり, 9, 2, "年月度")
        Call .FixedSet(中中, 右中, あり, 11, 3, "在庫数")
        Call .FixedSet(中中, 中中, あり, 5, 4, "実棚")
        Call .FixedSet(中中, 左中, あり, 40, 5, "変更事由")
        Call .FixedSet(中中, 左中, あり, 5, 6, "担当")
        Call .FixedSet(中中, 中中, あり, 12, 7, "変更日")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 8, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,2,3,4,5,6,7")
        .InitDisp
    
        SCol(1) = .FixedGet(1)  ' 区分
        SCol(2) = .FixedGet(2)  ' 年月度
        SCol(3) = .FixedGet(3)  ' 在庫数
        SCol(4) = .FixedGet(4)  ' 実棚日
        SCol(5) = .FixedGet(5)  ' 変更事由
        SCol(6) = .FixedGet(6)  ' 担当
        SCol(7) = .FixedGet(7)  ' 変更日
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
    Do Until Not ZKR_RDSTS
        With VsGrid1
            ID = ID + 1
            ' №
            .TextMatrix(ID, 0) = Format(ID, "###")
            ' 年月度
            .TextMatrix(ID, SCol(2)) = Mid(ZKR.年月度, 1, 4) & "/" & Mid(ZKR.年月度, 5, 2)
            ' 在庫数
            .TextMatrix(ID, SCol(3)) = Format(ZKR.在庫数, "###,###,###")
            ' 実棚日
            If RTrim(ZKR.実棚日) <> "" Then: .TextMatrix(ID, SCol(4)) = Mid(ZKR.実棚日, 7, 2)
            ' 変更事由
            .TextMatrix(ID, SCol(5)) = RTrim(ZKR.変更事由)
            ' 担当者
            .TextMatrix(ID, SCol(6)) = RTrim(ZKR.入力担当者)
            ' 変更日
            .TextMatrix(ID, SCol(7)) = Format(ZKR.入力日付, "yyyy/mm/dd")
        
        End With
        Call ZKRReadNext
    Loop
    
    With VsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub

'-----------------------------------------------------------------------------------' 2010.01.14 add start  ' 2010.01.29 upd start
Private Function GetZaikoSu() As Boolean
    GetZaikoSu = False
    Dim wZaikoSu    As Long: wZaikoSu = 0
    Dim wMasterSu   As Long: wMasterSu = 0

    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT * FROM 在庫マスタ"
    gSL_Select = gSL_Select & " WHERE 年月度 = '" & Trim(iH1_YM) & "'"
    gSL_Select = gSL_Select & "   AND 品番   = '" & Trim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND 取引先 = '" & Trim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND 受入   = '" & Trim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND 工順   = '" & Trim(iH1_Kojun) & "'"
    If Not ZKMRead(gSL_Select) Then
        wMasterSu = 0
        '---在庫マスタ：無
        gSL_Select = ""
''        gSL_Select = gSL_Select & "SELECT ISNULL(SUM(SN1.加工数),0) + ISNULL(SUM(JK1.実績数),0) - ISNULL(SUM(JK2.実績数),0) AS 加工数"
''        gSL_Select = gSL_Select & "      ,ISNULL(SUM(SN1.加工数),0) AS 加工数入, ISNULL(SUM(JK1.実績数),0) AS 実績数入, ISNULL(SUM(JK2.実績数),0) AS 実績数出"
        gSL_Select = gSL_Select & "SELECT ISNULL(SUM(SN.加工数),0) + ISNULL(SUM(JK1.実績数),0) - ISNULL(SUM(JK2.実績数),0) AS 加工数"
        gSL_Select = gSL_Select & "      ,ISNULL(SUM(SN.加工数),0) AS 加工数入, ISNULL(SUM(JK1.実績数),0) AS 実績数入, ISNULL(SUM(JK2.実績数),0) AS 実績数出"
        gSL_Select = gSL_Select & "      ,SN.部門, SN.品番, SN.受入, SN.工順"
        gSL_Select = gSL_Select & " FROM 作業日報テーブル SN"
''        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 作業日報テーブル SN1"
''        gSL_Select = gSL_Select & "        ON  SUBSTRING(SN1.作業日,1,6) >= '" & Trim(iH1_YM) & "'"
''        gSL_Select = gSL_Select & "        AND SN1.部門     = SN.部門"
''        gSL_Select = gSL_Select & "        AND SN1.品番     = SN.品番"
''        gSL_Select = gSL_Select & "        AND SN1.受入     = SN.受入"
''        gSL_Select = gSL_Select & "        AND SN1.工順     = SN.工順"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 実績管理テーブル JK1"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) >= '" & Trim(iH1_YM) & "'"
        gSL_Select = gSL_Select & "        AND JK1.実績先   = SN.部門"
        gSL_Select = gSL_Select & "        AND JK1.品番     = SN.品番"
        gSL_Select = gSL_Select & "        AND JK1.受入     = SN.受入"
        gSL_Select = gSL_Select & "        AND JK1.工順     = SN.工順"
        gSL_Select = gSL_Select & "        AND JK1.実績区分 = 1"
        gSL_Select = gSL_Select & "       LEFT OUTER JOIN 実績管理テーブル JK2"
        gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) >= '" & Trim(iH1_YM) & "'"
        gSL_Select = gSL_Select & "        AND JK2.実績先   = SN.部門"
        gSL_Select = gSL_Select & "        AND JK2.品番     = SN.品番"
        gSL_Select = gSL_Select & "        AND JK2.受入     = SN.受入"
        gSL_Select = gSL_Select & "        AND JK2.工順     = SN.工順"
        gSL_Select = gSL_Select & "        AND JK2.実績区分 <> 1"
        gSL_Select = gSL_Select & "        AND JK2.入力種類 <> 'A'"
''      gSL_Select = gSL_Select & "  WHERE SUBSTRING(SN.作業日,1,6) >= '" & Trim(iH1_YM) & "'"
''      gSL_Select = gSL_Select & "    AND SN.部門 = '" & Trim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "  WHERE SN.部門 = '" & Trim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "    AND SN.品番 = '" & Trim(iH1_Hinbn) & "'"
        gSL_Select = gSL_Select & "    AND SN.受入 = '" & Trim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & "    AND SN.工順 =  " & Trim(iH1_Kojun) & " "
        gSL_Select = gSL_Select & "    AND SUBSTRING(SN.作業日,1,6) >= '" & Trim(iH1_YM) & "'"
        gSL_Select = gSL_Select & "  GROUP BY SN.部門, SN.品番, SN.受入, SN.工順"
    Else
        '---在庫マスタ：有
        wMasterSu = ZKM.在庫数
        If Trim(ZKM.実棚日) = "" Then
            '---実棚日：空白
            gSL_Select = ""
''            gSL_Select = gSL_Select & "SELECT ISNULL(SUM(SN1.加工数),0) + ISNULL(SUM(JK1.実績数),0) - ISNULL(SUM(JK2.実績数),0) AS 加工数"
''            gSL_Select = gSL_Select & "      ,ISNULL(SUM(SN1.加工数),0) AS 加工数入, ISNULL(SUM(JK1.実績数),0) AS 実績数入, ISNULL(SUM(JK2.実績数),0) AS 実績数出"
            gSL_Select = gSL_Select & "SELECT ISNULL(SUM(SN.加工数),0) + ISNULL(SUM(JK1.実績数),0) - ISNULL(SUM(JK2.実績数),0) AS 加工数"
            gSL_Select = gSL_Select & "      ,ISNULL(SUM(SN.加工数),0) AS 加工数入, ISNULL(SUM(JK1.実績数),0) AS 実績数入, ISNULL(SUM(JK2.実績数),0) AS 実績数出"
            gSL_Select = gSL_Select & "      ,SN.部門, SN.品番, SN.受入, SN.工順"
            gSL_Select = gSL_Select & " FROM 作業日報テーブル SN"
''            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 作業日報テーブル SN1"
''            gSL_Select = gSL_Select & "        ON  SUBSTRING(SN1.作業日,1,6) > '" & Trim(ZKM.年月度) & "'"
''            gSL_Select = gSL_Select & "        AND SN1.部門     = SN.部門"
''            gSL_Select = gSL_Select & "        AND SN1.品番     = SN.品番"
''            gSL_Select = gSL_Select & "        AND SN1.受入     = SN.受入"
''            gSL_Select = gSL_Select & "        AND SN1.工順     = SN.工順"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 実績管理テーブル JK1"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(JK1.実績日,1,6) > '" & Trim(ZKM.年月度) & "'"
            gSL_Select = gSL_Select & "        AND JK1.実績先   = SN.部門"
            gSL_Select = gSL_Select & "        AND JK1.品番     = SN.品番"
            gSL_Select = gSL_Select & "        AND JK1.受入     = SN.受入"
            gSL_Select = gSL_Select & "        AND JK1.工順     = SN.工順"
            gSL_Select = gSL_Select & "        AND JK1.実績区分 = 1"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 実績管理テーブル JK2"
            gSL_Select = gSL_Select & "        ON  SUBSTRING(JK2.実績日,1,6) > '" & Trim(ZKM.年月度) & "'"
            gSL_Select = gSL_Select & "        AND JK2.実績先   = SN.部門"
            gSL_Select = gSL_Select & "        AND JK2.品番     = SN.品番"
            gSL_Select = gSL_Select & "        AND JK2.受入     = SN.受入"
            gSL_Select = gSL_Select & "        AND JK2.工順     = SN.工順"
            gSL_Select = gSL_Select & "        AND JK2.実績区分 <> 1"
            gSL_Select = gSL_Select & "        AND JK2.入力種類 <> 'A'"
''          gSL_Select = gSL_Select & "  WHERE SUBSTRING(SN.作業日,1,6) >= '" & Trim(ZKM.年月度) & "'"  ' ここ以外は[在庫マスタ：無]と同じSQL
''          gSL_Select = gSL_Select & "    AND SN.部門 = '" & Trim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "  WHERE SN.部門 = '" & Trim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "    AND SN.品番 = '" & Trim(iH1_Hinbn) & "'"
            gSL_Select = gSL_Select & "    AND SN.受入 = '" & Trim(iH1_Ukeir) & "'"
            gSL_Select = gSL_Select & "    AND SN.工順 =  " & Trim(iH1_Kojun) & " "
            
            gSL_Select = gSL_Select & "    AND SUBSTRING(SN.作業日,1,6) >= '" & Trim(ZKM.年月度) & "'"
            
            gSL_Select = gSL_Select & "  GROUP BY SN.部門, SN.品番, SN.受入, SN.工順"
        Else
            '---実棚日：入力済
            gSL_Select = ""
''            gSL_Select = gSL_Select & "SELECT ISNULL(SUM(SN1.加工数),0) + ISNULL(SUM(JK1.実績数),0) - ISNULL(SUM(JK2.実績数),0) AS 加工数"
''            gSL_Select = gSL_Select & "      ,ISNULL(SUM(SN1.加工数),0) AS 加工数入, ISNULL(SUM(JK1.実績数),0) AS 実績数入, ISNULL(SUM(JK2.実績数),0) AS 実績数出"
            gSL_Select = gSL_Select & "SELECT ISNULL(SUM(SN.加工数),0) + ISNULL(SUM(JK1.実績数),0) - ISNULL(SUM(JK2.実績数),0) AS 加工数"
            gSL_Select = gSL_Select & "      ,ISNULL(SUM(SN.加工数),0) AS 加工数入, ISNULL(SUM(JK1.実績数),0) AS 実績数入, ISNULL(SUM(JK2.実績数),0) AS 実績数出"
            gSL_Select = gSL_Select & "      ,SN.部門, SN.品番, SN.受入, SN.工順"
            gSL_Select = gSL_Select & " FROM 作業日報テーブル SN"
''            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 作業日報テーブル SN1"
''            gSL_Select = gSL_Select & "        ON  SN1.作業日 > '" & Trim(ZKM.実棚日) & "'"
''            gSL_Select = gSL_Select & "        AND SN1.部門     = SN.部門"
''            gSL_Select = gSL_Select & "        AND SN1.品番     = SN.品番"
''            gSL_Select = gSL_Select & "        AND SN1.受入     = SN.受入"
''            gSL_Select = gSL_Select & "        AND SN1.工順     = SN.工順"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 実績管理テーブル JK1"
            gSL_Select = gSL_Select & "        ON  JK1.実績日   > '" & Trim(ZKM.実棚日) & "'"
            gSL_Select = gSL_Select & "        AND JK1.実績先   = SN.部門"
            gSL_Select = gSL_Select & "        AND JK1.品番     = SN.品番"
            gSL_Select = gSL_Select & "        AND JK1.受入     = SN.受入"
            gSL_Select = gSL_Select & "        AND JK1.工順     = SN.工順"
            gSL_Select = gSL_Select & "        AND JK1.実績区分 = 1"
            gSL_Select = gSL_Select & "       LEFT OUTER JOIN 実績管理テーブル JK2"
            gSL_Select = gSL_Select & "        ON  JK2.実績日   > '" & Trim(ZKM.実棚日) & "'"
            gSL_Select = gSL_Select & "        AND JK2.実績先   = SN.部門"
            gSL_Select = gSL_Select & "        AND JK2.品番     = SN.品番"
            gSL_Select = gSL_Select & "        AND JK2.受入     = SN.受入"
            gSL_Select = gSL_Select & "        AND JK2.工順     = SN.工順"
            gSL_Select = gSL_Select & "        AND JK2.実績区分 <> 1"
            gSL_Select = gSL_Select & "        AND JK2.入力種類 <> 'A'"
''          gSL_Select = gSL_Select & "  WHERE SN.作業日 > '" & Trim(ZKM.実棚日) & "'"
''          gSL_Select = gSL_Select & "    AND SN.部門   = '" & Trim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "  WHERE SN.部門   = '" & Trim(iH1_Torcd) & "'"
            gSL_Select = gSL_Select & "    AND SN.品番   = '" & Trim(iH1_Hinbn) & "'"
            gSL_Select = gSL_Select & "    AND SN.受入   = '" & Trim(iH1_Ukeir) & "'"
            gSL_Select = gSL_Select & "    AND SN.工順   =  " & Trim(iH1_Kojun) & " "
            
            gSL_Select = gSL_Select & "    AND SN.作業日 > '" & Trim(ZKM.実棚日) & "'"
            
            gSL_Select = gSL_Select & "  GROUP BY SN.部門, SN.品番, SN.受入, SN.工順"
        End If
    End If
    If SGTRead(gSL_Select) Then
       wZaikoSu = SGT.加工数
    End If
    wZaikoSu = wMasterSu + wZaikoSu
    H1lb_Zaiko.Caption = Format(wZaikoSu, "#,###0")

    GetZaikoSu = True
End Function

Private Function DispMeisai() As Boolean
    DispMeisai = False
    
    H1lb_Kikai = ""
    H1lb_KikNo = ""
    H1lb_Okiba = ""
    H1lb_Zaiko = ""
    
    gSL_Select = ""
    gSL_Select = gSL_Select & "SELECT MM1.値名称 AS 機械名, MM2.値名称 AS 機械NO名, HK.*"
    gSL_Select = gSL_Select & " FROM 品番工順マスタ HK"
    gSL_Select = gSL_Select & "      LEFT OUTER JOIN 名称マスタ MM1"
    gSL_Select = gSL_Select & "       ON  MM1.区分名称 = '機械'"
    gSL_Select = gSL_Select & "       AND MM1.値       = HK.機械"
    gSL_Select = gSL_Select & "      LEFT OUTER JOIN 名称マスタ MM2"
    gSL_Select = gSL_Select & "       ON  MM2.区分名称 = '機械NO'"
    gSL_Select = gSL_Select & "       AND MM2.値       = HK.機械NO"
    gSL_Select = gSL_Select & " WHERE HK.品番   = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   AND HK.取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   AND HK.受入   = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   AND HK.工順   = '" & RTrim(iH1_Kojun) & "'"
    If Not HKTRead(gSL_Select, 1) Then
''''    With CisFun
''''         .MB_Lines = 4
''''         .MB_MSG(2) = "    品番工順マスタ  未登録            "
''''         .MB_Title = ""
''''         .MB_Button = Error
''''         .MBOX
''''    End With
''''    iH1_Hinbn.SetFocus
        Exit Function
    End If
'    If HKT.棚卸プレシート発行 <> 0 Then                                 '      2019/12/17 ｺﾒﾝﾄ ('一個のみ)
''''    With CisFun
''''         .MB_Lines = 4
''''         .MB_MSG(2) = "    在庫区分=有ではありません            "
''''         .MB_Title = ""
''''         .MB_Button = Error
''''         .MBOX
''''    End With
''''    iH1_Hinbn.SetFocus
'        DispMeisai = True
'        Exit Function
'    End If
    
    With HKT
        H1lb_Kikai = .機械名
        H1lb_KikNo = .機械NO名
        H1lb_Okiba = .置場
    End With
    Call GetZaikoSu
    
    DispMeisai = True
End Function
'-----------------------------------------------------------------------------------' 2010.01.14 add end    ' 2010.01.29 upd end

