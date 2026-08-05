VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CKH0110 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "かんばん増減管理入力 "
   ClientHeight    =   10980
   ClientLeft      =   -30
   ClientTop       =   525
   ClientWidth     =   15315
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "ＭＳ ゴシック"
      Size            =   11.25
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10980
   ScaleWidth      =   15315
   WindowState     =   2  '最大化
   Begin VB.Timer Timer3 
      Left            =   1080
      Top             =   450
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
      Height          =   1845
      Left            =   10200
      ScaleHeight     =   1785
      ScaleWidth      =   4335
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   465
      Width           =   4395
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2565
         Top             =   1200
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
         Left            =   3405
         Top             =   1200
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
         Left            =   3390
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
         Top             =   615
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   885
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
      Begin CisBtn_60.CisBtn PB_BACK 
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
      Begin CisBtn_60.CisBtn PB_Hakou 
         Height          =   585
         Left            =   2565
         Top             =   615
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
         BtnCaption      =   "発行"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   65535
         BtnShadowColor  =   0
         PFCaption       =   "( F11 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   0
         PFShadowColor   =   16777215
         PFBackColor     =   65535
      End
      Begin CisBtn_60.CisBtn PB_Naiji 
         Height          =   585
         Left            =   1710
         Top             =   15
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
         BtnType         =   12
         BtnCaption      =   "内示"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16711935
         BtnShadowColor  =   16777215
         PFCaption       =   "( F4 )"
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
         PFBackColor     =   16744703
      End
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   45
         Top             =   1200
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
         BtnType         =   7
         BtnCaption      =   "新規"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   65535
         BtnShadowColor  =   0
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
         PFForeColor     =   0
         PFShadowColor   =   16777215
         PFBackColor     =   65535
      End
      Begin CisBtn_60.CisBtn PB_MNT 
         Height          =   585
         Left            =   885
         Top             =   1200
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
         BtnType         =   8
         BtnCaption      =   "修正"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   16776960
         BtnShadowColor  =   0
         PFCaption       =   "( F2 )"
         BeginProperty PFFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFForeColor     =   0
         PFShadowColor   =   16777215
         PFBackColor     =   16776960
      End
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1725
         Top             =   1200
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
         BtnType         =   4
         BtnCaption      =   "削除"
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
      Begin CisBtn_60.CisBtn PB_Zogen 
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "増減"
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
         PFCaption       =   "( F9 )"
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
      Begin CisBtn_60.CisBtn PB_Hejun 
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
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnCaption      =   "平準化"
         BeginProperty BtnFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BtnForeColor    =   8388736
         BtnShadowColor  =   16777215
         PFCaption       =   "( F10 )"
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
         PFBackColor     =   8388736
      End
   End
   Begin VB.PictureBox T1_Area1 
      Height          =   1425
      Left            =   600
      ScaleHeight     =   1365
      ScaleWidth      =   10110
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   6015
      Width           =   10170
      Begin vsFlexLib.vsFlexArray vsGrid1 
         Height          =   1275
         Left            =   30
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   45
         Width           =   10020
         _Version        =   196608
         _ExtentX        =   17674
         _ExtentY        =   2249
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
         FormatString    =   "№|設定 |増減区分|開 始 日   |終 了 日   |設定枚数 |便当枚数 |累積枚数 |完 了 日  ||"
         Rows            =   4
         Cols            =   11
         BackColor       =   12648447
         BackColorBkg    =   14737632
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   300
         GridLines       =   3
         SelectionMode   =   1
         AllowBigSelection=   0   'False
         BackColorAlternate=   12648447
      End
   End
   Begin VB.PictureBox B3_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   2565
      Left            =   8760
      ScaleHeight     =   2505
      ScaleWidth      =   5805
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   8160
      Width           =   5865
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   2400
         Left            =   60
         Top             =   60
         Width           =   5715
         _ExtentX        =   10081
         _ExtentY        =   4233
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
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D CIS3D28 
            Height          =   615
            Left            =   30
            Top             =   15
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   1085
            Caption         =   "  収容変更日･便"
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
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D B3lb_SHYmd 
               Height          =   390
               Left            =   1815
               Top             =   105
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   688
               BackColor       =   16777152
               Caption         =   "9999/99/99"
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
               cAlingnment     =   7
               cBoderStyle     =   1
            End
            Begin Cis3D_v60.CIS3D B3lb_SHBin 
               Height          =   390
               Left            =   3285
               Top             =   105
               Width           =   450
               _ExtentX        =   794
               _ExtentY        =   688
               BackColor       =   16777152
               Caption         =   "99"
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
               cAlingnment     =   7
               cBoderStyle     =   1
            End
         End
         Begin Cis3D_v60.CIS3D BLB_Sur 
            Height          =   615
            Index           =   1
            Left            =   15
            Top             =   1140
            Width           =   5670
            _ExtentX        =   10001
            _ExtentY        =   1085
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   615
               Index           =   5
               Left            =   1935
               Top             =   0
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   1085
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B3lb_Syuyo 
                  Height          =   360
                  Index           =   1
                  Left            =   105
                  Top             =   135
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "##,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -60
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   615
               Index           =   6
               Left            =   420
               Top             =   0
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   1085
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B3lb_Syuki 
                  Height          =   360
                  Index           =   1
                  Left            =   90
                  Top             =   135
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXXX"
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
                  cPositionX      =   45
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   615
               Index           =   7
               Left            =   4365
               Top             =   0
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   1085
               ForeColor       =   16711680
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB3_HMaisu 
                  Height          =   360
                  Index           =   1
                  Left            =   390
                  TabIndex        =   18
                  Top             =   135
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   635
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
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   615
               Index           =   8
               Left            =   3135
               Top             =   0
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   1085
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B3lb_HNo 
                  Height          =   360
                  Index           =   1
                  Left            =   135
                  Top             =   135
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -60
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   615
               Index           =   0
               Left            =   15
               Top             =   0
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   1085
               BackColor       =   16777152
               Caption         =   "1"
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
               cAlingnment     =   7
            End
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   615
            Index           =   0
            Left            =   4005
            Top             =   15
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   1085
            ForeColor       =   0
            Caption         =   " 最終№"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
            Begin Cis3D_v60.CIS3D B3lb_EndNo 
               Height          =   390
               Left            =   840
               Top             =   105
               Width           =   630
               _ExtentX        =   1111
               _ExtentY        =   688
               BackColor       =   16777152
               Caption         =   "###"
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
               cAlingnment     =   7
               cBoderStyle     =   1
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D21 
            Height          =   510
            Left            =   30
            Top             =   630
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   900
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   510
            Index           =   2
            Left            =   1950
            Top             =   630
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   900
            ForeColor       =   0
            Caption         =   "収容数"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   510
            Index           =   50
            Left            =   435
            Top             =   630
            Width           =   1515
            _ExtentX        =   2672
            _ExtentY        =   900
            ForeColor       =   0
            Caption         =   "収容器"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   510
            Index           =   1
            Left            =   4380
            Top             =   630
            Width           =   1305
            _ExtentX        =   2302
            _ExtentY        =   900
            ForeColor       =   16711680
            Caption         =   "追加枚数"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   510
            Index           =   3
            Left            =   3150
            Top             =   630
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   900
            ForeColor       =   0
            Caption         =   "発行済"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D BLB_Sur 
            Height          =   630
            Index           =   0
            Left            =   15
            Top             =   1755
            Width           =   5670
            _ExtentX        =   10001
            _ExtentY        =   1111
            ForeColor       =   16711680
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   630
               Index           =   4
               Left            =   1935
               Top             =   0
               Width           =   1200
               _ExtentX        =   2117
               _ExtentY        =   1111
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B3lb_Syuyo 
                  Height          =   360
                  Index           =   2
                  Left            =   105
                  Top             =   135
                  Width           =   990
                  _ExtentX        =   1746
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "##,###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -60
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   630
               Index           =   9
               Left            =   420
               Top             =   0
               Width           =   1515
               _ExtentX        =   2672
               _ExtentY        =   1111
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B3lb_Syuki 
                  Height          =   360
                  Index           =   2
                  Left            =   90
                  Top             =   135
                  Width           =   1305
                  _ExtentX        =   2302
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "XXXXXXXXXX"
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
                  cPositionX      =   45
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   630
               Index           =   10
               Left            =   4365
               Top             =   0
               Width           =   1305
               _ExtentX        =   2302
               _ExtentY        =   1111
               ForeColor       =   16711680
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin CisText_V60.CisText iB3_HMaisu 
                  Height          =   360
                  Index           =   2
                  Left            =   390
                  TabIndex        =   19
                  Top             =   135
                  Width           =   525
                  _ExtentX        =   926
                  _ExtentY        =   635
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
                  cFdAutoFormat   =   1
                  cGFormat        =   "###"
                  cILength        =   3
               End
            End
            Begin Cis3D_v60.CIS3D BLB_DD 
               Height          =   630
               Index           =   11
               Left            =   3135
               Top             =   0
               Width           =   1080
               _ExtentX        =   1905
               _ExtentY        =   1111
               ForeColor       =   0
               Caption         =   "　"
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
               cAlingnment     =   6
               Begin Cis3D_v60.CIS3D B3lb_HNo 
                  Height          =   360
                  Index           =   2
                  Left            =   135
                  Top             =   135
                  Width           =   765
                  _ExtentX        =   1349
                  _ExtentY        =   635
                  BackColor       =   16777152
                  Caption         =   "###"
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
                  cAlingnment     =   4
                  cBoderStyle     =   1
                  cPositionX      =   -60
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D20 
               Height          =   630
               Index           =   1
               Left            =   15
               Top             =   0
               Width           =   405
               _ExtentX        =   714
               _ExtentY        =   1111
               BackColor       =   16777152
               Caption         =   "2"
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
               cAlingnment     =   7
            End
         End
      End
   End
   Begin VB.PictureBox T2_Area1 
      Height          =   1440
      Left            =   600
      ScaleHeight     =   1380
      ScaleWidth      =   7395
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   9315
      Width           =   7455
      Begin vsFlexLib.vsFlexArray vsGrid2 
         Height          =   1290
         Left            =   30
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   45
         Width           =   7320
         _Version        =   196608
         _ExtentX        =   12912
         _ExtentY        =   2275
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
         FormatString    =   "№|設定 |開 始 日   |終 了 日   |便当枚数|日当枚数|月当枚数|||"
         Rows            =   4
         Cols            =   11
         BackColor       =   12648447
         BackColorBkg    =   14737632
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   300
         GridLines       =   3
         SelectionMode   =   1
         AllowBigSelection=   0   'False
         BackColorAlternate=   12648447
      End
   End
   Begin VB.PictureBox B3_Area3 
      Height          =   1755
      Left            =   5685
      ScaleHeight     =   1695
      ScaleWidth      =   8850
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   2370
      Width           =   8910
      Begin vsFlexLib.vsFlexArray vsGrid3 
         Height          =   1275
         Left            =   30
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   405
         Width           =   8805
         _Version        =   196608
         _ExtentX        =   15531
         _ExtentY        =   2249
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
         FormatString    =   "№ |内示年月|内示数   |日当枚数 |回転枚数 |前月差|平準化日 |平準化便 |"
         Rows            =   4
         BackColor       =   16777152
         BackColorBkg    =   14737632
         FocusRect       =   0
         HighLight       =   0
         RowHeightMin    =   300
         GridLines       =   3
         SelectionMode   =   1
         Begin Cis3D_v60.CIS3D B3lb_Dummy 
            Height          =   300
            Index           =   0
            Left            =   2685
            Top             =   300
            Visible         =   0   'False
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            BackColor       =   16777152
            Caption         =   "999(999)"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Dummy 
            Height          =   300
            Index           =   1
            Left            =   3855
            Top             =   300
            Visible         =   0   'False
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            BackColor       =   16777152
            Caption         =   "999(999)"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Dummy 
            Height          =   300
            Index           =   2
            Left            =   5025
            Top             =   300
            Visible         =   0   'False
            Width           =   795
            _ExtentX        =   1402
            _ExtentY        =   529
            BackColor       =   16777152
            Caption         =   "-999"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Dummy 
            Height          =   300
            Index           =   4
            Left            =   5835
            Top             =   300
            Visible         =   0   'False
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            BackColor       =   16777152
            Caption         =   "999(999)"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
         Begin Cis3D_v60.CIS3D B3lb_Dummy 
            Height          =   300
            Index           =   6
            Left            =   7005
            Top             =   300
            Visible         =   0   'False
            Width           =   1140
            _ExtentX        =   2011
            _ExtentY        =   529
            BackColor       =   16777152
            Caption         =   "999(999)"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   345
         Left            =   30
         Top             =   15
         Width           =   8820
         _ExtentX        =   15558
         _ExtentY        =   609
         BackColor       =   0
         ForeColor       =   16777215
         Caption         =   "<<  内示情報  >>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   6
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'なし
      Height          =   405
      Index           =   1
      Left            =   12555
      ScaleHeight     =   405
      ScaleWidth      =   2055
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   7755
      Width           =   2055
      Begin Cis3D_v60.CIS3D UC_3D6 
         Height          =   345
         Left            =   30
         Top             =   15
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   609
         ForeColor       =   0
         Caption         =   "発行総枚数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   9.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   50
         Begin Cis3D_v60.CIS3D iT1_HMaisu 
            Height          =   225
            Left            =   1035
            Top             =   60
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   397
            BackColor       =   12648447
            ForeColor       =   0
            Caption         =   "ZZ,ZZ9"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Times New Roman"
               Size            =   9
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
   Begin Cis3D_v60.CIS3D B3_Area4 
      Height          =   2610
      Left            =   11235
      Top             =   4815
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   4604
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
      cBoderColor1    =   8421504
      cBoderColor2    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   780
         Left            =   60
         Top             =   60
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "サイクル"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B3lb_Cycle 
            Height          =   360
            Left            =   165
            Top             =   330
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "X-XX-XX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   780
         Left            =   1605
         Top             =   60
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "かんばん係数"
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
         cAlingnment     =   6
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B3lb_Keisu 
            Height          =   360
            Left            =   300
            Top             =   330
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "99"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D22 
         Height          =   780
         Left            =   60
         Top             =   840
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "社内背番"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B3lb_SSeban 
            Height          =   360
            Left            =   120
            Top             =   330
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D7 
         Height          =   780
         Index           =   5
         Left            =   1170
         Top             =   840
         Width           =   1020
         _ExtentX        =   1799
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "車型"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B3lb_Syakei 
            Height          =   360
            Left            =   120
            Top             =   330
            Width           =   780
            _ExtentX        =   1376
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D58 
         Height          =   780
         Left            =   2190
         Top             =   840
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   1376
         ForeColor       =   0
         Caption         =   "手配区分"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D B3lb_TKbn 
            Height          =   360
            Left            =   150
            Top             =   330
            Width           =   810
            _ExtentX        =   1429
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXX"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D60 
         Height          =   930
         Left            =   60
         Top             =   1620
         Width           =   3240
         _ExtentX        =   5715
         _ExtentY        =   1640
         ForeColor       =   0
         Caption         =   "品    名"
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
         cAlingnment     =   6
         cPositionY      =   15
         Begin Cis3D_v60.CIS3D B3lb_Hinnm1 
            Height          =   270
            Left            =   90
            Top             =   300
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   476
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
         End
         Begin Cis3D_v60.CIS3D B3lb_Hinnm2 
            Height          =   270
            Left            =   90
            Top             =   570
            Width           =   3045
            _ExtentX        =   5371
            _ExtentY        =   476
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   1
            cBoderStyle     =   1
            cPositionX      =   30
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
      Height          =   930
      Left            =   600
      ScaleHeight     =   870
      ScaleWidth      =   8850
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   1260
      Width           =   8910
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   780
         Left            =   60
         Top             =   60
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   1376
         ForeColor       =   192
         Caption         =   "背番号"
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
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   360
            Left            =   150
            TabIndex        =   1
            Top             =   330
            Width           =   750
            _ExtentX        =   1323
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
            Text            =   "XXXXX"
            MaxLength       =   5
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   780
         Left            =   5445
         Top             =   60
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "品  番"
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
         Begin CisText_V60.CisText iH1_Hinbn 
            Height          =   360
            Left            =   120
            TabIndex        =   4
            Top             =   330
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Left            =   1095
         Top             =   60
         Width           =   3645
         _ExtentX        =   6429
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "取 引 先"
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
         Begin CisText_V60.CisText iH1_Torcd 
            Height          =   360
            Left            =   60
            TabIndex        =   2
            Top             =   330
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1050
            Top             =   330
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
      Begin Cis3D_v60.CIS3D CIS3D15 
         Height          =   780
         Left            =   4740
         Top             =   60
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   1376
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
         cPositionY      =   30
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   150
            TabIndex        =   3
            Top             =   330
            Width           =   390
            _ExtentX        =   688
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
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   960
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   -1320
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   60
      Top             =   450
   End
   Begin VB.Timer Timer2 
      Left            =   540
      Top             =   450
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 かんばん管理入力(増減・平準化)  】"
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
      Begin Cis3D_v60.CIS3D K_Inkbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   "増減"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
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
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14355
         Top             =   75
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
         Left            =   11835
         Top             =   30
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   556
         BackColor       =   14737632
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
         cBoderStyle     =   1
         cPositionX      =   50
      End
   End
   Begin Cis3D_v60.CIS3D B3_Area2 
      Height          =   1740
      Left            =   600
      Top             =   2385
      Width           =   4950
      _ExtentX        =   8731
      _ExtentY        =   3069
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
      cBoderColor1    =   8421504
      cBoderColor2    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D33 
         Height          =   720
         Left            =   1140
         Top             =   405
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1270
         ForeColor       =   0
         Caption         =   "有効"
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
         cAlingnment     =   6
         Begin Cis3D_v60.CIS3D B3lb_Maisu 
            Height          =   330
            Index           =   1
            Left            =   90
            Top             =   285
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "#,###"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D35 
         Height          =   720
         Left            =   2085
         Top             =   405
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1270
         ForeColor       =   0
         Caption         =   "無効"
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
         cAlingnment     =   6
         Begin Cis3D_v60.CIS3D B3lb_Maisu 
            Height          =   330
            Index           =   2
            Left            =   90
            Top             =   285
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "#,###"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D37 
         Height          =   720
         Left            =   3030
         Top             =   405
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1270
         ForeColor       =   0
         Caption         =   "滞留"
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
         cAlingnment     =   6
         Begin Cis3D_v60.CIS3D B3lb_Maisu 
            Height          =   330
            Index           =   3
            Left            =   90
            Top             =   285
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "#,###"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D39 
         Height          =   720
         Left            =   3975
         Top             =   405
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   1270
         ForeColor       =   0
         Caption         =   "未流動"
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
         cAlingnment     =   6
         Begin Cis3D_v60.CIS3D B3lb_Maisu 
            Height          =   330
            Index           =   4
            Left            =   90
            Top             =   285
            Width           =   735
            _ExtentX        =   1296
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "#,###"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D51 
         Height          =   720
         Left            =   45
         Top             =   405
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   1270
         ForeColor       =   0
         Caption         =   "発行枚数"
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
         cAlingnment     =   6
         Begin Cis3D_v60.CIS3D B3lb_Maisu 
            Height          =   330
            Index           =   0
            Left            =   120
            Top             =   285
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "#,###"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D56 
         Height          =   345
         Left            =   45
         Top             =   45
         Width           =   4875
         _ExtentX        =   8599
         _ExtentY        =   609
         BackColor       =   0
         ForeColor       =   16777215
         Caption         =   "<<  流動管理情報  >>"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   6
      End
      Begin Cis3D_v60.CIS3D CIS3D8 
         Height          =   570
         Left            =   45
         Top             =   1125
         Width           =   2430
         _ExtentX        =   4286
         _ExtentY        =   1005
         ForeColor       =   0
         Caption         =   " 発注枚数"
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
         cAlingnment     =   1
         Begin Cis3D_v60.CIS3D B3lb_Maisu 
            Height          =   330
            Index           =   5
            Left            =   1320
            Top             =   105
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "#,###"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   570
         Left            =   2475
         Top             =   1125
         Width           =   2445
         _ExtentX        =   4313
         _ExtentY        =   1005
         ForeColor       =   0
         Caption         =   " 読取枚数"
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
         cAlingnment     =   1
         Begin Cis3D_v60.CIS3D B3lb_Maisu 
            Height          =   330
            Index           =   6
            Left            =   1245
            Top             =   90
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   582
            BackColor       =   16777152
            Caption         =   "#,###"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
   End
   Begin Cis3D_v60.CIS3D B2_Area1 
      Height          =   1140
      Left            =   600
      Top             =   8160
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   2011
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
      cBoderColor1    =   8421504
      cBoderColor2    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   510
         Index           =   2
         Left            =   1110
         Top             =   60
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "開始日"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin CisYMD_With_Btn.CisYMDwB iB2_SKYmd 
            Height          =   360
            Left            =   840
            TabIndex        =   13
            Top             =   75
            Width           =   1860
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
            Object.Height          =   360
            Object.Width           =   1860
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   720
         Index           =   6
         Left            =   4980
         Top             =   360
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "日当"
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
         cAlingnment     =   6
         Begin CisText_V60.CisText iB2_DaySu 
            Height          =   360
            Left            =   270
            TabIndex        =   16
            Top             =   285
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   635
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
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   1020
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1799
         ForeColor       =   16711680
         Caption         =   "設定"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   120
         Begin VB.CheckBox B2Chk_SKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "有効"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   12
            Top             =   540
            Width           =   810
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   720
         Index           =   5
         Left            =   3930
         Top             =   360
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "便当"
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
         cAlingnment     =   6
         Begin CisText_V60.CisText iB2_BinSu 
            Height          =   360
            Left            =   270
            TabIndex        =   15
            Top             =   285
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   635
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
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   300
         Index           =   4
         Left            =   3930
         Top             =   60
         Width           =   3450
         _ExtentX        =   6085
         _ExtentY        =   529
         ForeColor       =   16576
         Caption         =   "上限枚数"
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
         cAlingnment     =   6
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   720
         Index           =   7
         Left            =   6030
         Top             =   360
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   1270
         ForeColor       =   16711680
         Caption         =   "月当"
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
         cAlingnment     =   6
         Begin CisText_V60.CisText iB2_MonSu 
            Height          =   360
            Left            =   150
            TabIndex        =   17
            Top             =   285
            Width           =   1005
            _ExtentX        =   1773
            _ExtentY        =   635
            cFormat         =   "###,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Text            =   "###,###"
            MaxLength       =   7
            cDataType       =   1
            cDataReplace    =   1
            cFdAutoFormat   =   1
            cGFormat        =   "######"
            cILength        =   6
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   510
         Index           =   3
         Left            =   1110
         Top             =   570
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "終了日"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin CisYMD_With_Btn.CisYMDwB iB2_EKYmd 
            Height          =   360
            Left            =   840
            TabIndex        =   14
            Top             =   60
            Width           =   1860
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
            Object.Height          =   360
            Object.Width           =   1860
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D48 
      Height          =   420
      Index           =   2
      Left            =   600
      Top             =   7725
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   741
      BackColor       =   0
      ForeColor       =   65535
      Caption         =   "【　かんばん平準化管理　】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cAlingnment     =   7
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Index           =   2
         Left            =   15
         Top             =   15
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   661
         ForeColor       =   65535
         Caption         =   "新規"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cPositionX      =   30
      End
   End
   Begin Cis3D_v60.CIS3D B1_Area1 
      Height          =   1140
      Left            =   600
      Top             =   4815
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   2011
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
      cBoderColor1    =   8421504
      cBoderColor2    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1020
         Index           =   2
         Left            =   1110
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1799
         ForeColor       =   16711680
         Caption         =   "増減区分"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   30
         Begin VB.OptionButton B1Op_ZGKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "減"
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
            Index           =   0
            Left            =   240
            TabIndex        =   6
            Top             =   390
            Value           =   -1  'True
            Width           =   570
         End
         Begin VB.OptionButton B1Op_ZGKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "増"
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
            Index           =   1
            Left            =   240
            TabIndex        =   7
            Top             =   690
            Width           =   570
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   510
         Index           =   3
         Left            =   2160
         Top             =   60
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "開始日"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin CisYMD_With_Btn.CisYMDwB iB1_SKYmd 
            Height          =   360
            Left            =   840
            TabIndex        =   8
            Top             =   75
            Width           =   1860
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
            Object.Height          =   360
            Object.Width           =   1860
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1020
         Index           =   5
         Left            =   4980
         Top             =   60
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1799
         ForeColor       =   16711680
         Caption         =   "設定枚数"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   120
         Begin CisText_V60.CisText iB1_SetSu 
            Height          =   360
            Left            =   330
            TabIndex        =   10
            Top             =   510
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   635
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
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1020
         Index           =   7
         Left            =   7320
         Top             =   60
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1799
         ForeColor       =   0
         Caption         =   "累積枚数"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   120
         Begin Cis3D_v60.CIS3D B1lb_RuiSu 
            Height          =   360
            Left            =   120
            Top             =   510
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "#,###"
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
            cAlingnment     =   4
            cBoderStyle     =   1
            cPositionX      =   -60
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1020
         Index           =   1
         Left            =   60
         Top             =   60
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   1799
         ForeColor       =   16711680
         Caption         =   "設定"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   120
         Begin VB.CheckBox B1Chk_SKbn 
            BackColor       =   &H00C0C0C0&
            Caption         =   "有効"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   150
            TabIndex        =   5
            Top             =   540
            Width           =   810
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1020
         Index           =   6
         Left            =   6150
         Top             =   60
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1799
         ForeColor       =   16711680
         Caption         =   "便当枚数"
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
         cAlingnment     =   6
         cPositionX      =   30
         cPositionY      =   120
         Begin CisText_V60.CisText iB1_BinSu 
            Height          =   360
            Left            =   330
            TabIndex        =   11
            Top             =   510
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   635
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
            cFdAutoFormat   =   1
            cGFormat        =   "###"
            cILength        =   3
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   1020
         Index           =   8
         Left            =   8490
         Top             =   60
         Width           =   1605
         _ExtentX        =   2831
         _ExtentY        =   1799
         ForeColor       =   0
         Caption         =   "完 了 日"
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
         cAlingnment     =   6
         cPositionY      =   120
         Begin Cis3D_v60.CIS3D B1lb_KanYmd 
            Height          =   360
            Left            =   120
            Top             =   510
            Width           =   1350
            _ExtentX        =   2381
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "9999/99/99"
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   510
         Index           =   4
         Left            =   2160
         Top             =   570
         Width           =   2820
         _ExtentX        =   4974
         _ExtentY        =   900
         ForeColor       =   16711680
         Caption         =   "終了日"
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
         cAlingnment     =   1
         cPositionX      =   60
         Begin CisYMD_With_Btn.CisYMDwB iB1_EKYmd 
            Height          =   360
            Left            =   840
            TabIndex        =   9
            Top             =   75
            Width           =   1860
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
            Object.Height          =   360
            Object.Width           =   1860
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D48 
      Height          =   420
      Index           =   1
      Left            =   600
      Top             =   4365
      Width           =   10170
      _ExtentX        =   17939
      _ExtentY        =   741
      BackColor       =   0
      ForeColor       =   65535
      Caption         =   "【　かんばん増減管理　】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cAlingnment     =   7
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Index           =   1
         Left            =   30
         Top             =   15
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   661
         ForeColor       =   65535
         Caption         =   "新規"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cPositionX      =   30
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D48 
      Height          =   435
      Index           =   0
      Left            =   8775
      Top             =   7725
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   767
      BackColor       =   0
      ForeColor       =   65535
      Caption         =   "【　かんばん発行　】"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cAlingnment     =   7
      cPositionX      =   -630
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Index           =   3
         Left            =   15
         Top             =   15
         Width           =   750
         _ExtentX        =   1323
         _ExtentY        =   661
         ForeColor       =   65535
         Caption         =   "新規"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   7
         cPositionX      =   30
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D48 
      Height          =   420
      Index           =   3
      Left            =   11235
      Top             =   4365
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   741
      BackColor       =   0
      ForeColor       =   16777215
      Caption         =   "<<  マスタ情報　>>"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   14.25
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
      cAlingnment     =   7
   End
End
Attribute VB_Name = "CKH0110"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   かんばん管理入力（増減・平準化）
'**       フォームID    :   CKH0110
'**       処理概要      :
'**
'**       作  成  日    :   2004/01/23  By CIS
'**       変  更  日    :   2004/12/03  By CIS 内示かんばん情報照会との連携時不具合修正
'**                                            内示情報表示
'**       変  更  日    :   2005/03/01  By CIS 最終発行連番を流動管理テーブルより表示に変更
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl

    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim InputNo                 As Integer      ' 入力番号
    Dim InputKB                 As String       ' 入力区分
    Dim InputNM                 As String       ' 入力名称
    Dim CisVsGrid1              As New CisVsGrid3
    Dim CisVsGrid2              As New CisVsGrid3
    Dim CisVsGrid3              As New CisVsGrid3

'   <><><> Copy from CKN0220Bas <><><>
    Dim Kn_Jisya                As String
    Dim Kn_JisyaNM              As String
    Dim Kn_JisyaKJNM            As String

'   退避領域
    Dim Old_SKYmd               As String       ' 開始日（修正前）
    Dim New_EndNo               As Integer      ' 最終発行連番
    
    Dim mRType                  As Integer
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    If RTrim(Rv_Call2) = "" Then
    '   #-------------------#
    '   # 二 重 起 動 防 止  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
        If Not IniGet Then End
    End If
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
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
    Timer3.Interval = 500    ' 入力区分点滅間隔
    Timer3.Enabled = True    ' 入力区分点滅禁止

    Select Case App.Title
        Case "PKH0130"
            mRType = 1
        Case "PKH0140"
            mRType = 2
        Case Else
            mRType = 0
    End Select
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
'   かんばん増減管理
    With CisVsGrid1
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
        .PatanMax = 1
        .SelectCol = 10
        .InitGet = False
        .Init
    End With
    Call GridInit1
'   かんばん平準化管理
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
        .PatanMax = 1
        .SelectCol = 8
        .InitGet = False
        .Init
    End With
    Call GridInit2
'   内示情報
    With CisVsGrid3
        Set .GridObj = vsGrid3
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 300
        .PatanMax = 1
        .SelectCol = 9
        .InitGet = False
        .Init
    End With
    Call GridInit3

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    Call InputName("PB_Zogen")
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)

    If RTrim(Rv_Call2) = "" Then Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub

    B1Op_ZGKbn(0) = True

' 自社情報取得
    Kn_Jisya = ""
    Kn_JisyaNM = ""
    Kn_JisyaKJNM = ""
    
    Kn_Jisya = RTrim(gCompany)
' 自社伝票発行用(工場)
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 7
    If Not IRNRead("", 1) Then
        Call ItemsClearIRN
        IRN.bk = "Option": IRN.rno = 7
        IRN.oyian = "0"
        IRN.oyian2 = "00"
    End If
    ' 工場
    Kn_Jisya = Kn_Jisya & RTrim(IRN.oyian)
    
    gSL_Select = "Select * From 取引先マスタ "
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & Kn_Jisya & "'"
    If TRMRead(gSL_Select, 1) Then
       Kn_JisyaNM = TRM.略称
       Kn_JisyaKJNM = TRM.工場名
    End If

    If RTrim(Rv_Call2) <> "" Then
        iH1_Seban = RTrim(RV_Seban)
        iH1_Torcd = RTrim(RV_Torcd)
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RTrim(RV_Ukeir)
        iH1_Hinbn = RTrim(RV_Hinbn)
        Call PB_ENT_Click
    End If
    
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
           Case vbKeyF1:        If PB_ADD.Visible Then Call PB_ADD_Click        '【新規】
           Case vbKeyF2:        If PB_MNT.Visible Then Call PB_MNT_Click        '【修正】
           Case vbKeyF3:        If PB_DEL.Visible Then Call PB_DEL_Click        '【削除】
           Case vbKeyF4:        If PB_Naiji.Visible Then Call PB_Naiji_Click    '【内示】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click      '【検索】
           Case vbKeyF9:        If PB_Zogen.Visible Then Call PB_Zogen_Click    '【増減】
           Case vbKeyF10:       If PB_Hejun.Visible Then Call PB_Hejun_Click    '【平準化】
           Case vbKeyF11:       If PB_Hakou.Visible Then Call PB_Hakou_Click    '【発行】
           Case vbKeyPageUp:    If PB_BACK.Visible Then Call PB_Back_Click      '【後退】
           Case vbKeyPageDown:  If PB_NEXT.Visible Then Call PB_Next_Click      '【前進】
           Case vbKeyEscape:    If PB_CAN.Visible Then Call PB_CAN_Click        '【取消】
           Case vbKeyF12:       If PB_END.Visible Then Call PB_END_Click        '【終了】
           Case vbKeyReturn
                                If PB_ENT.Visible Then                          '【入力】
                                    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                    Call PB_ENT_Click
                                    Dummy.Enabled = False
                                End If
           Case Else:           Exit Function
    End Select

    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If RTrim(Rv_Call2) = "" Then
        Call CisDB.DBDISConnect
        Call CisFrm.UnLoadDisp
        End
    Else
        Me.Hide
    End If
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
    If ProcHB = "B1" Or _
       ProcHB = "B2" Then
        If SyoriKB = "ADD" Then
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call HeadBodyClear("B")
            Call InputName(InputKB)
            SyoriKB = "LOC"
            Call SyoriName(SyoriKB)
            iH1_Seban.SetFocus
            GoTo PB_CAN_Ed
        Else
            Select Case ProcHB
                Case "B1"
                    Call CisVsGrid1.vsColor(消去全部)
                    ProcHB = "T1"
                    Call DispChange(ProcHB)
                    Call SyoriName(SyoriKB)
                    vsGrid1.SetFocus
                Case "B2"
                    Call CisVsGrid2.vsColor(消去全部)
                    ProcHB = "T2"
                    Call DispChange(ProcHB)
                    Call SyoriName(SyoriKB)
                    vsGrid2.SetFocus
            End Select
            GoTo PB_CAN_Ed
        End If
    End If
    If ProcHB = "B3" Or _
       ProcHB = "T1" Or _
       ProcHB = "T2" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        Call InputName(InputKB)
        SyoriKB = "LOC"
        Call SyoriName(SyoriKB)
        iH1_Seban.SetFocus
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If ProcHB <> "H1" Then
        If SyoriKB <> "ADD" Then
            SyoriKB = "ADD"
            Select Case ProcHB
                Case "T1": ProcHB = "B1"
                Case "T2": ProcHB = "B2"
            End Select
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            Call HeadBodyClear("B", ProcHB)
        End If
        Select Case ProcHB
            Case "B1": B1Chk_SKbn.SetFocus
            Case "B2": B2Chk_SKbn.SetFocus
        End Select
    End If
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If ProcHB <> "H1" Then
        If SyoriKB <> "MNT" Then
            SyoriKB = "MNT"
            Select Case ProcHB
                Case "B1": ProcHB = "T1"
                Case "B2": ProcHB = "T2"
            End Select
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        End If
        Select Case ProcHB
            Case "T1": vsGrid1.SetFocus: Call BodySet1
            Case "T2": vsGrid2.SetFocus: Call BodySet2
        End Select
    End If
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If ProcHB <> "H1" Then
        If SyoriKB <> "DEL" Then
            SyoriKB = "DEL"
            Select Case ProcHB
                Case "B1": ProcHB = "T1"
                Case "B2": ProcHB = "T2"
            End Select
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        End If
        Select Case ProcHB
            Case "T1": vsGrid1.SetFocus: Call BodySet1
            Case "T2": vsGrid2.SetFocus: Call BodySet2
        End Select
    End If
End Sub
'+----------------------------+
'+    増  減  キ ー(F9)
'+----------------------------+
Private Sub PB_Zogen_Click()
    If ProcHB = "H1" Then
        If InputNo <> 1 Then
            Call InputName("PB_Zogen")
        End If
        iH1_Seban.SetFocus
    End If
End Sub
'+----------------------------+
'+    平準化  キ ー(F10)
'+----------------------------+
Private Sub PB_Hejun_Click()
    If ProcHB = "H1" Then
        If InputNo <> 2 Then
            Call InputName("PB_Hejun")
        End If
        iH1_Seban.SetFocus
    End If
End Sub
'+----------------------------+
'+    発  行  キ ー(F11)
'+----------------------------+
Private Sub PB_Hakou_Click()
    If ProcHB = "H1" Then
        If InputNo <> 3 Then
            Call InputName("PB_Hakou")
        End If
        iH1_Seban.SetFocus
    End If
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    If ProcHB = "H1" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Back")
        Call DBInput("Back")

        iH1_Seban.SetFocus

        Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    If ProcHB = "H1" Then
        Call CisFrm.MousePT(11)

        Call SqlSelectSet("Next")
        Call DBInput("Next")

        iH1_Seban.SetFocus

        Call CisFrm.MousePT(1)
    End If
End Sub
'+----------------------------+
'+    内示かんばん情報表示      +
'+----------------------------+
Private Sub PB_Naiji_Click()
    Rv_Call3 = "C"
    RV_Rtn = False

    CKH0410.Show vbModal
    Unload CKH0410
    Set CKH0410 = Nothing

    If RV_Rtn Then
        iH1_Seban = RV_Seban
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_Torcd
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
    End If

End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Seban" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_Ukeir: Exit Sub
'   日付
    If PB_Look.Tag = "iB1_SKYmd" Then iB1_SKYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB1_EKYmd" Then iB1_EKYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB2_SKYmd" Then iB2_SKYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iB2_EKYmd" Then iB2_EKYmd.ShowCalender: Exit Sub
End Sub
'+----------------------------+
'+      品番検索              +
'+----------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RVI_Hinmoku = 1
    RVI_Torcd = iH1_Torcd
    RVI_Ukeir = iH1_Ukeir
    RV_Call = ""

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Seban = RV_Seban
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        iH1_Ukeir = RV_Ukeir
        H1lb_Tornm = ""
        If TorNmGet(iH1_Torcd, 9) Then
           H1lb_Tornm = TRM_RName
        End If
    End If
End Sub
'+----------------------------+
'+      取引先検索            +
'+----------------------------+
Private Sub Look_Torcd()
    RV_Left = 0
    RV_Top = 0
    RV_TorKb = 9            ' [発注]手配先/契約先/社内検索
    RV_Call = "C"

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
    RV_Call = "C"

    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+-----------------------------+
'+  iH1_Seban
'+-----------------------------+
Private Sub iH1_Seban_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_Seban_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
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
'+-----------------------------+
'+  B1Chk_SKbn
'+-----------------------------+
Private Sub B1Chk_SKbn_Click()
    Select Case B1Chk_SKbn
        Case 0: B1Chk_SKbn.Caption = "有効"
        Case 1: B1Chk_SKbn.Caption = "無効"
    End Select
End Sub
'+-----------------------------+
'+  B1Op_ZGKbn
'+-----------------------------+
Private Sub B1Op_ZGKbn_Click(Index As Integer)
    If Index = 0 Then
        iB1_BinSu.Enabled = True
    Else
        iB1_BinSu.Enabled = False
        iB1_BinSu = ""
    End If
End Sub
'+-----------------------------+
'+  iB1_SKYmd
'+-----------------------------+
Private Sub iB1_SKYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_SKYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB1_EKYmd
'+-----------------------------+
Private Sub iB1_EKYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB1_EKYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  B2Chk_SKbn
'+-----------------------------+
Private Sub B2Chk_SKbn_Click()
    Select Case B2Chk_SKbn
        Case 0: B2Chk_SKbn.Caption = "有効"
        Case 1: B2Chk_SKbn.Caption = "無効"
    End Select
End Sub
'+-----------------------------+
'+  iB2_SKYmd
'+-----------------------------+
Private Sub iB2_SKYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB2_SKYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iB2_EKYmd
'+-----------------------------+
Private Sub iB2_EKYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iB2_EKYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'**************************************
'*                                    *
'*      入力区分により区分名表示        *
'*                                    *
'**************************************
Private Sub InputName(INKB As String)
    InputKB = INKB
    K_Sykbnm(1).Visible = False: K_Sykbnm(1).Caption = ""
    K_Sykbnm(2).Visible = False: K_Sykbnm(2).Caption = ""
    K_Sykbnm(3).Visible = False: K_Sykbnm(3).Caption = ""
    Select Case InputKB
        Case "PB_Zogen"
            InputNo = 1
            InputNM = "増減"
        Case "PB_Hejun"
            InputNo = 2
            InputNM = "平準化"
        Case "PB_Hakou"
            InputNo = 3
            InputNM = "発行"
        Case Else
            InputNo = 0
            SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
'       * PF-Key 表示設定 *
'        If gObj.Name = "PB_Naiji" Or
        If gObj.Name = "PB_Zogen" Or _
           gObj.Name = "PB_Hejun" Or _
           gObj.Name = "PB_Hakou" Then
           gObj.Visible = Not (gObj.Name = INKB)
        End If
    Next gObj
    Timer3.Enabled = False
    If InputNM = "" Then
        K_Inkbnm.Visible = False
        K_Inkbnm.Caption = ""
        Exit Sub
    End If
    K_Inkbnm.Caption = InputNM
    K_Inkbnm.Visible = True
    Timer3.Enabled = True
End Sub
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Static w_Color      As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
    Select Case SyoriKB
        Case "ADD"
            SyoriNM = "追加"
            K_Sykbnm(InputNo).ForeColor = &HFFFF&   '(黄色)
        Case "MNT"
            SyoriNM = "修正"
            K_Sykbnm(InputNo).ForeColor = &HFF0000  '(青色)
        Case "DEL"
            SyoriNM = "削除"
            K_Sykbnm(InputNo).ForeColor = &HFF&     '(赤色)
            w_Color = &HFF&
        Case "LOK"
            SyoriNM = "照会"
            K_Sykbnm(InputNo).ForeColor = &HFF00FF   '(ピンク)
        Case Else
            SyoriNM = ""
    End Select
    For Each gObj In Me.Controls
        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
'       * PF-Key 表示設定 *
        If SYKB <> "LOC" Then
            If gObj.Name = "PB_ADD" Or _
               gObj.Name = "PB_MNT" Or _
               gObj.Name = "PB_DEL" Or _
               gObj.Name = "PB_LOK" Then
               gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
            End If
        End If
    Next gObj
    Timer2.Enabled = False
    If SyoriNM = "" Then
        K_Sykbnm(InputNo).Visible = False
        K_Sykbnm(InputNo).Caption = ""
        Exit Sub
    End If
    K_Sykbnm(InputNo).Caption = SyoriNM
    K_Sykbnm(InputNo).Visible = True
    Timer2.Enabled = True
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*******************************
'*      処理区分名点滅制御      *
'*******************************
Private Sub Timer2_Timer()
    If K_Sykbnm(InputNo).Caption <> "" Then
       K_Sykbnm(InputNo).Caption = ""
    Else
       K_Sykbnm(InputNo).Caption = SyoriNM
    End If
End Sub
'*******************************
'*      入力区分名点滅制御      *
'*******************************
Private Sub Timer3_Timer()
    If K_Inkbnm.Caption <> "" Then
       K_Inkbnm.Caption = ""
    Else
       K_Inkbnm.Caption = InputNM
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
    Dim T1Mode      As Boolean
    Dim T1Color     As Long
    Dim B2Mode      As Boolean
    Dim B2Color     As Long
    Dim T2Mode      As Boolean
    Dim T2Color     As Long
    Dim B3Mode      As Boolean
    Dim B3Color     As Long

    H1Mode = False: H1Color = gPLostSel
    B1Mode = False: B1Color = gPLostSel
    T1Mode = False: T1Color = gPLostSel
    B2Mode = False: B2Color = gPLostSel
    T2Mode = False: T2Color = gPLostSel
    B3Mode = False: B3Color = gPLostSel
' ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then
       Dummy.Enabled = True
       Dummy.SetFocus
    End If
'
    Select Case Pro
        Case "H1": H1Mode = True: H1Color = gPGotSel
        Case "B1": B1Mode = True: B1Color = gPGotSel
        Case "T1": T1Mode = True: T1Color = gPGotSel
        Case "B2": B2Mode = True: B2Color = gPGotSel
        Case "T2": T2Mode = True: T2Color = gPGotSel
        Case "B3": B3Mode = True: B3Color = gPGotSel
    End Select
    H1_Area1.Enabled = H1Mode: H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode: B1_Area1.BackColor = B1Color
    T1_Area1.Enabled = T1Mode: T1_Area1.BackColor = T1Color
    B2_Area1.Enabled = B2Mode: B2_Area1.BackColor = B2Color
    T2_Area1.Enabled = T2Mode: T2_Area1.BackColor = T2Color
    B3_Area1.Enabled = B3Mode: B3_Area1.BackColor = B3Color

    PB_Naiji.Visible = H1Mode
    
    PB_Zogen.Visible = H1Mode
    PB_Hejun.Visible = H1Mode
    PB_Hakou.Visible = H1Mode
    PB_ADD.Visible = False
    PB_MNT.Visible = False
    PB_DEL.Visible = False
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_Look.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode

    Select Case Pro
        Case "T1": PB_ENT.BtnCaption = "選択"
        Case "T2": PB_ENT.BtnCaption = "選択"
        Case Else: PB_ENT.BtnCaption = "入力"
    End Select
    If RTrim(Rv_Call2) = "" Then
        PB_END.BtnCaption = "終了"
    Else
        PB_END.BtnCaption = "戻る"
        '+---------------------------------- 2004/12/03 Insert
        PB_END.Visible = True
        PB_Naiji.Visible = False
        '+---------------------------------- 2004/12/03 Insert End
    End If
    vsGrid1.TabStop = T1Mode
    vsGrid2.TabStop = T2Mode
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
    B1Op_ZGKbn(0) = True
    If HeadBody = "B" And ClrType = "ALL" Then
        Call CisVsGrid1.Clear
        Call CisVsGrid2.Clear
        Call CisVsGrid3.Clear
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        Select Case InputNo
            Case 1: ProcHB = "B1"
            Case 2: ProcHB = "B2"
            Case 3
                ProcHB = "B3"
                SyoriKB = "ADD"
                Call DispChange(ProcHB)
                Call SyoriName("LOC")
                GoTo ReturnPress_Ed
        End Select
        SyoriKB = "ADD"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B", ProcHB)
        Call GridSet1
        If SyoriKB = "MNT" Then
            ProcHB = "T1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            vsGrid1.SetFocus
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B2" Then
        If Not Body2Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B", ProcHB)
        Call GridSet2
        If SyoriKB = "MNT" Then
            ProcHB = "T2"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
            vsGrid2.SetFocus
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B3" Then
        If Not Body3Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call HeadBodyClear("B")
        Call InputName(InputKB)
        SyoriKB = "LOC"
        Call SyoriName(SyoriKB)
        iH1_Seban.SetFocus
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "T1" Then Call FromGrid1
    If ProcHB = "T2" Then Call FromGrid2
ReturnPress_Ed:
   Call CisFrm.MousePT(1)
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False

'   背番号
    If Trim(iH1_Seban) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    背番号を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Seban.SetFocus
       Exit Function
    End If

    gSL_Select = "select * from 品番取引先マスタ"
    gSL_Select = gSL_Select & " where 品番 = '" & RTrim(iH1_Hinbn) & "'"
    gSL_Select = gSL_Select & "   and 取引先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(iH1_Ukeir) & "'"
    gSL_Select = gSL_Select & "   and 背番号 = '" & RTrim(iH1_Seban) & "'"
    gSL_Select = gSL_Select & "   and 手配区分 <> '2'"
    If Not HTMRead(gSL_Select, 1, 1) Then
'< 2005.03.01 MNT START >--------------------------------------------------------------------
'
'        gSL_Select = ""
'        gSL_Select = gSL_Select & "Select"
'        gSL_Select = gSL_Select & "  htm.背番号"
'        gSL_Select = gSL_Select & ", htm.取引先"
'        gSL_Select = gSL_Select & ", htm.受入"
'        gSL_Select = gSL_Select & ", htm.品番"
'        gSL_Select = gSL_Select & ", hnm.表示品番"
'        gSL_Select = gSL_Select & ", hnm.社内背番"
'        gSL_Select = gSL_Select & ", htm.最終発行連番"
'        gSL_Select = gSL_Select & ", rtrim(trm.略称) 取引先名"
'        gSL_Select = gSL_Select & ", htm.手配区分"
'        gSL_Select = gSL_Select & " from 品番取引先マスタ htm"
'        gSL_Select = gSL_Select & "      left outer join 品番マスタ hnm"
'        gSL_Select = gSL_Select & "        on hnm.品番 = htm.品番"
'        gSL_Select = gSL_Select & "      left outer join 取引先マスタ trm"
'        gSL_Select = gSL_Select & "        on trm.取引先CD = htm.取引先"
'        gSL_Select = gSL_Select & "       and trm.取引先区分 = 1"
'        gSL_Select = gSL_Select & " where htm.背番号 = '" & iH1_Seban & "'"
'        gSL_Select = gSL_Select & "   and htm.品目 = 1"
'        gSL_Select = gSL_Select & "   and htm.SK区分 = 0"
'        gSL_Select = gSL_Select & "   and htm.手配区分 <> '2'"
'        gSL_Select = gSL_Select & " order by"
'        gSL_Select = gSL_Select & "  htm.品番"
'        gSL_Select = gSL_Select & ", htm.取引先"
'        gSL_Select = gSL_Select & ", htm.受入"

        gSL_Select = ""
        gSL_Select = gSL_Select & "Select"
        gSL_Select = gSL_Select & "  htm.背番号"
        gSL_Select = gSL_Select & ", htm.取引先"
        gSL_Select = gSL_Select & ", htm.受入"
        gSL_Select = gSL_Select & ", htm.品番"
        gSL_Select = gSL_Select & ", hnm.表示品番"
        gSL_Select = gSL_Select & ", hnm.社内背番"
        
        gSL_Select = gSL_Select & ", Case When htm.最終発行連番 > KN.発行NO"
        gSL_Select = gSL_Select & "  Then htm.最終発行連番 Else KN.発行NO  End 最終発行連番"
        
        gSL_Select = gSL_Select & ", rtrim(trm.略称) 取引先名"
        gSL_Select = gSL_Select & ", htm.手配区分"
        gSL_Select = gSL_Select & " from 品番取引先マスタ htm"
        gSL_Select = gSL_Select & "      left outer join 品番マスタ hnm"
        gSL_Select = gSL_Select & "        on hnm.品番 = htm.品番"
        gSL_Select = gSL_Select & "      left outer join 取引先マスタ trm"
        gSL_Select = gSL_Select & "        on trm.取引先CD = htm.取引先"
        gSL_Select = gSL_Select & "       and trm.取引先区分 = 1"
        gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 取引先, 受入, 品番, MAX(発行NO) 発行NO"
        gSL_Select = gSL_Select & " From かんばん流動管理テーブル"
        gSL_Select = gSL_Select & " GROUP BY 取引先, 受入, 品番 ) KN"
        gSL_Select = gSL_Select & " ON  KN.取引先 = htm.取引先"
        gSL_Select = gSL_Select & " AND KN.受入 = htm.受入"
        gSL_Select = gSL_Select & " AND KN.品番 = htm.品番"
        gSL_Select = gSL_Select & " where htm.背番号 = '" & iH1_Seban & "'"
        gSL_Select = gSL_Select & "   and htm.品目 = 1"
        gSL_Select = gSL_Select & "   and htm.SK区分 = 0"
        gSL_Select = gSL_Select & "   and htm.手配区分 <> '2'"
        gSL_Select = gSL_Select & " order by"
        gSL_Select = gSL_Select & "  htm.品番"
        gSL_Select = gSL_Select & ", htm.取引先"
        gSL_Select = gSL_Select & ", htm.受入"

'
'< 2005.03.01 MNT START >--------------------------------------------------------------------
 
        If HTMRead(gSL_Select, 0, 1) Then
            If CisDB.RecordCount > 1 Then
                CKN0215.Show vbModal
                Unload CKN0215
                Set CKN0215 = Nothing
                If RV_Rtn Then
                    iH1_Torcd = RV_TorcdK
                    iH1_Ukeir = RV_Ukeir
                    iH1_Hinbn = RV_Hinbn
                    Call iH1_Torcd_LostFocus
                Else
                    iH1_Seban.SetFocus
                    Exit Function
                End If
            Else
                iH1_Torcd = RTrim(HTM.取引先)
                iH1_Ukeir = RTrim(HTM.受入)
                iH1_Hinbn = RTrim(HTM.品番)
            End If
            Call HTMClose(1)
        Else
            Call HTMClose(1)

'           取引先
            If Trim(iH1_Torcd) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    取引先を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_Torcd.SetFocus
                Exit Function
            End If
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
'           受入
            If Trim(iH1_Ukeir) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    受入を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_Ukeir.SetFocus
                Exit Function
            End If
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
'           品番
            If Trim(iH1_Hinbn) = "" Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    品番を入力して下さい            "
                    .MB_Button = Error
                    .MBOX
                End With
                iH1_Hinbn.SetFocus
                Exit Function
            End If
'           <<< 品番取引先マスタ >>>
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    品番取引先マスタ未登録            "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iH1_Seban.SetFocus
            Exit Function

        End If
    End If

    Call SqlSelectSet("INV")
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
    Dim wSKYmd      As String
    Dim wEKYmd      As String

    Body1Chk = False

'   開始日
    If Trim(iB1_SKYmd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    開始日を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_SKYmd.SetFocus
       Exit Function
    Else
        If Not iB1_SKYmd.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    開始日を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_SKYmd.SetFocus
            Exit Function
        End If
    End If
'   終了日
    If Trim(iB1_EKYmd) <> "" Then
        If Not iB1_EKYmd.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    終了日を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_EKYmd.SetFocus
            Exit Function
        End If
    '   大小関係
        If iB1_SKYmd > iB1_EKYmd Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付を正しく入力して下さい（開始＞終了）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_SKYmd.SetFocus
            Exit Function
        End If
    End If
'   <><><><><> 重複ﾁｪｯｸ／設定：有効の時 <><><><><>
    If B1Chk_SKbn = 0 Then
        wSKYmd = iB1_SKYmd
        wEKYmd = iB1_EKYmd
        If Trim(iB1_EKYmd) = "" Then wEKYmd = "99991231"
        gSL_Select = ""
        gSL_Select = gSL_Select & "select kzt.開始日"
        gSL_Select = gSL_Select & "     , kzt.終了日"
        gSL_Select = gSL_Select & "  from かんばん増減テーブル kzt"
        gSL_Select = gSL_Select & " where kzt.取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   and kzt.受入 = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & "   and kzt.品番 = '" & RTrim(iH1_Hinbn) & "'"
        If SyoriKB = "MNT" Then
            gSL_Select = gSL_Select & "   and kzt.開始日 <> '" & RTrim(Old_SKYmd) & "'"
        End If
        gSL_Select = gSL_Select & "   and kzt.設定内容区分 = 0"
        gSL_Select = gSL_Select & "   and kzt.開始日 <= '" & wEKYmd & "'"
        gSL_Select = gSL_Select & "   and case when kzt.終了日 = ''"
        gSL_Select = gSL_Select & "            then '99991231'"
        gSL_Select = gSL_Select & "            else kzt.終了日"
        gSL_Select = gSL_Select & "        end >= '" & wSKYmd & "'"
        If KZTRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    下記の登録済日付と重複しています！            "
                .MB_MSG(4) = "        > 開始日 : " & Format(KZT.開始日, "0000/00/00")
                .MB_MSG(5) = "        > 終了日 : " & Format(KZT.終了日, "0000/00/00")
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB1_SKYmd.SetFocus
            Exit Function
        End If
    End If
'   設定枚数
    If CisFun.Val2(iB1_SetSu) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    設定枚数を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB1_SetSu.SetFocus
       Exit Function
    End If

    Body1Chk = True
End Function
'****************************
'*      Body2ﾁｪｯｸ処理
'****************************
Private Function Body2Chk() As Boolean
    Dim wSKYmd      As String
    Dim wEKYmd      As String
    Dim wCnt        As Integer

    Body2Chk = False

'   開始日
    If Trim(iB2_SKYmd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    開始日を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB2_SKYmd.SetFocus
       Exit Function
    Else
        If Not iB2_SKYmd.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    開始日を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB2_SKYmd.SetFocus
            Exit Function
        End If
    End If
'   終了日
    If Trim(iB2_EKYmd) <> "" Then
        If Not iB2_EKYmd.cChkResult Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    終了日を正しく入力して下さい            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB2_EKYmd.SetFocus
            Exit Function
        End If
    '   大小関係
        If iB2_SKYmd > iB2_EKYmd Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    日付を正しく入力して下さい（開始＞終了）            "
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB2_SKYmd.SetFocus
            Exit Function
        End If
    End If
'   <><><><><> 重複ﾁｪｯｸ／設定：有効の時 <><><><><>
    If B2Chk_SKbn = 0 Then
        wSKYmd = iB2_SKYmd
        wEKYmd = iB2_EKYmd
        If Trim(iB2_EKYmd) = "" Then wEKYmd = "99991231"
        gSL_Select = ""
        gSL_Select = gSL_Select & "select kjt.開始日"
        gSL_Select = gSL_Select & "     , kjt.終了日"
        gSL_Select = gSL_Select & "  from かんばん平準化テーブル kjt"
        gSL_Select = gSL_Select & " where kjt.取引先 = '" & RTrim(iH1_Torcd) & "'"
        gSL_Select = gSL_Select & "   and kjt.受入 = '" & RTrim(iH1_Ukeir) & "'"
        gSL_Select = gSL_Select & "   and kjt.品番 = '" & RTrim(iH1_Hinbn) & "'"
        If SyoriKB = "MNT" Then
            gSL_Select = gSL_Select & "   and kjt.開始日 <> '" & RTrim(Old_SKYmd) & "'"
        End If
        gSL_Select = gSL_Select & "   and kjt.設定内容区分 = 0"
        gSL_Select = gSL_Select & "   and kjt.開始日 <= '" & wEKYmd & "'"
        gSL_Select = gSL_Select & "   and case when kjt.終了日 = ''"
        gSL_Select = gSL_Select & "            then '99991231'"
        gSL_Select = gSL_Select & "            else kjt.終了日"
        gSL_Select = gSL_Select & "        end >= '" & wSKYmd & "'"
        If KJTRead(gSL_Select, 1) Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    下記の登録済日付と重複しています！            "
                .MB_MSG(4) = "        > 開始日 : " & Format(KJT.開始日, "0000/00/00")
                .MB_MSG(5) = "        > 終了日 : " & Format(KJT.終了日, "0000/00/00")
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
            iB2_SKYmd.SetFocus
            Exit Function
        End If
    End If
'   上限枚数
    wCnt = 0
    If CisFun.Val2(iB2_BinSu) <> 0 Then wCnt = wCnt + 1
    If CisFun.Val2(iB2_DaySu) <> 0 Then wCnt = wCnt + 1
    If CisFun.Val2(iB2_MonSu) <> 0 Then wCnt = wCnt + 1
    If wCnt <> 1 Then
        With CisFun
            .MB_Lines = 4
            Select Case wCnt
                Case 0:    .MB_MSG(2) = "    上限枚数を１箇所のみ入力して下さい            "
                Case Else: .MB_MSG(2) = "    上限枚数は１箇所のみ入力して下さい            "
            End Select
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB2_BinSu.SetFocus
       Exit Function
    End If

    Body2Chk = True
End Function
'****************************
'*      Body3ﾁｪｯｸ処理
'****************************
Private Function Body3Chk() As Boolean
    Body3Chk = False

'   追加枚数
    If CisFun.Val2(iB3_HMaisu(1)) = 0 And _
       CisFun.Val2(iB3_HMaisu(2)) = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    追加枚数を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iB3_HMaisu(1).SetFocus
       Exit Function
    End If

    Body3Chk = True
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
    If Not HTMRead(gSL_Select, 1) Then
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
        iH1_Seban = RTrim(HTM.背番号)
        iH1_Hinbn = RTrim(HTM.品番)
        iH1_Torcd = RTrim(HTM.取引先)
        iH1_Ukeir = RTrim(HTM.受入)
        H1lb_Tornm = RTrim(HTM.取引先名)
        If InputKey = "INV" Then
            If Not BodySet("ALL") Then GoTo DBInput_Ed
        End If
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
Private Sub SqlSelectSet(InputKey As String)
'
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  htm.品番"
    gSL_Select = gSL_Select & ", htm.取引先"
    gSL_Select = gSL_Select & ", htm.受入"
    gSL_Select = gSL_Select & ", htm.背番号"
    gSL_Select = gSL_Select & ", htm.か係数"
    gSL_Select = gSL_Select & ", htm.サイクル"
    gSL_Select = gSL_Select & ", htm.収容数1"
    gSL_Select = gSL_Select & ", htm.収容器1"
    gSL_Select = gSL_Select & ", htm.変更日"
    gSL_Select = gSL_Select & ", htm.変更便"
    gSL_Select = gSL_Select & ", htm.収容数2"
    gSL_Select = gSL_Select & ", htm.収容器2"
    
'-----------------------------------------------------< 2005.03.01 MNT START >
'    gSL_Select = gSL_Select & ", htm.最終発行連番"
    gSL_Select = gSL_Select & ", Case When htm.最終発行連番 > KN.発行NO"
    gSL_Select = gSL_Select & "  Then htm.最終発行連番 Else KN.発行NO  End 最終発行連番"
'-----------------------------------------------------< 2005.03.01 MNT  END  >
    
    gSL_Select = gSL_Select & ", isnull(trm.略称, '') 取引先名"
    gSL_Select = gSL_Select & ", isnull(hnm.社内背番, '') 社内背番"
    gSL_Select = gSL_Select & ", isnull(hnm.品名, '') 品名"
    gSL_Select = gSL_Select & ", isnull(hnm.車型, '') 車型"
    gSL_Select = gSL_Select & ", isnull(sym.値名称, '') 手配区分名"
    gSL_Select = gSL_Select & " from"
    gSL_Select = gSL_Select & "  品番取引先マスタ htm"
    gSL_Select = gSL_Select & "  left outer join 取引先マスタ trm"
    gSL_Select = gSL_Select & "    on trm.取引先CD = htm.取引先"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    
'< 2005.03.01 ADD START >----------------------------------------------------------------------------
'
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 取引先, 受入, 品番, MAX(発行NO) 発行NO"
    gSL_Select = gSL_Select & " From かんばん流動管理テーブル"
    gSL_Select = gSL_Select & " GROUP BY 取引先, 受入, 品番 ) KN"
    gSL_Select = gSL_Select & " ON  KN.取引先 = htm.取引先"
    gSL_Select = gSL_Select & " AND KN.受入 = htm.受入"
    gSL_Select = gSL_Select & " AND KN.品番 = htm.品番"
'
'< 2005.03.01 ADD  END  >----------------------------------------------------------------------------
    
    gSL_Select = gSL_Select & "  left outer join 品番マスタ hnm"
    gSL_Select = gSL_Select & "    on hnm.品番 = htm.品番"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sym"
    gSL_Select = gSL_Select & "    on sym.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   and sym.区分タイプ = ''"
    gSL_Select = gSL_Select & "   and sym.値 = htm.手配区分"
    gSL_Select = gSL_Select & " where htm.品目 = 1"
    gSL_Select = gSL_Select & "   and htm.SK区分 = 0"
    gSL_Select = gSL_Select & "   and htm.手配区分 <> '2'"
    gSL_Select = gSL_Select & "   and htm.背番号 <> ''"
    gSL_Select = gSL_Select & "   and"
    gSL_Select = gSL_Select & " ( htm.背番号 + space(5 - len(htm.背番号))"
    gSL_Select = gSL_Select & " + htm.品番 + space(25 - len(htm.品番))"
    gSL_Select = gSL_Select & " + htm.取引先 + space(7 - len(htm.取引先))"
    gSL_Select = gSL_Select & " + htm.受入 + space(2 - len(htm.受入))"
    gSL_Select = gSL_Select & " ) "
    Select Case InputKey
        Case "Back": gSL_Select = gSL_Select & "<"
        Case "Next": gSL_Select = gSL_Select & ">"
        Case Else:   gSL_Select = gSL_Select & "="
    End Select
    gSL_Select = gSL_Select & " ( '"
    gSL_Select = gSL_Select & iH1_Seban & Space(5 - CisFun.Len2(iH1_Seban))
    gSL_Select = gSL_Select & iH1_Hinbn & Space(25 - CisFun.Len2(iH1_Hinbn))
    gSL_Select = gSL_Select & iH1_Torcd & Space(7 - CisFun.Len2(iH1_Torcd))
    gSL_Select = gSL_Select & iH1_Ukeir & Space(2 - CisFun.Len2(iH1_Ukeir))
    gSL_Select = gSL_Select & "' )"
    Select Case InputKey
        Case "Back"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  htm.背番号 desc"
            gSL_Select = gSL_Select & ", htm.取引先 desc"
            gSL_Select = gSL_Select & ", htm.受入 desc"
            gSL_Select = gSL_Select & ", htm.品番 desc"
        Case "Next"
            gSL_Select = gSL_Select & " order by"
            gSL_Select = gSL_Select & "  htm.背番号 asc"
            gSL_Select = gSL_Select & ", htm.取引先 asc"
            gSL_Select = gSL_Select & ", htm.受入 asc"
            gSL_Select = gSL_Select & ", htm.品番 asc"
    End Select
'
End Sub
'**************************************
'*          BODY 内容ｾｯﾄ処理           *
'**************************************
Private Function BodySet(SetType As String) As Boolean
    BodySet = False
    '------------------------------------------------------- 2004/12/03 Insert
    Dim wCurrency       As Currency
    '------------------------------------------------------- 2004/12/03 Insert End

    With HTM
'       << マスタ情報 >>
        B3lb_TKbn = RTrim(.手配区分名)
        If Trim(.サイクル) <> "" Then
            B3lb_Cycle = CisFun.RSetFld(CisFun.Mid2(.サイクル, 1, 1), 1, "#") _
                 & "-" & CisFun.RSetFld(CisFun.Mid2(.サイクル, 2, 2), 2, "#") _
                 & "-" & CisFun.RSetFld(CisFun.Mid2(.サイクル, 4, 2), 2, "#")
        End If
        B3lb_Keisu = Format(.か係数, "0.00")
        B3lb_SSeban = RTrim(.社内背番)
        B3lb_Syakei = RTrim(.車型)
        B3lb_Hinnm1 = CisFun.Mid2(.品名, 1, 20)
        B3lb_Hinnm2 = CisFun.Mid2(.品名, 21, 20)
'       【かんばん発行】
        B3lb_Syuki(1) = RTrim(.収容器1)
        B3lb_Syuyo(1) = Format(.収容数1, "#,###")
        If Trim(.変更日) <> "" Then
            B3lb_SHYmd = Format(.変更日, "0000/00/00")
            B3lb_SHBin = RTrim(.変更便)
            B3lb_Syuki(2) = RTrim(.収容器2)
            B3lb_Syuyo(2) = Format(.収容数2, "#,###")
        End If
        B3lb_EndNo = Format(.最終発行連番, "#,###")
    End With
'   入力切替
    If B3lb_SHYmd = "" Then iB3_HMaisu(2).Enabled = False Else iB3_HMaisu(2).Enabled = True

'   <><><> かんばん流動管理テーブル <><><>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  sum(case when krt.発行区分 = 0 then 1 else 0 end) 未流件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 1 then 1 else 0 end) 有効件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 2 then 1 else 0 end) 無効件数"
    gSL_Select = gSL_Select & ", sum(case when krt.発行区分 = 3 then 1 else 0 end) 滞留件数"
    gSL_Select = gSL_Select & ", sum(case when krt.収容区分 = 1 then 1 else 0 end) 収容数1"
    gSL_Select = gSL_Select & ", sum(case when krt.収容区分 = 2 then 1 else 0 end) 収容数2"
    gSL_Select = gSL_Select & " from かんばん流動管理テーブル krt"
    gSL_Select = gSL_Select & " where krt.取引先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and krt.受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & "   and krt.品番 = '" & RTrim(HTM.品番) & "'"
    If KRTRead(gSL_Select, , 1) Then
        With KRT
'           << 流動管理情報 >>
            B3lb_Maisu(0) = Format(.有効件数 + .無効件数 + .滞留件数 + .未流件数, "#,###")
            B3lb_Maisu(1) = Format(.有効件数, "#,###")
            B3lb_Maisu(2) = Format(.無効件数, "#,###")
            B3lb_Maisu(3) = Format(.滞留件数, "#,###")
            B3lb_Maisu(4) = Format(.未流件数, "#,###")
'           【かんばん発行】
            B3lb_HNo(1) = Format(.収容数1, "#,###")
            B3lb_HNo(2) = Format(.収容数2, "#,###")
        End With
    End If
    Call KRTClose
'   <><><> 発注テーブル（入荷テーブル） <><><>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select sum(hct.枚数) 枚数"
    gSL_Select = gSL_Select & "  from 発注テーブル hct"
    gSL_Select = gSL_Select & " where hct.仕入先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and hct.受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & "   and hct.品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & "   and not exists"
    gSL_Select = gSL_Select & " ("
    gSL_Select = gSL_Select & " select nyt.発注管理NO"
    gSL_Select = gSL_Select & "   from 入荷テーブル nyt"
    gSL_Select = gSL_Select & "  where nyt.発注管理NO = hct.発注管理NO"
    gSL_Select = gSL_Select & " )"
    If HCTRead(gSL_Select, , 1) Then
        B3lb_Maisu(5) = Format(HCT.枚数, "#,###")
    End If
    Call HCTClose
'   <><><> 発注かんばん予定テーブル <><><>
    gSL_Select = ""
    gSL_Select = gSL_Select & "select sum(hky.枚数) 枚数"
    gSL_Select = gSL_Select & "  from 発注かんばん予定テーブル hky"
    gSL_Select = gSL_Select & " where hky.仕入先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and hky.受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & "   and hky.品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & "   and hky.処理区分 <> 2"
    If HKYRead(gSL_Select, , 1) Then
        B3lb_Maisu(6) = Format(HKY.枚数, "#,###")
    End If
    Call HKYClose

'   ■■■■■■■■■■■ << 内示情報 >> は 未設定 ■■■■■■■■■■
    '------------------------------------------------------- 2004/12/03 Insert
    Call CisVsGrid3.Clear
    ' 内示数
    gSL_Select = "select 内示年月,SUM(数量) 数量 from 発注内示テーブル"
    gSL_Select = gSL_Select & " Where 手配先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & "   and 内示年月 = 計画年月"
    gSL_Select = gSL_Select & "  Group by 内示年月 "
    gSL_Select = gSL_Select & " Order By 内示年月 Desc "
    gLong = 0
    If HNJRead(gSL_Select, , 1) Then
       With vsGrid3
           If CisVsGrid3.DispMax < CisDB.RecordCount + 1 Then
              .Rows = CisDB.RecordCount + 1
           End If
           Do Until Not HNJ_RDSTS
              gLong = gLong + 1
              .TextMatrix(gLong, 0) = gLong
              .TextMatrix(gLong, 1) = Mid(HNJ.内示年月, 1, 4) & "/" & Mid(HNJ.内示年月, 5, 2)
              .TextMatrix(gLong, 2) = Format(HNJ.数量, "#,##0")
              '年月
              .TextMatrix(gLong, 9) = HNJ.内示年月
              Call HNJReadNext(1)
           Loop
       End With
    End If
    Call HNJClose(1)
    ' 日量枚数計算
    gCnt1 = CisFun.Val2(B3lb_Syuyo(1))
    If gCnt1 <> 0 Then
        With vsGrid3
            For gLong = 1 To .Rows - 1
                If Trim(.TextMatrix(gLong, 1)) = "" Then Exit For
                
                gSL_Select = "Select * from カレンダマスタ "
                gSL_Select = gSL_Select & " Where 取引先区分 = 1"
                gSL_Select = gSL_Select & "   And 取引先 = ''"
                gSL_Select = gSL_Select & "   And 年 + 月 = '" & Trim(.TextMatrix(gLong, 9)) & "'"
                gSL_Select = gSL_Select & "   And 受入 = ''"
                If Not CLMRead(gSL_Select, 1, 1) Then
                   CLM.稼動日数 = 31
                End If
                wCurrency = CisFun.Val2(.TextMatrix(gLong, 2)) / CLM.稼動日数 / gCnt1
                gCnt2 = CisFun.HasuProc(CDbl(wCurrency), 切捨て)
                If gCnt2 = 0 And wCurrency <> 0 Then gCnt2 = 1
                .TextMatrix(gLong, 3) = gCnt2
                    
                    
            Next gLong
        End With
    End If
    ' かんばん回転枚数テーブル
    gSL_Select = "Select * From かんばん回転枚数テーブル "
    gSL_Select = gSL_Select & " Where 仕入先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & "   and 背番号 = '" & RTrim(HTM.背番号) & "'"
    gSL_Select = gSL_Select & " Order By 年月 Desc "
    gLong = 0
    If KKTRead(gSL_Select, , 1) Then
       With vsGrid3
           Do Until Not KKT_RDSTS
              For gLong = 1 To .Rows - 1
                  If Trim(.TextMatrix(gLong, 9)) = Trim(KKT.年月) Then
                    '回転枚数
                    .TextMatrix(gLong, 4) = CisFun.RSetFld(KKT.回転枚数, 3, "##0") & _
                                          "(" & CisFun.RSetFld(KKT.前回回転枚数, 3, "##0") & ")"
                    '前月差
                    .TextMatrix(gLong, 5) = Format(KKT.回転枚数 - KKT.前月回転枚数, "##0")
                    '平準化日
                    .TextMatrix(gLong, 6) = Format(KKT.日上限枚数, "###")
                    '平準化便
                    .TextMatrix(gLong, 7) = Format(KKT.便上限枚数, "###")
                    Exit For
                   End If
              Next gLong
              Call KKTReadNext(1)
          Loop
       End With
    End If
    Call KKTClose(1)
    '------------------------------------------------------- 2004/12/03 Insert End

'   【かんばん増減管理】
    Call GridSet1

'   【かんばん増減管理】
    Call GridSet2

    BodySet = True
End Function
'+-------------------------------------+
'+      ｸﾞﾘｯﾄﾞより内容セット後入力処理
'+-------------------------------------+
Private Sub FromGrid1()
    If SyoriKB = "MNT" Or _
       SyoriKB = "DEL" Then
        Call BodySet1
        If Trim(vsGrid1.TextMatrix(vsGrid1.Row, CisVsGrid1.FixedGet(0))) = "" Then
            vsGrid1.SetFocus
            Exit Sub
        End If
        If SyoriKB = "MNT" Then
            ProcHB = "B1"
            Call DispChange(ProcHB)
            Call CisVsGrid1.vsColor(消去個別)
            Call CisVsGrid1.vsColor
            B1Chk_SKbn.SetFocus
        End If
        If SyoriKB = "DEL" Then
            Call CisVsGrid1.vsColor
            If Not DBDelete Then
                Call CisVsGrid1.vsColor(消去個別)
            Else
                If Not GridSet1 Then
                '   表示明細が無い時は、新規にする
                    SyoriKB = "ADD"
                    ProcHB = "B1"
                    Call DispChange(ProcHB)
                    Call SyoriName(SyoriKB)
                    B1Chk_SKbn.SetFocus
                    Exit Sub
                End If
                Call BodySet1
            End If
            vsGrid1.SetFocus
        End If
    End If
End Sub
'+---------------------------+
'+      ｸﾞﾘｯﾄﾞより内容セット
'+---------------------------+
Private Sub BodySet1()
    With vsGrid1
        If .Row < 1 Then Exit Sub
        Select Case .TextMatrix(.Row, CisVsGrid1.FixedGet(1))
            Case "有効": B1Chk_SKbn = 0
            Case "無効": B1Chk_SKbn = 1
        End Select
        Select Case .TextMatrix(.Row, CisVsGrid1.FixedGet(2))
            Case "減": B1Op_ZGKbn(0) = True
            Case "増": B1Op_ZGKbn(1) = True
        End Select
        iB1_SKYmd = Format(.TextMatrix(.Row, CisVsGrid1.FixedGet(3)), "yyyymmdd")
        iB1_EKYmd = Format(.TextMatrix(.Row, CisVsGrid1.FixedGet(4)), "yyyymmdd")
        iB1_SetSu = .TextMatrix(.Row, CisVsGrid1.FixedGet(5))
        iB1_BinSu = .TextMatrix(.Row, CisVsGrid1.FixedGet(6))
        B1lb_RuiSu = .TextMatrix(.Row, CisVsGrid1.FixedGet(7))
        B1lb_KanYmd = .TextMatrix(.Row, CisVsGrid1.FixedGet(8))
    End With
'   退避
    Old_SKYmd = iB1_SKYmd
End Sub
'+-------------------------------------+
'+      ｸﾞﾘｯﾄﾞより内容セット後入力処理
'+-------------------------------------+
Private Sub FromGrid2()
    If SyoriKB = "MNT" Or _
       SyoriKB = "DEL" Then
        Call BodySet2
        If Trim(vsGrid2.TextMatrix(vsGrid2.Row, CisVsGrid2.FixedGet(0))) = "" Then
            vsGrid2.SetFocus
            Exit Sub
        End If
        If SyoriKB = "MNT" Then
            ProcHB = "B2"
            Call DispChange(ProcHB)
            Call CisVsGrid2.vsColor(消去個別)
            Call CisVsGrid2.vsColor
            B2Chk_SKbn.SetFocus
        End If
        If SyoriKB = "DEL" Then
            Call CisVsGrid2.vsColor
            If Not DBDelete Then
                Call CisVsGrid2.vsColor(消去個別)
            Else
                If Not GridSet2 Then
                '   表示明細が無い時は、新規にする
                    SyoriKB = "ADD"
                    ProcHB = "B2"
                    Call DispChange(ProcHB)
                    Call SyoriName(SyoriKB)
                    B2Chk_SKbn.SetFocus
                    Exit Sub
                End If
                Call BodySet2
            End If
            vsGrid2.SetFocus
        End If
    End If
End Sub
'+---------------------------+
'+      ｸﾞﾘｯﾄﾞより内容セット
'+---------------------------+
Private Sub BodySet2()
    With vsGrid2
        If .Row < 1 Then Exit Sub
        Select Case .TextMatrix(.Row, CisVsGrid2.FixedGet(1))
            Case "有効": B2Chk_SKbn = 0
            Case "無効": B2Chk_SKbn = 1
        End Select
        iB2_SKYmd = Format(.TextMatrix(.Row, CisVsGrid2.FixedGet(2)), "yyyymmdd")
        iB2_EKYmd = Format(.TextMatrix(.Row, CisVsGrid2.FixedGet(3)), "yyyymmdd")
        iB2_BinSu = .TextMatrix(.Row, CisVsGrid2.FixedGet(4))
        iB2_DaySu = .TextMatrix(.Row, CisVsGrid2.FixedGet(5))
        iB2_MonSu = .TextMatrix(.Row, CisVsGrid2.FixedGet(6))
    End With
'   退避
    Old_SKYmd = iB2_SKYmd
End Sub
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+      削 除 処 理      +
'+----------------------+
Private Function DBDelete()
    DBDelete = False

'  ( 確認メッセージ )
    K_Inkbnm.Caption = InputNM
    K_Inkbnm.Visible = True
    K_Sykbnm(InputNo).Caption = SyoriNM
    K_Sykbnm(InputNo).Visible = True
    With CisFun
         .MB_Lines = 5
         .MB_MSG(2) = "    削除処理を行います。            "
         .MB_MSG(4) = "     よろしいですか？            "
         .MB_Title = "削除確認"
         .MB_Button = OK_CAN
         If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBDelete_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 削除処理 )
    Select Case InputNo
        Case 1: Call DBSet_KZT: Call DBDelete_KZT
        Case 2: Call DBSet_KJT: Call DBDelete_KJT
    End Select

    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    DBDelete = True
    Exit Function
DBDelete_Err:
    CisFun.ErrorBox
    End
End Function
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False

'  ( 確認メッセージ )
    K_Inkbnm.Caption = InputNM
    K_Inkbnm.Visible = True
    K_Sykbnm(InputNo).Caption = SyoriNM
    K_Sykbnm(InputNo).Visible = True
    With CisFun
        .MB_Lines = 5
        If SyoriKB = "ADD" Then
            .MB_MSG(2) = "    追加処理を行います。            "
            .MB_Title = "追加確認"
        Else
            .MB_MSG(2) = "    修正処理を行います。            "
            .MB_Title = "修正確認"
        End If
        .MB_MSG(4) = "     よろしいですか？            "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo DBPut_Err

    Call CisDB.DBTran(TransBegin)   ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 更新処理 )
    If SyoriKB = "ADD" Then
        Select Case InputNo
            Case 1: Call DBSet_KZT: Call KZTInsert
            Case 2: Call DBSet_KJT: Call KJTInsert
            Case 3: Call DBSet_KHT: Call DBInsert_KHT
        End Select
    Else
        Select Case InputNo
            Case 1: Call DBSet_KZT: Call DBUpdate_KZT
            Case 2: Call DBSet_KJT: Call DBUpdate_KJT
        End Select
    End If

    CisDB.DBTran (TransCommit)      ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了

    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+------------------------------------+
'+      かんばん増減テーブル・セット
'+------------------------------------+
Private Function DBSet_KZT()
    DBSet_KZT = False
    Call ItemsClearKZT
    With KZT
        .取引先 = RTrim(iH1_Torcd)
        .受入 = RTrim(iH1_Ukeir)
        .品番 = RTrim(iH1_Hinbn)
        .開始日 = RTrim(iB1_SKYmd)
        .終了日 = RTrim(iB1_EKYmd)
        .設定内容区分 = B1Chk_SKbn
        If B1Op_ZGKbn(0) Then .増減区分 = 0
        If B1Op_ZGKbn(1) Then .増減区分 = 1
        .設定枚数 = CisFun.Val2(iB1_SetSu)
        .便当枚数 = CisFun.Val2(iB1_BinSu)
        .作成者 = gTanto
        .更新者 = gTanto
        .累積枚数 = CisFun.Val2(B1lb_RuiSu)
    End With
    DBSet_KZT = True
End Function
'*--------------------*'
'*     更新処理
'*--------------------*'
Public Function DBUpdate_KZT(Optional DBNo As Integer) As Boolean
    DBUpdate_KZT = False
    With KZT
        gSL_Select = ""
        gSL_Select = gSL_Select & "update かんばん増減テーブル"
        gSL_Select = gSL_Select & "   set 開始日 = '" & .開始日 & "'"
        gSL_Select = gSL_Select & "     , 終了日 = '" & .終了日 & "'"
        gSL_Select = gSL_Select & "     , 設定内容区分 = " & .設定内容区分 & ""
        gSL_Select = gSL_Select & "     , 増減区分 = " & .増減区分 & ""
        gSL_Select = gSL_Select & "     , 設定枚数 = " & .設定枚数 & ""
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
        gSL_Select = gSL_Select & "   and 開始日 = '" & RTrim(Old_SKYmd) & "'"
    End With
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBUpdate_KZT = True
End Function
'*--------------------*'
'*     削除処理
'*--------------------*'
Public Function DBDelete_KZT(Optional DBNo As Integer) As Boolean
    DBDelete_KZT = False
    With KZT
        gSL_Select = ""
        gSL_Select = gSL_Select & "delete かんばん増減テーブル"
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
    DBDelete_KZT = True
End Function
'+------------------------------------+
'+      かんばん平準化テーブル・セット
'+------------------------------------+
Private Function DBSet_KJT()
    DBSet_KJT = False
    Call ItemsClearKJT
    With KJT
        .取引先 = RTrim(iH1_Torcd)
        .受入 = RTrim(iH1_Ukeir)
        .品番 = RTrim(iH1_Hinbn)
        .開始日 = RTrim(iB2_SKYmd)
        .終了日 = RTrim(iB2_EKYmd)
        .設定内容区分 = B2Chk_SKbn
        .便上限枚数 = CisFun.Val2(iB2_BinSu)
        .日上限枚数 = CisFun.Val2(iB2_DaySu)
        .月上限枚数 = CisFun.Val2(iB2_MonSu)
        If .日上限枚数 <> 0 Then .平準化区分 = 1
        If .便上限枚数 <> 0 Then .平準化区分 = 2
        If .月上限枚数 <> 0 Then .平準化区分 = 3
        .作成者 = gTanto
        .更新者 = gTanto
    End With
    DBSet_KJT = True
End Function
'*--------------------*'
'*     更新処理
'*--------------------*'
Public Function DBUpdate_KJT(Optional DBNo As Integer) As Boolean
    DBUpdate_KJT = False
    With KJT
        gSL_Select = ""
        gSL_Select = gSL_Select & "update かんばん平準化テーブル"
        gSL_Select = gSL_Select & "   set 開始日 = '" & .開始日 & "'"
        gSL_Select = gSL_Select & "     , 終了日 = '" & .終了日 & "'"
        gSL_Select = gSL_Select & "     , 設定内容区分 = " & .設定内容区分 & ""
        gSL_Select = gSL_Select & "     , 平準化区分 = " & .平準化区分 & ""
        gSL_Select = gSL_Select & "     , 便上限枚数 = " & .便上限枚数 & ""
        gSL_Select = gSL_Select & "     , 日上限枚数 = " & .日上限枚数 & ""
        gSL_Select = gSL_Select & "     , 月上限枚数 = " & .月上限枚数 & ""
        gSL_Select = gSL_Select & "     , 更新日 = getdate()"
        gSL_Select = gSL_Select & "     , 更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & "     , 更新端末 = substring(host_name(),1,20)"
        gSL_Select = gSL_Select & " where 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(.受入) & "'"
        gSL_Select = gSL_Select & "   and 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   and 開始日 = '" & RTrim(Old_SKYmd) & "'"
    End With
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBUpdate_KJT = True
End Function
'*--------------------*'
'*     削除処理
'*--------------------*'
Public Function DBDelete_KJT(Optional DBNo As Integer) As Boolean
    DBDelete_KJT = False
    With KJT
        gSL_Select = ""
        gSL_Select = gSL_Select & "delete かんばん平準化テーブル"
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
    DBDelete_KJT = True
End Function
'+------------------------------------+
'+      かんばん発行テーブル・セット
'+------------------------------------+
Private Function DBSet_KHT()
    Dim wTSeban(4)  As String
    Dim wTorCD(4)   As String
    Dim wTorKJ(4)   As String
    
    DBSet_KHT = False

    Call ItemsClearKHT
    With KHT
        .品番 = RTrim(iH1_Hinbn)
        .取引先 = RTrim(iH1_Torcd)
        .受入 = RTrim(iH1_Ukeir)
        .自社CD = Kn_Jisya
        .自社名1 = CisFun.Mid2(Kn_JisyaNM, 1, 20)
        .自社名2 = CisFun.Mid2(Kn_JisyaKJNM, 1, 14)
        .作成者 = gTanto
    End With
'   品番取引先情報(得意先情報)
    gSL_Select = "Select *"
    gSL_Select = gSL_Select & " From 品番取引先マスタ "
    gSL_Select = gSL_Select & " Where 品番 = '" & RTrim(KHT.品番) & "'"
    gSL_Select = gSL_Select & " And   品目 = 0"
    gSL_Select = gSL_Select & " And   SK区分 = 0"
    gSL_Select = gSL_Select & " Order By 背番号 "
    If HTMRead(gSL_Select) Then
       gCnt1 = 1
       Do Until Not HTM_RDSTS
           If Trim(HTM.背番号) <> "" Then
              wTSeban(gCnt1) = Mid(Trim(HTM.背番号), 1, 3)
              wTorCD(gCnt1) = HTM.取引先
              gCnt1 = gCnt1 + 1
              If gCnt1 > 4 Then Exit Do
           End If
           Call HTMReadNext
       Loop
       For gCnt1 = 1 To 4
           If Trim(wTorCD(gCnt1)) <> "" Then
              gSL_Select = "Select * From 取引先マスタ "
              gSL_Select = gSL_Select & " Where 取引先区分 = 0"
              gSL_Select = gSL_Select & "   AND 取引先CD = '" & wTorCD(gCnt1) & "'"
              If TRMRead(gSL_Select, 1) Then
                 wTorKJ(gCnt1) = CisFun.Mid2(TRM.工場名, 1, 6)
              End If
           End If
           Select Case gCnt1
                Case 1
                    KHT.得意先背番1 = wTSeban(gCnt1)
                    KHT.得意先工場名1 = wTorKJ(gCnt1)
                Case 2
                    KHT.得意先背番2 = wTSeban(gCnt1)
                    KHT.得意先工場名2 = wTorKJ(gCnt1)
                Case 3
                    KHT.得意先背番3 = wTSeban(gCnt1)
                    KHT.得意先工場名3 = wTorKJ(gCnt1)
                Case 4
                    KHT.得意先背番4 = wTSeban(gCnt1)
                    KHT.得意先工場名4 = wTorKJ(gCnt1)
           End Select
       Next gCnt1
    End If
    Call HTMClose
'>> 品番取引先マスタ
'< 2005.03.01 MNT START >--------------------------------------------------------------------
'
'    gSL_Select = ""
'    gSL_Select = gSL_Select & "select htm.背番号"
'    gSL_Select = gSL_Select & "     , htm.品目"
'    gSL_Select = gSL_Select & "     , htm.手配区分"
'    gSL_Select = gSL_Select & "     , htm.納入先"
'    gSL_Select = gSL_Select & "     , htm.納入先受入"
'    gSL_Select = gSL_Select & "     , htm.サイクル"
'    gSL_Select = gSL_Select & "     , htm.置場"
'    gSL_Select = gSL_Select & "     , htm.基準工数"
'    gSL_Select = gSL_Select & "     , htm.最終発行連番"
'    gSL_Select = gSL_Select & "     , htm.アドレス"
'    gSL_Select = gSL_Select & "     , htm.備考"
'    gSL_Select = gSL_Select & "     , isnull(sym.値名称, '') 手配区分名"
'    gSL_Select = gSL_Select & "  from 品番取引先マスタ htm"
'    gSL_Select = gSL_Select & "  left outer join 名称マスタ sym"
'    gSL_Select = gSL_Select & "    on sym.区分名称 = '手配区分'"
'    gSL_Select = gSL_Select & "   and sym.区分タイプ = ''"
'    gSL_Select = gSL_Select & "   and sym.値 = htm.手配区分"
'    gSL_Select = gSL_Select & " where htm.品番 = '" & KHT.品番 & "'"
'    gSL_Select = gSL_Select & "   and htm.取引先 = '" & KHT.取引先 & "'"
'    gSL_Select = gSL_Select & "   and htm.受入 = '" & KHT.受入 & "'"

    gSL_Select = ""
    gSL_Select = gSL_Select & "select htm.背番号"
    gSL_Select = gSL_Select & "     , htm.品目"
    gSL_Select = gSL_Select & "     , htm.手配区分"
    gSL_Select = gSL_Select & "     , htm.納入先"
    gSL_Select = gSL_Select & "     , htm.納入先受入"
    gSL_Select = gSL_Select & "     , htm.サイクル"
    gSL_Select = gSL_Select & "     , htm.置場"
    gSL_Select = gSL_Select & "     , htm.基準工数"
    
    gSL_Select = gSL_Select & "     , Case When htm.最終発行連番 > KN.発行NO"
    gSL_Select = gSL_Select & "       Then htm.最終発行連番 Else KN.発行NO  End 最終発行連番"
    
    gSL_Select = gSL_Select & "     , htm.アドレス"
    gSL_Select = gSL_Select & "     , htm.備考"
    gSL_Select = gSL_Select & "     , isnull(sym.値名称, '') 手配区分名"
    gSL_Select = gSL_Select & "  from 品番取引先マスタ htm"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 取引先, 受入, 品番, MAX(発行NO) 発行NO"
    gSL_Select = gSL_Select & " From かんばん流動管理テーブル"
    gSL_Select = gSL_Select & " GROUP BY 取引先, 受入, 品番 ) KN"
    gSL_Select = gSL_Select & " ON  KN.取引先 = htm.取引先"
    gSL_Select = gSL_Select & " AND KN.受入 = htm.受入"
    gSL_Select = gSL_Select & " AND KN.品番 = htm.品番"
    gSL_Select = gSL_Select & "  left outer join 名称マスタ sym"
    gSL_Select = gSL_Select & "    on sym.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "   and sym.区分タイプ = ''"
    gSL_Select = gSL_Select & "   and sym.値 = htm.手配区分"
    gSL_Select = gSL_Select & " where htm.品番 = '" & KHT.品番 & "'"
    gSL_Select = gSL_Select & "   and htm.取引先 = '" & KHT.取引先 & "'"
    gSL_Select = gSL_Select & "   and htm.受入 = '" & KHT.受入 & "'"
'
'< 2005.03.01 MNT START >--------------------------------------------------------------------
 
    If HTMRead(gSL_Select, 1) Then
        KHT.背番号 = HTM.背番号
        KHT.品目 = HTM.品目
        KHT.手配区分 = HTM.手配区分
        KHT.手配区分名 = CisFun.Mid2(HTM.手配区分名, 1, 4)
        KHT.納入先 = HTM.納入先
        KHT.納入先受入 = HTM.納入先受入
        If Trim(HTM.サイクル) <> "" Then
            KHT.サイクル = CisFun.RSetFld(CisFun.Mid2(HTM.サイクル, 1, 1), 1, "#") _
                   & "-" & CisFun.RSetFld(CisFun.Mid2(HTM.サイクル, 2, 2), 2, "#") _
                   & "-" & CisFun.RSetFld(CisFun.Mid2(HTM.サイクル, 4, 2), 2, "#")
        End If
        KHT.置場 = HTM.置場
        If KHT.手配区分 = 2 Then
        '   << 社内の時 >>
            KHT.基準工数 = HTM.基準工数
        End If
        New_EndNo = HTM.最終発行連番
        KHT.アドレス = HTM.アドレス
        KHT.備考 = HTM.備考
    End If
'>> 品番マスタ
'   ■■■■■■■■■■　品番マスタより管理番号がなくなった？対応が必要！　■■■■■■■■■■
    gSL_Select = ""
    gSL_Select = gSL_Select & "select hnm.表示品番"
    gSL_Select = gSL_Select & "     , hnm.品名"
    gSL_Select = gSL_Select & "     , hnm.工程"
    gSL_Select = gSL_Select & "     , hnm.次工程"
    gSL_Select = gSL_Select & "     , hnm.IMAGE"
    gSL_Select = gSL_Select & "     , hnm.実数区分"
    gSL_Select = gSL_Select & "     , hnm.車型"
    gSL_Select = gSL_Select & "     , hnm.担当"
    gSL_Select = gSL_Select & "  from 品番マスタ hnm"
    gSL_Select = gSL_Select & " where hnm.品番 = '" & KHT.品番 & "'"
    If HNMRead(gSL_Select, 1) Then
        KHT.表示品番 = HNM.表示品番
        If Len(HNM.表示品番) < 15 Then
           KHT.品番14 = HNM.表示品番
        Else
           KHT.品番30 = HNM.表示品番
        End If
        KHT.品名 = HNM.品名
        KHT.IMAGE = HNM.IMAGE
        KHT.車型 = HNM.車型
        KHT.実数区分 = HNM.実数区分
        If KHT.手配区分 = 2 Then
        '   << 社内の時 >>
            KHT.工程 = HNM.工程
            KHT.次工程 = HNM.次工程
           'KHT.管理番号 = HNM.管理番号
            If KHT.工程 <> "" Then
                If MeisyoGet("工程管理", KHT.工程) Then KHT.工程名 = SYM_Meisyo
            End If
            If KHT.次工程 <> "" Then
                If MeisyoGet("工程管理", KHT.次工程) Then KHT.次工程名 = SYM_Meisyo
            End If
        End If
        If mRType = 2 Then
           ' 担当者名
           Call MeisyoGet("担当管理", HNM.担当)
           KHT.区分 = CisFun.Mid2(SYM_Meisyo, 1, 4)
        End If
    End If
'>> 取引先マスタ（取引先）
    gSL_Select = ""
    gSL_Select = gSL_Select & "select trm.略称"
    gSL_Select = gSL_Select & "     , trm.工場名"
    gSL_Select = gSL_Select & "     , trm.取引先"
    gSL_Select = gSL_Select & "     , trm.工場CD"
    gSL_Select = gSL_Select & "  from 取引先マスタ trm"
    gSL_Select = gSL_Select & " where trm.取引先CD = '" & KHT.取引先 & "'"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    If TRMRead(gSL_Select, 1) Then
        KHT.取引先CD = TRM.取引先
        If Trim(TRM.工場CD) <> "" Then
           KHT.取引先CD = TRM.取引先 & "-" & TRM.工場CD
        End If
        KHT.取引先名 = TRM.略称
        KHT.取引先名1 = CisFun.Mid2(TRM.略称, 1, 10)
        KHT.取引先名2 = CisFun.Mid2(TRM.略称, 11, 10)
        KHT.取引先工場名 = TRM.工場名
        KHT.取引先工場名1 = CisFun.Mid2(TRM.工場名, 1, 10)
        KHT.取引先工場名2 = CisFun.Mid2(TRM.工場名, 11, 10)
    End If
    ' 取引先受入情報
    gSL_Select = "Select * From 取引先受入マスタ "
    gSL_Select = gSL_Select & " Where 取引先区分 = 1"
    gSL_Select = gSL_Select & "   AND 取引先CD = '" & KHT.取引先 & "'"
    gSL_Select = gSL_Select & "   AND 受入 = '" & KHT.受入 & "'"
    If TUMRead(gSL_Select, 1) Then
       KHT.受入名 = TUM.受入名称
    End If
'>> 取引先マスタ（納入先）
    gSL_Select = ""
    gSL_Select = gSL_Select & "select trm.略称,trm.取引先,trm.工場CD"
    gSL_Select = gSL_Select & "  from 取引先マスタ trm"
    gSL_Select = gSL_Select & " where trm.取引先CD = '" & KHT.納入先 & "'"
    gSL_Select = gSL_Select & "   and trm.取引先区分 = 1"
    If TRMRead(gSL_Select, 1) Then
        KHT.納入先CD = TRM.取引先
        If Trim(TRM.工場CD) <> "" Then
           KHT.納入先CD = TRM.取引先 & "-" & TRM.工場CD
        End If
        KHT.納入先名 = TRM.略称
        KHT.納入先名1 = CisFun.Mid2(TRM.略称, 1, 10)
        KHT.納入先名2 = CisFun.Mid2(TRM.略称, 11, 10)
    End If
    If KHT.手配区分 = "1" Then
       KHT.納入先固定 = "直送先"
    End If
    ' 納入時間情報
    gSL_Select = "Select * From 発注サイクルマスタ"
    gSL_Select = gSL_Select & " Where 仕入先 = '" & RTrim(KHT.取引先) & "'"
    gSL_Select = gSL_Select & " And   サイクル = '" & RTrim(HTM.サイクル) & "'"
    gSL_Select = gSL_Select & " And   受入 = '" & RTrim(KHT.受入) & "'"
    gSL_Select = gSL_Select & " And   行識別 = 0"
    If HCMRead(gSL_Select) Then
       gCnt1 = 1
       Do Until Not HCM_RDSTS
           If Trim(HCM.入荷時間) <> "" And HCM.入荷時間 <> "0000" Then
              gStr = CisFun.RSetFld(Mid(HCM.入荷時間, 1, 2), 2, "##") & ":" & Mid(HCM.入荷時間, 3, 2)
              Select Case gCnt1
                    Case 1: KHT.納入時間1 = gStr
                    Case 2: KHT.納入時間2 = gStr
                    Case 3: KHT.納入時間3 = gStr
                    Case 4: KHT.納入時間4 = gStr
                    Case 5: KHT.納入時間5 = gStr
                    Case 6: KHT.納入時間6 = gStr
                    Case 7: KHT.納入時間7 = gStr
                    Case 8: KHT.納入時間8 = gStr
              End Select
              gCnt1 = gCnt1 + 1
              If gCnt1 > 8 Then Exit Do
           End If
           Call HCMReadNext
       Loop
    End If
    Call HCMClose
    ' 発行タイプ別
    If mRType = 1 Then
       gStr = Format(CisFun.Val2(KHT.アドレス), "#")
       KHT.アドレス = Space(6 - CisFun.Len2(gStr)) & gStr
       Select Case HTM.手配区分
          Case "1": KHT.区分 = "直送"
          Case "3": KHT.区分 = "自給"
          Case Else
            KHT.区分 = "自給"
       End Select
       If Mid(KHT.取引先, 1, 4) < "7000" Then
          gSL_Select = "Select * From 名称マスタ"
          gSL_Select = gSL_Select & " Where 区分名称 = 'かんばん発行ホッパ指定'"
          gSL_Select = gSL_Select & " And   区分タイプ = ''"
          gSL_Select = gSL_Select & " And   数字3 = 1"
          gSL_Select = gSL_Select & " Order By 値"
          If Not SYMRead(gSL_Select, 1) Then
             SYM.数字1 = 1: SYM.数字2 = 3
          End If
          KHT.発行種類 = 10
          KHT.StHNo = SYM.数字1 - 1
          KHT.EdHNo = SYM.数字2 - 1
          KHT.Stack1 = KHT.StHNo
          KHT.Stack2 = KHT.StHNo
       Else
          gSL_Select = "Select * From 名称マスタ"
          gSL_Select = gSL_Select & " Where 区分名称 = 'かんばん発行ホッパ指定'"
          gSL_Select = gSL_Select & " And   区分タイプ = ''"
          gSL_Select = gSL_Select & " And   数字3 = 2"
          gSL_Select = gSL_Select & " Order By 値"
          If Not SYMRead(gSL_Select, 1) Then
             SYM.数字1 = 4: SYM.数字2 = 4
          End If
          KHT.発行種類 = 11
          KHT.StHNo = SYM.数字1 - 1
          KHT.EdHNo = SYM.数字2 - 1
          KHT.Stack1 = KHT.StHNo
          KHT.Stack2 = KHT.StHNo
       End If
    End If

    With KHT
        .取引先名 = RepName(.取引先名, 20)
        .取引先名1 = RepName(.取引先名1, 10)
        .取引先名2 = RepName(.取引先名2, 10)
        .取引先工場名 = RepName(.取引先工場名, 20)
        .取引先工場名1 = RepName(.取引先工場名1, 10)
        .取引先工場名2 = RepName(.取引先工場名1, 10)
        .納入先名 = RepName(.納入先名, 20)
        .納入先名1 = RepName(.納入先名1, 10)
        .納入先名2 = RepName(.納入先名2, 10)
    End With

    DBSet_KHT = True
End Function
'+------------------------------+
'+  特殊文字置換え
'+------------------------------+
Private Function RepName(InName As String, Leng As Long) As String
    
    If CisFun.Len2(InName) + 2 <= Leng Then
       InName = Replace(InName, "㈱", "(株)")
       InName = Replace(InName, "㈲", "(有)")
    Else
       InName = Replace(InName, "㈱", "株")
       InName = Replace(InName, "㈲", "有")
    End If
    RepName = InName
End Function
'*--------------------*'
'*     追加処理
'*--------------------*'
Private Function DBInsert_KHT(Optional DBNo As Integer) As Boolean
    Dim wIndex      As Integer
    Dim wCount      As Integer
    Dim wMaisu      As Integer

    DBInsert_KHT = False

'   <><> かんばん発行テーブルの削除：発行済
    Call DBDelete_KHT
'   <><> かんばん発行テーブルの追加：追加枚数分
    For wIndex = 1 To 2
        wMaisu = CisFun.Val2(iB3_HMaisu(wIndex))
        For wCount = 1 To wMaisu
            New_EndNo = New_EndNo + 1
            With KHT
                .収容区分 = wIndex
                .発行NO = New_EndNo
                .収容器 = RTrim(B3lb_Syuki(wIndex))
                .収容数 = CisFun.Val2(B3lb_Syuyo(wIndex))
            End With
            Call KHTInsert
        Next wCount
        iT1_HMaisu = CisFun.Val2(iT1_HMaisu) + wMaisu
    Next wIndex
'   <><> 品番取引先マスタの更新：最終発行連番
    Call DBUpdate_HTM

    DBInsert_KHT = True
End Function
'*--------------------*'
'*     削除処理
'*--------------------*'
Private Function DBDelete_KHT(Optional DBNo As Integer) As Boolean
    DBDelete_KHT = False
    gSL_Select = ""
    gSL_Select = gSL_Select & "delete かんばん発行テーブル"
    gSL_Select = gSL_Select & " where かんばん種類 = " & KHT.かんばん種類 & ""
    gSL_Select = gSL_Select & "   and 発行済 = 1"
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBDelete_KHT = True
End Function
'*--------------------*'
'*     更新処理
'*--------------------*'
Public Function DBUpdate_HTM(Optional DBNo As Integer) As Boolean
    DBUpdate_HTM = False
    With HTM
        .品番 = KHT.品番
        .取引先 = KHT.取引先
        .受入 = KHT.受入
        .最終発行連番 = New_EndNo
        .作成者 = gTanto

        gSL_Select = ""
        gSL_Select = gSL_Select & "update 品番取引先マスタ"
        gSL_Select = gSL_Select & "   set 最終発行連番 = '" & .最終発行連番 & "'"
        gSL_Select = gSL_Select & "     , 更新日 = getdate()"
        gSL_Select = gSL_Select & "     , 更新者 = '" & RTrim(.更新者) & "'"
        gSL_Select = gSL_Select & "     , 更新端末 = substring(host_name(),1,20)"
        gSL_Select = gSL_Select & " where 品番 = '" & RTrim(.品番) & "'"
        gSL_Select = gSL_Select & "   and 取引先 = '" & RTrim(.取引先) & "'"
        gSL_Select = gSL_Select & "   and 受入 = '" & RTrim(.受入) & "'"
'
'< 2005.03.01 ADD START >--------------------------------------------------------------------
        gSL_Select = gSL_Select & "   and 最終発行連番 < '" & .最終発行連番 & "'"
'< 2005.03.01 ADD  END  >--------------------------------------------------------------------
    
    End With
    With CisDB
        .ConnectNo = DBNo
        .SQL = gSL_Select
        If Not .DBExec Then Exit Function
    End With
    DBUpdate_HTM = True
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
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - かんばん増減管理
'+----------------------------------------------------------+
Private Sub GridInit1()
    With CisVsGrid1
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "№")
        Call .FixedSet(中中, 中中, あり, 5, 1, "設定")
        Call .FixedSet(中中, 中中, あり, 8, 2, "増減区分")
        Call .FixedSet(中中, 中中, あり, 11, 3, "開 始 日")
        Call .FixedSet(中中, 中中, あり, 11, 4, "終 了 日")
        Call .FixedSet(中中, 右中, あり, 9, 5, "設定枚数")
        Call .FixedSet(中中, 右中, あり, 9, 6, "便当枚数")
        Call .FixedSet(中中, 右中, あり, 9, 7, "累積枚数")
        Call .FixedSet(中中, 中中, あり, 10, 8, "完 了 日")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 9, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8")
        .InitDisp
    End With
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - かんばん平準化管理
'+----------------------------------------------------------+
Private Sub GridInit2()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "№")
        Call .FixedSet(中中, 中中, あり, 5, 1, "設定")
        Call .FixedSet(中中, 中中, あり, 11, 2, "開 始 日")
        Call .FixedSet(中中, 中中, あり, 11, 3, "終 了 日")
        Call .FixedSet(中中, 右中, あり, 8, 4, "便当枚数")
        Call .FixedSet(中中, 右中, あり, 8, 5, "日当枚数")
        Call .FixedSet(中中, 右中, あり, 8, 6, "月当枚数")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 7, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
        .InitDisp
    End With
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ - 内示情報
'+----------------------------------------------------------+
Private Sub GridInit3()
    With CisVsGrid3
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 中中, あり, 8, 1, "内示年月")
        Call .FixedSet(中中, 右中, あり, 9, 2, "内示数")
        Call .FixedSet(中中, 右中, あり, 9, 3, "日当枚数")
        Call .FixedSet(中中, 右中, あり, 9, 4, "回転枚数")
        Call .FixedSet(中中, 右中, あり, 6, 5, "前月差")
        Call .FixedSet(中中, 右中, あり, 9, 6, "平準化日")
        Call .FixedSet(中中, 右中, あり, 9, 7, "平準化便")
        ' +----------------------------------------------------- 2004/12/03 Insert
        Call .FixedSet(中中, 右中, あり, 0, 8, "稼動日数")
        Call .FixedSet(中中, 右中, あり, 0, 9, "年月")
        ' +----------------------------------------------------- 2004/12/03 Insert End
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 8, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7")
        .InitDisp
    End With
End Sub
'****************************************************'
'*****      グリッド内容セット - かんばん増減管理
'****************************************************'
Private Function GridSet1() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet1 = False
    Call CisVsGrid1.Clear
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  kzt.設定内容区分"
    gSL_Select = gSL_Select & ", kzt.増減区分"
    gSL_Select = gSL_Select & ", kzt.開始日"
    gSL_Select = gSL_Select & ", kzt.終了日"
    gSL_Select = gSL_Select & ", kzt.設定枚数"
    gSL_Select = gSL_Select & ", kzt.便当枚数"
    gSL_Select = gSL_Select & ", kzt.累積枚数"
    gSL_Select = gSL_Select & ", kzt.完了日"
    gSL_Select = gSL_Select & " from かんばん増減テーブル kzt"
    gSL_Select = gSL_Select & " where kzt.取引先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and kzt.受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & "   and kzt.品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  kzt.開始日 desc"
    If KZTRead(gSL_Select, , 1) Then

'       抽出処理 初期ｾｯﾄ
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
        '設定
        Select Case KZT.設定内容区分
            Case 0: .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = "有効"
            Case 1: .TextMatrix(SetRow, CisVsGrid1.FixedGet(1)) = "無効"
        End Select
        '増減区分
        Select Case KZT.増減区分
            Case 0: .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = "減"
            Case 1: .TextMatrix(SetRow, CisVsGrid1.FixedGet(2)) = "増"
        End Select
        '開始日
        If KZT.開始日 <> "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(3)) = Format(KZT.開始日, "0000/00/00")
        End If
        '終了日
        If KZT.終了日 <> "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(4)) = Format(KZT.終了日, "0000/00/00")
        End If
        '設定枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(5)) = Format(KZT.設定枚数, "#,###")
        '便当枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(6)) = Format(KZT.便当枚数, "#,###")
        '累積枚数
        .TextMatrix(SetRow, CisVsGrid1.FixedGet(7)) = Format(KZT.累積枚数, "#,###")
        '完了日
        If KZT.完了日 <> "" Then
            .TextMatrix(SetRow, CisVsGrid1.FixedGet(8)) = Format(KZT.完了日, "0000/00/00")
        End If
    End With
End Sub
'******************************************'
'*****      行選択（クリック）
'******************************************'
Private Sub vsGrid1_Click()
    If ProcHB = "T1" Then
        Call FromGrid1
        Exit Sub
    Else
        SendKeys "{TAB}"
    End If
End Sub
'******************************************'
'*****      行選択（スペース）
'******************************************'
Private Sub vsGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If ProcHB = "T1" Then
        If KeyCode = vbKeySpace Then
            Call FromGrid1
            Exit Sub
        End If
    Else
        SendKeys "{TAB}"
    End If
End Sub
'+-------------------------------+
'+      行変更時　内容表示
'+-------------------------------+
Private Sub VsGrid1_RowColChange()
    If ProcHB = "T1" Then
        Call BodySet1
    End If
End Sub
'****************************************************'
'*****      グリッド内容セット - かんばん平準化管理
'****************************************************'
Private Function GridSet2() As Boolean
    Dim ID          As Long
    Dim Ttl_Kensu   As Long

    GridSet2 = False
    Call CisVsGrid2.Clear
    gSL_Select = ""
    gSL_Select = gSL_Select & "select"
    gSL_Select = gSL_Select & "  kjt.設定内容区分"
    gSL_Select = gSL_Select & ", kjt.開始日"
    gSL_Select = gSL_Select & ", kjt.終了日"
    gSL_Select = gSL_Select & ", kjt.便上限枚数"
    gSL_Select = gSL_Select & ", kjt.日上限枚数"
    gSL_Select = gSL_Select & ", kjt.月上限枚数"
    gSL_Select = gSL_Select & " from かんばん平準化テーブル kjt"
    gSL_Select = gSL_Select & " where kjt.取引先 = '" & RTrim(HTM.取引先) & "'"
    gSL_Select = gSL_Select & "   and kjt.受入 = '" & RTrim(HTM.受入) & "'"
    gSL_Select = gSL_Select & "   and kjt.品番 = '" & RTrim(HTM.品番) & "'"
    gSL_Select = gSL_Select & " order by"
    gSL_Select = gSL_Select & "  kjt.開始日 desc"
    If KJTRead(gSL_Select, , 1) Then

'       抽出処理 初期ｾｯﾄ
        Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
'       ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
        With vsGrid2
           .Redraw = False
            If Ttl_Kensu + 1 < CisVsGrid2.DispMax Then
                .Rows = CisVsGrid2.DispMax
            Else
                .Rows = Ttl_Kensu + 1
            End If
        End With
        ID = 0
        Do Until Not KJT_RDSTS
            ID = ID + 1

            Call KjtInfoToGrid2(ID)

            Call KJTReadNext(1)
        Loop
        With vsGrid2
             .Row = 1: .Col = 1: .ColSel = .Cols - 1
             .Redraw = True
        End With

        GridSet2 = True
    End If
    Call KJTClose
End Function
'+----------------------------------------------+
'+  かんばん平準化テーブルよりグリッドに転送
'+----------------------------------------------+
Private Sub KjtInfoToGrid2(SetRow As Long)
    With vsGrid2
        'No
        .TextMatrix(SetRow, CisVsGrid2.FixedGet(0)) = SetRow
        '設定
        Select Case KJT.設定内容区分
            Case 0: .TextMatrix(SetRow, CisVsGrid2.FixedGet(1)) = "有効"
            Case 1: .TextMatrix(SetRow, CisVsGrid2.FixedGet(1)) = "無効"
        End Select
        '開始日
        If KJT.開始日 <> "" Then
            .TextMatrix(SetRow, CisVsGrid2.FixedGet(2)) = Format(KJT.開始日, "0000/00/00")
        End If
        '終了日
        If KJT.終了日 <> "" Then
            .TextMatrix(SetRow, CisVsGrid2.FixedGet(3)) = Format(KJT.終了日, "0000/00/00")
        End If
        '便当枚数
        .TextMatrix(SetRow, CisVsGrid2.FixedGet(4)) = Format(KJT.便上限枚数, "#,###")
        '日当枚数
        .TextMatrix(SetRow, CisVsGrid2.FixedGet(5)) = Format(KJT.日上限枚数, "#,###")
        '月当枚数
        .TextMatrix(SetRow, CisVsGrid2.FixedGet(6)) = Format(KJT.月上限枚数, "#,###")
    End With
End Sub
'******************************************'
'*****      行選択（クリック）
'******************************************'
Private Sub vsGrid2_Click()
    If ProcHB = "T2" Then
        Call FromGrid2
        Exit Sub
    Else
        SendKeys "{TAB}"
    End If
End Sub
'******************************************'
'*****      行選択（スペース）
'******************************************'
Private Sub vsGrid2_KeyDown(KeyCode As Integer, Shift As Integer)
    If ProcHB = "T2" Then
        If KeyCode = vbKeySpace Then
            Call FromGrid2
            Exit Sub
        End If
    Else
        SendKeys "{TAB}"
    End If
End Sub
'+-------------------------------+
'+      行変更時　内容表示
'+-------------------------------+
Private Sub VsGrid2_RowColChange()
    If ProcHB = "T2" Then
        Call BodySet2
    End If
End Sub
'******************************************'
'*****      行選択（クリック） - 内示情報
'******************************************'
Private Sub vsGrid3_Click()
    SendKeys "{TAB}"
End Sub

