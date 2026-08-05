VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{BA5D7819-2B67-11D4-806A-00000EA57F9E}#3.0#0"; "CISYMD~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Begin VB.Form CXT0030 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "TNS送信対象取引先設定"
   ClientHeight    =   10020
   ClientLeft      =   2895
   ClientTop       =   3465
   ClientWidth     =   13860
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   10020
   ScaleWidth      =   13860
   WindowState     =   2  '最大化
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
      Height          =   945
      Left            =   3930
      ScaleHeight     =   885
      ScaleWidth      =   3765
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1680
      Width           =   3825
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   780
         Left            =   60
         Top             =   60
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   1376
         ForeColor       =   16711680
         Caption         =   "契 約 先"
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
            Left            =   90
            TabIndex        =   0
            Top             =   330
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
            cAutoEject      =   0   'False
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1080
            Top             =   330
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
      Left            =   7905
      ScaleHeight     =   660
      ScaleWidth      =   5115
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   1680
      Width           =   5175
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   2565
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
         Left            =   3405
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
         Left            =   4245
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
         Left            =   1725
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
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   -1320
      Width           =   585
   End
   Begin VB.Timer Timer1 
      Left            =   900
      Top             =   -1380
   End
   Begin VB.Timer Timer2 
      Left            =   1170
      Top             =   -1380
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   6135
      Left            =   3930
      ScaleHeight     =   6075
      ScaleWidth      =   6885
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2820
      Width           =   6945
      Begin Cis3D_v60.CIS3D UC_3D4 
         Height          =   5940
         Index           =   1
         Left            =   90
         Top             =   75
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   10478
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
         Begin vsFlexLib.vsFlexArray vsGrid 
            Height          =   5700
            Left            =   60
            TabIndex        =   7
            TabStop         =   0   'False
            Top             =   105
            Width           =   6585
            _Version        =   196608
            _ExtentX        =   11615
            _ExtentY        =   10054
            _StockProps     =   228
            Appearance      =   1
            ConvInfo        =   1418783674
            FormatString    =   "  |仕入先　　　　　　　　　　　　|受入|切替日付        |||"
            Rows            =   16
            Cols            =   7
            BackColor       =   12648384
            BackColorFixed  =   14737632
            BackColorBkg    =   12632256
            HighLight       =   2
            RowHeightMin    =   350
            FillStyle       =   1
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Input_Row 
               Height          =   330
               Left            =   4470
               Top             =   435
               Width           =   1770
               _ExtentX        =   3122
               _ExtentY        =   582
               BackColor       =   0
               ForeColor       =   16711680
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
               cBoderColor1    =   8421504
               cBoderColor2    =   -2147483643
               cFont3DColor1   =   -2147483643
               cFont3DColor2   =   8421504
               cAlingnment     =   6
               cBoderStyle     =   1
               Begin CisYMD_V60.CisYMD iB1_YMD 
                  Height          =   330
                  Left            =   15
                  TabIndex        =   1
                  Top             =   0
                  Width           =   1740
                  _ExtentX        =   3069
                  _ExtentY        =   582
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cChkResult      =   0   'False
               End
            End
         End
         Begin CisText_V60.CisText B_TabDummy 
            Height          =   165
            Left            =   4605
            TabIndex        =   3
            Top             =   180
            Width           =   285
            _ExtentX        =   503
            _ExtentY        =   291
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
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
      Caption         =   "【 TNS送信対象取引先設定 】[自社入荷データ]"
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
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   661
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   "XXXX"
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
         Left            =   14280
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
         Left            =   11790
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
      Begin Cis3D_v60.CIS3D CountDsp 
         Height          =   315
         Left            =   9855
         Top             =   30
         Width           =   1905
         _ExtentX        =   3360
         _ExtentY        =   556
         BackColor       =   14737632
         ForeColor       =   16711935
         Caption         =   "選択件数：999件"
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
Attribute VB_Name = "CXT0030"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   TNS送信対象取引先設定
'**       フォームID    :   CXT0030
'**       処理概要      :
'**
'**       作  成  日    :   2007/07/20  By CIS SIMPR-A
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)

    Dim ActObj          As Object
    Dim ErrObj          As Object
    Dim ActRow          As Integer
    Dim Ttl_Kensu   As Long
    
'   行
    Dim NewRow                  As Long
    
    Dim mCount          As Integer      '選択件数
    
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
    
    mCount = 0
    
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
        Set .GridObj = vsGrid
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 430
        .PatanMax = 1
        .SelectCol = 19
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = ""
    Call SyoriName(SyoriKB)
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
    iB1_YMD = Format(Now(), "yyyymmdd")
    With CisFun
        
        .DateE_BefVal = iB1_YMD
        .DateEdit J_YMD
    End With

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
           'Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click  '【後退】
           Case vbKeyPageUp:
                               If PB_BACK.Visible Then
                                    Call PB_Back_Click  '【後退】
                               Else
                                    Call iB1_YMD_KeyDown(Key_Code, 0)
                                    Exit Function
                               End If
           'Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click  '【前進】
           Case vbKeyPageDown:
                               If PB_NEXT.Visible Then
                                    Call PB_Next_Click  '【前進】
                               Else
                                    Call iB1_YMD_KeyDown(Key_Code, 0)
                                    Exit Function
                               End If
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
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
   
   If ProcHB = "H1" Then
'''      SyoriKB = "MNT"
'''      Call SyoriName(SyoriKB)
      PB_CAN.Visible = False
      iH1_Torcd.SetFocus
      GoTo PB_CAN_Ed
   End If
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
'''      Call SyoriName(SyoriKB)
      iH1_Torcd.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT HT.契約先,HT.取引先,HT.受入 "
    gSL_Select = gSL_Select & " FROM 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " WHERE HT.契約先 < '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " AND   HT.品目 = 1 "
    gSL_Select = gSL_Select & " AND   HT.手配区分 = 3 "                 '------ 手配区分：受給のみ
    gSL_Select = gSL_Select & " GROUP BY HT.契約先,HT.取引先,HT.受入"
    gSL_Select = gSL_Select & " ORDER BY HT.契約先 Desc"
    If Not HTMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     先頭レコードです            "
            .MB_Title = "【後退】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Back_Ed
    End If
    
    iH1_Torcd = RTrim(HTM.契約先)
    Call iH1_Torcd_LostFocus
    
    gSL_Select = " SELECT HT.契約先,HT.取引先,HT.受入,ISNULL(TR.略称,'') 取引先名,ISNULL(TS.切替日付,'') 切替日付,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(TS.契約先,'') = '' THEN '' ELSE '1' END 対象フラグ "
    gSL_Select = gSL_Select & " FROM 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN TNS送信対象取引先マスタ TS "
    gSL_Select = gSL_Select & " ON HT.取引先 = TS.取引先 "
    gSL_Select = gSL_Select & " AND HT.受入 = TS.受入 "
    gSL_Select = gSL_Select & " AND HT.契約先 = TS.契約先 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & " ON TR.取引先CD = HT.取引先 "
    gSL_Select = gSL_Select & " AND TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & " WHERE HT.契約先 = '" & RTrim(HTM.契約先) & "'"
    gSL_Select = gSL_Select & " AND   HT.品目 = 1 "
    gSL_Select = gSL_Select & " AND   HT.手配区分 = 3 "                 '------ 手配区分：受給のみ
    gSL_Select = gSL_Select & " GROUP BY HT.契約先,HT.取引先,HT.受入,TR.略称,TS.契約先,TS.切替日付"
    gSL_Select = gSL_Select & " ORDER BY HT.契約先,HT.取引先,HT.受入"
    
    Call DBInput("Back")
PB_Back_Ed:
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Torcd.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    Call CisFrm.MousePT(11)
    gSL_Select = " SELECT *"
    gSL_Select = " SELECT HT.契約先,HT.取引先,HT.受入 "
    gSL_Select = gSL_Select & " FROM 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " WHERE HT.契約先 > '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " AND   HT.品目 = 1 "
    gSL_Select = gSL_Select & " AND   HT.手配区分 = 3 "                 '------ 手配区分：受給のみ
    gSL_Select = gSL_Select & " GROUP BY HT.契約先,HT.取引先,HT.受入"
    gSL_Select = gSL_Select & " ORDER BY HT.契約先 "
    If Not HTMRead(gSL_Select, 1) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     最終レコードです            "
            .MB_Title = "【前進】"
            .MB_Button = OK
            .MBOX
        End With
        GoTo PB_Next_Ed
    End If
    
    iH1_Torcd = RTrim(HTM.契約先)
    Call iH1_Torcd_LostFocus
    
    gSL_Select = " SELECT HT.契約先,HT.取引先,HT.受入,ISNULL(TR.略称,'') 取引先名,ISNULL(TS.切替日付,'') 切替日付,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(TS.契約先,'') = '' THEN '' ELSE '1' END 対象フラグ "
    gSL_Select = gSL_Select & " FROM 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN TNS送信対象取引先マスタ TS "
    gSL_Select = gSL_Select & " ON HT.取引先 = TS.取引先 "
    gSL_Select = gSL_Select & " AND HT.受入 = TS.受入 "
    gSL_Select = gSL_Select & " AND HT.契約先 = TS.契約先 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & " ON TR.取引先CD = HT.取引先 "
    gSL_Select = gSL_Select & " AND TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & " WHERE HT.契約先 = '" & RTrim(HTM.契約先) & "'"
    gSL_Select = gSL_Select & " AND   HT.品目 = 1 "
    gSL_Select = gSL_Select & " AND   HT.手配区分 = 3 "                 '------ 手配区分：受給のみ
    gSL_Select = gSL_Select & " GROUP BY HT.契約先,HT.取引先,HT.受入,TR.略称,TS.契約先,TS.切替日付"
    gSL_Select = gSL_Select & " ORDER BY HT.契約先,HT.取引先,HT.受入"
    
    Call DBInput("Next")
PB_Next_Ed:
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_Torcd.SetFocus
    Dummy.Enabled = False
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_Torcd(0): Exit Sub
End Sub
'+----------------------------+
'+       取引先検索            +
'+----------------------------+
Private Sub Look_Torcd(tTKb As Byte)

    RV_Left = 0:    RV_Top = 0
    If tTKb = 0 Then
        RV_TorKb = 5
    Else
        RV_TorKb = 8
    End If
    RV_Call = "C"
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        If tTKb = 0 Then
            iH1_Torcd = RTrim(RV_TorcdK)
            Call iH1_Torcd_LostFocus
        End If
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
    If RTrim(iH1_Torcd) <> "" Then
        If TorNmGet(iH1_Torcd, 5, 0) Then
            H1lb_Tornm = RTrim(TRM.略称)
        End If
    End If
End Sub
'+-----------------------------+
'+  Dummy
'+-----------------------------+
Private Sub B_TabDummy_GotFocus()
    
    Call TextToVsGrid(Val(Input_Row.Tag))
    Call Grid_RowChange(vbKeyDown)
    SendKeys "{Tab}"
End Sub
'+-----------------------------+
'+  iB1_YMD
'+-----------------------------+
Private Sub iB1_YMD_GotFocus()
    Set ActObj = iB1_YMD
End Sub
Private Sub iB1_YMD_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub

'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
   Dim W_Color     As Long
   W_Color = &H80000008 '初期色ｾｯﾄ(黒)
   Select Case SyoriKB
     Case "ADD"
          SyoriNM = "追加"
          K_Sykbnm.ForeColor = &HFFFF&   '(黄色)
     Case "MNT"
          SyoriNM = "修正"
          K_Sykbnm.ForeColor = &HFF0000  '(青色)
     Case "DEL"
          SyoriNM = "削除"
          K_Sykbnm.ForeColor = &HFF&     '(赤色)
          W_Color = &HFF&
     Case "LOK"
          SyoriNM = "照会"
          K_Sykbnm.ForeColor = &HFF00FF   '(ピンク)
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
       If SyoriKB <> "LOC" Then
          If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
             gObj.Name = "PB_DEL" Then
             gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
          End If
       End If
   Next gObj
   Timer2.Enabled = False
   If SyoriNM = "" Then
      K_Sykbnm.Visible = False
      K_Sykbnm.Caption = ""
      Exit Sub
   End If
   K_Sykbnm.Caption = SyoriNM
   K_Sykbnm.Visible = True
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
    PB_Look.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    
    CountDsp.Visible = B1Mode
    
    If Pro = "H1" Then
       mCount = 0
       Input_Row.Visible = False
    End If
    
    If Pro = "B1" Then
       vsGrid.SetFocus
'       Input_Row.Visible = True
    End If
    
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
                '.Enabled = True
            End If
            If .Name Like "Chk" & wbc_Enm & "_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid.Clear
        CisVsGrid.NumberSet = All
        CisVsGrid.ReNumber
    End If
    
    mCount = 0: CountDsp = ""
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
'
        ProcHB = "B1"
        Call DispChange(ProcHB)
'        Call SyoriName(SyoriKB)
        vsGrid.Row = 1
        Call VsGridToText(vsGrid.Row)
        Input_Row.Visible = True
        iB1_YMD.SetFocus
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        Call TextToVsGrid(Input_Row.Tag)
        Input_Row.Visible = False
        If Not Body1Chk Then
           Input_Row.Visible = True
           GoTo ReturnPress_Ed
        End If
        If Not DBPut Then
           Input_Row.Visible = True
           GoTo ReturnPress_Ed
        End If
        Input_Row.Visible = True
        Call HeadBodyClear("B")
        vsGrid.Row = 1
        ProcHB = "H1"
        Call DispChange(ProcHB)
        SyoriKB = "MNT"
'''        Call SyoriName(SyoriKB)
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
'契約先
    If Trim(iH1_Torcd) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    If Not TorNmGet(iH1_Torcd, 5) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    契約先　未登録！（取引先マスタ）            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = " SELECT HT.契約先,HT.取引先,HT.受入,ISNULL(TR.略称,'') 取引先名,ISNULL(TS.切替日付,'') 切替日付,"
    gSL_Select = gSL_Select & " CASE WHEN ISNULL(TS.契約先,'') = '' THEN '' ELSE '1' END 対象フラグ "
    gSL_Select = gSL_Select & " FROM 品番取引先マスタ HT"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN TNS送信対象取引先マスタ TS "
    gSL_Select = gSL_Select & " ON HT.取引先 = TS.取引先 "
    gSL_Select = gSL_Select & " AND HT.受入 = TS.受入 "
    gSL_Select = gSL_Select & " AND HT.契約先 = TS.契約先 "
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & " ON TR.取引先CD = HT.取引先 "
    gSL_Select = gSL_Select & " AND TR.取引先区分 = 1 "
    gSL_Select = gSL_Select & " WHERE HT.契約先 = '" & RTrim(iH1_Torcd) & "'"
    gSL_Select = gSL_Select & " AND   HT.品目 = 1 "
    gSL_Select = gSL_Select & " AND   HT.手配区分 = 3 "                 '------ 手配区分：受給のみ
    gSL_Select = gSL_Select & " GROUP BY HT.契約先,HT.取引先,HT.受入,TR.略称,TS.契約先,TS.切替日付"
    gSL_Select = gSL_Select & " ORDER BY HT.契約先,HT.取引先,HT.受入"
    
    If DBInput("INV") Then
        If SyoriKB = "" Then
            SyoriKB = "MNT"
        End If
    Else
        If SyoriKB = "DEL" Then
             With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "    該当データが存在しません！        "
                 .MB_Title = ""
                 .MB_Button = Error
                 .MBOX
             End With
            iH1_Torcd.SetFocus
            Exit Function
        Else
            If CisDB.RecordCount = 0 Then
               With CisFun
                   .MB_Lines = 4
                   .MB_MSG(2) = "    対象仕入先が存在しません        "
                   .MB_MSG(3) = "      （品番取引先マスタ）          "
                   .MB_Title = ""
                   .MB_Button = Error
                   .MBOX
               End With
              iH1_Torcd.SetFocus
              Exit Function
            End If
        End If
        SyoriKB = "MNT"
    End If
       
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
    Call TextToVsGrid(Val(Input_Row.Tag))
    Call GridTextEnabled(False)
'
''ﾒｯｾｰｼﾞﾎﾞｯｸｽ設定
    With CisFun
        .MB_Lines = 4
        .MB_Title = ""
        .MB_Button = Error
    End With
''<< 切替日付チェック >>
    With vsGrid
        For gCnt1 = 1 To .Rows - 1
            If RTrim(.TextMatrix(gCnt1, 5)) <> "" Then          '選択
            '<< 選択時切替日付必須 >>
                If RTrim(.TextMatrix(gCnt1, 3)) = "" Then
                    CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　切替日付　未入力！】            "
                    CisFun.MBOX
                    Set ErrObj = iB1_YMD
                    GoTo Body1Chk_Err
                End If
                If RTrim(.TextMatrix(gCnt1, 3)) <> "" Then
                    iB1_YMD = Mid(RTrim(.TextMatrix(gCnt1, 3)), 1, 4) & Mid(RTrim(.TextMatrix(gCnt1, 3)), 6, 2) & Mid(RTrim(.TextMatrix(gCnt1, 3)), 9, 2)
                    If Not iB1_YMD.cChkResult Then
                        CisFun.MB_MSG(2) = "    " & gCnt1 & "行目【　切替日付　誤入力！】            "
                        CisFun.MBOX
                        Set ErrObj = iB1_YMD
                        GoTo Body1Chk_Err
                    End If
                End If
            End If
        Next gCnt1
    End With

    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    vsGrid.Row = gCnt1
    Input_Row.Visible = True
    Call VsGridToText(vsGrid.Row)
    ErrObj.SetFocus
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
    
    If Not TSTRead(gSL_Select, 0, 1) Then
        Call TSTClose(1)
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
        Call BodySet
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
Private Sub BodySet()


'Grid表示
    Call GridSet
    Call TSTClose(1)
        

End Sub
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim AddFlg  As Boolean
    Dim wYDJun As Byte

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     更新処理を行います      "
        .MB_Title = "更新処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    With CisDB
        On Error GoTo DBPut_Err
    ' ( 削除処理 )
            '
        .SQL = "DELETE FROM TNS送信対象取引先マスタ "
        .SQL = .SQL & " WHERE 契約先 = '" & RTrim(iH1_Torcd) & "'"
        .DBExec

        On Error GoTo 0
    End With
'   マスタ項目移送
    For gInt = 1 To vsGrid.Rows - 1
        If RTrim(vsGrid.TextMatrix(gInt, 5)) <> "" Then             '選択行のみ
            With TST
                .契約先 = RTrim(iH1_Torcd)
                .取引先 = RTrim(vsGrid.TextMatrix(gInt, 4))
                .受入 = RTrim(vsGrid.TextMatrix(gInt, 2))
                .切替日付 = Mid(RTrim(vsGrid.TextMatrix(gInt, 3)), 1, 4) & Mid(RTrim(vsGrid.TextMatrix(gInt, 3)), 6, 2) & Mid(RTrim(vsGrid.TextMatrix(gInt, 3)), 9, 2)
                .作成者 = gTanto
            End With
        ' ( 追加処理 )
            Call TSTInsert
        End If
    Next gInt
    
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
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 中中, あり, 2, 0, "№")
        Call .FixedSet(中中, 左中, あり, 29, 1, "仕 入 先")
        Call .FixedSet(中中, 中中, あり, 4, 2, "受入")
        Call .FixedSet(中中, 中中, あり, 14, 3, "送信終了日")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 中中, あり, 0, 4, "x取引先")
        Call .FixedSet(中中, 左中, なし, 0, 5, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3")
        .InitDisp
    End With
    
End Sub
'******************************************'
'*****       グリッド内容セット        *****
'******************************************'
Private Sub GridSet()
    Dim ID          As Integer
    Dim GridCol()   As Integer
' 抽出処理 初期ｾｯﾄ
'    B1lb_Kensu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid
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
    Do Until Not TST_RDSTS
        'Grid Row,Col
        ID = ID + 1
        With vsGrid
            .TextMatrix(ID, GridCol(0)) = ID
            If RTrim(TST.取引先名) <> "" Then
               .TextMatrix(ID, GridCol(1)) = RTrim(TST.取引先) & ":" & RTrim(TST.取引先名)
            Else
               .TextMatrix(ID, GridCol(1)) = RTrim(TST.取引先)
            End If
            .TextMatrix(ID, GridCol(2)) = RTrim(TST.受入)
            If RTrim(TST.切替日付) <> "" Then
               .TextMatrix(ID, GridCol(3)) = Mid(TST.切替日付, 1, 4) & "年" & Mid(TST.切替日付, 5, 2) & "月" & Mid(TST.切替日付, 7, 2) & "日"
            Else
               .TextMatrix(ID, GridCol(3)) = ""
            End If
            '---非表示
            .TextMatrix(ID, GridCol(4)) = RTrim(TST.取引先)
            If RTrim(TST.対象フラグ) = "" Then
               .TextMatrix(ID, GridCol(5)) = ""
'               .Col = 0: .Col = 2: .CellBackColor = &HC0FFC0
               
               .TextMatrix(ID, 5) = ""
               .Row = ID: .Col = 0: .ColSel = 2: .CellBackColor = &HC0FFC0
               .Col = 0: .CellBackColor = &HE0E0E0
            Else
               .TextMatrix(ID, GridCol(5)) = "1"
'               .Col = 0: .Col = 2: .CellBackColor = &HFFC0FF
               
               .Row = ID: .Col = 0: .ColSel = 2: .CellBackColor = &HFFC0FF
               .Col = 0: .CellBackColor = &HFFC0FF
               
               mCount = mCount + 1
            End If
        End With
        Call TSTReadNext(1)
    Loop

    With vsGrid
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    
    CountDsp = "選択件数：" & Format(mCount, "##0") & "件"
    
End Sub
Private Sub VSGrid_Click()
    Dim I   As Integer
    Dim wRow As Integer
    
    wRow = vsGrid.Row
    Call TextToVsGrid(Val(Input_Row.Tag))
    vsGrid.Row = wRow
    DoEvents
    If RTrim$(vsGrid.TextMatrix(vsGrid.Row, 1)) = "" Then Exit Sub
    Call GridTextEnabled(True)
    Call VsGridToText(vsGrid.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
    
    CountDsp = "選択件数：" & Format(mCount, "##0") & "件"
    
End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
'Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode <> vbKeySpace Then Exit Sub
'    Call VSGrid_Click
'End Sub
Private Sub VSGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown _
        Then Exit Sub
    
    Call TextToVsGrid(Val(Input_Row.Tag))
    Call VsGridToText(vsGrid.Row)
    iB1_YMD.SetFocus
    
End Sub
Private Sub vsGrid_Scroll()
    'スクロール時はテキストボックスを非表示にします。
'    Call GridTextEnabled(False)
End Sub
'==========================================================='
'   入力テキスト表示・非表示
'==========================================================='
Private Sub GridTextEnabled(EnabledFlg As Boolean)
    Input_Row.Visible = EnabledFlg
End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub


'テキストからグリッドに退避
    Call TextToVsGrid(Val(Input_Row.Tag))
    gInt = vsGrid.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > Ttl_Kensu Then
            GoTo Grid_RowChange_End
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
        If gInt + CisVsGrid.DispMax > vsGrid.Rows - 1 Then
            gInt = vsGrid.Rows - 1
        Else
            'gInt = gInt + CisVsGrid.DispMax
            gInt = (gInt + CisVsGrid.DispMax) - 1
        End If
    End If
    vsGrid.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText(vsGrid.Row)
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    Dim wStr        As String
    
    With vsGrid
'       >> 切替日付
        If RTrim(iB1_YMD) <> "" Then
            gStr = iB1_YMD
            iB1_YMD = gStr
            If iB1_YMD.cChkResult Then
                    
                With CisFun
                    .DateE_BefVal = iB1_YMD
                    .DateEdit J_YMD
                    vsGrid.TextMatrix(tRow, 3) = .DateE_AfVal
                End With
                .TextMatrix(tRow, 5) = "1"
                .Row = tRow: .Col = 0: .ColSel = 3: .CellBackColor = &HFFC0FF
                .Col = 0: .CellBackColor = &HFFC0FF
            Else
                .TextMatrix(tRow, 3) = ""
                .TextMatrix(tRow, 5) = ""
                .Row = tRow: .Col = 0: .ColSel = 3: .CellBackColor = &HC0FFC0
                .Col = 0: .CellBackColor = &HE0E0E0
                .Col = 1: .ColSel = 3
            End If
        Else
           .TextMatrix(tRow, 3) = ""
        
           .TextMatrix(tRow, 5) = ""
           .Row = tRow: .Col = 0: .ColSel = 3: .CellBackColor = &HC0FFC0
           .Col = 0: .CellBackColor = &HE0E0E0
           .Col = 1: .ColSel = 3
        End If
        
        
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With vsGrid
        Input_Row.Tag = tRow
        Input_Row.Move Input_Row.Left, vsGrid.CellTop, Input_Row.Width, vsGrid.CellHeight
        
        '切替日付
        If RTrim(.TextMatrix(tRow, 3)) <> "" Then
            With CisFun
                
                .DateE_AfVal = vsGrid.TextMatrix(tRow, 3)
                .DateEdit J_YMD
                iB1_YMD = .DateE_BefVal
            End With
'           iB1_YMD = Mid(RTrim(.TextMatrix(tRow, 3)), 1, 4) & Mid(RTrim(.TextMatrix(tRow, 3)), 6, 2) & Mid(RTrim(.TextMatrix(tRow, 3)), 9, 2)
        Else
           iB1_YMD = ""
        End If
    End With
End Sub


