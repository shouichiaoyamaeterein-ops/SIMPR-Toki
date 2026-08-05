VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Begin VB.Form CXM0145 
   BackColor       =   &H00808000&
   BorderStyle     =   3  '固定ﾀﾞｲｱﾛｸﾞ
   Caption         =   "取引先変換品番選択"
   ClientHeight    =   9750
   ClientLeft      =   4710
   ClientTop       =   3945
   ClientWidth     =   9495
   ForeColor       =   &H00C0C000&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9750
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   Begin Cis3D_v60.CIS3D lb_Wait 
      Height          =   930
      Left            =   2310
      Top             =   4995
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
   Begin Cis3D_v60.CIS3D CIS3D2 
      Height          =   450
      Left            =   180
      Top             =   2670
      Width           =   5370
      _ExtentX        =   9472
      _ExtentY        =   794
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
      Begin Cis3D_v60.CIS3D CIS3D4 
         Height          =   390
         Left            =   30
         Top             =   30
         Width           =   5310
         _ExtentX        =   9366
         _ExtentY        =   688
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
         Begin VB.OptionButton Op_DSPChg 
            BackColor       =   &H00C0C0C0&
            Caption         =   "選択分のみ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   2175
            TabIndex        =   23
            Top             =   90
            Width           =   1560
         End
         Begin VB.OptionButton Op_DSPChg 
            BackColor       =   &H00C0C0C0&
            Caption         =   "全件表示"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   3795
            TabIndex        =   24
            Top             =   90
            Width           =   1320
         End
         Begin VB.OptionButton Op_DSPChg 
            BackColor       =   &H00C0C0C0&
            Caption         =   "選択可品番のみ"
            BeginProperty Font 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   75
            TabIndex        =   22
            Top             =   90
            Value           =   -1  'True
            Width           =   2040
         End
      End
   End
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   450
      Left            =   5910
      Top             =   2670
      Width           =   3420
      _ExtentX        =   6033
      _ExtentY        =   794
      BackColor       =   -2147483633
      ForeColor       =   65535
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
      cBoderColor1    =   8421504
      cBoderColor2    =   -2147483643
      cFont3DColor1   =   16777215
      cFont3DColor2   =   12632256
      cAlingnment     =   7
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   405
         Left            =   1710
         Top             =   30
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
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
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   345
            Left            =   945
            Top             =   30
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   609
            BackColor       =   16777152
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
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   405
         Index           =   0
         Left            =   30
         Top             =   30
         Width           =   1680
         _ExtentX        =   2963
         _ExtentY        =   714
         Caption         =   "選択件数"
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
         cPositionX      =   30
         Begin Cis3D_v60.CIS3D B1lb_SKensu 
            Height          =   360
            Left            =   960
            Top             =   30
            Width           =   690
            _ExtentX        =   1217
            _ExtentY        =   635
            BackColor       =   0
            ForeColor       =   65535
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
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cFont3D         =   2
            cAlingnment     =   7
            cBoderStyle     =   1
         End
      End
   End
   Begin VB.PictureBox PNL_PF 
      Height          =   705
      Left            =   7455
      ScaleHeight     =   645
      ScaleWidth      =   1845
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   420
      Width           =   1905
      Begin CisBtn_60.CisBtn PB_Ent 
         Height          =   585
         Left            =   990
         Top             =   30
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
         BtnCaption      =   "更新"
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
         Left            =   90
         Top             =   30
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
         BtnCaption      =   "戻る"
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
   End
   Begin vsFlexLib.vsFlexArray vsGridS 
      Height          =   6405
      Left            =   180
      TabIndex        =   0
      Top             =   3165
      Width           =   9150
      _Version        =   196608
      _ExtentX        =   16140
      _ExtentY        =   11298
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
      FormatString    =   "№|品番                     |背番号|サイクル|置場    |収容数 |収容器    |"
      Rows            =   21
      Cols            =   9
      BackColor       =   12648447
      BackColorBkg    =   14737632
      FocusRect       =   0
      HighLight       =   2
      RowHeightMin    =   300
      FillStyle       =   1
      SelectionMode   =   1
      BackColorAlternate=   12648447
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   375
      Left            =   0
      Top             =   15
      Width           =   9480
      _ExtentX        =   16722
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 取引先変換品番選択  】"
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
      cPositionX      =   30
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   8370
         Top             =   60
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
   End
   Begin VB.Label H1lb_DaikoK 
      BackStyle       =   0  '透明
      Caption         =   "XXXXX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   26
      Top             =   840
      Width           =   780
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  '不透明
      Height          =   465
      Left            =   6315
      Shape           =   4  '丸みのある長方形
      Top             =   720
      Width           =   1080
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "代行契約先"
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
      Left            =   5040
      TabIndex        =   25
      Top             =   840
      Width           =   1200
   End
   Begin VB.Label H1lb_HiJisya 
      BackStyle       =   0  '透明
      Caption         =   "XXXXX"
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
      Left            =   8640
      TabIndex        =   21
      Top             =   1440
      Width           =   795
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   14
      Left            =   8535
      Shape           =   4  '丸みのある長方形
      Top             =   1365
      Width           =   825
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "被自社"
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
      Index           =   2
      Left            =   7755
      TabIndex        =   20
      Top             =   1455
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   11
      Left            =   7695
      Shape           =   4  '丸みのある長方形
      Top             =   1365
      Width           =   855
   End
   Begin VB.Label H1lb_DTorNm 
      BackStyle       =   0  '透明
      Caption         =   "XXXXXXXXX1XXXXXXXXX2XXXX"
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
      Left            =   5010
      TabIndex        =   19
      Top             =   2235
      Width           =   2880
   End
   Begin VB.Label H1lb_DUke 
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
      Height          =   255
      Left            =   4545
      TabIndex        =   18
      Top             =   2235
      Width           =   255
   End
   Begin VB.Label H1lb_DTorcd 
      BackStyle       =   0  '透明
      Caption         =   "XXXXX-XX"
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
      Left            =   3375
      TabIndex        =   17
      Top             =   2235
      Width           =   960
   End
   Begin VB.Label H1lb_STorcd2 
      BackStyle       =   0  '透明
      Caption         =   "XXXXX-XX:XX"
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
      Left            =   1710
      TabIndex        =   16
      Top             =   2235
      Width           =   1395
   End
   Begin VB.Label H1lb_KTorcd 
      BackStyle       =   0  '透明
      Caption         =   "XXXXX-XX"
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
      Left            =   330
      TabIndex        =   15
      Top             =   2235
      Width           =   960
   End
   Begin VB.Label H1lb_STorcd 
      BackStyle       =   0  '透明
      Caption         =   "XXXXX"
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
      Left            =   6945
      TabIndex        =   14
      Top             =   1455
      Width           =   600
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   8
      Left            =   4890
      Shape           =   4  '丸みのある長方形
      Top             =   2145
      Width           =   4470
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   7
      Left            =   4440
      Shape           =   4  '丸みのある長方形
      Top             =   2145
      Width           =   465
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   6
      Left            =   3300
      Shape           =   4  '丸みのある長方形
      Top             =   2145
      Width           =   1155
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "代行ｺｰﾄﾞ/受入/  代 行 名 称"
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
      Left            =   3360
      TabIndex        =   13
      Top             =   1845
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   5
      Left            =   3300
      Shape           =   4  '丸みのある長方形
      Top             =   1755
      Width           =   6060
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   3
      Left            =   1500
      Shape           =   4  '丸みのある長方形
      Top             =   2145
      Width           =   1815
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "直直支給:受入"
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
      Left            =   1620
      TabIndex        =   12
      Top             =   1845
      Width           =   1560
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   4
      Left            =   1500
      Shape           =   4  '丸みのある長方形
      Top             =   1755
      Width           =   1815
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   2
      Left            =   150
      Shape           =   4  '丸みのある長方形
      Top             =   2145
      Width           =   1365
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "契約先ｺｰﾄﾞ"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   1845
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   1
      Left            =   150
      Shape           =   4  '丸みのある長方形
      Top             =   1755
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   13
      Left            =   6810
      Shape           =   4  '丸みのある長方形
      Top             =   1365
      Width           =   900
   End
   Begin VB.Label H1lb_Tornm 
      BackStyle       =   0  '透明
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   2505
      TabIndex        =   10
      Top             =   840
      Width           =   2400
   End
   Begin VB.Label H1lb_Torcd 
      BackStyle       =   0  '透明
      Caption         =   "XXXXX-XX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1395
      TabIndex        =   9
      Top             =   840
      Width           =   1005
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
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
      Height          =   240
      Left            =   375
      TabIndex        =   8
      Top             =   840
      Width           =   720
   End
   Begin VB.Shape Shape6 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  '不透明
      Height          =   465
      Left            =   150
      Shape           =   4  '丸みのある長方形
      Top             =   720
      Width           =   1140
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "支給元"
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
      Left            =   6060
      TabIndex        =   7
      Top             =   1455
      Width           =   720
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   12
      Left            =   6030
      Shape           =   4  '丸みのある長方形
      Top             =   1365
      Width           =   795
   End
   Begin VB.Label H1lb_Uke 
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
      Height          =   255
      Left            =   5700
      TabIndex        =   6
      Top             =   1455
      Width           =   255
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   15
      Left            =   5580
      Shape           =   4  '丸みのある長方形
      Top             =   1365
      Width           =   465
   End
   Begin VB.Label Label3 
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
      Left            =   5115
      TabIndex        =   5
      Top             =   1455
      Width           =   480
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   10
      Left            =   5070
      Shape           =   4  '丸みのある長方形
      Top             =   1365
      Width           =   525
   End
   Begin VB.Label H1lb_JTornm 
      BackStyle       =   0  '透明
      Caption         =   "XXXXXXXXX1XXXXXXXXX2"
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
      Left            =   2640
      TabIndex        =   4
      Top             =   1455
      Width           =   2400
   End
   Begin VB.Label H1lb_JTorcd 
      BackStyle       =   0  '透明
      Caption         =   "XXXXX-XX"
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1620
      TabIndex        =   2
      Top             =   1455
      Width           =   1005
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '透明
      Caption         =   "自社取引先"
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
      Left            =   240
      TabIndex        =   1
      Top             =   1455
      Width           =   1200
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   0
      Left            =   150
      Shape           =   4  '丸みのある長方形
      Top             =   1365
      Width           =   1365
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00C0FFC0&
      BackStyle       =   1  '不透明
      Height          =   405
      Index           =   9
      Left            =   1500
      Shape           =   4  '丸みのある長方形
      Top             =   1365
      Width           =   3585
   End
   Begin VB.Shape Shape7 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  '不透明
      Height          =   465
      Left            =   1275
      Shape           =   4  '丸みのある長方形
      Top             =   720
      Width           =   3705
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFC0C0&
      BackStyle       =   1  '不透明
      Height          =   465
      Left            =   4965
      Shape           =   4  '丸みのある長方形
      Top             =   720
      Width           =   1365
   End
End
Attribute VB_Name = "CXM0145"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   取引先変換品番選択
'**       フォームID    :   CXM0145
'**       処理概要      :
'**
'**       作  成  日    :   2006/02/09  By CIS
'**       変  更  日    :   2006/03/15  By CIS  支給元２の表示を直直支給元に変更
'**                                             被自社に直直支給元を代入
'**                         2006/09/05  By CIS  バグ修正(削除がきかない）
'**                         2007/11/17  By CIS  代行対応（代行契約先追加）
'**                                             品番選択対象は手配区分='3'受給のみ
'**
'**********************************************************'
    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim FormCap         As String
    
    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid       As New CisVsGrid3
    
    Dim mCount          As Long
    Dim mLong           As Long

    Dim mClick_Cnt      As Long
    Dim selKeyIn        As Byte
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
'   #-------------------#
'   #  項 目 初 期 化    #
'   #-------------------#
    Set CisFrm = New CISFormContorl
    
    Set CisFrm.ThisForm = Me
'    Call CisFrm.InitFld
    
    '+-----------------------+
    '+  グリッド初期設定
    '+-----------------------+
    With CisVsGrid
        Set .GridObj = vsGridS
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 9
        .InitGet = False
        .Init
    End With
    Call GridInit
'   #------------------#
'   # 画面中央位置表示  #
'   #------------------#
    If RV_Left <> 0 Or RV_Top <> 0 Then
        Me.Move RV_Left, RV_Top
    Else
        Call CisFrm.Frm_Center(False, x1024y768)
    End If

'   +---------------------------------+
'   + パラメータ情報表示
'   +---------------------------------+
    Call Init_Para_Set
    
    RV_SelCnt = 0
    RV_CrtRep = 0
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
'
    FormCap = Me.Caption
    
    RV_Rtn = False
    
    FormAct = Not FormAct
    
    lb_Wait.Visible = True
    
    DoEvents
    
    Call Op_DSPChg_Click(0)
''
''   +---------------------------------+
''   + 対象品番読み込み
''   +---------------------------------+
'    If Not Hin_Dsp_Set Then
'       Me.Hide
'
'    End If
'
'    Call GridSet

    lb_Wait.Visible = False

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
           Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click    '【取消】
           Case vbKeyReturn
                               If PB_Ent.Visible Then                      '【入力】
'                                  Dummy.Enabled = True: Dummy.SetFocus: DoEvents
                                  Call PB_ENT_Click
'                                  Dummy.Enabled = False
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
    Call PB_Can_Click
End Sub
Private Sub Op_DSPChg_Click(Index As Integer)
    Op_DSPChg(0).BackColor = &HC0C0C0
    Op_DSPChg(1).BackColor = &HC0C0C0
    Op_DSPChg(2).BackColor = &HC0C0C0
    
    Op_DSPChg(Index).BackColor = &HC0FFC0

    Call CisFrm.MousePT(11)
    
    lb_Wait.Visible = True
    DoEvents
    CisVsGrid.Init
    Call GridInit
'   +---------------------------------+
'   + 対象品番読み込み
'   +---------------------------------+
    If Not Hin_Dsp_Set Then
        If Index = 0 Then
            Call PB_Can_Click                        '対象品番が存在しない場合は終了
            Call CisFrm.MousePT(1)
            Exit Sub
        End If
    End If
    
    Call GridSet

    lb_Wait.Visible = False
    vsGridS.SetFocus
    DoEvents
    
    Call CisFrm.MousePT(1)

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
Private Sub PB_Can_Click()
    Unload CXM0145
End Sub
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
        Call .FixedSet(中中, 左中, あり, 25, 1, "品        番")
        Call .FixedSet(中中, 中中, あり, 6, 2, "背番号")
        Call .FixedSet(中中, 中中, あり, 7, 3, "ｻｲｸﾙ")
        Call .FixedSet(中中, 中中, あり, 10, 4, "置場")
        Call .FixedSet(中中, 右中, あり, 6, 5, "収容数")
        Call .FixedSet(中中, 右中, あり, 10, 6, "収容器")
        Call .FixedSet(中中, 左中, あり, 1, 7, "select")
        Call .FixedSet(中中, 左中, あり, 25, 8, "品番")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")
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
    Dim DspChg_Flg  As Byte
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = mCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGridS
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
    
'    B1lb_Kensu = Format(Ttl_Kensu, "###,##0")
    ID = 0
    mClick_Cnt = 0
    
    Do Until Not HTM_RDSTS
        With vsGridS
            ID = ID + 1
            DspChg_Flg = 0
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            '選択
            .TextMatrix(ID, GridCol(7)) = "0"
            
            gSL_Select = "SELECT * FROM 取引先変換マスタ "
            gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(RV_TorcdK) & "'"
            gSL_Select = gSL_Select & " AND   代行契約先 = '" & RTrim(RV_DaikoK) & "'"                  ' 2007/12/17 ADD
            gSL_Select = gSL_Select & " AND   取引先 = '" & RTrim(RV_Torcd) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(契約取引先,'') = '" & RTrim(RV_KTorcd) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(契約工場,'') = '" & RTrim(RV_KTorkj) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(契約取引先CD,'') = '" & RTrim(RV_KTorcd) & RTrim(RV_KTorkj) & "'"
            gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(RV_Ukeir) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(支給元,'') = '" & RTrim(RV_STorcd) & "'"
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'            gSL_Select = gSL_Select & " AND   ISNULL(被自社,'') = '" & RTrim(RV_HiJisya) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(被自社,'') = '" & RTrim(RV_TyokTorcd) & "'"
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
            gSL_Select = gSL_Select & " AND   ISNULL(支給元2,'') = '" & RTrim(RV_STorcd2) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(支給元工場2,'') = '" & RTrim(RV_STorkj2) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(支給元2受入,'') = '" & RTrim(RV_STorUke2) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(代行,'') = '" & RTrim(RV_DTorcd) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(代行工場,'') = '" & RTrim(RV_DTorkj) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(代行受入,'') = '" & RTrim(RV_DUke) & "'"
            gSL_Select = gSL_Select & " AND   ISNULL(代行名称,'') = '" & RTrim(RV_DTorNm) & "'"
            gSL_Select = gSL_Select & " AND   品番 = '" & RTrim(HTM.品番) & "'"
            If THMRead(gSL_Select, 1, 1) Then
                .Row = ID
                .Col = 0
                .CellBackColor = &HFFFFC0
                .Col = 1: .ColSel = 6
                .CellBackColor = &HFFFFC0                      ' 設定済み &HFF8080
                .TextMatrix(ID, GridCol(7)) = "1"
                mClick_Cnt = mClick_Cnt + 1
            Else
                gSL_Select = "SELECT * FROM 取引先変換マスタ "
                gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(RV_TorcdK) & "'"
                gSL_Select = gSL_Select & " AND   代行契約先 = '" & RTrim(RV_DaikoK) & "'"              ' 2007/12/17 ADD
                gSL_Select = gSL_Select & " AND   取引先 = '" & RTrim(RV_Torcd) & "'"
                gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(RV_Ukeir) & "'"
                gSL_Select = gSL_Select & " AND   品番 = '" & RTrim(HTM.品番) & "'"
                gSL_Select = gSL_Select & " AND   ( ISNULL(契約取引先,'') <> '" & RTrim(RV_KTorcd) & "'"
                gSL_Select = gSL_Select & " OR    ISNULL(契約工場,'') <> '" & RTrim(RV_KTorkj) & "'"
                gSL_Select = gSL_Select & " OR    ISNULL(契約取引先CD,'') <> '" & RTrim(RV_KTorcd) & RTrim(RV_KTorkj) & "'"
                gSL_Select = gSL_Select & " OR    ISNULL(支給元,'') <> '" & RTrim(RV_STorcd) & "'"
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'                gSL_Select = gSL_Select & " OR    ISNULL(被自社,'') <> '" & RTrim(RV_HiJisya) & "'"
                gSL_Select = gSL_Select & " OR    ISNULL(被自社,'') <> '" & RTrim(RV_TyokTorcd) & "'"
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
                gSL_Select = gSL_Select & " OR    ISNULL(支給元2,'') <> '" & RTrim(RV_STorcd2) & "'"
                gSL_Select = gSL_Select & " OR    ISNULL(支給元工場2,'') <> '" & RTrim(RV_STorkj2) & "'"
                gSL_Select = gSL_Select & " AND   ISNULL(支給元2受入,'') <> '" & RTrim(RV_STorUke2) & "'"
                gSL_Select = gSL_Select & " OR    ISNULL(代行,'') <> '" & RTrim(RV_DTorcd) & "'"
                gSL_Select = gSL_Select & " OR    ISNULL(代行工場,'') <> '" & RTrim(RV_DTorkj) & "'"
                gSL_Select = gSL_Select & " OR    ISNULL(代行受入,'') <> '" & RTrim(RV_DUke) & "'"
                gSL_Select = gSL_Select & " OR    ISNULL(代行名称,'') <> '" & RTrim(RV_DTorNm) & "' )"
                If THMRead(gSL_Select, 1, 1) Then
                   If Op_DSPChg(1).Value = True Then
                      .Row = ID
                      .Col = 0
                      .CellBackColor = &HFFC0FF
                      .Col = 1: .ColSel = 6
                      .CellBackColor = &HFFC0FF             ' 他で設定済み &HFF8080
                      .TextMatrix(ID, GridCol(7)) = "2"
                   Else
                      ID = ID - 1
                      DspChg_Flg = 1
                   End If
                End If
            End If
            
            If DspChg_Flg = 0 Then
                '品番
                .TextMatrix(ID, GridCol(1)) = HTM.表示品番
                .TextMatrix(ID, 8) = HTM.品番
                '背番号
                .TextMatrix(ID, GridCol(2)) = HTM.背番号
                'サイクル
                .TextMatrix(ID, GridCol(3)) = Mid(HTM.サイクル, 1, 1) & "-" & Mid(HTM.サイクル, 2, 2) & "-" & Mid(HTM.サイクル, 4, 2)
                '置場
                .TextMatrix(ID, GridCol(4)) = HTM.置場
                '収容数
                If RTrim(HTM.変更日) < Format(Now, "YYYYMMDD") Then
                   .TextMatrix(ID, GridCol(5)) = HTM.収容数1
                Else
                   .TextMatrix(ID, GridCol(5)) = HTM.収容数2
                End If
                '収容器
                If RTrim(HTM.変更日) < Format(Now, "YYYYMMDD") Then
                   .TextMatrix(ID, GridCol(6)) = HTM.収容器1
                Else
                   .TextMatrix(ID, GridCol(6)) = HTM.収容器2
                End If
            End If
            
       End With
        
        Call HTMReadNext(0)
    Loop
    
    Call HTMClose(0)
    B1lb_SKensu = Format(mClick_Cnt, "###,###")
    B1lb_Kensu = Format(ID, "###,##0")
'    vsGridS.Rows = ID + 1
    
    With vsGridS
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
    selKeyIn = 0
End Sub
'+----------------------------+
'+  品番取引先選択
'+----------------------------+
Private Sub ReturnPress()
'    If CisFun.Val2(B1lb_SKensu) = 0 Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    品番を選択して下さい            "
'            .MB_Title = ""
'            .MB_Button = Error
'            .MBOX
'            Exit Sub
'        End With
'    End If
    
    If Not DBPut Then Exit Sub
    
    Call PB_Can_Click
End Sub
'+----------------------+
'+    追加･更新 処理     +
'+----------------------+
Private Function DBPut()
    DBPut = False
    Dim AddFlg  As Boolean
    Dim wYDJun As Byte

    With CisFun
        If .Val2(B1lb_SKensu) = 0 Then
            
'           < 削除 or 代表で更新 >
            If Not DB_DelOrUpd Then
                Exit Function
            Else
                GoTo DBPut_ED
            End If
        Else
            .MB_Lines = 4
            .MB_MSG(2) = "     更新処理を行います      "
            .MB_Title = "更新処理"
            .MB_MSG(3) = "     よろしいですか？      "
            .MB_Button = OK_CAN
            If Not .MBOX Then Exit Function
        End If
    End With
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    With CisDB
        On Error GoTo DBPut_Err
    ' ( 取引先変換マスタ削除処理 )
        .SQL = "DELETE FROM 取引先変換マスタ "
        .SQL = .SQL & " WHERE 契約先 = '" & RTrim(RV_TorcdK) & "'"
        .SQL = .SQL & " AND   代行契約先 = '" & RTrim(RV_DaikoK) & "'"                      ' 2007/12/17 ADD
        .SQL = .SQL & " AND   取引先 = '" & RTrim(RV_Torcd) & "'"
        .SQL = .SQL & " AND   ISNULL(契約取引先,'') = '" & RTrim(RV_KTorcd) & "'"
        .SQL = .SQL & " AND   ISNULL(契約工場,'') = '" & RTrim(RV_KTorkj) & "'"
        .SQL = .SQL & " AND   ISNULL(契約取引先CD,'') = '" & RTrim(RV_KTorcd) & RTrim(RV_KTorkj) & "'"
        .SQL = .SQL & " AND   受入 = '" & RTrim(RV_Ukeir) & "'"
        .SQL = .SQL & " AND   ISNULL(支給元,'') = '" & RTrim(RV_STorcd) & "'"
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'        .SQL = .SQL & " AND   ISNULL(被自社,'') = '" & RTrim(RV_HiJisya) & "'"             '2008/01/30
        .SQL = .SQL & " AND   ISNULL(被自社,'') = '" & RTrim(RV_TyokTorcd) & "'"
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
        .SQL = .SQL & " AND   ISNULL(支給元2,'') = '" & RTrim(RV_STorcd2) & "'"
        .SQL = .SQL & " AND   ISNULL(支給元工場2,'') = '" & RTrim(RV_STorkj2) & "'"
        .SQL = .SQL & " AND   ISNULL(支給元2受入,'') = '" & RTrim(RV_STorUke2) & "'"
        .SQL = .SQL & " AND   ISNULL(代行,'') = '" & RTrim(RV_DTorcd) & "'"
        .SQL = .SQL & " AND   ISNULL(代行工場,'') = '" & RTrim(RV_DTorkj) & "'"
        .SQL = .SQL & " AND   ISNULL(代行受入,'') = '" & RTrim(RV_DUke) & "'"
        .SQL = .SQL & " AND   ISNULL(代行名称,'') = '" & RTrim(RV_DTorNm) & "'"
'        .SQL = .SQL & " AND   品番 <> '' "
        .DBExec
        On Error GoTo 0
    End With


'   マスタ項目移送
    For gInt = 1 To vsGridS.Rows - 1
        If RTrim(vsGridS.TextMatrix(gInt, 7)) = "1" Then
            With THM
                .契約先 = RTrim(RV_TorcdK)
                .代行契約先 = RTrim(RV_DaikoK)                  '2007/12/17 ADD
                .取引先 = RTrim(RV_Torcd)
                .契約取引先 = RTrim(RV_KTorcd)
                .契約工場 = RTrim(RV_KTorkj)
                .契約取引先CD = RTrim(RV_KTorcd) & RTrim(RV_KTorkj)
                .受入 = RTrim(RV_Ukeir)
                .支給元 = RTrim(RV_STorcd)
                .支給元2 = RTrim(RV_STorcd2)
                .支給元工場2 = RTrim(RV_STorkj2)
                .支給元2受入 = RTrim(RV_STorUke2)
                .代行 = RTrim(RV_DTorcd)
                .代行工場 = RTrim(RV_DTorkj)
                .代行受入 = RTrim(RV_DUke)
                .代行名称 = RTrim(RV_DTorNm)
                .作成者 = gTanto
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'                .被自社 = RTrim(RV_HiJisya)
                .被自社 = RTrim(RV_TyokTorcd)
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
                .品番 = RTrim(vsGridS.TextMatrix(gInt, 8))
            End With
        ' ( 追加処理 )
            Call THMInsert
            
            RV_SelCnt = RV_SelCnt + 1
        End If
    Next gInt
    
'    '【代表】設定の確認及び追加
'    gSL_Select = "SELECT * FROM 取引先変換マスタ "
'    gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(RV_TorcdK) & "'"
'    gSL_Select = gSL_Select & " AND   取引先 = '" & RTrim(RV_Torcd) & "'"
'    gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(RV_Ukeir) & "'"
'    gSL_Select = gSL_Select & " AND   品番 = ''"
'    If Not THMRead(gSL_Select, 1) Then
'
'        '存在しない場合は追加
'        Call ItemsClearTHM
'        With THM
'            .契約先 = RTrim(RV_TorcdK)
'            .取引先 = RTrim(RV_Torcd)
'            .契約取引先 = RTrim(RV_KTorcd)
'            .契約工場 = RTrim(RV_KTorkj)
'            .契約取引先CD = RTrim(RV_KTorcd) & RTrim(RV_KTorkj)
'            .受入 = RTrim(RV_Ukeir)
'            .支給元 = RTrim(RV_STorcd)
'            .支給元2 = RTrim(RV_STorcd2)
'            .支給元工場2 = RTrim(RV_STorkj2)
'            .支給元2受入 = RTrim(RV_STorUke2)
'            .代行 = RTrim(RV_DTorcd)
'            .代行工場 = RTrim(RV_DTorkj)
'            .代行受入 = RTrim(RV_DUke)
'            .代行名称 = RTrim(RV_DTorNm)
'            .作成者 = gTanto
'            .被自社 = RTrim(RV_HiJisya)
'            .品番 = ""
'        End With
'
'        ' ( 追加処理 )
'        Call THMInsert
'    End If
    
    CisDB.DBTran (TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
DBPut_ED:
    RV_Rtn = True
    DBPut = True
    Exit Function

DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+---------------------------+
'+    削除又は、代表で更新
'+---------------------------+
Private Function DB_DelOrUpd()
    Static swRep            As Byte     '代表設定の有無 ( 0:有り、1:無し )
    Static swRProc          As Byte     '代表処理の有無 ( 0:有り、1:無し )
    
    DB_DelOrUpd = False
    swRep = 0: swRProc = 1

    gSL_Select = "SELECT 取引先 FROM 取引先変換マスタ "
    gSL_Select = gSL_Select & " WHERE 契約先 = '" & RTrim(RV_TorcdK) & "'"
    gSL_Select = gSL_Select & " AND   代行契約先 = '" & RTrim(RV_DaikoK) & "'"          ' 2007/12/17 ADD
    gSL_Select = gSL_Select & " AND   取引先 = '" & RTrim(RV_Torcd) & "'"
    gSL_Select = gSL_Select & " AND   受入 = '" & RTrim(RV_Ukeir) & "'"
    gSL_Select = gSL_Select & " AND   品番 = '' "
    If Not THMRead(gSL_Select, 1) Then swRep = 1
    
    If swRep = 1 Then
    
    ' ( HEAD変換情報　バックカラー変更：赤 )
        For gInt = 0 To 15
            Shape1(gInt).BackColor = vbWhite       ' 白
        Next gInt

'   (代表で残す)
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     『品番設定無し』で登録しますか？　      "
            .MB_MSG(3) = "     よろしいですか？      "
            .MB_Title = "更新処理"
            .MB_Button = OK_CAN
            If .MBOX Then
                swRProc = 0
            End If
        End With
    End If

    If swRProc = 1 Then
    
    ' ( HEAD変換情報　バックカラー変更：赤 )
        For gInt = 0 To 15
            Shape1(gInt).BackColor = &HC0C0FF       ' 赤
        Next gInt
    
'   (削除)
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "     削除処理を行います      "
            .MB_Title = "削除処理"
            .MB_MSG(3) = "     よろしいですか？      "
            .MB_Button = OK_CAN
            If Not .MBOX Then GoTo DB_DelOrUpd_ED
        End With
    End If
    
    With CisDB
        .DBTran (TransBegin)              ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
        On Error GoTo DB_DelOrUpd_Err
        
        .SQL = "DELETE FROM 取引先変換マスタ "
        .SQL = .SQL & " WHERE 契約先 = '" & RTrim(RV_TorcdK) & "'"
        .SQL = .SQL & " AND   代行契約先 = '" & RTrim(RV_DaikoK) & "'"                      ' 2007/12/17 ADD
        .SQL = .SQL & " AND   取引先 = '" & RTrim(RV_Torcd) & "'"
        .SQL = .SQL & " AND   ISNULL(契約取引先,'') = '" & RTrim(RV_KTorcd) & "'"
        .SQL = .SQL & " AND   ISNULL(契約工場,'') = '" & RTrim(RV_KTorkj) & "'"
        .SQL = .SQL & " AND   ISNULL(契約取引先CD,'') = '" & RTrim(RV_KTorcd) & RTrim(RV_KTorkj) & "'"
        .SQL = .SQL & " AND   受入 = '" & RTrim(RV_Ukeir) & "'"
        .SQL = .SQL & " AND   ISNULL(支給元,'') = '" & RTrim(RV_STorcd) & "'"
        '-*-*-*-*-*-*-*-* 2006/09/05 Update -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        '.SQL = .SQL & " AND   ISNULL(被自社,'') = '" & RTrim(RV_HiJisya) & "'"
        .SQL = .SQL & " AND   ISNULL(被自社,'') = '" & RTrim(RV_TyokTorcd) & "'"
        '-*-*-*-*-*-*-*-* 2006/09/05 Update -*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*
        .SQL = .SQL & " AND   ISNULL(支給元2,'') = '" & RTrim(RV_STorcd2) & "'"
        .SQL = .SQL & " AND   ISNULL(支給元工場2,'') = '" & RTrim(RV_STorkj2) & "'"
        .SQL = .SQL & " AND   ISNULL(支給元2受入,'') = '" & RTrim(RV_STorUke2) & "'"
        .SQL = .SQL & " AND   ISNULL(代行,'') = '" & RTrim(RV_DTorcd) & "'"
        .SQL = .SQL & " AND   ISNULL(代行工場,'') = '" & RTrim(RV_DTorkj) & "'"
        .SQL = .SQL & " AND   ISNULL(代行受入,'') = '" & RTrim(RV_DUke) & "'"
        .SQL = .SQL & " AND   ISNULL(代行名称,'') = '" & RTrim(RV_DTorNm) & "'"
        
        .DBExec
        
        If swRProc = 0 Then
'       (代表設定作成)
            With THM
                .契約先 = RTrim(RV_TorcdK)
                .代行契約先 = RTrim(RV_DaikoK)                                              ' 2007/12/17 ADD
                .取引先 = RTrim(RV_Torcd)
                .契約取引先 = RTrim(RV_KTorcd)
                .契約工場 = RTrim(RV_KTorkj)
                .契約取引先CD = RTrim(RV_KTorcd) & RTrim(RV_KTorkj)
                .受入 = RTrim(RV_Ukeir)
                .支給元 = RTrim(RV_STorcd)
                .支給元2 = RTrim(RV_STorcd2)
                .支給元工場2 = RTrim(RV_STorkj2)
                .支給元2受入 = RTrim(RV_STorUke2)
                .代行 = RTrim(RV_DTorcd)
                .代行工場 = RTrim(RV_DTorkj)
                .代行受入 = RTrim(RV_DUke)
                .代行名称 = RTrim(RV_DTorNm)
                .作成者 = gTanto
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'                .被自社 = RTrim(RV_HiJisya)
                .被自社 = RTrim(RV_TyokTorcd)
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
                .品番 = ""
            End With

            Call THMInsert
            RV_CrtRep = 1
        End If
        
        On Error GoTo 0
        .DBTran (TransCommit)              ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    End With
    
    DB_DelOrUpd = True

DB_DelOrUpd_ED:
' ( HEAD変換情報　バックカラー変更：緑[初期カラー] )
    For gInt = 0 To 15
        Shape1(gInt).BackColor = &HC0FFC0       ' 緑
    Next gInt
    Exit Function

DB_DelOrUpd_Err:
    CisFun.ErrorBox
    End
End Function
'+--------------------------------+
'+　品番取引先マスタ読み込み
'+--------------------------------+
Private Function Hin_Dsp_Set()
    Hin_Dsp_Set = False
    gSL_Select = ""
    
    If Not Op_DSPChg(2).Value Then
        ' 変換マスタ未登録対象
        gSL_Select = "SELECT 3 区分,HT.*,HN.表示品番 FROM 品番取引先マスタ HT" & vbCrLf
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN " & vbCrLf
        gSL_Select = gSL_Select & " ON HN.品番 = HT.品番 " & vbCrLf
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先変換マスタ TH1 " & vbCrLf
        gSL_Select = gSL_Select & " ON TH1.契約先 = HT.契約先 " & vbCrLf
        gSL_Select = gSL_Select & " AND TH1.代行契約先 = HT.代行契約先 " & vbCrLf
        gSL_Select = gSL_Select & " AND TH1.取引先 = HT.取引先 " & vbCrLf
        gSL_Select = gSL_Select & " AND TH1.受入 = HT.受入 " & vbCrLf
        gSL_Select = gSL_Select & " AND TH1.品番 = HT.品番 " & vbCrLf
        gSL_Select = gSL_Select & " WHERE HT.取引先 = '" & RTrim(RV_Torcd) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   AND HT.受入 = '" & RTrim(RV_Ukeir) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   AND HT.品目 = 1" & vbCrLf
        gSL_Select = gSL_Select & "   AND HT.契約先 = '" & RTrim(RV_TorcdK) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   AND HT.代行契約先 = '" & RTrim(RV_DaikoK) & "'" & vbCrLf          ' 2007/12/17 ADD
        gSL_Select = gSL_Select & "   AND HT.手配区分 = '3'" & vbCrLf                                   ' 2007/12/17 ADD
        gSL_Select = gSL_Select & "   AND ISNULl(TH1.契約先,'') = ''" & vbCrLf
        gSL_Select = gSL_Select & " UNION " & vbCrLf
    End If

'    '変換マスタ同一内容登録済対象
    gSL_Select = gSL_Select & " SELECT 2 区分,HT.*,HN.表示品番 FROM 品番取引先マスタ HT" & vbCrLf
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN " & vbCrLf
    gSL_Select = gSL_Select & " ON HN.品番 = HT.品番 " & vbCrLf
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先変換マスタ TH1 " & vbCrLf
    gSL_Select = gSL_Select & " ON TH1.契約先 = HT.契約先 " & vbCrLf
    gSL_Select = gSL_Select & " AND TH1.代行契約先 = HT.代行契約先 " & vbCrLf
    gSL_Select = gSL_Select & " AND TH1.取引先 = HT.取引先 " & vbCrLf
    gSL_Select = gSL_Select & " AND TH1.受入 = HT.受入 " & vbCrLf
    gSL_Select = gSL_Select & " AND TH1.品番 = HT.品番 " & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.契約取引先,'') = '" & RTrim(RV_KTorcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.契約工場,'') = '" & RTrim(RV_KTorkj) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.契約取引先CD,'') = '" & RTrim(RV_KTorcd) & RTrim(RV_KTorkj) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.支給元,'') = '" & RTrim(RV_STorcd) & "'" & vbCrLf
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'    gSL_Select = gSL_Select & " AND ISNULL(TH1.被自社,'') = '" & RTrim(RV_HiJisya) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.被自社,'') = '" & RTrim(RV_TyokTorcd) & "'" & vbCrLf
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
    gSL_Select = gSL_Select & " AND ISNULL(TH1.支給元2,'') = '" & RTrim(RV_STorcd2) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.支給元工場2,'') = '" & RTrim(RV_STorkj2) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.支給元2受入,'') = '" & RTrim(RV_STorUke2) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.代行,'') = '" & RTrim(RV_DTorcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.代行工場,'') = '" & RTrim(RV_DTorkj) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.代行受入,'') = '" & RTrim(RV_DUke) & "'" & vbCrLf
    gSL_Select = gSL_Select & " AND ISNULL(TH1.代行名称,'') = '" & RTrim(RV_DTorNm) & "'" & vbCrLf
    gSL_Select = gSL_Select & " WHERE HT.取引先 = '" & RTrim(RV_Torcd) & "'" & vbCrLf
    gSL_Select = gSL_Select & "   AND HT.受入 = '" & RTrim(RV_Ukeir) & "'" & vbCrLf
    gSL_Select = gSL_Select & "   AND HT.契約先 = '" & RTrim(RV_TorcdK) & "'" & vbCrLf
    gSL_Select = gSL_Select & "   AND HT.代行契約先 = '" & RTrim(RV_DaikoK) & "'" & vbCrLf                  ' 2007/12/17 ADD
    gSL_Select = gSL_Select & "   AND HT.品目 = 1" & vbCrLf
    gSL_Select = gSL_Select & "   AND ISNULl(TH1.契約先,'') <> ''    " & vbCrLf
    '変換マスタ異内容登録済対象
    If Op_DSPChg(1).Value Then
        gSL_Select = gSL_Select & " UNION " & vbCrLf
        gSL_Select = gSL_Select & " SELECT 1 区分,HT.*,HN.表示品番 FROM 品番取引先マスタ HT" & vbCrLf
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN " & vbCrLf
        gSL_Select = gSL_Select & " ON HN.品番 = HT.品番 " & vbCrLf
        gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先変換マスタ TH1 " & vbCrLf
        gSL_Select = gSL_Select & " ON TH1.契約先 = HT.契約先 " & vbCrLf
        gSL_Select = gSL_Select & " AND TH1.代行契約先 = HT.代行契約先 " & vbCrLf
        gSL_Select = gSL_Select & " AND TH1.取引先 = HT.取引先 " & vbCrLf
        gSL_Select = gSL_Select & " AND TH1.受入 = HT.受入 " & vbCrLf
        gSL_Select = gSL_Select & " AND TH1.品番 = HT.品番 " & vbCrLf
        gSL_Select = gSL_Select & " AND NOT ( ISNULL(TH1.契約取引先,'') = '" & RTrim(RV_KTorcd) & "'" & vbCrLf
        gSL_Select = gSL_Select & " AND ISNULL(TH1.契約工場,'') = '" & RTrim(RV_KTorkj) & "'" & vbCrLf
        gSL_Select = gSL_Select & " AND ISNULL(TH1.契約取引先CD,'') = '" & RTrim(RV_KTorcd) & RTrim(RV_KTorkj) & "'" & vbCrLf
        gSL_Select = gSL_Select & " AND ISNULL(TH1.支給元,'') = '" & RTrim(RV_STorcd) & "'" & vbCrLf
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'        gSL_Select = gSL_Select & " AND ISNULL(TH1.被自社,'') = '" & RTrim(RV_HiJisya) & "'" & vbCrLf
        gSL_Select = gSL_Select & " AND ISNULL(TH1.被自社,'') = '" & RTrim(RV_TyokTorcd) & "'" & vbCrLf
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
        gSL_Select = gSL_Select & " AND ISNULL(TH1.支給元2,'') = '" & RTrim(RV_STorcd2) & "'" & vbCrLf
        gSL_Select = gSL_Select & " AND ISNULL(TH1.支給元工場2,'') = '" & RTrim(RV_STorkj2) & "'" & vbCrLf
        gSL_Select = gSL_Select & " AND ISNULL(TH1.支給元2受入,'') = '" & RTrim(RV_STorUke2) & "'" & vbCrLf
        gSL_Select = gSL_Select & " AND ISNULL(TH1.代行,'') = '" & RTrim(RV_DTorcd) & "'" & vbCrLf
        gSL_Select = gSL_Select & " AND ISNULL(TH1.代行工場,'') = '" & RTrim(RV_DTorkj) & "'" & vbCrLf
        gSL_Select = gSL_Select & " AND ISNULL(TH1.代行受入,'') = '" & RTrim(RV_DUke) & "'" & vbCrLf
        gSL_Select = gSL_Select & " AND ISNULL(TH1.代行名称,'') = '" & RTrim(RV_DTorNm) & "' )" & vbCrLf
        gSL_Select = gSL_Select & " WHERE HT.取引先 = '" & RTrim(RV_Torcd) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   AND HT.受入 = '" & RTrim(RV_Ukeir) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   AND HT.契約先 = '" & RTrim(RV_TorcdK) & "'" & vbCrLf
        gSL_Select = gSL_Select & "   AND HT.代行契約先 = '" & RTrim(RV_DaikoK) & "'" & vbCrLf              ' 2007/12/17 ADD
        gSL_Select = gSL_Select & "   AND HT.品目 = 1" & vbCrLf
        gSL_Select = gSL_Select & "   AND ISNULl(TH1.契約先,'') <> ''    " & vbCrLf
        gSL_Select = gSL_Select & " ORDER BY 区分,HT.品番,HT.背番号" & vbCrLf & vbCrLf
    End If
    
    If Not HTMRead(gSL_Select) Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    選択対象品番が存在しません     "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
        Call HTMClose
        Exit Function
    End If
    
    mCount = CisDB.RecordCount
    
    Hin_Dsp_Set = True
End Function
'+--------------------------------+
'+　初期画面表示(パラメータセット)
'+--------------------------------+
Private Sub Init_Para_Set()

    '契約先
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1 "
    gSL_Select = gSL_Select & " AND 取引先CD = '" & RTrim(RV_TorcdK) & "'"
    If TRMRead(gSL_Select, 1) Then
       If RTrim(TRM.工場CD) <> "" Then
          H1lb_Torcd = RTrim(TRM.取引先) & "-" & RTrim(TRM.工場CD)
       Else
          H1lb_Torcd = RTrim(TRM.取引先)
       End If
       H1lb_Tornm = RTrim(TRM.略称)
    Else
       H1lb_Torcd = ""
       H1lb_Tornm = ""
    End If
    '代行契約先         2007/12/17 ADD
    H1lb_DaikoK = RTrim(RV_DaikoK)
    '自社取引先
    gSL_Select = "SELECT * FROM 取引先マスタ"
    gSL_Select = gSL_Select & " WHERE 取引先区分 = 1 "
    gSL_Select = gSL_Select & " AND 取引先CD = '" & RTrim(RV_Torcd) & "'"
    If TRMRead(gSL_Select, 1) Then
       If RTrim(TRM.工場CD) <> "" Then
          H1lb_JTorcd = RTrim(TRM.取引先) & "-" & RTrim(TRM.工場CD)
       Else
          H1lb_JTorcd = RTrim(TRM.取引先)
       End If
       H1lb_JTornm = RTrim(TRM.略称)
    Else
       H1lb_JTorcd = ""
       H1lb_JTornm = ""
    End If
    H1lb_Uke = RV_Ukeir                                                 '自社受入
    H1lb_STorcd = RV_STorcd                                             '支給元
    H1lb_HiJisya = RV_HiJisya                                           '被自社
    H1lb_KTorcd = RV_KTorcd                                             '契約先ｺｰﾄﾞ
    If RV_KTorkj <> "" Then
       H1lb_KTorcd = H1lb_KTorcd & "-" & RV_KTorkj
    End If
    
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'    If RTrim(RV_STorcd2) <> "" Then gStr = Format(RTrim(RV_STorcd2), "!@@@@@") Else gStr = Space(5)
    If RTrim(RV_TyokTorcd) <> "" Then gStr = Format(RTrim(RV_TyokTorcd), "!@@@@@") Else gStr = Space(5)
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
    If RTrim(RV_STorkj2) <> "" Then
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  Start
'        gStr = gStr & "-" & Format(RTrim(RV_STorkj2), "!@@")
       If RTrim(gStr) <> "" Then
          gStr = gStr & "-" & Format(RTrim(RV_STorkj2), "!@@")
       Else
          gStr = gStr & " " & Format(RTrim(RV_STorkj2), "!@@")
       End If
'[][][][][][][][][] 2006/03/15 [][][][][][][][][][][][][][][]  End
    Else
        gStr = gStr & Space(3)
    End If
    If RTrim$(RV_STorUke2) <> "" Then
        gStr = gStr & ":" & Format(RTrim(RV_STorUke2), "!@@")
    Else
        gStr = gStr & Space(3)
    End If
    H1lb_STorcd2 = gStr                                                 '支給元2
    
    H1lb_DTorcd = RV_DTorcd                                             '代行ｺｰﾄﾞ
    If RV_DTorkj <> "" Then
       H1lb_DTorcd = H1lb_DTorcd & "-" & RV_DTorkj
    End If
    H1lb_DUke = RV_DUke                                                 '代行受入
    H1lb_DTorNm = RV_DTorNm                                             '代行名称

End Sub
Private Sub VSGridS_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
End Sub
Private Sub VSGridS_Click()
    Static swST         As Long
    Static swED         As Long
    
    If selKeyIn = 0 Then
        If vsGridS.MouseRow = 0 And vsGridS.MouseCol <> 0 Then Exit Sub
    End If
    
    mLong = vsGridS.Row
    
'    vsGridS.Redraw = False
    If vsGridS.Row <= vsGridS.RowSel Then
        swST = vsGridS.Row
        swED = vsGridS.RowSel
    Else
        swST = vsGridS.RowSel
        swED = vsGridS.Row
    End If
    
    For gLong = swST To swED
    
        If vsGridS.TextMatrix(gLong, 1) = "" Then Exit For
        
        If vsGridS.TextMatrix(gLong, 7) <> "2" Then               '他取引先・受入にて設定済
        
           If vsGridS.TextMatrix(gLong, 7) = "1" Then
'           < 非選択 >
              vsGridS.TextMatrix(gLong, 7) = "0"
              vsGridS.Row = gLong
              vsGridS.Col = 0
              vsGridS.CellBackColor = &H8000000F                    ' 灰色(行№Cell:Col=0)
              vsGridS.Col = 1: vsGridS.ColSel = 6
              vsGridS.CellBackColor = &HC0FFFF                      ' 黄色(明細行)
              mClick_Cnt = mClick_Cnt - 1
              B1lb_SKensu = Format(mClick_Cnt, "###,###")
           Else
'           < 選択 >
              vsGridS.TextMatrix(gLong, 7) = "1"
              vsGridS.Row = gLong
              vsGridS.Col = 0
              vsGridS.CellBackColor = &HFFFFC0                      ' 選択行
              vsGridS.Col = 1: vsGridS.ColSel = 6
              vsGridS.CellBackColor = &HFFFFC0
              mClick_Cnt = mClick_Cnt + 1
              B1lb_SKensu = Format(mClick_Cnt, "###,###")
           End If
        Else
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    他の変換設定にて選択済の為、選択不可   "
                .MB_MSG(3) = "    【品番】 ：" & vsGridS.TextMatrix(gLong, 1)
                .MB_Title = ""
                .MB_Button = Error
                .MBOX
            End With
        End If
            
    Next gLong
    
'    vsGridS.Redraw = True
    
    vsGridS.Col = 0: vsGridS.Row = mLong
    
    On Error Resume Next
'    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
Private Sub VSGridS_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeySpace Then Exit Sub
    
    selKeyIn = 1
    Call VSGridS_Click
    selKeyIn = 0
    
End Sub
