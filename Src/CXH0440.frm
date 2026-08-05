VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~2.OCX"
Object = "{0D6234D1-DBA2-11D1-B5DF-0060976089D0}#6.0#0"; "TODG6.OCX"
Object = "{8F4FF474-3E18-11D4-AEC2-00000EA57F9E}#1.0#0"; "CISMSG~2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CisYMDwB_6.0.ocx"
Begin VB.Form CXH0440 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注入荷状況照会(支給検収)"
   ClientHeight    =   11070
   ClientLeft      =   90
   ClientTop       =   1485
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
   ScaleHeight     =   11070
   ScaleWidth      =   15375
   WindowState     =   2  '最大化
   Begin VB.PictureBox B1_Area3 
      Height          =   1170
      Left            =   12840
      ScaleHeight     =   1110
      ScaleWidth      =   2355
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   1245
      Width           =   2415
      Begin Cis3D_v60.CIS3D B1lb_HKei 
         Height          =   540
         Left            =   810
         Top             =   30
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   953
         BackColor       =   16761024
         Caption         =   "Z,ZZZ,ZZZ"
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
         cAlingnment     =   4
         cPositionX      =   -30
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   540
         Left            =   30
         Top             =   30
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   953
         BackColor       =   16761024
         ForeColor       =   0
         Caption         =   "発注計"
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
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   540
         Left            =   30
         Top             =   570
         Width           =   780
         _ExtentX        =   1376
         _ExtentY        =   953
         BackColor       =   12648384
         ForeColor       =   0
         Caption         =   "入荷計"
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
      Begin Cis3D_v60.CIS3D B1lb_NKei 
         Height          =   540
         Left            =   810
         Top             =   570
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   953
         BackColor       =   12648384
         Caption         =   "Z,ZZZ,ZZZ"
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
         cAlingnment     =   4
         cPositionX      =   -30
      End
   End
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   5280
      Top             =   4080
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
   Begin VB.PictureBox PNL_ExeMsg 
      Height          =   1800
      Left            =   900
      ScaleHeight     =   1740
      ScaleWidth      =   7425
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   5235
      Visible         =   0   'False
      Width           =   7485
      Begin Cis3D_v60.CIS3D lb_SKensu 
         Height          =   345
         Left            =   3870
         Top             =   480
         Width           =   3525
         _ExtentX        =   6218
         _ExtentY        =   609
         ForeColor       =   0
         Caption         =   "【 集計件数 ＝ ZZZZZ】"
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
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D UC_3D1 
         Height          =   405
         Left            =   45
         Top             =   1305
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   714
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   11.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cBoderStyle     =   1
         Begin Cis3D_v60.CIS3D FXLabel7 
            Height          =   345
            Left            =   30
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "発注数"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D lb_THSu 
            Height          =   345
            Left            =   1260
            Top             =   30
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
            BackColor       =   16777215
            ForeColor       =   0
            Caption         =   "ZZZ,ZZZ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cPositionX      =   -60
         End
         Begin Cis3D_v60.CIS3D lb_TNSu 
            Height          =   345
            Left            =   3690
            Top             =   30
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
            BackColor       =   16777215
            ForeColor       =   0
            Caption         =   "ZZZ,ZZZ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cPositionX      =   -60
         End
         Begin Cis3D_v60.CIS3D FXLabel8 
            Height          =   345
            Left            =   2460
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "入荷数"
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
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D lb_TNYSu 
            Height          =   345
            Left            =   6120
            Top             =   30
            Width           =   1200
            _ExtentX        =   2117
            _ExtentY        =   609
            BackColor       =   16777215
            ForeColor       =   0
            Caption         =   "ZZZ,ZZZ"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ 明朝"
               Size            =   11.25
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   4
            cPositionX      =   -60
         End
         Begin Cis3D_v60.CIS3D CIS3D6 
            Height          =   345
            Left            =   4890
            Top             =   30
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   609
            ForeColor       =   0
            Caption         =   "入荷予定数"
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
            cAlingnment     =   7
         End
      End
      Begin Cis3D_v60.CIS3D lb_TMei 
         Height          =   465
         Left            =   45
         Top             =   825
         Width           =   7350
         _ExtentX        =   12965
         _ExtentY        =   820
         ForeColor       =   16711680
         Caption         =   "品番：XXXXX-XXXXX-XX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   200
      End
      Begin Cis3D_v60.CIS3D PNL_ExeMsgX 
         Height          =   345
         Left            =   45
         Top             =   480
         Width           =   3825
         _ExtentX        =   6747
         _ExtentY        =   609
         ForeColor       =   0
         Caption         =   "【 集計単位 ＝ 品番】"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D FXLabel3 
         Height          =   1320
         Left            =   0
         Top             =   420
         Width           =   7425
         _ExtentX        =   13097
         _ExtentY        =   2328
         ForeColor       =   16576
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   9
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
      End
      Begin CisMsg_v60.CISMSG CISMSG2 
         Height          =   420
         Left            =   15
         TabIndex        =   13
         Top             =   0
         Width           =   7410
         _ExtentX        =   13070
         _ExtentY        =   741
         BackColor       =   -2147483633
         ForeColor       =   12583104
         Caption         =   "※  集 計 結 果  ※"
         BackStyle       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   14.25
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   8421504
         cBoderColor2    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   0
         cAlingnment     =   7
         cBoderStyle     =   1
         gGC2_Green      =   160
         gGradation      =   -1  'True
      End
   End
   Begin VB.PictureBox PNL_PF 
      BackColor       =   &H00808080&
      Height          =   720
      Left            =   10020
      ScaleHeight     =   660
      ScaleWidth      =   5175
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   450
      Width           =   5235
      Begin CisBtn_60.CisBtn PB_END 
         Height          =   585
         Left            =   4305
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
         Left            =   3450
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
         Left            =   2595
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
         Left            =   30
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
      Begin CisBtn_60.CisBtn PB_Csv 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   14
         BtnCaption      =   "出力"
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
         PFBackColor     =   192
      End
      Begin CisBtn_60.CisBtn PB_Prt 
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   13
         BtnCaption      =   "印刷"
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
         PFForeColor     =   0
         PFShadowColor   =   16777215
         PFBackColor     =   65535
      End
   End
   Begin VB.PictureBox B1_Area2 
      Height          =   8490
      Left            =   8295
      ScaleHeight     =   8430
      ScaleWidth      =   6900
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   2460
      Width           =   6960
      Begin TrueOleDBGrid60.TDBGrid DBGridH 
         Bindings        =   "CXH0440.frx":0000
         Height          =   8295
         Left            =   45
         OleObjectBlob   =   "CXH0440.frx":0015
         TabIndex        =   8
         Top             =   60
         Width           =   6825
      End
   End
   Begin VB.PictureBox B1_Area1 
      Height          =   8475
      Left            =   105
      ScaleHeight     =   8415
      ScaleWidth      =   8130
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2460
      Width           =   8190
      Begin TrueOleDBGrid60.TDBGrid DBGrid1 
         Bindings        =   "CXH0440.frx":155BD
         Height          =   8310
         Left            =   90
         OleObjectBlob   =   "CXH0440.frx":155D2
         TabIndex        =   7
         Top             =   60
         Width           =   7980
      End
   End
   Begin VB.PictureBox H1_Area1 
      Height          =   1155
      Left            =   120
      ScaleHeight     =   1095
      ScaleWidth      =   12630
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1260
      Width           =   12690
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   1005
         Left            =   6885
         Top             =   45
         Width           =   3285
         _ExtentX        =   5794
         _ExtentY        =   1773
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
            Left            =   60
            TabIndex        =   4
            Top             =   435
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
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   1005
         Left            =   10170
         Top             =   45
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "背番号"
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
         Begin CisText_V60.CisText iH1_Seban 
            Height          =   375
            Left            =   75
            TabIndex        =   5
            Top             =   420
            Width           =   1110
            _ExtentX        =   1958
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
            Text            =   "XXXXXXXX"
            MaxLength       =   8
            IMEMode         =   2
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D8 
         Height          =   1005
         Left            =   2760
         Top             =   45
         Width           =   3600
         _ExtentX        =   6350
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "仕入先"
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
            Left            =   75
            TabIndex        =   2
            Top             =   435
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
            Left            =   1065
            Top             =   435
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
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   1005
         Left            =   6360
         Top             =   45
         Width           =   525
         _ExtentX        =   926
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "受入"
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
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   360
            Left            =   75
            TabIndex        =   3
            Top             =   435
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
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   1005
         Left            =   11430
         Top             =   45
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   1773
         ForeColor       =   16711680
         Caption         =   "手配区"
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
         Begin VB.ComboBox H1Cmb_Tehai 
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
            Left            =   45
            Style           =   2  'ﾄﾞﾛｯﾌﾟﾀﾞｳﾝ ﾘｽﾄ
            TabIndex        =   6
            Top             =   405
            Width           =   1065
         End
      End
      Begin Cis3D_v60.CIS3D UC_3D2 
         Height          =   1005
         Index           =   0
         Left            =   45
         Top             =   45
         Width           =   2715
         _ExtentX        =   4789
         _ExtentY        =   1773
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
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cFont3D         =   2
         cAlingnment     =   6
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   585
            Left            =   75
            Top             =   240
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   1032
            ForeColor       =   16711680
            Caption         =   "日付"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cBoderStyle     =   2
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_SDate 
            Height          =   345
            Left            =   825
            TabIndex        =   0
            Top             =   120
            Width           =   1815
            _ExtentX        =   2990
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
            Object.Height          =   345
            Object.Width           =   1815
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
               Size            =   11.25
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
         Begin Cis3D_v60.CIS3D PB_SDate 
            Height          =   345
            Left            =   345
            Top             =   120
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   609
            Caption         =   "開始"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin Cis3D_v60.CIS3D PB_EDate 
            Height          =   345
            Left            =   345
            Top             =   555
            Width           =   465
            _ExtentX        =   820
            _ExtentY        =   609
            Caption         =   "終了"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   9.75
               Charset         =   128
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            cFont3DColor1   =   8421504
            cFont3DColor2   =   -2147483643
            cBoderWidth     =   2
            cFont3D         =   2
            cAlingnment     =   7
            cButton         =   -1  'True
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_EDate 
            Height          =   345
            Left            =   825
            TabIndex        =   1
            Top             =   555
            Width           =   1815
            _ExtentX        =   2990
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
            Object.Height          =   345
            Object.Width           =   1815
            cSize           =   1
            cChkResult      =   0   'False
            cTypeIntval     =   0
            BeginProperty cTextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ Ｐゴシック"
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
   Begin Cis3D_v60.CIS3D UC_3D31 
      Height          =   405
      Left            =   0
      Top             =   -15
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "　　　【 発注入荷状況照会(支給検収) 】"
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
      Begin Cis3D_v60.CIS3D UC_3D32 
         Height          =   255
         Left            =   6840
         Top             =   60
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
            Left            =   1155
            Top             =   0
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   255
         Left            =   9705
         Top             =   60
         Width           =   2955
         _ExtentX        =   5212
         _ExtentY        =   450
         BackColor       =   14737632
         Caption         =   "[              ]"
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
         Begin Cis3D_v60.CIS3D lb_Kensu2 
            Height          =   285
            Left            =   240
            Top             =   0
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
   End
   Begin VB.TextBox Dummy 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'ｵﾌ固定
      Left            =   240
      TabIndex        =   9
      Text            =   "D"
      Top             =   210
      Width           =   255
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   510
      Top             =   210
   End
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   1230
      Top             =   210
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   405
      Left            =   0
      Top             =   420
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSAdodcLib.Adodc Adodc2 
      Height          =   405
      Left            =   0
      Top             =   765
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   714
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc2"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.25
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
End
Attribute VB_Name = "CXH0440"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   発注入荷状況照会(受給検収)
'**       フォームID    :   CXH0440
'**       処理概要      :
'**
'**       作  成  日    :   2006/08/28  By CIS
'**       変  更  日    :   2006/12/12  By CIS  グリッドに納入番号追加
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim MaxCnt                  As Long
    Dim NowCnt                  As Long
    Dim mSL_SelectW             As String       ' SQL SELECT(Where)条件文
    Dim mSL_Order               As String       '
    Dim mSl_Select2             As String
    Dim mGridMoveFlg            As Boolean
    
    Dim CisDBGridH              As CisDBGrid ' TDBGridPro60ｸﾗｽ定義
    Dim CisComboTehai           As CisCombo
    
    Dim mGotColor               As Long
    Dim mLostColor              As Long
    
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》            +
'+----------------------------------------------------------+
Private Sub Form_Load()

    If Rv_Call3 = "" Then
    '   #-------------------#
    '   # 二 重 起 動 防 止 #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
        If Not IniGet Then End
    Else
        FormAct = False
    End If
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止
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

'+---------------------+
'+ TrueGrid60 初期設定  +
'+---------------------+

    Set CisDBGridH = New CisDBGrid
    With CisDBGridH
        Set .DBGridObj = DBGridH
        .FrmCaption = Me.Caption
        .MaxSort = 7                                    ' ﾚｲｱｳﾄｿｰﾄ最大数
        .LayOut = False                                 ' ﾚｲｱｳﾄ変更処理 True:有り,False:無し
        .LayOutMax = 0                                  ' ﾚｲｱｳﾄ情報最大数
        .LayOutColMax = 30                              ' ﾚｲｱｳﾄ表示可能項目最大数
        .LayOutPath = gIniExe                           ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ場所
        .LayOutName = CisFun.WSID & "H0440R.Lay"        ' ﾚｲｱｳﾄ情報格納先、ﾌｧｲﾙ名
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        .CsvName = Trim(Me.Caption) & "R.csv"
        .CsvRegKey = gRegKey
        .CsvValue = "CXH0440"
        .GridInitGet = False                            ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ
        .GridProcGet = False
        .Init
    End With
    Call GridInitSet

    '+---------------------+
    '+ ｺﾝﾎﾞﾎﾞｯｸｽ 初期設定  +
    '+---------------------+
    With CisFun
        .INI_Section = "発注処理手配"
        .INI_FileID = gIniExe & "Combo.ini"
        .INI_Default = ""
        .INI_Key = "TYPE"
        .GetIni
        If Trim(.INI_String) = "" Then
           .INI_Key = "TYPE"
           .INI_String = "SQL"
           .PutIni
           .INI_Key = "CNT"
           .INI_String = "4"
           .PutIni
           .INI_Key = "DATA1"
           .INI_String = "SELECT 値,値名称 FROM 名称マスタ"
           .PutIni
           .INI_Key = "DATA2"
           .INI_String = "WHERE 区分名称 = '手配区分'   AND 区分タイプ = ''"
           .PutIni
           .INI_Key = "DATA3"
           .INI_String = "And 値 IN('0','1','3')"
           .PutIni
           .INI_Key = "DATA4"
           .INI_String = "ORDER BY 値"
           .PutIni
        End If
    End With
    
    ' 手配区分
    Set CisComboTehai = New CisCombo
    With CisComboTehai
        Set .Connect = CisDB
        Set .CombObj = H1Cmb_Tehai
        .Control = "発注処理手配"
        .ComboName = ""
        .Ess = True
        .IniPath = gIniExe
        .SetSpace = 50
        .Ess = False
        .ComboInit
        .ComboSet
    End With
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    PNL_ExeMsg.Move 4600, 400

'    If Rv_Call3 = "" Then
        Me.Show
'    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
' Head部初期化
    
    Call HeadBodyClear("H")
    
' ADOｺﾝﾄﾛｰﾙ　接続情報　セット
    Adodc1.ConnectionString = CisDB.ConnectString
    Adodc2.ConnectionString = CisDB.ConnectString
    
    If Rv_Call3 = "K" Then
       iH1_SDate = RV_SYmd
       iH1_Hinbn = RV_Hinbn
       iH1_Seban = RV_Seban
       iH1_Torcd = RV_Torcd
       iH1_Ukeir = RV_Ukeir
       Call iH1_Torcd_LostFocus
       
       If Not Rv_CallObj2 Is Nothing Then
           Unload Rv_CallObj2
       End If
       
       Call ReturnPress
    End If
    
    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Rv_Call3 = "" Then
       Cancel = 1
    End If
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
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
           Case vbKeyF10:       If PB_Csv.Visible Then Call PB_Csv_Click    '【出力】
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

'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
    If Rv_Call3 = "" Then
    '+---------------------+
    '+   ﾃﾞｰﾀﾍﾞｰｽ切断
    '+---------------------+
        Call CisDB.DBDISConnect
    '+---------------------+
    '+   ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    '+---------------------+
        Call CisFrm.UnLoadDisp
        End
    Else
        Unload CXH0440
        If Not Rv_CallObj Is Nothing Then
           Rv_CallObj.Show
        End If
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
    
    If ProcHB = "H1" Then
       Call HeadBodyClear("H")
        iH1_SDate.SetFocus
        GoTo PB_CAN_Ed
    End If
    
    If ProcHB = "B1" Then
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call DBGridClr
        iH1_SDate.SetFocus
        GoTo PB_CAN_Ed
    End If
    If ProcHB = "B2" Then
        ProcHB = "B1"
        Call DispChange(ProcHB)
        GoTo PB_CAN_Ed
    End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    CSV出力           　     +
'+----------------------------+
Private Sub PB_Csv_Click()
    If Adodc2.Recordset.EOF Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データが存在しません｡            "
            .MB_Title = "出力処理"
            .MB_Button = Error
            .MBOX
        End With
    Else
        With CisDBGridH
            Set .Recordset = Adodc2.Recordset
            .PutCsv
        End With
    End If
End Sub
'+----------------------------+
'+    印刷             　     +
'+----------------------------+
Private Sub PB_Prt_Click()
    If Adodc2.Recordset.EOF Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データが存在しません｡            "
            .MB_Title = "印刷処理"
            .MB_Button = Error
            .MBOX
        End With
    Else
        With CisDBGridH
            Set .Recordset = Adodc2.Recordset
            .PutPrt
        End With
    End If
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_SDate" Then iH1_SDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_EDate" Then iH1_EDate.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SSTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd(0): Exit Sub
    If PB_Look.Tag = "iH1_Seban" Then Call Look_Hinbn(1): Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn(0): Exit Sub
End Sub
'+-----------------------------+
'+       取引先検索             +
'+-----------------------------+
Private Sub Look_SSTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 8
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Torcd = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+       取引先受入検索         +
'+-----------------------------+
Private Sub Look_UTorcd(KB As Integer)
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_Torcd = iH1_Torcd
    RV_TorKb = 8
    RV_Kojyo = ""
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Torcd = RV_TorRName
        iH1_Ukeir = RV_Ukeir
    End If
End Sub
'+-----------------------------+
'+       品番検索              +
'+-----------------------------+
Private Sub Look_Hinbn(KB As Integer)
    RV_Left = 0
    RV_Top = 0
        
    RVI_Hinmoku = 1 ' 発注
    RVI_SK = 9      ' かんばん
    RV_Call = ""
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        Select Case KB
            Case 1: iH1_Seban = RV_Seban
            Case Else
                iH1_Hinbn = RV_Hinbn
        End Select
    End If
End Sub
'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'+-----------------------------+
'+  iH1_SDate
'+-----------------------------+
Private Sub iH1_SDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SDate_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_EDate
'+-----------------------------+
Private Sub iH1_EDate_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_EDate_LostFocus()
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
    H1lb_Torcd = ""
    If TorNmGet(iH1_Torcd, 8) Then
        H1lb_Torcd = TRM_RName
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
'+  H1Cmb_Tehai
'+-----------------------------+
Private Sub H1Cmb_Tehai_GotFocus()
    H1Cmb_Tehai.BackColor = iH1_SDate.cGotColor
End Sub
Private Sub H1Cmb_Tehai_LostFocus()
    H1Cmb_Tehai.BackColor = iH1_SDate.cLostColor
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel: PB_ENT.BtnCaption = "入力"
        Case "B1":      B1Mode = True:  B1Color = gPGotSel: PB_ENT.BtnCaption = "詳細"
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area2.Enabled = B1Mode
    B1_Area3.Visible = B1Mode
    
    PB_Look.Visible = H1Mode
    PB_END.Visible = H1Mode
    If Not B1Mode Then
       If PNL_ExeMsg.Visible Then
          PNL_ExeMsg.Visible = False
       End If
    End If
    PB_ENT.Visible = H1Mode
    PB_Prt.Visible = B1Mode
    PB_Csv.Visible = B1Mode
    
    If Rv_Call3 = "K" Then
       PB_END.BtnCaption = "戻る"
       PB_END.Visible = True
    End If
    
'   ﾀﾞﾐｰ項目(入力禁止)
    Dummy.Enabled = False
End Function
'**************************************
'*      Head/Body 初期化処理           *
'**************************************
Private Sub HeadBodyClear(Optional HeadBody As String = "H", Optional ClrType As String = "ALL")
    Static wbc_Enm      As String
    
    If ClrType = "ALL" Then wbc_Enm = HeadBody & "*" Else wbc_Enm = Trim(ClrType)
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then .Text = "": .Enabled = True
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like "Op" & wbc_Enm & "*" Then
               .Value = False
               .BackColor = mLostColor
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "*Cmb*" Then .ListIndex = 0
        End With
    Next gObj
    If HeadBody = "H" Then
        iH1_SDate = Format(Date, "YYYYMMDD")
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
       If Not Head1Chk Then
          GoTo ReturnPress_Ed
       End If
       ProcHB = "B1"
       Call DispChange(ProcHB)
       If Not Adodc1.Recordset.EOF Then
          With DBGrid1
               .SetFocus
               SendKeys "{HOME}"
          End With
       End If
       If Not Adodc2.Recordset.EOF Then
          With DBGridH
               .SetFocus
               SendKeys "{HOME}"
          End With
       End If
        
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
    
    ' 日付
    If Trim(iH1_SDate) <> "" Then
       If CisFun.Val2(iH1_SDate.cDay) = 0 Then
          iH1_EDate = ""
       End If
    End If
    If Trim(iH1_SDate) <> "" And Trim(iH1_EDate) <> "" Then
       If iH1_SDate > iH1_EDate Then
            With CisFun
                .MB_Lines = 5
                .MB_MSG(2) = "    日付を正しく入力して下さい         "
                .MB_MSG(4) = "        開始　＞   終了             "
                .MB_Button = Error
                .MBOX
            End With
            iH1_SDate.SetFocus
            Exit Function
       End If
    End If
    
    If Not DBInput Then
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
'**************************************
'*                                    *
'*       データベース入力 ＆ 表示     *
'*                                    *
'**************************************
Private Function DBInput(Optional InType As String = "", Optional DispType As String = "") As Boolean
    Dim DateStr As String
    Dim DateSt  As String
    Dim DateEd  As String
    Dim BinSt   As String
    Dim BinEd   As String
    
    DBInput = True
    
    Call CisFrm.MousePT(11)
    
    On Error GoTo DBInput_Err:
' 抽出中メッセージ表示
    PNL_Wait.Visible = True
    PNL_PF.Enabled = False
    DoEvents
   
' 初期表示時、条件文作成(Where)
    If ProcHB = "H1" And InType = "" Then
'       表示条件 セット
       CisDBGridH.Select_Where = ""
       With CisFun
            .Where_CreateKB = NewDocument
            ' 日付
            DateStr = "受入年月"
            DateSt = "": DateEd = "": BinSt = "": BinEd = ""
            ' 日付 月指定
            If Trim(iH1_SDate) <> "" And CisFun.Val2(iH1_SDate.cDay) = 0 Then
               .Where_Create AlphaMe, "Substring(" & DateStr & "日,1,6)", iH1_SDate
            End If
            ' 日付開始編集
            If Trim(iH1_SDate) <> "" Then
               DateSt = iH1_SDate
            End If
            ' 日付終了編集
            If Trim(iH1_EDate) <> "" Then
               DateEd = iH1_EDate
            End If
            ' 日付開始のみ
            If Trim(DateSt) <> "" And Trim(DateEd) = "" And CisFun.Val2(iH1_SDate.cDay) <> 0 Then
               .Where_Create AlphaMe, DateStr & "日", DateSt
            End If
            ' 日付終了のみ
            If Trim(DateSt) = "" And Trim(DateEd) <> "" Then
               .Where_Create AlphaMe, DateStr & "日", DateEd, 以下
            End If
            ' 日付開始・終了
            If Trim(DateSt) <> "" And Trim(DateEd) <> "" Then
               .Where_Create AlphaMe, DateStr & "日", DateSt, Between, , DateEd
            End If
            ' 仕入先
            .Where_Create AlphaMe, "仕入先", iH1_Torcd
            ' 受入
            .Where_Create AlphaMe, "受入", iH1_Ukeir
            ' 品番
            .Where_Create AlphaMe, "品番", iH1_Hinbn
            ' 背番号
            .Where_Create AlphaMe, "背番号", iH1_Seban
            ' 手配区分
            Call CisComboTehai.CodeGet
            If Trim(CisComboTehai.Code) <> "" Then
               .Where_Create AlphaMe, "手配区分", CisComboTehai.Code
            End If
            
            CisDBGridH.Select_Where = .Where_Phrase
            
            ' 納品書形態
            
            mSL_SelectW = Replace(.Where_Phrase, "受入年月日", "納入日")
            CisDBGridH.Select_Where = Replace(CisDBGridH.Select_Where, "受入年月日", "入荷日")
       
            gSL_Select = "Select Sum(Case When 分納区分 = 0 Then 発注数 Else 0 End )"
            gSL_Select = gSL_Select & " From 発注入荷状況照会ビュー "
            gSL_Select = gSL_Select & mSL_SelectW
        
            With CisDB
                .SQL = gSL_Select
                If .DBRead(1, 0) Then
                    B1lb_HKei = Format(.Recordset(0), "###,##0")
                End If
                On Error Resume Next
                    .ConnectNo = 0
                    .Recordset.Close
                On Error GoTo 0
            End With
            
            gSL_Select = "Select Sum(入荷数)"

            gSL_Select = gSL_Select & " From 受給検収問合ビュー "

            gSL_Select = gSL_Select & CisDBGridH.Select_Where
                
            With CisDB
                .SQL = gSL_Select
                If .DBRead(1, 0) Then
                    B1lb_NKei = Format(.Recordset(0), "###,##0")
                End If
                On Error Resume Next
                    .ConnectNo = 0
                    .Recordset.Close
                On Error GoTo 0
            End With
            
       End With
    End If
    
    On Error GoTo DBInput_Err:
    
' ﾚｲｱｳﾄ表示内容よりﾃﾞｰﾀ連結名を求める
    CisDBGridH.LayOutGet

    gSL_Select = ""
    gSL_Select = gSL_Select & " Select * From   発注入荷状況照会ビュー "
    
    mSl_Select2 = ""
    mSl_Select2 = mSl_Select2 & " Select * From   受給検収問合ビュー "
' 分類内容(Order By句) ｾｯﾄ
    With CisDBGridH
        Call .OrderSet(True)
        If .Select_OrderBy = "" Then
            DBInput = False
            GoTo DBInput_Ed
        End If
    End With
    Call DBGridSort
' 抽出文　作成
    gSL_Select = gSL_Select & mSL_SelectW & mSL_Order
    
    mSl_Select2 = mSl_Select2 & CisDBGridH.Select_Where & CisDBGridH.Select_OrderBy
' グリッド内表示
    If Not DBGridDisp(DispType) Then
        DBInput = False
        GoTo DBInput_Ed
    End If
    On Error Resume Next
' 表示データなし
    On Error GoTo 0
    If Adodc1.Recordset.EOF And Adodc2.Recordset.EOF Then
        DBInput = False
        GoTo DBInput_Ed
    End If
    
'' 抽出件数表示
'    lb_Kensu.Visible = True
    
    On Error GoTo 0
DBInput_Ed:
    Call CisFrm.MousePT(0)
    
    PNL_Wait.Visible = False
    PNL_PF.Enabled = True
    Exit Function
DBInput_Err:
    Call CisFrm.MousePT(0)
    CisFun.ErrorBox
    PNL_Wait.Visible = False
    PNL_PF.Enabled = True
End Function
'+----------------------------------------------------------+
'+                                                          +
'+                                                          +
'+      グリッド処理(TrueDBGrid)                             +
'+                                                          +
'+                                                          +
'+----------------------------------------------------------+
'+----------------------------------------------------------+
'+   グリッド内容初期設定
'+----------------------------------------------------------+
Private Sub GridInitSet()
' Call .GridInitSet(ﾏｰｼﾞ処理,分類処理,ﾃﾞｰﾀ連結名,印刷見出し,印刷桁数,合計印刷)
    With CisDBGridH
         Call .GridInitSet(, , "編集受入年月日", "入 荷 日", 11)
         Call .GridInitSet(, , "納入番号", "納番", 8)
         Call .GridInitSet(, , "編集仕入先", "仕入先", 25)
         Call .GridInitSet(, , "表示品番", "品       番", 22)
         Call .GridInitSet(, , "入荷数", "入荷数", 8, False)
         Call .GridInitSet(, , "受入", "受", 3)
         Call .GridInitSet(, , "表示納入番号", "納  番", 10)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
         'Call .GridInitSet(, , "表示背番号", "背番号", 5)
         Call .GridInitSet(, , "表示背番号", "背番号", 8)
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
         Call .GridInitSet(, , "収容数", "収容数", 6, False)
         Call .GridInitSet(, , "手配区分名", "手配区", 9)
         Call .GridInitSet(, , "編集納入先", "納入先", 25)
         Call .GridInitSet(, , "SK区分名", "SK区分", 9)
         Call .GridInitSet(, , "置場", "置場", 11)
         Call .GridInitSet(, , "サイクル", "サイクル", 8)
         Call .GridInitSet(, , "編集契約先", "支  払  先", 25)
         Call .GridInitSet(, , "品名", "品        名", 40)
         Call .GridInitSet(, , "編集受信日", "受 信 日", 11)
         Call .GridInitSet(, , "伝票種別名", "伝 票 種 別", 20)
    End With
End Sub

'+----------------------------------------------------------+
'+      ADODC1(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub KensuDisp()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    lb_Kensu = ""
    NowCnt = Adodc1.Recordset.AbsolutePosition
    MaxCnt = Adodc1.Recordset.RecordCount
    If NowCnt < 0 Then Exit Sub
    lb_Kensu.Caption = Format(NowCnt, "###,##0") & " / " & _
                       Format(MaxCnt, "###,##0")

End Sub
'+----------------------------------------------------------+
'+      ADODC2(件数表示)                                     +
'+----------------------------------------------------------+
Private Sub KensuDisp2()
    Static NowCnt       As Long
    Static MaxCnt       As Long
    lb_Kensu2 = ""
    NowCnt = Adodc2.Recordset.AbsolutePosition
    MaxCnt = Adodc2.Recordset.RecordCount
    If NowCnt < 0 Then Exit Sub
    lb_Kensu2.Caption = Format(NowCnt, "###,##0") & " / " & _
                        Format(MaxCnt, "###,##0")

End Sub
'+---------------------+
'+  グリッド内容表示   +
'+---------------------+
Private Function DBGridDisp(DispType As String)
    
    DBGridDisp = False
    
    On Error GoTo DBGridDisp_ER:
    
    On Error Resume Next

    If DispType = "" Or DispType = "P" Then
        Adodc1.RecordSource = gSL_Select
        DBGrid1.ReBind
        Adodc1.Refresh
        DBGrid1.Refresh
    
        Call KensuDisp
        With DBGrid1
             .SetFocus
             SendKeys "{HOME}"
        End With
        
    End If
    

    If DispType = "" Or DispType = "H" Then
        Adodc2.RecordSource = mSl_Select2
        DBGridH.ReBind
        Adodc2.Refresh
        DBGridH.Refresh
    
        Call KensuDisp2
        
        With DBGridH
             .SetFocus
             SendKeys "{HOME}"
        End With
        
    End If
    
    DBGridDisp = True
    
    On Error GoTo 0
    Exit Function
DBGridDisp_ER:
    If Err <> 40002 Then        ' RDO タイムアウト
       CisFun.ErrorBox
    End If
End Function
'+-----------------------+
'+  グリッド内容初期化    +
'+-----------------------+
Private Sub DBGridClr()

    Call CisDBGridH.SplitsClr("", 0)   ' 分割ﾍﾟｲﾝ解除
    
    Adodc1.Recordset.Close
    DBGrid1.Close True
    
    Adodc2.Recordset.Close
    DBGridH.Close True
End Sub
Private Sub DBGridSort()
    Dim DBSpCols    As Object
    Dim DBSpCol     As Object
    Dim DBField()   As String
    Dim DBCols      As Long
    
    mSL_Order = ""
    
    DBCols = DBGrid1.Splits(0).Columns.Count
    ReDim DBField(DBCols)
    Set DBSpCols = DBGrid1.Splits(0).Columns
    For Each DBSpCol In DBSpCols
         With DBSpCol
             If .Visible Then
                DBField(.Order) = .DataField
             End If
         End With
    Next DBSpCol
    gStr = ""
    For gCnt1 = 0 To 5
        mSL_Order = mSL_Order & gStr & DBField(gCnt1)
        gStr = ","
    Next gCnt1
    
    If Trim(mSL_Order) <> "" Then
       mSL_Order = " Order By " & mSL_Order
    End If
End Sub
'+----------------------------------------------------------+
'+ キー入力にて行移動 《 KeyUp イベント   》                  +
'+----------------------------------------------------------+
Private Sub DBGrid1_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
        KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then
        Call KensuDisp
    Else
        Exit Sub
    End If
End Sub
Private Sub DBGrid1_GotFocus()
    B1_Area1.BackColor = gPGotSel
    B1_Area2.BackColor = gPLostSel
End Sub
Private Sub DBGridH_GotFocus()
    B1_Area1.BackColor = gPLostSel
    B1_Area2.BackColor = gPGotSel
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 ColMove イベント   》            +
'+----------------------------------------------------------+
Private Sub DBGrid1_ColMove(ByVal Position As Integer, Cancel As Integer)
    mGridMoveFlg = True
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 Paint イベント   》              +
'+----------------------------------------------------------+
Private Sub DBGrid1_Paint()
    If Not FormAct Then Exit Sub
    If mGridMoveFlg Then  ' 再表示(分類)
        mGridMoveFlg = False
        Call DBInput("", "P")
    End If
End Sub

'+----------------------------------------------------------+
'+      グリッド    《 MouseDown  イベント   》               +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    B1_Area1.BackColor = gPGotSel
    B1_Area2.BackColor = gPLostSel
'    If DBGrid1.RowContaining(Y) = -1 And _
'        DBGrid1.ColContaining(X) = -1 Then
'        Call CisDBGridP.SplitsClr("C", 0)          ' 分割ﾍﾟｲﾝのみ解除
'        Call DBInput("", "P")
'        Exit Sub
'    End If
'    If DBGrid1.RowContaining(Y) < 0 Then        ' 分割ﾍﾟｲﾝ表示
'        If Button = 2 Then
'            Call CisDBGridP.Splits(DBGrid1.ColContaining(X))
'            Call DBInput("", "P")
'        End If
'        Exit Sub
'    End If

    If DBGrid1.RowContaining(Y) < 0 Then Exit Sub
    If Button = 2 Then
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        Call TotalDsp(DBGrid1.ColContaining(X), "P")
    Else
        DBGrid1.Row = DBGrid1.RowContaining(Y)
        If PNL_ExeMsg.Visible Then
           PNL_ExeMsg.Visible = False
        End If
    End If

End Sub

'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGrid1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call KensuDisp
    If Button = 2 Then
       PNL_ExeMsgX = ""
       lb_THSu = "": lb_TNSu = "": lb_TNYSu = ""
       PNL_ExeMsg.Visible = False
    End If
End Sub

'+----------------------------------------------------------+
'+ キー入力にて行移動 《 KeyUp イベント   》                  +
'+----------------------------------------------------------+
Private Sub DBGridH_KeyUp(KeyCode As Integer, Shift As Integer)
    If (KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or _
        KeyCode = vbKeyPageUp Or KeyCode = vbKeyPageDown) Then
        Call KensuDisp2
    Else
        Exit Sub
    End If
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 ColMove イベント   》            +
'+----------------------------------------------------------+
Private Sub DBGridH_ColMove(ByVal Position As Integer, Cancel As Integer)
    CisDBGridH.ColMoveFlg = True
End Sub
'+----------------------------------------------------------+
'+ 表示項目移動後 分類表示 《 Paint イベント   》              +
'+----------------------------------------------------------+
Private Sub DBGridH_Paint()
    If Not FormAct Then Exit Sub
    If CisDBGridH.ColMoveFlg Then  ' 再表示(分類)
        CisDBGridH.ColMoveFlg = False
        Call DBInput("", "H")
    End If
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseDown  イベント   》               +
'+----------------------------------------------------------+
Private Sub DBGridH_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    B1_Area1.BackColor = gPLostSel
    B1_Area2.BackColor = gPGotSel
    If DBGridH.RowContaining(Y) = -1 And _
        DBGridH.ColContaining(X) = -1 Then
        Call CisDBGridH.SplitsClr("C", 0)          ' 分割ﾍﾟｲﾝのみ解除
        Call DBInput("", "H")
        Exit Sub
    End If
    If DBGridH.RowContaining(Y) < 0 Then        ' 分割ﾍﾟｲﾝ表示
        If Button = 2 Then
            Call CisDBGridH.SplitsAdd(DBGridH.ColContaining(X))
            Call DBInput("", "H")
        End If
        Exit Sub
    End If
    If DBGridH.RowContaining(Y) < 0 Then Exit Sub
    If Button = 2 Then
        DBGridH.Row = DBGridH.RowContaining(Y)
        Call TotalDsp(DBGridH.ColContaining(X), "H")
    Else
        DBGridH.Row = DBGridH.RowContaining(Y)
        If PNL_ExeMsg.Visible Then
           PNL_ExeMsg.Visible = False
        End If
    End If
End Sub
'+----------------------------------------------------------+
'+      グリッド    《 MouseUp  イベント   》                +
'+----------------------------------------------------------+
Private Sub DBGridH_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call KensuDisp2
    If Button = 2 Then
       PNL_ExeMsgX = ""
       lb_THSu = "": lb_TNSu = "": lb_TNYSu = ""
       PNL_ExeMsg.Visible = False
    End If
End Sub
'+----------------------+
'+      集計表示　       +
'+----------------------+
Private Sub TotalDsp(JCol As Long, SelGrid As String)
    Dim wWhere          As String
    Dim wMsg            As String
    Dim wValue          As String
    Dim wWhereH         As String
    Dim wCnt            As Long
    Dim wCntH           As Long
    
    
    wWhereH = "": wWhere = ""
    wCnt = 0: wCntH = 0
    If SelGrid = "H" Then
        With DBGridH
            Select Case JCol
              Case -1
                   wWhereH = "All"
                   wWhere = wWhereH
              Case 1, 2, 8, 9, 10, 11, 13
                   ' 仕入先(1),品番(2),手配区分(8),納入先(9),SK区分(10),置場(11),契約先(13)
                   wMsg = .Columns(JCol).Caption
                   wValue = .Columns(JCol).Text
                   wWhereH = .Columns(JCol).DataField & " = '" & RTrim(.Columns(JCol).Value) & "'"
                   wWhere = wWhereH
              Case 1            ' 仕入先･受入(1)
                   wMsg = "仕入先･受入"
                   wValue = .Columns(JCol - 1).Text & "-[" & .Columns(JCol).Text & "]"
                   If wValue = "-[]" Then wValue = ""
                   wWhereH = .Columns(JCol - 1).DataField & " + " & .Columns(JCol).DataField & " = '" & _
                            .Columns(JCol - 1).Text & .Columns(JCol).Text & "'"
                   wWhere = wWhereH
              Case 0            ' 入荷日(4)
                   wMsg = .Columns(JCol).Caption
                   wValue = .Columns(JCol).Text
                   wWhereH = .Columns(JCol).DataField & " = '" & RTrim(.Columns(JCol).Value) & "'"
                   wWhere = "Substring(納入日便,1,10) = '" & RTrim(.Columns(JCol).Value) & "'"
              Case 15, 16           ' 受信日(15),伝票種別(16)
                   wMsg = .Columns(JCol).Caption
                   wValue = .Columns(JCol).Text
                   wWhereH = .Columns(JCol).DataField & " = '" & RTrim(.Columns(JCol).Value) & "'"
                   wWhere = ""
               Case Else
                    Exit Sub
            End Select
        End With
    Else
        With DBGrid1
            Select Case JCol
              Case -1
                   wWhere = "All"
                   wWhereH = wWhere
              Case 1, 3, 15, 17, 19, 20, 22
                   ' 仕入先(0),品番(3),手配区分(15),納入先(17),SK区分(19),置場(20),契約先(22)
                   wMsg = .Columns(JCol).Caption
                   wValue = .Columns(JCol).Text
                   wWhere = .Columns(JCol).DataField & " = '" & RTrim(.Columns(JCol).Value) & "'"
                   wWhereH = wWhere
              Case 1            ' 仕入先･受入(1)
                   wMsg = "仕入先･受入"
                   wValue = .Columns(JCol - 1).Text & "-[" & .Columns(JCol).Text & "]"
                   If wValue = "-[]" Then wValue = ""
                   wWhere = .Columns(JCol - 1).DataField & " + " & .Columns(JCol).DataField & " = '" & _
                            .Columns(JCol - 1).Text & .Columns(JCol).Text & "'"
                   wWhereH = wWhere
              Case 0            ' 納入日(4)
                   wMsg = .Columns(JCol).Caption
                   wValue = .Columns(JCol).Text
                   wWhere = .Columns(JCol).DataField & " = '" & RTrim(.Columns(JCol).Value) & "'"
                   wWhereH = "編集受入年月日 = '" & Mid(RTrim(.Columns(JCol).Value), 1, 10) & "'"
               Case Else
                    Exit Sub
            End Select
        End With
    End If
    wMsg = Replace(wMsg, "　", "")
    wMsg = Replace(wMsg, " ", "")
    If JCol < 0 Then
        PNL_ExeMsgX = "【 集計単位 ＝ 全件対象 】"
        lb_TMei.Caption = ""
    Else
        PNL_ExeMsgX = "【 集計単位 ＝ " & wMsg & " 】"
        lb_TMei.Caption = wMsg & " ： " & wValue
    End If
    
    gSL_Select = "Select Sum(Case When 分納区分 = 0 Then 発注数 Else 0 End ),"
    gSL_Select = gSL_Select & " Count(*) "
    gSL_Select = gSL_Select & " From 発注入荷状況照会ビュー "
    
    If wWhere <> "" Then
        If wWhere <> "All" Then
            If Trim(mSL_SelectW) = "" Then
               gSL_Select = gSL_Select & " Where " & wWhere
            Else
               gSL_Select = gSL_Select & mSL_SelectW
               gSL_Select = gSL_Select & " And " & wWhere
            End If
        Else
            gSL_Select = gSL_Select & mSL_SelectW
        End If
    
        With CisDB
            .SQL = gSL_Select
            If .DBRead(1, 0) Then
                lb_THSu = Format(.Recordset(0), "###,##0")
                wCnt = CisFun.Val2(.Recordset(1))
            End If
            On Error Resume Next
                .ConnectNo = 0
                .Recordset.Close
            On Error GoTo 0
        End With
    End If
    
    gSL_Select = "Select Sum(入荷数),"
    gSL_Select = gSL_Select & " Count(*) "
    gSL_Select = gSL_Select & " From 受給検収問合ビュー "
    If wWhereH <> "" Then
        If wWhereH <> "All" Then
            If Trim(CisDBGridH.Select_Where) = "" Then
               gSL_Select = gSL_Select & " Where " & wWhereH
            Else
               gSL_Select = gSL_Select & CisDBGridH.Select_Where
               gSL_Select = gSL_Select & " And " & wWhereH
            End If
        Else
            gSL_Select = gSL_Select & CisDBGridH.Select_Where
        End If
        
        With CisDB
            .SQL = gSL_Select
            If .DBRead(1, 0) Then
                
                lb_TNSu = Format(.Recordset(0), "###,##0")
                wCntH = CisFun.Val2(.Recordset(1))
            End If
            On Error Resume Next
                .ConnectNo = 0
                .Recordset.Close
            On Error GoTo 0
        End With
    End If
    
    lb_TNYSu = Format(CisFun.Val2(lb_THSu) - CisFun.Val2(lb_TNSu), "###,##0")
    
    lb_SKensu = "【 集計件数 ＝ " & Format(wCnt, "#,##0") & " | " & Format(wCntH, "#,##0") & " 】"
    
    PNL_ExeMsg.Visible = True
'
End Sub

