VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "VSFLEX3.OCX"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Object = "{1BA4D453-0150-11CE-89B0-0000C037528B}#3.2#0"; "PDQCom32.OCX"
Object = "{490AF35D-478A-11D4-AEC2-00000EA57F9E}#4.0#0"; "CISMSG~2.OCX"
Begin VB.Form CMJ0010 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "前工程発行JAMA 入荷実績入力"
   ClientHeight    =   11115
   ClientLeft      =   1305
   ClientTop       =   3105
   ClientWidth     =   15240
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
   ForeColor       =   &H80000008&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11115
   ScaleWidth      =   15240
   WindowState     =   2  '最大化
   Begin VB.PictureBox PNL_Start 
      BackColor       =   &H00000000&
      Height          =   825
      Left            =   390
      ScaleHeight     =   765
      ScaleWidth      =   8235
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   525
      Visible         =   0   'False
      Width           =   8295
      Begin CisMsgRelate_v60.CISMsgRelate ProcMsg 
         Height          =   645
         Left            =   75
         Top             =   60
         Width           =   8085
         _ExtentX        =   14261
         _ExtentY        =   1138
         BackColor       =   65535
         Caption         =   "ホッパリーダの【ｽﾀｰﾄ】キーを押して下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ 明朝"
            Size            =   15.75
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         mrInterval      =   50
         mrStopCnt       =   40
      End
   End
   Begin Cis3D_v60.CIS3D H6_Area1 
      Height          =   1320
      Left            =   465
      Top             =   5415
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2328
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
      Begin VB.TextBox QRRDummy 
         Height          =   330
         Left            =   3645
         TabIndex        =   8
         Text            =   "Text1"
         Top             =   30
         Width           =   240
      End
      Begin Cis3D_v60.CIS3D QRRMsg 
         Height          =   1170
         Left            =   90
         Top             =   75
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   2064
         BackColor       =   16777152
         ForeColor       =   16711680
         Caption         =   "QRﾎｯﾊﾟﾘｰﾀﾞ読取"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   7
         cAlingnment     =   7
      End
   End
   Begin PdqcommLib.PDQComm PDQComm1 
      Height          =   480
      Left            =   1260
      TabIndex        =   7
      Top             =   660
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.26
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FastScroll      =   0   'False
      Object.Height          =   0
      Object.Width           =   0
   End
   Begin Cis3D_v60.CIS3D PNL_Wait 
      Height          =   930
      Left            =   4830
      Top             =   4290
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
   Begin Cis3D_v60.CIS3D CIS3D1 
      Height          =   885
      Left            =   7740
      Top             =   1890
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   1561
      BackColor       =   0
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
         Left            =   45
         Top             =   45
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   714
         Caption         =   "表示枚数"
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
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_Kensu 
            Height          =   345
            Left            =   1215
            Top             =   30
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   609
            BackColor       =   16777152
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D14 
         Height          =   405
         Left            =   2250
         Top             =   45
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   714
         Caption         =   "更新選択枚数"
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
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_KKensu 
            Height          =   375
            Left            =   1545
            Top             =   15
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   661
            BackColor       =   12648384
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D12 
         Height          =   405
         Left            =   2250
         Top             =   450
         Width           =   2520
         _ExtentX        =   4445
         _ExtentY        =   714
         ForeColor       =   0
         Caption         =   "未選択枚数"
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
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_MKensu 
            Height          =   330
            Left            =   1545
            Top             =   30
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   582
            BackColor       =   12648384
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   405
         Left            =   4770
         Top             =   45
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   714
         Caption         =   "更新済枚数"
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
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D Tlb_UpdCnt 
            Height          =   345
            Left            =   1410
            Top             =   30
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   609
            BackColor       =   16761024
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   405
         Left            =   45
         Top             =   450
         Width           =   2205
         _ExtentX        =   3889
         _ExtentY        =   714
         ForeColor       =   0
         Caption         =   "ｴﾗｰ含枚数"
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
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B1lb_EKensu 
            Height          =   330
            Left            =   1215
            Top             =   30
            Width           =   885
            _ExtentX        =   1561
            _ExtentY        =   582
            BackColor       =   16761087
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D17 
         Height          =   405
         Left            =   4770
         Top             =   450
         Width           =   2400
         _ExtentX        =   4233
         _ExtentY        =   714
         Caption         =   "ｴﾗｰ更新枚数"
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
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D Tlb_ErrCnt 
            Height          =   330
            Left            =   1410
            Top             =   45
            Width           =   900
            _ExtentX        =   1588
            _ExtentY        =   582
            BackColor       =   16761087
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS UI Gothic"
               Size            =   15.75
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
            cPositionX      =   -30
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
      Height          =   750
      Left            =   8730
      ScaleHeight     =   690
      ScaleWidth      =   6150
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   525
      Width           =   6210
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   3525
         Top             =   45
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
         Left            =   4380
         Top             =   45
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
         Left            =   5235
         Top             =   45
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
         Left            =   60
         Top             =   45
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         Enabled         =   -1  'True
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
      Begin CisBtn_60.CisBtn PB_CHG 
         Height          =   585
         Left            =   1815
         Top             =   45
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
         BtnType         =   10
         BtnCaption      =   "切替"
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
         PFCaption       =   "( F6 )"
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
      Begin CisBtn_60.CisBtn PB_UPD 
         Height          =   585
         Left            =   60
         Top             =   45
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
      Begin CisBtn_60.CisBtn PB_MEI 
         Height          =   585
         Left            =   1815
         Top             =   45
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
         BtnForeColor    =   16711935
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
         PFBackColor     =   16744703
      End
      Begin CisBtn_60.CisBtn PB_DEL 
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
      Begin CisBtn_60.CisBtn PB_Bun 
         Height          =   585
         Left            =   930
         Top             =   45
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
         BtnCaption      =   "個別"
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
      Begin CisBtn_60.CisBtn PB_Print 
         Height          =   585
         Left            =   2670
         Top             =   45
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
   Begin Cis3D_v60.CIS3D H4_Area1 
      Height          =   1320
      Left            =   3240
      Top             =   1455
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2328
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
      Begin VB.TextBox QRDummy 
         Height          =   330
         Left            =   3660
         TabIndex        =   1
         Text            =   "Text1"
         Top             =   30
         Width           =   240
      End
      Begin Cis3D_v60.CIS3D QRMsg 
         Height          =   1170
         Left            =   90
         Top             =   75
         Width           =   3915
         _ExtentX        =   6906
         _ExtentY        =   2064
         BackColor       =   12648384
         ForeColor       =   16711680
         Caption         =   "QRﾀｯﾁﾘｰﾀﾞ読取"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   18
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cFont3DColor1   =   16777215
         cFont3DColor2   =   8421504
         cBoderWidth     =   7
         cAlingnment     =   7
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
      Height          =   1275
      Left            =   390
      ScaleHeight     =   1215
      ScaleWidth      =   2625
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   1455
      Width           =   2685
      Begin Cis3D_v60.CIS3D CIS3D13 
         Height          =   1110
         Left            =   60
         Top             =   60
         Width           =   2505
         _ExtentX        =   4419
         _ExtentY        =   1958
         ForeColor       =   16711680
         Caption         =   "検 収 日"
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_NYmd 
            Height          =   360
            Left            =   285
            TabIndex        =   0
            Top             =   450
            Width           =   1860
            _ExtentX        =   3069
            _ExtentY        =   582
            cCalender       =   -1  'True
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
   Begin VB.PictureBox B1_Area1 
      Height          =   7725
      Left            =   390
      ScaleHeight     =   7665
      ScaleWidth      =   14505
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2850
      Width           =   14565
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   7560
         Left            =   60
         Top             =   45
         Width           =   14415
         _ExtentX        =   25426
         _ExtentY        =   13335
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
            Height          =   7350
            Left            =   60
            TabIndex        =   2
            TabStop         =   0   'False
            Top             =   105
            Width           =   14295
            _Version        =   196608
            _ExtentX        =   25215
            _ExtentY        =   12965
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
            FormatString    =   "№|処理 |納　期 　　-便|仕入先 |受入|出荷場|納入番号|頁|支給元1|支給元2|発区|ｶｰﾄﾞ|発行日|発注者|S納入日|S納入便|S頁|Select|||"
            Rows            =   24
            Cols            =   21
            BackColor       =   12648447
            BackColorBkg    =   14737632
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   300
            GridLines       =   3
            SelectionMode   =   1
            BackColorAlternate=   12648447
            Begin Cis3D_v60.CIS3D B2_Area1 
               Height          =   4695
               Left            =   6705
               Top             =   2085
               Visible         =   0   'False
               Width           =   7245
               _ExtentX        =   12779
               _ExtentY        =   8281
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
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   285
                  Left            =   120
                  Top             =   225
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "納 入 日"
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
               Begin vsFlexLib.vsFlexArray vsGrid2 
                  Height          =   3735
                  Left            =   135
                  TabIndex        =   11
                  TabStop         =   0   'False
                  Top             =   825
                  Width           =   6990
                  _Version        =   196608
                  _ExtentX        =   12330
                  _ExtentY        =   6588
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
                  FormatString    =   "№|品　番　　　   　|背番号|収容数 |箱数 |納入数  |エラー|"
                  Rows            =   11
                  Cols            =   8
                  BackColor       =   16777152
                  BackColorBkg    =   14737632
                  FocusRect       =   0
                  HighLight       =   2
                  ScrollBars      =   0
                  RowHeightMin    =   330
                  GridLines       =   3
                  SelectionMode   =   1
                  BackColorAlternate=   16777152
                  Begin Cis3D_v60.CIS3D CIS3D5 
                     Height          =   6270
                     Left            =   7050
                     Top             =   660
                     Width           =   6360
                     _ExtentX        =   11218
                     _ExtentY        =   11060
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
                  End
               End
               Begin Cis3D_v60.CIS3D B2lb_NYmd 
                  Height          =   285
                  Left            =   120
                  Top             =   510
                  Width           =   1785
                  _ExtentX        =   3149
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "9999/99/99-99"
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
               Begin Cis3D_v60.CIS3D CIS3D8 
                  Height          =   285
                  Left            =   1905
                  Top             =   225
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "納入番号"
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
               Begin Cis3D_v60.CIS3D B2lb_Nouno 
                  Height          =   285
                  Left            =   1905
                  Top             =   510
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "XXXXX"
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
               Begin Cis3D_v60.CIS3D CIS3D10 
                  Height          =   285
                  Left            =   3105
                  Top             =   225
                  Width           =   480
                  _ExtentX        =   847
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "頁"
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
               Begin Cis3D_v60.CIS3D B2lb_Page 
                  Height          =   285
                  Left            =   3105
                  Top             =   510
                  Width           =   480
                  _ExtentX        =   847
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "99"
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
               Begin Cis3D_v60.CIS3D CIS3D16 
                  Height          =   285
                  Left            =   3585
                  Top             =   225
                  Width           =   930
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "受注者"
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
               Begin Cis3D_v60.CIS3D B2lb_Torcd 
                  Height          =   285
                  Left            =   3585
                  Top             =   510
                  Width           =   930
                  _ExtentX        =   1640
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "XXXXX"
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
               Begin Cis3D_v60.CIS3D CIS3D18 
                  Height          =   285
                  Left            =   4515
                  Top             =   225
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "受入"
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
               Begin Cis3D_v60.CIS3D B2lb_Ukeir 
                  Height          =   285
                  Left            =   4515
                  Top             =   510
                  Width           =   615
                  _ExtentX        =   1085
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "99"
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
               Begin Cis3D_v60.CIS3D CIS3D20 
                  Height          =   285
                  Left            =   5130
                  Top             =   225
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "支給元1"
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
               Begin Cis3D_v60.CIS3D B2lb_SMoto 
                  Height          =   285
                  Left            =   5130
                  Top             =   510
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "XXXXX"
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
               Begin Cis3D_v60.CIS3D CIS3D7 
                  Height          =   285
                  Left            =   6135
                  Top             =   225
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "支給元2"
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
               Begin Cis3D_v60.CIS3D B2lb_SMoto2 
                  Height          =   285
                  Left            =   6135
                  Top             =   510
                  Width           =   1005
                  _ExtentX        =   1773
                  _ExtentY        =   503
                  BackColor       =   14737632
                  Caption         =   "XXXXX"
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
         End
      End
   End
   Begin VB.TextBox Dummy 
      Height          =   330
      IMEMode         =   3  'ｵﾌ固定
      Left            =   90
      TabIndex        =   3
      Text            =   "Dummy"
      Top             =   420
      Width           =   240
   End
   Begin VB.Timer Timer1 
      Left            =   900
      Top             =   -1380
   End
   Begin VB.Timer Timer2 
      Left            =   1170
      Top             =   -1380
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15390
      _ExtentX        =   27146
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 前工程発行JAMA 入荷実績入力 [一括)]  】"
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
      cPositionX      =   3000
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   375
         Left            =   0
         Top             =   0
         Width           =   2865
         _ExtentX        =   5054
         _ExtentY        =   661
         BackColor       =   0
         ForeColor       =   65535
         Caption         =   "XXXX"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   15.75
            Charset         =   128
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor2    =   12632256
         cFont3DColor1   =   8421504
         cFont3DColor2   =   -2147483643
         cBoderWidth     =   3
         cFont3D         =   2
         cAlingnment     =   7
      End
      Begin Cis3D_v60.CIS3D Pgid 
         Height          =   255
         Left            =   14340
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
         Left            =   11850
         Top             =   45
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
   Begin Cis3D_v60.CIS3D ErrMsg 
      Height          =   465
      Left            =   3315
      Top             =   945
      Visible         =   0   'False
      Width           =   3810
      _ExtentX        =   6720
      _ExtentY        =   820
      BackColor       =   16761087
      Caption         =   "エラー"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ 明朝"
         Size            =   14.25
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
      cPositionX      =   90
   End
   Begin PdqcommLib.PDQComm PDQComm2 
      Height          =   480
      Left            =   1785
      TabIndex        =   9
      Top             =   630
      Width           =   480
      _Version        =   196610
      _ExtentX        =   847
      _ExtentY        =   847
      _StockProps     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "ＭＳ ゴシック"
         Size            =   11.26
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FastScroll      =   0   'False
      Object.Height          =   0
      Object.Width           =   0
   End
End
Attribute VB_Name = "CMJ0010"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   前工程発行JAMA 入荷実績入力
'**       フォームID    :   CMJ0010
'**       処理概要      :
'**
'**       作  成  日    :   2007/07/27  By CIS      (ﾎｯﾊﾟｰﾘｰﾀﾞｰのテスト未）
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim CisVsGrid               As New CisVsGrid3
    Dim CisVsGrid2              As New CisVsGrid3
    Dim SL_Select               As String
    Dim mSL_Select              As String
    Dim mSL_Group               As String
    Dim mSL_Order               As String
    
    Dim mHeadProc               As String
    
    Dim mKanCnt                 As Long
    Dim mBunCnt                 As Long
    Dim mMiCnt                  As Long
    
    Dim QRReadFlg               As Boolean
    Dim CheckFlg                As Boolean
'+----------------------------------------------------------+
'+      フォーム    《    Ｌｏａｄイベント    》              +
'+----------------------------------------------------------+
Private Sub Form_Load()
    If Rv_NCall = "" Then
    '   #-------------------#
    '   # 二 重 起 動 防 止  #
    '   #-------------------#
        If Not Dupli_Start Then End
    '   #-------------------#
    '   # 初 期 内 容 取 得  #
    '   #-------------------#
        If Not IniGet Then End
    End If
    
    Set CisFrm = New CISFormContorl
    Set CisFrm.ThisForm = Me
    Dummy.Left = -1000
    Dummy.Enabled = False
    QRDummy.Left = -1000
    QRRDummy.Left = -1000
    
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 600    '
    Timer2.Enabled = True    '

    H6_Area1.Move H4_Area1.Left, H4_Area1.Top

    Set QRr_Object = PDQComm1
    Set QRr_MsgObject = QRMsg
    Set QRr_ErrMsgObj = ErrMsg
    Set TX2400_StartMsg = PNL_Start
    Set TX2400_ReadObj = QRRMsg
    Set TX2400_ErObj = ErrMsg
    '   ( QR Reader 環境 )
    Call QRInit
    Call TX2400InfoGet
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = VsGrid1
        .NoSpaceCol = 2
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 19
        .InitGet = False
        .Init
    End With
    Call GridInit
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid2
        Set .GridObj = vsGrid2
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
'        .SelectCol = 19
        .InitGet = False
        .Init
    End With
    Call Grid2Init
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

    
    K_Sykbnm.Tag = "0"
    Call PB_CHG_Click

'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    If Rv_NCall = "" Then
       Me.Show
    End If
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    
    Call ErrorDataCheck(False)
'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)

    FormAct = Not FormAct
End Sub
'+----------------------------------------------------------+
'+      フォーム    《 QueryUnloadイベント 》                +
'+----------------------------------------------------------+
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If Rv_NCall = "" Then
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
        Case vbKeyF1:       If PB_UPD.Visible Then Call PB_UPD_Click        '【更新】
        Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click        '【更新】
        Case vbKeyF4:       If PB_Bun.Visible Then Call PB_Bun_Click        '【個別】
        Case vbKeyF6:       If PB_CHG.Visible Then Call PB_CHG_Click        '【切替】
        Case vbKeyF7:       If PB_Look.Visible Then Call PB_Look_Click      '【検索】
        Case vbKeyF8:       If PB_Print.Visible Then Call PB_Print_Click    '【印刷】
        Case vbKeyF10:      If PB_MEI.Visible Then Call PB_MEI_Click        '【詳細表示】
        Case vbKeyEscape:   If PB_CAN.Visible Then Call PB_Can_Click        '【取消】
        Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click        '【終了】
        Case vbKeyReturn
                            If PB_ENT.Visible Then                          '【入力】
                                Call PB_ENT_Click
                            End If
        Case Else:          Exit Function
    End Select
    
    Key_Acc = True
End Function
'+---------------------+
'+    終  了  キ ー     +
'+---------------------+
Private Sub PB_END_Click()
        
    Call CisDB.DBDISConnect                         ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
    End
End Sub
'+----------------------------+
'+    入  力  キ ー(Enter)     +
'+----------------------------+
Private Sub PB_ENT_Click()
    Dummy.Enabled = True: Dummy.SetFocus: DoEvents
    Call ReturnPress
    Dummy.Enabled = False
End Sub
'+----------------------------+
'+    取  消  キ ー(Escape)    +
'+----------------------------+
Private Sub PB_Can_Click()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H4" Or ProcHB = "H6" Then
        If CisFun.Val2(B1lb_KKensu) <> 0 Or CisFun.Val2(B1lb_MKensu) <> 0 Then
            With CisFun
                .MB_Lines = 4
                .MB_MSG(2) = "    未更新の伝票があります。         "
                .MB_MSG(4) = "    消去してもよろしいですか？         "
                .MB_Button = OK_CAN
                If Not .MBOX Then
                    VsGrid1.SetFocus
                    GoTo PB_CAN_Ed
                End If
            End With
        End If
        ProcHB = "H1"
        Call HeadBodyClear("B")
        Call DispChange(ProcHB)
        If B2_Area1.Visible Then B2_Area1.Visible = False
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+---------------------+
'+    一括画面切替     +
'+---------------------+
Private Sub PB_Bun_Click()
    Rv_NCall = "K"
    Rv_NTtlSu = 0
        
    Timer1.Enabled = False
    Call QR_Comm_Close
    Call TX2400Close
    
    CMJ0020.Show vbModal
    Unload CMJ0020
    Set CMJ0020 = Nothing
    
    Timer1.Enabled = True
    Rv_NCall = ""
    
    '   ( QR Reader 環境 )
    Set QRr_Object = PDQComm1
    Set QRr_MsgObject = QRMsg
    Set QRr_ErrMsgObj = ErrMsg
    Set TX2400_StartMsg = PNL_Start
    Set TX2400_ReadObj = QRRMsg
    Set TX2400_ErObj = ErrMsg
    Call QRInit
    Call TX2400InfoGet
    Call DispChange(ProcHB)
    
End Sub
'+------------------+
'+    切替   (F6)   +
'+------------------+
Private Sub PB_CHG_Click()
    gLong = CisFun.Val2(K_Sykbnm.Tag)
    
    gLong = gLong + 1
    If gLong > 2 Then gLong = 1
    
    gStr = ""
    If IRN_NKRead = 1 Then
       gStr = gStr & "H4"
    Else
       gStr = gStr & "  "
    End If
    If IRN_NKReadHP = 1 Then
       gStr = gStr & "H6"
    Else
       gStr = gStr & "  "
    End If
    
    mHeadProc = ""
    For gInt = 1 To 2
        mHeadProc = Mid(gStr, gLong + (gLong - 1), 2)
        If Trim(mHeadProc) <> "" Then
           K_Sykbnm.Tag = gLong
           Exit For
        End If
        gLong = gLong + 1
        If gLong > 2 Then gLong = 1
    Next gInt
    
    H4_Area1.Visible = False
    H6_Area1.Visible = False
    
    Select Case mHeadProc
        Case "H4"
            K_Sykbnm.Caption = "【 QRﾀｯﾁﾘｰﾀﾞ 】"
            H4_Area1.Visible = True
        Case "H6"
            K_Sykbnm.Caption = "【QRﾎｯﾊﾟﾘｰﾀﾞ】"
            H6_Area1.Visible = True
        Case Else
            K_Sykbnm.Caption = "【 QRﾀｯﾁﾘｰﾀﾞ 】"
            H4_Area1.Visible = True
            QRMsg = "QR未接続"
    End Select
    If ProcHB <> "H1" Then
       ProcHB = mHeadProc
       Call DispChange(ProcHB)
    End If
End Sub
'+------------------+
'+    詳細表示 (F8)   +
'+------------------+
Private Sub PB_MEI_Click()
    If B2_Area1.Visible Then
        B2_Area1.Visible = False
    Else
        Call Grid2Set
        B2_Area1.Visible = True
    End If
    VsGrid1.SetFocus
End Sub
'+----------------------------+
'+    エラーリスト印刷 (F8)   +
'+----------------------------+
Private Sub PB_Print_Click()
    Call Error_List
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_NYmd" Then iH1_NYmd.ShowCalender
End Sub

'*******************************
'*      タイトルバー編集        *
'*******************************
Private Sub Timer1_Timer()
    CisFrm.TimeSet CisFun.INI_FileID, 110
End Sub
'*--------------------------------------*
'*      入 荷 日    ( iH1_NYmd )        *
'*--------------------------------------*
Private Sub iH1_NYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_NYmd_LostFocus()
    If PB_Look.Tag <> ActiveControl.Name Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
        ProcHB = mHeadProc
        Call DispChange(ProcHB)
        VsGrid1.SetFocus
        GoTo ReturnPress_Ed
    End If
    
    If ProcHB = "H4" Then
       If Not Head4Chk Then GoTo ReturnPress_Ed
       GoTo ReturnPress_Ed
    End If
    
ReturnPress_Ed:
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
    Dim H1Mode          As Boolean
    Dim H1Color         As Long
    Dim H2Mode          As Boolean
    Dim H2Color         As Long
    Dim H3Mode          As Boolean
    Dim H3Color         As Long
    Dim H4Mode          As Boolean
    Dim H4Color         As Long
    Dim H5Mode          As Boolean
    Dim H5Color         As Long
    Dim H6Mode          As Boolean
    Dim H6Color         As Long
    Dim B1Mode          As Boolean
    Dim B1Color         As Long
    
    H1Mode = False:     H1Color = gPLostSel
    H2Mode = False:     H2Color = gPLostSel
    H3Mode = False:     H3Color = gPLostSel
    H4Mode = False:     H4Color = gPLostSel
    H5Mode = False:     H5Color = gPLostSel
    H6Mode = False:     H6Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "H4":      H4Mode = True:  H4Color = gPGotSel
        Case "H6":      H6Mode = True:  H6Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    H4_Area1.Enabled = H4Mode
    H4_Area1.BackColor = H4Color
    H6_Area1.Enabled = H6Mode
    H6_Area1.BackColor = H6Color
    
    B1_Area1.Enabled = B1Mode Or H4Mode Or H6Mode
    B1_Area1.BackColor = B1Color
    
    PB_CAN.Visible = Not H1Mode
    If IRN_NKReadHP = 0 Then
        PB_CHG.Visible = False
    Else
        PB_CHG.Visible = H1Mode
    End If
    PB_Look.Visible = H1Mode
    PB_Bun.Visible = H1Mode
    PB_END.Visible = H1Mode
    PB_Print.Visible = H1Mode
    
    PB_UPD.Visible = B1Mode
    PB_DEL.Visible = B1Mode
    If H4Mode Or H6Mode Then
        If mKanCnt <> 0 Then
            PB_UPD.Visible = True
            PB_DEL.Visible = True
        End If
    End If
    
    PB_MEI.Visible = H4Mode Or H6Mode
    PB_ENT.Visible = Not (H4Mode Or H6Mode)
    
    If ProcHB = "H1" Then
        B1lb_Kensu = ""
        B1lb_KKensu = ""
        B1lb_MKensu = ""
        B1lb_EKensu = ""
        PB_Look.Visible = False
    End If
'
    Set QRr_Object = PDQComm1
    Set TX2400_Comm = PDQComm2
    Call QR_Comm_Close
    Call TX2400Close
    If ProcHB = "H4" Then
       Call QR_Comm_Open
    End If
    If ProcHB = "H6" Then
       If TX2400_ReadAuto Then
          ProcMsg.Caption = "用紙をホッパーにセットして下さい。"
       Else
          ProcMsg.Caption = "用紙をﾎｯﾊﾟｰにｾｯﾄし【ｽﾀｰﾄ】キーを押して下さい。"
       End If
       PNL_Start.Tag = ""
       Call TX2400OpenAuto
    End If
    PNL_Start.Visible = H6Mode
    Call TX2400QRRReadMsg("")

' ﾀﾞﾐｰ項目(入力禁止)
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
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    If HeadBody = "B" Then
       CisVsGrid.Clear
    End If
End Sub
'****************************
'*      HEAD1ﾁｪｯｸ処理
'****************************
Private Function Head1Chk() As Boolean
    Head1Chk = False
    
'( QR設定 )
    If IRN_NKRead = 0 And IRN_NKReadHP = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    QRリーダーの設定がされていません！         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_NYmd.SetFocus
        Exit Function
    End If
'( 検収日 )
    If RTrim$(iH1_NYmd) <> "" And Not iH1_NYmd.cChkResult Then
         With CisFun
             .MB_Lines = 4
             .MB_MSG(2) = "    正しい日付を入力して下さい         "
             .MB_Button = Error
             .MBOX
         End With
         iH1_NYmd.SetFocus
         Exit Function
    End If
    
    If RTrim$(iH1_NYmd) <> "" Then
        gInt = Kdou_Check(iH1_NYmd)
        If gInt = 1 Then
        '休日
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(1) = "     検収日が非稼働日です。     "
            CisFun.MB_MSG(3) = "     　よろしいですか？      "
            CisFun.MB_Button = OK_CAN
            If Not CisFun.MBOX Then
                iH1_NYmd.SetFocus
                Exit Function
            End If
        End If
        If gInt = 2 Then
        'カレンダ未登録
            CisFun.MB_Lines = 4
            CisFun.MB_MSG(1) = "     " & RTrim(iH1_NYmd.cYear) & "年" & RTrim(iH1_NYmd.cMonth) & "月のカレンダ未登録！     "
            CisFun.MB_MSG(3) = "     　　(　カレンダマスタ　)　"
            CisFun.MB_Button = Error
            CisFun.MBOX
            iH1_NYmd.SetFocus
            Exit Function
        End If
    End If
    ' 入荷実績処理ワーク削除
    With CisDB
        .SQL = "DELETE FROM 前工程入荷ワーク"
        .SQL = .SQL & "    Where 処理端末 = substring(HOST_NAME(),1,20)"
        .SQL = .SQL & "      AND 処理区分 = 0"
        .DBExec
    End With
    
    ' 表示件数カウント初期化
    mKanCnt = 0
    mBunCnt = 0
    mMiCnt = 0

'エラー退避データ表示
    Call CisVsGrid.Clear
    Call HeadBodyClear("B")
    Call ErrorDataCheck(True)
    
    
    Head1Chk = True
End Function
'*-----------------------------------------*
'*      HEAD4ﾁｪｯｸ処理 ( ＱＲ )             *
'*-----------------------------------------*
Private Function Head4Chk() As Boolean
    Head4Chk = False
    
    
    Head4Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Body1Chk = False
    
'    If iB1_???? = "" Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    ???を入力して下さい         "
'            .MB_Button = Error
'            .MBOX
'        End With
'        iB1_???.SetFocus
'        Exit Function
'    End If
    
'   < 印刷データ確認 >
'    gSL_Select = "SELECT * FROM ??テーブル"
'    gSL_Select = gSL_Select & " WHERE Key = '" & RTrim(iH1_NYmd) & "'"
    
'    If Not NKTRead(gSL_Select, 1) Then
'        With CisFun
'            .MB_Lines = 4
'            .MB_MSG(2) = "    該当データがありません            "
'            .MB_Button = Error
'            .MBOX
'        End With
''        iB1_???.SetFocus
'        Exit Function
'    End If
    
    Body1Chk = True
End Function
'****************************
'*      ボディ内容セット
'****************************
Private Function BodySet()

End Function
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆
'☆                         入　　荷　　実　　績　　更　　新　　(  F1  )
'☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
Private Sub PB_UPD_Click()
    
    gCnt1 = 0
    With VsGrid1
        For gCnt2 = 0 To .Rows - 1
            If Trim(.TextMatrix(gCnt2, 1)) = "更新" Then
               gCnt1 = 1
               Exit For
            End If
        Next gCnt2
    End With
        
    If gCnt1 = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "      更新行を選択して下さい               "
            .MB_Button = OK_CAN
            .MBOX
            Exit Sub
        End With
    End If
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "      入荷実績更新を行います。               "
        .MB_MSG(4) = "          よろしいですか？                      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            VsGrid1.SetFocus
            Exit Sub
        End If
    End With

    On Error GoTo PB_UPD_Er:
    
    If B2_Area1.Visible Then B2_Area1.Visible = False
    PNL_Wait.Visible = True
    Message = "【実績更新中】しばらくお待ち下さい"
    DoEvents
    '+---------------
    '+   完納処理
    '+---------------
    Set QRr_Object = PDQComm1
    Set TX2400_Comm = PDQComm2
    Call QR_Comm_Close
    Call TX2400Close
    
    Call CisDB.DBTran(TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With VsGrid1
        .Redraw = False
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) = "更新" Then
               With CisDB
                    
                '更新
                    .SQL = "前工程入荷データチェック"
                    .StoadoCount = 7
                    .ParaNo = 0: .ParaIO = Return用
                    .ParaNo = 1: .ParaIO = Input用           '処理区分(0:ﾁｪｯｸ、1:更新）
                    .ParaNo = 2: .ParaIO = Input用           '納入番号
                    .ParaNo = 3: .ParaIO = Input用           '納入日
                    .ParaNo = 4: .ParaIO = Input用           '納品書頁
                    .ParaNo = 5: .ParaIO = OutPut用          '作成件数
                    .ParaNo = 6: .ParaIO = OutPut用          'エラー件数
                    .ParaNo = 7: .ParaIO = OutPut用          'エラーコード
            '
                    .ParaNo = 1: .ParaValue = 1              '処理区分(0:ﾁｪｯｸ、1:更新）
                    .ParaNo = 2: .ParaValue = RTrim(VsGrid1.TextMatrix(gLong, 6))   '納入番号
                    .ParaNo = 3: .ParaValue = RTrim(VsGrid1.TextMatrix(gLong, 16))     '納入日
                    .ParaNo = 4: .ParaValue = CisFun.Val2(VsGrid1.TextMatrix(gLong, 18))   '納品書頁
                    
                    If Not .DBStored Then
                        MsgBox "ストアド異常 "
                        Exit Sub
                    Else
                        .ParaNo = 7
                        If .ParaValue <> 0 Then
                            MsgBox "ストアドエラー "
                            Exit Sub
                        End If
                    End If
                    .ParaNo = 6
                    If .ParaValue <> 0 Then
                       Tlb_ErrCnt = Format(CisFun.Val2(Tlb_ErrCnt) + 1, "#,###")
                    Else
                       Tlb_UpdCnt = Format(CisFun.Val2(Tlb_UpdCnt) + 1, "#,###")
                    End If
               End With
               .TextMatrix(gLong, 1) = "済"
               .Row = gLong
               Call CisVsGrid.vsColor(消去個別)
               mKanCnt = mKanCnt - 1
               B1lb_KKensu = Format(mKanCnt, "#,###")
               DoEvents
            End If
        Next gLong
        .Row = 1: .Col = 1: .TopRow = 1
        .Redraw = True
    End With
    Call CisDB.DBTran(TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ終了
    
PB_UPD_Ed:
    PNL_Wait.Visible = False
    PB_UPD.Visible = False
    PB_DEL.Visible = False
    '   ( QR Reader 環境 )
    Set QRr_Object = PDQComm1
    Set QRr_MsgObject = QRMsg
    Set QRr_ErrMsgObj = ErrMsg
    Set TX2400_StartMsg = PNL_Start
    Set TX2400_ReadObj = QRRMsg
    Set TX2400_ErObj = ErrMsg
    Call QRInit
    Call TX2400InfoGet
    
    Call DispChange(ProcHB)
'    Timer1.Enabled = True
    Exit Sub
PB_UPD_Er:
    VsGrid1.Redraw = True
    PNL_Wait.Visible = False
    Call CisDB.DBTran(TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ破棄

End Sub
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆
'☆                         入　　荷　　実　　績　　削　　除　　(  F3  )
'☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
'☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆★☆
Private Sub PB_DEL_Click()
    
    gCnt1 = 0
    With VsGrid1
        For gCnt2 = 0 To .Rows - 1
            If Trim(.TextMatrix(gCnt2, 1)) = "更新" Then
               gCnt1 = 1
               Exit For
            End If
        Next gCnt2
    End With
        
    If gCnt1 = 0 Then
        With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "      削除行を選択して下さい               "
            .MB_Button = OK_CAN
            .MBOX
            Exit Sub
        End With
    End If
    
    With CisFun
        .MB_Lines = 5
        .MB_MSG(2) = "      入荷データ削除を行います。               "
        .MB_MSG(4) = "          よろしいですか？                      "
        .MB_Button = OK_CAN
        If Not .MBOX Then
            VsGrid1.SetFocus
            Exit Sub
        End If
    End With

    On Error GoTo PB_DEL_Er:
    
    If B2_Area1.Visible Then B2_Area1.Visible = False
    PNL_Wait.Visible = True
    Message = "【入荷削除中】しばらくお待ち下さい"
    DoEvents
    '+---------------
    '+   完納処理
    '+---------------
    Set QRr_Object = PDQComm1
    Set TX2400_Comm = PDQComm2
    Call QR_Comm_Close
    Call TX2400Close
    
    Call CisDB.DBTran(TransBegin) ' ﾄﾗﾝｻﾞｸｼｮﾝ開始

    With VsGrid1
        .Redraw = False
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) = "更新" Then
               With CisDB
                    
                '更新
                    .SQL = "DELETE FROM 前工程入荷ワーク"
                    .SQL = .SQL & " WHERE 処理端末 = HOST_NAME()"
                    .SQL = .SQL & "   AND 納入番号 = '" & RTrim(VsGrid1.TextMatrix(gLong, 6)) & "'"
                    .SQL = .SQL & "   AND 納入日 = '" & RTrim(VsGrid1.TextMatrix(gLong, 16)) & "'"
                    .SQL = .SQL & "   AND 納品書頁 = " & CisFun.Val2(VsGrid1.TextMatrix(gLong, 18))
                    .DBExec
               End With
               .TextMatrix(gLong, 1) = "削除"
               .Row = gLong
               Call CisVsGrid.vsColor(消去個別)
               mKanCnt = mKanCnt - 1
               B1lb_KKensu = Format(mKanCnt, "#,###")
               DoEvents
            End If
        Next gLong
        .Row = 1: .Col = 1: .TopRow = 1
        .Redraw = True
    End With
    Call CisDB.DBTran(TransCommit) ' ﾄﾗﾝｻﾞｸｼｮﾝ終了
    
PB_DEL_Ed:
    PNL_Wait.Visible = False
    PB_UPD.Visible = False
    PB_DEL.Visible = False
    '   ( QR Reader 環境 )
    Set QRr_Object = PDQComm1
    Set QRr_MsgObject = QRMsg
    Set QRr_ErrMsgObj = ErrMsg
    Set TX2400_StartMsg = PNL_Start
    Set TX2400_ReadObj = QRRMsg
    Set TX2400_ErObj = ErrMsg
    Call QRInit
    Call TX2400InfoGet
    
    Call DispChange(ProcHB)
'    Timer1.Enabled = True
    Exit Sub
PB_DEL_Er:
    VsGrid1.Redraw = True
    PNL_Wait.Visible = False
    Call CisDB.DBTran(TransRollback) ' ﾄﾗﾝｻﾞｸｼｮﾝ破棄

End Sub
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit()
    With CisVsGrid
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 中中, あり, 5, 1, "処理")
        Call .FixedSet(中中, 中中, あり, 13, 2, "納入日-便")
        Call .FixedSet(中中, 中中, あり, 10, 3, "検収日")
        Call .FixedSet(中中, 中中, あり, 8, 4, "受注者")
        Call .FixedSet(中中, 中中, あり, 4, 5, "受入")
        Call .FixedSet(中中, 中中, あり, 8, 6, "納品書№")
        Call .FixedSet(中中, 中中, あり, 5, 7, "頁")
        Call .FixedSet(中中, 中中, あり, 7, 8, "支給元1")
        Call .FixedSet(中中, 中中, あり, 7, 9, "支給元2")
        Call .FixedSet(中中, 中中, あり, 4, 10, "発区")
        Call .FixedSet(中中, 中中, あり, 4, 11, "ｶｰﾄﾞ")
        Call .FixedSet(中中, 中中, あり, 6, 12, "出荷場")
        Call .FixedSet(中中, 中中, あり, 10, 13, "発行日")
        Call .FixedSet(中中, 中中, あり, 8, 14, "発注者")
        Call .FixedSet(中中, 中中, あり, 3, 15, "ｴﾗｰ")

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 0, 16, "S納入日")
        Call .FixedSet(中中, 左中, あり, 0, 17, "S納入便")
        Call .FixedSet(中中, 左中, あり, 0, 18, "S頁")
        Call .FixedSet(中中, 左中, なし, 0, 19, "Select")

'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10,11,12,13,14,15")

'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub Grid2Init()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容,分類ﾀｲﾌﾟ)
        Call .FixedSet(中中, 右中, あり, 2, 0, "№")
        Call .FixedSet(中中, 左中, あり, 17, 1, "品　番")
        Call .FixedSet(中中, 中中, あり, 6, 2, "背番号")
        Call .FixedSet(中中, 右中, あり, 7, 3, "収容数")
        Call .FixedSet(中中, 右中, あり, 5, 4, "箱数")
        Call .FixedSet(中中, 右中, あり, 8, 5, "納入数")
        Call .FixedSet(中中, 左中, あり, 6, 6, "エラー")

'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ

'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6")

'  編集内容セット
'       Call .EditSet(ﾃｷｽﾄｵﾌﾞｼﾞｪｸﾄ,ｶﾗﾑNO)
'        Call .EditSet(iH_Text1, 1)
        .InitDisp
    End With
End Sub
'******************************************'
'*****        行選択（クリック）       *****
'******************************************'
Private Sub VSGrid1_Click()
    
    With VsGrid1
        .Redraw = False
        For gLong = .Row To .RowSel
            If Trim(.TextMatrix(gLong, CisVsGrid.NoSpaceCol)) <> "" Then
                gStr = Trim(.TextMatrix(gLong, CisVsGrid.SelectCol))
                .Row = gLong
                Select Case Trim(.TextMatrix(gLong, 1))
                    Case ""
                        If gStr = "X" Then
                           Call CisVsGrid.vsColor(消去個別)
                        End If
                        Call CisVsGrid.vsColor(個別, B1lb_KKensu.BackColor)
                        .TextMatrix(gLong, 1) = "更新"
                        mKanCnt = mKanCnt + 1
                        mMiCnt = mMiCnt - 1
                    Case "更新"
                        If gStr = "X" Then
                           Call CisVsGrid.vsColor(消去個別)
                        End If
                        .TextMatrix(gLong, 1) = ""
                        mKanCnt = mKanCnt - 1
                        mMiCnt = mMiCnt + 1
                    Case Else: Exit For
                End Select
            End If
        Next gLong
        .Redraw = True
    End With
    
    B1lb_KKensu = Format(mKanCnt, "#,###")
    B1lb_MKensu = Format(mMiCnt, "#,###")
    
    If mKanCnt <> 0 Then
        PB_UPD.Visible = True
        PB_DEL.Visible = True
    Else
        PB_UPD.Visible = False
        PB_DEL.Visible = False
    End If

End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    Call VSGrid1_Click
End Sub
Private Sub vsGrid1_RowColChange()
    If ProcHB = "H4" And B2_Area1.Visible Then
        Call Grid2Set
    End If
End Sub
'+------------------------------------+
'+      グリッドセット（個別)
'+------------------------------------+
Private Function GridSetKobetsu(Optional MsgType As Boolean = False) As Boolean
    GridSetKobetsu = False
    
    With VsGrid1
        .Redraw = False
        
        If Trim(.TextMatrix(.Rows - 1, 2)) <> "" Then
           .Rows = .Rows + 1
        End If

        Call GridMove(.Rows - 1)
        .RowPosition(.Rows - 1) = 1
        .Row = 1
        If CheckFlg Then
            mMiCnt = mMiCnt + 1
            Call VSGrid1_Click
        End If
                
        Call CisVsGrid.ReNumber
        
        .Redraw = True
    
    End With

    B1lb_KKensu = Format(mKanCnt, "#,###")
    B1lb_MKensu = Format(mMiCnt, "#,###")
    If MKW.ERR伝票区分 <> 0 Then
        B1lb_EKensu = Format(CisFun.Val2(B1lb_EKensu) + 1, "#,###")
    End If
    
    B1lb_Kensu = Format(CisFun.Val2(B1lb_Kensu) + 1, "#,###")
    If mBunCnt <> 0 Or mKanCnt <> 0 Then
        PB_UPD.Visible = True
        PB_DEL.Visible = True
    Else
        PB_UPD.Visible = False
        PB_DEL.Visible = False
    End If

    GridSetKobetsu = True
End Function
'+------------------------------------+
'+      グリッド項目転送
'+------------------------------------+
Private Sub GridMove(mRow As Long)

    With VsGrid1
        
        With CisFun
            .DateE_BefVal = MKW.納入日
            .DateEdit
            gStr = CisFun.DateE_AfVal
            If Trim(MKW.納入便) <> "" Then
                gStr = gStr & "-" & MKW.納入便
            Else
                gStr = gStr & "   "
            End If
        End With
        .TextMatrix(mRow, 2) = gStr
        
        If RTrim(iH1_NYmd) = "" Then
            gStr = "=========="
        Else
            With CisFun
                .DateE_BefVal = RTrim(iH1_NYmd)
                .DateEdit
                gStr = CisFun.DateE_AfVal
            End With
        End If
        .TextMatrix(mRow, 3) = gStr
        
        
        gStr = MKW.受注者 & "-" & MKW.受注者工場
        .TextMatrix(mRow, 4) = gStr
        .TextMatrix(mRow, 5) = MKW.受入
        .TextMatrix(mRow, 6) = MKW.納入番号
        .TextMatrix(mRow, 7) = Format(MKW.納品書頁, "@@") & "/" & Format(MKW.総頁, "@@")
        .TextMatrix(mRow, 8) = MKW.支給元1
        .TextMatrix(mRow, 9) = MKW.支給元2
        .TextMatrix(mRow, 10) = MKW.発注区分
        .TextMatrix(mRow, 11) = MKW.カード
        .TextMatrix(mRow, 12) = MKW.出荷場
        
        With CisFun
            .DateE_BefVal = MKW.発行日
            .DateEdit
            gStr = CisFun.DateE_AfVal
        End With
        .TextMatrix(mRow, 13) = gStr
        
        .TextMatrix(mRow, 14) = MKW.発注者 & "-" & MKW.発注者工場
        If MKW.ERR伝票区分 <> 0 Then
            .TextMatrix(mRow, 15) = "ERR"
        End If
        .TextMatrix(mRow, 16) = MKW.納入日
        .TextMatrix(mRow, 17) = MKW.納入便
        .TextMatrix(mRow, 18) = MKW.納品書頁
    
    End With
End Sub
'+------------------------------------+
'+      グリッド2セット（個別)
'+------------------------------------+
Private Function Grid2Set() As Boolean
    Grid2Set = False
    Dim ID          As Long
    
    Call HeadBodyClear("B2")
    Call CisVsGrid2.Clear
    If RTrim(VsGrid1.TextMatrix(VsGrid1.Row, 2)) = "" Then Exit Function
        
    B2lb_NYmd = RTrim(VsGrid1.TextMatrix(VsGrid1.Row, 2))
    B2lb_Nouno = RTrim(VsGrid1.TextMatrix(VsGrid1.Row, 6))
    B2lb_Page = RTrim(VsGrid1.TextMatrix(VsGrid1.Row, 18))
    B2lb_Torcd = RTrim(VsGrid1.TextMatrix(VsGrid1.Row, 4))
    B2lb_Ukeir = RTrim(VsGrid1.TextMatrix(VsGrid1.Row, 5))
    B2lb_SMoto = RTrim(VsGrid1.TextMatrix(VsGrid1.Row, 8))
    B2lb_SMoto2 = RTrim(VsGrid1.TextMatrix(VsGrid1.Row, 9))
    
'( データ確認 ワーク)
    gSL_Select = "SELECT MK.*,HN.表示品番 FROM 前工程入荷ワーク MK"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "   ON HN.品番 = MK.品番"
    gSL_Select = gSL_Select & " WHERE 納入番号 = '" & RTrim(VsGrid1.TextMatrix(VsGrid1.Row, 6)) & "'"
    gSL_Select = gSL_Select & "   AND 納品書頁 = " & CisFun.Val2(VsGrid1.TextMatrix(VsGrid1.Row, 18))
    gSL_Select = gSL_Select & "   AND 納入日 = '" & RTrim(VsGrid1.TextMatrix(VsGrid1.Row, 16)) & "'"
    gSL_Select = gSL_Select & "   AND 処理端末 = SUBSTRING(HOST_NAME(),1,20)"
    gSL_Select = gSL_Select & " ORDER BY 納品書行"
    If Not MKWRead(gSL_Select, 0) Then Exit Function
    
    With vsGrid2
        .Redraw = False
    End With
    
    ID = 0
    Do Until Not MKW_RDSTS
        With vsGrid2
            ID = ID + 1
            If ID > 10 Then Exit Do
            'No
            .TextMatrix(ID, 0) = ID
            '品番
            If RTrim(MKW.表示品番) <> "" Then
                .TextMatrix(ID, 1) = MKW.表示品番
            Else
                .TextMatrix(ID, 1) = MKW.品番
            End If
            '背番号
            .TextMatrix(ID, 2) = MKW.背番号
            '収容数
            .TextMatrix(ID, 3) = Format(MKW.収容数, "#,###")
            '箱数
            .TextMatrix(ID, 4) = Format(MKW.枚数, "#,###")
            '納入数
            .TextMatrix(ID, 5) = Format(MKW.納入数, "#,###")
            'エラー
            Select Case MKW.ERRCD
                Case 0: gStr = ""
                Case 1: gStr = "品番"
                Case 2: gStr = "取変換"
                Case 3: gStr = "品番取"
                Case Else: gStr = "その他"
            End Select
            .TextMatrix(ID, 6) = gStr
        
        End With
        Call MKWReadNext
    Loop
    Call MKWClose
    With vsGrid2
'         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With


    Grid2Set = True
End Function
'*=======================================================================================================*
'*=======================================================================================================*
'*
'*                      Ｑ  Ｒ  リ  ー  ダ  (　Ｃｏｍｍ　)  処  理
'*
'*=======================================================================================================*
'*=======================================================================================================*
'+------------------------------------+
'+      読  取  処  理(QRﾀｯﾁﾘｰﾀﾞ)     +
'+------------------------------------+
Private Sub PDQComm1_OnComm()
    
    If QR_CommEvent <> "OK" Then Exit Sub
    
    With QRDataArea
        If .ステータス <> 0 Then
           Call QRReadMsg("E")
           Exit Sub
        End If
        
    ' JAMA伝票の納品書・支給書のみ有効
        If .バーコード種類 <> "J" Then
            .エラー内容 = "JAMA伝票のみ対象"
            Call QRReadMsg("E")
            Exit Sub
        End If
        
        If typQRAfter.カード丼 <> "7" And typQRAfter.カード丼 <> "8" Then
           .エラー内容 = "カードが対象外です"
           Call QRReadMsg("E")
           Exit Sub
        End If
        QRr_ReadData = ""
        
        ProcHB = "H4"
    End With
    
'( データ確認 テーブル)
    gSL_Select = "SELECT COUNT(*) 件数 FROM 受給検収テーブル"
    gSL_Select = gSL_Select & " WHERE 納入番号 = '" & typQRAfter.納入番号 & "'"
    gSL_Select = gSL_Select & "   AND 納品書頁 = " & typQRAfter.ページNO
    gSL_Select = gSL_Select & "   AND 納入日 = '" & typQRAfter.納入日 & "'"
    If JYKRead(gSL_Select, 1) Then
        If JYK.件数 <> 0 Then
           QRDataArea.エラー内容 = "入力済み伝票"
           Call QRReadMsg("E")
           Exit Sub
        End If
    End If
'( データ確認 ワーク)
    gSL_Select = "SELECT COUNT(*) 件数 FROM 前工程入荷ワーク"
    gSL_Select = gSL_Select & " WHERE 納入番号 = '" & typQRAfter.納入番号 & "'"
    gSL_Select = gSL_Select & "   AND 納品書頁 = " & typQRAfter.ページNO
    gSL_Select = gSL_Select & "   AND 納入日 = '" & typQRAfter.納入日 & "'"
    gSL_Select = gSL_Select & "   AND 処理端末 = SUBSTRING(HOST_NAME(),1,20)"
    If MKWRead(gSL_Select, 1) Then
        If MKW.件数 <> 0 Then
           QRDataArea.エラー内容 = "表示済み伝票"
           Call QRReadMsg("E")
           Exit Sub
        End If
    End If
    
'( ワークへ出力 )
    Call ItemsClearMKW
    With MKW
        .納入番号 = typQRAfter.納入番号
        .納品書頁 = typQRAfter.ページNO
        .納入日 = typQRAfter.納入日
        .納入便 = typQRAfter.納入便
        .受注者 = typQRAfter.仕入先
        .受注者工場 = typQRAfter.仕入先工区
        .発注区分 = typQRAfter.発注区分
        .カード = typQRAfter.カード丼
        .機番 = typQRAfter.機番
        .支給元体系 = typQRAfter.支給元体系
        .支給元1 = typQRAfter.支給元1
        .支給元2 = typQRAfter.支給元2
        .総頁 = typQRAfter.総ページ数
        .出荷場 = typQRAfter.出荷場
        .発注者 = typQRAfter.納入先
        .発注者工場 = typQRAfter.納入先工区
        .受入 = typQRAfter.受入
        .発行日 = typQRAfter.発行日
        
        If RTrim(iH1_NYmd) = "" Then
            .検収日 = .納入日
        Else
            .検収日 = iH1_NYmd
        End If
        .処理端末 = CisFun.Left2(CisFun.WSID, 20)
        For gInt = 1 To 10
            If RTrim(typQRAfter.品番(gInt)) <> "" Then
                .納品書行 = gInt
                .品番 = typQRAfter.品番(gInt)
                .背番号 = typQRAfter.背番号(gInt)
                .収容数 = typQRAfter.収容数(gInt)
                .枚数 = typQRAfter.箱数(gInt)
                .納入数 = typQRAfter.納入数(gInt)
                Call MKWInsert
            End If
        Next
        
    End With
'発注打切更新
    With CisDB
        .SQL = "前工程入荷データチェック"
        .StoadoCount = 7
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用           '処理区分(0:ﾁｪｯｸ、1:更新）
        .ParaNo = 2: .ParaIO = Input用           '納入番号
        .ParaNo = 3: .ParaIO = Input用           '納入日
        .ParaNo = 4: .ParaIO = Input用           '納品書頁
        .ParaNo = 5: .ParaIO = OutPut用          '作成件数
        .ParaNo = 6: .ParaIO = OutPut用          'エラー件数
        .ParaNo = 7: .ParaIO = OutPut用          'エラーコード
'
        .ParaNo = 1: .ParaValue = 0              '処理区分(0:ﾁｪｯｸ、1:更新）
        .ParaNo = 2: .ParaValue = MKW.納入番号   '納入番号
        .ParaNo = 3: .ParaValue = MKW.納入日     '納入日
        .ParaNo = 4: .ParaValue = MKW.納品書頁   '納品書頁
        
        If Not .DBStored Then
            MsgBox "ストアド異常 "
            Exit Sub
        Else
            .ParaNo = 7
            If .ParaValue <> 0 Then
                MsgBox "ストアドエラー "
                Exit Sub
            End If
        End If
        .ParaNo = 6
        If .ParaValue <> 0 Then
            MKW.ERR伝票区分 = 1
        Else
            MKW.ERR伝票区分 = 0
        End If
    End With
    
    If Not GridSetKobetsu(True) Then
       Call QRReadMsg("E")
       Exit Sub
    End If
    
End Sub
'+------------------------------------+
'+      読  取  処  理(QRﾎｯﾊﾟﾘｰﾀﾞ)     +
'+------------------------------------+
Private Sub PDQComm2_OnComm()

    Call TX2400Event
'   * エラー *
    If TX2400_Event = "ERR" Then
         QRDataArea.エラー内容 = TX2400_EventMsg
         Call TX2400QRRReadMsg("E")
         TX2400_ReadFlg = True
         Exit Sub            ' 通信エラー処理
    End If

    If TX2400_Event <> "RCV" Then Exit Sub   ' 受信ﾃﾞｰﾀ編集
    
    If Not TX2400SendProc Then Exit Sub    ' CR が存在しない時抜ける
    ' ｽﾃｰﾀｽﾁｪｯｸ
    If Not TX2400STSCheck Then
        If TX2400_STSNo <> 0 Then
           QRDataArea.エラー内容 = TX2400_STSMsg
        End If
        PNL_Start.Tag = ""
        GoTo PDQComm2_OnComm_ED
    End If
    
    If TX2400_STSNo = 1 Then
       Call TX2400OpenAuto
       Exit Sub
    End If
    If TX2400_STSNo = 2 Then
       Call TX2400ReadAuto
       Exit Sub
    End If
    
    If Trim(TX2400_CommData) = "" Then Exit Sub
    
    Call TX2400QRRReadMsg("R")
    
    ' 読取内容編集
    QRr_ReadData = TX2400_CommData
    If Not QRReadCheck Then
       TX2400_STSNo = 5
       GoTo PDQComm2_OnComm_ED
    End If
    
    With QRDataArea
        If .ステータス <> 0 Then
           TX2400_STSNo = 5
           GoTo PDQComm2_OnComm_ED
        End If
    ' JAMA伝票の納品書・支給書のみ有効
        If .バーコード種類 <> "J" Then
            .エラー内容 = "JAMA伝票のみ対象"
            TX2400_STSNo = 5
            GoTo PDQComm2_OnComm_ED
        End If
        
        If typQRAfter.カード丼 <> "7" And typQRAfter.カード丼 <> "8" Then
           .エラー内容 = "カードが対象外です"
           TX2400_STSNo = 5
           GoTo PDQComm2_OnComm_ED
        End If
        
    End With
'( データ確認 テーブル)
    gSL_Select = "SELECT COUNT(*) 件数 FROM 受給検収テーブル"
    gSL_Select = gSL_Select & " WHERE 納入番号 = '" & typQRAfter.納入番号 & "'"
    gSL_Select = gSL_Select & "   AND 納品書頁 = " & typQRAfter.ページNO
    gSL_Select = gSL_Select & "   AND 納入日 = '" & typQRAfter.納入日 & "'"
    If JYKRead(gSL_Select, 1) Then
        QRDataArea.エラー内容 = "入力済み伝票"
        TX2400_STSNo = 5
        GoTo PDQComm2_OnComm_ED
    End If
'( データ確認 ワーク)
    gSL_Select = "SELECT COUNT(*) 件数 FROM 前工程入荷ワーク"
    gSL_Select = gSL_Select & " WHERE 納入番号 = '" & typQRAfter.納入番号 & "'"
    gSL_Select = gSL_Select & "   AND 納品書頁 = " & typQRAfter.ページNO
    gSL_Select = gSL_Select & "   AND 納入日 = '" & typQRAfter.納入日 & "'"
    gSL_Select = gSL_Select & "   AND 処理端末 = SUBSTRING(HOST_NAME(),1,20)"
    If MKWRead(gSL_Select, 1) Then
        If MKW.件数 <> 0 Then
            QRDataArea.エラー内容 = "表示済み伝票"
            TX2400_STSNo = 5
            GoTo PDQComm2_OnComm_ED
        End If
    End If
'( ワークへ出力 )
    Call ItemsClearMKW
    With MKW
        .納入番号 = typQRAfter.納入番号
        .納品書頁 = typQRAfter.ページNO
        .納入日 = typQRAfter.納入日
        .納入便 = typQRAfter.納入便
        .受注者 = typQRAfter.仕入先
        .受注者工場 = typQRAfter.仕入先工区
        .発注区分 = typQRAfter.発注区分
        .カード = typQRAfter.カード丼
        .機番 = typQRAfter.機番
        .支給元体系 = typQRAfter.支給元体系
        .支給元1 = typQRAfter.支給元1
        .支給元2 = typQRAfter.支給元2
        .総頁 = typQRAfter.総ページ数
        .出荷場 = typQRAfter.出荷場
        .発注者 = typQRAfter.納入先
        .発注者工場 = typQRAfter.納入先工区
        .受入 = typQRAfter.受入
        .発行日 = typQRAfter.発行日
        
        If RTrim(iH1_NYmd) = "" Then
            .検収日 = .納入日
        Else
            .検収日 = iH1_NYmd
        End If
        .処理端末 = CisFun.Left2(CisFun.WSID, 20)
        For gInt = 1 To 10
            If RTrim(typQRAfter.品番(gInt)) <> "" Then
                .納品書行 = gInt
                .品番 = typQRAfter.品番(gInt)
                .背番号 = typQRAfter.背番号(gInt)
                .収容数 = typQRAfter.収容数(gInt)
                .枚数 = typQRAfter.箱数(gInt)
                .納入数 = typQRAfter.納入数(gInt)
                Call MKWInsert
            End If
        Next
    End With
'発注打切更新
    With CisDB
        .SQL = "前工程入荷データチェック"
        .StoadoCount = 7
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用           '処理区分(0:ﾁｪｯｸ、1:更新）
        .ParaNo = 2: .ParaIO = Input用           '納入番号
        .ParaNo = 3: .ParaIO = Input用           '納入日
        .ParaNo = 4: .ParaIO = Input用           '納品書頁
        .ParaNo = 5: .ParaIO = OutPut用          '作成件数
        .ParaNo = 6: .ParaIO = OutPut用          'エラー件数
        .ParaNo = 7: .ParaIO = OutPut用          'エラーコード
'
        .ParaNo = 1: .ParaValue = 0              '処理区分(0:ﾁｪｯｸ、1:更新）
        .ParaNo = 2: .ParaValue = MKW.納入番号   '納入番号
        .ParaNo = 3: .ParaValue = MKW.納入日     '納入日
        .ParaNo = 4: .ParaValue = MKW.納品書頁   '納品書頁
        
        If Not .DBStored Then
            MsgBox "ストアド異常 "
            Exit Sub
        Else
            .ParaNo = 7
            If .ParaValue <> 0 Then
                MsgBox "ストアドエラー "
                Exit Sub
            End If
        End If
        .ParaNo = 6
        If .ParaValue <> 0 Then
            MKW.ERR伝票区分 = 1
        Else
            MKW.ERR伝票区分 = 0
        End If
    End With
    
    If Not GridSetKobetsu(True) Then
       TX2400_STSNo = 5
       GoTo PDQComm2_OnComm_ED
    End If
    
    Call TX2400Stack(True)
    
    Exit Sub
PDQComm2_OnComm_ED:
    TX2400_ReadFlg = True
    Call TX2400QRRReadMsg("E")
    If TX2400_STSNo = 1 Then
       Call TX2400QRRReadMsg("")
       PNL_Start.Visible = True
       DoEvents
       Call TX2400ReadAuto
    End If
    If TX2400_STSNo = 2 Then
       Call TX2400Stack(False)
    End If
    If TX2400_STSNo = 4 Then
       Call TX2400Stack(False)
    End If
    If TX2400_STSNo = 5 Then
       Call TX2400Stack(False)
    End If
    If TX2400_STSNo = 3 Then
       PNL_Start.Visible = True
       DoEvents
    End If
End Sub
'+------------------------------------+
'+      エラーデータチェック
'+------------------------------------+
Private Function ErrorDataCheck(CheckKB As Boolean) As Boolean
    ErrorDataCheck = False
    
    CheckFlg = CheckKB
    ' 入荷実績処理ワーク削除
    With CisDB
        .SQL = "DELETE FROM 前工程入荷ワーク"
        .SQL = .SQL & "    Where 処理端末 = substring(HOST_NAME(),1,20)"
        .SQL = .SQL & "      AND 処理区分 = 0"
        .DBExec
    End With
    
'( データ確認 ワーク)
    gSL_Select = "SELECT COUNT(*) 件数 FROM 前工程入荷ワーク "
    gSL_Select = gSL_Select & " WHERE 処理端末 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 処理区分 = 1"
    gSL_Select = gSL_Select & "   AND ERRCD <> 0"
    If Not MKWRead(gSL_Select, 1) Then Exit Function
    If MKW.件数 = 0 Then Exit Function
'エラーデータチェック
   With CisDB
        
    '更新
        .SQL = "前工程入荷データチェック"
        .StoadoCount = 7
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用           '処理区分(0:ﾁｪｯｸ、1:更新）
        .ParaNo = 2: .ParaIO = Input用           '納入番号
        .ParaNo = 3: .ParaIO = Input用           '納入日
        .ParaNo = 4: .ParaIO = Input用           '納品書頁
        .ParaNo = 5: .ParaIO = OutPut用          '作成件数
        .ParaNo = 6: .ParaIO = OutPut用          'エラー件数
        .ParaNo = 7: .ParaIO = OutPut用          'エラーコード
'
        .ParaNo = 1: .ParaValue = 0              '処理区分(0:ﾁｪｯｸ、1:更新）
        .ParaNo = 2: .ParaValue = ""            '納入番号
        .ParaNo = 3: .ParaValue = ""            '納入日
        .ParaNo = 4: .ParaValue = 0             '納品書頁
        
        If Not .DBStored Then
            MsgBox "ストアド異常 "
            Exit Function
        Else
            .ParaNo = 7
            If .ParaValue <> 0 Then
                MsgBox "ストアドエラー "
                Exit Function
            End If
        End If
   End With
    
    gCnt1 = 0
    gSL_Select = "SELECT 納入日,納入便,検収日,受注者,受注者工場,受入,納入番号,納品書頁,総頁,支給元1,支給元2,"
    gSL_Select = gSL_Select & " 発注区分,カード,出荷場,発行日,発注者,発注者工場,処理区分,ERR伝票区分"
    gSL_Select = gSL_Select & "  FROM 前工程入荷ワーク "
    gSL_Select = gSL_Select & " WHERE 処理端末 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 処理区分 = 1"
    gSL_Select = gSL_Select & " GROUP BY 納入日,納入便,検収日,受注者,受注者工場,受入,納入番号,納品書頁,総頁,支給元1,支給元2,"
    gSL_Select = gSL_Select & "         発注区分,カード,出荷場,発行日,発注者,発注者工場,処理区分,ERR伝票区分"
    gSL_Select = gSL_Select & " ORDER BY 納入日,納入便,受注者,受注者工場,受入,納入番号"
    If Not MKWRead(gSL_Select, 0) Then Call MKWClose: Exit Function
    Do Until Not MKW_RDSTS
        gCnt1 = gCnt1 + 1
        If Not GridSetKobetsu(True) Then
           Call QRReadMsg("E")
           Exit Function
        End If
        
        
        Call MKWReadNext
    Loop
    Call MKWClose
    
    If Not CheckFlg Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(1) = "     　マスタ未登録のエラーデータが存在します。　　"
            .MB_MSG(3) = "     　　"
            .MB_Button = OK
            .MBOX
        End With
    End If

    ErrorDataCheck = True
    Exit Function:
ErrorDataCheck_Err:

End Function

'*-------------------------------------------------*
'*      エラーリスト
'*-------------------------------------------------*
Public Sub Error_List()
    
    gSL_Select = "SELECT CASE WHEN ERRCD = 1 THEN '品番マスタ未登録'"
    gSL_Select = gSL_Select & "    WHEN  ERRCD = 2 THEN '取引先変換マスタ未登録'"
    gSL_Select = gSL_Select & "    WHEN  ERRCD = 3 THEN '品番取引先マスタ未登録'"
    gSL_Select = gSL_Select & "    Else '' END エラー内容,"
    gSL_Select = gSL_Select & "    支給元1,支給元2,"
    gSL_Select = gSL_Select & "    受注者 + '-' + 受注者工場 受注者,受入,発注者 + '-' + 発注者工場 発注者,"
    gSL_Select = gSL_Select & "    CASE WHEN ISNULL(HN.表示品番,'') <> '' THEN HN.表示品番 ELSE MW.品番 END 品番,"
    gSL_Select = gSL_Select & "    背番号,収容数,枚数,納入数,検収日,カード,"
    gSL_Select = gSL_Select & "    納入番号 + '-' + CONVERT(VARCHAR(2),納品書頁) 納入番号"
    gSL_Select = gSL_Select & " FROM 前工程入荷ワーク MW"
    gSL_Select = gSL_Select & "  LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "    ON HN.品番 = MW.品番"
    gSL_Select = gSL_Select & " Where 処理端末 = HOST_NAME()"
    gSL_Select = gSL_Select & "   AND 処理区分 = 1"
    gSL_Select = gSL_Select & "   AND ERRCD <> 0"
    gSL_Select = gSL_Select & " ORDER BY ERRCD,支給元1,支給元2,受注者,受注者工場,受入,発注者,発注者工場"
    If Not MKWRead(gSL_Select) Then
       Call MKWClose
       Exit Sub
    End If
    
    Dim CisSimplePrint As New CisSimplePrint
    
    With CisSimplePrint
        .PrintTitle = "前工程発行JAMA入荷エラーリスト"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Landscape
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        
        
        .pHeadItemSet 0, 中, "エラー内容"
        .pHeadItemSet 1, 中, "支給元1"
        .pHeadItemSet 2, 中, "支給元2"
        .pHeadItemSet 3, 中, "受注者"
        .pHeadItemSet 4, 中, "受入"
        .pHeadItemSet 5, 中, "発注者"
        .pHeadItemSet 6, 中, "品　番"
        .pHeadItemSet 7, 中, "背番号"
        .pHeadItemSet 8, 右, "収容数"
        .pHeadItemSet 9, 右, "枚数"
        .pHeadItemSet 10, 右, "納入数"
        .pHeadItemSet 11, 中, "納入日"
        .pHeadItemSet 12, 中, "納入番号"
        .pHeadItemSet 13, 中, "#"
    
        .pBodyItemSet 0, 22
        .pBodyItemSet 1, 7
        .pBodyItemSet 2, 7
        .pBodyItemSet 3, 7
        .pBodyItemSet 4, 4
        .pBodyItemSet 5, 7
        .pBodyItemSet 6, 14
        .pBodyItemSet 7, 6
        .pBodyItemSet 8, 8, 右
        .pBodyItemSet 9, 6, 右
        .pBodyItemSet 10, 8, 右
        .pBodyItemSet 11, 10
        .pBodyItemSet 12, 8
        .pBodyItemSet 13, 2
        
        .BodyInit
        .MaxCount = CisDB.RecordCount
        
        If Not .CheckPrtCsv Then Exit Sub
        
        Do Until (Not MKW_RDSTS)
            
            .pBodyValueSet 0, MKW.エラー内容
            .pBodyValueSet 1, MKW.支給元1
            .pBodyValueSet 2, MKW.支給元2
            .pBodyValueSet 3, MKW.受注者
            .pBodyValueSet 4, MKW.受入
            .pBodyValueSet 5, MKW.発注者
            .pBodyValueSet 6, MKW.品番
            .pBodyValueSet 7, MKW.背番号
            .pBodyValueSet 8, Format(MKW.収容数, "#,###")
            .pBodyValueSet 9, Format(MKW.枚数, "#,###")
            .pBodyValueSet 10, Format(MKW.納入数, "#,###")
            With CisFun
                .DateE_BefVal = MKW.検収日
                .DateEdit
            End With
            .pBodyValueSet 11, CisFun.DateE_AfVal
            .pBodyValueSet 12, MKW.納入番号
            .pBodyValueSet 13, MKW.カード
            
            .BodySet
            
            Call MKWReadNext
        Loop
        
        .PrintStart
    
    End With
    
End Sub

