VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{1EE89F6D-4F21-4577-AD2D-097803C5570A}#1.3#0"; "CISYMD~3.OCX"
Begin VB.Form CXM1160 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "売上単価未登録一覧"
   ClientHeight    =   11400
   ClientLeft      =   1575
   ClientTop       =   1275
   ClientWidth     =   15405
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
   ScaleHeight     =   11400
   ScaleWidth      =   15405
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      Height          =   645
      Left            =   12270
      ScaleHeight     =   585
      ScaleWidth      =   2265
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1830
      Width           =   2325
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   570
         Left            =   15
         Top             =   0
         Width           =   2235
         _ExtentX        =   3942
         _ExtentY        =   1005
         Caption         =   " 表示件数"
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
         Begin Cis3D_v60.CIS3D B1lb_DispSu 
            Height          =   435
            Left            =   1185
            Top             =   60
            Width           =   960
            _ExtentX        =   1693
            _ExtentY        =   767
            BackColor       =   12648447
            Caption         =   "#,###"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
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
            cPositionX      =   -60
         End
      End
   End
   Begin Cis3D_v60.CIS3D B1_Area1 
      Height          =   8370
      Left            =   840
      Top             =   2520
      Width           =   13755
      _ExtentX        =   24262
      _ExtentY        =   14764
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
      Begin Cis3D_v60.CIS3D CIS3D18 
         Height          =   8100
         Left            =   135
         Top             =   120
         Width           =   13500
         _ExtentX        =   23813
         _ExtentY        =   14288
         ForeColor       =   0
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
         cAlingnment     =   6
         Begin vsFlexLib.vsFlexArray vsGrid1 
            DragIcon        =   "CXM1160.frx":0000
            Height          =   7965
            Left            =   90
            TabIndex        =   5
            Top             =   60
            Width           =   13335
            _Version        =   196608
            _ExtentX        =   23521
            _ExtentY        =   14049
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
            FormatString    =   "№ |処理|契約先|品番|日付|数量合計|契約先CD|品番CD|"
            Rows            =   25
            Cols            =   9
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            ScrollBars      =   2
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Message 
               Height          =   990
               Left            =   3195
               Top             =   1365
               Visible         =   0   'False
               Width           =   6420
               _ExtentX        =   11324
               _ExtentY        =   1746
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
               Begin Cis3D_v60.CIS3D MessageDisp 
                  Height          =   705
                  Left            =   120
                  Top             =   150
                  Width           =   6180
                  _ExtentX        =   10901
                  _ExtentY        =   1244
                  BackColor       =   12648384
                  ForeColor       =   0
                  Caption         =   "【単価決定処理中】しばらくお待ち下さい"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   14.25
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
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1545
      Left            =   840
      ScaleHeight     =   1485
      ScaleWidth      =   9180
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   750
      Width           =   9240
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   1290
         Left            =   5595
         Top             =   90
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   2275
         ForeColor       =   16711680
         Caption         =   "納 入 先"
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
            Left            =   75
            TabIndex        =   2
            Top             =   615
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
            Left            =   1065
            Top             =   615
            Width           =   2325
            _ExtentX        =   4101
            _ExtentY        =   635
            BackColor       =   16777152
            Caption         =   "XXXXXXXXX1XX"
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
      Begin Cis3D_v60.CIS3D CIS3D31 
         Height          =   1290
         Left            =   90
         Top             =   90
         Width           =   2910
         _ExtentX        =   5133
         _ExtentY        =   2275
         ForeColor       =   16711680
         Caption         =   "表示対象"
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
         Begin VB.OptionButton H1Op_TanKB 
            BackColor       =   &H00C0FFC0&
            Caption         =   "出荷データ"
            Height          =   405
            Index           =   0
            Left            =   150
            TabIndex        =   3
            Top             =   420
            Width           =   2685
         End
         Begin VB.OptionButton H1Op_TanKB 
            BackColor       =   &H00C0C0C0&
            Caption         =   "単価未決（売上実績)"
            Height          =   405
            Index           =   1
            Left            =   150
            TabIndex        =   4
            Top             =   810
            Width           =   2685
         End
      End
      Begin Cis3D_v60.CIS3D Back_Date 
         Height          =   1290
         Left            =   3000
         Top             =   90
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   2275
         ForeColor       =   16711680
         Caption         =   " 納 入 日"
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
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   345
            Index           =   1
            Left            =   90
            Top             =   405
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   609
            ForeColor       =   0
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin Cis3D_v60.CIS3D CIS3D10 
            Height          =   345
            Index           =   2
            Left            =   90
            Top             =   855
            Width           =   510
            _ExtentX        =   900
            _ExtentY        =   609
            ForeColor       =   0
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
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cAlingnment     =   7
         End
         Begin CisYMD_With_Btn.CisYMDwB iH1_ENYmd 
            Height          =   360
            Left            =   615
            TabIndex        =   1
            Top             =   840
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
         Begin CisYMD_With_Btn.CisYMDwB iH1_SNYmd 
            Height          =   360
            Left            =   615
            TabIndex        =   0
            Top             =   390
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
      Left            =   10155
      ScaleHeight     =   660
      ScaleWidth      =   4380
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   750
      Width           =   4440
      Begin CisBtn_60.CisBtn PB_Prt 
         Height          =   585
         Left            =   900
         Top             =   210
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
      Begin CisBtn_60.CisBtn PB_Chg 
         Height          =   585
         Left            =   930
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
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1815
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
         Left            =   2655
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
         Left            =   3495
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
      IMEMode         =   3  'ｵﾌ固定
      Left            =   960
      TabIndex        =   6
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
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   405
      Left            =   0
      Top             =   0
      Width           =   15375
      _ExtentX        =   27120
      _ExtentY        =   714
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 売上単価未登録一覧  】"
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
         Left            =   14400
         Top             =   75
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
         Height          =   315
         Left            =   11850
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
Attribute VB_Name = "CXM1160"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   売上単価未登録一覧
'**       フォームID    :   CXM1160
'**       処理概要      :
'**
'**       作  成  日    :   2005/06/10  By CIS
'**       変  更  日    :
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim CisVsGrid               As New CisVsGrid3
    

    Dim mJiseki                 As String
    Dim mSelColor               As Long
    Dim mUnSelColor             As Long
    Dim mGrpOrder               As String
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
    PB_Prt.Move PB_Chg.Left, PB_Chg.Top
' 売上処理対象ファイルを求める  1=出荷テーブル,1≠検収テーブル
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 50
    Call IRNRead("", 1)
    mJiseki = Mid(IRN.oyian, 1, 1)
    
    If mJiseki <> "1" Then
       Back_Date.Caption = "検 収 日"
       H1Op_TanKB(0).Caption = "検 収 情 報"
    Else
       Back_Date.Caption = "出 荷 日"
       H1Op_TanKB(0).Caption = "出 荷 情 報"
    End If
    mSelColor = H1Op_TanKB(0).BackColor
    mUnSelColor = H1Op_TanKB(1).BackColor
    H1Op_TanKB(0).Value = True
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 3
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 8
'        Set .SelectDispObj = B1lb_SelSu
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
    
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
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
           Case vbKeyF6:        If PB_Chg.Visible Then Call PB_Chg_Click    '【切替】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
           Case vbKeyF8:        If PB_Prt.Visible Then Call PB_Prt_Click    '【印刷】
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
   If ProcHB = "H1" Then
      Call HeadBodyClear("H")
      GoTo PB_CAN_Ed
   End If
    
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
'+    切　替  キ ー(F6)       +
'+----------------------------+
Private Sub PB_Chg_Click()
    Rv_Call3 = "C"
    Rv_Call3Rtn = 0
    
    CXM0160.Show vbModal
    Unload CXM0160
    Set CXM0160 = Nothing
End Sub
'+----------------------------+
'+    印　刷  キ ー(F8)       +
'+----------------------------+
Private Sub PB_Prt_Click()
    Dim CisSimplePrint As New CisSimplePrint
    
    With CisSimplePrint
        If H1Op_TanKB(0).Value Then
           gStr = H1Op_TanKB(0).Caption
        Else
           gStr = H1Op_TanKB(1).Caption
        End If
        .PrintTitle = "単価未決リスト【" & gStr & "】"
        .Preservation = App.Title
        .PrintSize = A4
        .PrintOrientation = Portrait
        .PrintHeadLine = あり
        .PrintBodyLine = あり
        .PrintBodyGyo = 10
        
        .pHeadItemSet 0, 中, "№"
        .pHeadItemSet 1, 中, "納　入  先"
        .pHeadItemSet 2, 中, "品  番"
        If mJiseki <> "1" Then
           .pHeadItemSet 3, 中, "検 収 日"
           .pHeadItemSet 4, 右, "検収数計"
        Else
           .pHeadItemSet 3, 中, "出 荷 日"
           .pHeadItemSet 4, 右, "出荷数計"
        End If
    
        .pBodyItemSet 0, 5, 右
        .pBodyItemSet 1, 20
        .pBodyItemSet 2, 25
        .pBodyItemSet 3, 22
        .pBodyItemSet 4, 8, 右
        
        .BodyInit
        .MaxCount = CisFun.Val2(B1lb_DispSu)
        
        If Not .CheckPrtCsv Then Exit Sub
        
        For gInt = 1 To .MaxCount
            
            .pBodyValueSet 0, Format(gInt, "#,###")
            .pBodyValueSet 1, Trim(vsGrid1.TextMatrix(gInt, CisVsGrid.FixedGet(2)))
            .pBodyValueSet 2, Trim(vsGrid1.TextMatrix(gInt, CisVsGrid.FixedGet(3)))
            .pBodyValueSet 3, vsGrid1.TextMatrix(gInt, CisVsGrid.FixedGet(4))
            .pBodyValueSet 4, vsGrid1.TextMatrix(gInt, CisVsGrid.FixedGet(5))
            
            .BodySet
            
        Next gInt
        
        .PrintStart
    
    End With

End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_SNYmd" Then Call iH1_SNYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_ENYmd" Then Call iH1_ENYmd.ShowCalender: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SSTorcd: Exit Sub
End Sub
'+-----------------------------+
'+       取引先検索             +
'+-----------------------------+
Private Sub Look_SSTorcd()
    RV_Left = 0
    RV_Top = 0
    RV_Call = "C"
    RV_TorKb = 3
    
    CKK0025.Show vbModal
    Unload CKK0025
    Set CKK0025 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_TorcdK
        H1lb_Tornm = RV_TorRName
    End If
End Sub
'+-----------------------------+
'+  H1Op_TanKB
'+-----------------------------+
Private Sub H1Op_TanKB_Click(Index As Integer)
    If Index = 0 Then
       Back_Date.Enabled = True
       H1Op_TanKB(0).BackColor = mSelColor
       H1Op_TanKB(1).BackColor = mUnSelColor
    Else
       Back_Date.Enabled = False
       H1Op_TanKB(0).BackColor = mUnSelColor
       H1Op_TanKB(1).BackColor = mSelColor
       iH1_SNYmd = ""
       iH1_ENYmd = ""
    End If
End Sub

'+-----------------------------+
'+  iH1_SNYmd
'+-----------------------------+
Private Sub iH1_SNYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_SNYmd_LostFocus()
    If ActiveControl.Name <> "PB_Look" Then
        PB_Look.Tag = ""
        PB_Look.Visible = False
    End If
End Sub
'+-----------------------------+
'+  iH1_ENYmd
'+-----------------------------+
Private Sub iH1_ENYmd_GotFocus()
    PB_Look.Tag = ActiveControl.Name
    PB_Look.Visible = True
End Sub
Private Sub iH1_ENYmd_LostFocus()
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
    If TorNmGet(iH1_Torcd, 3) Then
        H1lb_Tornm = TRM_RName
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_Look.Visible = H1Mode
    PB_Chg.Visible = H1Mode
    PB_CAN.Visible = B1Mode
    PB_Prt.Visible = B1Mode
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
    On Error GoTo HeadBodyClear_ER
    For Each gObj In Me.Controls
        With gObj
            If .Name Like "i" & wbc_Enm & "_*" Then
                .Text = ""
              '  .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then
                .Value = False
            End If
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
            If .Name Like wbc_Enm & "Comb_*" Then .Clear
        End With
    Next gObj
    If HeadBody = "B" Then
        CisVsGrid.Clear
    End If
    Exit Sub
HeadBodyClear_ER:
End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    If ProcHB = "H1" Then
       If Not Head1Chk Then
          Message.Visible = False
          GoTo ReturnPress_Ed
       End If
       ProcHB = "B1"
       Call DispChange(ProcHB)
       vsGrid1.SetFocus
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
       Call CallProc
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
    
    B1lb_DispSu = ""
    
    If H1Op_TanKB(0).Value Then
        gStr = Replace(Back_Date.Caption, " ", "")
        gStr = Replace(gStr, "　", "")
        
        If Trim(iH1_SNYmd) = "" Then
           With CisFun
               .MB_Lines = 4
               .MB_MSG(2) = "    " & gStr & "開始を入力して下さい            "
               .MB_Title = ""
               .MB_Button = Error
               .MBOX
           End With
           iH1_SNYmd.SetFocus
           Exit Function
        End If
        If CisFun.Val2(iH1_SNYmd.cDay) <> 0 Then
           If Not iH1_SNYmd.cChkResult Then
              With CisFun
                  .MB_Lines = 4
                  .MB_MSG(2) = "    " & gStr & "開始を正しい日付にて入力して下さい            "
                  .MB_Title = ""
                  .MB_Button = Error
                  .MBOX
              End With
              iH1_SNYmd.SetFocus
              Exit Function
           End If
        End If
        If Trim(iH1_ENYmd) <> "" Then
           If CisFun.Val2(iH1_SNYmd.cDay) <> 0 And _
              CisFun.Val2(iH1_ENYmd.cDay) = 0 Then
              With CisFun
                  .MB_Lines = 4
                  .MB_MSG(2) = "    " & gStr & "終了を年月日にて入力して下さい            "
                  .MB_Title = ""
                  .MB_Button = Error
                  .MBOX
              End With
              iH1_ENYmd.SetFocus
              Exit Function
           End If
           If CisFun.Val2(iH1_ENYmd.cDay) <> 0 Then
              If Not iH1_ENYmd.cChkResult Then
                 With CisFun
                     .MB_Lines = 4
                     .MB_MSG(2) = "    " & gStr & "終了を正しい日付にて入力して下さい            "
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                 End With
                 iH1_ENYmd.SetFocus
                 Exit Function
              End If
              If iH1_SNYmd > iH1_ENYmd Then
                 With CisFun
                     .MB_Lines = 5
                     .MB_MSG(2) = "    " & gStr & "終了を正しい日付にて入力して下さい            "
                     .MB_MSG(4) = "                【 開始　＞　終了 】"
                     .MB_Title = ""
                     .MB_Button = Error
                     .MBOX
                 End With
                 iH1_ENYmd.SetFocus
                 Exit Function
              End If
            End If
        End If
    End If
    
    MessageDisp.Caption = "【単価決定処理中】しばらくお待ち下さい"
    MessageDisp.BackColor = &HC0FFC0
    Message.BackColor = &HC0FFC0
    Message.Visible = True: DoEvents
    If Not TankaProc Then
       Message.Visible = False: DoEvents
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
       iH1_SNYmd.SetFocus
       Exit Function
    End If
    gSL_Select = "SELECT HN.表示品番,TW.取引先,"
    gSL_Select = gSL_Select & "TW.品番,"
    gSL_Select = gSL_Select & "MIN(TW.日付) 日付から,"
    gSL_Select = gSL_Select & "MAX(TW.日付) 日付まで,"
    gSL_Select = gSL_Select & "SUM(TW.数量) 数量,"
    gSL_Select = gSL_Select & " ISNULL(TR.略称,'') 取引先名"
    gSL_Select = gSL_Select & "  From 単価未登録一覧ワーク TW"
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "    ON HN.品番 = TW.品番"
    gSL_Select = gSL_Select & "   LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "    ON  TR.取引先CD = TW.取引先"
    gSL_Select = gSL_Select & "    And TR.取引先区分 = 0"
    gSL_Select = gSL_Select & " Where TW.処理端末 = Host_Name() "
    gSL_Select = gSL_Select & " GROUP BY TW.取引先,HN.表示品番,TW.品番,TR.略称"
    gSL_Select = gSL_Select & " Order By TW.取引先,TW.品番"
    
    If Not TMWRead(gSL_Select) Then
       Call TMWClose
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
       iH1_SNYmd.SetFocus
       Exit Function
    End If
    
    MessageDisp.Caption = "【表示中】しばらくお待ち下さい"
    MessageDisp.BackColor = &HFFC0C0
    Message.BackColor = &HFFC0C0
    Message.Visible = True
    DoEvents
    
    Call GridSet
    
    With CisDB
         .SQL = "DELETE 単価未登録一覧ワーク"
         .SQL = .SQL & " Where 処理端末 = Host_Name() "
         .DBExec
    End With
    
    Message.Visible = False
    
    Head1Chk = True
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'//////////////////////////////////////
'//
'//     売上単価未決データ作成
'//
'//////////////////////////////////////
Private Function TankaProc() As Boolean
    Dim wSYMD       As String
    Dim wEYMD       As String
    
    TankaProc = False
    
    On Error GoTo TankaProc_ER
    gStr = ""
    With CisDB
        .SQL = "売上単価未決作成"
        .StoadoCount = 6
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = Input用
        .ParaNo = 3: .ParaIO = Input用
        .ParaNo = 4: .ParaIO = Input用
        .ParaNo = 5: .ParaIO = OutPut用
        .ParaNo = 6: .ParaIO = OutPut用
        
        wSYMD = iH1_SNYmd
        If CisFun.Val2(iH1_SNYmd.cDay) = 0 Then
           wSYMD = wSYMD & "00"
        End If
        If Trim(iH1_ENYmd) = "" Then
           wEYMD = iH1_SNYmd
           If CisFun.Val2(iH1_SNYmd.cDay) = 0 Then
              wEYMD = wEYMD & "99"
           End If
        Else
           wEYMD = iH1_ENYmd
           If CisFun.Val2(iH1_ENYmd.cDay) = 0 Then
              wEYMD = wEYMD & "99"
           End If
        End If
        
        .ParaNo = 1
        If H1Op_TanKB(0).Value Then
           .ParaValue = 0               ' 単価未決
        Else
           .ParaValue = 1               ' 出荷Or検収
        End If
        .ParaNo = 2: .ParaValue = wSYMD
        .ParaNo = 3: .ParaValue = wEYMD
        .ParaNo = 4: .ParaValue = Trim(iH1_Torcd)
        
        If Not .DBStored Then GoTo TankaProc_ER
        
        ' ストアド処理エラー
        .ParaNo = 5: gCnt1 = CisFun.Val2(.ParaValue)
        If gCnt1 <> 0 Then GoTo TankaProc_ER
        ' 未決件数確認
        .ParaNo = 6: gCnt2 = CisFun.Val2(.ParaValue)
        If gCnt2 = 0 Then Exit Function
        
    End With

    TankaProc = True
    On Error GoTo 0
    Exit Function
TankaProc_ER:
    Select Case gLong
        Case 1: gStr = "単価未登録一覧ワーク 削除エラー"
        Case 13: gStr = "単価未登録一覧ワーク 作成エラー"
        Case Else
             gStr = "その他ｴﾗｰ:" & gStr
    End Select
    gStr = StrConv(Format(gLong, "00"), vbWide) & ":" & gStr
    Call CisFun.ErrorBox(gStr)

End Function
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇
'◇                             グ　リ　ッ　ド　処　理　関　連
'◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
'◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇◆◇
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
        Call .FixedSet(中中, 右中, あり, 3, 0, "№")
        Call .FixedSet(中中, 中中, あり, 5, 1, "処理")
        Call .FixedSet(中中, 左中, あり, 30, 2, "納 入 先")
        Call .FixedSet(中中, 左中, あり, 30, 3, "品       番")
        If mJiseki <> "1" Then
           Call .FixedSet(中中, 左中, あり, 26, 4, "検　収　日")
        Else
           Call .FixedSet(中中, 左中, あり, 26, 4, "出　荷　日")
        End If
        Call .FixedSet(中中, 右中, あり, 10, 5, "数量合計")

        Call .FixedSet(中中, 右中, あり, 0, 6, "納入先CD")
        Call .FixedSet(中中, 右中, あり, 0, 7, "品番CD")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'        Call .FixedSet(中中, 左中, なし, 0,8,"Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5")
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
' 抽出処理 初期ｾｯﾄ
    B1lb_DispSu = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
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
        
        ID = 0
        Do Until Not TMW_RDSTS
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            
            .TextMatrix(ID, GridCol(6)) = TMW.取引先
            .TextMatrix(ID, GridCol(7)) = TMW.品番
            '納入先
            gStr = TMW.取引先
            If Trim(TMW.取引先名) <> "" Then
               gStr = gStr & ":" & TMW.取引先名
            End If
            .TextMatrix(ID, GridCol(2)) = gStr
            ' 品番
            .TextMatrix(ID, GridCol(3)) = TMW.表示品番
            ' 日付
            With CisFun
                .DateE_BefVal = TMW.日付から
                .DateEdit
                gStr = .DateE_AfVal
            
                If TMW.日付から = TMW.日付まで Then
                   gStr = gStr & "             "
                Else
                   .DateE_BefVal = TMW.日付まで
                   .DateEdit
                   gStr = gStr & " ～ " & .DateE_AfVal
                End If
            End With
            .TextMatrix(ID, GridCol(4)) = " " & gStr
            
            ' 数量計
            .TextMatrix(ID, GridCol(5)) = Format(TMW.数量, "#,###")
            Call TMWReadNext
        Loop
        Call TMWClose
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .Redraw = True
    End With
End Sub
'******************************************'
'*****      行選択（ダブルクリック）  *****
'******************************************'
Private Sub vsGrid1_DblClick()
    Call CallProc
End Sub

'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    Call CallProc
End Sub


Private Sub CallProc()
    With vsGrid1
        If Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(1))) <> "" Then Exit Sub
        If Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(2))) = "" Then Exit Sub
    
        RV_Torcd = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(6)))
        RV_Hinbn = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(7)))
        RV_SYmd = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(4)))
        RV_EYmd = Trim(.TextMatrix(.Row, CisVsGrid.FixedGet(5)))
        Rv_Call3 = "X"
        Rv_Call3Rtn = 0
        
        CXM0160.Show vbModal
        Unload CXM0160
        Set CXM0160 = Nothing
    
        If Rv_Call3Rtn = 1 Then
           .TextMatrix(.Row, CisVsGrid.FixedGet(1)) = "済み"
        End If
    End With
End Sub

