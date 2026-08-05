VERSION 5.00
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CCT0070 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "内示テキストファイルレイアウト設定"
   ClientHeight    =   11175
   ClientLeft      =   1920
   ClientTop       =   2880
   ClientWidth     =   15255
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
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11175
   ScaleWidth      =   15255
   WindowState     =   2  '最大化
   Begin VB.PictureBox H2_Area1 
      BeginProperty Font 
         Name            =   "ＭＳ Ｐゴシック"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1035
      Left            =   3060
      ScaleHeight     =   975
      ScaleWidth      =   1470
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   1230
      Visible         =   0   'False
      Width           =   1530
      Begin Cis3D_v60.CIS3D CIS3D9 
         Height          =   900
         Left            =   45
         Top             =   45
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   1588
         ForeColor       =   16711680
         Caption         =   "複写元№"
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
         Begin CisText_V60.CisText iH2_CpyNo 
            Height          =   375
            Left            =   480
            TabIndex        =   1
            Top             =   390
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   661
            cFormat         =   "00"
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
            Text            =   "00"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
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
      Height          =   795
      Left            =   5850
      ScaleHeight     =   735
      ScaleWidth      =   7815
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   1470
      Width           =   7875
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   5205
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   6060
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   6915
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_NEXT 
         Height          =   585
         Left            =   4350
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   3495
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   1785
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_ADD 
         Height          =   585
         Left            =   75
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
         Left            =   930
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
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
      Begin CisBtn_60.CisBtn PB_CPY 
         Height          =   585
         Left            =   2640
         Top             =   75
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   1032
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ ゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PFSize          =   0
         BtnType         =   4
         BtnCaption      =   "複写"
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
         PFCaption       =   "( F5 )"
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
         PFBackColor     =   16711935
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
      Height          =   1035
      Left            =   2025
      ScaleHeight     =   975
      ScaleWidth      =   930
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   1230
      Width           =   990
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   900
         Left            =   45
         Top             =   45
         Width           =   840
         _ExtentX        =   1482
         _ExtentY        =   1588
         ForeColor       =   16711680
         Caption         =   "№"
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
         Begin CisText_V60.CisText iH1_No 
            Height          =   375
            Left            =   210
            TabIndex        =   0
            Top             =   390
            Width           =   405
            _ExtentX        =   714
            _ExtentY        =   661
            cFormat         =   "00"
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
            Text            =   "00"
            MaxLength       =   2
            cDataType       =   1
            cDataReplace    =   1
            cFaZero         =   0
            cFdAutoFormat   =   1
            cGFormat        =   "##"
            cILength        =   2
         End
      End
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
      Height          =   8415
      Left            =   2010
      ScaleHeight     =   8355
      ScaleWidth      =   11655
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2340
      Width           =   11715
      Begin VB.PictureBox Picture1 
         Height          =   8250
         Left            =   60
         ScaleHeight     =   8190
         ScaleWidth      =   11490
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   60
         Width           =   11550
         Begin VB.Frame Frame2 
            Caption         =   "データベース設定"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7560
            Left            =   4005
            TabIndex        =   21
            Top             =   600
            Width           =   6555
            Begin Cis3D_v60.CIS3D CIS3D1 
               Height          =   270
               Left            =   2190
               Top             =   225
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   476
               BackColor       =   -2147483633
               ForeColor       =   255
               Caption         =   "※【品番】【受入】は必須入力"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483633
               cBoderColor2    =   -2147483633
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
            Begin Cis3D_v60.CIS3D PB_Text 
               Height          =   780
               Left            =   135
               Top             =   240
               Width           =   2025
               _ExtentX        =   3572
               _ExtentY        =   1376
               Caption         =   "TEXTﾚｲｱｳﾄ確認"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderWidth     =   5
               cAlingnment     =   7
               cButton         =   -1  'True
            End
            Begin vsFlexLib.vsFlexArray VsGrid4 
               Height          =   6375
               Left            =   150
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   1050
               Width           =   5730
               _Version        =   196608
               _ExtentX        =   10107
               _ExtentY        =   11245
               _StockProps     =   228
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   1
               ConvInfo        =   1418783674
               FormatString    =   "　|                 変換先項目|連番|開始|桁数|変換先属性|変換先桁数"
               Rows            =   77
               Cols            =   14
               BackColor       =   12648447
               BackColorFixed  =   12632256
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   2
               RowHeightMin    =   300
               FillStyle       =   1
               GridColor       =   0
               SelectionMode   =   1
               AutoResizeColumns=   0   'False
               AutoSizeMode    =   1
               Begin CisText_V60.CisText iB1_Start 
                  Height          =   300
                  Left            =   4215
                  TabIndex        =   7
                  Top             =   285
                  Width           =   585
                  _ExtentX        =   1032
                  _ExtentY        =   529
                  cFormat         =   "####"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "####"
                  MaxLength       =   4
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "####"
                  cILength        =   4
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_Keta 
                  Height          =   300
                  Left            =   4785
                  TabIndex        =   8
                  Top             =   285
                  Width           =   600
                  _ExtentX        =   1058
                  _ExtentY        =   529
                  cFormat         =   "####"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "####"
                  MaxLength       =   4
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "####"
                  cILength        =   4
                  cAlignment      =   1
               End
               Begin CisText_V60.CisText iB1_LayNo 
                  Height          =   300
                  Left            =   3645
                  TabIndex        =   6
                  Top             =   285
                  Width           =   585
                  _ExtentX        =   1032
                  _ExtentY        =   529
                  cFormat         =   "####"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Text            =   "####"
                  MaxLength       =   4
                  cDataType       =   1
                  cDataReplace    =   1
                  cFdAutoFormat   =   1
                  cGFormat        =   "####"
                  cILength        =   4
                  cAlignment      =   1
               End
            End
            Begin Cis3D_v60.CIS3D CIS3D7 
               Height          =   540
               Left            =   2190
               Top             =   510
               Width           =   4215
               _ExtentX        =   7435
               _ExtentY        =   953
               BackColor       =   -2147483633
               ForeColor       =   255
               Caption         =   "※【内示年月】指定の場合、ｶﾚﾝﾀﾞﾚｺｰﾄﾞ必須"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   9.75
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   -2147483633
               cBoderColor2    =   -2147483633
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
               Begin Cis3D_v60.CIS3D CIS3D8 
                  Height          =   270
                  Left            =   300
                  Top             =   255
                  Width           =   4395
                  _ExtentX        =   7752
                  _ExtentY        =   476
                  BackColor       =   -2147483633
                  ForeColor       =   255
                  Caption         =   "未指定の場合、変換時自動取得"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9.75
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cBoderColor1    =   -2147483633
                  cBoderColor2    =   -2147483633
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderStyle     =   2
               End
            End
         End
         Begin Cis3D_v60.CIS3D CIS3D4 
            Height          =   480
            Left            =   60
            Top             =   45
            Width           =   5100
            _ExtentX        =   8996
            _ExtentY        =   847
            ForeColor       =   16711680
            Caption         =   "変換名称"
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
            cPositionX      =   100
            Begin CisText_V60.CisText iB1_FName 
               Height          =   375
               Left            =   1125
               TabIndex        =   5
               Top             =   60
               Width           =   3975
               _ExtentX        =   7011
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
               Text            =   "XXXXXXXXXXXXXXXXXXXXXXXXXXXXXX"
               MaxLength       =   30
               IMEMode         =   1
            End
         End
         Begin VB.Frame Frame1 
            Caption         =   "テキストレイアウト"
            BeginProperty Font 
               Name            =   "ＭＳ ゴシック"
               Size            =   12
               Charset         =   128
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   7560
            Left            =   90
            TabIndex        =   20
            Top             =   600
            Width           =   3885
            Begin VB.PictureBox Picture2 
               BackColor       =   &H00000000&
               Height          =   1035
               Left            =   90
               ScaleHeight     =   975
               ScaleWidth      =   3630
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   255
               Width           =   3690
               Begin Cis3D_v60.CIS3D CIS3D11 
                  Height          =   435
                  Left            =   45
                  Top             =   495
                  Width           =   3540
                  _ExtentX        =   6244
                  _ExtentY        =   767
                  ForeColor       =   16711680
                  Caption         =   "テキストレコード長"
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
                  cPositionX      =   50
                  Begin CisText_V60.CisText iB1_Len 
                     Height          =   330
                     Left            =   2880
                     TabIndex        =   11
                     Top             =   45
                     Width           =   615
                     _ExtentX        =   1085
                     _ExtentY        =   582
                     cFormat         =   "####"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   11.25
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Text            =   "####"
                     MaxLength       =   4
                     cDataType       =   1
                     cDataReplace    =   1
                     cFdAutoFormat   =   1
                     cGFormat        =   "####"
                     cILength        =   4
                     cAlignment      =   1
                  End
               End
               Begin Cis3D_v60.CIS3D CIS3D2 
                  Height          =   435
                  Left            =   45
                  Top             =   45
                  Width           =   3540
                  _ExtentX        =   6244
                  _ExtentY        =   767
                  ForeColor       =   16711680
                  Caption         =   "ﾃﾞｰﾀ基本識別"
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
                  cPositionX      =   50
                  Begin CisText_V60.CisText iB1_KSikibetu 
                     Height          =   330
                     Left            =   2715
                     TabIndex        =   10
                     Top             =   45
                     Width           =   780
                     _ExtentX        =   1376
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
                     Text            =   "XXXXXXXXXXXXXXXXXXXX"
                     MaxLength       =   20
                  End
                  Begin Cis3D_v60.CIS3D CIS3D10 
                     Height          =   345
                     Left            =   2130
                     Top             =   45
                     Width           =   570
                     _ExtentX        =   1005
                     _ExtentY        =   609
                     ForeColor       =   16711680
                     Caption         =   "桁目～"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   9
                        Charset         =   128
                        Weight          =   400
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     cBoderColor1    =   12632256
                     cBoderColor2    =   12632256
                     cFont3DColor1   =   8421504
                     cFont3DColor2   =   -2147483643
                     cFont3D         =   2
                     cAlingnment     =   7
                     cBoderStyle     =   2
                  End
                  Begin CisText_V60.CisText iB1_KSStart 
                     Height          =   345
                     Left            =   1560
                     TabIndex        =   9
                     Top             =   45
                     Width           =   480
                     _ExtentX        =   847
                     _ExtentY        =   609
                     cFormat         =   "###"
                     BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                        Name            =   "ＭＳ ゴシック"
                        Size            =   11.25
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
            End
            Begin Cis3D_v60.CIS3D CIS3D5 
               Height          =   405
               Left            =   60
               Top             =   7065
               Width           =   3735
               _ExtentX        =   6588
               _ExtentY        =   714
               ForeColor       =   16711680
               Caption         =   "識別"
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
               cPositionX      =   50
               Begin CisText_V60.CisText iB1_Sikibetu 
                  Height          =   330
                  Left            =   1605
                  TabIndex        =   14
                  Top             =   30
                  Width           =   2145
                  _ExtentX        =   3784
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
                  Text            =   "XXXXXXXXXXXXXXXXXXXX"
                  MaxLength       =   20
               End
               Begin Cis3D_v60.CIS3D CIS3D6 
                  Height          =   345
                  Left            =   1020
                  Top             =   30
                  Width           =   570
                  _ExtentX        =   1005
                  _ExtentY        =   609
                  ForeColor       =   16711680
                  Caption         =   "桁目～"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   9
                     Charset         =   128
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cBoderColor1    =   12632256
                  cBoderColor2    =   12632256
                  cFont3DColor1   =   8421504
                  cFont3DColor2   =   -2147483643
                  cFont3D         =   2
                  cAlingnment     =   7
                  cBoderStyle     =   2
               End
               Begin CisText_V60.CisText iB1_SKaisi 
                  Height          =   345
                  Left            =   540
                  TabIndex        =   13
                  Top             =   30
                  Width           =   480
                  _ExtentX        =   847
                  _ExtentY        =   609
                  cFormat         =   "###"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   11.25
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
            Begin Cis3D_v60.CIS3D Back_Cln 
               Height          =   390
               Left            =   75
               Top             =   6675
               Width           =   3720
               _ExtentX        =   6562
               _ExtentY        =   688
               ForeColor       =   16711680
               Caption         =   "カレンダレコード区分"
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
               cPositionX      =   50
               Begin VB.CheckBox B1Chk_Cln 
                  BackColor       =   &H00C0C0C0&
                  Caption         =   "Yes"
                  Height          =   270
                  Left            =   2625
                  TabIndex        =   12
                  Top             =   60
                  Width           =   1050
               End
            End
            Begin VB.PictureBox Back_LayOut 
               Height          =   405
               Left            =   90
               ScaleHeight     =   345
               ScaleWidth      =   3645
               TabIndex        =   17
               TabStop         =   0   'False
               Top             =   1350
               Width           =   3705
               Begin Cis3D_v60.CIS3D PB_LayOut 
                  Height          =   345
                  Index           =   0
                  Left            =   0
                  Top             =   0
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   609
                  Caption         =   "①"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   3
                  cAlingnment     =   7
                  cButton         =   -1  'True
               End
               Begin Cis3D_v60.CIS3D PB_LayOut 
                  Height          =   345
                  Index           =   1
                  Left            =   1215
                  Top             =   0
                  Width           =   1200
                  _ExtentX        =   2117
                  _ExtentY        =   609
                  Caption         =   "②"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   3
                  cAlingnment     =   7
                  cButton         =   -1  'True
               End
               Begin Cis3D_v60.CIS3D PB_LayOut 
                  Height          =   345
                  Index           =   2
                  Left            =   2415
                  Top             =   0
                  Width           =   1230
                  _ExtentX        =   2170
                  _ExtentY        =   609
                  Caption         =   "③"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "ＭＳ ゴシック"
                     Size            =   12
                     Charset         =   128
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  cFont3DColor1   =   16777215
                  cFont3DColor2   =   8421504
                  cBoderWidth     =   3
                  cAlingnment     =   7
                  cButton         =   -1  'True
               End
            End
            Begin vsFlexLib.vsFlexArray VsGrid3 
               Height          =   4860
               Left            =   90
               TabIndex        =   16
               TabStop         =   0   'False
               Top             =   1785
               Width           =   3720
               _Version        =   196608
               _ExtentX        =   6562
               _ExtentY        =   8572
               _StockProps     =   228
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Appearance      =   1
               ConvInfo        =   1418783674
               FormatString    =   "ｶﾗﾑ |    項目名        |桁数|"
               Rows            =   1001
               Cols            =   4
               BackColor       =   12648384
               BackColorFixed  =   12632256
               BackColorBkg    =   12632256
               FocusRect       =   0
               HighLight       =   2
               RowHeightMin    =   300
               GridColor       =   0
               SelectionMode   =   1
               MergeCells      =   2
               AutoResizeColumns=   0   'False
               AutoSizeMode    =   1
            End
         End
      End
   End
   Begin VB.Timer Timer1 
      Left            =   2640
      Top             =   270
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
      Height          =   285
      IMEMode         =   3  'ｵﾌ固定
      Left            =   0
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   480
      Width           =   225
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
      Caption         =   "【 内示テキストファイルレイアウト設定  】"
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
         Top             =   90
         Width           =   885
         _ExtentX        =   1561
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
         Left            =   11880
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
   End
End
Attribute VB_Name = "CCT0070"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   TEXTﾌｧｲﾙ変換設定
'**       フォームID    :   CCT0040
'**       処理概要      :
'**
'**       作  成  日    :   2005/12/05  By CIS
'**       変  更  日    :
'**
'**********************************************************'

    Dim FormAct         As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB          As String * 2   ' 入力域切替
    Dim SyoriKB         As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM         As String       ' 処理名称(新規･修正･削除 等)
   
    Dim mDelColor       As Long         ' 削除時、ｸﾞﾘｯﾄﾞ色
    Dim mUnDelColor     As Long         ' 削除、設定以外
    Dim mSetColor       As Long         ' ｸﾞﾙｰﾌﾟ設定時、ｸﾞﾘｯﾄﾞ色(BackColor)
    Dim mGrpColor       As Long         ' ｸﾞﾙｰﾌﾟ設定時、ｸﾞﾘｯﾄﾞ色
    Dim mWSID           As String       ' ワークステーション名

    Dim CisFrm          As CISFormContorl
    Dim CisVsGrid2      As New CisVsGrid3
    Dim CisVsGrid3      As New CisVsGrid3

    Dim mX              As Single
    Dim mY              As Single
    Dim mItemMax        As Long
    Dim mItemFlg        As Boolean
    
    Dim mSaveRow        As Long
    Dim mOccColor       As Long
    Dim mOccUnColor     As Long
    Dim mVsGrid2Color   As Long

    Dim Now_LayoutNo    As Integer
    Dim Now_VS4Row      As Integer
    
    Dim ActObj          As Object
    Dim ActRow          As Integer

    Dim mItem1()        As String
    Dim mKeta1()        As Long
    Dim mStart1()       As Long
    Dim mCnt1           As Long
    Dim mItem2()        As String
    Dim mKeta2()        As Long
    Dim mStart2()       As Long
    Dim mCnt2           As Long
    Dim mItem3()        As String
    Dim mKeta3()        As Long
    Dim mStart3()       As Long
    Dim mCnt3           As Long

    Dim mSKaisi(2)      As Long
    Dim mSikibetu(2)    As String
    Dim mClnFlg         As Byte

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
    Call IniGet
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
    
'
    Dummy.Left = -1000
    Dummy.Enabled = False
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
'+---------------------+
'+ VSGrid 初期設定  +
'+---------------------+
' 項目追加　ｸﾞﾘｯﾄﾞ初期設定
' ﾚｲｱｳﾄ設定
    With CisVsGrid2
        Set .GridObj = VsGrid3              ' ｸﾞﾘｯﾄﾞｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .FixedHeight = 300                  ' 固定行(見出し)高さ設定(省略時は300)
        .PatanMax = 0                       ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ最大数
        .PatanInit = 0                      ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ初期値
' ﾌﾟﾛｸﾞﾗﾑ作成時、以下のﾌﾗｸﾞをｾｯﾄし、ｸﾘｯﾌﾟﾎﾞｰﾄﾞよりｺｰﾄﾞの貼り付けを行う。(GridInit)
        .InitGet = True                   ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ
        .Init                               ' ｸﾞﾘｯﾄﾞ初期処理 実行
    End With
    Call GridInit3                          ' ｸﾞﾘｯﾄﾞ初期表示
    
    ' 変換テキスト項目設定
    With CisVsGrid3
        Set .GridObj = VsGrid4              ' ｸﾞﾘｯﾄﾞｵﾌﾞｼﾞｪｸﾄ名 ｾｯﾄ
        .FixedHeight = 300                  ' 固定行(見出し)高さ設定(省略時は300)
        .PatanMax = 2                       ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ最大数
        .PatanInit = 0                      ' ﾚｲｱｳﾄﾊﾟﾀｰﾝ初期値
' ﾌﾟﾛｸﾞﾗﾑ作成時、以下のﾌﾗｸﾞをｾｯﾄし、ｸﾘｯﾌﾟﾎﾞｰﾄﾞよりｺｰﾄﾞの貼り付けを行う。(GridInit)
        .InitGet = False                   ' ｸﾞﾘｯﾄﾞ内容初期設定 ｸﾘｯﾌﾟﾎﾞｰﾄﾞ作成ﾌﾗｸﾞ
        .Init                               ' ｸﾞﾘｯﾄﾞ初期処理 実行
    End With
    Call GridInit4                          ' ｸﾞﾘｯﾄﾞ初期表示

'   #------------------#
'   #    初期入力設定   #
'   #------------------#
    ProcHB = "H1"
    Call DispChange(ProcHB)
'   #------------------#
'   #  処理区分初期表示 #
'   #------------------#
    SyoriKB = "ADD"
    Call SyoriName(SyoriKB)

'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True, x1024y768)
    
'   コンピュータ名の獲得
    mWSID = CisFun.WSID
    '使用できない文字を削除
    mWSID = Replace(mWSID, "-", "")
    mWSID = Replace(mWSID, "*", "")
    mWSID = Replace(mWSID, "+", "")
    mWSID = Replace(mWSID, "/", "")
'   レイアウト表示用ワーク作成
    Call Work_Crt
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
        
    RV_Rtn = False
    
    Now_LayoutNo = 0
    Now_VS4Row = 1
    
    Erase mSKaisi: Erase mSikibetu
    mClnFlg = 0
    
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
           Case vbKeyF1:       If PB_ADD.Visible Then Call PB_ADD_Click               '【新規】
           Case vbKeyF2:       If PB_MNT.Visible Then Call PB_MNT_Click               '【修正】
           Case vbKeyF3:       If PB_DEL.Visible Then Call PB_DEL_Click               '【削除】
           Case vbKeyF5:       If PB_CPY.Visible Then Call PB_CPY_Click             '【複写】
           Case vbKeyPageUp:   If PB_BACK.Visible Then Call PB_Back_Click             '【後退】
           Case vbKeyPageDown: If PB_NEXT.Visible Then Call PB_Next_Click             '【前進】
           Case vbKeyEscape
                               If PB_CAN.Visible Then Call PB_CAN_Click               '【取消】
           Case vbKeyF12:      If PB_END.Visible Then Call PB_END_Click               '【終了】
           Case vbKeyReturn
                               If PB_ENT.Visible Then                                 '【入力】
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
'   レイアウト表示用ワーク削除
    Call Work_DEL
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
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      Call VsGridToText(1)
      iH1_No.SetFocus
      GoTo PB_CAN_Ed
   End If

PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    新  規  キ ー(F1)       +
'+----------------------------+
Private Sub PB_ADD_Click()
    If SyoriKB <> "ADD" Then
        SyoriKB = "ADD"
        Call SyoriName(SyoriKB)
    End If
    iH1_No.SetFocus
End Sub
'+----------------------------+
'+    修  正  キ ー(F2)       +
'+----------------------------+
Private Sub PB_MNT_Click()
    If SyoriKB <> "MNT" Then
        SyoriKB = "MNT"
        Call SyoriName(SyoriKB)
    End If
    iH1_No.SetFocus
End Sub
'+----------------------------+
'+    削  除  キ ー(F3)       +
'+----------------------------+
Private Sub PB_DEL_Click()
    If SyoriKB <> "DEL" Then
        SyoriKB = "DEL"
    Else
        SyoriKB = ""
    End If
    Call SyoriName(SyoriKB)
    iH1_No.SetFocus
End Sub
'+----------------------------+
'+    複　写  キ ー(F5)       +
'+----------------------------+
Private Sub PB_CPY_Click()
    If SyoriKB <> "CPY" Then
        SyoriKB = "CPY"
        ProcHB = "H1"
        Call DispChange(ProcHB)
    End If
    Call SyoriName(SyoriKB)
End Sub
Private Sub iH1_No_GotFocus()
    iH1_No.Tag = "iH1_No"
End Sub
Private Sub iH1_No_LostFocus()
    iH1_No.Tag = ""
End Sub
'+----------------------------+
'+    後  退  キ ー(PageUp  ) +
'+----------------------------+
Private Sub PB_Back_Click()
    Call CisFrm.MousePT(11)
    
    gSL_Select = "select * from TEXT設定テーブル"
    gSL_Select = gSL_Select & " where レイアウト番号 < "
    gSL_Select = gSL_Select & " " & CisFun.Val2(iH1_No)
    gSL_Select = gSL_Select & " order by レイアウト番号 desc"
    Call DBInput("Back")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_No.SetFocus
    Dummy.Enabled = False
    
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    前  進  キ ー(PageDown) +
'+----------------------------+
Private Sub PB_Next_Click()
    
    Call CisFrm.MousePT(11)
    
    gSL_Select = "select * from TEXT設定テーブル"
    gSL_Select = gSL_Select & " where レイアウト番号 > "
    gSL_Select = gSL_Select & " " & CisFun.Val2(iH1_No)
    gSL_Select = gSL_Select & " order by レイアウト番号"
    Call DBInput("Next")
    Dummy.Enabled = True: Dummy.SetFocus
    iH1_No.SetFocus
    Dummy.Enabled = False
    
    Call CisFrm.MousePT(1)
End Sub
'+---------------------------------------------------------+
'+          グリッド編集エリア
'+---------------------------------------------------------+
'+---------------------------------------------+
'+  iB1_LayNo    連番
'+---------------------------------------------+
Private Sub iB1_LayNo_GotFocus()
    Set ActObj = iB1_LayNo
End Sub
Private Sub iB1_LayNo_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_LayNo_LostFocus()
    With VsGrid4
        ActRow = CisFun.Val2(iB1_LayNo.Tag)
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 2) = CisFun.Val2(iB1_LayNo)
    End With
End Sub
'+---------------------------------------------+
'+  iB1_Start    開始
'+---------------------------------------------+
Private Sub iB1_Start_GotFocus()
    Set ActObj = iB1_Start
End Sub
Private Sub iB1_Start_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Start_LostFocus()
    With VsGrid4
        ActRow = CisFun.Val2(iB1_LayNo.Tag)
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 3) = CisFun.Val2(iB1_Start)
    End With
End Sub
'+---------------------------------------------+
'+  iB1_Keta    桁数
'+---------------------------------------------+
Private Sub iB1_Keta_GotFocus()
    Set ActObj = iB1_Keta
End Sub
Private Sub iB1_Keta_KeyDown(KeyCode As Integer, Shift As Integer)
    Call Grid_RowChange(KeyCode)
End Sub
Private Sub iB1_Keta_LostFocus()
    With VsGrid4
        ActRow = CisFun.Val2(iB1_LayNo.Tag)
        If ActRow > .Rows - 1 Then Exit Sub
        .TextMatrix(ActRow, 4) = CisFun.Val2(iB1_Keta)
    End With
End Sub
''+---------------------------------------------+
''+  iB1_Kotei    固定文字
''+---------------------------------------------+
'Private Sub iB1_Kotei_GotFocus()
'    Set ActObj = iB1_Kotei
'End Sub
'Private Sub iB1_Kotei_KeyDown(KeyCode As Integer, Shift As Integer)
'    Call Grid_RowChange(KeyCode)
'End Sub
'Private Sub iB1_Kotei_LostFocus()
'    With VsGrid4
'        ActRow = CisFun.Val2(iB1_LayNo.Tag)
'        If ActRow > .Rows - 1 Then Exit Sub
'        .TextMatrix(ActRow, 5) = RTrim(iB1_Kotei)
'    End With
'End Sub
Private Sub Grid_RowChange(KeyCode As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown And _
        KeyCode <> vbKeyPageUp And KeyCode <> vbKeyPageDown Then Exit Sub
    
'テキストからグリッドに退避
    Call TextToVsGrid(Val(iB1_LayNo.Tag))
    gInt = VsGrid4.Row
    '↑
    If KeyCode = vbKeyUp Then
        If gInt - 1 < 1 Then GoTo Grid_RowChange_End
        gInt = gInt - 1
    End If
    '↓
    If KeyCode = vbKeyDown Then
        If gInt + 1 > VsGrid4.Rows - 1 Then
            If Trim(VsGrid4.TextMatrix(VsGrid4.Rows - 1, 1)) = "" Then GoTo Grid_RowChange_End
            VsGrid4.Rows = VsGrid4.Rows + 1
        End If
        gInt = gInt + 1
    End If
    'PageUp
    If KeyCode = vbKeyPageUp Then
        If gInt - CisVsGrid3.DispMax < 1 Then
            gInt = 1
        Else
            gInt = gInt - CisVsGrid3.DispMax
        End If
    End If
    'PageDown
'    If KeyCode = vbKeyPageDown Then
'        If gInt + CisVsGrid3.DispMax > Ttl_Kensu - 1 Then
'            gInt = Ttl_Kensu - 1
'        Else
'            gInt = gInt + CisVsGrid3.DispMax
'        End If
'    End If
    
    VsGrid4.Row = gInt
'グリッドからテキストに表示
    Call VsGridToText(VsGrid4.Row)
    
Grid_RowChange_End:
    KeyCode = 0
    Me.Refresh
    DoEvents
End Sub
'==========================================================='
'   入力テキストからグリッドに移す
'==========================================================='
Private Sub TextToVsGrid(tRow As Integer)
    With VsGrid4
        .TextMatrix(tRow, 2) = iB1_LayNo
        .TextMatrix(tRow, 3) = iB1_Start
        .TextMatrix(tRow, 4) = iB1_Keta
'        .TextMatrix(tRow, 5) = iB1_Kotei
    End With
End Sub
'==========================================================='
'   グリッドから入力テキストに移す、入力テキストの移動
'==========================================================='
Private Sub VsGridToText(tRow As Integer)
    With VsGrid4
        '開始
        iB1_LayNo.Tag = tRow
        iB1_LayNo = RTrim(.TextMatrix(tRow, 2))
        iB1_LayNo.Move iB1_LayNo.Left, VsGrid4.CellTop, iB1_LayNo.Width, VsGrid4.CellHeight
        '開始
        iB1_Start = RTrim(.TextMatrix(tRow, 3))
        iB1_Start.Move iB1_Start.Left, VsGrid4.CellTop, iB1_Start.Width, VsGrid4.CellHeight
        '桁
        iB1_Keta = RTrim(.TextMatrix(tRow, 4))
        iB1_Keta.Move iB1_Keta.Left, VsGrid4.CellTop, iB1_Keta.Width, VsGrid4.CellHeight
        '固定文字
'        iB1_Kotei = RTrim(.TextMatrix(tRow, 5))
'        iB1_Kotei.Move iB1_Kotei.Left, VsGrid4.CellTop, iB1_Kotei.Width, VsGrid4.CellHeight
    End With
End Sub
'+-----------------------------------+
'+    レイアウト設定                  +
'+-----------------------------------+
Private Sub PB_Layout_Click(Index As Integer)
    Call CisFrm.MousePT(11)
        
    '現在の情報を退避
    mSKaisi(Now_LayoutNo) = CisFun.Val2(iB1_SKaisi)
    mSikibetu(Now_LayoutNo) = RTrim(iB1_Sikibetu)
    If Now_LayoutNo = 0 Then mClnFlg = B1Chk_Cln.Value
    
    PB_LayOut(0).BackColor = &H8000000F
    PB_LayOut(1).BackColor = &H8000000F
    PB_LayOut(2).BackColor = &H8000000F
    PB_LayOut(Index).BackColor = &HC0FFC0
    
    Select Case Index
           Case 0: Call TextGrid_Dsp1
           Case 1: Call TextGrid_Dsp2
           Case 2: Call TextGrid_Dsp3
    End Select
        
    Now_LayoutNo = Index
    
    iB1_SKaisi = Format(mSKaisi(Now_LayoutNo), "###")
    iB1_Sikibetu = RTrim(mSikibetu(Now_LayoutNo))
    If Now_LayoutNo = 0 Then
       Back_Cln.Caption = "カレンダレコード区分"
       B1Chk_Cln.Visible = True
       B1Chk_Cln.Value = mClnFlg
    Else
       Back_Cln.Caption = ""
       B1Chk_Cln.Visible = False
    End If
    
    Call CisFrm.MousePT(1)
End Sub
'+-----------------------------------+
'+    最新ﾃｷｽﾄﾚｲｱｳﾄに更新
'+-----------------------------------+
Private Sub PB_Text_Click()
    Call MoveItemLayout
    
    '現在のレイアウトＮｏを表示
    Select Case Now_LayoutNo
        Case 0: Call PB_Layout_Click(0)
        Case 1: Call PB_Layout_Click(1)
        Case 2: Call PB_Layout_Click(2)
    End Select
        
End Sub
Private Sub B1Chk_Cln_Click()
    mClnFlg = B1Chk_Cln.Value
End Sub
Private Sub B1Chk_Cln_LostFocus()
    mClnFlg = B1Chk_Cln.Value
End Sub
Private Sub iB1_SKaisi_LostFocus()
    mSKaisi(Now_LayoutNo) = CisFun.Val2(iB1_SKaisi)
End Sub
Private Sub iB1_Sikibetu_LostFocus()
    mSikibetu(Now_LayoutNo) = RTrim(iB1_Sikibetu)
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
     Case "CPY"
          SyoriNM = "複写"
          K_Sykbnm.ForeColor = &HFF00FF      '(紫)
     Case Else
          SyoriNM = ""
   End Select
   For Each gObj In Me.Controls
       If gObj.Name Like "i*" Then gObj.ForeColor = W_Color
       If gObj.Name Like "*Comb*" Then gObj.ForeColor = W_Color
'      * PF-Key 表示設定 *
       If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
          gObj.Name = "PB_DEL" Or gObj.Name = "PB_CPY" Then
           gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
       End If
   Next gObj
   If SyoriNM = "" Then
      K_Sykbnm.Visible = False
      K_Sykbnm.Caption = ""
      Exit Sub
   End If
   K_Sykbnm.Caption = SyoriNM
   K_Sykbnm.Visible = True
   If SyoriKB = "CPY" Then
      H2_Area1.Visible = True
   Else
      H2_Area1.Visible = False
   End If
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
    If SyoriKB = "CPY" Then
       H2_Area1.Enabled = H1Mode
       H2_Area1.BackColor = H1Color
    End If
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    PB_NEXT.Visible = H1Mode
    PB_BACK.Visible = H1Mode
    PB_ADD.Visible = H1Mode
    PB_MNT.Visible = H1Mode
    PB_DEL.Visible = H1Mode
    PB_CPY.Visible = H1Mode
    
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
                .Enabled = True
            End If
            If .Name Like wbc_Enm & "Chk_*" Then .Value = 0
            If .Name Like wbc_Enm & "Ck_*" Then .Value = 0
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj
    
    Call CisVsGrid2.Clear
    Call CisVsGrid3.Clear

    Now_LayoutNo = 0
    Now_VS4Row = 1
    
    Erase mSKaisi: Erase mSikibetu
    mClnFlg = 0
    
    Erase mItem1
    Erase mKeta1
    Erase mStart1
    Erase mItem2
    Erase mKeta2
    Erase mStart2
    Erase mItem3
    Erase mKeta3
    Erase mStart3
    mCnt1 = 0: mCnt2 = 0: mCnt3 = 0

End Sub
'****************************
'*      入力処理(Enter)      *
'****************************
Private Sub ReturnPress()
    Call CisFrm.MousePT(11)
    DoEvents
    If ProcHB = "H1" Then
        If Not Head1Chk Then GoTo ReturnPress_Ed
' << 削除処理 >>
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            Call HeadBodyClear("B")
            ProcHB = "H1"
            Call DispChange(ProcHB)
            Call SyoriName(SyoriKB)
        Else
            ProcHB = "B1"
            Call VsGridToText(1)
            Call DispChange(ProcHB)
            iB1_FName.SetFocus
        End If
        GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
        Call CisFrm.MousePT(11)
        If Not Body1Chk Then GoTo ReturnPress_Ed
        If Not DBPut Then GoTo ReturnPress_Ed
        ProcHB = "H1"
        Call HeadBodyClear("B")
        Call DispChange(ProcHB)
        Call VsGridToText(1)
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
    If iH1_No = "" Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  ナンバーを入力して下さい  "
            .MB_MSG(3) = "　"
            .MB_Title = " "
            .MB_Button = Error
            .MBOX
            iH1_No.SetFocus
            Exit Function
       End With
    End If
    
    gSL_Select = "SELECT * FROM TEXT設定テーブル"
    gSL_Select = gSL_Select & " Where レイアウト番号 = " & CisFun.Val2(iH1_No)
    If TXSRead(gSL_Select, 1) Then
       If SyoriKB = "ADD" Or SyoriKB = "CPY" Then
          With CisFun
               .MB_Lines = 5
               .MB_MSG(2) = "     TEXT設定　登録済       "
               .MB_MSG(4) = "　"
               .MB_Title = " "
               .MB_Button = Error
               .MBOX
               iH1_No.SetFocus
               Exit Function
          End With
       End If
    Else
       If SyoriKB = "MNT" Or SyoriKB = "DEL" Then
          With CisFun
               .MB_Lines = 5
               .MB_MSG(2) = "     TEXT設定  未登録       "
               .MB_MSG(4) = "　"
               .MB_Title = " "
               .MB_Button = Error
               .MBOX
               iH1_No.SetFocus
               Exit Function
          End With
       End If
    End If
    
    If SyoriKB = "CPY" Then
       If iH2_CpyNo = "" Then
          With CisFun
               .MB_Lines = 5
               .MB_MSG(2) = "  複写元ナンバーを入力して下さい  "
               .MB_MSG(3) = "　"
               .MB_Title = " "
               .MB_Button = Error
               .MBOX
               iH2_CpyNo.SetFocus
               Exit Function
          End With
       End If
    End If

' マスタ入力 SELECT文 ｾｯﾄ
    gSL_Select = "SELECT * FROM TEXT設定テーブル"
    If SyoriKB <> "CPY" Then
       gSL_Select = gSL_Select & " Where レイアウト番号 = " & CisFun.Val2(iH1_No)
    Else
       gSL_Select = gSL_Select & " Where レイアウト番号 = " & CisFun.Val2(iH2_CpyNo)
    End If
    
    If DBInput("INV") Then
       If SyoriKB = "ADD" Then
          With CisFun
               .MB_Lines = 5
               .MB_MSG(2) = "     TEXT設定　登録済       "
               .MB_MSG(4) = "　"
               .MB_Title = " "
               .MB_Button = Error
               .MBOX
               iH1_No.SetFocus
               Exit Function
          End With
       End If
    Else
       If SyoriKB <> "ADD" Then
          With CisFun
               .MB_Lines = 5
               .MB_MSG(2) = "     TEXT設定  未登録       "
               .MB_MSG(4) = "　"
               .MB_Title = " "
               .MB_Button = Error
               .MBOX
               iH1_No.SetFocus
               Exit Function
          End With
       End If
    End If
    
    Head1Chk = True

End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    Dim wChk_Flg        As Boolean
    Dim wSikibetu_Flg1  As Byte
    Dim wSikibetu_Flg2  As Byte
    Dim wSikibetu_Flg3  As Byte
    
    Body1Chk = False
    wChk_Flg = False
    
    If RTrim(iB1_FName) = "" Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  変換名称を入力して下さい  "
            .MB_MSG(3) = "　"
            .MB_Title = " "
            .MB_Button = Error
            .MBOX
            iB1_FName.SetFocus
            Exit Function
       End With
    End If
    If CisFun.Val2(iB1_KSStart) = 0 Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  基本識別開始位置を入力して下さい  "
            .MB_MSG(3) = "　"
            .MB_Title = " "
            .MB_Button = Error
            .MBOX
            iB1_KSStart.SetFocus
            Exit Function
       End With
    End If
    If RTrim(iB1_KSikibetu) = 0 Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  基本識別文字を入力して下さい  "
            .MB_MSG(3) = "　"
            .MB_Title = " "
            .MB_Button = Error
            .MBOX
            iB1_KSikibetu.SetFocus
            Exit Function
       End With
    End If
    
    With VsGrid4
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) <> "" Then
               If CisFun.Val2(.TextMatrix(gLong, 2)) = 0 And (CisFun.Val2(.TextMatrix(gLong, 3)) <> 0 Or CisFun.Val2(.TextMatrix(gLong, 4)) <> 0) Then
                  CisFun.MB_Lines = 4
                  CisFun.MB_MSG(2) = "【" & Trim(.TextMatrix(gLong, 1)) & "】" & " 連番未入力 "
                  CisFun.MB_Title = ""
                  CisFun.MB_Button = Error
                  CisFun.MBOX
                  iB1_LayNo.SetFocus
                  GoTo Body1Chk_Err
               End If
               If CisFun.Val2(.TextMatrix(gLong, 2)) > 3 Then
                  CisFun.MB_Lines = 4
                  CisFun.MB_MSG(2) = "【" & Trim(.TextMatrix(gLong, 1)) & "】" & " 連番上限オーバー "
                  CisFun.MB_Title = ""
                  CisFun.MB_Button = Error
                  CisFun.MBOX
                  iB1_LayNo.SetFocus
                  GoTo Body1Chk_Err
               End If
               
               If CisFun.Val2(.TextMatrix(gLong, 2)) = 1 Then wSikibetu_Flg1 = 1
               If CisFun.Val2(.TextMatrix(gLong, 2)) = 2 Then wSikibetu_Flg2 = 1
               If CisFun.Val2(.TextMatrix(gLong, 2)) = 3 Then wSikibetu_Flg3 = 1
               
               If CisFun.Val2(.TextMatrix(gLong, 3)) = 0 And CisFun.Val2(.TextMatrix(gLong, 4)) <> 0 Then
                  If RTrim(.TextMatrix(gLong, 1)) <> "納入先" Then
                     CisFun.MB_Lines = 4
                     CisFun.MB_MSG(2) = "【" & Trim(.TextMatrix(gLong, 1)) & "】" & " 開始位置未入力 "
                     CisFun.MB_Title = ""
                     CisFun.MB_Button = Error
                     CisFun.MBOX
                     iB1_Start.SetFocus
                     GoTo Body1Chk_Err
                  End If
               End If
               If CisFun.Val2(.TextMatrix(gLong, 3)) <> 0 And CisFun.Val2(.TextMatrix(gLong, 4)) = 0 Then
                  CisFun.MB_Lines = 4
                  CisFun.MB_MSG(2) = "【" & Trim(.TextMatrix(gLong, 1)) & "】" & " 桁数未入力 "
                  CisFun.MB_Title = ""
                  CisFun.MB_Button = Error
                  CisFun.MBOX
                  iB1_Keta.SetFocus
                  GoTo Body1Chk_Err
               End If
'               If (CisFun.Val2(.TextMatrix(gLong, 3)) <> 0 Or CisFun.Val2(.TextMatrix(gLong, 4)) <> 0) And Trim(.TextMatrix(gLong, 5)) <> "" Then
'                  CisFun.MB_Lines = 4
'                  CisFun.MB_MSG(2) = "【" & Trim(.TextMatrix(gLong, 1)) & "】" & " テキストレイアウトと固定文字の同時入力は不可 "
'                  CisFun.MB_Title = ""
'                  CisFun.MB_Button = Error
'                  CisFun.MBOX
'                  iB1_LayNo.SetFocus
'                  GoTo Body1Chk_Err
'               End If
                            
'               If CisFun.Val2(.TextMatrix(gLong, 3)) <> 0 Or CisFun.Val2(.TextMatrix(gLong, 4)) <> 0 Or Trim(.TextMatrix(gLong, 5)) <> "" Then
               If CisFun.Val2(.TextMatrix(gLong, 3)) <> 0 Or CisFun.Val2(.TextMatrix(gLong, 4)) <> 0 Then
                  wChk_Flg = True
               End If
               
               If RTrim(.TextMatrix(gLong, 1)) = "内示年月" Then
                  If CisFun.Val2(.TextMatrix(gLong, 2)) <> 0 Or CisFun.Val2(.TextMatrix(gLong, 3)) <> 0 Or CisFun.Val2(.TextMatrix(gLong, 4)) <> 0 Then
                     If mClnFlg = 0 Then
                        CisFun.MB_Lines = 4
                        CisFun.MB_MSG(2) = "【" & Trim(.TextMatrix(gLong, 1)) & "】" & " カレンダレコードチェックが必要です "
                        CisFun.MB_Title = ""
                        CisFun.MB_Button = Error
                        CisFun.MBOX
                        iB1_LayNo.SetFocus
                        GoTo Body1Chk_Err
                     End If
                  End If
               End If
            
'               If RTrim(.TextMatrix(gLong, 1)) = "品番" Or RTrim(.TextMatrix(gLong, 1)) = "納入先" Or RTrim(.TextMatrix(gLong, 1)) = "受入" Then
               If RTrim(.TextMatrix(gLong, 1)) = "品番" Or RTrim(.TextMatrix(gLong, 1)) = "受入" Then
                  If CisFun.Val2(.TextMatrix(gLong, 2)) = 0 Or CisFun.Val2(.TextMatrix(gLong, 3)) = 0 Or CisFun.Val2(.TextMatrix(gLong, 4)) = 0 Then
                     If mClnFlg = 0 Then
                        CisFun.MB_Lines = 4
                        CisFun.MB_MSG(2) = "【" & Trim(.TextMatrix(gLong, 1)) & "】" & " 必須入力です "
                        CisFun.MB_Title = ""
                        CisFun.MB_Button = Error
                        CisFun.MBOX
                        iB1_LayNo.SetFocus
                        GoTo Body1Chk_Err
                     End If
                  End If
               End If
               
            End If
        Next gLong
    End With
    
    If Not wChk_Flg Then
       With CisFun
            .MB_Lines = 5
            .MB_MSG(2) = "  データベース設定を行ってください  "
            .MB_MSG(3) = "　"
            .MB_Title = " "
            .MB_Button = Error
            .MBOX
            VsGrid4.SetFocus
            Exit Function
       End With
    End If
    
'    If wSikibetu_Flg1 = 1 Then
'       If mSKaisi(0) = 0 Then
'          With CisFun
'               .MB_Lines = 5
'               .MB_MSG(2) = "  テキストレイアウト①識別開始位置を入力して下さい  "
'               .MB_MSG(3) = "　"
'               .MB_Title = " "
'               .MB_Button = Error
'               .MBOX
'               iB1_SKaisi.SetFocus
'               Call PB_Layout_Click(0)
'               Exit Function
'          End With
'       End If
'       If RTrim(mSikibetu(0)) = "" Then
'          With CisFun
'               .MB_Lines = 5
'               .MB_MSG(2) = "  テキストレイアウト①識別文字を入力して下さい  "
'               .MB_MSG(3) = "　"
'               .MB_Title = " "
'               .MB_Button = Error
'               .MBOX
'               iB1_Sikibetu.SetFocus
'               Call PB_Layout_Click(0)
'               Exit Function
'          End With
'       End If
'    End If
'    If wSikibetu_Flg2 = 1 Then
'       If mSKaisi(1) = 0 Then
'          With CisFun
'               .MB_Lines = 5
'               .MB_MSG(2) = "  テキストレイアウト②識別開始位置を入力して下さい  "
'               .MB_MSG(3) = "　"
'               .MB_Title = " "
'               .MB_Button = Error
'               .MBOX
'               iB1_SKaisi.SetFocus
'               Call PB_Layout_Click(1)
'               Exit Function
'          End With
'       End If
'       If RTrim(mSikibetu(1)) = "" Then
'          With CisFun
'               .MB_Lines = 5
'               .MB_MSG(2) = "  テキストレイアウト②識別文字を入力して下さい  "
'               .MB_MSG(3) = "　"
'               .MB_Title = " "
'               .MB_Button = Error
'               .MBOX
'               iB1_Sikibetu.SetFocus
'               Call PB_Layout_Click(1)
'               Exit Function
'          End With
'       End If
'    End If
'    If wSikibetu_Flg3 = 1 Then
'       If mSKaisi(2) = 0 Then
'          With CisFun
'               .MB_Lines = 5
'               .MB_MSG(2) = "  テキストレイアウト③識別開始位置を入力して下さい  "
'               .MB_MSG(3) = "　"
'               .MB_Title = " "
'               .MB_Button = Error
'               .MBOX
'               iB1_SKaisi.SetFocus
'               Call PB_Layout_Click(2)
'               Exit Function
'          End With
'       End If
'       If RTrim(mSikibetu(2)) = "" Then
'          With CisFun
'               .MB_Lines = 5
'               .MB_MSG(2) = "  テキストレイアウト③識別文字を入力して下さい  "
'               .MB_MSG(3) = "　"
'               .MB_Title = " "
'               .MB_Button = Error
'               .MBOX
'               iB1_Sikibetu.SetFocus
'               Call PB_Layout_Click(2)
'               Exit Function
'          End With
'       End If
'    End If
    
    Body1Chk = True
    Exit Function
Body1Chk_Err:
    VsGrid4.Row = gLong
    Call VsGridToText(VsGrid4.Row)
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
    
    Call ItemsClearTXS
    If Not TXSRead(gSL_Select, 1) Then
        If InputKey = "Back" Then
            With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "     先頭レコードです            "
                 .MB_Title = "【後退】"
                 .MB_Button = OK
                 .MBOX
                 iH1_No.SetFocus
                 GoTo DBInput_Ed
            End With
        End If
        If InputKey = "Next" Then
            With CisFun
                 .MB_Lines = 4
                 .MB_MSG(2) = "     最終レコードです            "
                 .MB_Title = "【前進】"
                 .MB_Button = OK
                 .MBOX
                 iH1_No.SetFocus
                 GoTo DBInput_Ed
            End With
        End If
        Call BodySet("")
    Else
        If SyoriKB <> "CPY" Then
           iH1_No = Format(TXS.レイアウト番号, "00")
        End If
        Call BodySet("")
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
    Dim wID     As Long
'
    gSL_Select = "Select * From 名称マスタ "
    gSL_Select = gSL_Select & " Where 区分名称 = 'レイアウト情報内示'"
    If SyoriKB <> "CPY" Then
       gSL_Select = gSL_Select & "   And 値 = '" & CisFun.Val2(iH1_No) & "'"
    Else
       gSL_Select = gSL_Select & "   And 値 = '" & CisFun.Val2(iH2_CpyNo) & "'"
    End If
    If Not SYMRead(gSL_Select, 1) Then
       Call ItemsClearSYM
    End If
    
    iB1_FName = RTrim(SYM.値名称)
    iB1_Len = Format(SYM.数字3, "####")
    
    Call ItemDisp2           'DBレイアウト表示
    Call MoveItemLayout
    Call PB_Text_Click
    
End Sub
Private Sub ItemDisp2()
    Dim wID As Long
    Call DBGridDisp              '項目設定用テーブル内容表示
    
    gLong = 1
    With VsGrid4
        For wID = 1 To .Rows - 1
            If RTrim(.TextMatrix(wID, 1)) <> "" Then
            
               gSL_Select = "Select * From TEXT設定テーブル "
               If SyoriKB <> "CPY" Then
                  gSL_Select = gSL_Select & " Where レイアウト番号 = " & CisFun.Val2(iH1_No)
               Else
                  gSL_Select = gSL_Select & " Where レイアウト番号 = " & CisFun.Val2(iH2_CpyNo)
               End If
               gSL_Select = gSL_Select & " And   区分 = 11"
               gSL_Select = gSL_Select & " And   データ項目 = '" & RTrim(.TextMatrix(wID, 1)) & "'"
               
               If TXSRead(gSL_Select, 1) Then
            
                  .TextMatrix(wID, 2) = TXS.連番
                  .TextMatrix(wID, 3) = TXS.開始位置
                  .TextMatrix(wID, 4) = TXS.長さ
'                  .TextMatrix(wID, 5) = TXS.固定文字
               
               End If
            End If
        Next wID
        
        .Row = 1: .Col = 1
    End With
    
    gSL_Select = "Select * From TEXT設定テーブル "
    If SyoriKB <> "CPY" Then
       gSL_Select = gSL_Select & " Where レイアウト番号 = " & CisFun.Val2(iH1_No)
    Else
       gSL_Select = gSL_Select & " Where レイアウト番号 = " & CisFun.Val2(iH2_CpyNo)
    End If
    gSL_Select = gSL_Select & " And   区分 = 10"
    gSL_Select = gSL_Select & " Order By 連番 "
    If TXSRead(gSL_Select) Then
       wID = 0
       Do Until Not TXS_RDSTS
          mSKaisi(wID) = TXS.開始位置
          mSikibetu(wID) = TXS.項目
          
          Call TXSReadNext
          wID = wID + 1
       Loop
       Call TXSClose
    Else
       Call TXSClose
    End If
    
    iB1_SKaisi = Format(mSKaisi(0), "###")
    iB1_Sikibetu = RTrim(mSikibetu(0))

    '基本情報＆カレンダレコードの有無判断
    gSL_Select = "Select * From TEXT設定テーブル "
    If SyoriKB <> "CPY" Then
       gSL_Select = gSL_Select & " Where レイアウト番号 = " & CisFun.Val2(iH1_No)
    Else
       gSL_Select = gSL_Select & " Where レイアウト番号 = " & CisFun.Val2(iH2_CpyNo)
    End If
    gSL_Select = gSL_Select & " And   区分 = 1"
    If TXSRead(gSL_Select, 1) Then
       If TXS.開始レコード <> 1 Then
          B1Chk_Cln.Value = 1
          mClnFlg = 1
       End If
    
       iB1_KSStart = TXS.開始位置
       iB1_KSikibetu = TXS.項目
    Else
       B1Chk_Cln.Value = 0
       mClnFlg = 0
    
       iB1_KSStart = ""
       iB1_KSikibetu = ""
    End If
    
End Sub
'+----------------------------------------------------------+
'+   項目設定用テーブル内容表示
'+----------------------------------------------------------+
Private Sub DBGridDisp()
    Dim wID As Long
            
    With CisDB
        .SQL = "TEXT変換システムオブジェクト"
        .StoadoCount = 2
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用
        .ParaNo = 2: .ParaIO = OutPut用
    '
        .ParaNo = 1: .ParaValue = "内示TEXTテーブル"
    '
        If Not .DBStored Then
            MsgBox "ストアドエラー "
            Exit Sub
        End If
    End With
    
    gSL_Select = "Select * FROM TEXT変換システムワーク"
    gSL_Select = gSL_Select & " WHERE cname <> 'FILECODE' and cname <> 'HEAD作成日' and cname <> 'HEADSEQ' "
    gSL_Select = gSL_Select & " and cname <> '連番' and cname <> '対象年月' and cname <> '号補区分' "
    gSL_Select = gSL_Select & " and cname <> '置換工区' and cname <> '作成日' and cname <> '作成端末' "
    gSL_Select = gSL_Select & " and cname <> '契約先' and cname <> '変更回数' "
    
    Call TXHRead(gSL_Select)
    
    VsGrid4.Rows = CisDB.RecordCount + 1
    
    gLong = 1
    With VsGrid4
        Do Until Not TXH_RDSTS
            .TextMatrix(gLong, 0) = Format(gLong, "##")
            .TextMatrix(gLong, 1) = RTrim(TXH.cname)
            .TextMatrix(gLong, 5) = RTrim(TXH.tname)
            .TextMatrix(gLong, 6) = TXH.prec
           
           gLong = gLong + 1
           
           Call TXHReadNext
        Loop
        .Row = 1: .Col = 1
    End With
    Call TXHClose
    
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
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    With CisFun
        .MB_Lines = 5
        .MB_MSG(1) = "     【 削除処理確認 】            "
        .MB_MSG(3) = "    № ＝ " & StrConv(iH1_No, vbWide)
        .MB_MSG(4) = "    名称 ＝ " & StrConv(iB1_FName, vbWide)
        .MB_Title = "削除処理"
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    Call CisDB.DBTran(TransBegin)              'ﾄﾗﾝｻﾞｸｼｮﾝ開始

' ( 削除処理 )
    With CisDB
        .SQL = "DELETE FROM TEXT設定テーブル"
        .SQL = .SQL & " WHERE レイアウト番号 = " & CisFun.Val2(iH1_No)
        .DBExec
    End With
    
    With CisDB
        .SQL = "DELETE FROM 名称マスタ"
        .SQL = .SQL & " WHERE 区分名称 = 'レイアウト情報内示'"
        .SQL = .SQL & " AND   値 = " & CisFun.Val2(iH1_No)
        .DBExec
    End With
    
    Call CisDB.DBTran(TransCommit)              'ﾄﾗﾝｻﾞｸｼｮﾝ終了
    
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
    Dim UpdFlg  As Boolean
    Dim wIndex  As Byte
    Dim wCnt1   As Byte
    Dim wCnt2   As Byte
    Dim wCnt3   As Byte
    Dim SYM_Flg As Byte
    
    DBPut = False
    UpdFlg = False
    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 5
        If SyoriKB = "ADD" Or SyoriKB = "CPY" Then
           .MB_MSG(1) = "     【 追加処理確認 】            "
           .MB_Title = "追加処理"
        Else
           .MB_MSG(1) = "     【 修正処理確認 】            "
           .MB_Title = "修正処理"
        End If
        .MB_MSG(3) = "    № ＝ " & StrConv(iH1_No, vbWide)
        .MB_MSG(4) = "    名称 ＝ " & StrConv(iB1_FName, vbWide)
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    On Error GoTo DBPut_Err
    
    Call CisDB.DBTran(TransBegin)              'ﾄﾗﾝｻﾞｸｼｮﾝ開始
    
    gSL_Select = "SELECT * FROM 名称マスタ "
    gSL_Select = gSL_Select & " WHERE 区分名称 = 'レイアウト情報内示'"
    gSL_Select = gSL_Select & " AND 値 = " & CisFun.Val2(iH1_No)
    gSL_Select = gSL_Select & " AND 区分タイプ = '' "
    If SYMRead(gSL_Select, 1) Then
       SYM_Flg = 1
    Else
       SYM_Flg = 0
    End If
    Call ItemsClearSYM
    Call SYMClose
    With SYM
       .区分名称 = "レイアウト情報内示"
       .区分タイプ = ""
       .値 = CisFun.Val2(iH1_No)
       .値名称 = RTrim(iB1_FName)
       .区分桁数 = 1
       .数字1 = 1
       .数字3 = CisFun.Val2(iB1_Len)
       .制御区分 = "X"
    End With
    
    If SYM_Flg = 0 Then
       Call SYMInsert
    Else
       Call SYMUpdate
    End If
    
'+---------------------------------------+
'+ TEXT設定テーブル更新      +
'+---------------------------------------+
    wCnt1 = 0: wCnt2 = 0: wCnt3 = 0
    
    With CisDB
        .SQL = "DELETE TEXT設定テーブル "
        .SQL = .SQL & " Where レイアウト番号 = " & CisFun.Val2(iH1_No)
        .SQL = .SQL & " And   区分 = 11"
        .DBExec
    End With
            
    With VsGrid4
        For gLong = 1 To .Rows - 1
'            If Trim(.TextMatrix(gLong, 1)) <> "" And (CisFun.Val2(.TextMatrix(gLong, 2)) <> 0 Or Trim(.TextMatrix(gLong, 5)) <> "") Then
            If Trim(.TextMatrix(gLong, 1)) <> "" And CisFun.Val2(.TextMatrix(gLong, 2)) <> 0 Then
               '/_/_/_/ データレコード項目情報出力【区分11】 /_/_/_/
               Call ItemsClearTXS
               '各レイアウト毎のレコード番号計算
               Select Case CisFun.Val2(.TextMatrix(gLong, 2))
                    Case 1: wCnt1 = wCnt1 + 1
                    Case 2: wCnt2 = wCnt2 + 1
                    Case 3: wCnt3 = wCnt3 + 1
               End Select
               TXS.レイアウト番号 = CisFun.Val2(iH1_No)
               TXS.区分 = 11
               TXS.連番 = CisFun.Val2(.TextMatrix(gLong, 2))
               Select Case CisFun.Val2(.TextMatrix(gLong, 2))
                    Case 1: TXS.レコード番号 = wCnt1
                    Case 2: TXS.レコード番号 = wCnt2
                    Case 3: TXS.レコード番号 = wCnt3
               End Select
               TXS.開始位置 = CisFun.Val2(.TextMatrix(gLong, 3))
               TXS.長さ = CisFun.Val2(.TextMatrix(gLong, 4))
               TXS.データ項目 = RTrim(.TextMatrix(gLong, 1))
               Select Case RTrim(.TextMatrix(gLong, 5))
                      Case "char", "varchar"
                           TXS.データタイプ = 0
                      Case Else
                           TXS.データタイプ = 1
               End Select
'               If RTrim(.TextMatrix(gLong, 5)) <> "" Or (RTrim(.TextMatrix(gLong, 1)) = "納入先" And CisFun.Val2(.TextMatrix(gLong, 3)) = 0) Then
               If RTrim(.TextMatrix(gLong, 1)) = "納入先" And CisFun.Val2(.TextMatrix(gLong, 3)) = 0 Then
                  TXS.データタイプ = 2
'                  TXS.固定文字 = RTrim(.TextMatrix(gLong, 5))
               End If
               If RTrim(.TextMatrix(gLong, 5)) = "char" Or RTrim(.TextMatrix(gLong, 5)) = "varchar" Then
                  TXS.出力長さ = CisFun.Val2(.TextMatrix(gLong, 6))
               End If
               
               Call TXSInsert
            End If
        Next gLong
        '/_/_/_/ データレコード基本情報出力【区分10】 /_/_/_/
        With CisDB
            .SQL = "DELETE TEXT設定テーブル "
            .SQL = .SQL & " Where レイアウト番号 = " & CisFun.Val2(iH1_No)
            .SQL = .SQL & " And   区分 = 10"
            .DBExec
        End With
        If wCnt1 <> 0 Then
           Call ItemsClearTXS
           TXS.レイアウト番号 = CisFun.Val2(iH1_No)
           TXS.区分 = 10
           TXS.連番 = 1
           TXS.レコード定義件数 = wCnt1
           TXS.項目 = RTrim(mSikibetu(0))
           TXS.開始位置 = CisFun.Val2(mSKaisi(0))
           TXS.長さ = Len(RTrim(mSikibetu(0)))
           
           Call TXSInsert
         
        End If
        If wCnt2 <> 0 Then
           Call ItemsClearTXS
           TXS.レイアウト番号 = CisFun.Val2(iH1_No)
           TXS.区分 = 10
           TXS.連番 = 2
           TXS.レコード定義件数 = wCnt2
           TXS.項目 = RTrim(mSikibetu(1))
           TXS.開始位置 = CisFun.Val2(mSKaisi(1))
           TXS.長さ = Len(RTrim(mSikibetu(1)))
           
           Call TXSInsert
         
        End If
        If wCnt3 <> 0 Then
           Call ItemsClearTXS
           TXS.レイアウト番号 = CisFun.Val2(iH1_No)
           TXS.区分 = 10
           TXS.連番 = 3
           TXS.レコード定義件数 = wCnt3
           TXS.項目 = RTrim(mSikibetu(2))
           TXS.開始位置 = CisFun.Val2(mSKaisi(2))
           TXS.長さ = Len(RTrim(mSikibetu(2)))
           
           Call TXSInsert
         
        End If
        '/_/_/_/ 基本情報出力【区分1】 /_/_/_/
        With CisDB
            .SQL = "DELETE TEXT設定テーブル "
            .SQL = .SQL & " Where レイアウト番号 = " & CisFun.Val2(iH1_No)
            .SQL = .SQL & " And   区分 = 1"
            .DBExec
        End With
        Call ItemsClearTXS
        TXS.レイアウト番号 = CisFun.Val2(iH1_No)
        TXS.区分 = 1
        If wCnt1 <> 0 Then TXS.レコード定義件数 = 1
        If wCnt2 <> 0 Then TXS.レコード定義件数 = 2
        If wCnt3 <> 0 Then TXS.レコード定義件数 = 3
        If mClnFlg = 0 Then
           TXS.開始レコード = 1
        Else
           TXS.開始レコード = 2
        End If
        
        If wCnt3 <> 0 Then TXS.最終レコード = 3
        If wCnt2 <> 0 And wCnt3 = 0 Then TXS.最終レコード = 2
        If wCnt1 <> 0 And wCnt2 = 0 And wCnt3 = 0 Then TXS.最終レコード = 1
        TXS.項目 = RTrim(iB1_KSikibetu)
        TXS.開始位置 = CisFun.Val2(iB1_KSStart)
        TXS.長さ = Len(iB1_KSikibetu)
        
        Call TXSInsert
        
    End With
    
    Call CisDB.DBTran(TransCommit)              'ﾄﾗﾝｻﾞｸｼｮﾝ終了
    
    DBPut = True
    Exit Function
DBPut_Err:
    CisFun.ErrorBox
    End
End Function
'+==========================================================+
'+      グリッド処理(レイアウト設定）                         +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit3()
    With CisVsGrid2
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear ("NO")
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容)
         Call .FixedSet(中中, 右中, なし, 0, 0, "ｶﾗﾑ")
         Call .FixedSet(中中, 左中, あり, 0, 1, "項目名")
         Call .FixedSet(中中, 右中, あり, 0, 2, "桁数")
         Call .FixedSet(中中, 左中, あり, 0, 3, "NO")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'         Call .FixedSet(中中, 左中, なし, 0, 0, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2")
        .InitDisp
    End With
End Sub
'+----------------------------------------------------------+
'+   DB項目設定よりレイアウト設定へ項目移送
'+----------------------------------------------------------+
Private Sub MoveItemLayout()
    Dim wFlg    As Boolean
    Dim iD5     As Long
    Dim id2     As Long
'
    Dim wStart  As Long
    Dim wLayNo  As Byte
    
    wFlg = False
    
    With CisDB
        .SQL = "DELETE TEXTレイアウト設定ワーク" & RTrim(mWSID)
        .DBExec
    End With
    
    With VsGrid4
        gInt = 1
        For gLong = 1 To .Rows - 1
            If Trim(.TextMatrix(gLong, 1)) <> "" And CisFun.Val2(.TextMatrix(gLong, 2)) <> 0 And CisFun.Val2(.TextMatrix(gLong, 3)) <> 0 And CisFun.Val2(.TextMatrix(gLong, 4)) <> 0 Then
               CisDB.SQL = "INSERT INTO TEXTレイアウト設定ワーク" & RTrim(mWSID)
               CisDB.SQL = CisDB.SQL & " ( NO,項目名,連番,桁数,開始 ) VALUES ( "
               CisDB.SQL = CisDB.SQL & gInt & ",'" & RTrim(.TextMatrix(gLong, 1)) & "'," & CisFun.Val2(.TextMatrix(gLong, 2)) & "," & CisFun.Val2(.TextMatrix(gLong, 4)) & "," & CisFun.Val2(.TextMatrix(gLong, 3)) & " )"
               CisDB.DBExec
               wFlg = True
               gInt = gInt + 1
            End If
        Next gLong
    End With
    
    If wFlg Then
       With CisDB
           .SQL = "SELECT * FROM TEXTレイアウト設定ワーク" & RTrim(mWSID)
           .SQL = .SQL & " ORDER BY 連番,開始"
           .DBRead
       End With
    Else
       Exit Sub
    End If
    
    ReDim mItem1(VsGrid4.Rows + 1)
    ReDim mKeta1(VsGrid4.Rows + 1)
    ReDim mStart1(VsGrid4.Rows + 1)
    ReDim mItem2(VsGrid4.Rows + 1)
    ReDim mKeta2(VsGrid4.Rows + 1)
    ReDim mStart2(VsGrid4.Rows + 1)
    ReDim mItem3(VsGrid4.Rows + 1)
    ReDim mKeta3(VsGrid4.Rows + 1)
    ReDim mStart3(VsGrid4.Rows + 1)
    wStart = 1: mCnt1 = 0: mCnt2 = 0: mCnt3 = 0
    
    With VsGrid3
        wLayNo = CisDB.Recordset(2)
        iD5 = 0
        'TEXTレイアウト情報を退避
        Do Until CisDB.Recordset.EOF
           
           If wLayNo <> CisDB.Recordset(2) Then
              iD5 = 0
           End If
           
           If CisDB.Recordset(2) = 1 Then
              mItem1(iD5) = CisDB.Recordset(1)
              mKeta1(iD5) = CisDB.Recordset(3)
              mStart1(iD5) = CisDB.Recordset(4)
              mCnt1 = mCnt1 + 1
           End If
           If CisDB.Recordset(2) = 2 Then
              mItem2(iD5) = CisDB.Recordset(1)
              mKeta2(iD5) = CisDB.Recordset(3)
              mStart2(iD5) = CisDB.Recordset(4)
              mCnt2 = mCnt2 + 1
           End If
           If CisDB.Recordset(2) = 3 Then
              mItem3(iD5) = CisDB.Recordset(1)
              mKeta3(iD5) = CisDB.Recordset(3)
              mStart3(iD5) = CisDB.Recordset(4)
              mCnt3 = mCnt3 + 1
           End If
           
           wLayNo = CisDB.Recordset(2)
           iD5 = iD5 + 1
           
           CisDB.Recordset.MoveNext
                       
        Loop
    
    End With

End Sub
'+----------------------------------------------------------+
'+   テキストレイアウト表示①
'+----------------------------------------------------------+
Private Sub TextGrid_Dsp1()
    Dim id1     As Long
    Dim id2     As Long
    Dim wRow    As Long
    Dim wStart  As Long

    'TEXTレイアウト情報を退避
        
    Call CisVsGrid2.Clear
    
    If CisFun.Val2(iB1_Len) + 1 > VsGrid3.Rows Then
        VsGrid3.Rows = CisFun.Val2(iB1_Len) + 1
    End If
    For id2 = 1 To VsGrid3.Rows - 1
        VsGrid3.TextMatrix(id2, 0) = id2
    Next id2
    
    With VsGrid3
    
        wStart = 1
        For wRow = 0 To mCnt1 - 1
           
           '未設定領域が存在する場合
           If wStart < mStart1(wRow) Then
              For id1 = wStart To mStart1(wRow)                             '未設定領域の表示
                  .TextMatrix(id1, 1) = "*未設定*"
                  .TextMatrix(id1, 2) = mStart1(wRow) - wStart
              Next id1
              For id1 = mStart1(wRow) To (mStart1(wRow) + mKeta1(wRow)) - 1     '設定領域の表示
                  .TextMatrix(id1, 1) = RTrim(mItem1(wRow))
'                  .TextMatrix(id1, 2) = mStart1(wRow) - wStart
                  .TextMatrix(id1, 2) = mKeta1(wRow)
              Next id1
           Else
           '設定領域のみ表示
              For id1 = mStart1(wRow) To (mStart1(wRow) + mKeta1(wRow)) - 1     '設定領域の表示
                  .TextMatrix(id1, 1) = RTrim(mItem1(wRow))
'                  .TextMatrix(id1, 2) = mStart1(wRow) - wStart
                  .TextMatrix(id1, 2) = mKeta1(wRow)
              Next id1
           End If
           wStart = mStart1(wRow) + mKeta1(wRow)
           
        Next wRow
        
        '最後、レコード長まで表示
        If wStart < CisFun.Val2(iB1_Len) Then
            For id1 = wStart To CisFun.Val2(iB1_Len)                              '未設定領域の表示
                .TextMatrix(id1, 1) = "*未設定*"
                .TextMatrix(id1, 2) = (CisFun.Val2(iB1_Len) - wStart) + 1
            Next id1
        End If

    End With
    
End Sub
'+----------------------------------------------------------+
'+   テキストレイアウト表示②
'+----------------------------------------------------------+
Private Sub TextGrid_Dsp2()
    Dim id1     As Long
    Dim id2     As Long
    Dim wRow    As Long
    Dim wStart  As Long

    'TEXTレイアウト情報を退避
        
    Call CisVsGrid2.Clear
    
    If CisFun.Val2(iB1_Len) + 1 > VsGrid3.Rows Then
        VsGrid3.Rows = CisFun.Val2(iB1_Len) + 1
    End If
    For id2 = 1 To VsGrid3.Rows - 1
        VsGrid3.TextMatrix(id2, 0) = id2
    Next id2
    
    With VsGrid3
    
        wStart = 1
        For wRow = 0 To mCnt2 - 1
           
           '未設定領域が存在する場合
           If wStart < mStart2(wRow) Then
              For id1 = wStart To mStart2(wRow)                             '未設定領域の表示
                  .TextMatrix(id1, 1) = "*未設定*"
                  .TextMatrix(id1, 2) = mStart2(wRow) - wStart
              Next id1
              For id1 = mStart2(wRow) To (mStart2(wRow) + mKeta2(wRow)) - 1     '設定領域の表示
                  .TextMatrix(id1, 1) = RTrim(mItem2(wRow))
'                  .TextMatrix(id1, 2) = mStart2(wRow) - wStart
                  .TextMatrix(id1, 2) = mKeta2(wRow)
              Next id1
           Else
           '設定領域のみ表示
              For id1 = mStart2(wRow) To (mStart2(wRow) + mKeta2(wRow)) - 1     '設定領域の表示
                  .TextMatrix(id1, 1) = RTrim(mItem2(wRow))
'                  .TextMatrix(id1, 2) = mStart2(wRow) - wStart
                  .TextMatrix(id1, 2) = mKeta2(wRow)
              Next id1
           End If
           wStart = mStart2(wRow) + mKeta2(wRow)
           
        Next wRow
        
        '最後、レコード長まで表示
        If wStart < CisFun.Val2(iB1_Len) Then
            For id1 = wStart To CisFun.Val2(iB1_Len)                              '未設定領域の表示
                .TextMatrix(id1, 1) = "*未設定*"
                .TextMatrix(id1, 2) = (CisFun.Val2(iB1_Len) - wStart) + 1
            Next id1
        End If

    End With
    
End Sub
'+----------------------------------------------------------+
'+   テキストレイアウト表示③
'+----------------------------------------------------------+
Private Sub TextGrid_Dsp3()
    Dim id1     As Long
    Dim id2     As Long
    Dim wRow    As Long
    Dim wStart  As Long

    'TEXTレイアウト情報を退避
        
    Call CisVsGrid2.Clear
    
    If CisFun.Val2(iB1_Len) + 1 > VsGrid3.Rows Then
        VsGrid3.Rows = CisFun.Val2(iB1_Len) + 1
    End If
    For id2 = 1 To VsGrid3.Rows - 1
        VsGrid3.TextMatrix(id2, 0) = id2
    Next id2
    
    With VsGrid3
    
        wStart = 1
        For wRow = 0 To mCnt3 - 1
           
           '未設定領域が存在する場合
           If wStart < mStart3(wRow) Then
              For id1 = wStart To mStart3(wRow)                             '未設定領域の表示
                  .TextMatrix(id1, 1) = "*未設定*"
                  .TextMatrix(id1, 2) = mStart3(wRow) - wStart
              Next id1
              For id1 = mStart3(wRow) To (mStart3(wRow) + mKeta3(wRow)) - 1     '設定領域の表示
                  .TextMatrix(id1, 1) = RTrim(mItem3(wRow))
'                  .TextMatrix(id1, 2) = mStart3(wRow) - wStart
                  .TextMatrix(id1, 2) = mKeta3(wRow)
              Next id1
           Else
           '設定領域のみ表示
              For id1 = mStart3(wRow) To (mStart3(wRow) + mKeta3(wRow)) - 1     '設定領域の表示
                  .TextMatrix(id1, 1) = RTrim(mItem3(wRow))
'                  .TextMatrix(id1, 2) = mStart3(wRow) - wStart
                  .TextMatrix(id1, 2) = mKeta3(wRow)
              Next id1
           End If
           wStart = mStart3(wRow) + mKeta3(wRow)
           
        Next wRow
        
        '最後、レコード長まで表示
        If wStart < CisFun.Val2(iB1_Len) Then
            For id1 = wStart To CisFun.Val2(iB1_Len)                              '未設定領域の表示
                .TextMatrix(id1, 1) = "*未設定*"
                .TextMatrix(id1, 2) = (CisFun.Val2(iB1_Len) - wStart) + 1
            Next id1
        End If

    End With
    
End Sub

'+==========================================================+
'+      グリッド処理(変換項目設定)                            +
'+==========================================================+
'+----------------------------------------------------------+
'+   見出しｾｯﾄ & 初期状態ｾｯﾄ
'+----------------------------------------------------------+
Private Sub GridInit4()
    With CisVsGrid3
'　ｸﾞﾘｯﾄﾞ初期化処理
'        .Clear ("NO")
'       Call .FixedSet(見出し位置設定,列位置設定,ﾏｰｼﾞ処理,
'                      ｾﾙ桁数,列番号,見出し内容)
         Call .FixedSet(中中, 右中, あり, 0, 0, "№")
         Call .FixedSet(中中, 左中, あり, 0, 1, "変換先項目")
         Call .FixedSet(中中, 右中, あり, 0, 2, "連番")
         Call .FixedSet(中中, 右中, あり, 0, 3, "開始")
         Call .FixedSet(中中, 右中, あり, 0, 4, "桁数")
         Call .FixedSet(中中, 左中, あり, 0, 5, "変換先属性")
         Call .FixedSet(中中, 左中, あり, 0, 6, "変換先桁数")
    
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
'         Call .FixedSet(中中, 左中, なし, 0, 0, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
         Call .PatanSet(0, "0,1,2,3,4")
         .InitDisp
    End With
End Sub


Private Sub VSGrid4_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyUp And KeyCode <> vbKeyDown Then Exit Sub
    Call TextToVsGrid(Val(iB1_LayNo.Tag))
    Call VsGridToText(VsGrid4.Row)
End Sub
Private Sub VSGrid4_Click()
    Call TextToVsGrid(Val(iB1_LayNo.Tag))
    Call VsGridToText(VsGrid4.Row)
    On Error Resume Next
    If ActObj.Visible Then ActObj.SetFocus
    On Error GoTo 0
End Sub
'Private Sub VsGrid4_Click()
'    Dim wKey        As String
'    Dim wRow        As Integer
'
'    wRow = VsGrid4.Row
'
''    If RTrim(VsGrid4.TextMatrix(wRow, 1)) = "" Then
''       B1Comb_Text.Text = "": iB1_Kotei = "": iB1_YukoKeta = ""
''
''       B1Comb_Text.Enabled = False
''       B1Op_TxtKotei(0).Enabled = False
''       B1Op_TxtKotei(1).Enabled = False
''       iB1_Kotei.Enabled = False
''    End If
'
'    VsGrid4.Row = Now_VS4Row
'    VsGrid4.Col = 1: VsGrid4.ColSel = 5
'    VsGrid4.CellBackColor = &HC0FFFF
'    VsGrid4.CellForeColor = &H80000008
'
'    VsGrid4.Row = wRow
'    VsGrid4.Col = 1: VsGrid4.ColSel = 5
'    VsGrid4.CellBackColor = &H800000
'    VsGrid4.CellForeColor = &H8000000E
''
'    '変換元情報をVS4Gridに入力
'    '--項目区分
'    If B1Op_TxtKotei(0).Value Then
'       VsGrid4.TextMatrix(Now_VS4Row, 7) = "0"
'    Else
'       VsGrid4.TextMatrix(Now_VS4Row, 7) = "1"
'    End If
'
'    If Not B1Op_TxtKotei(0).Value Then              'テキスト項目
'       '--固定文字
'       VsGrid4.TextMatrix(Now_VS4Row, 5) = RTrim(iB1_Kotei)
'       '--変換元テキスト
'       VsGrid4.TextMatrix(Now_VS4Row, 4) = ""
'       '--ﾚｲｱｳﾄ連番,連番,繰返連番
'       VsGrid4.TextMatrix(Now_VS4Row, 8) = ""                              'ﾚｲｱｳﾄ連番
'       VsGrid4.TextMatrix(Now_VS4Row, 9) = ""                              '連番
'       VsGrid4.TextMatrix(Now_VS4Row, 12) = ""                             '繰返連番
'       '--年変換
'       VsGrid4.TextMatrix(Now_VS4Row, 11) = "0"
'       '--有効桁数
'       VsGrid4.TextMatrix(Now_VS4Row, 10) = ""
'       '--左右区分
'       VsGrid4.TextMatrix(Now_VS4Row, 13) = "0"
'    End If
'
'    If Trim(VsGrid4.TextMatrix(VsGrid4.Row, 1)) <> "" Then
'       B1lb_DBKoumoku = Trim(VsGrid4.TextMatrix(VsGrid4.Row, 1))
'       If Val2(VsGrid4.TextMatrix(VsGrid4.Row, 7)) = 0 Then
'         B1Op_TxtKotei(0).Value = True
'         iB1_Kotei = ""
'         iB1_Kotei.Enabled = False
'
'         B1Comb_Text.Enabled = True
'         iB1_YukoKeta.Enabled = True
'         Back_YearConv.Enabled = True
'         Back_RLKB.Enabled = True
'
'         If Trim(VsGrid4.TextMatrix(VsGrid4.Row, 9)) <> "" Then
'            wKey = Format(Val2(VsGrid4.TextMatrix(VsGrid4.Row, 8)), "0") & VsGrid4.TextMatrix(VsGrid4.Row, 9) & VsGrid4.TextMatrix(VsGrid4.Row, 12)
'            With mCmbText
'                .Code = RTrim(wKey)
'                Call .NameGet
'            End With
'            iB1_YukoKeta = VsGrid4.TextMatrix(VsGrid4.Row, 10)
'
'            Call B1Op_YearConv_Click(Val2(VsGrid4.TextMatrix(VsGrid4.Row, 11)))
'            B1Op_YearConv(Val2(VsGrid4.TextMatrix(VsGrid4.Row, 11))).Value = True
'
'            Call B1Op_RLKB_Click(Val2(VsGrid4.TextMatrix(VsGrid4.Row, 13)))
'            B1Op_RLKB(Val2(VsGrid4.TextMatrix(VsGrid4.Row, 13))).Value = True
'         Else
'            With mCmbText
'                .Code = ""
'                Call .NameGet
'            End With
'            iB1_YukoKeta = ""
'            Call B1Op_YearConv_Click(0)
'            B1Op_YearConv(0).Value = True
'            Call B1Op_RLKB_Click(0)
'            B1Op_RLKB(0).Value = True
'         End If
'       Else
'         B1Op_TxtKotei(1).Value = True
'
'         With mCmbText
'             .Code = ""
'             Call .NameGet
'         End With
'         iB1_YukoKeta = ""
'         Call B1Op_YearConv_Click(0)
'         B1Op_YearConv(0).Value = True
'         Call B1Op_RLKB_Click(0)
'         B1Op_RLKB(0).Value = True
'
'         B1Comb_Text.Enabled = False
'         iB1_YukoKeta.Enabled = False
'         Back_YearConv.Enabled = False
'         Back_RLKB.Enabled = False
'
'         iB1_Kotei.Enabled = True
'         iB1_Kotei = VsGrid4.TextMatrix(VsGrid4.Row, 5)
'
'       End If
'
'    End If
'
'    Now_VS4Row = VsGrid4.Row
'
'End Sub
'Private Sub Grid4_Move()
'    Dim wKey        As String
'
''    VsGrid4.Row = 1
''    VsGrid4.Col = 1: VsGrid4.ColSel = 5
''    VsGrid4.CellBackColor = &H800000
''    VsGrid4.CellForeColor = &H8000000E
'
'    VsGrid4.Row = Now_VS4Row
'    VsGrid4.Col = 1: VsGrid4.ColSel = 5
'    VsGrid4.CellBackColor = &H800000
'    VsGrid4.CellForeColor = &H8000000E
'
'    If Trim(VsGrid4.TextMatrix(VsGrid4.Row, 1)) <> "" Then
'       B1lb_DBKoumoku = Trim(VsGrid4.TextMatrix(VsGrid4.Row, 1))
'       If Val2(VsGrid4.TextMatrix(VsGrid4.Row, 7)) = 0 Then
'         B1Op_TxtKotei(0).Value = True
'         iB1_Kotei = ""
'         iB1_Kotei.Enabled = False
'
'         B1Comb_Text.Enabled = True
'         iB1_YukoKeta.Enabled = True
'         Back_YearConv.Enabled = True
'         Back_RLKB.Enabled = True
'
'         If Trim(VsGrid4.TextMatrix(VsGrid4.Row, 9)) <> "" Then
'            wKey = Format(Val2(VsGrid4.TextMatrix(VsGrid4.Row, 8)), "0") & VsGrid4.TextMatrix(VsGrid4.Row, 9) & VsGrid4.TextMatrix(VsGrid4.Row, 12)
'            With mCmbText
'                .Code = RTrim(wKey)
'                Call .NameGet
'            End With
'            iB1_YukoKeta = VsGrid4.TextMatrix(VsGrid4.Row, 10)
'            Call B1Op_YearConv_Click(Val2(VsGrid4.TextMatrix(VsGrid4.Row, 11)))
'            B1Op_YearConv(Val2(VsGrid4.TextMatrix(VsGrid4.Row, 11))).Value = True
'
'            Call B1Op_RLKB_Click(Val2(VsGrid4.TextMatrix(VsGrid4.Row, 13)))
'            B1Op_RLKB(Val2(VsGrid4.TextMatrix(VsGrid4.Row, 13))).Value = True
'         Else
'            With mCmbText
'                .Code = ""
'                Call .NameGet
'            End With
'            iB1_YukoKeta = ""
'            Call B1Op_YearConv_Click(0)
'            B1Op_YearConv(0).Value = True
'
'            Call B1Op_RLKB_Click(0)
'            B1Op_RLKB(0).Value = True
'         End If
'       Else
'         B1Op_TxtKotei(1).Value = True
'
'         With mCmbText
'             .Code = ""
'             Call .NameGet
'         End With
'         iB1_YukoKeta = ""
'         Call B1Op_YearConv_Click(0)
'         B1Op_YearConv(0).Value = True
'         Call B1Op_RLKB_Click(0)
'         B1Op_RLKB(0).Value = True
'
'         B1Comb_Text.Enabled = False
'         iB1_YukoKeta.Enabled = False
'         Back_YearConv.Enabled = False
'         Back_RLKB.Enabled = False
'
'         iB1_Kotei.Enabled = True
'         iB1_Kotei = VsGrid4.TextMatrix(VsGrid4.Row, 5)
'
'       End If
'
'    End If
'
'End Sub
'Private Sub Comb_text_Move()
'
'    '変換元情報をVS4Gridに入力
'    '--項目区分
'    If B1Op_TxtKotei(0).Value Then
'       VsGrid4.TextMatrix(Now_VS4Row, 7) = "0"
'    Else
'       VsGrid4.TextMatrix(Now_VS4Row, 7) = "1"
'    End If
'
'    If B1Op_TxtKotei(0).Value Then              'テキスト項目
'       '--変換元テキスト
'       VsGrid4.TextMatrix(Now_VS4Row, 4) = RTrim(B1Comb_Text.Text)
'       '--ﾚｲｱｳﾄ連番,連番,繰返連番
'       Call mCmbText.CodeGet
'       VsGrid4.TextMatrix(Now_VS4Row, 8) = Mid(mCmbText.Code, 1, 1)        'ﾚｲｱｳﾄ連番
'       VsGrid4.TextMatrix(Now_VS4Row, 9) = Mid(mCmbText.Code, 2, 3)        '連番
'       VsGrid4.TextMatrix(Now_VS4Row, 12) = Mid(mCmbText.Code, 5, 2)       '繰返連番
'       '--年変換
'       If B1Op_YearConv(0).Value Then
'          VsGrid4.TextMatrix(Now_VS4Row, 11) = "0"
'       Else
'          VsGrid4.TextMatrix(Now_VS4Row, 11) = "1"
'       End If
'       '--左右区分
'       If B1Op_RLKB(0).Value Then
'          VsGrid4.TextMatrix(Now_VS4Row, 13) = "0"
'       Else
'          VsGrid4.TextMatrix(Now_VS4Row, 13) = "1"
'       End If
'       '--有効桁数
'       VsGrid4.TextMatrix(Now_VS4Row, 10) = iB1_YukoKeta
'       '--固定文字
'       VsGrid4.TextMatrix(Now_VS4Row, 5) = ""
'    Else                                        '固定文字
'       '--固定文字
'       VsGrid4.TextMatrix(Now_VS4Row, 5) = RTrim(iB1_Kotei)
'       '--変換元テキスト
'       VsGrid4.TextMatrix(Now_VS4Row, 4) = ""
'       '--ﾚｲｱｳﾄ連番,連番,繰返連番
'       VsGrid4.TextMatrix(Now_VS4Row, 8) = ""                              'ﾚｲｱｳﾄ連番
'       VsGrid4.TextMatrix(Now_VS4Row, 9) = ""                              '連番
'       VsGrid4.TextMatrix(Now_VS4Row, 12) = ""                             '繰返連番
'       '--年変換
'       VsGrid4.TextMatrix(Now_VS4Row, 11) = "0"
'       '--左右区分
'       VsGrid4.TextMatrix(Now_VS4Row, 13) = "0"
'       '--有効桁数
'       VsGrid4.TextMatrix(Now_VS4Row, 10) = ""
'    End If
'
'
'End Sub
'*********************************************************
'* レイアウト表示ワーク作成
'*********************************************************
Private Sub Work_Crt()
    
    With CisDB
        .SQL = "CREATE TABLE [dbo].[TEXTレイアウト設定ワーク" & RTrim(mWSID) & "] ("
        .SQL = .SQL & " [NO] [int] NULL ,"
        .SQL = .SQL & " [項目名] [varchar](30) NULL ,"
        .SQL = .SQL & " [連番] [tinyint] NULL ,"
        .SQL = .SQL & " [桁数] [int] NULL ,"
        .SQL = .SQL & " [開始] [int] NULL ,"
        .SQL = .SQL & " ) "
        .DBExec
    End With
    
End Sub
'*********************************************************
'* レイアウト表示ワーク削除
'*********************************************************
Private Sub Work_DEL()
    
    With CisDB
        .SQL = "DROP TABLE [dbo].[TEXTレイアウト設定ワーク" & RTrim(mWSID) & "] "
        .DBExec
    End With
    
End Sub
