VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Begin VB.Form CXH0040 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "内示発注　確定処理"
   ClientHeight    =   10050
   ClientLeft      =   -150
   ClientTop       =   2595
   ClientWidth     =   14760
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
   ScaleHeight     =   10050
   ScaleWidth      =   14760
   WindowState     =   2  '最大化
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00000000&
      Height          =   405
      Left            =   1185
      ScaleHeight     =   345
      ScaleWidth      =   4230
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   1230
      Width           =   4290
      Begin Cis3D_v60.CIS3D CIS3D3 
         Height          =   345
         Left            =   15
         Top             =   0
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   609
         Caption         =   " 選択件数"
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
         Begin Cis3D_v60.CIS3D B1lb_SelSu 
            Height          =   270
            Left            =   1095
            Top             =   30
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   476
            BackColor       =   12648384
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
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   345
         Left            =   2115
         Top             =   0
         Width           =   2100
         _ExtentX        =   3704
         _ExtentY        =   609
         Caption         =   " 表示件数"
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
         Begin Cis3D_v60.CIS3D B1lb_DspSu 
            Height          =   270
            Left            =   1095
            Top             =   30
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   476
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
   Begin Cis3D_v60.CIS3D B1_Area1 
      Height          =   7065
      Left            =   1155
      Top             =   2985
      Width           =   12585
      _ExtentX        =   22199
      _ExtentY        =   12462
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
         Height          =   6915
         Left            =   75
         Top             =   60
         Width           =   12420
         _ExtentX        =   21908
         _ExtentY        =   12197
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
            DragIcon        =   "CXH0040.frx":0000
            Height          =   6720
            Left            =   105
            TabIndex        =   4
            Top             =   105
            Width           =   12195
            _Version        =   196608
            _ExtentX        =   21511
            _ExtentY        =   11853
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
            FormatString    =   "№  |処理|仕入先 |仕入先名　 　　　　　|受入|品　番                   |納 入 日   |発注数 |手配区||Select|"
            Rows            =   21
            Cols            =   20
            BackColor       =   12648447
            BackColorBkg    =   12632256
            FocusRect       =   0
            HighLight       =   2
            RowHeightMin    =   320
            GridLines       =   2
            SelectionMode   =   1
            Begin Cis3D_v60.CIS3D Message 
               Height          =   930
               Left            =   2775
               Top             =   555
               Visible         =   0   'False
               Width           =   6855
               _ExtentX        =   12091
               _ExtentY        =   1640
               BackColor       =   12582912
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
               Begin Cis3D_v60.CIS3D lb_Message 
                  Height          =   660
                  Left            =   135
                  Top             =   150
                  Width           =   6615
                  _ExtentX        =   11668
                  _ExtentY        =   1164
                  BackColor       =   16761024
                  ForeColor       =   0
                  Caption         =   "【表示中】しばらくお待ち下さい"
                  BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                     Name            =   "HG丸ｺﾞｼｯｸM-PRO"
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
      Height          =   1125
      Left            =   1155
      ScaleHeight     =   1065
      ScaleWidth      =   9915
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1710
      Width           =   9975
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   960
         Left            =   6405
         Top             =   60
         Width           =   3435
         _ExtentX        =   6059
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "品   番"
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
            Left            =   135
            TabIndex        =   2
            Top             =   420
            Width           =   3210
            _ExtentX        =   5662
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
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   960
         Left            =   2175
         Top             =   60
         Width           =   3510
         _ExtentX        =   6191
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "仕 入 先"
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
            Left            =   120
            TabIndex        =   0
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
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1095
            Top             =   435
            Width           =   2295
            _ExtentX        =   4048
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
         Height          =   960
         Left            =   5685
         Top             =   60
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   1693
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
            Left            =   165
            TabIndex        =   1
            Top             =   420
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
      Begin Cis3D_v60.CIS3D CIS3D31 
         Height          =   960
         Left            =   45
         Top             =   60
         Width           =   2130
         _ExtentX        =   3757
         _ExtentY        =   1693
         ForeColor       =   16711680
         Caption         =   "内示年月"
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
         Begin CisYM_V60.CisYM iH1_NJYm 
            Height          =   405
            Left            =   195
            TabIndex        =   3
            Top             =   405
            Width           =   1785
            _ExtentX        =   2328
            _ExtentY        =   582
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "ＭＳ ゴシック"
               Size            =   14.25
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
      Left            =   9255
      ScaleHeight     =   660
      ScaleWidth      =   4290
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   825
      Width           =   4350
      Begin CisBtn_60.CisBtn PB_Dsp 
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
         BtnType         =   12
         BtnCaption      =   "照会"
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
      Begin CisBtn_60.CisBtn PB_CAN 
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
      TabIndex        =   5
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
      Caption         =   "【 内示発注　確定処理  】"
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
Attribute VB_Name = "CXH0040"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   内示発注　確定処理
'**       フォームID    :   CXH0040
'**       処理概要      :
'**
'**       作  成  日    :   2004/03/25  By CIS
'**       変  更  日    :   2004/04/08  By CIS - 画面中央位置表示の追加
'**       変  更  日    :   2005/11/07  By CIS オプション追加（発注打切りﾘｽﾄ発行）
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    
    Dim CisVsGrid               As New CisVsGrid3
    
    Dim mWhere                  As String
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
'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 3
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 11
        Set .SelectDispObj = B1lb_SelSu
        .InitGet = False
        .Init
    End With
    Call GridInit
    
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
    FormAct = Not FormAct
    
' 内示展開処理年月
    Call GetNJProc_YM(3)
    iH1_NJYm = Trim(gn_HNYM)
    
'''    '発注完了済リスト印刷
'        RV_SYmd = RTrim(iH1_NJYm)
'        If Not CXH0040LPrint Then Exit Sub


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
           Case vbKeyF4:        If PB_Dsp.Visible Then Call PB_Dsp_Click    '【照会】
           Case vbKeyF7:        If PB_Look.Visible Then Call PB_Look_Click  '【検索】
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
      iH1_Torcd.SetFocus
      GoTo PB_CAN_Ed
   End If
    
   If ProcHB = "B1" Then
      ProcHB = "H1"
      Call DispChange(ProcHB)
      Call HeadBodyClear("B")
      Call SyoriName(SyoriKB)
      iH1_Torcd.SetFocus
      GoTo PB_CAN_Ed
   End If
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
End Sub
'+----------------------------+
'+    発注処理照会(F4)       +
'+----------------------------+
Private Sub PB_Dsp_Click()
    
    RV_Call = "C"
    RV_SYmd = RTrim(iH1_NJYm)
    RV_Torcd = RTrim(iH1_Torcd)
    
    CXH0045.Show vbModal
    Unload CXH0045
    Set CXH0045 = Nothing

    If RV_Rtn Then
        iH1_Torcd = RV_Torcd
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
        iH1_Hinbn = RV_Hinbn
    End If

End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SSTorcd: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_UTorcd(0): Exit Sub
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
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
        H1lb_Tornm = RV_TorRName
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
    RVI_SK = 1
    RVI_Hinmoku = 1

    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        Call iH1_Torcd_LostFocus
        iH1_Ukeir = RV_Ukeir
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
    If TorNmGet(iH1_Torcd, 8) Then
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
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Static w_Color      As Long
    w_Color = &H80000008 '初期色ｾｯﾄ(黒)
    
    Select Case SyoriKB
        Case "DEL"
                            SyoriNM = "削除"
                            K_Sykbnm.ForeColor = &HFF&              '(赤色)
                            w_Color = &HFF&
        Case Else:          SyoriNM = ""
    End Select
    
'    For Each gObj In Me.Controls
'        If gObj.Name Like "i*" Then gObj.ForeColor = w_Color
'        If gObj.Name Like "*Comb*" Then gObj.ForeColor = w_Color
'
'''       * PF-Key 表示設定 *
''        If gObj.Name = "PB_ADD" Or gObj.Name = "PB_MNT" Or _
''           gObj.Name = "PB_DEL" Then
''                gObj.Visible = Not (UCase(Mid(gObj.Name, 4, 3)) = SyoriKB)
''        End If
'    Next gObj
    
    Timer2.Enabled = False
    K_Sykbnm = SyoriNM
    K_Sykbnm.Visible = Not (SyoriNM = "")
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
    If RTrim$(K_Sykbnm) <> "" Then K_Sykbnm = "" Else K_Sykbnm = SyoriNM
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
    
'    PB_CAN.Visible = Not H1Mode
    PB_Look.Visible = H1Mode
    PB_Dsp.Visible = Not B1Mode
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
        B1lb_SelSu = ""
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
       If Not Head1Chk Then GoTo ReturnPress_Ed
       ProcHB = "B1"
       Call DispChange(ProcHB)
       vsGrid1.SetFocus
       GoTo ReturnPress_Ed
    End If
    If ProcHB = "B1" Then
       If Not Body1Chk Then GoTo ReturnPress_Ed
    ' << 発注処理 >>
        If Not HCProc Then GoTo ReturnPress_Ed
        Call HeadBodyClear("B")
        ProcHB = "H1"
        Call DispChange(ProcHB)
        Call SyoriName(SyoriKB)
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
    
    If RTrim(iH1_NJYm) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    内示年月を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       Exit Function
    End If
    If Not iH1_NJYm.cChkResult Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    正しい内示年月を入力して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       Exit Function
    End If

' 発注内示テーブル SELECT文 ｾｯﾄ
    gSL_Select = "SELECT NH.手配先,NH.受入,NH.品番,NH.納期,NH.数量,"
    gSL_Select = gSL_Select & "     HT.SK区分,HT.手配区分,HT.背番号,HN.表示品番 品番編集,"
    gSL_Select = gSL_Select & "     ISNULL(TR.略称,'') 手配先名,ISNULL(SY.値名称,'') 手配区分名"
    gSL_Select = gSL_Select & "  FROM 発注内示テーブル NH"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN"
    gSL_Select = gSL_Select & "     ON  HN.品番 = NH.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番取引先マスタ HT"
    gSL_Select = gSL_Select & "     ON  HT.品番 = NH.品番"
    gSL_Select = gSL_Select & "     AND  HT.取引先 = NH.手配先"
    gSL_Select = gSL_Select & "     AND  HT.受入 = NH.受入"
    gSL_Select = gSL_Select & "     AND  HT.品目 = 1"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 内示発注チェックテーブル NC"
    gSL_Select = gSL_Select & "     ON  NC.区分 = 0"
    gSL_Select = gSL_Select & "     AND  NC.発注年月 = '" & RTrim(iH1_NJYm) & "'"
    gSL_Select = gSL_Select & "     AND  NC.手配先 = NH.手配先"
    gSL_Select = gSL_Select & "     AND  NC.受入 = NH.受入"
    gSL_Select = gSL_Select & "     AND  NC.品番 = NH.品番"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR"
    gSL_Select = gSL_Select & "     ON  TR.取引先区分 = 1"
    gSL_Select = gSL_Select & "     AND  TR.取引先CD = NH.手配先"
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 名称マスタ SY"
    gSL_Select = gSL_Select & "     ON  SY.区分名称 = '手配区分'"
    gSL_Select = gSL_Select & "     AND  SY.値 = HT.手配区分"
    
    ' 抽出条件
    mWhere = ""
    With CisFun
        .Where_CreateKB = NewDocument
        .Where_Create AlphaMe, "NH.内示年月", iH1_NJYm
        .Where_Create AlphaMe, "NH.計画年月", iH1_NJYm
        .Where_Create Numeric, "ISNULL(HT.SK区分,0)", 1
        .Where_Create Numeric, "ISNULL(HT.内示発注,0)", 1
        .Where_Create Numeric, "ISNULL(TR.社内区分,0)", 0, , Create
        .Where_Create AlphaMe, "NC.発注処理日", "", Nul
        ' 仕入先
        .Where_Create AlphaMe, "NH.手配先", iH1_Torcd
        ' 受入
        .Where_Create AlphaMe, "NH.受入", iH1_Ukeir
        ' 品番
        .Where_Create AlphaMe, "NH.品番", iH1_Hinbn
        
        gSL_Select = gSL_Select & .Where_Phrase
    End With
    
    gSL_Select = gSL_Select & " Order By NH.手配先,NH.受入,NH.品番,NH.納期"
    
    If Not HNJRead(gSL_Select) Then
       Call HNJClose
       With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    該当データがありません            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
       End With
       iH1_Torcd.SetFocus
       Exit Function
    End If
    
    lb_Message = "【表示中】しばらくお待ち下さい"
    Message.Visible = True
    DoEvents
    
    Call GridSet
    
    Message.Visible = False
    
    Head1Chk = True
End Function
'****************************
'*      Body1ﾁｪｯｸ処理
'****************************
Private Function Body1Chk() As Boolean
    
    Body1Chk = False
    
    If CisVsGrid.SelectCount = 0 Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    処理対象行を選択して下さい            "
            .MB_Title = ""
            .MB_Button = Error
            .MBOX
        End With
       Exit Function
    End If

    Body1Chk = True
End Function
'**************************************
'*                                    *
'*          データベース更新           *
'*                                    *
'**************************************
'+----------------------+
'+    発注処理 確認     +
'+----------------------+
Private Function HCProc() As Boolean
    HCProc = False
    Dim PrtCnt3     As Long
    Dim PrtCnt4     As Long
    Dim PrtCntD     As Long
    Dim OpUchiPrt   As Byte

    K_Sykbnm.Caption = SyoriNM
    K_Sykbnm.Visible = True
    
    With CisFun
        .MB_Lines = 4
        .MB_MSG(2) = "     発注処理を行います      "
        .MB_Title = "発注処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    On Error GoTo HCProc_Er

    lb_Message = "【 確定発注 処理中 】しばらくお待ち下さい！"
    Message.Visible = True
    DoEvents
    
    Call CisDB.DBTran(TransBegin)  ' ﾄﾗﾝｻﾞｸｼｮﾝ開始
'発注画面入力ワーク削除
    gSL_Select = "DELETE FROM 発注画面入力ワーク"
    gSL_Select = gSL_Select & " WHERE 処理端末 = HOST_NAME() "
    gSL_Select = gSL_Select & "   AND 作成種別 = 'N' "
    CisDB.SQL = gSL_Select
    CisDB.DBExec
    
' 発注画面ワークに出力
    With vsGrid1
        For gLong = 1 To .Rows - 1

            If Trim(.TextMatrix(gLong, CisVsGrid.FixedGet(CisVsGrid.SelectCol))) = "X" Then
                Call HGWInsert_Set
            End If
        Next gLong
    End With
    
    Call CisDB.DBTran(TransCommit)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 終了
    
'発注打切更新
    With CisDB
        .SQL = "内示発注確定処理"
        .StoadoCount = 7
        .ParaNo = 0: .ParaIO = Return用
        .ParaNo = 1: .ParaIO = Input用            '内示年月
        .ParaNo = 2: .ParaIO = OutPut用            '入力件数
        .ParaNo = 3: .ParaIO = OutPut用            '発注打切テーブル更新件数
        .ParaNo = 4: .ParaIO = OutPut用            '発注打切警告リスト件数
        .ParaNo = 5: .ParaIO = OutPut用            '発注打切リスト件数
        .ParaNo = 6: .ParaIO = OutPut用            '発注打切完了済リスト件数
        .ParaNo = 7: .ParaIO = OutPut用            'エラーコード
'
        .ParaNo = 1: .ParaValue = RTrim(iH1_NJYm)  '内示年月
        
        If Not .DBStored Then
            MsgBox "ストアドエラー "
            GoTo HCProc_STErr
        Else
            .ParaNo = 7
            If .ParaValue <> 0 Then GoTo HCProc_STErr
            .ParaNo = 4
            PrtCnt3 = .ParaValue
            .ParaNo = 5
            PrtCnt4 = .ParaValue
            .ParaNo = 6
            PrtCntD = .ParaValue
        End If
    End With

'+---------------------------------------------------------------------------------  2005/11/07 Insert
    ' 発注打切リスト　発行
    Call ItemsClearIRN
    IRN.bk = "Option": IRN.rno = 36
    Call IRNRead("", 1)
    OpUchiPrt = CisFun.Val2(Mid(IRN.oyian, 1, 1))
'+---------------------------------------------------------------------------------  2005/11/07 Insert End
'発注打切警告リスト印刷
    If PrtCnt3 <> 0 And OpUchiPrt = 0 Then                                               '2005/11/07追加
        lb_Message = "【 発注打切警告リスト 印刷中 】"
        DoEvents
        CHH0930_CallType = ""
        If Not CKH0930LPrint Then Exit Function
        DoEvents
    End If
'発注打切リスト印刷
    If PrtCnt4 <> 0 And OpUchiPrt = 0 Then                                               '2005/11/07追加
        lb_Message = "【 発注打切リスト 印刷中 】"
        DoEvents
        CHH0940_CallType = ""
        If Not CKH0940LPrint Then Exit Function
        DoEvents
    End If
'発注完了済リスト印刷
    If PrtCntD <> 0 And OpUchiPrt = 0 Then                                               '2005/11/07追加
        lb_Message = "【 発注打切完了済リスト 印刷中 】"
        DoEvents
        RV_SYmd = RTrim(iH1_NJYm)
        If Not CXH0040LPrint Then Exit Function
        DoEvents
    End If
    
    Message.Visible = False
    DoEvents

' 発注処理
    CisDB.PassLink
    With CisFun
        .ShellApl = gIniExe & "PXH0110.exe"
        .ShellAplPara = "G"
        .ShellTask NoFreeze
    End With
    
    If PrtCntD <> 0 Then
        With CisFun
            .MB_Lines = 6
            .MB_MSG(1) = "     打切完了済の品番が　" & PrtCntD & " 件存在しました。      "
            .MB_MSG(3) = "     打切完了済の品番は発注されません！      "
            .MB_MSG(5) = "     発注打切完了済リストで確認して下さい      "
            .MB_Title = "発注打切済確認"
            .MB_Button = OK
            .MBOX
        End With
    End If


    HCProc = True
    Exit Function
HCProc_Er:
    Call CisDB.DBTran(TransRollback)  ' ﾄﾗﾝｻﾞｸｼｮﾝ 破棄
    CisFun.ErrorBox
    Exit Function
HCProc_STErr:
    Select Case CisDB.ParaValue
        Case 1: gStr = "    【　かんばん読取エラーテーブル　削除エラー！】"
        Case 2: gStr = "    【　帳票出力エラーテーブル　削除エラー！】"
        Case 3: gStr = "    【　発注打切テーブル(完了)　更新エラー！】"
        Case 5: gStr = "    【　発注打切テーブル(未完)　更新エラー！】"
        Case 6: gStr = "    【　帳票出力テーブル(警告)　出力エラー！】"
        Case 7: gStr = "    【　帳票出力テーブル(打切)　出力エラー！】"
        Case 8: gStr = "    【　帳票出力テーブル(打切完了済)　出力エラー！】"
        Case 9: gStr = "    【　かんばん読取エラーテーブル　出力エラー！】"
        Case 10: gStr = "    【　かんばん読取エラーテーブル(今回発注数)更新エラー！】"
        Case 11: gStr = "    【　かんばん読取エラーテーブル(発注済数)更新エラー！】"
        Case 12: gStr = "    【　発注画面入力ワーク(処理)　更新エラー！】"
        Case 60: gStr = "    【　エラー管理番号　獲得エラー！】"
    End Select
    MsgBox gStr & "  Code = " & CisDB.ParaValue & "     "

End Function

Private Function HGWInsert_Set() As Boolean
    HGWInsert_Set = False

    Call ItemsClearHGW
    With HGW
        .発注区分 = 5
        .仕入先 = RTrim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(1)))
        .受入 = RTrim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(3)))
        .納入日 = RTrim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(10)))
        .品番 = RTrim(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(9)))
        .発注数 = CisFun.Val2(vsGrid1.TextMatrix(gLong, CisVsGrid.FixedGet(6)))
        .作成種別 = "N"
        .処理区分 = 0
    End With
    Call HGWInsert

    HGWInsert_Set = True
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
        Call .FixedSet(中中, 右中, あり, 4, 0, "№")
        Call .FixedSet(中中, 左中, あり, 7, 1, "仕入先")
        Call .FixedSet(中中, 左中, あり, 19, 2, "仕入先名")
        Call .FixedSet(中中, 中中, あり, 4, 3, "受入")
        Call .FixedSet(中中, 左中, あり, 25, 4, "品　番")
        Call .FixedSet(中中, 中中, あり, 11, 5, "納 入 日")
        Call .FixedSet(中中, 右中, あり, 8, 6, "発注数")
        Call .FixedSet(中中, 中中, あり, 6, 7, "手配区")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE START
        'Call .FixedSet(中中, 中中, あり, 6, 8, "背番号")
        Call .FixedSet(中中, 中中, あり, 8, 8, "背番号")
'-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-*-* 2008/04/11 UPDATE END
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, あり, 25, 9, "X品番")
        Call .FixedSet(中中, 左中, なし, 0, 10, "X納期")
        Call .FixedSet(中中, 左中, なし, 0, 11, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8")
        
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
    
    CisVsGrid.Clear
' 抽出処理 初期ｾｯﾄ
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避
    B1lb_DspSu = Format(CisDB.RecordCount, "#,###")
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
    End With
    ID = 0
    Do Until Not HNJ_RDSTS
        With vsGrid1
            ID = ID + 1
            '№
            .TextMatrix(ID, GridCol(0)) = Format(ID, "####")
            '仕入先・名
            .TextMatrix(ID, GridCol(1)) = RTrim(HNJ.手配先)
            .TextMatrix(ID, GridCol(2)) = RTrim(HNJ.手配先名)
            '受入
            .TextMatrix(ID, GridCol(3)) = RTrim(HNJ.受入)
            '品番
            .TextMatrix(ID, GridCol(4)) = RTrim(HNJ.品番編集)
            ' 納入日
            If Trim(HNJ.納期) <> "" Then
               With CisFun
                    .DateE_BefVal = HNJ.納期
                    .DateEdit
                    gStr = .DateE_AfVal
               End With
               .TextMatrix(ID, GridCol(5)) = gStr
            End If
            ' 発注数
            .TextMatrix(ID, GridCol(6)) = Format(HNJ.数量, "#,###")
            ' 手配区分
            .TextMatrix(ID, GridCol(7)) = RTrim(HNJ.手配区分名)
            ' 背番号
            .TextMatrix(ID, GridCol(8)) = RTrim(HNJ.背番号)
            '品番
            .TextMatrix(ID, GridCol(9)) = RTrim(HNJ.品番)
            '納期
            .TextMatrix(ID, GridCol(10)) = RTrim(HNJ.納期)
            
        End With
        Call HNJReadNext
    Loop
    Call HNJClose
    With vsGrid1
         .Row = 1: .Col = 1: .ColSel = .Cols - 1
         .RowSel = Ttl_Kensu
          Call CisVsGrid.vsColor
         .Redraw = True
    End With
End Sub
'******************************************'
'*****        行選択（クリック）       *****
'******************************************'
Private Sub VSGrid1_Click()
    If SyoriKB = "" Then
       Call CisVsGrid.vsColor
    Else
       Call CisVsGrid.vsColor(, &HC0C0FF)
    End If
End Sub
'******************************************'
'*****        行選択（スペースｷｰ)      *****
'******************************************'
Private Sub VSGrid1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    Call VSGrid1_Click
End Sub

'Private Sub vsGrid1_RowColChange()
'    Call GridSet2
'End Sub
'******************************************'
'*****       移動選択項目選択          *****
'******************************************'
Private Sub VSGrid1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vsGrid1
        .Tag = ""
    
        If Button <> vbLeftButton Then Exit Sub
        If .MouseRow <> 0 Or .MouseCol = 0 Then Exit Sub
    
        .Tag = Str(.MouseCol)
        .Drag vbBeginDrag
    End With
End Sub
'******************************************'
'*****        表 示 項 目 移 動        *****
'******************************************'
Private Sub VSGrid1_DragDrop(Source As Control, X As Single, Y As Single)
    With vsGrid1
        If .MouseCol = 0 Then Exit Sub
        If .Tag = "" Then Exit Sub
    
        .Redraw = False
    
        .ColPosition(CisFun.Val2(.Tag)) = .MouseCol
        .Col = 1: .ColSel = .Cols - 1
    
        .Redraw = True
    
        CisVsGrid.VSSort
    End With
End Sub
