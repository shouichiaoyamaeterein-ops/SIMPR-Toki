VERSION 5.00
Object = "{8FF8DA3B-2BD7-11D4-806A-00000EA57F9E}#4.0#0"; "CIS3D_~1.OCX"
Object = "{B5AA0C17-2C0D-11D4-831F-009027089EFF}#2.0#0"; "CISBTN~1.OCX"
Object = "{90954F85-2B4C-11D4-806A-00000EA57F9E}#7.0#0"; "CISTEX~1.OCX"
Object = "{021DB354-2B52-11D4-806A-00000EA57F9E}#2.1#0"; "CISYM_~1.OCX"
Object = "{8AE029D0-08E3-11D1-BAA2-444553540000}#3.0#0"; "Vsflex3.ocx"
Begin VB.Form CNJ0050 
   BackColor       =   &H00808000&
   BorderStyle     =   1  '固定(実線)
   Caption         =   "発注内示変更一覧"
   ClientHeight    =   11235
   ClientLeft      =   1665
   ClientTop       =   825
   ClientWidth     =   15390
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
   PaletteMode     =   1  'Z ｵｰﾀﾞｰ
   ScaleHeight     =   11235
   ScaleWidth      =   15390
   Begin Cis3D_v60.CIS3D Aplexec_MSG 
      Height          =   1260
      Left            =   4440
      Top             =   5220
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2223
      BackColor       =   8421631
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
      cBoderWidth     =   7
      Begin Cis3D_v60.CIS3D Exe_MSG 
         Height          =   915
         Left            =   195
         Top             =   165
         Width           =   5910
         _ExtentX        =   10425
         _ExtentY        =   1614
         BackColor       =   14737632
         ForeColor       =   12582912
         Caption         =   "[発注内示入力]を起動しています"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   15.75
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
         cBoderWidth     =   7
         cAlingnment     =   7
         cBoderStyle     =   1
      End
   End
   Begin Cis3D_v60.CIS3D Proc_MSG 
      Height          =   1170
      Left            =   9525
      Top             =   2880
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   2064
      BackColor       =   16744576
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
      cBoderWidth     =   5
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D5 
         Height          =   930
         Left            =   135
         Top             =   120
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   1640
         ForeColor       =   8421376
         Caption         =   "『データ抽出中』しばらくお待ち下さい"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐ明朝"
            Size            =   14.25
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
      End
   End
   Begin Cis3D_v60.CIS3D H1_Area2 
      Height          =   645
      Left            =   120
      Top             =   510
      Width           =   2835
      _ExtentX        =   5001
      _ExtentY        =   1138
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
      cFont3DColor1   =   16777215
      cFont3DColor2   =   8421504
      cBoderStyle     =   1
      Begin Cis3D_v60.CIS3D CIS3D11 
         Height          =   555
         Left            =   45
         Top             =   45
         Width           =   2745
         _ExtentX        =   4842
         _ExtentY        =   979
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
         cAlingnment     =   1
         cPositionX      =   180
         Begin Cis3D_v60.CIS3D H1lb_NJYM 
            Height          =   390
            Left            =   1230
            Top             =   75
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   688
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   7
            cBoderStyle     =   1
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
      Height          =   1005
      Left            =   105
      ScaleHeight     =   945
      ScaleWidth      =   12240
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   1260
      Width           =   12300
      Begin Cis3D_v60.CIS3D CIS3D10 
         Height          =   825
         Left            =   8955
         Top             =   75
         Width           =   3210
         _ExtentX        =   5662
         _ExtentY        =   1455
         ForeColor       =   16711680
         Caption         =   "増減率"
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
         cPositionX      =   90
         cPositionY      =   30
         Begin Cis3D_v60.CIS3D H1lb_Ritu 
            Height          =   690
            Left            =   870
            Top             =   75
            Width           =   2265
            _ExtentX        =   3995
            _ExtentY        =   1217
            BackColor       =   16761087
            ForeColor       =   16711680
            Caption         =   "90%以下 110%以上"
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
            cFont3DColor1   =   16777215
            cFont3DColor2   =   8421504
            cAlingnment     =   8
            cBoderStyle     =   1
            cPositionY      =   -30
            Begin Cis3D_v60.CIS3D CIS3D4 
               Height          =   285
               Left            =   645
               Top             =   60
               Width           =   1035
               _ExtentX        =   1826
               _ExtentY        =   503
               BackColor       =   16761087
               ForeColor       =   16711680
               Caption         =   "表示範囲"
               BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
                  Name            =   "ＭＳ ゴシック"
                  Size            =   11.25
                  Charset         =   128
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               cBoderColor1    =   16761087
               cBoderColor2    =   16761087
               cFont3DColor1   =   16777215
               cFont3DColor2   =   8421504
               cBoderStyle     =   2
            End
         End
         Begin CisText_V60.CisText iH1_Ritu 
            Height          =   375
            Left            =   180
            TabIndex        =   5
            Top             =   360
            Width           =   510
            _ExtentX        =   900
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
            cILength        =   3
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0C0C0&
            Caption         =   "%"
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
            Left            =   705
            TabIndex        =   10
            Top             =   480
            Width           =   165
         End
      End
      Begin Cis3D_v60.CIS3D CIS3D1 
         Height          =   825
         Left            =   45
         Top             =   75
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   1455
         ForeColor       =   16711680
         Caption         =   "対象年月"
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
         Begin CisYM_V60.CisYM iH1_YM 
            Height          =   375
            Left            =   60
            TabIndex        =   1
            Top             =   360
            Width           =   1350
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
      Begin Cis3D_v60.CIS3D BLB_HINBN 
         Height          =   825
         Left            =   1515
         Top             =   75
         Width           =   3300
         _ExtentX        =   5821
         _ExtentY        =   1455
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
            Height          =   375
            Left            =   60
            TabIndex        =   2
            Top             =   360
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
      Begin Cis3D_v60.CIS3D CIS3D2 
         Height          =   825
         Left            =   4815
         Top             =   75
         Width           =   4140
         _ExtentX        =   7303
         _ExtentY        =   1455
         ForeColor       =   16711680
         Caption         =   "手 配 先"
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
            Height          =   375
            Left            =   90
            TabIndex        =   3
            Top             =   360
            Width           =   990
            _ExtentX        =   1746
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
         End
         Begin Cis3D_v60.CIS3D H1lb_Tornm 
            Height          =   360
            Left            =   1080
            Top             =   360
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
         Begin Cis3D_v60.CIS3D BLB_DD 
            Height          =   300
            Index           =   47
            Left            =   3525
            Top             =   30
            Width           =   540
            _ExtentX        =   953
            _ExtentY        =   529
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
            cBoderColor1    =   12632256
            cBoderColor2    =   12632256
            cFont3DColor1   =   -2147483643
            cFont3DColor2   =   8421504
            cBoderWidth     =   2
            cAlingnment     =   7
            cBoderStyle     =   2
         End
         Begin CisText_V60.CisText iH1_Ukeir 
            Height          =   375
            Left            =   3615
            TabIndex        =   4
            Top             =   360
            Width           =   360
            _ExtentX        =   635
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
         End
      End
   End
   Begin VB.PictureBox B1_Area1 
      BackColor       =   &H00C0C0C0&
      Height          =   8670
      Left            =   105
      ScaleHeight     =   8610
      ScaleWidth      =   15075
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   2430
      Width           =   15135
      Begin vsFlexLib.vsFlexArray vsGrid1 
         Height          =   8445
         Left            =   60
         TabIndex        =   6
         Top             =   75
         Width           =   14940
         _Version        =   196608
         _ExtentX        =   26352
         _ExtentY        =   14896
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
         ConvInfo        =   1418783674
         FormatString    =   "№  |品番                  |手配先                |受入|今回計画数|前回計画数|計画差   |率  |今回必要数|前回必要数|必要差   |"
         Rows            =   28
         Cols            =   13
         BackColor       =   12648447
         ForeColorFixed  =   128
         BackColorBkg    =   12632256
         FocusRect       =   0
         HighLight       =   2
         RowHeightMin    =   300
         SelectionMode   =   1
         BackColorAlternate=   12648447
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
      Height          =   765
      Left            =   10920
      ScaleHeight     =   705
      ScaleWidth      =   4260
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   420
      Width           =   4320
      Begin CisBtn_60.CisBtn PB_CAN 
         Height          =   585
         Left            =   1710
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
         Left            =   2535
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
         Left            =   3360
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
         Left            =   885
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
      Begin CisBtn_60.CisBtn PB_DEL 
         Height          =   585
         Left            =   60
         Top             =   60
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
   End
   Begin VB.TextBox Dummy 
      BeginProperty Font 
         Name            =   "ＭＳ ゴシック"
         Size            =   8.25
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      IMEMode         =   3  'ｵﾌ固定
      Left            =   105
      TabIndex        =   0
      Text            =   "Dummy"
      Top             =   450
      Width           =   180
   End
   Begin VB.Timer Timer1 
      Left            =   3240
      Top             =   750
   End
   Begin VB.Timer Timer2 
      Left            =   3255
      Top             =   450
   End
   Begin Cis3D_v60.CIS3D PGName 
      Height          =   375
      Left            =   0
      Top             =   0
      Width           =   15360
      _ExtentX        =   27093
      _ExtentY        =   661
      BackColor       =   14737632
      ForeColor       =   0
      Caption         =   "【 発注内示変更一覧  】"
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
         Left            =   14430
         Top             =   60
         Width           =   825
         _ExtentX        =   1455
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
         Top             =   30
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
      Begin Cis3D_v60.CIS3D B1_Area2 
         Height          =   330
         Left            =   10200
         Top             =   15
         Width           =   1590
         _ExtentX        =   2805
         _ExtentY        =   582
         Caption         =   "件数"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "ＭＳ Ｐゴシック"
            Size            =   12
            Charset         =   128
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         cBoderColor1    =   -2147483643
         cFont3DColor1   =   -2147483643
         cFont3DColor2   =   8421504
         cAlingnment     =   1
         cPositionX      =   90
         Begin Cis3D_v60.CIS3D B2lb_Cnt 
            Height          =   270
            Left            =   630
            Top             =   30
            Width           =   915
            _ExtentX        =   1614
            _ExtentY        =   476
            Caption         =   "99,999"
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
            cPositionX      =   -30
         End
      End
      Begin Cis3D_v60.CIS3D K_Sykbnm 
         Height          =   345
         Left            =   0
         Top             =   0
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
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
Attribute VB_Name = "CNJ0050"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************'
'**
'**       フォーム名    :   発注内示変更一覧
'**       フォームID    :   CNJ0050
'**       処理概要      :
'**
'**       作  成  日    :   2004/04/07  By CIS
'**       変  更  日    :   2004/06/28  By CIS - 手配先の検索は発注全てを対象にする
'**       変  更  日    :   2004/06/29  By CIS - 手配先の検索を実行した時、取引先CDがｾｯﾄされない＆検索項目クリア
'**       変  更  日    :   2005/02/10  By CIS - 品番検索時、品目のデフォルト値を修正（受注→発注）
'**
'**********************************************************'
    Dim CisFrm                  As CISFormContorl
    
    Dim FormAct                 As Boolean      ' ﾌｫｰﾑｱｸﾃｨﾌﾞ ﾌﾗｸﾞ
    Dim ProcHB                  As String * 2   ' 入力域切替
    Dim SyoriKB                 As String       ' 処理区分(新規･修正･削除 等)
    Dim SyoriNM                 As String       ' 処理名称(新規･修正･削除 等)
    Dim CisVsGrid               As New CisVsGrid3

    Dim Ritu_Up                 As Integer      ' 増減率上限値
    Dim Ritu_Dw                 As Integer      ' 増減率下限値

    Dim sv_Hinbn                As String
    Dim sv_Torcd                As String
    Dim sv_Ukeir                As String
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
    K_Sykbnm = ""
    Timer1.Interval = 1000   ' 時刻表示間隔
    Timer1.Enabled = True    ' 時刻表示開始
    Timer2.Interval = 500    ' 処理区分点滅間隔
    Timer2.Enabled = False   ' 処理区分点滅禁止

'( 抽出中メッセージ )
    Set CisFrm.ObjetItem = Proc_MSG
    CisFrm.OBJPosition
    Proc_MSG.Visible = False

'( メンテ起動中メッセージ )
    Set CisFrm.ObjetItem = Aplexec_MSG
    CisFrm.OBJPosition

'   *-------------------*
'   *   ｸﾞﾘｯﾄﾞ初期処理
'   *-------------------*
    With CisVsGrid
        Set .GridObj = vsGrid1
        .NoSpaceCol = 1
        .NumberSet = NoSpace
        .FixedHeight = 330
        .PatanMax = 1
        .SelectCol = 12
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
    
'   #-------------------#
'   # 画面中央位置表示   #
'   #-------------------#
    Call CisFrm.Frm_Center(True)
    
    Me.Show
End Sub
'+-------------------------------------------------------------+
'+      フォーム    《Ａｃｔｉｖａｔｅ イベント》                +
'+-------------------------------------------------------------+
Private Sub Form_Activate()
    If FormAct Then Exit Sub
    FormAct = Not FormAct
    
'   < 表示範囲 >
    Call DispRange

'   < 発注内示対象年月 >
    Call GetNJProc_YM(3)
    If RTrim$(gn_HNYM) = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    発注内示対象年月を獲得出来ません"
            .MB_Button = Error
            .MBOX
            Call PB_END_Click
            Exit Sub
        End With
    End If
    
    H1lb_NJYM = Format(gn_HNYM, "@@@@年@@月")
    iH1_YM = gn_HNYM

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
           Case vbKeyF3:        If PB_DEL.Visible Then Call PB_DEL_Click    '【削除】
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
    Call CisDB.DBDISConnect                            ' ﾃﾞｰﾀﾍﾞｰｽ切断
    Call CisFrm.UnLoadDisp                          ' ﾌｫｰﾑｵﾌﾞｼﾞｪｸﾄ消去
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
        SyoriKB = ""
        Call SyoriName(SyoriKB)
        iH1_YM.SetFocus
        GoTo PB_CAN_Ed
    End If
    
PB_CAN_Ed:
    Call CisFrm.MousePT(1)
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
    vsGrid1.SetFocus
End Sub
'+----------------------------+
'+    検  索  キ ー(F7)       +
'+----------------------------+
Private Sub PB_Look_Click()
    If PB_Look.Tag = "iH1_Hinbn" Then Call Look_Hinbn: Exit Sub
    If PB_Look.Tag = "iH1_Torcd" Then Call Look_SKTable: Exit Sub
    If PB_Look.Tag = "iH1_Ukeir" Then Call Look_SKTable: Exit Sub
End Sub
'+-----------------------------------+
'+       品番検索                     +
'+-----------------------------------+
Private Sub Look_Hinbn()
    RV_Left = 0
    RV_Top = 0
    RV_Call = ""
'*---------------------------------------------------* ▽▽▽ 2005/02/10 Update Start ▽▽▽
'   RVI_Hinmoku = 0
    RVI_Hinmoku = 1         '>発注
'*---------------------------------------------------* △△△ 2005/02/10 Update End   △△△
    RVI_SK = 9
    
    CKK0035.Show vbModal
    Unload CKK0035
    Set CKK0035 = Nothing

    If RV_Rtn Then
        iH1_Hinbn = RV_Hinbn
        iH1_Torcd = RV_TorcdK
        iH1_Ukeir = RV_Ukeir
        
        Call TorNmGet(iH1_Torcd, 0)
        H1lb_Tornm = RTrim(TRM_RName)
    End If
End Sub
'+-----------------------------------+
'+       手配先検索
'+-----------------------------------+
Private Sub Look_SKTable()
    RV_Left = 0
    RV_Top = 0
'*---------------------------------------------<< 2004/06/28 Update Start >>
'   RV_TorKb = 7
    RV_TorKb = 9
'*---------------------------------------------<< 2004/06/28 Update End >>
'*---------------------------------------------<< 2004/06/29 Insert Start >>
    RV_Torcd = ""
    RV_Kojyo = ""
'*---------------------------------------------<< 2004/06/29 Insert Start >>
    
    CKK0028.Show vbModal
    Unload CKK0028
    Set CKK0028 = Nothing

    If RV_Rtn Then
'*---------------------------------------------<< 2004/06/29 Update Start >>
'       iH1_Torcd = RV_Torcd
        iH1_Torcd = RV_TorcdK
'*---------------------------------------------<< 2004/06/29 Update End >>
        H1lb_Tornm = RTrim$(RV_TorRName)
        iH1_Ukeir = RV_Ukeir
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

    Call TorNmGet(iH1_Torcd, 0)
    H1lb_Tornm = RTrim(TRM_RName)
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
'+  iH1_Ritu
'+-----------------------------+
Private Sub iH1_Ritu_LostFocus()
    Call DispRange
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
'**************************************
'*                                    *
'*      処理区分により区分名表示        *
'*                                    *
'**************************************
Private Sub SyoriName(SYKB As String)
    Dim W_Color     As Long
    W_Color = &HC0FFFF    '初期色ｾｯﾄ(１番薄い黄色)
    Select Case SyoriKB
      Case "DEL"
           SyoriNM = "削除"
           K_Sykbnm.ForeColor = &HFF&              '(赤色)
           W_Color = &HC0C0FF                   '(１番薄い赤色)
      Case Else
           SyoriNM = ""
    End Select

    vsGrid1.BackColor = W_Color
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
    
    H1Mode = False:     H1Color = gPLostSel
    B1Mode = False:     B1Color = gPLostSel
    B2Mode = False
    
'   ﾀﾞﾐｰ項目(画面のチラツキ防止)
    If Dummy.Visible Then Dummy.Enabled = True: Dummy.SetFocus
    
    Select Case Pro
        Case "H1":      H1Mode = True:  H1Color = gPGotSel
        Case "B1":      B1Mode = True:  B1Color = gPGotSel
        Case "B2":      B2Mode = True
    End Select
    
    H1_Area1.Enabled = H1Mode
    H1_Area1.BackColor = H1Color
    B1_Area1.Enabled = B1Mode
    B1_Area1.BackColor = B1Color
    B1_Area2.Visible = B1Mode
    
    Aplexec_MSG.Visible = B2Mode
    If H1Mode Then
        vsGrid1.HighLight = flexHighlightWithFocus
    Else
        vsGrid1.HighLight = flexHighlightAlways
    End If
    
    PB_Look.Visible = False
    PB_CAN.Visible = Not H1Mode
    PB_END.Visible = H1Mode
    PB_DEL.Visible = B1Mode
    
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
            If .Name Like wbc_Enm & "Op_*" Then .Value = False
            If .Name Like wbc_Enm & "lb_*" Then .Caption = ""
        End With
    Next gObj

    If HeadBody = "B1" Or HeadBody = "B" Then CisVsGrid.Clear
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
        GoTo ReturnPress_Ed
    End If

    If ProcHB = "B1" Then
        If Not Body1Chk Then GoTo ReturnPress_Ed
        
        If SyoriKB = "DEL" Then
            If Not DBDelete Then GoTo ReturnPress_Ed
            If Not BodySet Then
                With CisFun
                    .MB_Lines = 4
                    .MB_MSG(2) = "    対象データがなくなりました         "
                    .MB_Button = Error
                    .MBOX
                End With
                Call PB_CAN_Click
                GoTo ReturnPress_Ed
            End If
        Else
            ProcHB = "B2"
            Call DispChange(ProcHB): DoEvents
    
'           < メンテ起動 >
            Call AplExec
        
'           < メンテデータ再表示 >
            Call MNTData_Disp
        
            ProcHB = "B1"
            Call DispChange(ProcHB): DoEvents
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
    
' < 対象年月 >
    If iH1_YM = "" Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    対象年月を入力して下さい         "
            .MB_Button = Error
            .MBOX
        End With
        iH1_YM.SetFocus
        Exit Function
     End If

' < 増減値 >
    Call DispRange
    
    If Not BodySet Then
        With CisFun
            .MB_Lines = 4
            .MB_MSG(2) = "    対象データがありません         "
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
    
    With vsGrid1
        If RTrim$(.TextMatrix(.Row, 11)) = "" Then Exit Function         '品番
    End With
    
    Body1Chk = True
End Function
'*-----------------------------------------------------------------------*
'*
'*          発  注  内  示　入　力 　( PNJ0030.EXE ) 　起　動
'*
'*-----------------------------------------------------------------------*
Private Sub AplExec()
    
    With vsGrid1
        sv_Hinbn = RTrim$(.TextMatrix(.Row, 11))
        sv_Torcd = RTrim$(Mid$(.TextMatrix(.Row, 2), 1, 7))
        sv_Ukeir = RTrim$(.TextMatrix(.Row, 3))
    End With
    
    gStr = ""
    gStr = gStr & gn_HNYM & ","
    gStr = gStr & RTrim$(iH1_YM) & ","
    gStr = gStr & sv_Hinbn & ","
    gStr = gStr & sv_Torcd & ","
    gStr = gStr & sv_Ukeir
    
'< 発注内示入力起動 >
    With CisFun
        .ShellApl = gIniExe & "PNJ0030.Exe"
        .ShellAplPara = gStr
        .ShellTask NoFreeze
    End With

End Sub
'*-----------------------------------------------------------------------*
'*
'*                  メ　ン　テ　デ　ー　タ　再　表　示
'*
'*-----------------------------------------------------------------------*
Private Sub MNTData_Disp()
    
'   ( Select 文 作成 )
    Call CrtSelect(1)
    
    Call HNJRead(gSL_Select, 1)

    With vsGrid1
        
        .TextMatrix(.Row, 4) = Format(HNJ.今回計画数, "#,###")
        .TextMatrix(.Row, 5) = Format(HNJ.前回計画数, "#,###")
        .TextMatrix(.Row, 6) = Format(HNJ.計画差, "#,###")
        
        If HNJ.増減率 > 999 Then
            .TextMatrix(.Row, 7) = "999"
        Else
            .TextMatrix(.Row, 7) = Format(HNJ.増減率, "#")
        End If
    
    End With

End Sub
'*----------------------*
'*       BODY表示       *
'*----------------------*
Private Function BodySet() As Boolean
    BodySet = False
    Proc_MSG.Visible = True: DoEvents
    Call HeadBodyClear("B")
    
' ( SELECT文 作成 )
    Call CrtSelect(0)
    On Error GoTo BodySet_Err
    
    If HNJRead(gSL_Select) Then
        Call GridSet
        BodySet = True
    End If
    
    On Error GoTo 0
    Proc_MSG.Visible = False: DoEvents
    Exit Function

BodySet_Err:
     CisFun.ErrorBox
End Function
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
        .MB_Lines = 4
        .MB_MSG(2) = "     削除処理を行います      "
        .MB_Title = "削除処理"
        .MB_MSG(3) = "     よろしいですか？      "
        .MB_Button = OK_CAN
        If Not .MBOX Then Exit Function
    End With
    
    With vsGrid1
        sv_Hinbn = RTrim$(.TextMatrix(.Row, 11))
        sv_Torcd = RTrim$(Mid$(.TextMatrix(.Row, 2), 1, 7))
        sv_Ukeir = RTrim$(.TextMatrix(.Row, 3))
    End With
        
    With CisDB
        On Error GoTo DBDelete_Err
        .DBTran (TransBegin)

        .SQL = "DELETE 発注内示テーブル "
        .SQL = .SQL & " Where 内示年月 = '" & RTrim$(gn_HNYM) & "'"
        .SQL = .SQL & " And   SUBSTRING(納期,1,6) = '" & RTrim$(iH1_YM) & "'"
        .SQL = .SQL & " And   品番 = '" & RTrim$(sv_Hinbn) & "'"
        .SQL = .SQL & " And   手配先 = '" & RTrim$(sv_Torcd) & "'"
        .SQL = .SQL & " And   受入 = '" & RTrim$(sv_Ukeir) & "'"
        .DBExec
        
        .SQL = "DELETE 発注内示前回テーブル "
        .SQL = .SQL & " Where 内示年月 = '" & RTrim$(gn_HNYM) & "'"
        .SQL = .SQL & " And   SUBSTRING(納期,1,6) = '" & RTrim$(iH1_YM) & "'"
        .SQL = .SQL & " And   品番 = '" & RTrim$(sv_Hinbn) & "'"
        .SQL = .SQL & " And   手配先 = '" & RTrim$(sv_Torcd) & "'"
        .SQL = .SQL & " And   受入 = '" & RTrim$(sv_Ukeir) & "'"
        .DBExec
        
        .DBTran (TransCommit)
        
        On Error GoTo 0
    End With
    
    DBDelete = True
    Exit Function
DBDelete_Err:
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
        Call .FixedSet(中中, 右中, あり, 4, 0, "№")
        Call .FixedSet(中中, 左中, あり, 22, 1, "品番")
        Call .FixedSet(中中, 左中, あり, 22, 2, "手配先")
        Call .FixedSet(中中, 中中, あり, 4, 3, "受入")
        Call .FixedSet(中中, 右中, あり, 10, 4, "今回計画数")
        Call .FixedSet(中中, 右中, あり, 10, 5, "前回計画数")
        Call .FixedSet(中中, 右中, あり, 9, 6, "計画差")
        Call .FixedSet(中中, 右中, あり, 4, 7, "率")
        Call .FixedSet(中中, 右中, あり, 10, 8, "今回必要数")
        Call .FixedSet(中中, 右中, あり, 10, 9, "前回必要数")
        Call .FixedSet(中中, 右中, あり, 9, 10, "必要差")
        Call .FixedSet(中中, 左中, あり, 0, 11, "品番編集前")
'　ｸﾞﾘｯﾄﾞ選択処理用　退避ｴﾘｱ
        Call .FixedSet(中中, 左中, なし, 0, 12, "Select")
'  表示ﾊﾟﾀｰﾝｾｯﾄ
'       Call .PatanSet(ﾊﾟﾀｰﾝNO,表示パターン)
        Call .PatanSet(0, "0,1,2,3,4,5,6,7,8,9,10")
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
    Dim Ttl_Kensu   As Long

' 抽出処理 初期ｾｯﾄ
    B2lb_Cnt = Format(CisDB.RecordCount, "#,##0")
    Ttl_Kensu = CisDB.RecordCount       ' 抽出件数(合計)退避

' ｸﾞﾘｯﾄﾞ行再ｾｯﾄ & 列位置取得
    With vsGrid1
        .Redraw = False
         If Ttl_Kensu + 1 < CisVsGrid.DispMax Then
             .Rows = CisVsGrid.DispMax
         Else
             .Rows = Ttl_Kensu + 1
         End If
    End With
    
    ID = 0
    Do Until Not HNJ_RDSTS
        ID = ID + 1
            
        With vsGrid1
            .TextMatrix(ID, 0) = ID
            
            If RTrim$(HNJ.表示品番) <> "" Then
                .TextMatrix(ID, 1) = RTrim$(HNJ.表示品番)
            Else
                .TextMatrix(ID, 1) = RTrim$(HNJ.品番)
            End If
            
            If RTrim$(HNJ.手配先) = "" Then
                .TextMatrix(ID, 2) = ""
            Else
                .TextMatrix(ID, 2) = RTrim$(HNJ.手配先) & Space(7 - Len(RTrim$(HNJ.手配先))) & ":" & RTrim$(HNJ.手配先名)
            End If
            .TextMatrix(ID, 3) = RTrim$(HNJ.受入)
            .TextMatrix(ID, 4) = Format(HNJ.今回計画数, "#,###")
            .TextMatrix(ID, 5) = Format(HNJ.前回計画数, "#,###")
            .TextMatrix(ID, 6) = Format(HNJ.計画差, "#,###")
            
            If HNJ.増減率 > 999 Then
                .TextMatrix(ID, 7) = "999"
            Else
                .TextMatrix(ID, 7) = Format(HNJ.増減率, "#")
            End If
        
            .TextMatrix(ID, 8) = Format(HNJ.今回必要数, "#,###")
            .TextMatrix(ID, 9) = Format(HNJ.前回必要数, "#,###")
            .TextMatrix(ID, 10) = Format(HNJ.必要差, "#,###")
            .TextMatrix(ID, 11) = RTrim$(HNJ.品番)
        End With
        
        Call HNJReadNext
    Loop
    
    With vsGrid1
        .Row = 1: .Col = 1: .ColSel = .Cols - 1
        .Redraw = True
    End With
    
End Sub
'*--------------------------*
'*      vsGrid1             *
'*--------------------------*
Private Sub vsGrid1_DblClick()
    Call ReturnPress
End Sub
'*---------------------------------------------------------------------------------------------------------------*
'*
'*
'*                                  そ　の　他　フ　ァ　ン　ク　シ　ョ　ン
'*
'*---------------------------------------------------------------------------------------------------------------*
'*---------------------------------------------------------*
'*      Select 文作成   ( pKB : 0/Headより、1/Bodyより )    *
'*---------------------------------------------------------*
Private Sub CrtSelect(pKB As Byte)
    Static sw_Hinbn             As String
    Static sw_Torcd             As String
    Static sw_Ukeir             As String
    
    sw_Hinbn = "":  sw_Torcd = "":  sw_Ukeir = ""
    
    If pKB = 0 Then
'       < 品番 >
        If RTrim$(iH1_Hinbn) <> "" Then
            sw_Hinbn = "                          "
            If InStr(iH1_Hinbn, "*") = 0 Then
                sw_Hinbn = sw_Hinbn & "AND   品番 = '" & RTrim$(iH1_Hinbn) & "'" & vbCrLf
            Else
                sw_Hinbn = sw_Hinbn & "AND   品番 Like '" & Replace(RTrim$(iH1_Hinbn), "*", "%") & "'" & vbCrLf
            End If
        End If
    
'       < 手配先 >
        If RTrim$(iH1_Torcd) <> "" Then
            sw_Torcd = "                          "
            If InStr(iH1_Torcd, "*") = 0 Then
                sw_Torcd = sw_Torcd & "AND   手配先 = '" & RTrim$(iH1_Torcd) & "'" & vbCrLf
            Else
                sw_Torcd = sw_Torcd & "AND   手配先 Like '" & Replace(RTrim$(iH1_Torcd), "*", "%") & "'" & vbCrLf
            End If
        End If
    
'       < 受入 >
        If RTrim$(iH1_Ukeir) <> "" Then
            sw_Ukeir = "                          "
            If InStr(iH1_Ukeir, "*") = 0 Then
                sw_Ukeir = sw_Ukeir & "AND   受入 = '" & RTrim$(iH1_Ukeir) & "'" & vbCrLf
            Else
                sw_Ukeir = sw_Ukeir & "AND   受入 Like '" & Replace(RTrim$(iH1_Ukeir), "*", "%") & "'" & vbCrLf
            End If
        End If
    Else
'       < 品番 >
        sw_Hinbn = "                          AND   品番 = '" & RTrim$(sv_Hinbn) & "'" & vbCrLf
    
'       < 手配先 >
        If RTrim$(sv_Torcd) <> "" Then
            sw_Torcd = "                          AND   手配先 = '" & RTrim$(sv_Torcd) & "'" & vbCrLf
        End If
    
'       < 受入 >
        If RTrim$(sv_Ukeir) <> "" Then
            sw_Ukeir = "                          AND   受入 = '" & RTrim$(sv_Ukeir) & "'" & vbCrLf
        End If
    End If
    
    gSL_Select = "SELECT SK.内示年月, SK.品番, SK.手配先, ISNULL( TR.略称, '' ) 手配先名, SK.受入, ISNULL( HN.表示品番, '' ) 表示品番," & vbCrLf
    gSL_Select = gSL_Select & " SK.対象年月, SK.数量 今回計画数," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( SZ.数量, 0 ) 前回計画数," & vbCrLf
    gSL_Select = gSL_Select & " SK.数量 - ISNULL( SZ.数量, 0 ) 計画差," & vbCrLf
    gSL_Select = gSL_Select & " CEILING( ( ISNULL( SZ.数量, 0 ) / ( SK.数量 * 1.0 ) ) * 100 ) 増減率," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( SH.数量, 0 ) 今回必要数," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( HZ.数量, 0 ) 前回必要数," & vbCrLf
    gSL_Select = gSL_Select & " ISNULL( SH.数量, 0 ) - ISNULL( HZ.数量, 0 ) 必要差" & vbCrLf

'[ 読込基 ( 生産計画テーブル ) ]
    gSL_Select = gSL_Select & " FROM ( SELECT 内示年月, 品番, 手配先, 受入," & vbCrLf
    gSL_Select = gSL_Select & "                          SUBSTRING( 納期, 1, 6 ) 対象年月, SUM(数量) 数量" & vbCrLf
    gSL_Select = gSL_Select & "                          FROM 発注内示テーブル" & vbCrLf
    gSL_Select = gSL_Select & "                          WHERE 内示年月 = '" & gn_HNYM & "'" & vbCrLf
    gSL_Select = gSL_Select & "                          AND   SUBSTRING( 納期, 1, 6 ) = '" & iH1_YM & "'" & vbCrLf
    gSL_Select = gSL_Select & sw_Hinbn & sw_Torcd & sw_Ukeir
    gSL_Select = gSL_Select & "                          GROUP BY 内示年月, 品番, 手配先, 受入, SUBSTRING( 納期, 1, 6 )" & vbCrLf
    gSL_Select = gSL_Select & "      ) SK" & vbCrLf

'[ 生産計画前回テーブル ]
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 内示年月, 品番, 手配先, 受入," & vbCrLf
    gSL_Select = gSL_Select & "                          SUBSTRING( 納期, 1, 6 ) 対象年月, SUM(数量) 数量" & vbCrLf
    gSL_Select = gSL_Select & "                          FROM 発注内示前回テーブル" & vbCrLf
    gSL_Select = gSL_Select & "                          WHERE 内示年月 = '" & gn_HNYM & "'" & vbCrLf
    gSL_Select = gSL_Select & "                          AND   SUBSTRING( 納期, 1, 6 ) = '" & iH1_YM & "'" & vbCrLf
    gSL_Select = gSL_Select & sw_Hinbn & sw_Torcd & sw_Ukeir
    gSL_Select = gSL_Select & "                          GROUP BY 内示年月, 品番, 手配先, 受入, SUBSTRING( 納期, 1, 6 )" & vbCrLf
    gSL_Select = gSL_Select & "                 ) SZ" & vbCrLf
    gSL_Select = gSL_Select & " ON  SZ.内示年月 = SK.内示年月" & vbCrLf
    gSL_Select = gSL_Select & " AND SZ.品番     = SK.品番" & vbCrLf
    gSL_Select = gSL_Select & " AND SZ.手配先   = SK.手配先" & vbCrLf
    gSL_Select = gSL_Select & " AND SZ.受入     = SK.受入" & vbCrLf
    gSL_Select = gSL_Select & " AND SZ.対象年月 = SK.対象年月" & vbCrLf

'[ 生産計画必要数テーブル ]
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 内示年月, 品番, 手配先, 受入," & vbCrLf
    gSL_Select = gSL_Select & "                          SUBSTRING( 納期, 1, 6 ) 対象年月, SUM(数量) 数量" & vbCrLf
    gSL_Select = gSL_Select & "                          FROM 発注内示必要数テーブル" & vbCrLf
    gSL_Select = gSL_Select & "                          WHERE 内示年月 = '" & gn_HNYM & "'" & vbCrLf
    gSL_Select = gSL_Select & "                          AND   SUBSTRING( 納期, 1, 6 ) = '" & iH1_YM & "'" & vbCrLf
    gSL_Select = gSL_Select & sw_Hinbn & sw_Torcd & sw_Ukeir
    gSL_Select = gSL_Select & "                          GROUP BY 内示年月, 品番, 手配先, 受入, SUBSTRING( 納期, 1, 6 )" & vbCrLf
    gSL_Select = gSL_Select & "                 ) SH" & vbCrLf
    gSL_Select = gSL_Select & " ON  SH.内示年月 = SK.内示年月" & vbCrLf
    gSL_Select = gSL_Select & " AND SH.品番     = SK.品番" & vbCrLf
    gSL_Select = gSL_Select & " AND SH.手配先   = SK.手配先" & vbCrLf
    gSL_Select = gSL_Select & " AND SH.受入     = SK.受入" & vbCrLf
    gSL_Select = gSL_Select & " AND SH.対象年月 = SK.対象年月" & vbCrLf

'[ 生産計画必要数前回テーブル ]
    gSL_Select = gSL_Select & " LEFT OUTER JOIN ( SELECT 内示年月, 品番, 手配先, 受入," & vbCrLf
    gSL_Select = gSL_Select & "                          SUBSTRING( 納期, 1, 6 ) 対象年月, SUM(数量) 数量" & vbCrLf
    gSL_Select = gSL_Select & "                          FROM 発注内示必要数前回テーブル" & vbCrLf
    gSL_Select = gSL_Select & "                          WHERE 内示年月 = '" & gn_HNYM & "'" & vbCrLf
    gSL_Select = gSL_Select & "                          AND   SUBSTRING( 納期, 1, 6 ) = '" & iH1_YM & "'" & vbCrLf
    gSL_Select = gSL_Select & sw_Hinbn & sw_Torcd & sw_Ukeir
    gSL_Select = gSL_Select & "                          GROUP BY 内示年月, 品番, 手配先, 受入, SUBSTRING( 納期, 1, 6 )" & vbCrLf
    gSL_Select = gSL_Select & "                 ) HZ" & vbCrLf
    gSL_Select = gSL_Select & " ON  HZ.内示年月 = SK.内示年月" & vbCrLf
    gSL_Select = gSL_Select & " AND HZ.品番     = SK.品番" & vbCrLf
    gSL_Select = gSL_Select & " AND HZ.手配先   = SK.手配先" & vbCrLf
    gSL_Select = gSL_Select & " AND HZ.受入     = SK.受入" & vbCrLf
    gSL_Select = gSL_Select & " AND HZ.対象年月 = SK.対象年月" & vbCrLf

'[ 表示品番 (品番マスタ) ]
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 品番マスタ HN" & vbCrLf
    gSL_Select = gSL_Select & " ON  HN.品番 = SK.品番" & vbCrLf

'[ 手配先名 (取引先マスタ) ]
    gSL_Select = gSL_Select & " LEFT OUTER JOIN 取引先マスタ TR" & vbCrLf
    gSL_Select = gSL_Select & " ON  TR.取引先CD = SK.手配先" & vbCrLf
    gSL_Select = gSL_Select & " AND TR.取引先区分 = 1" & vbCrLf

'[ 増減率 ]
    If pKB = 0 Then
        If CisFun.Val2(iH1_Ritu) <> 0 Then
            gSL_Select = gSL_Select & " WHERE CEILING( ( ISNULL( SZ.数量, 0 ) / ( SK.数量 * 1.0 ) ) * 100 ) <= " & Ritu_Dw & vbCrLf
            gSL_Select = gSL_Select & " OR    CEILING( ( ISNULL( SZ.数量, 0 ) / ( SK.数量 * 1.0 ) ) * 100 ) >= " & Ritu_Up & vbCrLf
        End If
    
        gSL_Select = gSL_Select & " Order By SK.品番, SK.手配先, SK.受入"
    End If
    
End Sub
'*-----------------------*
'*    表示範囲項目表示    *
'*-----------------------*
Private Sub DispRange()
    Ritu_Dw = 100 - CisFun.Val2(iH1_Ritu)
    Ritu_Up = 100 + CisFun.Val2(iH1_Ritu)

    If Ritu_Dw < 0 Then Ritu_Dw = 0
    If Ritu_Up > 999 Then Ritu_Up = 999

    If CisFun.Val2(iH1_Ritu) = 0 Then
        H1lb_Ritu = "全　て"
    Else
        H1lb_Ritu = Ritu_Dw & "%以下 " & Ritu_Up & "%以上"
    End If
    DoEvents
End Sub
